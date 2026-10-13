"""
barcode_app.views

Extended Django views for HAWB Automation Pro.
This version keeps the existing Excel/XML flow and adds:
- HTML-to-PDF barcode generation for HAWB labels
- PDF to TIF conversion (ported from PDF-To-Tiff.py)
- full-column filtering and sortable record headers
"""

import io
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from urllib.parse import urlencode

import openpyxl
from django.contrib import messages
from django.db import IntegrityError
from django.db.models import Q
from django.http import HttpResponse
from django.shortcuts import redirect, render
from django.views.decorators.http import require_POST
from openpyxl.styles import Alignment, Font, PatternFill

from .forms import GenerateHAWBForm, PDFToTifForm, UploadXMLForm, BagLabelForm
from .models import HAWBRecord, HAWBState

import os
import fitz          # PyMuPDF
from PIL import Image
import base64
from django.template.loader import render_to_string

# Attempt to load label_config
try:
    import label_config
except ImportError:
    class label_config:
        CONSIGNOR_NAME = "SPK COURIER PVT LTD"
        CONSIGNOR_ADDRESS_LINE1 = "KHASRA NO 629, TATA TELCO ROAD"
        CONSIGNOR_ADDRESS_LINE2 = "RANGPURI, NEW DELHI-110037"


# ── PDF-To-Tiff constants (mirrored from PDF-To-Tiff.py) ──────────────────── #
PDF_MAX_BYTES = 1.9 * 1024 * 1024   # 1.9 MB hard ceiling
PDF_DEFAULT_DPI = 150
PDF_MIN_DPI = 72


BASE_DIR = Path(__file__).resolve().parent.parent
DOCX_TEMPLATE_DIR = Path(__file__).resolve().parent / "templates" / "docx_templates"

XML_TAG_ALIASES = {
    "hawb_number": ["HAWB_Number", "HAWBNumber", "HawbNumber", "hawb_number", "HAWB"],
    "no_of_bags": ["No_of_Bags_Pkgs_Pieces_ULD", "No_of_Bags", "NoOfBags", "Pieces", "Pkgs", "ULD", "Bags"],
    "weight": ["Declared_Weight", "Weight", "DeclaredWeight", "Wt"],
    "destination": ["Airport_of_Destination", "Destination", "Dest", "AirportOfDestination", "Airport_Destination"],
    "iec": ["Import_Export_Code", "IEC", "ImportExportCode", "Iec"],
    "mhbs_no": ["MHBS_NO", "MHBS_No", "MHBSNo", "MHBS"],
    "consignor": ["Consignor_Name", "Consignor", "ConsignorName", "ShipperName", "Shipper_Name"],
    "consignee": ["Consignee_Name", "Consignee", "ConsigneeName"],
    "invoice_no": ["Inv_No", "Invoice_No", "InvoiceNo", "InvNo", "Invoice_Number", "InvoiceNumber", "Inv_Number"],
    "invoice_date": ["Inv_Dt", "Invoice_Date", "InvoiceDate", "InvDate", "Inv_Date", "InvoiceDt"],
    "gstin": ["Gstin_id", "GSTIN", "Gstin", "GstinId", "GST_No", "GSTNo"],
    "flight_no": ["Flight_Number", "FlightNumber", "Flight_No", "FlightNo", "Flight"],
    "departure_date": ["Date_of_Departure", "DepartureDate", "Departure_Date", "DateOfDeparture", "Dep_Date", "DepDate", "DepDt"],
}

DISPLAY_COLUMNS = [
    "MAWB",
    "Run No",
    "HAWB",
    "Pkg/Bags",
    "Weight",
    "Dest",
    "IEC",
    "MHBS",
    "Consignor",
    "Consignee",
    "Invoice No",
    "Inv Date",
    "GSTIN",
    "Flight",
    "Dep Date",
    "Special",
    "Status",
]

SORT_FIELD_MAP = {
    "mawb": "mawb",
    "run_no": "run_no",
    "hawb": "hawb_number",
    "no_of_bags": "no_of_bags",
    "weight": "weight",
    "destination": "destination",
    "iec": "iec",
    "mhbs_no": "mhbs_no",
    "consignor": "consignor",
    "consignee": "consignee",
    "invoice_no": "invoice_no",
    "invoice_date": "invoice_date",
    "gstin": "gstin",
    "flight_no": "flight_no",
    "departure_date": "departure_date",
    "special": "is_special",
    "status": "status",
}

SORT_HEADERS = [
    {"label": "MAWB", "sort": "mawb"},
    {"label": "Run No", "sort": "run_no"},
    {"label": "HAWB #", "sort": "hawb"},
    {"label": "Pkg/Bags", "sort": "no_of_bags"},
    {"label": "Weight", "sort": "weight"},
    {"label": "Dest", "sort": "destination"},
    {"label": "IEC", "sort": "iec"},
    {"label": "MHBS No", "sort": "mhbs_no"},
    {"label": "Consignor", "sort": "consignor"},
    {"label": "Consignee", "sort": "consignee"},
    {"label": "Invoice No", "sort": "invoice_no"},
    {"label": "Inv Date", "sort": "invoice_date"},
    {"label": "GSTIN", "sort": "gstin"},
    {"label": "Flight", "sort": "flight_no"},
    {"label": "Dep Date", "sort": "departure_date"},
    {"label": "Type", "sort": "special"},
    {"label": "Status", "sort": "status"},
]

DATE_FORMATS = (
    "%Y-%m-%d",
    "%d-%m-%Y",
    "%d/%m/%Y",
    "%d.%m.%Y",
    "%d %b %Y",
    "%d %B %Y",
    "%Y/%m/%d",
    "%m/%d/%Y",
)


def _resolve_tag(element, key, default=""):
    for tag in XML_TAG_ALIASES.get(key, []):
        val = element.findtext(f".//{tag}")
        if val is not None:
            return val.strip()
    return default


def _get_next_hawb():
    state = HAWBState.get()
    max_db = HAWBRecord.objects.filter(is_special=False).values_list("hawb_number", flat=True)
    nums = [int(h) for h in max_db if h.isdigit() and len(h) == 9]
    base = max(state.last_generated, max(nums) if nums else 0)
    return str(base + 1) if base > 0 else "221001279"


def _set_last_generated(last_value):
    state = HAWBState.get()
    state.last_generated = last_value
    state.save(update_fields=["last_generated"])


def _stream_bytes(data, filename, content_type):
    response = HttpResponse(data, content_type=content_type)
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response


def _build_hawb_barcode_pdf(start_hawb, count):
    """
    Build a multipage A5 PDF with barcode labels.
    Uses HTML templates rendered via PyMuPDF (fitz) for pixel-perfect results.
    """
    import fitz
    import io
    from barcode import Code128
    from barcode.writer import ImageWriter

    pdf_buffer = io.BytesIO()
    doc = fitz.open()
    
    # 148mm x 150mm dimensions in points
    WIDTH, HEIGHT = 419.52, 425.20
    
    for i in range(count):
        hawb = str(start_hawb + i)
        page = doc.new_page(width=WIDTH, height=HEIGHT)
        
        # Generate barcode image (Code128 as requested)
        img_buf = io.BytesIO()
        Code128(hawb, writer=ImageWriter()).write(img_buf, options={
            "module_width": 1.0,   # Very wide bars
            "module_height": 33,   # Taller bars (+1mm)
            "quiet_zone": 1,
            "font_size": 18,
            "text_distance": 9,
        })
        img_buf.seek(0)
        # Convert image to 1-bit monochrome to significantly reduce file size
        img = Image.open(img_buf).convert("1")
        opt_buf = io.BytesIO()
        img.save(opt_buf, format="PNG", optimize=True)
        barcode_b64 = base64.b64encode(opt_buf.getvalue()).decode()
        
        # Render HTML template for this page
        html_content = render_to_string("barcode_app/barcode_pdf_template.html", {
            "hawb": hawb,
            "barcode_base64": barcode_b64,
        })
        
        # Insert HTML into the PDF page
        page.insert_htmlbox(page.rect, html_content)
        
    doc.save(pdf_buffer, deflate=True, garbage=4)
    doc.close()
    return pdf_buffer.getvalue()






def _build_hawb_excel(start, count):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "HAWB Numbers"
    worksheet.column_dimensions["A"].width = 20

    header_cell = worksheet["A1"]
    header_cell.value = "HAWB Number"
    header_cell.font = Font(bold=True, color="FFFFFF", size=12)
    header_cell.fill = PatternFill("solid", fgColor="1558B0")
    header_cell.alignment = Alignment(horizontal="center")

    for offset in range(count):
        worksheet.append([str(start + offset)])

    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer.read()


def _build_query(params, overrides=None, remove=None):
    data = {}
    for key, value in params.items():
        if value not in ("", None):
            data[key] = value
    overrides = overrides or {}
    remove = remove or []
    for key, value in overrides.items():
        if value in ("", None):
            data.pop(key, None)
        else:
            data[key] = value
    for key in remove:
        data.pop(key, None)
    return urlencode(data)


def _parse_flexible_date(value):
    if not value:
        return None
    value = str(value).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return None


def _filter_by_date_range(queryset, from_date, to_date):
    if not from_date and not to_date:
        return queryset

    start_date = _parse_flexible_date(from_date)
    end_date = _parse_flexible_date(to_date)
    matching_ids = []

    for row_id, departure_value in queryset.values_list("id", "departure_date"):
        parsed = _parse_flexible_date(departure_value)
        if not parsed:
            continue
        if start_date and parsed < start_date:
            continue
        if end_date and parsed > end_date:
            continue
        matching_ids.append(row_id)

    return queryset.filter(id__in=matching_ids)


def _apply_record_filters(request, queryset):
    filters = {
        "q": request.GET.get("q", "").strip(),
        "mawb": request.GET.get("mawb", "").strip(),
        "run_no": request.GET.get("run_no", "").strip(),
        "hawb": request.GET.get("hawb", "").strip(),
        "no_of_bags": request.GET.get("no_of_bags", "").strip(),
        "weight": request.GET.get("weight", "").strip(),
        "destination": request.GET.get("destination", "").strip(),
        "iec": request.GET.get("iec", "").strip(),
        "mhbs_no": request.GET.get("mhbs_no", "").strip(),
        "consignor": request.GET.get("consignor", "").strip(),
        "consignee": request.GET.get("consignee", "").strip(),
        "invoice_no": request.GET.get("invoice_no", "").strip(),
        "invoice_date": request.GET.get("invoice_date", "").strip(),
        "gstin": request.GET.get("gstin", "").strip(),
        "flight_no": request.GET.get("flight_no", "").strip(),
        "departure_date": request.GET.get("departure_date", "").strip(),
        "special": request.GET.get("special", "").strip(),
        "from_date": request.GET.get("from_date", "").strip(),
        "to_date": request.GET.get("to_date", "").strip(),
        "status": request.GET.get("status", "").strip(),
        "sort": request.GET.get("sort", "").strip(),
        "order": request.GET.get("order", "").strip().lower(),
    }

    if filters["q"]:
        search = filters["q"]
        queryset = queryset.filter(
            Q(hawb_number__icontains=search)
            | Q(mawb__icontains=search)
            | Q(run_no__icontains=search)
            | Q(no_of_bags__icontains=search)
            | Q(weight__icontains=search)
            | Q(destination__icontains=search)
            | Q(iec__icontains=search)
            | Q(mhbs_no__icontains=search)
            | Q(consignor__icontains=search)
            | Q(consignee__icontains=search)
            | Q(invoice_no__icontains=search)
            | Q(invoice_date__icontains=search)
            | Q(gstin__icontains=search)
            | Q(flight_no__icontains=search)
            | Q(departure_date__icontains=search)
        )

    field_filters = {
        "mawb": "mawb__icontains",
        "run_no": "run_no__icontains",
        "hawb": "hawb_number__icontains",
        "no_of_bags": "no_of_bags__icontains",
        "weight": "weight__icontains",
        "destination": "destination__icontains",
        "iec": "iec__icontains",
        "mhbs_no": "mhbs_no__icontains",
        "consignor": "consignor__icontains",
        "consignee": "consignee__icontains",
        "invoice_no": "invoice_no__icontains",
        "invoice_date": "invoice_date__icontains",
        "gstin": "gstin__icontains",
        "flight_no": "flight_no__icontains",
        "departure_date": "departure_date__icontains",
        "status": "status__icontains",
    }
    for key, lookup in field_filters.items():
        if filters[key]:
            queryset = queryset.filter(**{lookup: filters[key]})

    if filters["special"] == "1":
        queryset = queryset.filter(is_special=True)
    elif filters["special"] == "0":
        queryset = queryset.filter(is_special=False)

    queryset = _filter_by_date_range(queryset, filters["from_date"], filters["to_date"])

    sort_key = filters["sort"]
    sort_order = "desc" if filters["order"] == "desc" else "asc"
    if sort_key in SORT_FIELD_MAP:
        sort_field = SORT_FIELD_MAP[sort_key]
        queryset = queryset.order_by(f"-{sort_field}" if sort_order == "desc" else sort_field)

    return queryset, filters


def _build_sort_headers(filters):
    headers = []
    for header in SORT_HEADERS:
        current = filters.get("sort") == header["sort"]
        current_order = filters.get("order", "asc")
        next_order = "desc" if current and current_order == "asc" else "asc"
        query_string = _build_query(filters, overrides={"sort": header["sort"], "order": next_order})
        headers.append(
            {
                "label": header["label"],
                "url": f"?{query_string}" if query_string else "?",
                "active": current,
                "order": current_order if current else "",
            }
        )
    return headers


def home(request):
    state = HAWBState.get()
    total = HAWBRecord.objects.count()
    total_special = HAWBRecord.objects.filter(is_special=True).count()
    last_record = HAWBRecord.objects.filter(is_special=False).order_by("-hawb_number").first()
    next_hawb = _get_next_hawb()

    return render(
        request,
        "barcode_app/home.html",
        {
            "next_hawb": next_hawb,
            "total": total,
            "total_special": total_special,
            "last_generated": state.last_generated or "-",
            "last_record": last_record,
        },
    )


def generate_hawb(request):
    next_hawb = _get_next_hawb()

    if request.method == "POST":
        form = GenerateHAWBForm(request.POST)
        if form.is_valid():
            start = int(form.cleaned_data["start_hawb"])
            count = form.cleaned_data["count"]
            last = start + count - 1
            action = request.POST.get("action", "excel")

            try:
                _set_last_generated(last)
                
                state = HAWBState.get()
                state.last_batch_start = start
                state.last_batch_count = count
                state.save(update_fields=["last_batch_start", "last_batch_count"])

                if action == "excel":
                    excel_bytes = _build_hawb_excel(start, count)
                    return _stream_bytes(
                        excel_bytes,
                        f"HAWB_{start}_to_{last}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                # HTML to PDF conversion (replaces DOCX)
                pdf_bytes = _build_hawb_barcode_pdf(start, count)
                return _stream_bytes(
                    pdf_bytes,
                    f"HAWB_Barcodes_{start}_to_{last}.pdf",
                    "application/pdf",
                )
            except Exception as exc:
                messages.error(request, str(exc))
    else:
        form = GenerateHAWBForm(initial={"start_hawb": next_hawb})

    return render(
        request,
        "barcode_app/generate_hawb.html",
        {
            "form": form,
            "next_hawb": next_hawb,
            "last_batch_start": HAWBState.get().last_batch_start,
            "last_batch_count": HAWBState.get().last_batch_count,
            "last_batch_end": HAWBState.get().last_batch_start + HAWBState.get().last_batch_count - 1 if HAWBState.get().last_batch_count > 0 else 0,
        },
    )


# ── PDF → TIFF Conversion (from PDF-To-Tiff.py) ───────────────────────────── #

def _sizeof_fmt(num, suffix="B"):
    """Human-readable file size (ported from PDF-To-Tiff.py sizeof_fmt)."""
    for unit in ["", "K", "M", "G"]:
        if abs(num) < 1024.0:
            return "%3.1f%s%s" % (num, unit, suffix)
        num /= 1024.0
    return "%.1f%s%s" % (num, "T", suffix)


def _pdf_to_multipage_tiff(pdf_path, out_tiff_path,
                           dpi=PDF_DEFAULT_DPI,
                           max_bytes=PDF_MAX_BYTES,
                           auto_optimize=True,
                           progress_callback=None):
    """
    Convert a PDF to a multipage TIFF, auto-tuning DPI to stay within
    max_bytes (default 1.9 MB).  Exact port of pdf_to_multipage_tiff()
    from PDF-To-Tiff.py.

    Returns (page_count, file_size_bytes, used_dpi, was_adjusted).
    """
    doc = fitz.open(pdf_path)
    page_count = doc.page_count

    def render_and_save(current_dpi, temp_path=None):
        zoom = current_dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        frames = []
        for i in range(page_count):
            page = doc.load_page(i)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            frames.append(img)
            if progress_callback:
                progress_callback(i + 1, page_count)
        save_path = temp_path or out_tiff_path
        frames[0].save(
            save_path,
            format="TIFF",
            save_all=True,
            append_images=frames[1:],
            compression="tiff_deflate",
        )
        return os.path.getsize(save_path)

    current_dpi = dpi
    size = render_and_save(current_dpi)
    adjusted = False

    if auto_optimize:
        # ↓ Reduce DPI until under limit (same logic as PDF-To-Tiff.py)
        while size > max_bytes and current_dpi > PDF_MIN_DPI:
            current_dpi = int(current_dpi * 0.85)
            size = render_and_save(current_dpi)
            adjusted = True

        # ↑ Try increasing quality if well under limit
        while size < (max_bytes * 0.75) and current_dpi < 400:
            next_dpi = int(current_dpi * 1.15)
            test_path = str(out_tiff_path) + ".tmp"
            test_size = render_and_save(next_dpi, temp_path=test_path)
            if test_size <= max_bytes:
                os.replace(test_path, out_tiff_path)
                current_dpi = next_dpi
                size = test_size
            else:
                if os.path.exists(test_path):
                    os.remove(test_path)
                break

    return page_count, size, current_dpi, adjusted


def convert_pdf_to_tif_view(request):
    """
    Django view for the PDF → TIFF converter.  Incorporates all logic from
    the standalone PDF-To-Tiff.py application:
      - Multi-file upload
      - Auto-optimize DPI (stay within 1.9 MB)
      - Custom DPI override (mutual exclusivity with auto-optimize)
      - Duplicate filename handling
      - Color-coded result messages (green / yellow / red)
      - Single TIF download or ZIP for multiple files
    """
    if request.method == "POST":
        form = PDFToTifForm(request.POST, request.FILES)

        if form.is_valid():
            pdf_files   = form.cleaned_data.get("pdf_files", [])
            auto_opt    = form.cleaned_data.get("auto_optimize", True)
            use_custom  = form.cleaned_data.get("use_custom_dpi", False)
            custom_dpi  = form.cleaned_data.get("custom_dpi") or PDF_DEFAULT_DPI

            # Mirror desktop app: custom DPI wins over auto-optimize
            dpi = custom_dpi if use_custom else PDF_DEFAULT_DPI
            if use_custom:
                auto_opt = False

            if not pdf_files:
                messages.error(request, "No PDF files uploaded.")
                return render(request, "barcode_app/convert_pdf_to_tif.html", {"form": form})

            tif_results = []   # list of (filename, bytes)
            errors      = []

            with tempfile.TemporaryDirectory(prefix="pdf_to_tif_") as tmp_dir:
                tmp = Path(tmp_dir)
                used_names = set()

                for pdf_file in pdf_files:
                    base_name  = Path(pdf_file.name).stem
                    safe_name  = "".join(
                        ch if ch.isalnum() or ch in ("-", "_") else "_"
                        for ch in base_name
                    ).strip("_") or "converted"

                    # Duplicate handling (same as desktop app)
                    final_name = safe_name
                    counter = 1
                    while final_name in used_names:
                        final_name = f"{safe_name}_{counter}"
                        counter += 1
                    used_names.add(final_name)

                    pdf_path = tmp / f"{final_name}.pdf"
                    tif_path = tmp / f"{final_name}.tif"

                    try:
                        # Save uploaded file to temp directory
                        with pdf_path.open("wb") as f:
                            for chunk in pdf_file.chunks():
                                f.write(chunk)

                        pages, size, used_dpi, adjusted = _pdf_to_multipage_tiff(
                            str(pdf_path),
                            str(tif_path),
                            dpi=dpi,
                            max_bytes=PDF_MAX_BYTES,
                            auto_optimize=auto_opt,
                        )

                        ratio    = size / PDF_MAX_BYTES
                        size_str = _sizeof_fmt(size)

                        # Mirror color-coded log messages from desktop app
                        if adjusted:
                            tag = "🔴"
                            msg = f"{tag} {pdf_file.name} → {size_str} @ {used_dpi} DPI ({pages} pages, DPI reduced to fit)"
                        # Warn if custom DPI produced oversized file (same as desktop app)
                        if use_custom and size > PDF_MAX_BYTES:
                            messages.error(
                                request,
                                f"⚠️ {pdf_file.name} is {_sizeof_fmt(size)} — over 1.9 MB. "
                                "Reduce DPI or enable Auto Optimize.",
                                extra_tags='tif-log'
                            )

                        tif_results.append((tif_path.name, tif_path.read_bytes()))

                    except Exception as exc:
                        errors.append(f"{pdf_file.name}: {exc}")

                # Summary message instead of per-file messages to prevent Dashboard clutter
                if tif_results:
                    success_count = len(tif_results)
                    if success_count == 1:
                        messages.success(request, f"🟢 {pdf_files[0].name} converted successfully.", extra_tags='tif-log')
                    else:
                        messages.success(request, f"🟢 Successfully converted {success_count} files.", extra_tags='tif-log')

                # Report any errors
                for err in errors:
                    messages.error(request, f"❌ {err}", extra_tags='tif-log')

            if not tif_results:
                return render(request, "barcode_app/convert_pdf_to_tif.html", {"form": form})

            # Single file → direct download
            if len(tif_results) == 1:
                filename, data = tif_results[0]
                response = HttpResponse(data, content_type="image/tiff")
                response["Content-Disposition"] = f'attachment; filename="{filename}"'
                response.set_cookie("download_complete", "true", max_age=30)
                return response

            # Multiple files → ZIP (same as desktop app's output folder)
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for filename, data in tif_results:
                    zf.writestr(filename, data)
            zip_buf.seek(0)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            response = HttpResponse(zip_buf.read(), content_type="application/zip")
            response["Content-Disposition"] = f'attachment; filename="converted_tifs_{ts}.zip"'
            response.set_cookie("download_complete", "true", max_age=30)
            return response

    else:
        form = PDFToTifForm(initial={"auto_optimize": True})

    return render(request, "barcode_app/convert_pdf_to_tif.html", {"form": form})


def upload_xml(request):
    if request.method == "POST":
        confirm_action = request.POST.get("confirm_action")
        
        if confirm_action == "execute":
            # --- PHASE 2: RESOLUTION ---
            records_to_process = request.session.get("pending_xml_data", [])
            mawb = request.POST.get("mawb", "")
            run_no = request.POST.get("run_no", "")
            dup_mode = request.POST.get("dup_mode", "skip")
            header_flight = request.session.get("pending_header_flight", "")
            header_dep_dt = request.session.get("pending_header_dep_dt", "")
            
            if not records_to_process:
                messages.error(request, "Session expired or no data found. Please upload again.")
                return redirect("upload_xml")
        else:
            # --- PHASE 1: INITIAL UPLOAD ---
            form = UploadXMLForm(request.POST, request.FILES)
            if not form.is_valid():
                return render(request, "barcode_app/upload_xml.html", {"form": form, "supported_tags": XML_TAG_ALIASES})
            
            mawb = form.cleaned_data["mawb"]
            run_no = form.cleaned_data["run_no"]
            xml_file = request.FILES["xml_file"]

            try:
                root = ET.parse(xml_file).getroot()
            except ET.ParseError as exc:
                messages.error(request, f"Invalid XML file: {exc}")
                return render(request, "barcode_app/upload_xml.html", {"form": form, "supported_tags": XML_TAG_ALIASES})

            # Parse headers and tags
            header_flight = ""
            header_dep_dt = ""
            header_el = root.find(".//Header")
            if header_el is not None:
                for tag in ["Flight_Number", "FlightNumber", "Flight_No", "Flight"]:
                    value = header_el.findtext(tag, "").strip()
                    if value:
                        header_flight = value
                        break
                for tag in ["Date_of_Departure", "DepartureDate", "Departure_Date", "DateOfDeparture", "Dep_Date", "DepDate"]:
                    value = header_el.findtext(tag, "").strip()
                    if value:
                        header_dep_dt = value
                        break

            hawb_tag = None
            for alias in XML_TAG_ALIASES["hawb_number"]:
                if root.findall(f".//{alias}"):
                    hawb_tag = alias
                    break

            if not hawb_tag:
                messages.error(request, "No HAWB records found in this XML.")
                return render(request, "barcode_app/upload_xml.html", {"form": form, "supported_tags": XML_TAG_ALIASES})

            hawb_elements = root.findall(f".//{hawb_tag}")
            all_hawbs = []
            for element in hawb_elements:
                hawb = (element.text or element.get("value", "")).strip()
                if hawb and hawb.isdigit():
                    print(f"DEBUG: Found HAWB in XML: {hawb}")
                    all_hawbs.append(hawb)

            if not all_hawbs:
                messages.error(request, "No numeric HAWB numbers found.")
                return render(request, "barcode_app/upload_xml.html", {"form": form, "supported_tags": XML_TAG_ALIASES})

            normal_hawbs = {hawb for hawb in all_hawbs if len(hawb) == 9}
            special_hawbs = {hawb for hawb in all_hawbs if hawb.isdigit() and len(hawb) != 9}

            # Block intra-XML duplicates
            seen = {}
            intra_dups = []
            for h in all_hawbs:
                if h in seen: intra_dups.append(h)
                seen[h] = True
            if intra_dups:
                messages.error(request, f"Duplicate HAWBs inside XML: {', '.join(set(intra_dups))}")
                return render(request, "barcode_app/upload_xml.html", {"form": form, "supported_tags": XML_TAG_ALIASES})

            print(f"DEBUG: Found {len(hawb_elements)} HAWB elements in XML.")
            # Prepare records
            records_to_process = []
            for element in hawb_elements:
                hawb = (element.text or element.get("value", "")).strip()
                if not hawb or not hawb.isdigit(): continue
                is_special = hawb in special_hawbs
                is_normal = hawb in normal_hawbs
                if not is_special and not is_normal: continue
                
                records_to_process.append({
                    "hawb": hawb,
                    "no_of_bags": _resolve_tag(element, "no_of_bags"),
                    "weight": _resolve_tag(element, "weight"),
                    "destination": _resolve_tag(element, "destination"),
                    "iec": _resolve_tag(element, "iec"),
                    "mhbs_no": _resolve_tag(element, "mhbs_no"),
                    "consignor": _resolve_tag(element, "consignor"),
                    "consignee": _resolve_tag(element, "consignee"),
                    "invoice_no": _resolve_tag(element, "invoice_no"),
                    "invoice_date": _resolve_tag(element, "invoice_date"),
                    "gstin": _resolve_tag(element, "gstin"),
                    "is_special": is_special,
                    "is_normal": is_normal,
                })

            # Check DB for duplicates
            hawb_list = [r["hawb"] for r in records_to_process]
            existing_hawbs = set(HAWBRecord.objects.filter(hawb_number__in=hawb_list).values_list("hawb_number", flat=True))

            if existing_hawbs:
                request.session["pending_xml_data"] = records_to_process
                request.session["pending_header_flight"] = header_flight
                request.session["pending_header_dep_dt"] = header_dep_dt
                return render(request, "barcode_app/confirm_upload.html", {
                    "mawb": mawb,
                    "run_no": run_no,
                    "new_count": len(records_to_process) - len(existing_hawbs),
                    "dup_count": len(existing_hawbs),
                })
            
            dup_mode = "skip"

        # --- Execution Phase ---
        # --- EXECUTION PHASE ---
        inserted_normal, inserted_special, updated_count, skipped_dup = 0, 0, 0, 0
        xml_normal_nums = []

        import os
        from datetime import datetime
        from django.conf import settings
        log_file = os.path.join(settings.BASE_DIR, "debug_log.txt")
        with open(log_file, "a") as f:
            f.write(f"\n--- Upload execution started at {datetime.now()} ---\n")
            f.write(f"Executing with {len(records_to_process)} records. Dup mode: {dup_mode}\n")

        print(f"DEBUG: Starting execution phase for {len(records_to_process)} records.")
        for i, r in enumerate(records_to_process):
            hawb = r["hawb"]
            is_special = r["is_special"]
            if not is_special: xml_normal_nums.append(int(hawb))

            if dup_mode == "replace":
                # Update or Create
                try:
                    obj, created = HAWBRecord.objects.update_or_create(
                        hawb_number=hawb,
                        defaults={
                            "mawb": mawb,
                            "run_no": run_no,
                            "no_of_bags": r["no_of_bags"],
                            "weight": r["weight"],
                            "destination": r["destination"],
                            "iec": r["iec"],
                            "mhbs_no": r["mhbs_no"],
                            "consignor": r["consignor"],
                            "consignee": r["consignee"],
                            "invoice_no": r["invoice_no"],
                            "invoice_date": r["invoice_date"],
                            "gstin": r["gstin"],
                            "flight_no": header_flight,
                            "departure_date": header_dep_dt,
                            "is_special": is_special,
                        }
                    )
                    if created:
                        if is_special: inserted_special += 1
                        else: inserted_normal += 1
                    else:
                        updated_count += 1
                except Exception as e:
                    with open(log_file, "a") as f:
                        f.write(f"Error replacing {hawb}: {e}\n")
                    skipped_dup += 1
            else:
                # Skip Mode (Default)
                try:
                    HAWBRecord.objects.create(
                        mawb=mawb,
                        run_no=run_no,
                        hawb_number=hawb,
                        no_of_bags=r["no_of_bags"],
                        weight=r["weight"],
                        destination=r["destination"],
                        iec=r["iec"],
                        mhbs_no=r["mhbs_no"],
                        consignor=r["consignor"],
                        consignee=r["consignee"],
                        invoice_no=r["invoice_no"],
                        invoice_date=r["invoice_date"],
                        gstin=r["gstin"],
                        flight_no=header_flight,
                        departure_date=header_dep_dt,
                        is_special=is_special,
                    )
                    if is_special: inserted_special += 1
                    else: inserted_normal += 1
                    with open(log_file, "a") as f:
                        f.write(f"Index {i}: Successfully saved {hawb}\n")
                except Exception as e:
                    print(f"DEBUG: Error saving HAWB {hawb}: {e}")
                    with open(log_file, "a") as f:
                        f.write(f"Index {i}: Error saving {hawb} - {e}\n")
                    skipped_dup += 1

        with open(log_file, "a") as f:
            f.write(f"Finished execution. Normal: {inserted_normal}, Special: {inserted_special}, Updated: {updated_count}, Skipped: {skipped_dup}\n")

            if xml_normal_nums:
                _set_last_generated(max(xml_normal_nums))
                state = HAWBState.get()
                if state.last_batch_start > 0:
                    state.last_batch_start = 0
                    state.last_batch_count = 0
                    state.save(update_fields=['last_batch_start', 'last_batch_count'])

            # Cleanup session
            request.session.pop("pending_xml_data", None)
            request.session.pop("pending_header_flight", None)
            request.session.pop("pending_header_dep_dt", None)

            summary = f"Upload complete - Normal: {inserted_normal} | Special: {inserted_special}"
            if updated_count > 0: summary += f" | Updated: {updated_count}"
            if skipped_dup > 0: summary += f" | Skipped: {skipped_dup}"

            messages.success(request, summary)
            return redirect("view_records")
    else:
        form = UploadXMLForm()

    return render(
        request,
        "barcode_app/upload_xml.html",
        {"form": form, "supported_tags": XML_TAG_ALIASES},
    )


def view_records(request):
    queryset, filters = _apply_record_filters(request, HAWBRecord.objects.all())
    total = HAWBRecord.objects.count()
    showing = queryset.count()

    return render(
        request,
        "barcode_app/view_records.html",
        {
            "records": queryset,
            "total": total,
            "showing": showing,
            "filters": filters,
            "sort_headers": _build_sort_headers(filters),
            "export_query": _build_query(filters),
        },
    )


def export_records(request):
    queryset, _filters = _apply_record_filters(request, HAWBRecord.objects.all())

    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "HAWB Records"

    for column_index, header in enumerate(DISPLAY_COLUMNS, 1):
        cell = worksheet.cell(row=1, column=column_index, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1558B0")
        cell.alignment = Alignment(horizontal="center")

    for record in queryset:
        worksheet.append(
            [
                record.mawb,
                record.run_no,
                record.hawb_number,
                record.no_of_bags,
                record.weight,
                record.destination,
                record.iec,
                record.mhbs_no,
                record.consignor,
                record.consignee,
                record.invoice_no,
                record.invoice_date,
                record.gstin,
                record.flight_no,
                record.departure_date,
                "Special" if record.is_special else "",
                record.status,
            ]
        )

    for column in worksheet.columns:
        max_len = max((len(str(cell.value or "")) for cell in column), default=10)
        worksheet.column_dimensions[column[0].column_letter].width = min(max_len + 4, 40)

    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    return _stream_bytes(
        buffer.read(),
        f"HAWB_Records_{timestamp}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@require_POST
def delete_records(request):
    ids = request.POST.getlist("ids")
    if ids:
        deleted, _detail = HAWBRecord.objects.filter(pk__in=ids).delete()
        messages.success(request, f"{deleted} record(s) deleted.")
    else:
        messages.warning(request, "No records selected for deletion.")
    return redirect("view_records")


# ── Bag Label Generator ───────────────────────────────────────────────────── #

def _build_bag_barcode_pdf(run_number, count):
    """
    Build a multipage A5 PDF with Bag Labels.
    Uses HTML template rendered via PyMuPDF (fitz) for perfect layout.
    """
    import fitz
    import io
    from barcode import Code128
    from barcode.writer import ImageWriter

    pdf_buffer = io.BytesIO()
    doc = fitz.open()

    # 148mm x 150mm dimensions in points
    WIDTH, HEIGHT = 419.52, 425.20

    for i in range(1, count + 1):
        bag_number = f"{run_number}BAG{i}"
        page = doc.new_page(width=WIDTH, height=HEIGHT)

        # Generate barcode image
        img_buf = io.BytesIO()
        Code128(bag_number, writer=ImageWriter()).write(img_buf, options={
            "module_width": 1.0,   # Even wider bars
            "module_height": 33,   # Taller bars (+1mm)
            "quiet_zone": 1,
            "font_size": 18,       # Larger text under barcode
            "text_distance": 9,
        })
        img_buf.seek(0)
        # Convert image to 1-bit monochrome to significantly reduce file size
        img = Image.open(img_buf).convert("1")
        opt_buf = io.BytesIO()
        img.save(opt_buf, format="PNG", optimize=True)
        barcode_b64 = base64.b64encode(opt_buf.getvalue()).decode()

        # Render HTML template for this page
        html_content = render_to_string("barcode_app/bag_label_template.html", {
            "bag_number": bag_number,
            "barcode_base64": barcode_b64,
            "consignor_name": getattr(label_config, "CONSIGNOR_NAME", "SPK COURIER PVT LTD"),
            "consignor_address_line1": getattr(label_config, "CONSIGNOR_ADDRESS_LINE1", "KHASRA NO 629, TATA TELCO ROAD"),
            "consignor_address_line2": getattr(label_config, "CONSIGNOR_ADDRESS_LINE2", "RANGPURI, NEW DELHI-110037"),
        })

        # Insert HTML into the PDF page
        page.insert_htmlbox(page.rect, html_content)

    doc.save(pdf_buffer, deflate=True, garbage=4)
    doc.close()
    return pdf_buffer.getvalue()


def generate_bag_labels(request):
    if request.method == "POST":
        form = BagLabelForm(request.POST)
        if form.is_valid():
            run_number = form.cleaned_data["run_number"]
            count = form.cleaned_data["count"]

            try:
                pdf_bytes = _build_bag_barcode_pdf(run_number, count)
                return _stream_bytes(
                    pdf_bytes,
                    f"{run_number}_Bag_Labels_1_to_{count}.pdf",
                    "application/pdf",
                )
            except Exception as exc:
                messages.error(request, f"Error generating labels: {str(exc)}")
    else:
        form = BagLabelForm()

    # Last Run Summary logic
    all_dests = HAWBRecord.objects.values_list('destination', flat=True)
    unique_dests = sorted(list(set(d.strip().upper() for d in all_dests if d)))
    
    run_summary = []
    for d in unique_dests:
        latest = HAWBRecord.objects.filter(destination__iexact=d).order_by('-run_no').first()
        if latest:
            run_summary.append({'dest': d, 'run_no': latest.run_no})

    return render(request, "barcode_app/generate_bag_labels.html", {"form": form, "run_summary": run_summary})


def download_status_template(request):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Status Update"
    worksheet.column_dimensions["A"].width = 20
    worksheet.column_dimensions["B"].width = 20

    headers = ["HAWB Number", "Status"]
    for col_num, header in enumerate(headers, 1):
        cell = worksheet.cell(row=1, column=col_num, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1558B0")
        cell.alignment = Alignment(horizontal="center")

    worksheet.append(["10000001", "BTT"])

    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return _stream_bytes(
        buffer.read(),
        "Status_Update_Template.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@require_POST
def bulk_update_status(request):
    excel_file = request.FILES.get("status_excel")
    if not excel_file:
        messages.error(request, "No Excel file uploaded.")
        return redirect("view_records")
    
    try:
        workbook = openpyxl.load_workbook(excel_file, data_only=True)
        worksheet = workbook.active
        
        updated_count = 0
        not_found_count = 0
        
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
                
            hawb = str(row[0]).strip()
            # Handle float representation of strings if Excel format is weird
            if hawb.endswith('.0'):
                hawb = hawb[:-2]
                
            status = str(row[1]).strip() if row[1] is not None else ""
            
            if hawb:
                try:
                    record = HAWBRecord.objects.get(hawb_number=hawb)
                    record.status = status
                    record.save(update_fields=['status'])
                    updated_count += 1
                except HAWBRecord.DoesNotExist:
                    not_found_count += 1
                    
        summary = f"Status updated for {updated_count} records."
        if not_found_count > 0:
            summary += f" {not_found_count} records from Excel were not found in the database."
        messages.success(request, summary)
        
    except Exception as e:
        messages.error(request, f"Error processing Excel file: {e}")
        
    return redirect("view_records")

def redownload_hawb_batch(request):
    state = HAWBState.get()
    start = state.last_batch_start
    count = state.last_batch_count
    if start == 0 or count == 0:
        messages.error(request, "No batch available for re-download.")
        return redirect("generate_hawb")
        
    action = request.GET.get("action", "excel")
    last = start + count - 1
    
    try:
        if action == "excel":
            excel_bytes = _build_hawb_excel(start, count)
            return _stream_bytes(
                excel_bytes,
                f"HAWB_{start}_to_{last}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            pdf_bytes = _build_hawb_barcode_pdf(start, count)
            return _stream_bytes(
                pdf_bytes,
                f"HAWB_Barcodes_{start}_to_{last}.pdf",
                "application/pdf",
            )
    except Exception as exc:
        messages.error(request, str(exc))
        return redirect("generate_hawb")