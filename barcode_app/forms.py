from django import forms


class GenerateHAWBForm(forms.Form):
    start_hawb = forms.CharField(
        label='Start HAWB Number',
        max_length=9,
        widget=forms.TextInput(attrs={
            'placeholder': 'e.g. 221001279',
            'class': 'form-input',
            'autocomplete': 'off',
        }),
    )
    count = forms.IntegerField(
        label='Count',
        min_value=1,
        max_value=100000,
        widget=forms.NumberInput(attrs={
            'placeholder': 'e.g. 50',
            'class': 'form-input',
        }),
    )

    def clean_start_hawb(self):
        val = self.cleaned_data['start_hawb'].strip()
        if not val.isdigit():
            raise forms.ValidationError('Start HAWB must contain only digits.')
        if len(val) != 9:
            raise forms.ValidationError(
                f'Start HAWB must be exactly 9 digits (you entered {len(val)}).'
            )
        return val

    def clean(self):
        cleaned = super().clean()
        start = cleaned.get('start_hawb')
        count = cleaned.get('count')
        if start and count:
            last = int(start) + count - 1
            if last > 999_999_999:
                raise forms.ValidationError(
                    f'Range exceeds 9-digit limit. '
                    f'Max count from {start} is {999_999_999 - int(start) + 1}.'
                )
        return cleaned


class UploadXMLForm(forms.Form):
    mawb = forms.CharField(max_length=20)
    run_no = forms.CharField(max_length=50)
    xml_file = forms.FileField()


# ── PDF → TIFF Converter Form ─────────────────────────────────────────────── #

class MultipleFileInput(forms.ClearableFileInput):
    """File input that allows multiple files to be selected."""
    allow_multiple_selected = True

    def __init__(self, attrs=None):
        defaults = {'multiple': True}
        if attrs:
            defaults.update(attrs)
        super().__init__(attrs=defaults)

    def value_from_datadict(self, data, files, name):
        return files.getlist(name)


class MultipleFileField(forms.FileField):
    """FileField that accepts and validates a list of files."""

    def __init__(self, *args, **kwargs):
        kwargs.setdefault('widget', MultipleFileInput())
        super().__init__(*args, **kwargs)

    def clean(self, data, initial=None):
        if not data:
            if self.required:
                raise forms.ValidationError(self.error_messages['required'])
            return []
        result = []
        for f in data:
            result.append(super().clean(f, initial))
        return result


class PDFToTifForm(forms.Form):
    """
    Mirrors the options available in the standalone PDF-To-Tiff.py GUI:
      - Multi-file PDF upload
      - Auto Optimize DPI (reduce/increase to stay within 1.9 MB)
      - Custom DPI override
    """
    DEFAULT_DPI = 150

    pdf_files = MultipleFileField(
        label='PDF File(s)',
        widget=MultipleFileInput(attrs={
            'class': 'form-input',
            'accept': '.pdf,application/pdf',
        }),
    )

    auto_optimize = forms.BooleanField(
        label='Auto Optimize DPI (target ≤1.9 MB)',
        required=False,
        initial=True,
        widget=forms.CheckboxInput(attrs={'id': 'id_auto_optimize'}),
    )

    use_custom_dpi = forms.BooleanField(
        label='Use Custom DPI',
        required=False,
        initial=False,
        widget=forms.CheckboxInput(attrs={'id': 'id_use_custom_dpi'}),
    )

    custom_dpi = forms.IntegerField(
        label='Custom DPI',
        required=False,
        min_value=72,
        max_value=600,
        initial=DEFAULT_DPI,
        widget=forms.NumberInput(attrs={
            'class': 'form-input',
            'id': 'id_custom_dpi',
            'placeholder': '150',
            'style': 'width:90px;',
        }),
    )

    def clean(self):
        cleaned = super().clean()
        use_custom = cleaned.get('use_custom_dpi')
        auto = cleaned.get('auto_optimize')
        dpi = cleaned.get('custom_dpi')

        # Mutual exclusivity: custom DPI and auto-optimize cannot both be active
        if use_custom and auto:
            # Custom DPI wins — same logic as the desktop app
            cleaned['auto_optimize'] = False

        if use_custom and not dpi:
            raise forms.ValidationError('Please enter a DPI value when "Use Custom DPI" is selected.')

        return cleaned


# ── Bag Label Generator Form ──────────────────────────────────────────────── #

class BagLabelForm(forms.Form):
    run_number = forms.CharField(
        label='Run Number',
        max_length=30,
        widget=forms.TextInput(attrs={
            'placeholder': 'e.g. SPKLHR103',
            'class': 'form-input mono',
            'autocomplete': 'off',
        }),
    )
    count = forms.IntegerField(
        label='Number of Bags',
        min_value=1,
        max_value=500,
        widget=forms.NumberInput(attrs={
            'placeholder': 'e.g. 50',
            'class': 'form-input',
        }),
    )

    def clean_run_number(self):
        val = self.cleaned_data['run_number'].strip().upper()
        if not val:
            raise forms.ValidationError('Run number is required.')
        return val
