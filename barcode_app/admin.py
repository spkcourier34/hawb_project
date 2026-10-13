from django.contrib import admin
from .models import HAWBRecord, HAWBState


@admin.register(HAWBRecord)
class HAWBRecordAdmin(admin.ModelAdmin):
    list_display  = ('hawb_number', 'run_no', 'mawb', 'consignor', 'destination',
                     'flight_no', 'departure_date', 'is_special')
    list_filter   = ('is_special', 'run_no', 'destination')
    search_fields = ('hawb_number', 'mawb', 'run_no', 'consignor', 'consignee',
                     'invoice_no', 'iec', 'gstin')
    ordering      = ('is_special', 'hawb_number')


@admin.register(HAWBState)
class HAWBStateAdmin(admin.ModelAdmin):
    list_display = ('pk', 'last_generated')
