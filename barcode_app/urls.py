from django.urls import path
from . import views

urlpatterns = [
    path('',                 views.home,                    name='home'),
    path('generate/',        views.generate_hawb,           name='generate_hawb'),
    path('generate/bag/',    views.generate_bag_labels,     name='generate_bag_labels'),
    path('pdf-to-tif/',      views.convert_pdf_to_tif_view, name='convert_pdf_to_tif'),
    path('upload/',          views.upload_xml,              name='upload_xml'),
    path('records/',         views.view_records,            name='view_records'),
    path('records/export/',  views.export_records,          name='export_records'),
    path('records/delete/',  views.delete_records,          name='delete_records'),
    path('records/status/template/', views.download_status_template, name='download_status_template'),
    path('records/status/upload/',   views.bulk_update_status,       name='bulk_update_status'),
    path('generate/redownload/',     views.redownload_hawb_batch,    name='redownload_hawb_batch'),
]
