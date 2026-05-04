from django.urls import path
from . import views

urlpatterns = [
    # --- MAIN DASHBOARD & STATS ---
    path('', views.home, name='home'),
    path('run-sheet/', views.run_sheet, name='run_sheet'),
    path('stats/', views.stats, name='stats'),
    path('history/', views.run_sheet_history, name='run_sheet_history'),

    # --- RUN SHEET ACTIONS ---
    path('add-to-run-sheet/', views.add_to_run_sheet, name='add_to_run_sheet'),
    path('entry-form/<str:customer_id>/', views.entry_form, name='entry_form'),
    path('order/edit/<int:pk>/', views.edit_order, name='edit_order'),
    path('delete-stop/<int:pk>/', views.delete_stop, name='delete_stop'),
    path('finalize/', views.finalize_run_sheet, name='finalize_run_sheet'),
    path('export-excel/', views.export_run_sheet_excel, name='export_run_sheet_excel'),
    path('upload/', views.upload_run_sheet, name='upload_run_sheet'),
    path('clear-sheet/', views.clear_run_sheet, name='clear_run_sheet'),
    path('commit-day/', views.commit_and_clear_day, name='commit_and_clear_day'),

    # --- CUSTOMER DATABASE ---
    path('customers/', views.customer_list, name='customer_list'),
    path('database/', views.manage_customers, name='manage_customers'),
    path('database/add/', views.edit_customer, name='add_customer'),
    path('database/edit/<str:pk>/', views.edit_customer, name='edit_customer'),

    # --- RETURNS ---
    path('add-return/', views.add_return, name='add_return'),
    path('select-return/', views.select_customer_return, name='select_customer_return'),
    path('return-form/<str:customer_id>/', views.add_return_form, name='add_return_form'),

    # --- VENDOR PICKUPS (TRUCKS) ---
    path('vendors/', views.manage_vendors, name='manage_vendors'),
    path('vendors/add/', views.edit_vendor, name='add_vendor'),
    path('vendors/edit/<int:pk>/', views.edit_vendor, name='edit_vendor'),
    path('add-pickup/', views.add_pickup, name='add_pickup'),
    path('select-vendor/', views.select_vendor_pickup, name='select_vendor_pickup'),
    path('pickup-form/<int:vendor_id>/', views.add_pickup_form, name='add_pickup_form'),

    # --- CUSTOMER PICKUPS (COUNTER LOG) ---
    path('pickups/', views.pickup_log_list, name='pickup_log_list'),
    path('pickups/select/', views.select_customer_pickup, name='select_customer_pickup'),
    path('pickups/add/<str:customer_id>/', views.add_pickup_order, name='add_pickup_order'),
    path('pickups/day/<str:date>/', views.daily_pickup_detail, name='daily_pickup_detail'),
    path('pickups/edit/<int:pk>/', views.edit_pickup_order, name='edit_pickup_order'),
    path('pickups/delete/<int:pk>/', views.delete_pickup_order, name='delete_pickup_order'),

    # --- OUTBOUND PHOTOS ---
    path('outbound/', views.outbound_list, name='outbound_list'),
    path('outbound/add/', views.add_outbound_photos, name='add_outbound_photos'),
    path('outbound/<int:pk>/', views.outbound_detail, name='outbound_detail'),
    path('outbound/quick-add/<int:pk>/', views.upload_more_outbound_photos, name='upload_more_outbound_photos'),
    path('outbound/delete/<int:pk>/', views.delete_outbound_load, name='delete_outbound_load'),
    path('outbound/photo/delete/<int:photo_id>/', views.delete_outbound_photo, name='delete_outbound_photo'),

    # --- CONTAINER PHOTOS ---
    path('containers/', views.container_list, name='container_list'),
    path('containers/add/', views.add_container, name='add_container'),
    path('containers/<int:pk>/', views.container_detail, name='container_detail'),
    path('container/quick-add/<int:pk>/', views.upload_more_container_photos, name='upload_more_container_photos'),
    path('containers/delete/<int:pk>/', views.delete_container, name='delete_container'),
    path('containers/photo/delete/<int:photo_id>/', views.delete_container_photo, name='delete_container_photo'),

    # --- PICKUP PHOTOS ---
    path('pickup-photos/', views.pickup_photo_list, name='pickup_photo_list'),
    path('pickup-photos/add/', views.add_pickup_photos, name='add_pickup_photos'),
    path('pickup-photos/<int:pk>/', views.pickup_photo_detail, name='pickup_photo_detail'),
    path('pickup-photos/quick-add/<int:pk>/', views.upload_more_pickup_photos, name='upload_more_pickup_photos'),
    path('pickup-photos/delete-log/<int:pk>/', views.delete_pickup_photo_log, name='delete_pickup_photo_log'),
    path('pickup-photos/delete-photo/<int:photo_id>/', views.delete_pickup_individual_photo, name='delete_pickup_individual_photo'),
]