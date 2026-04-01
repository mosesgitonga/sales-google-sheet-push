from django.urls import path
from .views import SpreadsheetListView, TabListView, SheetStateView, PushRowsView

urlpatterns = [
    path('spreadsheets/',                              SpreadsheetListView.as_view(), name='spreadsheets'),
    path('spreadsheets/<str:spreadsheet_id>/tabs/',    TabListView.as_view(),         name='tabs'),
    path('spreadsheets/<str:spreadsheet_id>/state/',   SheetStateView.as_view(),      name='sheet_state'),
    path('spreadsheets/<str:spreadsheet_id>/push/',    PushRowsView.as_view(),        name='push_rows'),
]
