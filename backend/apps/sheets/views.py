import traceback
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework.permissions import IsAuthenticated

from apps.sheets.services import (
    list_spreadsheets, list_tabs,
    read_sheet_state, push_rows,
)


class SpreadsheetListView(APIView):
    permission_classes = [IsAuthenticated]

    def get(self, request):
        try:
            return Response(list_spreadsheets(request.user))
        except Exception as e:
            return Response({'error': str(e)}, status=400)


class TabListView(APIView):
    permission_classes = [IsAuthenticated]

    def get(self, request, spreadsheet_id):
        try:
            return Response(list_tabs(request.user, spreadsheet_id))
        except Exception as e:
            return Response({'error': str(e)}, status=400)


class SheetStateView(APIView):
    permission_classes = [IsAuthenticated]

    def get(self, request, spreadsheet_id):
        tab = request.query_params.get('tab')
        if not tab:
            return Response({'error': 'tab query param is required.'}, status=400)
        try:
            return Response(read_sheet_state(request.user, spreadsheet_id, tab))
        except Exception as e:
            return Response({'error': str(e)}, status=400)


class PushRowsView(APIView):
    permission_classes = [IsAuthenticated]

    def post(self, request, spreadsheet_id):
        print('=== PUSH REQUEST ===')
        print('sheet_id:   ', request.data.get('sheet_id'))
        print('sheet_title:', request.data.get('sheet_title'))
        print('page_name:  ', request.data.get('page_name'))
        print('rows count: ', len(request.data.get('rows', [])))
        print('====================')

        sheet_id    = request.data.get('sheet_id')
        sheet_title = request.data.get('sheet_title', '').strip()
        page_name   = request.data.get('page_name', '').strip()
        rows        = request.data.get('rows', [])

        if not sheet_title:
            return Response({'error': 'sheet_title is required.'}, status=400)
        if not page_name:
            return Response({'error': 'page_name is required.'}, status=400)
        if sheet_id is None:
            return Response({'error': 'sheet_id (numeric gid) is required.'}, status=400)
        if not rows:
            return Response({'error': 'No rows provided.'}, status=400)

        try:
            result = push_rows(
                request.user,
                spreadsheet_id,
                int(sheet_id),
                sheet_title,
                page_name,
                rows,
            )
            return Response(result)
        except Exception as e:
            print('=== PUSH ERROR ===')
            traceback.print_exc()
            print('==================')
            return Response({'error': str(e)}, status=400)