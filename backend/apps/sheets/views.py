from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework.permissions import IsAuthenticated

from .services import list_spreadsheets, list_tabs, read_sheet_state, push_rows


class SpreadsheetListView(APIView):
    """List all Google Sheets files in the user's Drive."""
    permission_classes = [IsAuthenticated]

    def get(self, request):
        try:
            files = list_spreadsheets(request.user)
            return Response(files)
        except Exception as e:
            return Response({'error': str(e)}, status=400)


class TabListView(APIView):
    """List all tabs in a spreadsheet."""
    permission_classes = [IsAuthenticated]

    def get(self, request, spreadsheet_id):
        try:
            tabs = list_tabs(request.user, spreadsheet_id)
            return Response(tabs)
        except Exception as e:
            return Response({'error': str(e)}, status=400)


class SheetStateView(APIView):
    """Return the current section state of a sheet tab."""
    permission_classes = [IsAuthenticated]

    def get(self, request, spreadsheet_id):
        sheet_title = request.query_params.get('tab')
        if not sheet_title:
            return Response({'error': 'tab query param required'}, status=400)
        try:
            state = read_sheet_state(request.user, spreadsheet_id, sheet_title)
            return Response(state)
        except Exception as e:
            return Response({'error': str(e)}, status=400)


class PushRowsView(APIView):
    """Push parsed rows to the correct section for a page."""
    permission_classes = [IsAuthenticated]

    def post(self, request, spreadsheet_id):
        data         = request.data
        sheet_id     = data.get('sheet_id')      # numeric gid
        sheet_title  = data.get('sheet_title')   # tab name
        page_name    = data.get('page_name', '').strip()
        rows         = data.get('rows', [])

        if not sheet_title:
            return Response({'error': 'sheet_title is required'}, status=400)
        if not page_name:
            return Response({'error': 'page_name is required'}, status=400)
        if not rows:
            return Response({'error': 'No rows provided'}, status=400)
        if sheet_id is None:
            return Response({'error': 'sheet_id (gid) is required'}, status=400)

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
            return Response({'error': str(e)}, status=400)
