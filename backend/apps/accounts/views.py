from django.contrib.auth import login, logout
from django.http import JsonResponse
from django.views import View
from django.shortcuts import redirect
from django.utils.decorators import method_decorator
from django.views.decorators.csrf import csrf_exempt

from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework.permissions import IsAuthenticated, AllowAny

from .services import get_authorization_url, exchange_code, get_or_create_user
from .models import UserPage


class GoogleLoginView(View):
    """Redirect user to Google OAuth consent screen."""
    def get(self, request):
        auth_url, state = get_authorization_url()
        request.session['oauth_state'] = state
        return redirect(auth_url)


class GoogleCallbackView(View):
    """Handle Google OAuth callback, create/login user, redirect to app."""
    def get(self, request):
        code  = request.GET.get('code')
        state = request.GET.get('state')

        if not code:
            return redirect('/?error=no_code')

        try:
            token_data = exchange_code(code, state)
            user       = get_or_create_user(token_data)
            login(request, user, backend='django.contrib.auth.backends.ModelBackend')
            return redirect('/')
        except Exception as e:
            return redirect(f'/?error=auth_failed&detail={str(e)[:100]}')


@method_decorator(csrf_exempt, name='dispatch')
class LogoutView(View):
    def post(self, request):
        logout(request)
        return JsonResponse({'status': 'ok'})


class MeView(APIView):
    permission_classes = [IsAuthenticated]

    def get(self, request):
        user = request.user
        cred = getattr(user, 'google_credential', None)
        return Response({
            'id':         user.id,
            'email':      user.email,
            'name':       user.get_full_name() or user.email,
            'avatar_url': cred.avatar_url if cred else '',
        })


class PagesView(APIView):
    """List and create user pages (creator names)."""
    permission_classes = [IsAuthenticated]

    def get(self, request):
        pages = request.user.pages.values('id', 'name', 'created_at')
        return Response(list(pages))

    def post(self, request):
        name = request.data.get('name', '').strip()
        if not name:
            return Response({'error': 'Name is required.'}, status=400)

        page, created = UserPage.objects.get_or_create(user=request.user, name=name)
        if not created:
            return Response({'error': f'Page "{name}" already exists.'}, status=409)

        return Response({'id': page.id, 'name': page.name}, status=201)


class PageDetailView(APIView):
    permission_classes = [IsAuthenticated]

    def delete(self, request, pk):
        try:
            page = request.user.pages.get(pk=pk)
            page.delete()
            return Response({'status': 'deleted'})
        except UserPage.DoesNotExist:
            return Response({'error': 'Not found.'}, status=404)
