import json
import requests
from datetime import datetime, timezone

from django.conf import settings
from django.contrib.auth.models import User

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request as GoogleRequest

from .models import GoogleCredential


def build_oauth_flow() -> Flow:
    client_config = {
        "web": {
            "client_id":                  settings.GOOGLE_CLIENT_ID,
            "client_secret":              settings.GOOGLE_CLIENT_SECRET,
            "redirect_uris":              [settings.GOOGLE_REDIRECT_URI],
            "auth_uri":                   "https://accounts.google.com/o/oauth2/auth",
            "token_uri":                  "https://oauth2.googleapis.com/token",
        }
    }
    flow = Flow.from_client_config(
        client_config,
        scopes=settings.GOOGLE_SCOPES,
        redirect_uri=settings.GOOGLE_REDIRECT_URI,
    )
    return flow


def get_authorization_url() -> tuple[str, str]:
    """Return (auth_url, state)."""
    flow = build_oauth_flow()
    auth_url, state = flow.authorization_url(
        access_type='offline',
        include_granted_scopes='true',
        prompt='consent',
    )
    return auth_url, state


def exchange_code(code: str, state: str) -> dict:
    """Exchange auth code for tokens and fetch user info."""
    flow = build_oauth_flow()
    flow.fetch_token(code=code)
    creds = flow.credentials

    # Fetch Google profile
    resp = requests.get(
        'https://www.googleapis.com/oauth2/v2/userinfo',
        headers={'Authorization': f'Bearer {creds.token}'},
        timeout=10,
    )
    resp.raise_for_status()
    profile = resp.json()

    return {
        'access_token':  creds.token,
        'refresh_token': creds.refresh_token,
        'token_expiry':  creds.expiry,
        'google_id':     profile.get('id', ''),
        'email':         profile.get('email', ''),
        'name':          profile.get('name', ''),
        'avatar_url':    profile.get('picture', ''),
    }


def get_or_create_user(token_data: dict) -> User:
    """Create or update a Django user from Google profile data."""
    email = token_data['email']
    name  = token_data.get('name', '')

    user, created = User.objects.get_or_create(
        username=email,
        defaults={
            'email':      email,
            'first_name': name.split(' ')[0] if name else '',
            'last_name':  ' '.join(name.split(' ')[1:]) if ' ' in name else '',
        }
    )
    if not created and name:
        parts = name.split(' ', 1)
        user.first_name = parts[0]
        user.last_name  = parts[1] if len(parts) > 1 else ''
        user.save(update_fields=['first_name', 'last_name'])

    # Upsert credentials
    GoogleCredential.objects.update_or_create(
        user=user,
        defaults={
            'access_token':  token_data['access_token'],
            'refresh_token': token_data.get('refresh_token') or '',
            'token_expiry':  token_data.get('token_expiry'),
            'google_id':     token_data.get('google_id', ''),
            'avatar_url':    token_data.get('avatar_url', ''),
        }
    )
    return user


def get_valid_credentials(user: User) -> Credentials:
    """Return a valid (possibly refreshed) Credentials object for the user."""
    cred = user.google_credential

    credentials = Credentials(
        token=cred.access_token,
        refresh_token=cred.refresh_token,
        token_uri='https://oauth2.googleapis.com/token',
        client_id=settings.GOOGLE_CLIENT_ID,
        client_secret=settings.GOOGLE_CLIENT_SECRET,
        scopes=settings.GOOGLE_SCOPES,
    )

    if credentials.expired and credentials.refresh_token:
        credentials.refresh(GoogleRequest())
        cred.access_token = credentials.token
        cred.token_expiry = credentials.expiry
        cred.save(update_fields=['access_token', 'token_expiry'])

    return credentials
