from django.db import models
from django.contrib.auth.models import User


class GoogleCredential(models.Model):
    """Stores Google OAuth tokens for a user."""
    user         = models.OneToOneField(User, on_delete=models.CASCADE, related_name='google_credential')
    access_token  = models.TextField()
    refresh_token = models.TextField(blank=True, null=True)
    token_expiry  = models.DateTimeField(blank=True, null=True)
    google_id     = models.CharField(max_length=128, blank=True)
    avatar_url    = models.URLField(blank=True)

    class Meta:
        verbose_name = 'Google Credential'

    def __str__(self):
        return f'{self.user.email} — Google credential'


class UserPage(models.Model):
    """Pages (creator names) a user has added to their account."""
    user       = models.ForeignKey(User, on_delete=models.CASCADE, related_name='pages')
    name       = models.CharField(max_length=200)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        unique_together = ('user', 'name')
        ordering = ['name']

    def __str__(self):
        return f'{self.user.email} — {self.name}'
