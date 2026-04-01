from django.urls import path
from .views import GoogleLoginView, GoogleCallbackView, LogoutView, MeView, PagesView, PageDetailView

urlpatterns = [
    path('login/',    GoogleLoginView.as_view(),    name='google_login'),
    path('callback/', GoogleCallbackView.as_view(), name='google_callback'),
    path('logout/',   LogoutView.as_view(),         name='logout'),
    path('me/',       MeView.as_view(),             name='me'),
    path('pages/',    PagesView.as_view(),           name='pages'),
    path('pages/<int:pk>/', PageDetailView.as_view(), name='page_detail'),
]
