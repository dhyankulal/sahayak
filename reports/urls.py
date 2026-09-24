from django.urls import path
from . import views

urlpatterns = [
    path("", views.create_report, name="create_report"),
    path("fetch-supabase/", views.fetch_external_supabase, name="fetch_external_supabase"),
]