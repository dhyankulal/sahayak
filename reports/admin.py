from django.contrib import admin
from .models import PoliceReport

@admin.register(PoliceReport)
class PoliceReportAdmin(admin.ModelAdmin):
    list_display = ("case_number", "type_of_case", "police_name", "phone_number", "created_at")
    search_fields = ("case_number", "type_of_case", "police_name", "volunteer_name", "phone_number")
    list_filter = ("created_at",)
