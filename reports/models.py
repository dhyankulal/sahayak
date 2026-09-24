from django.db import models

class PoliceReport(models.Model):
    case_number = models.CharField(max_length=100)
    type_of_case = models.CharField(max_length=150)
    age_of_person = models.IntegerField(blank=True, null=True)
    police_name = models.CharField(max_length=150)
    address = models.CharField(max_length=300)
    volunteer_name = models.CharField(max_length=150, blank=True, null=True)
    phone_number = models.CharField(max_length=50)
    type_of_help_needed = models.TextField()
    actions_taken = models.TextField()
    date_of_report = models.DateField(auto_now_add=True)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"Case {self.case_number} - {self.type_of_case}"
