from django import forms
from .models import PoliceReport

# 1. Original Volunteer / Incident Report Form
class VolunteerReportForm(forms.ModelForm):
    class Meta:
        model = PoliceReport
        fields = [
            "case_number",
            "type_of_case",
            "age_of_person",
            "police_name",
            "address",
            "volunteer_name",
            "phone_number",
            "type_of_help_needed",
            "actions_taken",
        ]
        labels = {
            "case_number": "Case Number",
            "type_of_case": "Type of Case / Incident",
            "age_of_person": "Age of the Person",
            "police_name": "Police Officer Name",
            "address": "Address",
            "volunteer_name": "Volunteer Name",
            "phone_number": "Contact Phone Number",
            "type_of_help_needed": "What Type of Help Needed / Details of Event",
            "actions_taken": "Actions Taken",
        }
        widgets = {
            "case_number": forms.TextInput(attrs={"placeholder": "e.g. CAS-2026-089"}),
            "type_of_case": forms.TextInput(attrs={"placeholder": "e.g. Missing Person / Medical Emergency"}),
            "police_name": forms.TextInput(attrs={"placeholder": "e.g. Officer R. Sharma"}),
            "volunteer_name": forms.TextInput(attrs={"placeholder": "e.g. Anish Kumar"}),
            "type_of_help_needed": forms.Textarea(attrs={"rows": 4}),
            "actions_taken": forms.Textarea(attrs={"rows": 4}),
        }

# Backwards compatibility alias
CaseDetailsForm = VolunteerReportForm

# 2. Sample Police Complaint Report Form (FIR / Complaint format)
class ComplaintReportForm(forms.Form):
    police_station = forms.CharField(label="Police Station", max_length=150, required=False, widget=forms.TextInput(attrs={"placeholder": "e.g. Shirva Police Station"}))
    district = forms.CharField(label="District", max_length=100, required=False, widget=forms.TextInput(attrs={"placeholder": "e.g. Udupi"}))
    state = forms.CharField(label="State", max_length=100, initial="Karnataka", required=False)
    date_of_complaint = forms.CharField(label="Date of Complaint", max_length=50, required=False, widget=forms.TextInput(attrs={"placeholder": "DD / MM / YYYY"}))
    time_of_complaint = forms.CharField(label="Time of Complaint", max_length=50, required=False, widget=forms.TextInput(attrs={"placeholder": "e.g. 10:30 AM"}))

    complainant_name = forms.CharField(label="Complainant Name", max_length=150, required=False)
    complainant_age = forms.CharField(label="Age", max_length=20, required=False)
    complainant_gender = forms.CharField(label="Gender", max_length=30, required=False)
    complainant_address = forms.CharField(label="Address", max_length=300, required=False)
    complainant_mobile = forms.CharField(label="Mobile Number", max_length=50, required=False)

    subject = forms.CharField(label="Subject (Complaint regarding)", max_length=200, required=False)
    date_of_incident = forms.CharField(label="Date of Incident", max_length=50, required=False, widget=forms.TextInput(attrs={"placeholder": "DD / MM / YYYY"}))
    approx_time = forms.CharField(label="Approximate Time", max_length=50, required=False)
    place_of_incident = forms.CharField(label="Place of Incident", max_length=200, required=False)

    location_at_time = forms.CharField(label="Location at Time of Incident", max_length=200, required=False)
    time_of_occurrence = forms.CharField(label="Occurrence Time", max_length=50, required=False)
    incident_description = forms.CharField(label="Detailed Incident Description", widget=forms.Textarea(attrs={"rows": 4}), required=False)

    accused_name = forms.CharField(label="Accused Name (if known)", max_length=150, required=False)
    accused_address = forms.CharField(label="Accused Address (if known)", max_length=200, required=False)
    accused_description = forms.CharField(label="Identification / Description", max_length=200, required=False)

    witness1_name = forms.CharField(label="Witness 1 Name", max_length=150, required=False)
    witness1_contact = forms.CharField(label="Witness 1 Contact", max_length=50, required=False)
    witness2_name = forms.CharField(label="Witness 2 Name", max_length=150, required=False)
    witness2_contact = forms.CharField(label="Witness 2 Contact", max_length=50, required=False)

    evidence_submitted = forms.CharField(label="Evidence / Documents Submitted", widget=forms.Textarea(attrs={"rows": 2}), required=False)

    declaration_place = forms.CharField(label="Declaration Place", max_length=100, required=False)
    declaration_date = forms.CharField(label="Declaration Date", max_length=50, required=False)
    declaration_name = forms.CharField(label="Declaration Signatory Name", max_length=150, required=False)

    police_received_by = forms.CharField(label="Complaint Received By", max_length=150, required=False)
    police_rank = forms.CharField(label="Rank / Designation", max_length=100, required=False)
    police_date_time = forms.CharField(label="Date & Time Received", max_length=100, required=False)
    police_diary_no = forms.CharField(label="Complaint / CSR / Diary No.", max_length=100, required=False)
    police_action_taken = forms.CharField(label="Action Taken / Remarks", widget=forms.Textarea(attrs={"rows": 3}), required=False)