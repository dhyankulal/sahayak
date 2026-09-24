from django.shortcuts import render
from django.http import HttpResponse, JsonResponse

from .forms import VolunteerReportForm, ComplaintReportForm
from .services import (
    prepare_volunteer_report_data,
    prepare_complaint_report_data,
    create_volunteer_word_document,
    create_complaint_word_document,
    save_police_report_to_supabase,
    fetch_from_external_supabase
)

def create_report(request):
    doc_type = request.GET.get("type", "volunteer") # 'volunteer' or 'complaint'

    if request.method == "POST":
        form_type = request.POST.get("form_type", "volunteer")
        
        if form_type == "volunteer":
            form = VolunteerReportForm(request.POST)
            if form.is_valid():
                form.save() # save to SQLite
                data = form.cleaned_data
                report_data = prepare_volunteer_report_data(data)
                output = create_volunteer_word_document(report_data)

                try:
                    save_police_report_to_supabase(report_data)
                except Exception as e:
                    print(f"Supabase sync notice: {e}")

                case_no = data["case_number"]
                safe_ref = "".join(c if c.isalnum() else "_" for c in case_no)
                response = HttpResponse(
                    output.getvalue(),
                    content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                response["Content-Disposition"] = f'attachment; filename="Volunteer_Police_Incident_Report_{safe_ref}.docx"'
                return response

        elif form_type == "complaint":
            form = ComplaintReportForm(request.POST)
            if form.is_valid():
                data = form.cleaned_data
                report_data = prepare_complaint_report_data(data)
                output = create_complaint_word_document(report_data)

                try:
                    save_police_report_to_supabase(report_data)
                except Exception as e:
                    print(f"Supabase sync notice: {e}")

                subj = data.get("subject") or "Complaint"
                safe_ref = "".join(c if c.isalnum() else "_" for c in subj)
                response = HttpResponse(
                    output.getvalue(),
                    content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                response["Content-Disposition"] = f'attachment; filename="Sample_Police_Complaint_Report_{safe_ref}.docx"'
                return response
    
    volunteer_form = VolunteerReportForm()
    complaint_form = ComplaintReportForm()

    return render(
        request,
        "index.html",
        {
            "volunteer_form": volunteer_form,
            "complaint_form": complaint_form,
            "doc_type": doc_type
        }
    )

def fetch_external_supabase(request):
    if request.method == "POST":
        target_url = request.POST.get("supabase_url", "").strip()
        target_key = request.POST.get("supabase_key", "").strip()
        table_name = request.POST.get("table_name", "police_reports").strip() or "police_reports"

        if not target_url or not target_key:
            return JsonResponse({"status": "error", "message": "Please enter both Supabase URL and Key."}, status=400)

        try:
            records = fetch_from_external_supabase(target_url, target_key, table_name)
            return JsonResponse({"status": "success", "count": len(records), "data": records})
        except Exception as e:
            return JsonResponse({"status": "error", "message": f"Failed to fetch from external Supabase: {str(e)}"}, status=500)

    return JsonResponse({"status": "error", "message": "Only POST request allowed."}, status=405)