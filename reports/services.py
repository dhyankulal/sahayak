import os
import json
import urllib.request
import urllib.error
from io import BytesIO
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn, nsmap
from supabase import create_client
from dotenv import load_dotenv

nsmap['v'] = 'urn:schemas-microsoft-com:vml'
nsmap['o'] = 'urn:schemas-microsoft-com:office:office'

load_dotenv()

SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
SUPABASE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
WATERMARK_PATH = os.path.join(ASSETS_DIR, "watermark_fade.png")
LOGO_PATH = os.path.join(ASSETS_DIR, "karnataka_logo.png")
VOLUNTEER_BADGE_PATH = os.path.join(ASSETS_DIR, "volunteer_badge.png")

# Helper XML styling functions for Word tables & borders
def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=120, bottom=120, left=160, right=160):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def add_box_border(paragraph_or_cell, color_hex="1E3A8A", sz="12"):
    pPr = paragraph_or_cell._p.get_or_add_pPr() if hasattr(paragraph_or_cell, '_p') else paragraph_or_cell._tc.get_or_add_tcPr()
    bdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:top w:val="single" w:sz="{sz}" w:space="8" w:color="{color_hex}"/><w:left w:val="single" w:sz="{sz}" w:space="8" w:color="{color_hex}"/><w:bottom w:val="single" w:sz="{sz}" w:space="8" w:color="{color_hex}"/><w:right w:val="single" w:sz="{sz}" w:space="8" w:color="{color_hex}"/></w:pBdr>')
    pPr.append(bdr)

def prepare_volunteer_report_data(form_data):
    return {
        "case_number": form_data.get("case_number") or "N/A",
        "type_of_case": form_data.get("type_of_case") or "N/A",
        "age_of_person": form_data.get("age_of_person") or "N/A",
        "police_name": form_data.get("police_name") or "N/A",
        "address": form_data.get("address") or "N/A",
        "volunteer_name": form_data.get("volunteer_name") or "N/A",
        "phone_number": form_data.get("phone_number") or "N/A",
        "type_of_help_needed": form_data.get("type_of_help_needed") or "",
        "actions_taken": form_data.get("actions_taken") or "",
    }

def prepare_complaint_report_data(form_data):
    return {
        "police_station": form_data.get("police_station") or "______________________________",
        "district": form_data.get("district") or "______________________________",
        "state": form_data.get("state") or "Karnataka",
        "date_of_complaint": form_data.get("date_of_complaint") or "____ / ____ / ______",
        "time_of_complaint": form_data.get("time_of_complaint") or "__________",
        "complainant_name": form_data.get("complainant_name") or "______________________________",
        "complainant_age": form_data.get("complainant_age") or "__________",
        "complainant_gender": form_data.get("complainant_gender") or "__________________",
        "complainant_address": form_data.get("complainant_address") or "______________________________",
        "complainant_mobile": form_data.get("complainant_mobile") or "________________________",
        "subject": form_data.get("subject") or "______________________________",
        "date_of_incident": form_data.get("date_of_incident") or "____ / ____ / ______",
        "approx_time": form_data.get("approx_time") or "__________",
        "place_of_incident": form_data.get("place_of_incident") or "______________________________",
        "location_at_time": form_data.get("location_at_time") or "______________________________",
        "time_of_occurrence": form_data.get("time_of_occurrence") or "__________",
        "incident_description": form_data.get("incident_description") or "",
        "accused_name": form_data.get("accused_name") or "______________________________",
        "accused_address": form_data.get("accused_address") or "____________________________",
        "accused_description": form_data.get("accused_description") or "_____________________",
        "witness1_name": form_data.get("witness1_name") or "______________________________",
        "witness1_contact": form_data.get("witness1_contact") or "______________________________",
        "witness2_name": form_data.get("witness2_name") or "______________________________",
        "witness2_contact": form_data.get("witness2_contact") or "______________________________",
        "evidence_submitted": form_data.get("evidence_submitted") or "",
        "declaration_place": form_data.get("declaration_place") or "______________________________",
        "declaration_date": form_data.get("declaration_date") or "____ / ____ / ______",
        "declaration_name": form_data.get("declaration_name") or "______________________________",
        "police_received_by": form_data.get("police_received_by") or "______________________________",
        "police_rank": form_data.get("police_rank") or "______________________________",
        "police_date_time": form_data.get("police_date_time") or "______________________________",
        "police_diary_no": form_data.get("police_diary_no") or "________________________",
        "police_action_taken": form_data.get("police_action_taken") or "",
    }

def add_bottom_border_to_paragraph(paragraph, color="000000", sz="18"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), sz)
    bottom.set(qn('w:space'), '4')
    bottom.set(qn('w:color'), color)
    pBdr.append(bottom)
    pPr.append(pBdr)

def attach_header_watermark_and_logo(doc):
    header = doc.sections[0].header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Watermark background image
    if os.path.exists(WATERMARK_PATH):
        rId, _ = header.part.get_or_add_image(WATERMARK_PATH)
        vml_xml = f'''<w:r {nsdecls("w")}>
            <w:rPr><w:noProof/></w:rPr>
            <w:pict {nsdecls("v", "o")}>
                <v:shape id="WaterMark" style="position:absolute;margin-left:0;margin-top:0;width:450pt;height:450pt;z-index:-251657728;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" coordsize="21600,21600" fillcolor="none" stroked="f">
                    <v:imagedata r:id="{rId}" {nsdecls("r")}/>
                </v:shape>
            </w:pict>
        </w:r>'''
        hp._p.append(parse_xml(vml_xml))

    # Official Logo in header
    if os.path.exists(LOGO_PATH):
        r_logo = hp.add_run()
        r_logo.add_picture(LOGO_PATH, width=Inches(0.95))

# ==========================================
# 2. BEAUTIFUL VOLUNTEER CERTIFICATE & INCIDENT REPORT GENERATOR
# ==========================================

def create_volunteer_word_document(report_data):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    attach_header_watermark_and_logo(doc)

    # Main Title: POLICE INCIDENT REPORT
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    r_title = p_title.add_run("POLICE INCIDENT REPORT")
    r_title.bold = True
    r_title.font.name = "Georgia"
    r_title.font.size = Pt(20)
    r_title.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A) # Slate Dark Blue
    
    add_bottom_border_to_paragraph(p_title, color="1E3A8A", sz="14")

    def add_section_header(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(title)
        r.bold = True
        r.font.name = "Georgia"
        r.font.size = Pt(11)
        r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    def add_field_line(label, value, label2=None, value2=None):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.05
        
        r_lbl1 = p.add_run(f"{label}: ")
        r_lbl1.bold = True
        r_lbl1.font.name = "Calibri"
        r_lbl1.font.size = Pt(9.5)

        r_val1 = p.add_run(f"{value}")
        r_val1.font.name = "Calibri"
        r_val1.font.size = Pt(9.5)

        if label2:
            p.add_run("\t\t")
            r_lbl2 = p.add_run(f"{label2}: ")
            r_lbl2.bold = True
            r_lbl2.font.name = "Calibri"
            r_lbl2.font.size = Pt(9.5)

            r_val2 = p.add_run(f"{value2}")
            r_val2.font.name = "Calibri"
            r_val2.font.size = Pt(9.5)

    add_section_header("Person(s) Involved")
    add_field_line("Case Number", report_data["case_number"], "Age of Person", report_data["age_of_person"])
    add_field_line("Contact Phone", report_data["phone_number"], "Address", report_data["address"])
    add_field_line("Police Officer", report_data["police_name"], "Volunteer Name", report_data["volunteer_name"])

    add_section_header("Incident Details")
    add_field_line("Type of Case / Incident", report_data["type_of_case"])

    p_desc_lbl = doc.add_paragraph()
    p_desc_lbl.paragraph_format.space_before = Pt(2)
    p_desc_lbl.paragraph_format.space_after = Pt(1)
    r_desc = p_desc_lbl.add_run("What Type of Help Needed / Details of Event:")
    r_desc.bold = True
    r_desc.font.name = "Calibri"
    r_desc.font.size = Pt(9.5)

    p_desc_val = doc.add_paragraph()
    p_desc_val.paragraph_format.space_after = Pt(4)
    r_desc_val = p_desc_val.add_run(report_data["type_of_help_needed"])
    r_desc_val.font.name = "Calibri"
    r_desc_val.font.size = Pt(9.5)

    add_section_header("Actions Taken")
    p_act_val = doc.add_paragraph()
    p_act_val.paragraph_format.space_after = Pt(8)
    r_act_val = p_act_val.add_run(report_data["actions_taken"])
    r_act_val.font.name = "Calibri"
    r_act_val.font.size = Pt(9.5)

    # ====================================================
    # BEAUTIFUL ELEGANT VOLUNTEER CERTIFICATE BOX
    # ====================================================
    cert_table = doc.add_table(rows=1, cols=1)
    cert_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cert_cell = cert_table.cell(0, 0)
    
    # 1. Background shading & double border styling
    set_cell_background(cert_cell, "F8FAFC")
    set_cell_margins(cert_cell, top=140, bottom=140, left=180, right=180)
    
    # Double golden/navy border for certificate feel
    tcPr = cert_cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="double" w:sz="18" w:space="0" w:color="1E3A8A"/><w:left w:val="double" w:sz="18" w:space="0" w:color="1E3A8A"/><w:bottom w:val="double" w:sz="18" w:space="0" w:color="1E3A8A"/><w:right w:val="double" w:sz="18" w:space="0" w:color="1E3A8A"/></w:tcBorders>')
    tcPr.append(borders)

    # 2. Certificate Header inside the box
    p_cert_title = cert_cell.paragraphs[0]
    p_cert_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cert_title.paragraph_format.space_after = Pt(2)
    r_cert_title = p_cert_title.add_run("CERTIFICATE OF VOLUNTEER SERVICE")
    r_cert_title.bold = True
    r_cert_title.font.name = "Georgia"
    r_cert_title.font.size = Pt(14)
    r_cert_title.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    p_cert_sub = cert_cell.add_paragraph()
    p_cert_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cert_sub.paragraph_format.space_after = Pt(6)
    r_cert_sub = p_cert_sub.add_run("AWARDED IN RECOGNITION OF COMMUNITY ASSISTANCE")
    r_cert_sub.bold = True
    r_cert_sub.font.name = "Calibri"
    r_cert_sub.font.size = Pt(8.5)
    r_cert_sub.font.color.rgb = RGBColor(0xD9, 0x77, 0x06)

    # 3. Volunteer Image Badge (if available)
    if os.path.exists(VOLUNTEER_BADGE_PATH):
        p_img = cert_cell.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_after = Pt(6)
        r_img = p_img.add_run()
        r_img.add_picture(VOLUNTEER_BADGE_PATH, width=Inches(1.2))

    # 4. Certificate Wording Text
    p_cert_body = cert_cell.add_paragraph()
    p_cert_body.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cert_body.paragraph_format.space_after = Pt(8)
    p_cert_body.paragraph_format.line_spacing = 1.15

    r_body1 = p_cert_body.add_run("This is to proudly certify that volunteer ")
    r_body1.font.name = "Calibri"
    r_body1.font.size = Pt(9.5)

    r_vname = p_cert_body.add_run(f" {report_data['volunteer_name']} ")
    r_vname.bold = True
    r_vname.font.name = "Georgia"
    r_vname.font.size = Pt(11)
    r_vname.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    r_vname.underline = True

    r_body2 = p_cert_body.add_run(" has rendered invaluable voluntary service and support to the Police Department under Officer ")
    r_body2.font.name = "Calibri"
    r_body2.font.size = Pt(9.5)

    r_pname = p_cert_body.add_run(f" {report_data['police_name']} ")
    r_pname.bold = True
    r_pname.font.name = "Calibri"
    r_pname.font.size = Pt(9.5)

    r_body3 = p_cert_body.add_run(" for Case No. ")
    r_body3.font.name = "Calibri"
    r_body3.font.size = Pt(9.5)

    r_cno = p_cert_body.add_run(f"{report_data['case_number']}")
    r_cno.bold = True
    r_cno.font.name = "Calibri"
    r_cno.font.size = Pt(9.5)

    r_body4 = p_cert_body.add_run(f" ({report_data['type_of_case']}). Their dedication to community welfare and active assistance is hereby commended.")
    r_body4.font.name = "Calibri"
    r_body4.font.size = Pt(9.5)

    # 5. Certificate Signatures Block
    p_cert_sig = cert_cell.add_paragraph()
    p_cert_sig.paragraph_format.space_before = Pt(6)
    p_cert_sig.paragraph_format.space_after = Pt(2)
    
    r_sig1 = p_cert_sig.add_run("_______________________________\t\t\t_______________________________")
    r_sig1.font.name = "Calibri"
    r_sig1.font.size = Pt(9)
    
    p_sig_lbl = cert_cell.add_paragraph()
    p_sig_lbl.paragraph_format.space_after = Pt(2)
    r_lbl_p = p_sig_lbl.add_run(f"Signature of Police Officer\t\t\tSignature of Issuing Authority\n({report_data['police_name']})\t\t\t(Police Department Seal)")
    r_lbl_p.font.name = "Calibri"
    r_lbl_p.font.size = Pt(8.5)
    r_lbl_p.bold = True

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# ==========================================
# 3. SAMPLE POLICE COMPLAINT REPORT GENERATOR
# ==========================================

def create_complaint_word_document(report_data):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.65)
        section.right_margin = Inches(0.65)

    attach_header_watermark_and_logo(doc)

    # Document Header Titles
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(2)
    r_title = p_title.add_run("SAMPLE POLICE COMPLAINT REPORT")
    r_title.bold = True
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(18)
    r_title.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(10)
    r_sub = p_sub.add_run("FOR PRACTICE / DEMONSTRATION ONLY — NOT AN OFFICIAL POLICE DOCUMENT")
    r_sub.bold = True
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(9.5)
    r_sub.font.color.rgb = RGBColor(0xDC, 0x26, 0x26)

    # General Station & Date Information
    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_after = Pt(8)
    p_info.paragraph_format.line_spacing = 1.2
    
    def append_field(p, label, value, end_space="\t"):
        r_lbl = p.add_run(f"{label}: ")
        r_lbl.bold = True
        r_lbl.font.name = "Calibri"
        r_lbl.font.size = Pt(10.5)
        r_lbl.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)
        
        r_val = p.add_run(f"{value}{end_space}")
        r_val.font.name = "Calibri"
        r_val.font.size = Pt(10.5)

    append_field(p_info, "Police Station", report_data["police_station"], "\t\t")
    append_field(p_info, "District", report_data["district"], "\n")
    append_field(p_info, "State", report_data["state"], "\t\t\t")
    append_field(p_info, "Date of Complaint", report_data["date_of_complaint"], "\t\t")
    append_field(p_info, "Time", report_data["time_of_complaint"], "")

    def add_section_header(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(title)
        r.bold = True
        r.font.name = "Calibri"
        r.font.size = Pt(11.5)
        r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
        add_bottom_border_to_paragraph(p, color="1E3A8A", sz="12")

    # 1. Complainant Details
    add_section_header("1. Complainant Details")
    p_c = doc.add_paragraph()
    p_c.paragraph_format.space_after = Pt(6)
    p_c.paragraph_format.line_spacing = 1.2
    append_field(p_c, "Name", report_data["complainant_name"], "\n")
    append_field(p_c, "Age", report_data["complainant_age"], "\t\t\t")
    append_field(p_c, "Gender", report_data["complainant_gender"], "\n")
    append_field(p_c, "Address", report_data["complainant_address"], "\n")
    append_field(p_c, "Mobile Number", report_data["complainant_mobile"], "")

    # 2. Details of the Complaint
    add_section_header("2. Details of the Complaint")
    p_d = doc.add_paragraph()
    p_d.paragraph_format.space_after = Pt(6)
    p_d.paragraph_format.line_spacing = 1.2
    append_field(p_d, "Subject", f"Complaint regarding {report_data['subject']}", "\n")
    
    p_stmt = doc.add_paragraph()
    p_stmt.paragraph_format.space_after = Pt(6)
    p_stmt.paragraph_format.line_spacing = 1.2
    r_stmt = p_stmt.add_run(f"I, {report_data['complainant_name']}, residing at the above-mentioned address, hereby submit this complaint regarding the following incident:")
    r_stmt.font.name = "Calibri"
    r_stmt.font.size = Pt(10.5)

    p_d_meta = doc.add_paragraph()
    p_d_meta.paragraph_format.space_after = Pt(6)
    p_d_meta.paragraph_format.line_spacing = 1.2
    append_field(p_d_meta, "Date of Incident", report_data["date_of_incident"], "\t\t")
    append_field(p_d_meta, "Approximate Time", report_data["approx_time"], "\n")
    append_field(p_d_meta, "Place of Incident", report_data["place_of_incident"], "")

    # 3. Description of Incident
    add_section_header("3. Description of Incident")
    p_desc = doc.add_paragraph()
    p_desc.paragraph_format.space_after = Pt(4)
    p_desc.paragraph_format.line_spacing = 1.2
    r_desc_intro = p_desc.add_run(f"On the above-mentioned date and time, I was at {report_data['location_at_time']}. At approximately {report_data['time_of_occurrence']}, the following incident occurred:")
    r_desc_intro.font.name = "Calibri"
    r_desc_intro.font.size = Pt(10.5)

    p_desc_body = doc.add_paragraph()
    p_desc_body.paragraph_format.space_after = Pt(6)
    p_desc_body.paragraph_format.line_spacing = 1.15
    r_desc_text = p_desc_body.add_run(report_data["incident_description"] or "________________________________________________________________________________________________________________________________________________________________________________________________________________________")
    r_desc_text.font.name = "Calibri"
    r_desc_text.font.size = Pt(10.5)

    p_req = doc.add_paragraph()
    p_req.paragraph_format.space_after = Pt(6)
    p_req.paragraph_format.line_spacing = 1.2
    r_req = p_req.add_run("I request the concerned police authorities to kindly take note of my complaint and take appropriate action in accordance with the applicable law.")
    r_req.font.name = "Calibri"
    r_req.font.size = Pt(10.5)
    r_req.italic = True

    # 4. Accused / Suspect Details
    add_section_header("4. Accused / Suspect Details")
    p_ac = doc.add_paragraph()
    p_ac.paragraph_format.space_after = Pt(6)
    p_ac.paragraph_format.line_spacing = 1.2
    append_field(p_ac, "Name (if known)", report_data["accused_name"], "\n")
    append_field(p_ac, "Address (if known)", report_data["accused_address"], "\n")
    append_field(p_ac, "Description / Identification", report_data["accused_description"], "")

    # 5. Witness Details
    add_section_header("5. Witness Details")
    p_w = doc.add_paragraph()
    p_w.paragraph_format.space_after = Pt(6)
    p_w.paragraph_format.line_spacing = 1.2
    append_field(p_w, "Witness 1", report_data["witness1_name"], "\t\t")
    append_field(p_w, "Contact", report_data["witness1_contact"], "\n")
    append_field(p_w, "Witness 2", report_data["witness2_name"], "\t\t")
    append_field(p_w, "Contact", report_data["witness2_contact"], "")

    # 6. Evidence / Documents Submitted
    add_section_header("6. Evidence / Documents Submitted")
    p_ev = doc.add_paragraph()
    p_ev.paragraph_format.space_after = Pt(6)
    p_ev.paragraph_format.line_spacing = 1.2
    r_ev = p_ev.add_run(report_data["evidence_submitted"] or "1. _________________________________________\n2. _________________________________________")
    r_ev.font.name = "Calibri"
    r_ev.font.size = Pt(10.5)

    # 7. Complainant's Declaration
    add_section_header("7. Complainant's Declaration")
    p_decl = doc.add_paragraph()
    p_decl.paragraph_format.space_after = Pt(8)
    p_decl.paragraph_format.line_spacing = 1.2
    r_decl = p_decl.add_run("I declare that the information provided above is true and correct to the best of my knowledge and belief.")
    r_decl.font.name = "Calibri"
    r_decl.font.size = Pt(10.5)

    p_sig = doc.add_paragraph()
    p_sig.paragraph_format.space_after = Pt(8)
    p_sig.paragraph_format.line_spacing = 1.2
    append_field(p_sig, "Place", report_data["declaration_place"], "\t\t\t\t\t")
    append_field(p_sig, "Signature of Complainant", "", "\n")
    append_field(p_sig, "Date", report_data["declaration_date"], "\t\t\t\t\t")
    append_field(p_sig, "Name", report_data["declaration_name"], "")

    # For Police Station Use Only
    add_section_header("For Police Station Use Only")
    p_pol = doc.add_paragraph()
    p_pol.paragraph_format.space_after = Pt(6)
    p_pol.paragraph_format.line_spacing = 1.2
    append_field(p_pol, "Complaint Received By", report_data["police_received_by"], "\n")
    append_field(p_pol, "Rank / Designation", report_data["police_rank"], "\n")
    append_field(p_pol, "Date & Time Received", report_data["police_date_time"], "\n")
    append_field(p_pol, "Complaint / CSR / Diary No.", report_data["police_diary_no"], "\n")
    append_field(p_pol, "Action Taken / Remarks", report_data["police_action_taken"] or "__________________________________________________", "\n\n")
    append_field(p_pol, "Signature of Receiving Officer", "", "\t\t\t")
    append_field(p_pol, "Police Station Seal", "", "")

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

create_word_document = create_volunteer_word_document

def save_police_report_to_supabase(report_data):
    url = f"{SUPABASE_URL.rstrip('/')}/rest/v1/police_reports"
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
        "Prefer": "return=minimal"
    }
    
    case_no = report_data.get("case_number") or report_data.get("police_diary_no") or "N/A"
    case_type = report_data.get("type_of_case") or report_data.get("subject") or "N/A"
    p_name = report_data.get("police_name") or report_data.get("police_received_by") or "N/A"
    addr = report_data.get("address") or report_data.get("complainant_address") or "N/A"
    v_name = report_data.get("volunteer_name") or report_data.get("complainant_name") or "N/A"
    phone = report_data.get("phone_number") or report_data.get("complainant_mobile") or "N/A"
    help_needed = report_data.get("type_of_help_needed") or report_data.get("incident_description") or ""
    actions = report_data.get("actions_taken") or report_data.get("police_action_taken") or ""

    payload = {
        "case_number": str(case_no),
        "type_of_case": str(case_type),
        "police_name": str(p_name),
        "address": str(addr),
        "volunteer_name": str(v_name),
        "phone_number": str(phone),
        "type_of_help_needed": str(help_needed),
        "actions_taken": str(actions),
    }

    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers=headers,
        method="POST"
    )

    try:
        with urllib.request.urlopen(req) as resp:
            return resp.status
    except Exception as e:
        fallback_url = f"{SUPABASE_URL.rstrip('/')}/rest/v1/report_documents"
        fallback_payload = {
            "case_reference": str(case_no),
            "title": f"Report - {case_type}",
            "storage_path": f"Volunteer/Person:{v_name} | Phone:{phone}"
        }
        fb_req = urllib.request.Request(
            fallback_url,
            data=json.dumps(fallback_payload).encode("utf-8"),
            headers=headers,
            method="POST"
        )
        with urllib.request.urlopen(fb_req) as fb_resp:
            return fb_resp.status

def fetch_from_external_supabase(target_url, target_key, table_name="police_reports", select="*"):
    cleaned_url = target_url.strip().rstrip('/')
    url = f"{cleaned_url}/rest/v1/{table_name}?select={select}"
    
    headers = {
        "apikey": target_key.strip(),
        "Authorization": f"Bearer {target_key.strip()}",
        "Content-Type": "application/json"
    }

    req = urllib.request.Request(url, headers=headers, method="GET")
    with urllib.request.urlopen(req) as resp:
        content = resp.read().decode("utf-8")
        return json.loads(content)