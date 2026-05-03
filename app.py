"""
app.py — IIRER Unified Intake Chatbot
Projects: SWTCIE · Pathways · CIP
Output: Appends row to master Excel in Azure Blob Storage
Pipeline: Streamlit → Append to master Excel in Blob → Blob Event Trigger → ADF Data Flow → Power BI
"""
from __future__ import annotations
import html, re, io, os
from datetime import date, datetime
from typing import Any, Callable, Dict, List, Optional

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

load_dotenv()

st.set_page_config(page_title="IIRER Intake Chatbot", page_icon="💬", layout="centered", initial_sidebar_state="collapsed")

AZURE_CONNECTION_STRING = os.getenv("AZURE_STORAGE_CONNECTION_STRING", "")
AZURE_CONTAINER_NAME    = os.getenv("AZURE_BLOB_CONTAINER_NAME", "iirer-submissions")
MASTER_BLOB_NAME        = os.getenv("AZURE_MASTER_BLOB_NAME", "IIRER_Sample_Data_950.xlsx")  # single master file
MASTER_SHEET_NAME       = "All_Projects"

DARK_BLUE = "#13294B"; ORANGE = "#E84A27"; LIGHT_BG = "#F0F3F8"

# ── Option lists ───────────────────────────────────────────────────────────────
RACE_OPTIONS        = ["American Indian or Alaska Native","Asian","Black or African American","Hispanic or Latino","Native Hawaiian or Other Pacific Islander","White","Two or more races","Prefer not to say"]
GENDER_OPTIONS      = ["Woman","Man","Non-binary / gender non-conforming","Prefer not to say"]
DISABILITY_OPTIONS  = ["Physical / Mobility","Visual","Hearing","Cognitive / Intellectual","Psychiatric / Mental Health","Autism Spectrum","Traumatic Brain Injury","Other"]
EDUCATION_OPTIONS   = ["Less than high school","High school diploma / GED","Some college","Associate degree","Bachelor's degree","Graduate degree"]
AGENCIES_SWTCIE     = ["Centerstone","IMPACT CIL","Access Living","Disability Resource Center – Will Grundy","Lake County CIL","Other"]
CILS_PATHWAYS       = ["Access Living","Disability Resource Center – Will Grundy CIL","IMPACT CIL","Lake County CIL","PACE CIL","Progress CIL","SILC Illinois","Other"]
UNDERSERVED_CATS    = ["Living in poverty / high-poverty school","Student of color","English learner","Disconnected youth","Experiencing homelessness / housing insecurity","In foster care","Impacted by justice system","Migrant student","LGBTQI+ student","First-generation postsecondary student","Other"]
PATHWAYS_SERVICES   = ["Employment services","Education / curriculum services","Post-secondary education (PSE) planning","Career exploration / Transfr VR headset","Benefits planning","Family training / wellness session","Transition planning support","Budgeting / life skills workshop","Other"]
VR_SERVICES         = ["Counseling and Guidance","Information and Guidance","Job Placement","On-the-Job Training","Transition Services"]
PROF_ROLES          = ["Vocational Rehabilitation Counselor","VR Supervisor / Administrator","Stakeholder / Partner Agency Staff","Educator / Academic","Employer Representative","Other"]
TRAINING_FORMATS    = ["In-person seminar","Webinar","Conference presentation","Online / self-paced","Podcast","Other"]
US_STATES           = ["Alabama","Alaska","Arizona","Arkansas","California","Colorado","Connecticut","Delaware","Florida","Georgia","Hawaii","Idaho","Illinois","Indiana","Iowa","Kansas","Kentucky","Louisiana","Maine","Maryland","Massachusetts","Michigan","Minnesota","Mississippi","Missouri","Montana","Nebraska","Nevada","New Hampshire","New Jersey","New Mexico","New York","North Carolina","North Dakota","Ohio","Oklahoma","Oregon","Pennsylvania","Rhode Island","South Carolina","South Dakota","Tennessee","Texas","Utah","Vermont","Virginia","Washington","West Virginia","Wisconsin","Wyoming"]

# ── Master column order (must match IIRER_Sample_Data_950.xlsx exactly) ────────
MASTER_COLUMNS = [
    "submitted_at","project","first_name","last_name","participant_email","dob",
    "gender","race_ethnicity","disability_category","education","consent_signed",
    "consent_date","drs_application_date","submitting_agency","worked_14c",
    "currently_employed","employment_type","hours_subminimum","wage_subminimum",
    "hours_competitive","wage_competitive","hours_per_week","hourly_wage",
    "previous_subminimum","contemplating_subminimum","receives_benefits_sw",
    "benefits_planning_offered_sw","service_type","vr_outcome","obtained_cie",
    "cie_start_date","job_title","job_duties","cie_hours_per_week","cie_wage",
    "closure_date","closure_status","closure_smw_hours","closure_smw_wage",
    "closure_cie_hours","closure_cie_wage","cil","school_lea","age_group",
    "contact_method","iep_status","underserved","underserved_categories",
    "pathways_services","pathways_svc_effectiveness","pathways_obtained_cie",
    "pathways_cie_start","pathways_cie_title","pathways_cie_hours","pathways_cie_wage",
    "pse_status","pse_institution","pathways_closure_date","pathways_closure_status",
    "pathways_closure_pse_inst","organization","professional_role","state",
    "training_type","training_format","training_topic","training_date",
    "training_completed","prev_cip","quality_rating","relevance_rating",
    "usefulness_rating","knowledge_increase","confidence_increase",
    "skill_adoption_increase","qualitative_notes","receives_benefits_pw",
    "benefits_planning_offered_pw",
]

def is_swtcie(a):   return a.get("project") == "SWTCIE"
def is_pathways(a): return a.get("project") == "Pathways"
def is_cip(a):      return a.get("project") == "CIP"

ALL_STEPS: List[Dict[str, Any]] = [
    {"key":"project","q":"👋 Welcome to the **IIRER Intake Assistant**!\n\nI'll guide you through participant data collection step by step.\n\nFirst, which **project** is this submission for?","type":"dropdown","options":["SWTCIE","Pathways","CIP"],"required":True},
    # ── Personal Details (shared) ─────────────────────────────────────────────
    {"key":"first_name","q":"What is the participant's **first name**?","type":"text","required":True},
    {"key":"last_name","q":"What is the participant's **last name**?","type":"text","required":True},
    {"key":"participant_email","q":"What is the participant's **email address**?","type":"email","required":True,"validation":["email"]},
    {"key":"dob","q":"What is the participant's **date of birth**?","type":"date","required":True,"validation":["no_future"]},
    {"key":"gender","q":"What is the participant's **gender**?","type":"dropdown","options":GENDER_OPTIONS},
    {"key":"race_ethnicity","q":"What is the participant's **race / ethnicity**?\n\n*(Select all that apply.)*","type":"multiselect","options":RACE_OPTIONS},
    {"key":"disability_category","q":"What is the participant's **primary disability category**?","type":"dropdown","options":DISABILITY_OPTIONS},
    {"key":"education","q":"What is the participant's **highest level of education** completed?","type":"dropdown","options":EDUCATION_OPTIONS},
    # ── Consent (shared) ──────────────────────────────────────────────────────
    {"key":"consent_signed","q":"Has the participant signed the **Consent and Release of Information** form?","type":"yesno","required":True},
    {"key":"consent_date","q":"What is the **date the participant signed** the consent form?","type":"date","validation":["no_future"],"show_if":lambda a:a.get("consent_signed")=="Yes"},
    # ── SWTCIE ────────────────────────────────────────────────────────────────
    {"key":"drs_application_date","q":"What is the **DRS application date**?","type":"date","required":True,"validation":["no_future"],"show_if":is_swtcie},
    {"key":"submitting_agency","q":"Which **agency** is submitting this record?","type":"dropdown","options":AGENCIES_SWTCIE,"required":True,"show_if":is_swtcie},
    {"key":"worked_14c","q":"Has the participant **ever worked under subminimum wage** or with a **14(c) certificate holder**?","type":"yesno","required":True,"show_if":is_swtcie},
    {"key":"currently_employed","q":"Is the participant **currently employed** at the time of application?","type":"yesno","required":True,"show_if":is_swtcie},
    {"key":"employment_type","q":"What **type of employment** does the participant currently have?","type":"dropdown","options":["Subminimum wage job only","Competitive integrated employment only","Both"],"required":True,"show_if":lambda a:is_swtcie(a) and a.get("currently_employed")=="Yes"},
    {"key":"hours_subminimum","q":"How many **hours per week** does the participant work in the **subminimum wage job**?","type":"number","validation":[">0"],"show_if":lambda a:is_swtcie(a) and a.get("employment_type")=="Both"},
    {"key":"wage_subminimum","q":"What is the **hourly wage** in the **subminimum wage job**?\n\n*(If piece rate: divide last paycheck by hours worked.)*","type":"decimal","validation":[">=0"],"show_if":lambda a:is_swtcie(a) and a.get("employment_type")=="Both"},
    {"key":"hours_competitive","q":"How many **hours per week** does the participant work in the **competitive integrated job**?","type":"number","validation":[">0"],"show_if":lambda a:is_swtcie(a) and a.get("employment_type")=="Both"},
    {"key":"wage_competitive","q":"What is the **hourly wage** in the **competitive integrated job**?","type":"decimal","validation":[">=0"],"show_if":lambda a:is_swtcie(a) and a.get("employment_type")=="Both"},
    {"key":"hours_per_week","q":"How many **hours per week** does the participant work?","type":"number","validation":[">0"],"show_if":lambda a:is_swtcie(a) and a.get("currently_employed")=="Yes" and a.get("employment_type") in ["Subminimum wage job only","Competitive integrated employment only"]},
    {"key":"hourly_wage","q":"What is the participant's **hourly wage**?","type":"decimal","validation":[">=0"],"show_if":lambda a:is_swtcie(a) and a.get("currently_employed")=="Yes" and a.get("employment_type") in ["Subminimum wage job only","Competitive integrated employment only"]},
    {"key":"previous_subminimum","q":"Has the participant **previously worked** in a subminimum wage job?","type":"yesno","required":True,"show_if":lambda a:is_swtcie(a) and a.get("currently_employed")=="No"},
    {"key":"contemplating_subminimum","q":"Is the participant **considering a subminimum wage job** for the first time?","type":"yesno","show_if":lambda a:is_swtcie(a) and a.get("currently_employed")=="No" and a.get("previous_subminimum")=="No"},
    {"key":"receives_benefits_sw","q":"Does the participant receive **SSI, SSDI, or waiver funding**?","type":"yesno","required":True,"show_if":is_swtcie},
    {"key":"benefits_planning_offered_sw","q":"Has **benefits planning** been offered to the participant?","type":"yesno","show_if":lambda a:is_swtcie(a) and a.get("receives_benefits_sw")=="Yes"},
    {"key":"service_type","q":"What **VR services** has the participant received?\n\n*(Select all that apply.)*","type":"multiselect","options":VR_SERVICES,"required":True,"show_if":is_swtcie},
    {"key":"vr_outcome","q":"How **effective** were the VR services overall?","type":"dropdown","options":["Satisfactory","Unsatisfactory"],"show_if":is_swtcie},
    {"key":"obtained_cie","q":"Did the participant **obtain competitive integrated employment (CIE)** at any time?","type":"yesno","show_if":is_swtcie},
    {"key":"cie_start_date","q":"What is the **CIE employment start date**?","type":"date","show_if":lambda a:is_swtcie(a) and a.get("obtained_cie")=="Yes"},
    {"key":"job_title","q":"What is the participant's **job title**?","type":"text","show_if":lambda a:is_swtcie(a) and a.get("obtained_cie")=="Yes"},
    {"key":"job_duties","q":"What are the participant's **primary job duties**?","type":"text","show_if":lambda a:is_swtcie(a) and a.get("obtained_cie")=="Yes"},
    {"key":"cie_hours_per_week","q":"How many **hours per week** does the participant work in this CIE position?","type":"number","validation":[">0"],"show_if":lambda a:is_swtcie(a) and a.get("obtained_cie")=="Yes"},
    {"key":"cie_wage","q":"What is the participant's **hourly wage** in this CIE position?","type":"decimal","validation":[">=0"],"show_if":lambda a:is_swtcie(a) and a.get("obtained_cie")=="Yes"},
    {"key":"closure_date","q":"What is the **DRS case closure date**?","type":"date","required":True,"validation":["no_future"],"show_if":is_swtcie},
    {"key":"closure_status","q":"What is the participant's **employment status at case closure**?","type":"dropdown","options":["Employed in subminimum wage job only","Employed in CIE job only","Employed in both subminimum wage and CIE","Not employed"],"required":True,"show_if":is_swtcie},
    {"key":"closure_smw_hours","q":"At closure — **hours per week** (subminimum wage job)?","type":"number","validation":[">=0"],"show_if":lambda a:is_swtcie(a) and a.get("closure_status") in ["Employed in subminimum wage job only","Employed in both subminimum wage and CIE"]},
    {"key":"closure_smw_wage","q":"At closure — **hourly wage** (subminimum wage job)?","type":"decimal","validation":[">=0"],"show_if":lambda a:is_swtcie(a) and a.get("closure_status") in ["Employed in subminimum wage job only","Employed in both subminimum wage and CIE"]},
    {"key":"closure_cie_hours","q":"At closure — **hours per week** (CIE job)?","type":"number","validation":[">=0"],"show_if":lambda a:is_swtcie(a) and a.get("closure_status") in ["Employed in CIE job only","Employed in both subminimum wage and CIE"]},
    {"key":"closure_cie_wage","q":"At closure — **hourly wage** (CIE job)?","type":"decimal","validation":[">=0"],"show_if":lambda a:is_swtcie(a) and a.get("closure_status") in ["Employed in CIE job only","Employed in both subminimum wage and CIE"]},
    # ── Pathways ──────────────────────────────────────────────────────────────
    {"key":"cil","q":"Which **Center for Independent Living (CIL)** is submitting this record?","type":"dropdown","options":CILS_PATHWAYS,"required":True,"show_if":is_pathways},
    {"key":"school_lea","q":"What is the name of the **partner school / LEA**?","type":"text","required":True,"show_if":is_pathways},
    {"key":"age_group","q":"What **age group** does this participant fall into?","type":"dropdown","options":["Child (ages 10–13)","Youth (ages 14–24)"],"required":True,"show_if":is_pathways},
    {"key":"contact_method","q":"How was this participant **first contacted**?","type":"dropdown","options":["In-person (school visit / classroom)","Email / phone outreach","IEP meeting","Referral from teacher / school staff","Outreach event (transition fair, kick-off, open house)","Paperwork sent home / mailed","Word of mouth","Other"],"show_if":is_pathways},
    {"key":"iep_status","q":"Does the participant have an **active IEP**?","type":"dropdown","options":["Yes","No","Unknown"],"show_if":is_pathways},
    {"key":"underserved","q":"Is this participant from an **underserved community**?","type":"yesno","required":True,"show_if":is_pathways},
    {"key":"underserved_categories","q":"Which **underserved categories** apply?\n\n*(Select all that apply.)*","type":"multiselect","options":UNDERSERVED_CATS,"show_if":lambda a:is_pathways(a) and a.get("underserved")=="Yes"},
    {"key":"receives_benefits_pw","q":"Does the participant receive **SSI, SSDI, or waiver funding**?","type":"yesno","required":True,"show_if":is_pathways},
    {"key":"benefits_planning_offered_pw","q":"Has **benefits planning** been offered to the participant?","type":"yesno","show_if":lambda a:is_pathways(a) and a.get("receives_benefits_pw")=="Yes"},
    {"key":"pathways_services","q":"Which **project services** has the participant received?\n\n*(Select all that apply.)*","type":"multiselect","options":PATHWAYS_SERVICES,"required":True,"show_if":is_pathways},
    {"key":"pathways_svc_effectiveness","q":"How **effective** were the services overall?","type":"dropdown","options":["Satisfactory","Not Satisfactory"],"show_if":is_pathways},
    {"key":"pathways_obtained_cie","q":"Did the participant obtain **competitive integrated employment (CIE)**?","type":"yesno","show_if":is_pathways},
    {"key":"pathways_cie_start","q":"What is the **CIE employment start date**?","type":"date","show_if":lambda a:is_pathways(a) and a.get("pathways_obtained_cie")=="Yes"},
    {"key":"pathways_cie_title","q":"What is the participant's **job title**?","type":"text","show_if":lambda a:is_pathways(a) and a.get("pathways_obtained_cie")=="Yes"},
    {"key":"pathways_cie_hours","q":"How many **hours per week** in this CIE position?","type":"number","validation":[">0"],"show_if":lambda a:is_pathways(a) and a.get("pathways_obtained_cie")=="Yes"},
    {"key":"pathways_cie_wage","q":"What is the **hourly wage** in this CIE position?","type":"decimal","validation":[">=0"],"show_if":lambda a:is_pathways(a) and a.get("pathways_obtained_cie")=="Yes"},
    {"key":"pse_status","q":"Did the participant **enroll in or plan to enroll in post-secondary education (PSE)**?","type":"dropdown","options":["Yes","No","In progress / planning"],"show_if":is_pathways},
    {"key":"pse_institution","q":"What is the **name of the PSE institution**?","type":"text","show_if":lambda a:is_pathways(a) and a.get("pse_status")=="Yes"},
    {"key":"pathways_closure_date","q":"What is the **case closure date**?","type":"date","required":True,"validation":["no_future"],"show_if":is_pathways},
    {"key":"pathways_closure_status","q":"What is the participant's **status at case closure**?","type":"dropdown","options":["Employed in CIE only","Enrolled in PSE only","Both employed and enrolled in PSE","Not employed / not enrolled"],"required":True,"show_if":is_pathways},
    {"key":"pathways_closure_pse_inst","q":"What **PSE institution** is the participant enrolled in at closure?","type":"text","show_if":lambda a:is_pathways(a) and a.get("pathways_closure_status") in ["Enrolled in PSE only","Both employed and enrolled in PSE"]},
    # ── CIP ───────────────────────────────────────────────────────────────────
    {"key":"organization","q":"What **organization or agency** does this participant work for?","type":"text","required":True,"show_if":is_cip},
    {"key":"professional_role","q":"What is the participant's **professional role**?","type":"dropdown","options":PROF_ROLES,"required":True,"show_if":is_cip},
    {"key":"state","q":"Which **state** is the participant based in?","type":"dropdown","options":US_STATES,"required":True,"show_if":is_cip},
    {"key":"training_type","q":"What **type of CIP training** did the participant attend?\n\n*(Intensive = badging/credential · Targeted = webinar/seminar · Universal = fact sheet/toolkit)*","type":"dropdown","options":["Intensive","Targeted","Universal"],"required":True,"show_if":is_cip},
    {"key":"training_format","q":"What was the **training delivery format**?","type":"dropdown","options":TRAINING_FORMATS,"required":True,"show_if":is_cip},
    {"key":"training_topic","q":"What was the **training topic or subject area**?","type":"text","required":True,"show_if":is_cip},
    {"key":"training_date","q":"What was the **date of the training**?","type":"date","validation":["no_future"],"show_if":is_cip},
    {"key":"training_completed","q":"Did the participant **complete** the training?","type":"yesno","required":True,"show_if":is_cip},
    {"key":"prev_cip","q":"Has this participant **previously attended CIP training**?","type":"yesno","show_if":is_cip},
    {"key":"quality_rating","q":"**Quality rating** — clarity, structure, effectiveness of training.\n\n*(Score 0–100. Project target: ≥ 80)*","type":"number","validation":[">=0","<=100"],"show_if":is_cip},
    {"key":"relevance_rating","q":"**Relevance rating** — usability and enhancement of practice.\n\n*(Score 0–100. Project target: ≥ 80)*","type":"number","validation":[">=0","<=100"],"show_if":is_cip},
    {"key":"usefulness_rating","q":"**Usefulness rating** — application and confidence gained.\n\n*(Score 0–100. Project target: ≥ 80)*","type":"number","validation":[">=0","<=100"],"show_if":is_cip},
    {"key":"knowledge_increase","q":"Did the participant report an **increase in knowledge**?\n\n*(Project target: ≥ 90%)*","type":"yesno","required":True,"show_if":is_cip},
    {"key":"confidence_increase","q":"Did the participant report an **increase in confidence**?\n\n*(Project target: ≥ 90%)*","type":"yesno","required":True,"show_if":is_cip},
    {"key":"skill_adoption_increase","q":"Did the participant report an **increase in skill adoption**?\n\n*(Project target: ≥ 80%)*","type":"yesno","required":True,"show_if":is_cip},
    {"key":"qualitative_notes","q":"Any **additional notes or qualitative feedback**?","type":"text","show_if":is_cip},
]

SECTION_PILLS: Dict[str,str] = {
    "project":                  "Project Selection",
    "first_name":               "Section 1 · Personal Details",
    "consent_signed":           "Section 2 · Consent & Application",
    "drs_application_date":     "Section 2 · DRS Application",
    "cil":                      "Section 2 · CIL & School",
    "organization":             "Section 2 · Professional Details",
    "worked_14c":               "Section 3 · Employment Background",
    "underserved":              "Section 3 · Underserved Status",
    "receives_benefits_sw":     "Section 4 · Benefits",
    "receives_benefits_pw":     "Section 4 · Benefits",
    "service_type":             "Section 5 · VR Services",
    "pathways_services":        "Section 5 · Services Received",
    "training_type":            "Section 5 · Training Details",
    "quality_rating":           "Section 6 · QRU Ratings",
    "knowledge_increase":       "Section 7 · KCS Outcomes",
    "obtained_cie":             "Section 6 · CIE Employment Outcome",
    "closure_date":             "Section 7 · Case Closure",
    "pathways_obtained_cie":    "Section 6 · Employment & PSE Outcomes",
    "pathways_closure_date":    "Section 7 · Case Closure",
}

# ── Navigation ─────────────────────────────────────────────────────────────────
def is_step_visible(step, answers):
    fn = step.get("show_if"); return fn(answers) if fn else True

def find_next_step_idx(from_idx, answers):
    for i in range(from_idx + 1, len(ALL_STEPS)):
        if is_step_visible(ALL_STEPS[i], answers): return i
    return None

def count_visible_steps(answers):
    return sum(1 for s in ALL_STEPS if is_step_visible(s, answers))

# ── Validation ─────────────────────────────────────────────────────────────────
def validate(step, value):
    is_empty = (value is None or value == "" or value == [] or (isinstance(value, str) and not value.strip()))
    if step.get("required") and is_empty: return "⚠️ This field is required — please provide an answer before continuing."
    if is_empty: return None
    for rule in step.get("validation", []):
        if rule == "email":
            if not re.match(r"^[^\s@]+@[^\s@]+\.[^\s@]+$", str(value).strip()): return "⚠️ Please enter a valid email address."
        elif rule == "no_future":
            d = value if isinstance(value, date) else None
            if d and d > date.today(): return "⚠️ This date cannot be in the future."
        elif rule == ">0":
            try:
                if float(value) <= 0: return "⚠️ Value must be greater than 0."
            except: return "⚠️ Please enter a valid number."
        elif rule == ">=0":
            try:
                if float(value) < 0: return "⚠️ Value must be 0 or greater."
            except: return "⚠️ Please enter a valid number."
        elif rule == "<=100":
            try:
                if float(value) > 100: return "⚠️ Value must be 100 or less."
            except: return "⚠️ Please enter a valid number."
    return None

# ── Build flat row (aligned to MASTER_COLUMNS) ─────────────────────────────────
def build_row(answers: Dict[str, Any]) -> Dict[str, Any]:
    # Start with all master columns set to None
    row = {col: None for col in MASTER_COLUMNS}
    row["submitted_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for step in ALL_STEPS:
        k = step["key"]
        if not is_step_visible(step, answers): continue
        val = answers.get(k)
        if isinstance(val, list):  val = ", ".join(str(v) for v in val)
        elif isinstance(val, date): val = val.strftime("%Y-%m-%d")
        if k in row:
            row[k] = val
    return row

# ── Azure Blob: download master Excel, append row, re-upload ───────────────────
def append_row_to_master_blob(new_row: Dict[str, Any]):
    """
    1. Download IIRER_Master_Data.xlsx from Blob (if it exists).
    2. Append the new row.
    3. Re-upload to the same blob name (overwrite).
    Returns (success: bool, message: str, row_count: int)
    """
    if not AZURE_CONNECTION_STRING:
        return False, "AZURE_STORAGE_CONNECTION_STRING not set in .env file.", 0

    try:
        from azure.storage.blob import BlobServiceClient
        client    = BlobServiceClient.from_connection_string(AZURE_CONNECTION_STRING)
        container = client.get_container_client(AZURE_CONTAINER_NAME)

        # Ensure container exists
        try:
            container.create_container()
        except Exception:
            pass  # already exists

        blob_client = container.get_blob_client(MASTER_BLOB_NAME)

        # ── Step 1: Download existing master file (or create fresh) ────────────
        try:
            download_stream = blob_client.download_blob()
            existing_bytes  = download_stream.readall()
            existing_df     = pd.read_excel(
                io.BytesIO(existing_bytes),
                sheet_name=MASTER_SHEET_NAME,
                dtype=str,       # read everything as string to avoid type coercion
            )
            # Ensure all master columns exist (handles schema drift gracefully)
            for col in MASTER_COLUMNS:
                if col not in existing_df.columns:
                    existing_df[col] = None
            existing_df = existing_df[MASTER_COLUMNS]  # enforce column order
        except Exception:
            # File doesn't exist yet — start fresh
            existing_df = pd.DataFrame(columns=MASTER_COLUMNS)

        # ── Step 2: Append the new row ─────────────────────────────────────────
        new_df     = pd.DataFrame([new_row])[MASTER_COLUMNS]
        updated_df = pd.concat([existing_df, new_df], ignore_index=True)
        row_count  = len(updated_df)

        # ── Step 3: Write back to buffer and upload ────────────────────────────
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            updated_df.to_excel(writer, index=False, sheet_name=MASTER_SHEET_NAME)
        buffer.seek(0)

        # overwrite=True replaces the existing blob — this is what triggers the
        # Blob Event (Microsoft.Storage.BlobCreated) that fires the ADF pipeline
        blob_client.upload_blob(buffer, overwrite=True)

        return True, MASTER_BLOB_NAME, row_count

    except ImportError:
        return False, "azure-storage-blob not installed. Run: pip install azure-storage-blob", 0
    except Exception as e:
        return False, str(e), 0

# ── Fallback: build local Excel for manual download ───────────────────────────
def build_local_excel(row: Dict[str, Any], project: str) -> io.BytesIO:
    df     = pd.DataFrame([row])[MASTER_COLUMNS]
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=MASTER_SHEET_NAME)
    buffer.seek(0)
    return buffer

# ── Session state ──────────────────────────────────────────────────────────────
def init_state():
    defaults = {
        "phase": "collecting",
        "step_idx": 0,
        "answers": {},
        "history": [],
        "error_msg": None,
        "history_initialized": False,
        "upload_result": None,   # dict with success, message, row_count, local_buffer
    }
    for k, v in defaults.items():
        if k not in st.session_state: st.session_state[k] = v
    if not st.session_state.history_initialized:
        st.session_state.history.append(("bot", _q_to_html(ALL_STEPS[0]["q"])))
        st.session_state.history_initialized = True

# ── Utilities ──────────────────────────────────────────────────────────────────
def _q_to_html(q):
    safe = html.escape(q)
    safe = re.sub(r"\*\*(.*?)\*\*", r"<strong>\1</strong>", safe)
    safe = re.sub(r"\*(.*?)\*",     r"<em>\1</em>",          safe)
    return safe.replace("\n", "<br>")

def _fmt_answer(key, value):
    if value is None: return "—"
    if isinstance(value, list): return ", ".join(str(v) for v in value)
    if isinstance(value, date): return value.strftime("%B %d, %Y")
    return str(value)

def _push(role, msg): st.session_state.history.append((role, msg))

def _build_summary_html(answers):
    project = answers.get("project","Unknown")
    bc = {"SWTCIE":"#1E5799","Pathways":"#1B7A4A","CIP":"#5B21B6"}.get(project,"#333")
    rows = [f'<strong>📋 Summary — <span style="background:{bc};color:#fff;padding:2px 10px;border-radius:12px;font-size:0.85em">{html.escape(project)}</span></strong><br><br>']
    for step in ALL_STEPS:
        k = step["key"]
        if k == "project" or k not in answers or not is_step_visible(step, answers): continue
        label = step["q"].split("\n")[0].replace("**","")
        rows.append(f"<strong>{html.escape(label)}</strong><br>{html.escape(_fmt_answer(k, answers[k]))}<br><br>")
    return "".join(rows)

def _reset():
    for k in list(st.session_state.keys()): del st.session_state[k]
    st.rerun()

# ── CSS ────────────────────────────────────────────────────────────────────────
def apply_css():
    st.markdown(f"""
<style>
html,body,.stApp{{background:{LIGHT_BG};font-family:'Segoe UI',Arial,sans-serif;}}
#MainMenu,footer,header{{visibility:hidden;}}
.cb-header{{background:{DARK_BLUE};color:white;padding:1rem 1.5rem;border-radius:14px 14px 0 0;display:flex;align-items:center;gap:0.75rem;}}
.cb-icon{{width:46px;height:46px;background:{ORANGE};border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:1.4rem;flex-shrink:0;}}
.cb-header h1{{margin:0;font-size:1.1rem;font-weight:700;}}
.cb-header p{{margin:0;font-size:0.75rem;opacity:0.7;}}
.proj-badge{{margin-left:auto;padding:0.25rem 0.85rem;border-radius:20px;font-size:0.75rem;font-weight:700;}}
.proj-swtcie{{background:#3b7dd8;color:#fff;}}.proj-pathways{{background:#1B7A4A;color:#fff;}}.proj-cip{{background:#7C3AED;color:#fff;}}.proj-none{{background:#555;color:#fff;}}
.cb-chat{{background:#fff;border:1px solid #dce3ef;border-top:none;padding:1.25rem 1.25rem 0.75rem;min-height:340px;height:auto !important;max-height:none !important;overflow:visible !important;}}
.cb-row-bot{{display:flex;justify-content:flex-start;margin-bottom:0.65rem;align-items:flex-end;gap:0.45rem;}}
.cb-row-user{{display:flex;justify-content:flex-end;margin-bottom:0.65rem;}}
.cb-av{{width:32px;height:32px;background:{DARK_BLUE};color:white;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:0.72rem;font-weight:700;flex-shrink:0;}}
.cb-bot{{background:{DARK_BLUE};color:white;padding:0.65rem 1rem;border-radius:4px 14px 14px 14px;max-width:78%;line-height:1.55;font-size:0.91rem;}}
.cb-user{{background:{ORANGE};color:white;padding:0.65rem 1rem;border-radius:14px 4px 14px 14px;max-width:78%;line-height:1.55;font-size:0.91rem;text-align:right;}}
.cb-input{{background:#fff;border:1px solid #dce3ef;border-top:3px solid {ORANGE};border-radius:0 0 14px 14px;padding:1rem 1.25rem;}}
.stProgress>div>div>div>div{{background:{ORANGE} !important;}}
.stButton>button{{background:{ORANGE};color:white;border:none;border-radius:8px;padding:0.5rem 1.5rem;font-weight:700;font-size:0.9rem;}}
.stButton>button:hover{{background:#c93d1f;}}
.sec-pill{{display:inline-block;background:{DARK_BLUE}1a;color:{DARK_BLUE};font-size:0.68rem;font-weight:700;padding:0.15rem 0.55rem;border-radius:20px;letter-spacing:0.06em;text-transform:uppercase;margin-bottom:0.35rem;}}
div[data-testid="stRadio"] > div{{flex-direction:row !important;gap:0.75rem;}}

div[data-testid="stRadio"] label{{
    background:white !important;
    color:{DARK_BLUE} !important;
    border:1px solid #dce3ef !important;
    border-radius:20px !important;
    padding:8px 18px !important;
    font-weight:600 !important;
}}

div[data-testid="stRadio"] label p{{
    color:{DARK_BLUE} !important;
    margin:0 !important;
}}

div[data-testid="stRadio"] label:has(input:checked){{
    background:{ORANGE} !important;
    border-color:{ORANGE} !important;
}}

div[data-testid="stRadio"] label:has(input:checked) p{{
    color:white !important;
}}
.stTextInput input,.stNumberInput input{{border-radius:8px;border-color:#c0cde0;}}
.stSelectbox>div>div,.stMultiSelect>div>div{{border-radius:8px !important;}}
</style>""", unsafe_allow_html=True)


def render_history():
    parts = ['<div class="cb-chat">']
    for role, msg in st.session_state.history:
        if role == "bot":
            parts.append(
                f'<div class="cb-row-bot"><div class="cb-av">AI</div><div class="cb-bot">{msg}</div></div>'
            )
        else:
            parts.append(
                f'<div class="cb-row-user"><div class="cb-user">{msg}</div></div>'
            )

    parts.append("</div>")
    st.markdown("".join(parts), unsafe_allow_html=True)

# ── Input widget ───────────────────────────────────────────────────────────────
def render_input(step):
    key = step["key"]; kind = step["type"]; wk = f"w_{key}"
    if key in SECTION_PILLS: st.markdown(f'<div class="sec-pill">{SECTION_PILLS[key]}</div>', unsafe_allow_html=True)
    if kind in ("text","email"): return st.text_input("", key=wk, placeholder="Type here…", label_visibility="collapsed")
    if kind == "date": return st.date_input("", key=wk, min_value=date(1900,1,1), max_value=date.today(), value=None, label_visibility="collapsed")
    if kind == "yesno": return st.radio("", ["Yes","No"], key=wk, index=None, horizontal=True, label_visibility="collapsed")
    if kind == "dropdown":
        opts = ["— select —"] + (step.get("options") or [])
        raw  = st.selectbox("", opts, key=wk, label_visibility="collapsed")
        return raw if raw != "— select —" else None
    if kind == "multiselect": return st.multiselect("", step.get("options") or [], key=wk, label_visibility="collapsed")
    if kind == "number":
        max_v = 100 if "rating" in key else 99999
        raw = st.number_input("", min_value=0, max_value=max_v, step=1, key=wk, label_visibility="collapsed")
        return int(raw) if raw is not None else None
    if kind == "decimal":
        raw = st.number_input("", min_value=0.0, step=0.01, format="%.2f", key=wk, label_visibility="collapsed")
        return float(raw) if raw is not None else None
    return None

# ── Phases ─────────────────────────────────────────────────────────────────────
def phase_collecting():
    answers  = st.session_state.answers
    step_idx = st.session_state.step_idx
    step     = ALL_STEPS[step_idx]
    done     = sum(1 for s in ALL_STEPS if s["key"] in answers and is_step_visible(s, answers))
    total    = count_visible_steps(answers)
    st.progress(done / max(total, 1))
    st.caption(f"Question {done + 1} of ~{total}")
    render_history()
    st.markdown('<div class="cb-input">', unsafe_allow_html=True)
    value = render_input(step)
    if st.session_state.error_msg:
        st.markdown(f'<p style="color:{ORANGE};font-size:0.85rem;font-weight:600">{html.escape(st.session_state.error_msg)}</p>', unsafe_allow_html=True)
    col_s, col_k = st.columns([3, 1])
    with col_s:
        if st.button("Submit →", key="btn_submit", use_container_width=True): _on_submit(step, value, answers)
    with col_k:
        if not step.get("required"):
            if st.button("Skip", key="btn_skip", use_container_width=True): _advance(step, None, "(skipped)", answers)
    st.markdown("</div>", unsafe_allow_html=True)

def _on_submit(step, value, answers):
    err = validate(step, value)
    if err: st.session_state.error_msg = err; st.rerun(); return
    st.session_state.error_msg = None
    _advance(step, value, _fmt_answer(step["key"], value), answers)

def _advance(step, value, display, answers):
    if value is not None: answers[step["key"]] = value
    st.session_state.answers = answers
    _push("user", html.escape(display))
    next_idx = find_next_step_idx(st.session_state.step_idx, answers)
    if next_idx is None:
        _push("bot", _build_summary_html(answers))
        _push("bot", "✅ That's all the questions!<br><br>Please review your responses above. Click <strong>Confirm & Append to Master File</strong> to save, or <strong>Start Over</strong> to restart.")
        st.session_state.phase = "confirming"
    else:
        st.session_state.step_idx = next_idx
        _push("bot", _q_to_html(ALL_STEPS[next_idx]["q"]))
    st.rerun()

def phase_confirming():
    render_history()
    st.markdown('<div class="cb-input">', unsafe_allow_html=True)
    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("☁️ Confirm & Append to Master File", use_container_width=True):
            with st.spinner("Appending row to master Excel in Azure Blob…"):
                row             = build_row(st.session_state.answers)
                success, message, row_count = append_row_to_master_blob(row)
                proj            = st.session_state.answers.get("project","")
                local_buffer    = None if success else build_local_excel(row, proj)

            st.session_state.upload_result = {
                "success":      success,
                "message":      message,
                "row_count":    row_count,
                "project":      proj,
                "local_buffer": local_buffer,
            }

            if success:
                _push("bot",
                    f"🎉 <strong>Row appended successfully!</strong><br><br>"
                    f"<strong>Project:</strong> {html.escape(proj)}<br>"
                    f"<strong>Master file:</strong> <code>iirer-submissions/{html.escape(MASTER_BLOB_NAME)}</code><br>"
                    f"<strong>Total rows in file:</strong> {row_count}<br><br>"
                    f"The Blob event trigger will fire ADF automatically — "
                    f"your data will flow through validation and appear in Power BI shortly."
                )
            else:
                _push("bot",
                    f"⚠️ <strong>Azure upload failed:</strong> {html.escape(message)}<br><br>"
                    "No problem — download the Excel file below and upload it manually to Blob Storage."
                )
            st.session_state.phase = "done"
            st.rerun()
    with col_b:
        if st.button("🔄 Start Over", use_container_width=True): _reset()
    st.markdown("</div>", unsafe_allow_html=True)

def phase_done():
    render_history()
    st.markdown('<div class="cb-input">', unsafe_allow_html=True)
    result = st.session_state.get("upload_result", {})
    if not result.get("success") and result.get("local_buffer"):
        result["local_buffer"].seek(0)
        proj      = result.get("project","UNKNOWN")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname     = f"{proj}_submission_{timestamp}.xlsx"
        st.download_button(
            label="⬇️ Download Excel (manual upload fallback)",
            data=result["local_buffer"],
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    if st.button("➕ Submit Another Participant", use_container_width=True): _reset()
    st.markdown("</div>", unsafe_allow_html=True)

# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    apply_css(); init_state()
    project   = st.session_state.answers.get("project","")
    badge_cls = {"SWTCIE":"proj-swtcie","Pathways":"proj-pathways","CIP":"proj-cip"}.get(project,"proj-none")
    badge_txt = project or "Select Project"
    st.markdown(f"""
<div class="cb-header">
  <div class="cb-icon">💬</div>
  <div><h1>IIRER Intake Assistant</h1><p>Illinois Institute for Rehabilitation and Employment Research · UIUC</p></div>
  <span class="proj-badge {badge_cls}">{html.escape(badge_txt)}</span>
</div>""", unsafe_allow_html=True)
    phase = st.session_state.phase
    if   phase == "collecting": phase_collecting()
    elif phase == "confirming": phase_confirming()
    elif phase == "done":       phase_done()
    else: st.error(f"Unknown phase: {phase}")

if __name__ == "__main__":
    main()
