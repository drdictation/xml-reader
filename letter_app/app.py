import csv
import html
import io
import json
import os
import re
import subprocess
import zipfile
from datetime import datetime
from xml.etree import ElementTree as ET

import fitz  # PyMuPDF
from bs4 import BeautifulSoup
from flask import Flask, jsonify, render_template, request, send_file, url_for
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

app = Flask(__name__)

BASE_DIR = os.path.dirname(__file__)
ROOT_DIR = os.path.abspath(os.path.join(BASE_DIR, ".."))
XML_DIR = "/Volumes/Chamara Bac/FINAL 09042026"
TEMPLATE_DIR = os.path.join(ROOT_DIR, "Letter templates")
PATHOLOGY_TEMPLATE = os.path.join(ROOT_DIR, "pathology.pdf")
RADIOLOGY_TEMPLATE_PDF = os.path.join(ROOT_DIR, "radioltemplate.pdf")
RADIOLOGY_TEMPLATE_JPEG = os.path.join(ROOT_DIR, "radioltemplate.jpeg")
IUS_TEMPLATE = os.path.join(ROOT_DIR, "IUS_template.pdf")
MANOMETRY_TEMPLATE = os.path.join(ROOT_DIR, "Manometry_Referral_Form.docx")
PRIVATE_IRON_TEMPLATE = os.path.join(ROOT_DIR, "private Iron infusion.pdf")
PUBLIC_IRON_TEMPLATE = os.path.join(ROOT_DIR, "public iron infusion.pdf")
PUBLIC_ENDOSCOPY_RECIPIENT = "Endoscopy.Waitinglist@svhm.org.au"
GENERATED_ENDOSCOPY_DIR = os.path.join(BASE_DIR, "generated_endoscopy")

ENDOSCOPY_PAGE_1_TEMPLATES = {
    "gastroscopy": os.path.join(ROOT_DIR, "1st_gas.pdf"),
    "colonoscopy": os.path.join(ROOT_DIR, "1st_colon.pdf"),
    "gascolon": os.path.join(ROOT_DIR, "1st_gascolon.pdf"),
}

ENDOSCOPY_PAGE_2_TEMPLATES = {
    "mains": {
        "gastroscopy": os.path.join(ROOT_DIR, "2nd_mains_gas.pdf"),
        "colonoscopy": os.path.join(ROOT_DIR, "2nd_mains_colon.pdf"),
        "gascolon": os.path.join(ROOT_DIR, "2nd_mains_gascolon.pdf"),
    },
    "svop": {
        "gastroscopy": os.path.join(ROOT_DIR, "2nd_svop_gas.pdf"),
        "colonoscopy": os.path.join(ROOT_DIR, "2nd_svop_colon.pdf"),
        "gascolon": os.path.join(ROOT_DIR, "2nd_svop_gas_and_colon.pdf"),
    },
}
PATIENT_DATABASE_FILE = os.path.join(ROOT_DIR, "output", "patient-database.csv")
PATIENT_INDEX_FILE = os.path.join(ROOT_DIR, "output", "patient-index.json")
CACHE_FILE = os.path.join(BASE_DIR, "data_cache.json")
GENERATED_LETTERS_DIR = os.path.join(BASE_DIR, "generated_letters")
GENERATED_PATHOLOGY_DIR = os.path.join(BASE_DIR, "generated_pathology")
GENERATED_RADIOLOGY_DIR = os.path.join(BASE_DIR, "generated_radiology")
GENERATED_IUS_DIR = os.path.join(BASE_DIR, "generated_ius")
GENERATED_MANOMETRY_DIR = os.path.join(BASE_DIR, "generated_manometry")
GENERATED_IRON_DIR = os.path.join(BASE_DIR, "generated_iron")
DATA_CACHE_VERSION = 3
SENDER_EMAIL = "drchamarabasnayake@gmail.com"
IRON_SENDER_EMAIL = "chamara.basnayake@svha.org.au"
PRACTICE_CC = "office@focusgastro.com.au"
PUBLIC_IRON_RECIPIENT = "infusioncentre.AAU@svha.org.au"
PRIVATE_IRON_RECIPIENT = {
    "name": "SVPHM Onc.infusion - SVPHM",
    "email": "SVPHM.onc.infusion@svha.org.au",
}

PATHOLOGY_LAYOUT = {
    "top_last_address": (14, 93, 140, 157),
    "top_given_names": (142, 93, 346, 121),
    "top_sex": (350, 93, 379, 121),
    "top_dob": (383, 93, 470, 121),
    "top_phone_home": (384, 123, 489, 152),
    "top_phone_bus": (497, 123, 580, 152),
    "top_tests": (14, 171, 430, 314),
    "top_notes": (14, 309, 366, 369),
    "top_copy_reports": (14, 494, 208, 540),
    "top_request_date_point": (438, 357),
    "bottom_surname": (14, 641, 140, 665),
    "bottom_given_names": (144, 641, 348, 665),
    "bottom_sex": (350, 641, 379, 665),
    "bottom_dob": (383, 641, 470, 665),
    "bottom_address": (14, 677, 350, 706),
    "bottom_phone_home": (384, 677, 489, 706),
    "bottom_phone_bus": (497, 677, 580, 706),
    "bottom_tests": (14, 742, 350, 787),
}

RADIOLOGY_IMAGE_SIZE = (2138, 1487)
RADIOLOGY_LAYOUT = {
    "name": (330, 300, 1340, 350),
    "dob": (1650, 300, 2080, 350),
    "address": (330, 410, 1340, 505),
    "medicare": (1650, 410, 2080, 465),
    "phone": (330, 560, 1040, 615),
    "request_for": (90, 655, 990, 905),
    "clinical_details": (1080, 655, 1865, 905),
    "date": (1090, 1260, 1450, 1320),
    "category_private": (1484, 942),
    "category_wc": (1484, 992),
    "category_pension": (1484, 1041),
    "category_vet_aff": (1622, 942),
    "category_tac": (1622, 992),
}

IUS_LAYOUT = {
    "date": (132, 236, 289, 256),
    "urgent_yes": (454, 250),
    "urgent_no": (499, 250),
    "patient_sticker": (72, 291, 539, 367),
    "reason_symptoms": (79, 512),
    "reason_ibd": (228, 512),
    "reason_incomplete_colonoscopy": (410, 512),
    "medical_history": (74, 544, 535, 621),
    "abnormal_markers_result": (340, 650, 535, 670),
    "faecal_calprotectin_result": (340, 673, 535, 693),
    "pregnant_yes": (173, 691),
    "pregnant_no": (213, 691),
    "gestation_weeks": (326, 692, 535, 713),
}

data_cache = {
    "patients": [],
    "doctors": [],
    "loaded": False,
    "version": DATA_CACHE_VERSION,
}
patient_xml_field_cache = {}
xml_file_lookup_cache = None
WORD_NAMESPACE = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def clean_text(value):
    if value is None:
        return ""
    return html.unescape(str(value)).replace("\r", "").strip()


def unescape_xml(text):
    return clean_text(text)


def extract_tag(block, tag):
    match = re.search(f"<{tag}>(.*?)</{tag}>", block, re.IGNORECASE | re.DOTALL)
    return unescape_xml(match.group(1)) if match else ""


def normalize_name(value):
    return re.sub(r"\s+", " ", clean_text(value).lower())


def normalize_iso_date(value):
    text = clean_text(value)
    if not text or text.startswith("0000"):
        return ""
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return text


def format_display_date(value):
    iso_value = normalize_iso_date(value)
    if not iso_value:
        return ""
    try:
        return datetime.strptime(iso_value, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return iso_value


def format_medicare_value(number, ref_number=""):
    base = clean_text(number)
    ref = clean_text(ref_number)
    if base and ref:
        return f"{base} / {ref}"
    return base or ref


def compact_lines(value):
    lines = [line.strip() for line in clean_text(value).splitlines()]
    return "\n".join(line for line in lines if line)


def compact_single_line(value):
    return " ".join(compact_lines(value).split())


def build_safe_filename(stem, suffix, extension):
    name = re.sub(r"[^\w\s-]", "", stem or "Unknown").strip().replace(" ", "_")
    return f"{name}_{suffix}_{datetime.now().strftime('%Y-%m-%d')}.{extension}"


def escape_applescript_string(value):
    return clean_text(value).replace("\\", "\\\\").replace('"', '\\"')


def normalize_email_recipient(recipient):
    if isinstance(recipient, dict):
        return clean_text(recipient.get("name")), clean_text(recipient.get("email") or recipient.get("address"))
    return "", clean_text(recipient)


def write_outlook_draft(subject, output_path, *, to_recipients=None, cc_recipients=None, sender_email=SENDER_EMAIL):
    applescript_path = os.path.join(BASE_DIR, "send_email.applescript")
    to_recipients = to_recipients or []
    cc_recipients = cc_recipients or []

    with open(applescript_path, "w", encoding="utf-8") as fh:
        fh.write(
            f'''
tell application "Microsoft Outlook"
    set newMessage to make new outgoing message with properties {{subject:"{escape_applescript_string(subject)}"}}
    try
        set theAccount to (first account whose email address is "{escape_applescript_string(sender_email)}")
        set account of newMessage to theAccount
    end try
'''
        )

        for recipient in to_recipients:
            name, address = normalize_email_recipient(recipient)
            if not address:
                continue
            if name:
                fh.write(
                    f'    make new recipient at newMessage with properties {{email address:{{name:"{escape_applescript_string(name)}", address:"{escape_applescript_string(address)}"}}}}\n'
                )
            else:
                fh.write(
                    f'    make new recipient at newMessage with properties {{email address:{{address:"{escape_applescript_string(address)}"}}}}\n'
                )

        for recipient in cc_recipients:
            name, address = normalize_email_recipient(recipient)
            if not address:
                continue
            if name:
                fh.write(
                    f'    make new cc recipient at newMessage with properties {{email address:{{name:"{escape_applescript_string(name)}", address:"{escape_applescript_string(address)}"}}}}\n'
                )
            else:
                fh.write(
                    f'    make new cc recipient at newMessage with properties {{email address:{{address:"{escape_applescript_string(address)}"}}}}\n'
                )

        fh.write(
            f'''
    make new attachment at newMessage with properties {{file:POSIX file "{escape_applescript_string(output_path)}"}}
    open newMessage
end tell
'''
        )

    subprocess.run(["osascript", applescript_path], check=True)


def read_cached_data():
    if not os.path.exists(CACHE_FILE):
        return {}
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception as exc:
        print("Cache read failed", exc)
        return {}


def read_patient_index_lookup():
    if not os.path.exists(PATIENT_INDEX_FILE):
        return {}

    try:
        with open(PATIENT_INDEX_FILE, "r", encoding="utf-8") as fh:
            entries = json.load(fh)
    except Exception as exc:
        print("Could not read patient-index.json", exc)
        return {}

    lookup = {}
    for entry in entries:
        name = clean_text(entry.get("name"))
        if not name:
            continue
        dob = normalize_iso_date(entry.get("dob"))
        key = f"{normalize_name(name)}|{dob}"
        lookup[key] = entry
        lookup.setdefault(normalize_name(name), entry)
    return lookup


def get_xml_file_lookup():
    global xml_file_lookup_cache
    if xml_file_lookup_cache is not None:
        return xml_file_lookup_cache

    lookup = {}
    if os.path.exists(XML_DIR):
        for root, _, files in os.walk(XML_DIR):
            for file_name in files:
                if file_name.lower().endswith(".xml"):
                    lookup.setdefault(file_name, os.path.join(root, file_name))

    xml_file_lookup_cache = lookup
    return lookup


def resolve_patient_xml_path(file_path="", file_name=""):
    path = clean_text(file_path)
    if path and os.path.exists(path):
        return path

    fallback_name = os.path.basename(clean_text(file_name) or path)
    if not fallback_name:
        return ""
    return get_xml_file_lookup().get(fallback_name, "")


def read_patient_xml_fields(file_path="", file_name=""):
    path = resolve_patient_xml_path(file_path, file_name)
    if not path:
        return {}

    cached = patient_xml_field_cache.get(path)
    if cached is not None:
        return cached

    fields = {}
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as fh:
                content = fh.read()
            match = re.search(r"<patient>(.*?)</patient>", content, re.IGNORECASE | re.DOTALL)
            if match:
                block = match.group(1)
                medicare_number = extract_tag(block, "medicarenum")
                medicare_ref_number = extract_tag(block, "medicarerefnum")
                fields = {
                    "email": extract_tag(block, "emailaddress"),
                    "mobile_phone": extract_tag(block, "mobilephone"),
                    "home_phone": extract_tag(block, "homephone"),
                    "phone": extract_tag(block, "mobilephone") or extract_tag(block, "homephone"),
                    "medicare_number": medicare_number,
                    "medicare_ref_number": medicare_ref_number,
                    "medicare": format_medicare_value(medicare_number, medicare_ref_number),
                    "urn": extract_tag(block, "urnum") or extract_tag(block, "urn") or extract_tag(block, "externalid"),
                }
        except Exception as exc:
            print(f"Could not read patient XML fields from {path}: {exc}")

    patient_xml_field_cache[path] = fields
    return fields


def infer_name_parts(full_name):
    name = clean_text(full_name)
    if not name:
        return "", ""

    parts = name.split()
    if parts and parts[0].rstrip(".").lower() in {
        "mr",
        "mrs",
        "ms",
        "miss",
        "dr",
        "prof",
        "doctor",
        "a/prof",
    }:
        parts = parts[1:]

    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], parts[0]
    return " ".join(parts[:-1]), parts[-1]


def load_patients_from_exports():
    if not os.path.exists(PATIENT_DATABASE_FILE):
        return []

    index_lookup = read_patient_index_lookup()
    patients = []

    try:
        with open(PATIENT_DATABASE_FILE, "r", encoding="utf-8", newline="") as fh:
            reader = csv.DictReader(fh)
            for row in reader:
                name = clean_text(row.get("patient_name"))
                if not name:
                    continue
                if re.fullmatch(r"\d+_\d+", name):
                    continue

                dob = normalize_iso_date(row.get("date_of_birth"))
                lookup_key = f"{normalize_name(name)}|{dob}"
                index_entry = index_lookup.get(lookup_key) or index_lookup.get(normalize_name(name), {})
                xml_fields = read_patient_xml_fields(index_entry.get("filePath"), index_entry.get("file"))

                inferred_first, inferred_surname = infer_name_parts(name)
                full_address = clean_text(row.get("full_address"))
                mobile_phone = clean_text(row.get("mobile_phone")) or clean_text(xml_fields.get("mobile_phone"))
                home_phone = clean_text(row.get("home_phone")) or clean_text(xml_fields.get("home_phone"))
                medicare_number = clean_text(xml_fields.get("medicare_number"))
                medicare_ref_number = clean_text(xml_fields.get("medicare_ref_number"))

                patients.append(
                    {
                        "name": name,
                        "firstname": clean_text(index_entry.get("firstname")) or inferred_first,
                        "surname": clean_text(index_entry.get("surname")) or inferred_surname,
                        "dob": dob,
                        "sex": clean_text(row.get("sex")) or clean_text(index_entry.get("sex")),
                        "address": full_address,
                        "full_address": full_address,
                        "suburb": clean_text(row.get("suburb")) or clean_text(index_entry.get("suburb")),
                        "mobile_phone": mobile_phone,
                        "home_phone": home_phone,
                        "phone": mobile_phone or home_phone or clean_text(row.get("phone")) or clean_text(xml_fields.get("phone")) or clean_text(index_entry.get("phone")),
                        "email": clean_text(row.get("email")) or clean_text(xml_fields.get("email")),
                        "medicare_number": medicare_number,
                        "medicare_ref_number": medicare_ref_number,
                        "medicare": format_medicare_value(medicare_number, medicare_ref_number),
                        "urn": clean_text(index_entry.get("urn")) or clean_text(xml_fields.get("urn")),
                        "usual_gp": clean_text(row.get("usual_gp")),
                        "referring_doctor": clean_text(row.get("referring_doctor")),
                        "ref_clinic": clean_text(row.get("ref_clinic")),
                        "ref_address": clean_text(row.get("ref_address")),
                        "ref_phone": clean_text(row.get("ref_phone")),
                        "ref_email": clean_text(row.get("ref_email")),
                        "referral_date": normalize_iso_date(row.get("referral_date")),
                        "referral_expiry": normalize_iso_date(row.get("referral_expiry")),
                        "referral_status": clean_text(row.get("referral_status")),
                    }
                )
    except Exception as exc:
        print("Could not read patient-database.csv", exc)
        return []

    patients.sort(key=lambda item: item["name"])
    return patients


def scan_xml_directory():
    patients_list = []
    doctors_list = []
    seen_patients = set()
    seen_doctors = set()

    if not os.path.exists(XML_DIR):
        print(f"Directory not found: {XML_DIR}")
        return patients_list, doctors_list

    print("Scanning XML files...")
    for root, _, files in os.walk(XML_DIR):
        for file_name in files:
            if not file_name.lower().endswith(".xml"):
                continue

            filepath = os.path.join(root, file_name)
            try:
                with open(filepath, "r", encoding="utf-8", errors="ignore") as fh:
                    content = fh.read()

                if "<addressbook>" in content:
                    for ab_block in re.finditer(r"<addressbook>(.*?)</addressbook>", content, re.IGNORECASE | re.DOTALL):
                        block = ab_block.group(1)
                        name = extract_tag(block, "fullname")
                        if not name:
                            title = extract_tag(block, "title")
                            fname = extract_tag(block, "firstname")
                            sname = extract_tag(block, "surname")
                            name = f"{title} {fname} {sname}".strip()

                        email = extract_tag(block, "emailaddress")
                        clinic = extract_tag(block, "clinic")

                        if name and name not in seen_doctors:
                            seen_doctors.add(name)
                            doctors_list.append({"name": name, "email": email, "clinic": clinic})

                if "<patient>" not in content and "<patient " not in content:
                    continue

                for pat_block in re.finditer(r"<patient>(.*?)</patient>", content, re.IGNORECASE | re.DOTALL):
                    block = pat_block.group(1)
                    name = extract_tag(block, "fullname")
                    if not name:
                        fname = extract_tag(block, "firstname")
                        sname = extract_tag(block, "surname")
                        name = f"{fname} {sname}".strip()

                    if not name or name in seen_patients:
                        continue

                    full_address = ", ".join(
                        part
                        for part in [
                            extract_tag(block, "addressline1"),
                            extract_tag(block, "addressline2"),
                            extract_tag(block, "suburb"),
                            extract_tag(block, "state"),
                            extract_tag(block, "postcode"),
                        ]
                        if part
                    )

                    seen_patients.add(name)
                    medicare_number = extract_tag(block, "medicarenum")
                    medicare_ref_number = extract_tag(block, "medicarerefnum")
                    patients_list.append(
                        {
                            "name": name,
                            "firstname": extract_tag(block, "firstname"),
                            "surname": extract_tag(block, "surname"),
                            "dob": normalize_iso_date(extract_tag(block, "dob")),
                            "sex": extract_tag(block, "sex"),
                            "address": full_address,
                            "full_address": full_address,
                            "mobile_phone": extract_tag(block, "mobilephone"),
                            "home_phone": extract_tag(block, "homephone"),
                            "phone": extract_tag(block, "mobilephone") or extract_tag(block, "homephone"),
                            "email": extract_tag(block, "emailaddress"),
                            "medicare_number": medicare_number,
                            "medicare_ref_number": medicare_ref_number,
                            "medicare": format_medicare_value(medicare_number, medicare_ref_number),
                            "urn": extract_tag(block, "urnum") or extract_tag(block, "urn") or extract_tag(block, "externalid"),
                            "referring_doctor": extract_tag(block, "referringdoctor"),
                            "referral_date": normalize_iso_date(extract_tag(block, "referraldate")),
                        }
                    )
            except Exception as exc:
                print(f"Error reading {file_name}: {exc}")

    patients_list.sort(key=lambda item: item["name"])
    doctors_list.sort(key=lambda item: item["name"])
    return patients_list, doctors_list


def merge_doctors(existing_doctors, patients):
    merged = {}

    for doctor in existing_doctors or []:
        name = clean_text(doctor.get("name"))
        if not name:
            continue
        merged[normalize_name(name)] = {
            "name": name,
            "email": clean_text(doctor.get("email")),
            "clinic": clean_text(doctor.get("clinic")),
        }

    for patient in patients or []:
        name = clean_text(patient.get("referring_doctor"))
        if not name:
            continue
        key = normalize_name(name)
        current = merged.get(key, {"name": name, "email": "", "clinic": ""})
        if not current.get("clinic"):
            current["clinic"] = clean_text(patient.get("ref_clinic"))
        if not current.get("email"):
            current["email"] = clean_text(patient.get("ref_email"))
        merged[key] = current

    return sorted(merged.values(), key=lambda item: item["name"])


def load_data():
    if data_cache["loaded"]:
        return

    cached = read_cached_data()
    cached_doctors = cached.get("doctors", []) if isinstance(cached, dict) else []
    cached_patients = cached.get("patients", []) if isinstance(cached, dict) else []

    patients_list = load_patients_from_exports()
    doctors_list = cached_doctors

    if not doctors_list or not patients_list:
        scanned_patients, scanned_doctors = scan_xml_directory()
        if not patients_list:
            patients_list = scanned_patients
        if not doctors_list:
            doctors_list = scanned_doctors

    if not patients_list and cached_patients:
        patients_list = cached_patients

    doctors_list = merge_doctors(doctors_list, patients_list)

    data_cache["patients"] = patients_list
    data_cache["doctors"] = doctors_list
    data_cache["loaded"] = True
    data_cache["version"] = DATA_CACHE_VERSION

    print(f"Loaded {len(patients_list)} patients and {len(doctors_list)} doctors.")
    with open(CACHE_FILE, "w", encoding="utf-8") as fh:
        json.dump(
            {
                "version": DATA_CACHE_VERSION,
                "patients": patients_list,
                "doctors": doctors_list,
                "loaded": True,
            },
            fh,
        )


def html_to_story(html_content, style_normal, style_bullet):
    soup = BeautifulSoup(html_content, "html.parser")
    flowables = []

    class State:
        buffer = ""

    def flush_buffer():
        text = State.buffer.strip()
        if not text:
            State.buffer = ""
            return

        m_bullet = re.match(r"^((?:<[^>]+>|\s)*)([•·\uf0b7●◦\-])((?:<[^>]+>|\s)*)(.*)$", text)
        m_numbered = re.match(r"^((?:<[^>]+>|\s)*)(\d+)[.)]((?:<[^>]+>|\s)*)(.*)$", text)

        if m_bullet:
            content = m_bullet.group(1) + m_bullet.group(3) + m_bullet.group(4)
            flowables.append(Paragraph(f"<bullet>&bull;</bullet>{content.strip()}", style_bullet))
        elif m_numbered:
            idx = m_numbered.group(2) + "."
            content = m_numbered.group(1) + m_numbered.group(3) + m_numbered.group(4)
            flowables.append(Paragraph(f"<bullet>{idx}</bullet>{content.strip()}", style_bullet))
        else:
            flowables.append(Paragraph(text, style_normal))
        flowables.append(Spacer(1, 4))
        State.buffer = ""

    class ListState:
        ol_counters = {}

    def walk(node):
        if not hasattr(node, "name"):
            return

        if node.name is None:
            text = str(node).replace("\n", " ").replace("\r", "")
            text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            State.buffer += text
            return

        if node.name in {"div", "p"}:
            flush_buffer()
            for child in node.children:
                walk(child)
            flush_buffer()
            return

        if node.name == "br":
            State.buffer += "<br/>"
            return

        if node.name in {"ul", "ol"}:
            flush_buffer()
            for child in node.children:
                walk(child)
            flush_buffer()
            return

        if node.name == "li":
            flush_buffer()
            previous_buffer = State.buffer
            State.buffer = ""

            def walk_li(child):
                if not hasattr(child, "name") or child.name is None:
                    text = str(child).replace("\n", " ").replace("\r", "")
                    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                    State.buffer += text
                    return

                if child.name == "br":
                    State.buffer += "<br/>"
                    return

                if child.name in {"strong", "b"}:
                    State.buffer += "<b>"
                    for grandchild in child.children:
                        walk_li(grandchild)
                    State.buffer += "</b>"
                    return

                if child.name in {"em", "i"}:
                    State.buffer += "<i>"
                    for grandchild in child.children:
                        walk_li(grandchild)
                    State.buffer += "</i>"
                    return

                if child.name == "u":
                    State.buffer += "<u>"
                    for grandchild in child.children:
                        walk_li(grandchild)
                    State.buffer += "</u>"
                    return

                if child.name == "span":
                    style = child.get("style", "")
                    if "bold" in style:
                        State.buffer += "<b>"
                    if "underline" in style:
                        State.buffer += "<u>"
                    for grandchild in child.children:
                        walk_li(grandchild)
                    if "underline" in style:
                        State.buffer += "</u>"
                    if "bold" in style:
                        State.buffer += "</b>"
                    return

                for grandchild in child.children:
                    walk_li(grandchild)

            for child in node.children:
                walk_li(child)

            li_text = State.buffer.strip()
            if li_text:
                parent = node.parent
                is_ordered = False
                while parent:
                    if parent.name == "ol":
                        is_ordered = True
                        break
                    if parent.name == "ul":
                        break
                    parent = parent.parent

                bullet = "&bull;"
                if is_ordered:
                    parent_id = id(parent) if parent else 0
                    idx = ListState.ol_counters.get(parent_id, 1)
                    bullet = f"{idx}."
                    ListState.ol_counters[parent_id] = idx + 1

                flowables.append(Paragraph(f"<bullet>{bullet}</bullet>{li_text}", style_bullet))
                flowables.append(Spacer(1, 4))

            State.buffer = previous_buffer
            return

        if node.name in {"strong", "b"}:
            State.buffer += "<b>"
            for child in node.children:
                walk(child)
            State.buffer += "</b>"
            return

        if node.name in {"em", "i"}:
            State.buffer += "<i>"
            for child in node.children:
                walk(child)
            State.buffer += "</i>"
            return

        if node.name == "u":
            State.buffer += "<u>"
            for child in node.children:
                walk(child)
            State.buffer += "</u>"
            return

        if node.name == "span":
            style = node.get("style", "")
            if "bold" in style:
                State.buffer += "<b>"
            if "underline" in style:
                State.buffer += "<u>"
            for child in node.children:
                walk(child)
            if "underline" in style:
                State.buffer += "</u>"
            if "bold" in style:
                State.buffer += "</b>"
            return

        for child in node.children:
            walk(child)

    for child in soup.children:
        walk(child)

    flush_buffer()
    if not flowables:
        flowables.append(Paragraph(" ", style_normal))
    return flowables


def split_patient_name(patient):
    firstname = clean_text(patient.get("firstname"))
    surname = clean_text(patient.get("surname"))
    if firstname or surname:
        return firstname, surname
    return infer_name_parts(patient.get("name"))


def build_letter_patient_detail_lines(patient):
    name = clean_text(patient.get("name"))
    dob = format_display_date(patient.get("dob"))
    address = clean_text(patient.get("full_address")) or clean_text(patient.get("address"))
    phone = clean_text(patient.get("phone")) or clean_text(patient.get("mobile_phone")) or clean_text(patient.get("home_phone"))
    email = clean_text(patient.get("email"))
    medicare = clean_text(patient.get("medicare_number")) or clean_text(patient.get("medicare"))
    medicare_ref = clean_text(patient.get("medicare_ref_number"))

    lines = []
    if name:
        lines.append(f"Re: {name}")

    summary_parts = []
    if dob:
        summary_parts.append(f"DOB: {dob}")
    if phone:
        summary_parts.append(f"Phone: {phone}")
    if summary_parts:
        lines.append("   ".join(summary_parts))

    if address:
        lines.append(f"Address: {address}")
    if email:
        lines.append(f"Email: {email}")

    medicare_text = format_medicare_value(medicare, medicare_ref) if medicare_ref and "/" not in medicare else medicare
    if medicare_text:
        lines.append(f"Medicare: {medicare_text}")

    return lines


def draw_textbox(page, rect_key, text, *, fontsize=9.0, min_fontsize=6.0, align=0):
    value = compact_lines(text)
    if not value:
        return

    rect = fitz.Rect(*PATHOLOGY_LAYOUT[rect_key])
    size = fontsize
    while size >= min_fontsize:
        result = page.insert_textbox(
            rect,
            value,
            fontsize=size,
            fontname="helv",
            color=(0, 0, 0),
            align=align,
        )
        if result >= 0:
            return
        size -= 0.5

    page.insert_textbox(
        rect,
        value,
        fontsize=min_fontsize,
        fontname="helv",
        color=(0, 0, 0),
        align=align,
    )


def draw_text(page, point_key, text, *, fontsize=9.0):
    value = clean_text(text)
    if not value:
        return
    x, y = PATHOLOGY_LAYOUT[point_key]
    page.insert_text(fitz.Point(x, y), value, fontsize=fontsize, fontname="helv", color=(0, 0, 0))


def build_patient_name(patient):
    explicit_name = clean_text(patient.get("name"))
    if explicit_name:
        return explicit_name

    firstname, surname = split_patient_name(patient)
    return " ".join(part for part in [firstname, surname] if part).strip()


def radiology_page_width():
    return A4[1]


def radiology_page_height():
    return A4[0]


def radiology_scale_x(value):
    return value * radiology_page_width() / RADIOLOGY_IMAGE_SIZE[0]


def radiology_scale_y(value):
    return value * radiology_page_height() / RADIOLOGY_IMAGE_SIZE[1]


def radiology_rect(layout_key):
    x0, y0, x1, y1 = RADIOLOGY_LAYOUT[layout_key]
    return fitz.Rect(
        radiology_scale_x(x0),
        radiology_scale_y(y0),
        radiology_scale_x(x1),
        radiology_scale_y(y1),
    )


def radiology_point(layout_key):
    x, y = RADIOLOGY_LAYOUT[layout_key]
    return fitz.Point(radiology_scale_x(x), radiology_scale_y(y))


def draw_radiology_textbox(page, layout_key, text, *, fontsize=12.0, min_fontsize=7.0, align=0):
    value = compact_lines(text)
    if not value:
        return

    rect = radiology_rect(layout_key)
    size = fontsize
    while size >= min_fontsize:
        result = page.insert_textbox(
            rect,
            value,
            fontsize=size,
            fontname="helv",
            color=(0, 0, 0),
            align=align,
        )
        if result >= 0:
            return
        size -= 0.5

    page.insert_textbox(
        rect,
        value,
        fontsize=min_fontsize,
        fontname="helv",
        color=(0, 0, 0),
        align=align,
    )


def draw_radiology_checkbox(page, layout_key, checked):
    if not checked:
        return
    point = radiology_point(layout_key)
    page.insert_text(
        fitz.Point(point.x - 4, point.y + 5),
        "X",
        fontsize=12.0,
        fontname="helv",
        color=(0, 0, 0),
    )


def draw_filled_rect(page, rect, *, fill=(1, 1, 1), border=(0.85, 0.85, 0.85)):
    shape = page.new_shape()
    shape.draw_rect(rect)
    shape.finish(color=border, fill=fill, width=0.8)
    shape.commit()


def draw_rect_textbox(page, rect, text, *, fontsize=10.0, min_fontsize=7.0, align=0, rotate=0):
    value = compact_lines(text)
    if not value:
        return

    size = fontsize
    while size >= min_fontsize:
        shape = page.new_shape()
        result = shape.insert_textbox(
            rect,
            value,
            fontsize=size,
            fontname="helv",
            color=(0, 0, 0),
            align=align,
            rotate=rotate,
        )
        if result >= 0:
            shape.commit()
            return
        size -= 0.5

    shape = page.new_shape()
    shape.insert_textbox(
        rect,
        value,
        fontsize=min_fontsize,
        fontname="helv",
        color=(0, 0, 0),
        align=align,
        rotate=rotate,
    )
    shape.commit()


def draw_checkbox_at(page, point, checked, *, fontsize=11.0):
    if not checked:
        return
    x, y = point
    page.insert_text(
        fitz.Point(x - 2, y + 4),
        "X",
        fontsize=fontsize,
        fontname="helv",
        color=(0, 0, 0),
    )


def build_radiology_background():
    doc = fitz.open()
    page = doc.new_page(width=radiology_page_width(), height=radiology_page_height())

    if os.path.exists(RADIOLOGY_TEMPLATE_PDF):
        template_doc = fitz.open(RADIOLOGY_TEMPLATE_PDF)
        page.show_pdf_page(page.rect, template_doc, 0, rotate=270, overlay=False)
        template_doc.close()
        return doc

    if os.path.exists(RADIOLOGY_TEMPLATE_JPEG):
        page.insert_image(page.rect, filename=RADIOLOGY_TEMPLATE_JPEG, overlay=False)
        return doc

    raise FileNotFoundError("Could not find radioltemplate.pdf or radioltemplate.jpeg")


def normalize_radiology_category(value):
    category = clean_text(value).lower().replace("/", "_").replace(" ", "_")
    aliases = {
        "private": "category_private",
        "w_c": "category_wc",
        "wc": "category_wc",
        "pension": "category_pension",
        "vet_aff": "category_vet_aff",
        "vetaff": "category_vet_aff",
        "tac": "category_tac",
    }
    return aliases.get(category, "")


def normalize_yes_no(value):
    option = clean_text(value).lower()
    if option in {"yes", "y", "true"}:
        return "yes"
    if option in {"no", "n", "false"}:
        return "no"
    return ""


def build_ius_patient_sticker_text(patient):
    patient_name = build_patient_name(patient)
    dob = format_display_date(patient.get("dob"))
    urn = clean_text(patient.get("urn"))
    medicare = clean_text(patient.get("medicare_number")) or clean_text(patient.get("medicare"))
    address = clean_text(patient.get("full_address")) or clean_text(patient.get("address"))
    phone = clean_text(patient.get("phone")) or clean_text(patient.get("home_phone")) or clean_text(patient.get("mobile_phone"))
    email = clean_text(patient.get("email"))

    lines = []
    if patient_name:
        lines.append(f"Name: {patient_name}")

    second_line_parts = []
    if dob:
        second_line_parts.append(f"DOB: {dob}")
    if urn:
        second_line_parts.append(f"URN: {urn}")
    if medicare:
        second_line_parts.append(f"Medicare: {medicare}")
    if second_line_parts:
        lines.append("   ".join(second_line_parts))

    if address:
        lines.append(f"Address: {address}")

    contact_parts = []
    if phone:
        contact_parts.append(f"Phone: {phone}")
    if email:
        contact_parts.append(f"Email: {email}")
    if contact_parts:
        lines.append("   ".join(contact_parts))

    return "\n".join(lines)


def get_patient_address_lines(patient):
    address = compact_lines(patient.get("full_address") or patient.get("address") or "")
    if not address:
        return "", ""

    lines = [line.strip() for line in address.splitlines() if line.strip()]
    if len(lines) >= 2:
        return lines[0], " ".join(lines[1:])

    single = lines[0]
    if len(single) <= 42:
        return single, ""

    parts = single.split(",")
    if len(parts) > 1:
        return parts[0].strip(), ", ".join(part.strip() for part in parts[1:] if part.strip())

    return single[:42].rstrip(), single[42:].lstrip()


def iter_word_paragraphs(root):
    yield from root.findall(".//w:p", WORD_NAMESPACE)


def get_word_paragraph_text(paragraph):
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", WORD_NAMESPACE))


def set_word_paragraph_text(paragraph, text):
    text_nodes = paragraph.findall(".//w:t", WORD_NAMESPACE)
    if text_nodes:
        text_nodes[0].text = text
        for node in text_nodes[1:]:
            node.text = ""
        return

    run = paragraph.find("w:r", WORD_NAMESPACE)
    if run is None:
        run = ET.SubElement(paragraph, "{%s}r" % WORD_NAMESPACE["w"])
    text_node = ET.SubElement(run, "{%s}t" % WORD_NAMESPACE["w"])
    if text.startswith(" ") or text.endswith(" "):
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = text


def update_word_paragraph_by_prefix(paragraphs, prefix, value):
    for paragraph in paragraphs:
        current_text = get_word_paragraph_text(paragraph)
        if current_text.startswith(prefix):
            set_word_paragraph_text(paragraph, value)
            return True
    return False


def update_word_checkbox_label(paragraphs, label, checked):
    marker = "☒" if checked else "☐"
    for paragraph in paragraphs:
        current_text = get_word_paragraph_text(paragraph).strip()
        if current_text == label.strip():
            set_word_paragraph_text(paragraph, f"{marker} {label.strip()}")
            return True
    return False


def update_word_contains(paragraphs, needle, value):
    for paragraph in paragraphs:
        current_text = get_word_paragraph_text(paragraph)
        if needle in current_text:
            set_word_paragraph_text(paragraph, value)
            return True
    return False


def update_word_exact_line(paragraphs, expected_text, value):
    for paragraph in paragraphs:
        current_text = get_word_paragraph_text(paragraph).strip()
        if current_text == expected_text.strip():
            set_word_paragraph_text(paragraph, value)
            return True
    return False


def generate_manometry_docx(
    patient,
    request_date,
    tests,
    on_ppi,
    indications,
    other_indication,
    clinical_details,
    last_gastroscopy,
    gastroscopy_results,
    surgery_details,
    referrer,
    medications,
):
    if not os.path.exists(MANOMETRY_TEMPLATE):
        raise FileNotFoundError("Could not find Manometry_Referral_Form.docx")

    with zipfile.ZipFile(MANOMETRY_TEMPLATE, "r") as src:
        archive_entries = {name: src.read(name) for name in src.namelist()}

    document_xml = archive_entries.get("word/document.xml")
    if not document_xml:
        raise ValueError("Manometry template is missing word/document.xml")

    root = ET.fromstring(document_xml)
    paragraphs = list(iter_word_paragraphs(root))

    patient_name = build_patient_name(patient)
    patient_phone = clean_text(patient.get("home_phone") or patient.get("mobile_phone") or patient.get("phone"))
    patient_email = clean_text(patient.get("email"))
    patient_urn = clean_text(patient.get("urn"))
    address_line_1, address_line_2 = get_patient_address_lines(patient)
    referral_date_text = format_display_date(request_date) or format_display_date(datetime.now().strftime("%Y-%m-%d"))

    update_word_exact_line(paragraphs, "URN………………………………………………………………", f"URN {patient_urn}")
    update_word_exact_line(paragraphs, "NAME……………………………………………………………", f"NAME {patient_name}")
    update_word_exact_line(paragraphs, "Date of Birth………………………………………………….", f"Date of Birth {format_display_date(patient.get('dob'))}")
    update_word_exact_line(paragraphs, "Address………………………………………………………….", f"Address {address_line_1}")
    update_word_exact_line(paragraphs, "Phone: ………………………………………………………….", f"Phone: {patient_phone}")
    update_word_exact_line(paragraphs, "Email……………………………………………………………..", f"Email {patient_email}")
    update_word_exact_line(paragraphs, "Date of Referral:  .............................................", f"Date of Referral: {referral_date_text}")
    update_word_exact_line(paragraphs, "………………………………………………………………………", address_line_2)

    test_values = {clean_text(item).lower() for item in (tests or [])}
    update_word_checkbox_label(paragraphs, "OESOPHAGEAL MANOMETRY", "oesophageal_manometry" in test_values)
    update_word_checkbox_label(paragraphs, "24 HOUR PH STUDY (WITH MANOMETRY)", "24_hour_ph_study" in test_values)
    update_word_contains(paragraphs, "ON PPI", f"{'☒' if clean_text(on_ppi).lower() == 'on' else '☐'} ON PPI")
    update_word_contains(paragraphs, "OFF PPI", f"{'☒' if clean_text(on_ppi).lower() == 'off' else '☐'} OFF PPI")
    update_word_checkbox_label(paragraphs, "24 PH IMPEDENCE", "24_ph_impedance" in test_values)

    indication_values = {clean_text(item).lower() for item in (indications or [])}
    indication_labels = {
        "dysphagia": "Dysphagia",
        "consideration_for_poem": "Consideration for POEM",
        "hx_of_oesophageal_spasm": "Hx of Oesophageal Spasm",
        "heart_burn": "Heart burn",
        "erosive_oesophagitis_despite_ppi": "Erosive Oesophagitis despite PPI",
        "hx_of_achalasia": "Hx of Achalasia",
        "regurgitation": "Regurgitation",
        "consideration_for_fundoplication": "Consideration for fundoplication",
        "hx_of_motility_disorder": "Hx of “Motility Disorder”",
        "atypical_chest_pain": "Atypical Chest pain",
        "consideration_for_hellers_myotomy": "Consideration for Hellers Myotomy",
        "post_fundoplication": "POST Fundoplication",
    }
    for key, label in indication_labels.items():
        update_word_contains(paragraphs, label, f"{'☒' if key in indication_values else '☐'} {label}")

    other_text = clean_text(other_indication)
    update_word_contains(
        paragraphs,
        "Other",
        f"{'☒' if other_text else '☐'} Other {other_text}".rstrip(),
    )
    update_word_contains(
        paragraphs,
        "CLINICAL DETAILS:",
        f"CLINICAL DETAILS: {compact_single_line(clinical_details)}",
    )
    update_word_contains(
        paragraphs,
        "When was the last Gastroscopy?",
        f"When was the last Gastroscopy? {clean_text(last_gastroscopy)}",
    )
    update_word_contains(
        paragraphs,
        "What were the results?",
        f"What were the results? {compact_single_line(gastroscopy_results)}",
    )
    update_word_contains(
        paragraphs,
        "Prior Gastric or Oesophageal Surgery details:",
        f"Prior Gastric or Oesophageal Surgery details: {compact_single_line(surgery_details)}",
    )

    referrer_name = clean_text(referrer.get("name")) or "Dr Chamara Basnayake"
    referrer_address = compact_single_line(referrer.get("clinic") or referrer.get("address"))
    referrer_phone = clean_text(referrer.get("phone"))
    referrer_department = clean_text(referrer.get("department")) or "Gastroenterology"
    referrer_provider_number = clean_text(referrer.get("provider_number"))
    referrer_fax = clean_text(referrer.get("fax"))
    referrer_signature_date = clean_text(referrer.get("signature_date")) or referral_date_text

    update_word_contains(paragraphs, "NAME:", f"NAME: {referrer_name}")
    update_word_contains(paragraphs, "ADDRESS", f"ADDRESS {referrer_address}")
    update_word_contains(paragraphs, "PHONE", f"PHONE {referrer_phone}")
    update_word_contains(paragraphs, "DEPARTMENT", f"DEPARTMENT {referrer_department}")
    update_word_contains(paragraphs, "PROVIDER NUMBER", f"PROVIDER NUMBER {referrer_provider_number}")
    update_word_contains(paragraphs, "FAX", f"FAX {referrer_fax}")
    update_word_contains(paragraphs, "SIGNATURE & DATE", f"SIGNATURE & DATE {referrer_signature_date}")
    update_word_contains(paragraphs, "Medication name:", f"Medication name: {compact_single_line(medications)}")

    archive_entries["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    os.makedirs(GENERATED_MANOMETRY_DIR, exist_ok=True)
    output_name = build_safe_filename(patient_name, "manometry", "docx")
    output_path = os.path.join(GENERATED_MANOMETRY_DIR, output_name)

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as dest:
        for name, content in archive_entries.items():
            dest.writestr(name, content)

    return output_name, output_path


def build_copy_reports_text(copy_doctor, fallback_name):
    lines = []
    if isinstance(copy_doctor, dict):
        name = clean_text(copy_doctor.get("name"))
        clinic = clean_text(copy_doctor.get("clinic"))
        if name:
            lines.append(name)
        if clinic:
            lines.append(clinic)
    else:
        name = ""

    if not lines and fallback_name:
        lines.append(clean_text(fallback_name))
    return "\n".join(line for line in lines if line)


def generate_pathology_pdf(patient, tests_required, clinical_notes, request_date, copy_doctor, copy_reports_to):
    if not os.path.exists(PATHOLOGY_TEMPLATE):
        raise FileNotFoundError("Could not find pathology.pdf")

    doc = fitz.open(PATHOLOGY_TEMPLATE)
    page = doc[0]

    firstname, surname = split_patient_name(patient)
    surname_or_name = surname or clean_text(patient.get("name"))
    full_address = clean_text(patient.get("full_address")) or clean_text(patient.get("address"))
    sex = clean_text(patient.get("sex"))
    dob = format_display_date(patient.get("dob"))
    home_phone = clean_text(patient.get("home_phone")) or clean_text(patient.get("mobile_phone")) or clean_text(patient.get("phone"))
    business_phone = ""
    request_date_text = format_display_date(request_date) or datetime.now().strftime("%d/%m/%Y")
    copy_reports_text = build_copy_reports_text(copy_doctor, copy_reports_to or patient.get("referring_doctor"))

    top_last_address = "\n".join(part for part in [surname_or_name, full_address] if part)
    draw_textbox(page, "top_last_address", top_last_address, fontsize=8.3)
    draw_textbox(page, "top_given_names", firstname, fontsize=8.8)
    draw_textbox(page, "top_sex", sex, fontsize=9.0, min_fontsize=8.0, align=1)
    draw_textbox(page, "top_dob", dob, fontsize=8.8, min_fontsize=8.0, align=1)
    draw_textbox(page, "top_phone_home", home_phone, fontsize=8.0)
    draw_textbox(page, "top_phone_bus", business_phone, fontsize=8.0)
    draw_textbox(page, "top_tests", tests_required, fontsize=10.0, min_fontsize=7.0)
    draw_textbox(page, "top_notes", clinical_notes, fontsize=8.0, min_fontsize=6.0)
    draw_textbox(page, "top_copy_reports", copy_reports_text, fontsize=8.0, min_fontsize=6.0)
    draw_text(page, "top_request_date_point", request_date_text, fontsize=9.0)

    draw_textbox(page, "bottom_surname", surname_or_name, fontsize=8.5)
    draw_textbox(page, "bottom_given_names", firstname, fontsize=8.5)
    draw_textbox(page, "bottom_sex", sex, fontsize=9.0, min_fontsize=8.0, align=1)
    draw_textbox(page, "bottom_dob", dob, fontsize=8.5, min_fontsize=8.0, align=1)
    draw_textbox(page, "bottom_address", full_address, fontsize=8.0, min_fontsize=6.5)
    draw_textbox(page, "bottom_phone_home", home_phone, fontsize=8.0)
    draw_textbox(page, "bottom_phone_bus", business_phone, fontsize=8.0)
    draw_textbox(page, "bottom_tests", tests_required, fontsize=9.2, min_fontsize=6.5)

    os.makedirs(GENERATED_PATHOLOGY_DIR, exist_ok=True)
    patient_name_safe = re.sub(r"[^\w\s-]", "", clean_text(patient.get("name")) or "Unknown").strip().replace(" ", "_")
    output_name = f"{patient_name_safe}_pathology_{datetime.now().strftime('%Y-%m-%d')}.pdf"
    output_path = os.path.join(GENERATED_PATHOLOGY_DIR, output_name)
    doc.save(output_path)
    doc.close()
    return output_name, output_path


def generate_radiology_pdf(patient, request_for, clinical_details, request_date, patient_category):
    doc = build_radiology_background()
    page = doc[0]

    patient_name = build_patient_name(patient)
    dob = format_display_date(patient.get("dob"))
    address = clean_text(patient.get("full_address")) or clean_text(patient.get("address"))
    phone = (
        clean_text(patient.get("phone"))
        or clean_text(patient.get("home_phone"))
        or clean_text(patient.get("mobile_phone"))
    )
    medicare = clean_text(patient.get("medicare_number")) or clean_text(patient.get("medicare"))
    request_date_text = format_display_date(request_date) or datetime.now().strftime("%d/%m/%Y")
    category_key = normalize_radiology_category(patient_category)

    draw_radiology_textbox(page, "name", patient_name, fontsize=13.0, min_fontsize=8.0)
    draw_radiology_textbox(page, "dob", dob, fontsize=13.0, min_fontsize=8.0)
    draw_radiology_textbox(page, "address", address, fontsize=12.0, min_fontsize=7.0)
    draw_radiology_textbox(page, "medicare", medicare, fontsize=12.0, min_fontsize=8.0)
    draw_radiology_textbox(page, "phone", phone, fontsize=12.0, min_fontsize=8.0)
    draw_radiology_textbox(page, "request_for", request_for, fontsize=14.0, min_fontsize=8.0)
    draw_radiology_textbox(page, "clinical_details", clinical_details, fontsize=12.0, min_fontsize=7.0)
    draw_radiology_textbox(page, "date", request_date_text, fontsize=12.0, min_fontsize=8.0)
    draw_radiology_checkbox(page, category_key, bool(category_key))

    os.makedirs(GENERATED_RADIOLOGY_DIR, exist_ok=True)
    patient_name_safe = re.sub(r"[^\w\s-]", "", patient_name or "Unknown").strip().replace(" ", "_")
    output_name = f"{patient_name_safe}_radiology_{datetime.now().strftime('%Y-%m-%d')}.pdf"
    output_path = os.path.join(GENERATED_RADIOLOGY_DIR, output_name)
    doc.save(output_path)
    doc.close()
    return output_name, output_path


def generate_ius_pdf(
    patient,
    request_date,
    urgent_referral,
    reasons,
    medical_history,
    abnormal_markers,
    abnormal_markers_result,
    faecal_calprotectin,
    faecal_calprotectin_result,
    pregnant,
    gestation_weeks,
):
    if not os.path.exists(IUS_TEMPLATE):
        raise FileNotFoundError("Could not find IUS_template.pdf")

    doc = fitz.open(IUS_TEMPLATE)
    page = doc[0]

    sticker_rect = fitz.Rect(*IUS_LAYOUT["patient_sticker"])
    draw_filled_rect(page, sticker_rect)
    draw_rect_textbox(
        page,
        fitz.Rect(sticker_rect.x0 + 10, sticker_rect.y0 + 8, sticker_rect.x1 - 10, sticker_rect.y1 - 8),
        build_ius_patient_sticker_text(patient),
        fontsize=9.6,
        min_fontsize=7.0,
    )

    request_date_text = format_display_date(request_date) or datetime.now().strftime("%d/%m/%Y")
    draw_rect_textbox(page, fitz.Rect(*IUS_LAYOUT["date"]), request_date_text, fontsize=10.0, min_fontsize=8.0)

    urgent_option = normalize_yes_no(urgent_referral)
    draw_checkbox_at(page, IUS_LAYOUT["urgent_yes"], urgent_option == "yes")
    draw_checkbox_at(page, IUS_LAYOUT["urgent_no"], urgent_option == "no")

    reason_values = {clean_text(reason).lower() for reason in (reasons or [])}
    draw_checkbox_at(page, IUS_LAYOUT["reason_symptoms"], "symptoms" in reason_values)
    draw_checkbox_at(page, IUS_LAYOUT["reason_ibd"], "ibd" in reason_values)
    draw_checkbox_at(page, IUS_LAYOUT["reason_incomplete_colonoscopy"], "incomplete_colonoscopy" in reason_values)

    draw_rect_textbox(page, fitz.Rect(*IUS_LAYOUT["medical_history"]), medical_history, fontsize=10.0, min_fontsize=7.0)
    abnormal_markers_line = " | ".join(part for part in [abnormal_markers, abnormal_markers_result] if part)
    draw_rect_textbox(
        page,
        fitz.Rect(*IUS_LAYOUT["abnormal_markers_result"]),
        abnormal_markers_line,
        fontsize=9.4,
        min_fontsize=7.0,
    )
    faecal_calprotectin_line = " | ".join(part for part in [faecal_calprotectin, faecal_calprotectin_result] if part)
    draw_rect_textbox(
        page,
        fitz.Rect(*IUS_LAYOUT["faecal_calprotectin_result"]),
        faecal_calprotectin_line,
        fontsize=9.4,
        min_fontsize=7.0,
    )

    pregnant_option = normalize_yes_no(pregnant)
    draw_checkbox_at(page, IUS_LAYOUT["pregnant_yes"], pregnant_option == "yes")
    draw_checkbox_at(page, IUS_LAYOUT["pregnant_no"], pregnant_option == "no")
    draw_rect_textbox(page, fitz.Rect(*IUS_LAYOUT["gestation_weeks"]), gestation_weeks, fontsize=9.4, min_fontsize=7.0)

    os.makedirs(GENERATED_IUS_DIR, exist_ok=True)
    patient_name_safe = re.sub(r"[^\w\s-]", "", build_patient_name(patient) or "Unknown").strip().replace(" ", "_")
    output_name = f"{patient_name_safe}_ius_{datetime.now().strftime('%Y-%m-%d')}.pdf"
    output_path = os.path.join(GENERATED_IUS_DIR, output_name)
    doc.save(output_path)
    doc.close()
    return output_name, output_path


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/pathology")
def pathology():
    return render_template("pathology.html")


@app.route("/radiology")
def radiology():
    return render_template("radiology.html")


@app.route("/ius")
def ius():
    return render_template("ius.html")


@app.route("/manometry")
def manometry():
    return render_template("manometry.html")


@app.route("/generated-pathology/<path:filename>")
def generated_pathology_file(filename):
    safe_name = os.path.basename(filename)
    path = os.path.join(GENERATED_PATHOLOGY_DIR, safe_name)
    if not os.path.exists(path):
        return jsonify({"error": "File not found"}), 404
    return send_file(path, mimetype="application/pdf", as_attachment=False, download_name=safe_name)


@app.route("/generated-radiology/<path:filename>")
def generated_radiology_file(filename):
    safe_name = os.path.basename(filename)
    path = os.path.join(GENERATED_RADIOLOGY_DIR, safe_name)
    if not os.path.exists(path):
        return jsonify({"error": "File not found"}), 404
    return send_file(path, mimetype="application/pdf", as_attachment=False, download_name=safe_name)


@app.route("/generated-ius/<path:filename>")
def generated_ius_file(filename):
    safe_name = os.path.basename(filename)
    path = os.path.join(GENERATED_IUS_DIR, safe_name)
    if not os.path.exists(path):
        return jsonify({"error": "File not found"}), 404
    return send_file(path, mimetype="application/pdf", as_attachment=False, download_name=safe_name)


@app.route("/generated-manometry/<path:filename>")
def generated_manometry_file(filename):
    safe_name = os.path.basename(filename)
    path = os.path.join(GENERATED_MANOMETRY_DIR, safe_name)
    if not os.path.exists(path):
        return jsonify({"error": "File not found"}), 404
    return send_file(
        path,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=False,
        download_name=safe_name,
    )


@app.route("/api/data")
def api_data():
    load_data()
    return jsonify({"patients": data_cache["patients"], "doctors": data_cache["doctors"]})


@app.route("/api/templates")
def api_templates():
    if not os.path.exists(TEMPLATE_DIR):
        return jsonify([])
    templates = [name for name in os.listdir(TEMPLATE_DIR) if name.lower().endswith(".pdf")]
    return jsonify(sorted(templates))


@app.route("/api/generate", methods=["POST"])
def api_generate():
    data = request.get_json(silent=True) or {}
    template_name = data.get("template")
    patient = data.get("patient", {})
    referrer = data.get("referrer", {})
    cc_doctors = data.get("cc_doctors", [])
    body_text = data.get("body_text", "")
    pathology_text = data.get("pathology_text", "")

    if not template_name:
        return jsonify({"error": "No template selected"}), 400

    template_path = os.path.join(TEMPLATE_DIR, template_name)
    if not os.path.exists(template_path):
        return jsonify({"error": "Template not found"}), 404

    overlay_pdf_stream = io.BytesIO()
    doc = SimpleDocTemplate(
        overlay_pdf_stream,
        pagesize=A4,
        rightMargin=50,
        leftMargin=50,
        topMargin=130,
        bottomMargin=160,
    )

    styles = getSampleStyleSheet()
    style_normal = styles["Normal"]
    style_normal.fontSize = 10
    style_normal.leading = 13
    style_normal.fontName = "Helvetica"
    styles.add(
        ParagraphStyle(
            name="PatientBlock",
            parent=style_normal,
            fontName="Helvetica-Bold",
            leading=12,
            spaceAfter=1,
        )
    )
    styles.add(ParagraphStyle(name="BulletList", parent=style_normal, leftIndent=24, bulletIndent=12))

    story = [
        Paragraph(datetime.now().strftime("%d %B %Y"), style_normal),
        Spacer(1, 10),
        Paragraph(referrer.get("name", "To whom it may concern"), style_normal),
    ]

    if referrer.get("clinic"):
        story.append(Paragraph(referrer.get("clinic", ""), style_normal))
    story.append(Spacer(1, 10))

    ref_name = clean_text(referrer.get("name"))
    parts = ref_name.split()
    dear_name = "Colleague"
    if len(parts) >= 2 and parts[0].lower() in {"dr", "dr.", "mr", "mrs", "ms", "a/prof", "prof", "doctor"}:
        dear_name = f"{parts[0]} {parts[-1]}"
    elif ref_name:
        dear_name = ref_name

    story.extend(
        [
            Paragraph(f"Dear {dear_name},", style_normal),
            Spacer(1, 10),
        ]
    )
    for line in build_letter_patient_detail_lines(patient):
        story.append(Paragraph(f"<b>{line}</b>", styles["PatientBlock"]))
    story.append(Spacer(1, 10))

    story.extend(html_to_story(body_text, style_normal, styles["BulletList"]))
    story.extend(
        [
            Spacer(1, 20),
            Paragraph("Yours sincerely,", style_normal),
            Spacer(1, 22),
            Paragraph("<b>A/Prof Chamara Basnayake</b>", style_normal),
            Paragraph("Gastroenterologist", style_normal),
        ]
    )

    if cc_doctors:
        names = ", ".join(doc.get("name", "") for doc in cc_doctors if doc.get("name"))
        if names:
            story.extend([Spacer(1, 14), Paragraph(f"CC: {names}", style_normal)])

    pathology_plain = BeautifulSoup(pathology_text, "html.parser").get_text().strip() if pathology_text else ""
    if pathology_plain:
        if "PathologyHeading" not in styles:
            styles.add(
                ParagraphStyle(
                    name="PathologyHeading",
                    parent=style_normal,
                    fontName="Helvetica-Bold",
                    fontSize=10,
                    spaceBefore=0,
                    spaceAfter=2,
                )
            )
        story.extend(
            [
                Spacer(1, 16),
                Paragraph("Investigation Results:", styles["PathologyHeading"]),
                Spacer(1, 4),
            ]
        )
        story.extend(html_to_story(pathology_text, style_normal, styles["BulletList"]))

    def draw_white_background(canvas_obj, _doc):
        canvas_obj.saveState()
        canvas_obj.setFillColorRGB(1, 1, 1)
        canvas_obj.rect(0, 150, A4[0], 590, stroke=0, fill=1)
        canvas_obj.restoreState()

    doc.build(story, onFirstPage=draw_white_background, onLaterPages=draw_white_background)
    overlay_pdf_stream.seek(0)

    doc_generated = fitz.open(stream=overlay_pdf_stream, filetype="pdf")
    doc_template = fitz.open(template_path)
    for page in doc_generated:
        page.show_pdf_page(page.rect, doc_template, 0, overlay=False)

    os.makedirs(GENERATED_LETTERS_DIR, exist_ok=True)
    patient_name_safe = re.sub(r"[^\w\s-]", "", clean_text(patient.get("name")) or "Unknown").strip().replace(" ", "_")
    date_str = datetime.now().strftime("%Y-%m-%d")
    output_path = os.path.join(GENERATED_LETTERS_DIR, f"{patient_name_safe}_{date_str}.pdf")
    doc_generated.save(output_path)
    doc_generated.close()
    doc_template.close()

    doc_list = []
    if referrer.get("email"):
        doc_list.append(referrer.get("email"))
    for cc_doctor in cc_doctors:
        if cc_doctor.get("email"):
            doc_list.append(cc_doctor.get("email"))

    cc_emails = [PRACTICE_CC]
    to_email = doc_list[0] if doc_list else ""
    if to_email:
        doc_list.pop(0)
    cc_emails.extend(doc_list)

    try:
        write_outlook_draft(
            f"Patient Letter: {patient.get('name', '')}",
            output_path,
            to_recipients=[to_email],
            cc_recipients=cc_emails,
            sender_email=SENDER_EMAIL,
        )
    except Exception as exc:
        print(f"Error opening Outlook: {exc}")

    return jsonify({"success": True, "message": "PDF Generated!"})


@app.route("/api/generate-pathology", methods=["POST"])
def api_generate_pathology():
    load_data()
    data = request.get_json(silent=True) or {}
    patient = data.get("patient") or {}
    tests_required = clean_text(data.get("tests_required"))
    clinical_notes = clean_text(data.get("clinical_notes"))
    request_date = data.get("request_date")
    copy_doctor = data.get("copy_doctor") or {}
    copy_reports_to = data.get("copy_reports_to")

    if not patient:
        return jsonify({"error": "Please select a patient."}), 400
    if not tests_required:
        return jsonify({"error": "Please enter the requested tests."}), 400

    try:
        output_name, output_path = generate_pathology_pdf(
            patient,
            tests_required=tests_required,
            clinical_notes=clinical_notes,
            request_date=request_date,
            copy_doctor=copy_doctor,
            copy_reports_to=copy_reports_to,
        )
    except Exception as exc:
        print("Error generating pathology request:", exc)
        return jsonify({"error": str(exc)}), 500

    return jsonify(
        {
            "success": True,
            "message": "Pathology request generated.",
            "url": url_for("generated_pathology_file", filename=output_name),
            "output_path": output_path,
        }
    )


@app.route("/api/generate-radiology", methods=["POST"])
def api_generate_radiology():
    load_data()
    data = request.get_json(silent=True) or {}
    patient = data.get("patient") or {}
    request_for = clean_text(data.get("request_for"))
    clinical_details = clean_text(data.get("clinical_details"))
    request_date = data.get("request_date")
    patient_category = data.get("patient_category")

    if not build_patient_name(patient):
        return jsonify({"error": "Please enter the patient name."}), 400
    if not request_for:
        return jsonify({"error": "Please enter what the scan is requested for."}), 400

    try:
        output_name, output_path = generate_radiology_pdf(
            patient,
            request_for=request_for,
            clinical_details=clinical_details,
            request_date=request_date,
            patient_category=patient_category,
        )
    except Exception as exc:
        print("Error generating radiology request:", exc)
        return jsonify({"error": str(exc)}), 500

    return jsonify(
        {
            "success": True,
            "message": "Radiology request generated.",
            "url": url_for("generated_radiology_file", filename=output_name),
            "output_path": output_path,
        }
    )


@app.route("/api/generate-ius", methods=["POST"])
def api_generate_ius():
    load_data()
    data = request.get_json(silent=True) or {}
    patient = data.get("patient") or {}
    request_date = data.get("request_date")
    urgent_referral = data.get("urgent_referral")
    reasons = data.get("reasons") or []
    medical_history = clean_text(data.get("medical_history"))
    abnormal_markers = clean_text(data.get("abnormal_markers"))
    abnormal_markers_result = clean_text(data.get("abnormal_markers_result"))
    faecal_calprotectin = clean_text(data.get("faecal_calprotectin"))
    faecal_calprotectin_result = clean_text(data.get("faecal_calprotectin_result"))
    pregnant = data.get("pregnant")
    gestation_weeks = clean_text(data.get("gestation_weeks"))

    if not build_patient_name(patient):
        return jsonify({"error": "Please enter the patient name."}), 400

    try:
        output_name, output_path = generate_ius_pdf(
            patient,
            request_date=request_date,
            urgent_referral=urgent_referral,
            reasons=reasons,
            medical_history=medical_history,
            abnormal_markers=abnormal_markers,
            abnormal_markers_result=abnormal_markers_result,
            faecal_calprotectin=faecal_calprotectin,
            faecal_calprotectin_result=faecal_calprotectin_result,
            pregnant=pregnant,
            gestation_weeks=gestation_weeks,
        )
    except Exception as exc:
        print("Error generating IUS request:", exc)
        return jsonify({"error": str(exc)}), 500

    return jsonify(
        {
            "success": True,
            "message": "IUS request generated.",
            "url": url_for("generated_ius_file", filename=output_name),
            "output_path": output_path,
        }
    )


@app.route("/api/generate-manometry", methods=["POST"])
def api_generate_manometry():
    load_data()
    data = request.get_json(silent=True) or {}
    patient = data.get("patient") or {}
    request_date = data.get("request_date")
    tests = data.get("tests") or []
    on_ppi = data.get("on_ppi")
    indications = data.get("indications") or []
    other_indication = clean_text(data.get("other_indication"))
    clinical_details = clean_text(data.get("clinical_details"))
    last_gastroscopy = clean_text(data.get("last_gastroscopy"))
    gastroscopy_results = clean_text(data.get("gastroscopy_results"))
    surgery_details = clean_text(data.get("surgery_details"))
    medications = clean_text(data.get("medications"))
    referrer = data.get("referrer") or {}

    if not build_patient_name(patient):
        return jsonify({"error": "Please enter the patient name."}), 400

    try:
        output_name, output_path = generate_manometry_docx(
            patient,
            request_date=request_date,
            tests=tests,
            on_ppi=on_ppi,
            indications=indications,
            other_indication=other_indication,
            clinical_details=clinical_details,
            last_gastroscopy=last_gastroscopy,
            gastroscopy_results=gastroscopy_results,
            surgery_details=surgery_details,
            referrer=referrer,
            medications=medications,
        )
    except Exception as exc:
        print("Error generating manometry referral:", exc)
        return jsonify({"error": str(exc)}), 500

    return jsonify(
        {
            "success": True,
            "message": "Manometry referral generated.",
            "url": url_for("generated_manometry_file", filename=output_name),
            "output_path": output_path,
        }
    )


import urllib.request
import urllib.error

def load_env():
    env_path = os.path.join(ROOT_DIR, ".env")
    if os.path.exists(env_path):
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" in line:
                    key, val = line.split("=", 1)
                    key = key.strip()
                    val = val.strip().strip("'\"")
                    os.environ[key] = val

def run_swift_ocr(image_path):
    binary_path = os.path.join(BASE_DIR, "ocr")
    swift_script = os.path.join(BASE_DIR, "ocr.swift")
    try:
        if os.path.exists(binary_path) and os.access(binary_path, os.X_OK):
            cmd = [binary_path, image_path]
        elif os.path.exists(swift_script):
            cmd = ["swift", swift_script, image_path]
        else:
            print("OCR script not found at", swift_script)
            return {}
        result = subprocess.run(cmd, capture_output=True, text=True, check=True)
        return json.loads(result.stdout)
    except Exception as e:
        print("Swift OCR failed:", e)
        return {}


def parse_genie_demographics(ocr_data):
    if not isinstance(ocr_data, dict):
        return None
    items = ocr_data.get("items", [])
    raw_text = ocr_data.get("raw_text", "")

    labels_present = sum(1 for label in ["First Name:", "Surname:", "Address 1:", "DOB:", "Medicare No:"] if label in raw_text)
    if labels_present < 2:
        return None

    def get_right_of(label, max_ydiff=0.025, max_x=None):
        matches = [it for it in items if it["text"].startswith(label)]
        if not matches:
            return ""
        lbl = matches[0]
        lbl_right = lbl["x"] + lbl["w"]
        lbl_y = lbl["y"]
        limit_x = max_x if max_x is not None else (lbl_right + 0.22)
        candidates = [
            it for it in items
            if abs(it["y"] - lbl_y) <= max_ydiff
            and it["x"] >= (lbl["x"] + lbl["w"] - 0.01)
            and it["x"] <= limit_x
            and it["text"] != lbl["text"]
        ]
        candidates.sort(key=lambda it: it["x"])
        if candidates:
            val = candidates[0]["text"].strip()
            if val.endswith(":"):
                return ""
            return val
        return ""

    first_name = get_right_of("First Name:", max_x=0.29)
    surname = get_right_of("Surname:", max_x=0.29)
    title = get_right_of("Title:", max_x=0.29)
    address1 = get_right_of("Address 1:", max_x=0.54)
    address2 = get_right_of("Address 2:", max_x=0.54)
    suburb = get_right_of("Suburb:", max_x=0.54)
    state = get_right_of("State:", max_x=0.45)
    dob_raw = get_right_of("DOB:", max_x=0.72)

    full_name = f"{first_name} {surname}".strip()

    postcode = ""
    state_matches = [it for it in items if it["text"].startswith("State:")]
    if state_matches:
        lbl = state_matches[0]
        same_line = [it for it in items if abs(it["y"] - lbl["y"]) <= 0.025 and it["x"] > lbl["x"]]
        for it in same_line:
            m = re.search(r"\b(\d{4})\b", it["text"])
            if m:
                postcode = m.group(1)
                break

    address_parts = [p for p in [address1, address2, suburb, state, postcode] if p]
    full_address = ", ".join(address_parts) if address_parts else ""

    dob = ""
    if dob_raw:
        m = re.search(r"(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})", dob_raw)
        if m:
            d, mon, y = m.groups()
            dob = f"{y}-{int(mon):02d}-{int(d):02d}"

    phone = ""
    for it in items:
        m = re.search(r"(?:04\d{2}[\s.-]?\d{3}[\s.-]?\d{3}|0\d{9})", it["text"])
        if m:
            phone = m.group(0).replace(" ", "").replace("-", "")
            break

    email = ""
    for it in items:
        m = re.search(r"([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)", it["text"])
        if m:
            email = m.group(1)
            break

    medicare = get_right_of("Medicare No:")
    if medicare:
        m = re.search(r"(\d{4}\s?\d{5}\s?\d|\d{10})", medicare)
        if m:
            medicare = m.group(1)
    else:
        for it in items:
            m = re.search(r"\b(\d{4}\s\d{5}\s\d)\b", it["text"])
            if m:
                medicare = m.group(1)
                break

    medicare_ref = ""
    for it in items:
        m = re.search(r"Ref:\s*(\d+)", it["text"])
        if m:
            medicare_ref = m.group(1)
            break

    account_type = get_right_of("Account Type:", max_x=0.30)
    private_fund = get_right_of("Private Fund:", max_x=0.30)
    private_health = private_fund if (private_fund and not private_fund.startswith("Select") and not private_fund.startswith("UPI")) else account_type

    gp = ""
    gp_matches = [it for it in items if it["text"].startswith("Usual GP:")]
    if gp_matches:
        lbl = gp_matches[0]
        same_line = [it for it in items if abs(it["y"] - lbl["y"]) <= 0.03 and it["x"] > lbl["x"] and "Select" not in it["text"]]
        if same_line:
            gp = same_line[-1]["text"].strip()

    urn = get_right_of("Chart:")

    return {
        "name": full_name,
        "firstname": first_name,
        "surname": surname,
        "title": title,
        "dob": dob,
        "address": full_address,
        "street": address1,
        "suburb": suburb,
        "state": state,
        "postcode": postcode,
        "phone": phone,
        "mobile_phone": phone,
        "email": email,
        "medicare": medicare,
        "medicare_number": medicare,
        "medicare_ref_number": medicare_ref,
        "private_health": private_health,
        "urn": urn,
        "referring_doctor": gp,
        "gp": gp,
        "provider": "Apple Vision (Native)",
    }


def call_gemini_api(ocr_text, api_key):
    prompt = f"""You are a precise data extraction tool. Extract the patient demographics and referral details from the following raw OCR or referral text.
Format the output as a JSON object with the following keys:
- "name" (Full Name of the patient)
- "dob" (Date of birth in YYYY-MM-DD format. If only DD/MM/YYYY or similar is found, convert to YYYY-MM-DD. If not present, use null)
- "urn" (URN / Patient ID / MRN, use null if not present)
- "medicare" (Medicare number, use null if not present)
- "private_health" (Private Health Insurance provider or membership/policy number, use null if not present)
- "phone" (Phone number / contact number, use null if not present)
- "email" (Email address, use null if not present)
- "address" (Full address, use null if not present)
- "request_for" (Requested scan, imaging, procedure, or investigation if mentioned, e.g. "CT Abdomen and Pelvis", use null if not present)
- "clinical_details" (Clinical indication, reason for referral/scan, symptoms, history, or provisional diagnosis, use null if not present)

OCR Text:
---
{ocr_text}
---

Return ONLY the raw JSON object. Do not include markdown code block formatting (like ```json ... ```) or any other text.
"""
    model = os.environ.get("GEMINI_MODEL", "gemini-3.1-flash-lite")
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
    payload = {
        "contents": [{
            "parts": [{"text": prompt}]
        }],
        "generationConfig": {
            "responseMimeType": "application/json"
        }
    }
    
    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        },
        method="POST"
    )
    
    try:
        with urllib.request.urlopen(req, timeout=15) as res:
            response_data = json.loads(res.read().decode("utf-8"))
            text_content = response_data["candidates"][0]["content"]["parts"][0]["text"].strip()
            if text_content.startswith("```"):
                text_content = re.sub(r"^```(?:json)?\n|```$", "", text_content, flags=re.MULTILINE).strip()
            return json.loads(text_content)
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8")
        raise Exception(f"Gemini API HTTP {e.code}: {error_body}")

def call_groq_api(ocr_text, api_key):
    prompt = f"""You are a precise data extraction tool. Extract the patient demographics and referral details from the following raw OCR or referral text.
Format the output as a JSON object with the following keys:
- "name" (Full Name of the patient)
- "dob" (Date of birth in YYYY-MM-DD format. If only DD/MM/YYYY or similar is found, convert to YYYY-MM-DD. If not present, use null)
- "urn" (URN / Patient ID / MRN, use null if not present)
- "medicare" (Medicare number, use null if not present)
- "private_health" (Private Health Insurance provider or membership/policy number, use null if not present)
- "phone" (Phone number / contact number, use null if not present)
- "email" (Email address, use null if not present)
- "address" (Full address, use null if not present)
- "request_for" (Requested scan, imaging, procedure, or investigation if mentioned, e.g. "CT Abdomen and Pelvis", use null if not present)
- "clinical_details" (Clinical indication, reason for referral/scan, symptoms, history, or provisional diagnosis, use null if not present)

OCR Text:
---
{ocr_text}
---

Return ONLY the raw JSON object. Do not include markdown code block formatting (like ```json ... ```) or any other text.
"""
    model = os.environ.get("GROQ_MODEL", "meta-llama/llama-4-scout-17b-16e-instruct")
    url = "https://api.groq.com/openai/v1/chat/completions"
    payload = {
        "model": model,
        "messages": [
            {"role": "user", "content": prompt}
        ],
        "response_format": {"type": "json_object"}
    }
    
    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        },
        method="POST"
    )
    
    try:
        with urllib.request.urlopen(req, timeout=15) as res:
            response_data = json.loads(res.read().decode("utf-8"))
            text_content = response_data["choices"][0]["message"]["content"].strip()
            if text_content.startswith("```"):
                text_content = re.sub(r"^```(?:json)?\n|```$", "", text_content, flags=re.MULTILINE).strip()
            return json.loads(text_content)
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8")
        raise Exception(f"Groq API HTTP {e.code}: {error_body}")


def extract_demographics_from_text(ocr_text):
    gemini_key = os.environ.get("GEMINI_API_KEY")
    groq_key = os.environ.get("GROQ_API_KEY")
    
    errors = []
    
    if gemini_key:
        try:
            print("Attempting extraction with Gemini...")
            result = call_gemini_api(ocr_text, gemini_key)
            result["provider"] = "Gemini"
            return result
        except Exception as e:
            err_msg = f"Gemini extraction failed: {str(e)}"
            print(err_msg)
            errors.append(err_msg)
            
    if groq_key:
        try:
            print("Attempting extraction with Groq...")
            result = call_groq_api(ocr_text, groq_key)
            result["provider"] = "Groq"
            return result
        except Exception as e:
            err_msg = f"Groq extraction failed: {str(e)}"
            print(err_msg)
            errors.append(err_msg)
            
    return {
        "error": "Failed to extract demographics using LLM providers.",
        "details": errors or ["No API keys configured. Please add GEMINI_API_KEY or GROQ_API_KEY to your .env file."]
    }

@app.route("/api/ocr-capture")
def api_ocr_capture():
    load_env()
    temp_img = "/tmp/referral_screenshot.png"
    if os.path.exists(temp_img):
        try:
            os.remove(temp_img)
        except Exception:
            pass
            
    try:
        # screencapture -i triggers interactive crop interface on macOS
        subprocess.run(["screencapture", "-i", temp_img], check=True)
    except Exception as e:
        return jsonify({"error": "Screen capture cancelled or failed: " + str(e)}), 400
        
    if not os.path.exists(temp_img):
        return jsonify({"error": "No screenshot was taken"}), 400
        
    ocr_data = run_swift_ocr(temp_img)
    if not ocr_data or not ocr_data.get("raw_text"):
        return jsonify({"error": "OCR failed to extract any text from the screenshot"}), 500
        
    parsed = parse_genie_demographics(ocr_data)
    if parsed and parsed.get("name"):
        return jsonify(parsed)

    result = extract_demographics_from_text(ocr_data.get("raw_text", ""))
    if "error" in result:
        return jsonify(result), 500
    return jsonify(result)

@app.route("/api/ocr-upload", methods=["POST"])
def api_ocr_upload():
    load_env()
    if "image" not in request.files:
        return jsonify({"error": "No image file provided"}), 400
    file = request.files["image"]
    if file.filename == "":
        return jsonify({"error": "No selected file"}), 400
        
    temp_img = "/tmp/referral_upload.png"
    try:
        file.save(temp_img)
    except Exception as e:
        return jsonify({"error": "Failed to save file: " + str(e)}), 500
    
    ocr_data = run_swift_ocr(temp_img)
    if not ocr_data or not ocr_data.get("raw_text"):
        return jsonify({"error": "OCR failed to extract any text from the image"}), 500
        
    parsed = parse_genie_demographics(ocr_data)
    if parsed and parsed.get("name"):
        return jsonify(parsed)

    result = extract_demographics_from_text(ocr_data.get("raw_text", ""))
    if "error" in result:
        return jsonify(result), 500
    return jsonify(result)

@app.route("/api/ocr-text", methods=["POST"])
def api_ocr_text():
    load_env()
    data = request.get_json(silent=True) or {}
    text = data.get("text", "").strip()
    if not text:
        return jsonify({"error": "No text provided"}), 400
        
    result = extract_demographics_from_text(text)
    if "error" in result:
        return jsonify(result), 500
    return jsonify(result)


def build_iron_patient_sticker_text(patient, is_small=False):
    patient_name = build_patient_name(patient)
    dob = format_display_date(patient.get("dob"))
    urn = clean_text(patient.get("urn"))
    medicare = clean_text(patient.get("medicare_number")) or clean_text(patient.get("medicare"))
    address = clean_text(patient.get("full_address")) or clean_text(patient.get("address"))
    phone = clean_text(patient.get("phone")) or clean_text(patient.get("home_phone")) or clean_text(patient.get("mobile_phone"))
    email = clean_text(patient.get("email"))
    private_health = clean_text(patient.get("private_health"))

    lines = []
    if is_small:
        if patient_name:
            lines.append(f"Name: {patient_name}")
        if dob or urn:
            lines.append(f"DOB: {dob or 'N/A'}  URN: {urn or 'N/A'}")
        if medicare:
            lines.append(f"Medicare: {medicare}")
        if private_health:
            lines.append(f"Private Health: {private_health}")
        if phone:
            lines.append(f"Phone: {phone}")
        if email:
            lines.append(f"Email: {email}")
        if address:
            lines.append(f"Address: {address}")
    else:
        if patient_name:
            lines.append(f"Patient Name: {patient_name}")
        
        second_line = []
        if dob:
            second_line.append(f"DOB: {dob}")
        if urn:
            second_line.append(f"URN: {urn}")
        if second_line:
            lines.append("   |   ".join(second_line))
            
        third_line = []
        if medicare:
            third_line.append(f"Medicare No: {medicare}")
        if private_health:
            third_line.append(f"Private Health: {private_health}")
        if third_line:
            lines.append("   |   ".join(third_line))
            
        if phone:
            lines.append(f"Phone: {phone}")
        if email:
            lines.append(f"Email: {email}")
        if address:
            lines.append(f"Address: {address}")
            
    return "\n".join(lines)


def open_iron_outlook_draft(patient, output_path, infusion_type):
    patient_name = build_patient_name(patient)
    subject = f"{patient_name} iron infusion"
    if infusion_type == "public":
        to_recipients = [PUBLIC_IRON_RECIPIENT]
    else:
        to_recipients = [PRIVATE_IRON_RECIPIENT]

    write_outlook_draft(
        subject,
        output_path,
        to_recipients=to_recipients,
        cc_recipients=[PRACTICE_CC],
        sender_email=IRON_SENDER_EMAIL,
    )


def generate_public_iron_pdf(patient, request_date=None):
    if not os.path.exists(PUBLIC_IRON_TEMPLATE):
        raise FileNotFoundError("Could not find public iron infusion.pdf")
        
    doc = fitz.open(PUBLIC_IRON_TEMPLATE)
    page = doc[0]
    
    sticker_rect = fitz.Rect(page.rect.x1 - 235, page.rect.y0 + 18, page.rect.x1 - 23, page.rect.y0 + 128)
    draw_filled_rect(page, sticker_rect)
    
    draw_rect_textbox(
        page,
        fitz.Rect(sticker_rect.x0 + 8, sticker_rect.y0 + 7, sticker_rect.x1 - 8, sticker_rect.y1 - 7),
        build_iron_patient_sticker_text(patient, is_small=True),
        fontsize=8.0,
        min_fontsize=5.5,
    )
    
    os.makedirs(GENERATED_IRON_DIR, exist_ok=True)
    patient_name_safe = re.sub(r"[^\w\s-]", "", build_patient_name(patient) or "Unknown").strip().replace(" ", "_")
    output_name = f"{patient_name_safe}_public_iron_{datetime.now().strftime('%Y-%m-%d')}.pdf"
    output_path = os.path.join(GENERATED_IRON_DIR, output_name)
    doc.save(output_path)
    doc.close()
    return output_name, output_path


def generate_private_iron_pdf(patient, request_date=None):
    if not os.path.exists(PRIVATE_IRON_TEMPLATE):
        raise FileNotFoundError("Could not find private Iron infusion.pdf")
        
    doc = fitz.open(PRIVATE_IRON_TEMPLATE)
    
    # Page 1
    page1 = doc[0]
    sticker_rect1 = fitz.Rect(230, 470, 1600, 715)
    draw_filled_rect(page1, sticker_rect1)
    draw_rect_textbox(
        page1,
        fitz.Rect(sticker_rect1.x0 + 15, sticker_rect1.y0 + 10, sticker_rect1.x1 - 15, sticker_rect1.y1 - 10),
        build_iron_patient_sticker_text(patient, is_small=False),
        fontsize=20.0,
        min_fontsize=12.0,
    )
    
    # Page 2
    page2 = doc[1]
    sticker_rect2 = fitz.Rect(1000, 1850, 1420, 2250)
    draw_filled_rect(page2, sticker_rect2)
    draw_rect_textbox(
        page2,
        fitz.Rect(sticker_rect2.x0 + 10, sticker_rect2.y0 + 10, sticker_rect2.x1 - 10, sticker_rect2.y1 - 10),
        build_iron_patient_sticker_text(patient, is_small=True),
        fontsize=18.0,
        min_fontsize=12.0,
        rotate=270,
    )
    
    os.makedirs(GENERATED_IRON_DIR, exist_ok=True)
    patient_name_safe = re.sub(r"[^\w\s-]", "", build_patient_name(patient) or "Unknown").strip().replace(" ", "_")
    output_name = f"{patient_name_safe}_private_iron_{datetime.now().strftime('%Y-%m-%d')}.pdf"
    output_path = os.path.join(GENERATED_IRON_DIR, output_name)
    doc.save(output_path)
    doc.close()
    return output_name, output_path


@app.route("/public-iron")
def public_iron():
    return render_template("public_iron.html")


@app.route("/private-iron")
def private_iron():
    return render_template("private_iron.html")


@app.route("/generated-iron/<path:filename>")
def generated_iron_file(filename):
    safe_name = os.path.basename(filename)
    path = os.path.join(GENERATED_IRON_DIR, safe_name)
    if not os.path.exists(path):
        return jsonify({"error": "File not found"}), 404
    return send_file(path, mimetype="application/pdf", as_attachment=False, download_name=safe_name)


@app.route("/api/generate-public-iron", methods=["POST"])
def api_generate_public_iron():
    load_data()
    data = request.get_json(silent=True) or {}
    patient = data.get("patient") or {}
    request_date = data.get("request_date")

    if not build_patient_name(patient):
        return jsonify({"error": "Please enter the patient name."}), 400

    try:
        output_name, output_path = generate_public_iron_pdf(patient, request_date)
        open_iron_outlook_draft(patient, output_path, "public")
    except Exception as exc:
        print("Error generating Public Iron request:", exc)
        return jsonify({"error": str(exc)}), 500

    return jsonify({
        "success": True,
        "message": "Public Iron request generated and opened in Outlook.",
        "url": url_for("generated_iron_file", filename=output_name),
        "output_path": output_path
    })


@app.route("/api/generate-private-iron", methods=["POST"])
def api_generate_private_iron():
    load_data()
    data = request.get_json(silent=True) or {}
    patient = data.get("patient") or {}
    request_date = data.get("request_date")

    if not build_patient_name(patient):
        return jsonify({"error": "Please enter the patient name."}), 400

    try:
        output_name, output_path = generate_private_iron_pdf(patient, request_date)
        open_iron_outlook_draft(patient, output_path, "private")
    except Exception as exc:
        print("Error generating Private Iron request:", exc)
        return jsonify({"error": str(exc)}), 500

    return jsonify({
        "success": True,
        "message": "Private Iron request generated and opened in Outlook.",
        "url": url_for("generated_iron_file", filename=output_name),
        "output_path": output_path
    })


def build_endoscopy_patient_sticker_text(patient):
    patient_name = build_patient_name(patient)
    dob = format_display_date(patient.get("dob"))
    phone = clean_text(patient.get("mobile_phone")) or clean_text(patient.get("phone")) or clean_text(patient.get("home_phone"))
    urn = clean_text(patient.get("urn"))
    medicare = clean_text(patient.get("medicare_number")) or clean_text(patient.get("medicare"))

    lines = []
    if patient_name:
        lines.append(f"Name: {patient_name}")
    dob_line_parts = []
    if dob:
        dob_line_parts.append(f"DOB: {dob}")
    if urn:
        dob_line_parts.append(f"URN: {urn}")
    if dob_line_parts:
        lines.append("  ".join(dob_line_parts))
    if phone:
        lines.append(f"Mob: {phone}")
    if medicare:
        lines.append(f"Medicare: {medicare}")
    return "\n".join(lines)


def open_endoscopy_outlook_draft(patient, output_path):
    patient_name = build_patient_name(patient)
    subject = f"{patient_name} - SVHM Endoscopy Referral"
    write_outlook_draft(
        subject,
        output_path,
        to_recipients=[PUBLIC_ENDOSCOPY_RECIPIENT],
        cc_recipients=[PRACTICE_CC],
        sender_email=IRON_SENDER_EMAIL,
    )


def generate_endoscopy_pdf(patient, location="mains", procedure="colonoscopy", diabetes_plan="", blood_thinners_plan="", request_date=None):
    loc_key = (location or "mains").strip().lower()
    proc_key = (procedure or "colonoscopy").strip().lower()

    if proc_key not in ENDOSCOPY_PAGE_1_TEMPLATES:
        proc_key = "colonoscopy"
    if loc_key not in ENDOSCOPY_PAGE_2_TEMPLATES:
        loc_key = "mains"

    p1_template = ENDOSCOPY_PAGE_1_TEMPLATES[proc_key]
    p2_template = ENDOSCOPY_PAGE_2_TEMPLATES[loc_key][proc_key]

    if not os.path.exists(p1_template):
        raise FileNotFoundError(f"Could not find template {p1_template}")
    if not os.path.exists(p2_template):
        raise FileNotFoundError(f"Could not find template {p2_template}")

    doc1 = fitz.open(p1_template)
    doc2 = fitz.open(p2_template)

    page1 = doc1[0]
    page2 = doc2[0]

    sticker_text = build_endoscopy_patient_sticker_text(patient)
    sticker_rect = fitz.Rect(440, 36, 575, 140)
    sticker_text_rect = fitz.Rect(445, 41, 570, 135)

    # 1. Page 1 Top-right sticker
    draw_filled_rect(page1, sticker_rect)
    draw_rect_textbox(page1, sticker_text_rect, sticker_text, fontsize=8.5, min_fontsize=6.0)

    # 2. Page 1 Referring doctor date
    formatted_date = format_display_date(request_date) or datetime.now().strftime("%d/%m/%Y")
    ref_date_rect = fitz.Rect(355, 175, 450, 192)
    draw_filled_rect(page1, ref_date_rect, fill=(1, 1, 1), border=(1, 1, 1))
    draw_rect_textbox(page1, ref_date_rect, formatted_date, fontsize=9.0)

    # 3. Page 1 Diabetes meds (bottom left)
    if diabetes_plan and diabetes_plan.strip():
        diab_rect = fitz.Rect(80, 715, 348, 746)
        draw_filled_rect(page1, diab_rect)
        draw_rect_textbox(
            page1,
            fitz.Rect(diab_rect.x0 + 4, diab_rect.y0 + 3, diab_rect.x1 - 4, diab_rect.y1 - 3),
            diabetes_plan.strip(),
            fontsize=8.0,
            min_fontsize=5.5,
        )

    # 4. Page 1 Blood thinners (bottom right)
    if blood_thinners_plan and blood_thinners_plan.strip():
        blood_rect = fitz.Rect(362, 693, 578, 796)
        draw_filled_rect(page1, blood_rect)
        draw_rect_textbox(
            page1,
            fitz.Rect(blood_rect.x0 + 5, blood_rect.y0 + 5, blood_rect.x1 - 5, blood_rect.y1 - 5),
            blood_thinners_plan.strip(),
            fontsize=8.0,
            min_fontsize=5.5,
        )

    # 5. Page 2 Top-right sticker
    draw_filled_rect(page2, sticker_rect)
    draw_rect_textbox(page2, sticker_text_rect, sticker_text, fontsize=8.5, min_fontsize=6.0)

    # 6. Page 2 Consent name and date
    patient_name = build_patient_name(patient)
    name_rect = fitz.Rect(110, 680, 280, 698)
    draw_rect_textbox(page2, name_rect, patient_name, fontsize=9.5)

    date_rect = fitz.Rect(455, 740, 550, 756)
    draw_rect_textbox(page2, date_rect, formatted_date, fontsize=9.5)

    # Merge page 2 into doc1
    doc1.insert_pdf(doc2)

    os.makedirs(GENERATED_ENDOSCOPY_DIR, exist_ok=True)
    patient_name_safe = re.sub(r"[^\w\s-]", "", patient_name or "Unknown").strip().replace(" ", "_")
    output_name = f"{patient_name_safe}_SVHM_endoscopy_{datetime.now().strftime('%Y-%m-%d')}.pdf"
    output_path = os.path.join(GENERATED_ENDOSCOPY_DIR, output_name)

    doc1.save(output_path)
    doc1.close()
    doc2.close()
    return output_name, output_path


@app.route("/endoscopy")
def endoscopy():
    return render_template("endoscopy.html")


@app.route("/generated-endoscopy/<path:filename>")
def generated_endoscopy_file(filename):
    safe_name = os.path.basename(filename)
    path = os.path.join(GENERATED_ENDOSCOPY_DIR, safe_name)
    if not os.path.exists(path):
        return jsonify({"error": "File not found"}), 404
    return send_file(path, mimetype="application/pdf", as_attachment=False, download_name=safe_name)


@app.route("/api/generate-endoscopy", methods=["POST"])
def api_generate_endoscopy():
    load_data()
    data = request.get_json(silent=True) or {}
    patient = data.get("patient") or {}
    location = data.get("location") or "mains"
    procedure = data.get("procedure") or "colonoscopy"
    diabetes_plan = data.get("diabetes_plan") or ""
    blood_thinners_plan = data.get("blood_thinners_plan") or ""
    request_date = data.get("request_date")

    if not build_patient_name(patient):
        return jsonify({"error": "Please enter the patient name."}), 400

    try:
        output_name, output_path = generate_endoscopy_pdf(
            patient,
            location=location,
            procedure=procedure,
            diabetes_plan=diabetes_plan,
            blood_thinners_plan=blood_thinners_plan,
            request_date=request_date,
        )
        open_endoscopy_outlook_draft(patient, output_path)
    except Exception as exc:
        print("Error generating Endoscopy referral:", exc)
        return jsonify({"error": str(exc)}), 500

    return jsonify({
        "success": True,
        "message": "Endoscopy referral generated and opened in Outlook.",
        "url": url_for("generated_endoscopy_file", filename=output_name),
        "output_path": output_path,
    })


if __name__ == "__main__":
    app.run(debug=True, port=5050)
