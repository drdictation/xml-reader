import csv
import html
import io
import json
import os
import re
import subprocess
from datetime import datetime

import fitz  # PyMuPDF
from bs4 import BeautifulSoup
from flask import Flask, jsonify, render_template, request, send_file, url_for
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

app = Flask(__name__)

BASE_DIR = os.path.dirname(__file__)
ROOT_DIR = os.path.abspath(os.path.join(BASE_DIR, ".."))
XML_DIR = "/Users/cbasnayake/Documents/BACKUP CMG XML/2017 onwards"
TEMPLATE_DIR = os.path.join(ROOT_DIR, "Letter templates")
PATHOLOGY_TEMPLATE = os.path.join(ROOT_DIR, "pathology.pdf")
RADIOLOGY_TEMPLATE_PDF = os.path.join(ROOT_DIR, "radioltemplate.pdf")
RADIOLOGY_TEMPLATE_JPEG = os.path.join(ROOT_DIR, "radioltemplate.jpeg")
IUS_TEMPLATE = os.path.join(ROOT_DIR, "IUS_template.pdf")
PATIENT_DATABASE_FILE = os.path.join(ROOT_DIR, "output", "patient-database.csv")
PATIENT_INDEX_FILE = os.path.join(ROOT_DIR, "output", "patient-index.json")
CACHE_FILE = os.path.join(BASE_DIR, "data_cache.json")
GENERATED_LETTERS_DIR = os.path.join(BASE_DIR, "generated_letters")
GENERATED_PATHOLOGY_DIR = os.path.join(BASE_DIR, "generated_pathology")
GENERATED_RADIOLOGY_DIR = os.path.join(BASE_DIR, "generated_radiology")
GENERATED_IUS_DIR = os.path.join(BASE_DIR, "generated_ius")
DATA_CACHE_VERSION = 2
SENDER_EMAIL = "drchamarabasnayake@gmail.com"
PRACTICE_CC = "office@focusgastro.com.au"

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


def compact_lines(value):
    lines = [line.strip() for line in clean_text(value).splitlines()]
    return "\n".join(line for line in lines if line)


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

                inferred_first, inferred_surname = infer_name_parts(name)
                full_address = clean_text(row.get("full_address"))
                mobile_phone = clean_text(row.get("mobile_phone"))
                home_phone = clean_text(row.get("home_phone"))

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
                        "phone": mobile_phone or home_phone or clean_text(index_entry.get("phone")),
                        "email": clean_text(row.get("email")),
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
        flowables.append(Spacer(1, 6))
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
                flowables.append(Spacer(1, 6))

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


def draw_rect_textbox(page, rect, text, *, fontsize=10.0, min_fontsize=7.0, align=0):
    value = compact_lines(text)
    if not value:
        return

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
    style_normal.fontSize = 11
    style_normal.leading = 16
    style_normal.fontName = "Helvetica"
    styles.add(ParagraphStyle(name="PatientBlock", parent=style_normal, fontName="Helvetica-Bold"))
    styles.add(ParagraphStyle(name="BulletList", parent=style_normal, leftIndent=30, bulletIndent=15))

    story = [
        Paragraph(datetime.now().strftime("%d %B %Y"), style_normal),
        Spacer(1, 16),
        Paragraph(referrer.get("name", "To whom it may concern"), style_normal),
    ]

    if referrer.get("clinic"):
        story.append(Paragraph(referrer.get("clinic", ""), style_normal))
    story.append(Spacer(1, 16))

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
            Spacer(1, 16),
            Paragraph(f"<b>Re: {patient.get('name', '')}</b>", styles["PatientBlock"]),
            Paragraph(f"<b>DOB: {patient.get('dob', '')}</b>", styles["PatientBlock"]),
            Paragraph(f"<b>Address: {patient.get('address', '')}</b>", styles["PatientBlock"]),
            Spacer(1, 16),
        ]
    )

    story.extend(html_to_story(body_text, style_normal, styles["BulletList"]))
    story.extend(
        [
            Spacer(1, 30),
            Paragraph("Yours sincerely,", style_normal),
            Spacer(1, 30),
            Paragraph("<b>A/Prof Chamara Basnayake</b>", style_normal),
            Paragraph("Gastroenterologist", style_normal),
        ]
    )

    if cc_doctors:
        names = ", ".join(doc.get("name", "") for doc in cc_doctors if doc.get("name"))
        if names:
            story.extend([Spacer(1, 20), Paragraph(f"CC: {names}", style_normal)])

    pathology_plain = BeautifulSoup(pathology_text, "html.parser").get_text().strip() if pathology_text else ""
    if pathology_plain:
        if "PathologyHeading" not in styles:
            styles.add(
                ParagraphStyle(
                    name="PathologyHeading",
                    parent=style_normal,
                    fontName="Helvetica-Bold",
                    fontSize=11,
                    spaceBefore=0,
                    spaceAfter=4,
                )
            )
        story.extend(
            [
                Spacer(1, 24),
                Paragraph("Investigation Results:", styles["PathologyHeading"]),
                Spacer(1, 6),
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

    applescript_path = os.path.join(BASE_DIR, "send_email.applescript")
    with open(applescript_path, "w", encoding="utf-8") as fh:
        fh.write(
            f'''
tell application "Microsoft Outlook"
    set newMessage to make new outgoing message with properties {{subject:"Patient Letter: {patient.get('name', '')}"}}
    try
        set theAccount to (first account whose email address is "{SENDER_EMAIL}")
        set account of newMessage to theAccount
    end try
'''
        )
        if to_email:
            fh.write(
                f'    make new recipient at newMessage with properties {{email address:{{address:"{to_email}"}}}}\n'
            )
        for cc_email in cc_emails:
            fh.write(
                f'    make new cc recipient at newMessage with properties {{email address:{{address:"{cc_email}"}}}}\n'
            )
        fh.write(
            f'''
    make new attachment at newMessage with properties {{file:POSIX file "{output_path}"}}
    open newMessage
end tell
'''
        )

    try:
        subprocess.run(["osascript", applescript_path], check=True)
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


if __name__ == "__main__":
    app.run(debug=True, port=5050)
