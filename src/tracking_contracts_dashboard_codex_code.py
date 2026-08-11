import csv
import json
import math
import re
from datetime import date, timedelta
from pathlib import Path

import pandas as pd


ROOT = Path("/Users/tspphupha/Documents/Codex/2026-07-11/files-mentioned-by-the-user-tracking")
SOURCE_HTML = Path("/Users/tspphupha/Downloads/tracking_contracts_dashboard_person_mixed_y_r_mockup_code.txt")
SOURCE_XLSX = Path("/Users/tspphupha/Downloads/Contract_SLA_Input_With_Contract_ID_Design.xlsx")
SOURCE_CONTRACT_MASTER_CSV = ROOT / "outputs" / "contract_master_sla_separated_2_export.csv"
SOURCE_CONTRACT_TYPE_MASTER_V2_CSV = Path("/Users/tspphupha/Downloads/contract_type_master_v2.csv")
OUTPUT_HTML = ROOT / "outputs" / "tracking_contracts_dashboard_real_excel_dropdowns.html"
OUTPUT_CONTRACTS_CSV = ROOT / "outputs" / "tracking_contracts_contracts_db.csv"
OUTPUT_LOGS_CSV = ROOT / "outputs" / "tracking_contracts_log_db.csv"
OUTPUT_TYPE_MASTER_CSV = ROOT / "outputs" / "tracking_contracts_type_master_db.csv"
OUTPUT_DEPARTMENT_MASTER_CSV = ROOT / "outputs" / "tracking_contracts_department_master_db.csv"
OUTPUT_PEOPLE_MASTER_CSV = ROOT / "outputs" / "tracking_contracts_people_master_db.csv"
OUTPUT_CONTRACT_TEMPLATE_CSV = ROOT / "outputs" / "tracking_contracts_contract_template_master_db.csv"
OUTPUT_ACTION_SLA_MASTER_CSV = ROOT / "outputs" / "tracking_contracts_action_sla_master_db.csv"
OUTPUT_CONTRACT_TYPE_MASTER_V2_CSV = ROOT / "outputs" / "contract_type_master_v2.csv"
OUTPUT_CODE = ROOT / "outputs" / "tracking_contracts_dashboard_codex_code.py"
OUTPUT_README = ROOT / "outputs" / "tracking_contracts_database_readme.txt"
OUTPUT_ATTACHMENT_APPS_SCRIPT = ROOT / "outputs" / "tracking_contracts_attachment_upload_apps_script.js"
DRIVE_FOLDER_ID = "1JH3z-QrsjhiHxc2h8IUGKTf-jxML1igj"
DRIVE_FOLDER_URL = f"https://drive.google.com/drive/folders/{DRIVE_FOLDER_ID}"
ATTACHMENT_CLOUD_FOLDER_ID = "1sU2_6KlvRSWZ3Rv-9bF9AEU7PvYBF4pJ"
ATTACHMENT_CLOUD_FOLDER_URL = f"https://drive.google.com/drive/folders/{ATTACHMENT_CLOUD_FOLDER_ID}"
ATTACHMENT_CLOUD_FOLDER_NAME = "Attachments Files"
ATTACHMENT_UPLOAD_ENDPOINT = "https://script.google.com/macros/s/AKfycbzhIbrLVvD-Cwxh3wqEWqjSaIESGgXfhdJ2cWUhepiSIsAyG8yQafG392kkjnSvjT_N/exec"
RESET_CONTRACT_AND_LOG_DATA = True
ACTIVE_UPDATE_ACTIONS = ["Submit to Review", "Return", "Resubmit", "Forward"]
STANDARD_SLA_DATA_VERSION = "2026-08-05-sla-config-updated1-v1"
DEPARTMENT_DATA_VERSION = "2026-07-23-nonzero-departments-v1"
ACTION_DATA_VERSION = "2026-08-05-sla-config-updated1-v1"

STANDARD_SLA_ADJUSTED_DAYS = {
    ("Day-to-day Work", "Lease & Rental Agreement", "Lease Agreement"): "20",
    ("Day-to-day Work", "Lease & Rental Agreement", "Sub Lease Agreement"): "20",
    ("Day-to-day Work", "Lease & Rental Agreement", "Lease Asset Agreement"): "20",
    ("Day-to-day Work", "Lease & Rental Agreement", "Rental Agreement"): "22",
    ("Day-to-day Work", "Service Agreement", ""): "22",
    ("Day-to-day Work", "Amendment Agreement", ""): "20",
    ("Day-to-day Work", "Sale and Purchase Agreement", ""): "24",
    ("Day-to-day Work", "Outsourcing Agreement", ""): "24",
    ("Day-to-day Work", "Commercial Agreement", "Consultancy Agreement"): "24",
    ("Day-to-day Work", "Commercial Agreement", "Confidentiality Agreement"): "20",
    ("Day-to-day Work", "Others", ""): "20",
    ("Confidential", "Preliminary Agreement", "Memorandum of Understanding"): "24",
    ("Confidential", "Preliminary Agreement", "Term Sheet"): "24",
    ("Confidential", "Commercial Agreement", "Consultancy Agreement"): "24",
    ("Confidential", "Commercial Agreement", "Confidentiality Agreement"): "20",
    ("Confidential", "Commercial Agreement", "Management Agreement"): "24",
    ("Confidential", "Commercial Agreement", "Loan Agreement"): "28",
    ("Confidential", "Commercial Agreement", "Mergers and Acquisitions Agreement"): "28",
    ("Confidential", "Commercial Agreement", "Shareholders’ Agreement"): "28",
    ("Confidential", "", "Others"): "20",
}

ACTION_SLA_ADJUSTED_DAYS = {
    "Submit to Review": 7,
    "Return": 3,
    "Resubmit": 5,
    "Forward": 5,
}

ACTION_DESCRIPTION_CONFIG = {
    "Submit to Review": {
        "descriptionEn": "Submit the contract to the review process.",
        "descriptionTh": "ส่งสัญญาเข้าสู่กระบวนการตรวจสอบ",
        "slaRuleEn": "7 working days (starts at Add Case)",
        "slaRuleTh": "7 วันทำการ (เริ่มเมื่อ Add Case)",
    },
    "Return": {
        "descriptionEn": "Return the contract for correction or additional information.",
        "descriptionTh": "ส่งสัญญากลับเพื่อแก้ไขหรือเพิ่มเติมข้อมูล",
        "slaRuleEn": "3 working days (starts when Return is received)",
        "slaRuleTh": "3 วันทำการ (เมื่อได้รับ Return)",
    },
    "Resubmit": {
        "descriptionEn": "Resubmit the corrected contract to the review process.",
        "descriptionTh": "ส่งสัญญากลับเข้าสู่กระบวนการตรวจสอบอีกครั้งหลังแก้ไข",
        "slaRuleEn": "5 working days (starts when Resubmit is received)",
        "slaRuleTh": "5 วันทำการ (เมื่อได้รับ Resubmit)",
    },
    "Forward": {
        "descriptionEn": "Forward the contract to the next department or responsible person.",
        "descriptionTh": "ส่งต่อสัญญาไปยังหน่วยงานหรือผู้รับผิดชอบลำดับถัดไป",
        "slaRuleEn": "5 working days (starts when Forward is received)",
        "slaRuleTh": "5 วันทำการ (เมื่อได้รับ Forward)",
    },
}

PRACTICAL_EXAMPLE_CONFIG = {
    "Day-to-day Work|Lease Agreement": {
        "en": "Lease office, restaurant or commercial space from a landlord.",
        "th": "เช่าสำนักงาน ร้านอาหาร หรือพื้นที่เชิงพาณิชย์จากผู้ให้เช่า",
    },
    "Day-to-day Work|Sub Lease Agreement": {
        "en": "Sublease part of leased premises to an affiliate, operator or tenant.",
        "th": "ให้บริษัทในเครือ ผู้ดำเนินการ หรือผู้เช่ารายอื่นเช่าช่วงพื้นที่",
    },
    "Day-to-day Work|Lease Asset Agreement": {
        "en": "Lease equipment, vehicles, POS, CCTV or other business assets.",
        "th": "เช่าอุปกรณ์ รถยนต์ POS CCTV หรือทรัพย์สินสำหรับธุรกิจ",
    },
    "Day-to-day Work|Rental Agreement": {
        "en": "Short-term rental of a venue, vehicle, equipment or temporary space.",
        "th": "เช่าสถานที่ รถ อุปกรณ์ หรือพื้นที่ชั่วคราวระยะสั้น",
    },
    "Day-to-day Work|Service Agreement": {
        "en": "IT maintenance, cleaning, marketing, logistics or recurring business services.",
        "th": "บริการดูแล IT ทำความสะอาด การตลาด ขนส่ง หรือบริการประจำ",
    },
    "Day-to-day Work|Amendment Agreement": {
        "en": "Extend a term or change price, scope, party details or selected clauses.",
        "th": "ขยายระยะเวลา หรือแก้ราคา ขอบเขต ข้อมูลคู่สัญญา หรือข้อกำหนดบางข้อ",
    },
    "Day-to-day Work|Sale and Purchase Agreement": {
        "en": "Purchase IT hardware, kitchen equipment, furniture, materials or other goods.",
        "th": "ซื้ออุปกรณ์ IT เครื่องครัว เฟอร์นิเจอร์ วัสดุ หรือสินค้า",
    },
    "Day-to-day Work|Outsourcing Agreement": {
        "en": "Engage a contractor, outsourced staff or project-based service provider.",
        "th": "จ้างผู้รับเหมา บุคลากร Outsource หรือผู้ให้บริการเป็นโครงการ",
    },
    "Day-to-day Work|Consultancy Agreement": {
        "en": "Business, IT, design, marketing or operational advisory for routine projects.",
        "th": "ที่ปรึกษาธุรกิจ IT ออกแบบ การตลาด หรือปฏิบัติการสำหรับโครงการทั่วไป",
    },
    "Day-to-day Work|Confidentiality Agreement": {
        "en": "Share information with a vendor or partner before a proposal, pilot or discussion.",
        "th": "เปิดเผยข้อมูลให้ผู้ขายหรือพันธมิตรก่อนเสนอราคา ทดลองระบบ หรือหารือ",
    },
    "Day-to-day Work|Others": {
        "en": "A routine agreement not matching an existing category, such as a trial or event.",
        "th": "ข้อตกลงทั่วไปที่ยังไม่เข้าหมวด เช่น ทดลองใช้หรืองานกิจกรรม",
    },
    "Confidential|Memorandum of Understanding": {
        "en": "A framework for cooperation, joint study, pilot or future business relationship.",
        "th": "กรอบความร่วมมือ ศึกษาร่วม ทดลองโครงการ หรือความสัมพันธ์ทางธุรกิจในอนาคต",
    },
    "Confidential|Term Sheet": {
        "en": "Record preliminary commercial terms for an investment, lease, financing or strategic deal.",
        "th": "สรุปเงื่อนไขเบื้องต้นของการลงทุน เช่า เงินทุน หรือธุรกรรมเชิงกลยุทธ์",
    },
    "Confidential|Consultancy Agreement": {
        "en": "Strategic, financial, transaction or sensitive-project advisory.",
        "th": "ที่ปรึกษากลยุทธ์ การเงิน ธุรกรรม หรือโครงการที่มีข้อมูลอ่อนไหว",
    },
    "Confidential|Confidentiality Agreement": {
        "en": "A deal-level NDA for an investment, acquisition, financing or restricted project.",
        "th": "NDA ระดับธุรกรรมสำหรับลงทุน ซื้อกิจการ เงินทุน หรือโครงการจำกัดสิทธิ",
    },
    "Confidential|Management Agreement": {
        "en": "Appoint a party to manage a restaurant, property, operation or business function.",
        "th": "แต่งตั้งผู้บริหารร้านอาหาร ทรัพย์สิน งานปฏิบัติการ หรือหน่วยธุรกิจ",
    },
    "Confidential|Loan Agreement": {
        "en": "Intercompany or external borrowing, shareholder loan or revolving facility.",
        "th": "เงินกู้ระหว่างบริษัท เงินกู้ผู้ถือหุ้น หรือวงเงินหมุนเวียน",
    },
    "Confidential|Mergers and Acquisitions Agreement": {
        "en": "Acquire or dispose of shares, a business or assets, including merger transactions.",
        "th": "ซื้อหรือขายหุ้น ธุรกิจ หรือทรัพย์สิน รวมถึงการควบรวม",
    },
    "Confidential|Shareholders’ Agreement": {
        "en": "Set governance and rights among founders, investors or joint-venture shareholders.",
        "th": "กำหนดธรรมาภิบาลและสิทธิระหว่างผู้ก่อตั้ง นักลงทุน หรือผู้ถือหุ้นร่วมทุน",
    },
    "Confidential|Others": {
        "en": "A sensitive investigation, restructuring, strategic project or restricted transaction.",
        "th": "งานสอบสวน ปรับโครงสร้าง โครงการกลยุทธ์ หรือธุรกรรมจำกัดสิทธิ",
    },
}

TODAY = date(2026, 7, 11)

CONTRACT_NAME_OVERRIDES = {
    "Project Beacon – Strategic Consultancy Agreement": "Code Name run number – Strategic Consultancy Agreement",
}

CUSTOM_CONTRACT_INPUT_ROWS = [
    ("Head Office Lease Agreement", "Lease Agreement", "Normal"),
    ("General Legal Service Agreement", "Service Agreement", "Normal"),
    ("Retail Space Sub Lease Agreement", "Sub Lease Agreement", "Normal"),
    ("Office Equipment Lease Asset Agreement", "Lease Asset Agreement", "Normal"),
    ("Amendment to Existing Agreement", "Amendment", "Normal"),
    ("Laptop Rental Agreement", "Rental Agreement", "Normal"),
    ("Equipment Sale and Purchase Agreement", "Sale and Purchase Agreement", "Normal"),
    ("Outsourced Service Provider Agreement", "Outsourcing Agreement", "Normal"),
    ("Confident-001 – Memorandum of Understanding", "Memorandum of Understanding", "Confidential"),
    ("Confident-002 – Consultancy Agreement", "Consultancy Agreement", "Confidential"),
    ("Confident-003 – Confidentiality Agreement", "Confidentiality Agreement", "Confidential"),
    ("Confident-004  – Management Agreement", "Management Agreement", "Confidential"),
    ("Confident-005  – Loan Agreement", "Loan Agreement", "Confidential"),
    ("Confident-006 – M&A Agreement", "Mergers and Acquisitions Agreement", "Confidential"),
    ("Confident-007 – Shareholder Agreement", "Sharer Holder Agreement", "Confidential"),
    ("Confident-008  – Term Sheet", "Term Sheet", "Confidential"),
]

DEPARTMENT_CODE_CONFIG = {
    "ADMIN": "ADMIN",
    "Information Technology": "IT",
    "IT": "IT",
    "Legal": "LEGAL",
    "Finance": "FIN",
    "Accounting": "ACC",
    "Human Resources": "HR",
    "Administration": "ADMIN",
    "Business Development": "BD",
    "Marketing": "MKT",
    "MKT": "MKT",
    "Operation": "OP",
    "Procurement": "PC",
    "Project Management": "PM",
    "Project Manager": "PM",
    "PM": "PM",
    "Baan Turtle": "BT",
    "Table 1749": "T1749",
    "FSQ": "FSQ",
}

DEPARTMENT_NAME_ALIASES = {
    "ADMIN": "Administration",
    "Information Technology": "IT",
    "MKT": "Marketing",
    "Project Management": "Project Manager",
    "PM": "Project Manager",
}

ACTIVE_DEPARTMENT_ORDER = [
    "Administration",
    "Business Development",
    "Finance",
    "IT",
    "Legal",
    "Marketing",
    "Operation",
    "Procurement",
    "Project Manager",
]


def clean(value):
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    if pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


def number(value, default=0):
    text = clean(value)
    if not text:
        return default
    try:
        value = float(text)
    except ValueError:
        return default
    return int(value) if value.is_integer() else value


def business_days_after(start, days):
    cursor = start
    remaining = int(days or 0)
    while remaining > 0:
        cursor += timedelta(days=1)
        if cursor.weekday() < 5:
            remaining -= 1
    return cursor.isoformat()


def read_table(sheet_name, header):
    df = pd.read_excel(SOURCE_XLSX, sheet_name=sheet_name, header=header, dtype=object)
    df = df.dropna(how="all").dropna(axis=1, how="all")
    return df


def records(df):
    normalized_rows = []
    for row in df.to_dict("records"):
        normalized = {clean(k): clean(v) for k, v in row.items()}
        contract_name = normalized.get("Contract Name / ชื่อสัญญา")
        if contract_name in CONTRACT_NAME_OVERRIDES:
            normalized["Contract Name / ชื่อสัญญา"] = CONTRACT_NAME_OVERRIDES[contract_name]
        normalized_rows.append(normalized)
    return normalized_rows


def first_sla_number(value, default=""):
    match = re.search(r"\d+(?:\.\d+)?", clean(value))
    if not match:
        return default
    number_value = float(match.group(0))
    return int(number_value) if number_value.is_integer() else number_value


def read_contract_master_rows():
    if not SOURCE_CONTRACT_MASTER_CSV.exists():
        return []
    with SOURCE_CONTRACT_MASTER_CSV.open("r", encoding="utf-8-sig", newline="") as handle:
        return [
            {clean(key): clean(value) for key, value in row.items()}
            for row in csv.DictReader(handle)
            if clean(row.get("Contract Name / ชื่อสัญญามาตรฐาน")) and clean(row.get("Type of Contract / ประเภทสัญญา"))
        ]


def read_contract_type_master_v2_rows():
    if not SOURCE_CONTRACT_TYPE_MASTER_V2_CSV.exists():
        return []
    with SOURCE_CONTRACT_TYPE_MASTER_V2_CSV.open("r", encoding="utf-8-sig", newline="") as handle:
        rows = []
        for row in csv.DictReader(handle):
            normalized = {clean(key): clean(value) for key, value in row.items()}
            if normalized.get("Type of Contract EN") or normalized.get("Sub Type of Contract EN"):
                rows.append(normalized)
        rows = apply_standard_sla_adjustments(rows)
        classification_order = {
            "day-to-day work": 0,
            "confidential": 1,
        }
        return sorted(rows, key=lambda item: (
            classification_order.get(item.get("Contract Classification EN", "").lower(), 9),
            number(item.get("Type Sort Order"), 999),
            number(item.get("Sub Type Sort Order"), 999),
            item.get("Type of Contract EN", ""),
            item.get("Sub Type of Contract EN", ""),
        ))


def apply_standard_sla_adjustments(rows):
    for row in rows:
        if clean(row.get("Type of Contract EN")) == "Service Provider Agreement":
            row["Type of Contract EN"] = "Outsourcing Agreement"
            row["Type of Contract TH"] = "สัญญาจ้าง"
        key = (
            clean(row.get("Contract Classification EN")),
            clean(row.get("Type of Contract EN")),
            clean(row.get("Sub Type of Contract EN")),
        )
        adjusted_days = STANDARD_SLA_ADJUSTED_DAYS.get(key)
        if adjusted_days is not None:
            row["Standard SLA"] = adjusted_days
    return rows


def bilingual_value(en, th):
    values = [clean(en), clean(th)]
    return " / ".join(value for value in values if value)


def type_rows_from_contract_type_master_v2(rows):
    type_rows = []
    for row in rows:
        classification = bilingual_value(row.get("Contract Classification EN"), row.get("Contract Classification TH"))
        type_value = bilingual_value(row.get("Type of Contract EN"), row.get("Type of Contract TH"))
        sub_type_value = bilingual_value(row.get("Sub Type of Contract EN"), row.get("Sub Type of Contract TH"))
        standard_sla = first_sla_number(
            row.get("Standard SLA")
            or row.get("Standard SLA / SLA รวม")
            or row.get("Standard SLA (Working Days)")
            or row.get("Total SLA / SLA รวม"),
            "",
        )
        type_rows.append({
            "Active": "Yes",
            "Contract Classification": classification,
            "Category": classification,
            "Type of Contract": type_value,
            "Sub Type of Contract": sub_type_value,
            "Fixed SLA (Working Days)": standard_sla,
            "Standard SLA Version": STANDARD_SLA_DATA_VERSION,
            "Description / คำอธิบาย": type_value,
        })
    return type_rows


def norm_contract_key(value):
    return re.sub(r"[^a-z0-9]+", "", clean(value).split("|")[0].split("/")[0].casefold())


def contract_classification_from_template(row):
    text = " ".join([
        row.get("Access Level / ระดับการเข้าถึง", ""),
        row.get("Category", ""),
        contract_group(row),
        row.get("หมายเหตุ", ""),
        row.get("Contract Name / ชื่อสัญญามาตรฐาน", ""),
    ]).casefold()
    return "Confidential" if "confident" in text or "confidential" in text or "สัญญาลับ" in text else "Day-to-day Work"


def contract_type_alias_value(type_value, classification="", context=""):
    text = " ".join([type_value or "", classification or "", context or ""]).casefold()
    normalized = norm_contract_key(text)
    if not normalized:
        return ""
    is_confidential = "confidential" in text or "confident" in text or "สัญญาลับ" in text
    if "non-disclosure" in text or "nondisclosure" in text or "ไม่เปิดเผย" in text:
        return "Confidentiality Agreement" if not is_confidential else "Confidentiality Agreement"
    if "mou" in text or "memorandum" in text:
        return "Memorandum of Understanding"
    if "term sheet" in text:
        return "Term Sheet"
    if "loan" in text or "เงินกู้" in text:
        return "Loan Agreement"
    if "shareholder" in text:
        return "Shareholders’ Agreement"
    if "merger" in text or "acquisition" in text or "m&a" in text:
        return "Mergers and Acquisitions Agreement"
    if "management" in text and is_confidential:
        return "Management Agreement"
    if "outsourc" in text or "contractor" in text or "construction contract" in text or "ผู้รับเหมา" in text:
        return "Outsourcing Agreement"
    if "consult" in text or "design" in text or "ที่ปรึกษา" in text or "ออกแบบ" in text:
        return "Consultancy Agreement"
    if "lease asset" in text:
        return "Lease Asset Agreement"
    if "sublease" in text or "sub lease" in text or "เช่าช่วง" in text:
        return "Sub Lease Agreement"
    if "rental" in text or "เช่าพื้นที่" in text or "เช่าทั่วไป" in text or "lease and service" in text or "booth rental" in text:
        return "Rental Agreement"
    if "lease" in text or "สัญญาเช่า" in text:
        return "Lease Agreement"
    if "sale and purchase" in text or "ซื้อขาย" in text:
        return "Sale and Purchase Agreement"
    if "service provider" in text or "provider" in text or "partner service" in text or "delivery platform" in text:
        return "Outsourcing Agreement"
    if "service" in text or "บริการ" in text or "management fee" in text or "management services" in text:
        return "Service Agreement"
    if "amendment" in text or "แก้ไขเพิ่มเติม" in text:
        return "Amendment Agreement"
    return "Others"


def contract_type_flow_from_v2(rows, type_value, classification="", context=""):
    normalized = norm_contract_key(type_value)
    if not normalized:
        return {}
    classification = clean(classification)
    alias = contract_type_alias_value(type_value, classification, context)
    aliases = [alias] if alias else []
    candidates = []
    for row in rows:
        if classification and row.get("Contract Classification EN") != classification:
            continue
        type_display = bilingual_value(row.get("Type of Contract EN"), row.get("Type of Contract TH"))
        sub_type_display = bilingual_value(row.get("Sub Type of Contract EN"), row.get("Sub Type of Contract TH"))
        match_values = [sub_type_display, row.get("Sub Type of Contract EN"), row.get("Sub Type of Contract TH")]
        if not sub_type_display:
            match_values.extend([type_display, row.get("Type of Contract EN"), row.get("Type of Contract TH")])
        score = 0
        for value in match_values:
            if norm_contract_key(value) == normalized:
                score = max(score, 100)
            if any(norm_contract_key(value) == norm_contract_key(alias_value) for alias_value in aliases):
                score = max(score, 90)
        if score:
            candidates.append((score, row, type_display, sub_type_display))
    if not candidates:
        return {}
    _, row, type_display, sub_type_display = sorted(candidates, key=lambda item: item[0], reverse=True)[0]
    classification_display = bilingual_value(row.get("Contract Classification EN"), row.get("Contract Classification TH"))
    standard_sla = first_sla_number(
        row.get("Standard SLA")
        or row.get("Standard SLA / SLA รวม")
        or row.get("Standard SLA (Working Days)")
        or row.get("Total SLA / SLA รวม"),
        "",
    )
    return {
        "classification": classification_display,
        "typeGroup": type_display,
        "subType": sub_type_display,
        "type": sub_type_display or type_display,
        "fixedSla": standard_sla,
        "accessLevel": "Confidential" if row.get("Contract Classification EN") == "Confidential" else "Normal",
    }


def contract_group(row):
    return (
        row.get("Group Contract / กลุ่มสัญญา")
        or row.get("Group Contract / กลุ่มสัญญา ")
        or row.get("กลุ่มสัญญา")
        or row.get("Category")
        or ""
    )


def js(value):
    return json.dumps(value, ensure_ascii=False, indent=6)


def replace_between(text, start_marker, end_marker, replacement):
    start = text.index(start_marker)
    end = text.index(end_marker, start)
    return text[:start] + replacement + text[end:]


def normalized_lookup(value):
    return clean(value).casefold()


def canonical_department_name(value):
    name = clean(value)
    if not name or name == "0":
        return ""
    direct = DEPARTMENT_NAME_ALIASES.get(name)
    if direct:
        return direct
    normalized = normalized_lookup(name)
    for alias, canonical in DEPARTMENT_NAME_ALIASES.items():
        if normalized_lookup(alias) == normalized:
            return canonical
    return name


def department_code_for(department):
    department = canonical_department_name(department)
    if not department:
        return ""
    direct = DEPARTMENT_CODE_CONFIG.get(department)
    if direct:
        return direct
    normalized = normalized_lookup(department)
    for name, code in DEPARTMENT_CODE_CONFIG.items():
        if normalized_lookup(name) == normalized:
            return code
    return ""


def access_level_code(access_level):
    return "C" if clean(access_level) == "Confidential" else "N"


def generate_contract_ids(rows):
    counters = {}
    id_map = {}
    for row in rows:
        old_id = row.get("Contract ID", "")
        access_level = row.get("Access Level / ระดับการเข้าถึง", "") or "Normal"
        department = row.get("Department / Restaurant", "")
        department_code = department_code_for(department)
        if not department_code:
            continue
        group_key = (access_level_code(access_level), department_code)
        counters[group_key] = counters.get(group_key, 0) + 1
        new_id = f"CT-{group_key[0]}-{group_key[1]}-{str(counters[group_key]).zfill(3)}"
        if old_id:
            id_map[old_id] = new_id
        row["Contract ID"] = new_id
    return id_map


def remap_contract_id_cells(rows, id_map):
    updated_rows = []
    for row in rows:
        updated_rows.append([id_map.get(cell, cell) for cell in row])
    return updated_rows


def write_csv(path, rows, headers):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def main():
    register_df = read_table("Contract Register", 3)
    sla_input_df = read_table("Contract SLA Input", 3)
    type_master_df = read_table("Contract Type Master", 0)
    action_sla_df = read_table("Fix Action SLA", 2)
    rules_raw_df = pd.read_excel(SOURCE_XLSX, sheet_name="Contract ID Rules", header=None, dtype=object)

    register_rows = records(register_df)
    sla_input_rows = [
        row for row in records(sla_input_df)
        if row.get("Contract Name / ชื่อสัญญา") and row.get("Type of Contract / ประเภทสัญญา")
    ]
    type_rows = records(type_master_df)
    for row in type_rows:
        row.setdefault("Active", "Yes")
    action_rows = [
        row for row in records(action_sla_df)
        if row.get("Action", "") in ACTIVE_UPDATE_ACTIONS
    ]
    contract_master_rows = read_contract_master_rows()
    contract_type_master_v2_rows = read_contract_type_master_v2_rows()
    rules_rows = [
        [clean(cell) for cell in row]
        for row in rules_raw_df.fillna("").values.tolist()
        if any(clean(cell) for cell in row)
    ]
    contract_id_map = generate_contract_ids(register_rows)
    rules_rows = remap_contract_id_cells(rules_rows, contract_id_map)
    rules_rows = [
        [
            cell
                .replace("CT-N-YY-###", "CT-N-[Department Code]-###")
                .replace("CT-C-YY-###", "CT-C-[Department Code]-###")
            if isinstance(cell, str) else cell
            for cell in row
        ]
        for row in rules_rows
    ]

    if contract_type_master_v2_rows:
        type_rows = type_rows_from_contract_type_master_v2(contract_type_master_v2_rows)
    elif contract_master_rows:
        type_by_name = {}
        for row in contract_master_rows:
            contract_type = row.get("Type of Contract / ประเภทสัญญา", "")
            if not contract_type:
                continue
            group = contract_group(row)
            sla_text = row.get("Total SLA / SLA รวม", "")
            current = type_by_name.setdefault(contract_type, {
                "Active": "Yes",
                "Category": group or "Day-to-day",
                "Description / คำอธิบาย": group,
                "Fixed SLA (Working Days)": "",
                "Type of Contract": contract_type,
            })
            if group and not current.get("Description / คำอธิบาย"):
                current["Description / คำอธิบาย"] = group
            if sla_text:
                values = [item.strip() for item in str(current.get("Fixed SLA (Working Days)", "")).split(" | ") if item.strip()]
                if sla_text not in values:
                    values.append(sla_text)
                current["Fixed SLA (Working Days)"] = " | ".join(values)
        type_rows = list(type_by_name.values())

    type_descriptions = {
        row["Type of Contract"]: row.get("Description / คำอธิบาย", "")
        for row in type_rows
        if row.get("Type of Contract")
    }
    type_categories = {
        row["Type of Contract"]: row.get("Category", "")
        for row in type_rows
        if row.get("Type of Contract")
    }
    if contract_master_rows:
        sla_input_rows = [
            {
                "Contract Name / ชื่อสัญญา": row.get("Contract Name / ชื่อสัญญามาตรฐาน", ""),
                "Type of Contract / ประเภทสัญญา": row.get("Type of Contract / ประเภทสัญญา", ""),
                "Access Level / ระดับการเข้าถึง": "Normal",
                "Category": contract_group(row),
                "Department / Restaurant": row.get("Department / Restaurant", ""),
                "Vendor / Counter Party / ผู้ขาย–คู่สัญญา": row.get("Vendor / Counter Party / ผู้ขาย–คู่สัญญา", ""),
                "Total SLA / SLA รวม": row.get("Total SLA / SLA รวม", ""),
            }
            for row in contract_master_rows
        ]
    else:
        sla_input_rows = [
            {
                "Contract Name / ชื่อสัญญา": name,
                "Type of Contract / ประเภทสัญญา": contract_type,
                "Access Level / ระดับการเข้าถึง": access_level,
                "Category": access_level,
            }
            for name, contract_type, access_level in CUSTOM_CONTRACT_INPUT_ROWS
        ]

    contracts = []
    log_records = []
    fixed_sla = {"Other": 3}
    contract_catalog = []
    department_master_rows = [
        {
            "Department / Restaurant": name,
            "Department Code": code,
            "Active": "Yes",
        }
        for name, code in DEPARTMENT_CODE_CONFIG.items()
    ]
    known_departments = {row["Department / Restaurant"] for row in department_master_rows}
    for row in contract_master_rows:
        department = row.get("Department / Restaurant", "")
        if department and department not in known_departments:
            department_master_rows.append({
                "Department / Restaurant": department,
                "Department Code": department_code_for(department) or department.upper().replace(" ", "_"),
                "Active": "Yes",
            })
            known_departments.add(department)
    source_html_for_people = SOURCE_HTML.read_text(encoding="utf-8")
    people_match = re.search(r"const employeeDirectory = (\[[\s\S]*?\]);", source_html_for_people)
    people_master_rows = []
    if people_match:
        try:
            people_master_rows = json.loads(people_match.group(1))
        except json.JSONDecodeError:
            people_master_rows = []
    for row in people_master_rows:
        row["department"] = canonical_department_name(row.get("department"))
        row.setdefault("active", "Yes")

    departments_with_people = {
        row.get("department", "")
        for row in people_master_rows
        if row.get("department") and clean(row.get("department")) != "0"
    }
    ordered_active_departments = [
        name for name in ACTIVE_DEPARTMENT_ORDER if name in departments_with_people
    ]
    ordered_active_departments.extend(sorted(
        name for name in departments_with_people if name not in ordered_active_departments
    ))
    department_master_rows = [
        {
            "Department / Restaurant": name,
            "Department Code": department_code_for(name),
            "Department Data Version": DEPARTMENT_DATA_VERSION,
            "Active": "Yes",
        }
        for name in ordered_active_departments
        if name and department_code_for(name)
    ]

    for index, row in enumerate(register_rows, start=1):
        contract_id = row.get("Contract ID", "")
        name = row.get("Contract Name / ชื่อสัญญา", "")
        contract_type = row.get("Type of Contract / ประเภทสัญญา", "")
        department = row.get("Department / Restaurant", "")
        total_sla = number(row.get("Total SLA / SLA รวม"), 0)
        access_level = row.get("Access Level / ระดับการเข้าถึง", "")
        visibility = row.get("Visibility / สิทธิ์การมองเห็น", "")
        vendor = "Confidential Party" if access_level == "Confidential" else ""
        due = business_days_after(TODAY, total_sla)
        owner = department or "Contract Owner"
        station = f"From {owner} >> To Legal"

        fixed_sla[contract_type] = total_sla
        contract_catalog.append({
            "name": name,
            "type": contract_type,
            "workType": contract_type,
            "contractId": contract_id,
            "accessLevel": access_level,
            "category": type_categories.get(contract_type, ""),
        })
        contracts.append({
            "id": contract_id,
            "name": name,
            "department": department,
            "owner": owner,
            "type": contract_type,
            "vendor": vendor,
            "stage": "Draft Created",
            "cycle": 1,
            "returns": 0,
            "status": "Green >>G=On Track",
            "station": station,
            "due": due,
            "workType": contract_type,
            "totalSla": total_sla,
            "used": 0,
            "balance": total_sla,
            "alert": "Green >>G=On Track",
            "remark": "Imported from Contract Register",
            "accessLevel": access_level,
            "visibility": visibility,
            "category": type_categories.get(contract_type, ""),
            "typeDescription": type_descriptions.get(contract_type, ""),
        })
        log_records.append([
            contract_id,
            1,
            1,
            station,
            owner,
            "Legal",
            TODAY.isoformat(),
            "",
            total_sla,
            0,
            "Green >>G=On Track",
            "",
            "Draft Created",
            "Imported from Contract Register",
            "OK",
            "",
            "Import Note",
            "Initial record from Excel workbook",
            "",
            "",
            "",
            "Excel Import",
            f"{TODAY.isoformat()}T09:00:00+07:00",
            "DRAFT_CREATED",
            "สร้างรายการสัญญา",
            "Draft Created",
            "สร้างรายการสัญญาจากทะเบียนสัญญาใน Excel",
            "Create an initial contract tracking record from the Excel register.",
            total_sla,
            "หมายเหตุ",
            "Import Note",
            [],
            [],
        ])

    if contract_master_rows:
        contract_catalog = []
        name_counts = {}
        for row in contract_master_rows:
            name = row.get("Contract Name / ชื่อสัญญามาตรฐาน", "")
            name_counts[name] = name_counts.get(name, 0) + 1
        for row in contract_master_rows:
            name = row.get("Contract Name / ชื่อสัญญามาตรฐาน", "")
            contract_type = row.get("Type of Contract / ประเภทสัญญา", "")
            sla_text = row.get("Total SLA / SLA รวม", "")
            vendor = row.get("Vendor / Counter Party / ผู้ขาย–คู่สัญญา", "")
            template_classification = contract_classification_from_template(row)
            classification_fallback = bilingual_value(
                template_classification,
                "สัญญาลับ" if template_classification == "Confidential" else "งานดำเนินงานทั่วไป",
            )
            type_flow = contract_type_flow_from_v2(
                contract_type_master_v2_rows,
                contract_type,
                template_classification,
                " ".join([name, contract_group(row), row.get("หมายเหตุ", "")]),
            )
            selection_label = name
            if name_counts.get(name, 0) > 1:
                selection_label = " — ".join([value for value in [name, vendor, f"SLA {type_flow.get('fixedSla', '')}" if type_flow.get("fixedSla") else ""] if value])
            if contract_type:
                fixed_sla[contract_type] = first_sla_number(sla_text, fixed_sla.get(contract_type, ""))
            contract_catalog.append({
                "classification": type_flow.get("classification") or classification_fallback,
                "typeGroup": type_flow.get("typeGroup", ""),
                "subType": type_flow.get("subType", ""),
                "name": name,
                "selectionLabel": selection_label,
                "sourceRow": row.get("ลำดับ", ""),
                "type": type_flow.get("type", contract_type),
                "workType": type_flow.get("type", contract_type),
                "contractId": "",
                "accessLevel": type_flow.get("accessLevel", template_classification),
                "category": type_flow.get("classification") or contract_group(row) or template_classification,
                "department": canonical_department_name(row.get("Department / Restaurant", "")),
                "vendor": vendor,
                "group": type_flow.get("typeGroup") or contract_group(row),
                "fixedSla": type_flow.get("fixedSla", ""),
                "remark": row.get("หมายเหตุ", ""),
                "active": "Yes",
            })
    else:
        contract_catalog = [
            {
                "name": name,
                "type": contract_type,
                "workType": contract_type,
                "contractId": "",
                "accessLevel": access_level,
                "category": access_level,
                "fixedSla": fixed_sla.get(contract_type, ""),
                "active": "Yes",
            }
            for name, contract_type, access_level in CUSTOM_CONTRACT_INPUT_ROWS
        ]
    for row in contract_catalog:
        row["slaVersion"] = STANDARD_SLA_DATA_VERSION
    if not contract_type_master_v2_rows:
        for row in type_rows:
            contract_type = row.get("Type of Contract", "")
            row["Fixed SLA (Working Days)"] = row.get("Fixed SLA (Working Days)") or fixed_sla.get(contract_type, "")
    if RESET_CONTRACT_AND_LOG_DATA:
        contracts = []
        log_records = []

    action_sla = {}
    sla_steps = []
    for row in action_rows:
        action = row.get("Action", "")
        days = ACTION_SLA_ADJUSTED_DAYS.get(action, number(row.get("Fix SLA (Working Days)"), 0))
        desc = row.get("Description", "")
        if not action:
            continue
        action_sla[action] = days
        sla_steps.append(["All Contract Types", action, "Current Owner", "Next Owner", f"{days} days", desc])
    action_sla_master_rows = [
        {
            "Action": action,
            "Description / รายละเอียด": ACTION_DESCRIPTION_CONFIG[action]["descriptionTh"],
            "Fixed SLA (Working Days)": action_sla.get(action, 0),
            "SLA Rule / วิธีนับ": ACTION_DESCRIPTION_CONFIG[action]["slaRuleTh"],
            "Action Data Version": ACTION_DATA_VERSION,
            "Active": "Yes",
        }
        for action in ACTIVE_UPDATE_ACTIONS
    ]

    sla_summary = []
    seen_types = set()
    for row in register_rows:
        contract_type = row.get("Type of Contract / ประเภทสัญญา", "")
        if not contract_type or contract_type in seen_types:
            continue
        seen_types.add(contract_type)
        sla_summary.append([
            contract_type,
            type_categories.get(contract_type, "Imported"),
            number(row.get("Total SLA / SLA รวม"), 0),
            type_descriptions.get(contract_type, ""),
            1,
            "Yes",
            TODAY.isoformat(),
            "Imported from Contract Register",
        ])

    # Preserve the deployed SLA-detail order so generator runs stay reproducible.
    deployed_type_order = [
        "Term Sheet",
        "Sale and Purchase Agreement",
        "Confidentiality Agreement",
        "Sub Lease Agreement",
        "Rental Agreement",
        "Service Agreement",
        "Service Provider Agreement",
        "Amendment",
        "Memorandum of Understanding",
        "Management Agreement",
        "Loan Agreement",
        "Lease Agreement",
        "Sharer Holder Agreement",
        "Mergers and Acquisitions Agreement",
        "Lease Asset Agreement",
        "Consultancy Agreement",
    ]
    ordered_sla_types = [item for item in deployed_type_order if item in seen_types]
    ordered_sla_types.extend(row[0] for row in sla_summary if row[0] not in ordered_sla_types)

    sla_detail = []
    for type_index, contract_type in enumerate(ordered_sla_types, start=1):
        for action_index, row in enumerate(action_rows, start=1):
            action = row.get("Action", "")
            days = ACTION_SLA_ADJUSTED_DAYS.get(action, number(row.get("Fix SLA (Working Days)"), 0))
            desc = row.get("Description", "")
            sla_detail.append([
                f"{contract_type}|{action}|Any",
                contract_type,
                action_index,
                action,
                "Current Owner",
                "Any",
                days,
                "Fixed Action SLA",
                "Yes",
                desc,
            ])

    list_rows = [
        [row.get("Action", ""), "Any", "", "", "Action", "Once", "Yes"]
        for row in action_rows
        if row.get("Action")
    ]

    workbook_data = {
        "sourceFile": SOURCE_XLSX.name,
        "importedAt": TODAY.isoformat(),
        "sheets": {
            "Contract Register": register_rows,
            "Contract SLA Input": sla_input_rows,
            "Contract Type Master": type_rows,
            "Contract Type Master V2": contract_type_master_v2_rows,
            "Fix Action SLA": action_rows,
            "Contract ID Rules": rules_rows,
        },
    }

    html = SOURCE_HTML.read_text(encoding="utf-8")

    html = replace_between(
        html,
        "    const contracts = [",
        "\n\n    const dashboardData = ",
        f"    const contracts = {js(contracts)};\n\n",
    )
    html = replace_between(
        html,
        "    const dashboardData = [",
        "\n\n    const logItems = ",
        "    const dashboardData = [];\n\n",
    )
    html = replace_between(
        html,
        "    const slaSteps = [",
        "\n\n    const notificationQueue = ",
        f"    const slaSteps = {js(sla_steps)};\n\n",
    )
    html = replace_between(
        html,
        "    const logRecords = [",
        "\n\n    const listRows = ",
        f"    const logRecords = {js(log_records)};\n\n",
    )
    html = replace_between(
        html,
        "    const listRows = [",
        "\n\n    const recipients = ",
        f"    const listRows = {js(list_rows)};\n\n",
    )
    html = replace_between(
        html,
        "    const slaSummaryConfig = [",
        "\n\n    const slaDetailConfig = ",
        f"    const slaSummaryConfig = {js(sla_summary)};\n\n",
    )
    html = replace_between(
        html,
        "    const slaDetailConfig = [",
        "\n\n    const logRecords = ",
        f"    const slaDetailConfig = {js(sla_detail)};\n\n",
    )
    html = replace_between(
        html,
        "    const actionSlaConfig = Object.freeze({",
        "\n    });\n\n    // Fixed SLA values imported",
        f"    const actionSlaConfig = Object.freeze({js(action_sla)});\n\n    // Fixed SLA values imported",
    )
    html = replace_between(
        html,
        "    const fixedSlaConfig = Object.freeze({",
        "\n    });\n\n    // Contract examples imported",
        f"    const fixedSlaConfig = Object.freeze({js(fixed_sla)});\n\n    // Contract examples imported",
    )
    html = replace_between(
        html,
        "    const contractInputCatalog = [",
        "\n    ];\n    const requestedRole",
        f"    const contractInputCatalog = {js(contract_catalog)};\n    const contractTypeMasterV2 = Object.freeze({js(contract_type_master_v2_rows)});\n    const practicalExampleConfig = Object.freeze({js(PRACTICAL_EXAMPLE_CONFIG)});\n    const standardSlaDataVersion = {js(STANDARD_SLA_DATA_VERSION)};\n    const departmentDataVersion = {js(DEPARTMENT_DATA_VERSION)};\n    const actionDataVersion = {js(ACTION_DATA_VERSION)};\n    const actionDescriptionConfig = Object.freeze({js(ACTION_DESCRIPTION_CONFIG)});\n    const departmentNameAliases = Object.freeze({js(DEPARTMENT_NAME_ALIASES)});\n    const realWorkbookData = Object.freeze({js(workbook_data)});\n    const driveDatabaseConfig = Object.freeze({js({ 'folderId': DRIVE_FOLDER_ID, 'folderUrl': DRIVE_FOLDER_URL, 'contractsCsv': OUTPUT_CONTRACTS_CSV.name, 'logsCsv': OUTPUT_LOGS_CSV.name, 'typeMasterCsv': OUTPUT_TYPE_MASTER_CSV.name, 'departmentMasterCsv': OUTPUT_DEPARTMENT_MASTER_CSV.name, 'peopleMasterCsv': OUTPUT_PEOPLE_MASTER_CSV.name, 'contractTemplateCsv': OUTPUT_CONTRACT_TEMPLATE_CSV.name, 'actionSlaCsv': OUTPUT_ACTION_SLA_MASTER_CSV.name })});\n    const attachmentCloudConfig = Object.freeze({js({ 'folderId': ATTACHMENT_CLOUD_FOLDER_ID, 'folderUrl': ATTACHMENT_CLOUD_FOLDER_URL, 'folderName': ATTACHMENT_CLOUD_FOLDER_NAME, 'uploadEndpoint': ATTACHMENT_UPLOAD_ENDPOINT })});\n    const requestedRole",
    )
    html = html.replace(
        "    const requestedRole",
        f"""    const masterData = {{
	      departments: {js(department_master_rows)},
	      people: employeeDirectory.map(item => ({{ ...item, active: item.active || "Yes" }})),
	      contractTypes: (realWorkbookData?.sheets?.["Contract Type Master"] || []).map(item => ({{ ...item, Active: item.Active || "Yes" }})),
	      contractTemplates: contractInputCatalog.map(item => ({{ ...item, active: item.active || "Yes" }})),
	      actionSla: {js(action_sla_master_rows)}
	    }};

    const requestedRole""",
        1,
    )
    html = html.replace(
        """    const currentUser = {
      name: requestedRole === "viewer" ? "Viewer" : "Kira",
      role: requestedRole === "viewer" ? "viewer" : "admin"
    };""",
        """    const accessRoleConfig = Object.freeze({
      viewer: { level: 1, code: "level1", nameEn: "Contract Viewer", nameTh: "ผู้ดูข้อมูลสัญญา", userName: "Viewer" },
      user: { level: 2, code: "level2", nameEn: "Contract User", nameTh: "ผู้ดำเนินการสัญญา", userName: "Contract User" },
      confidential: { level: 3, code: "level3", nameEn: "Confidential Contract User", nameTh: "ผู้ดำเนินการสัญญาลับ", userName: "Confidential User" },
      admin: { level: 4, code: "level4", nameEn: "System Administrator", nameTh: "ผู้ดูแลระบบ", userName: "Kira" }
    });
    const requestedRoleKey = ({
      level1: "viewer",
      "contract-viewer": "viewer",
      level2: "user",
      "contract-user": "user",
      level3: "confidential",
      "confidential-user": "confidential",
      level4: "admin",
      "system-administrator": "admin"
    })[String(requestedRole || "").trim().toLowerCase()] || String(requestedRole || "").trim().toLowerCase();
    const activeRoleKey = accessRoleConfig[requestedRoleKey] ? requestedRoleKey : "viewer";
    const activeRole = accessRoleConfig[activeRoleKey];
    const currentUser = {
      name: activeRole.userName,
      role: activeRoleKey,
      level: activeRole.level,
      accessLevel: activeRole.code,
      roleNameEn: activeRole.nameEn,
      roleNameTh: activeRole.nameTh
    };""",
    )
    html = html.replace(
        """    const updateActionList = updateActionDefinitions.map(item => item.nameEn);
    const updateActionByName = Object.freeze(Object.fromEntries(updateActionDefinitions.map(item => [item.nameEn, item])));""",
        f"""    const actionDefinitionOverrides = Object.freeze({js(ACTION_DESCRIPTION_CONFIG)});
    updateActionDefinitions.forEach(item => {{
      const configured = actionDefinitionOverrides[item.nameEn];
      if (!configured) return;
      item.shortEn = configured.descriptionEn;
      item.shortTh = configured.descriptionTh;
      item.descriptionEn = configured.descriptionEn;
      item.descriptionTh = configured.descriptionTh;
      item.sla = Number({js(ACTION_SLA_ADJUSTED_DAYS)}[item.nameEn]) || item.sla;
    }});
    const activeUpdateActionNames = new Set({js(ACTIVE_UPDATE_ACTIONS)});
    const activeUpdateActionDefinitions = updateActionDefinitions.filter(item => activeUpdateActionNames.has(item.nameEn));
    const updateActionList = activeUpdateActionDefinitions.map(item => item.nameEn);
    const updateActionByName = Object.freeze(Object.fromEntries(activeUpdateActionDefinitions.map(item => [item.nameEn, item])));""",
    )
    html = html.replace(
        """    const actionList = ["Draft Created", "Submit to Review", "Return", "Resubmit", "Forward", "Approved", "Signed / Completed", "Cancelled"];""",
        """    const actionList = ["Draft Created", "Submit to Review", "Return", "Resubmit", "Forward", "Signed / Completed", "Cancelled"];""",
    )
    html = html.replace('                  <option value="Approved">Approved</option>\n', "")
    html = html.replace('                      <option value="Approved">Approved</option>\n', "")
    html = html.replace(
        """    const updateActionReasonConfig = Object.freeze(Object.fromEntries(updateActionDefinitions.map(item => [item.nameEn, {""",
        """    const updateActionReasonConfig = Object.freeze(Object.fromEntries(activeUpdateActionDefinitions.map(item => [item.nameEn, {""",
    )

    old_summary = """    function contractSummaryMarkup(contract) {
      if (!contract) return "";
      return `
        <div><span>Contract Name</span><strong>${escapeHtml(contract.name || "-")}</strong></div>
        <div><span>Type of Contract</span><strong>${escapeHtml(contract.type || "-")}</strong></div>
        <div><span>Work Type</span><strong>${escapeHtml(contract.workType || "-")}</strong></div>
        <div><span>Total SLA</span><strong>${totalSlaFor(contract.workType)} Working Days</strong></div>`;
    }
"""
    new_summary = """    function contractSummaryMarkup(contract, options = {}) {
      if (!contract) return "";
      return `
        <div><span>Contract ID</span><strong>${escapeHtml(contract.id || "-")}</strong></div>
        <div><span>Contract Name</span><strong>${escapeHtml(contract.name || "-")}</strong></div>
        <div><span>Type of Contract</span><strong>${escapeHtml(contract.type || "-")}</strong></div>
        <div><span>Department / Restaurant</span><strong>${escapeHtml(contract.department || "-")}</strong></div>
        <div><span>Contract Owner</span><strong>${escapeHtml(contract.owner || "-")}</strong></div>
        ${options.showStatusUpdate ? `<div><span>Status Update</span><strong>${contractStatusBadge(contract.status || contract.alert || "")}</strong></div>` : ""}`;
    }
"""
    html = html.replace(old_summary, new_summary)
    html = html.replace(
        'document.querySelector("#updateContractSummary").innerHTML = contractSummaryMarkup(updateContract);',
        'document.querySelector("#updateContractSummary").innerHTML = contractSummaryMarkup(updateContract, { showStatusUpdate: true });',
    )
    html = html.replace(
        """      // Contract selectors must show only Contract ID and Contract Name.
      // Owner, Department, Stage and Status remain linked in the background only.
      const contractOptions = contracts.map(item => ({
        value: item.id,
        label: `${String(item.id || "").trim()} / ${String(item.name || "").trim()}`
      }));""",
        """      // Contract selectors show Contract ID, Type of Contract, Contract Name, and Department for easier search.
      // Owner, Department, Stage and Status remain linked in the background only.
      const contractOptions = contracts.map(item => ({
        value: item.id,
        label: [item.id, contractPrimaryTypeDisplay(item.type), item.name, item.department].map(value => String(value || "").trim()).filter(Boolean).join(" / ")
      }));""",
    )
    html = html.replace(
        """          if (contract) option.textContent = `${String(contract.id || "").trim()} / ${String(contract.name || "").trim()}`;""",
        """          if (contract) option.textContent = [contract.id, contractPrimaryTypeDisplay(contract.type), contract.name, contract.department].map(value => String(value || "").trim()).filter(Boolean).join(" / ");""",
    )
    html = html.replace(
        """<span class="bilingual-stack update-bilingual-label"><span class="en">Contract ID / Contract Name<span class="field-required-mark" aria-hidden="true">*</span></span><span class="th">รหัสสัญญา / ชื่อสัญญา<span class="field-required-mark" aria-hidden="true">*</span></span></span>
                    </span>
                    <select class="select" name="contractId" id="updateContract" required aria-label="Contract ID / Contract Name / รหัสสัญญา / ชื่อสัญญา"></select>""",
        """<span class="bilingual-stack update-bilingual-label"><span class="en">Contract ID / Type of Contract<span class="field-required-mark" aria-hidden="true">*</span> / Contract Name / Department</span><span class="th">รหัสสัญญา / ประเภทสัญญา<span class="field-required-mark" aria-hidden="true">*</span> / ชื่อสัญญา / แผนก</span></span>
                    </span>
                    <select class="select" name="contractId" id="updateContract" required aria-label="Contract ID / Type of Contract / Contract Name / Department / รหัสสัญญา / ประเภทสัญญา / ชื่อสัญญา / แผนก"></select>""",
    )
    html = html.replace(
        """<span class="bilingual-stack update-bilingual-label"><span class="en">Contract ID / Contract Name / Type of Contract<span class="field-required-mark" aria-hidden="true">*</span></span><span class="th">รหัสสัญญา / ชื่อสัญญา / ประเภทสัญญา<span class="field-required-mark" aria-hidden="true">*</span></span></span>
                    </span>
                    <select class="select" name="contractId" id="updateContract" required aria-label="Contract ID / Contract Name / Type of Contract / รหัสสัญญา / ชื่อสัญญา / ประเภทสัญญา"></select>""",
        """<span class="bilingual-stack update-bilingual-label"><span class="en">Contract ID / Type of Contract<span class="field-required-mark" aria-hidden="true">*</span> / Contract Name / Department</span><span class="th">รหัสสัญญา / ประเภทสัญญา<span class="field-required-mark" aria-hidden="true">*</span> / ชื่อสัญญา / แผนก</span></span>
                    </span>
                    <select class="select" name="contractId" id="updateContract" required aria-label="Contract ID / Type of Contract / Contract Name / Department / รหัสสัญญา / ประเภทสัญญา / ชื่อสัญญา / แผนก"></select>""",
    )
    html = html.replace(
        """            <div class="panel-header">
              <div>
                <h2>Contract Master</h2>
                <small>ติดตาม Contract Owner, cycle, return และสถานะล่าสุด</small>
              </div>
            </div>
            <div class="table-wrap">""",
        """            <div class="panel-header">
              <div>
                <h2>Contract Master</h2>
                <small>ติดตาม Contract Owner, cycle, return และสถานะล่าสุด</small>
              </div>
            </div>
            <div class="table-wrap">""",
    )
    html = html.replace(
        """<span class="bilingual-label"><span>Contract ID / Contract Name<span class="field-required-mark" aria-hidden="true">*</span></span><span>รหัสสัญญา / ชื่อสัญญา<span class="field-required-mark" aria-hidden="true">*</span></span></span>
                    <select class="select" name="contractId" id="closeContract" required></select>""",
        """<span class="bilingual-label"><span>Contract ID / Type of Contract<span class="field-required-mark" aria-hidden="true">*</span> / Contract Name / Department</span><span>รหัสสัญญา / ประเภทสัญญา<span class="field-required-mark" aria-hidden="true">*</span> / ชื่อสัญญา / แผนก</span></span>
                    <select class="select" name="contractId" id="closeContract" required aria-label="Contract ID / Type of Contract / Contract Name / Department / รหัสสัญญา / ประเภทสัญญา / ชื่อสัญญา / แผนก"></select>""",
    )
    html = html.replace(
        """<span class="bilingual-label"><span>Contract ID / Contract Name<span class="field-required-mark" aria-hidden="true">*</span></span><span>รหัสสัญญา / ชื่อสัญญา<span class="field-required-mark" aria-hidden="true">*</span></span></span>
                    <select class="select" name="contractId" id="adjustDueContract" required></select>""",
        """<span class="bilingual-label"><span>Contract ID / Type of Contract<span class="field-required-mark" aria-hidden="true">*</span> / Contract Name / Department</span><span>รหัสสัญญา / ประเภทสัญญา<span class="field-required-mark" aria-hidden="true">*</span> / ชื่อสัญญา / แผนก</span></span>
                    <select class="select" name="contractId" id="adjustDueContract" required aria-label="Contract ID / Type of Contract / Contract Name / Department / รหัสสัญญา / ประเภทสัญญา / ชื่อสัญญา / แผนก"></select>""",
    )
    html = html.replace(
        """      [...contractInputCatalog, ...contracts.map(item => ({ name: item.name, type: item.type, workType: item.workType }))]
        .forEach(item => {
          const name = String(item?.name || "").trim();
          if (!name) return;
          if (!merged.has(name)) merged.set(name, { name, type: item.type || "", workType: item.workType || "Other" });
        });""",
        """      contractInputCatalog
        .forEach(item => {
          const name = String(item?.name || "").trim();
          if (!name) return;
          if (!merged.has(name)) merged.set(name, {
            name,
            type: item.type || "",
            workType: item.workType || "Other",
            contractId: item.contractId || "",
            accessLevel: item.accessLevel || "",
            category: item.category || ""
          });
        });""",
    )
    html = html.replace(
        """                  <div class="case-form-section-title"><span class="section-index">1</span>Contract classification flow <small>Contract Name → Type of Contract → Vendor / Counter party</small></div>
                  <div class="linked-contract-flow full">
                    <label class="form-field full linked-flow-field" data-flow-order="1">
                      <span class="linked-flow-head">
                        <span class="linked-flow-number">1</span>
                        <span class="bilingual-label"><span>Contract Name<span class="field-required-mark" aria-hidden="true">*</span></span><span>ชื่อสัญญา<span class="field-required-mark" aria-hidden="true">*</span></span></span>
                      </span>
                      <input class="input" name="name" id="addContractName" list="contractNameOptions" required autocomplete="off" placeholder="Select an existing Contract Name or type a new name">
                      <small class="linked-flow-status" id="linkedNameStatus">Start here · เริ่มกรอกชื่อสัญญา</small>
                    </label>""",
        """                  <div class="case-form-section-title"><span class="section-index">1</span>Contract classification flow <small>Group Contract → Contract Name → Type of Contract → Vendor / Counter party</small></div>
                  <div class="linked-contract-flow full">
                    <label class="form-field full linked-flow-field" data-flow-order="0">
                      <span class="linked-flow-head">
                        <span class="bilingual-label"><span>Group Contract</span><span>กลุ่มสัญญา</span></span>
                      </span>
                      <input class="input" name="group" id="addContractGroup" list="contractGroupOptions" autocomplete="off" placeholder="Select Group Contract to filter / เลือกกลุ่มสัญญาเพื่อค้นหา">
                      <small class="linked-flow-status" id="linkedGroupStatus">All groups · แสดงทุกกลุ่ม</small>
                    </label>
                    <label class="form-field full linked-flow-field" data-flow-order="1">
                      <span class="linked-flow-head">
                        <span class="linked-flow-number">1</span>
                        <span class="bilingual-label"><span>Contract Name<span class="field-required-mark" aria-hidden="true">*</span></span><span>ชื่อสัญญา<span class="field-required-mark" aria-hidden="true">*</span></span></span>
                      </span>
                      <input class="input" name="name" id="addContractName" list="contractNameOptions" required autocomplete="off" placeholder="Select an existing Contract Name or type a new name">
                      <small class="linked-flow-status" id="linkedNameStatus">Select Contract Name · เลือกชื่อสัญญา</small>
                    </label>""",
    )
    html = html.replace(
        """                    <label class="form-field full linked-flow-field" data-flow-order="3">
                      <span class="linked-flow-head">
                        <span class="linked-flow-number">3</span>
                        <span class="bilingual-label"><span>Vendor / Counter party</span><span>ผู้ขาย / คู่สัญญา</span></span>
                      </span>
                      <input class="input" name="vendor" id="addVendor" autocomplete="off" placeholder="Enter Vendor or Counter party / ระบุผู้ขายหรือคู่สัญญา">
                      <small class="linked-flow-status">Enter the contract counter party · ระบุผู้ขายหรือคู่สัญญา</small>
                    </label>""",
        """                    <label class="form-field full linked-flow-field" data-flow-order="3">
                      <span class="linked-flow-head">
                        <span class="linked-flow-number">3</span>
                        <span class="bilingual-label"><span>Vendor / Counter party</span><span>ผู้ขาย / คู่สัญญา</span></span>
                      </span>
                      <input class="input" name="vendor" id="addVendor" autocomplete="off" placeholder="Enter Vendor or Counter party / ระบุผู้ขายหรือคู่สัญญา">
                      <small class="linked-flow-status">Enter the contract counter party · ระบุผู้ขายหรือคู่สัญญา</small>
                    </label>""",
    )
    html = html.replace(
        """                      <input class="input" name="type" id="addContractType" list="contractTypeOptions" required autocomplete="off" placeholder="Waiting for Contract Name / รอชื่อสัญญา" disabled>
                      <small class="linked-flow-status waiting" id="linkedTypeStatus">Waiting for Contract Name · รอข้อมูล Contract Name</small>""",
        """                      <input class="input" name="type" id="addContractType" list="contractTypeOptions" required autocomplete="off" placeholder="Waiting for Contract Name / รอชื่อสัญญา" disabled>
                      <input type="hidden" name="accessLevel" id="addAccessLevel" value="">
                      <div class="inline-access-status" aria-label="Access Level / ระดับการเข้าถึง">
                        <small class="access-level-chip waiting" id="addAccessLevelBadge">Waiting</small>
                        <small class="linked-flow-status waiting compact-access-status" id="linkedAccessStatus">Waiting for Type · รอประเภท</small>
                      </div>
                      <small class="linked-flow-status waiting" id="linkedTypeStatus">Waiting for Contract Name · รอชื่อสัญญา</small>""",
    )
    html = html.replace(
        """<small class="linked-flow-status" id="linkedNameStatus">Start here · เริ่มกรอกชื่อสัญญา</small>""",
        """<small class="linked-flow-status" id="linkedNameStatus">Select Contract Name · เลือกชื่อสัญญา</small>""",
    )
    html = html.replace(
        """                    <input type="hidden" name="workType" id="addWorkType" value="">
                  </div>

                  <div class="case-form-section-title"><span class="section-index">2</span>Ownership <small>ผู้รับผิดชอบสัญญา</small></div>""",
        """                    <input type="hidden" name="workType" id="addWorkType" value="">
                  </div>

                  <div class="add-case-smart-summary" id="addCaseSmartSummary" aria-live="polite">
                    <div class="summary-cell">
                      <span>Classification</span>
                      <strong id="addSummaryClassification">Day-to-day Work</strong>
                      <small id="addSummaryClassificationTh">งานดำเนินงานทั่วไป</small>
                    </div>
                    <div class="summary-cell">
                      <span>Type of Contract</span>
                      <strong id="addSummaryType">-</strong>
                      <small id="addSummaryTypeTh">เลือกประเภทสัญญา</small>
                    </div>
                    <div class="summary-cell">
                      <span>Sub Type</span>
                      <strong id="addSummarySubType">-</strong>
                      <small id="addSummarySubTypeTh">ประเภทย่อย</small>
                    </div>
                    <div class="summary-cell">
                      <span>Total SLA - Day / Due Date</span>
                      <strong id="addSummarySla">-</strong>
                      <small id="addSummaryDue">-</small>
                    </div>
                    <div class="summary-cell">
                      <span>Contract Name</span>
                      <strong id="addSummaryContractName">-</strong>
                      <small>ชื่อสัญญา</small>
                    </div>
                    <div class="summary-cell">
                      <span>Department</span>
                      <strong id="addSummaryDepartment">-</strong>
                      <small>แผนก / ร้านอาหาร</small>
                    </div>
                    <div class="summary-cell">
                      <span>Vendor / Counter party</span>
                      <strong id="addSummaryVendor">-</strong>
                      <small>ผู้ขาย / คู่สัญญา</small>
                    </div>
                    <div class="summary-cell">
                      <span>Contract Owner</span>
                      <strong id="addSummaryOwner">-</strong>
                      <small>เจ้าของสัญญา</small>
                    </div>
                    <div class="summary-cell">
                      <span>New Contract ID</span>
                      <strong id="addSummaryContractId">-</strong>
                      <small>รันตาม Department + Access</small>
                    </div>
                  </div>

                  <div class="case-form-section-title"><span class="section-index">2</span>Ownership <small>ผู้รับผิดชอบสัญญา</small></div>""",
    )
    html = html.replace(
        """                  <div class="case-form-section-title"><span class="section-index">1</span>Contract classification flow <small>Group Contract → Contract Name → Type of Contract → Vendor / Counter party</small></div>
                  <div class="linked-contract-flow full">
                    <label class="form-field full linked-flow-field" data-flow-order="0">
                      <span class="linked-flow-head">
                        <span class="bilingual-label"><span>Group Contract</span><span>กลุ่มสัญญา</span></span>
                      </span>
                      <input class="input" name="group" id="addContractGroup" list="contractGroupOptions" autocomplete="off" placeholder="Select Group Contract to filter / เลือกกลุ่มสัญญาเพื่อค้นหา">
                      <small class="linked-flow-status" id="linkedGroupStatus">All groups · แสดงทุกกลุ่ม</small>
                    </label>
                    <label class="form-field full linked-flow-field" data-flow-order="1">
                      <span class="linked-flow-head">
                        <span class="linked-flow-number">1</span>
                        <span class="bilingual-label"><span>Contract Name<span class="field-required-mark" aria-hidden="true">*</span></span><span>ชื่อสัญญา<span class="field-required-mark" aria-hidden="true">*</span></span></span>
                      </span>
                      <input class="input" name="name" id="addContractName" list="contractNameOptions" required autocomplete="off" placeholder="Select an existing Contract Name or type a new name">
                      <small class="linked-flow-status" id="linkedNameStatus">Select Contract Name · เลือกชื่อสัญญา</small>
                    </label>
                    <label class="form-field full linked-flow-field" data-flow-order="2">
                      <span class="linked-flow-head">
                        <span class="linked-flow-number">2</span>
                        <span class="bilingual-label"><span>Type of Contract<span class="field-required-mark" aria-hidden="true">*</span></span><span>ประเภทสัญญา<span class="field-required-mark" aria-hidden="true">*</span></span></span>
                      </span>
                      <input class="input" name="type" id="addContractType" list="contractTypeOptions" required autocomplete="off" placeholder="Waiting for Contract Name / รอชื่อสัญญา" disabled>
                      <input type="hidden" name="accessLevel" id="addAccessLevel" value="">
                      <div class="inline-access-status" aria-label="Access Level / ระดับการเข้าถึง">
                        <small class="access-level-chip waiting" id="addAccessLevelBadge">Waiting</small>
                        <small class="linked-flow-status waiting compact-access-status" id="linkedAccessStatus">Waiting for Type · รอประเภท</small>
                      </div>
                      <small class="linked-flow-status waiting" id="linkedTypeStatus">Waiting for Contract Name · รอชื่อสัญญา</small>
                    </label>
                    <label class="form-field full linked-flow-field" data-flow-order="3">
                      <span class="linked-flow-head">
                        <span class="linked-flow-number">3</span>
                        <span class="bilingual-label"><span>Vendor / Counter party</span><span>ผู้ขาย / คู่สัญญา</span></span>
                      </span>
                      <input class="input" name="vendor" id="addVendor" autocomplete="off" placeholder="Enter Vendor or Counter party / ระบุผู้ขายหรือคู่สัญญา">
                      <small class="linked-flow-status">Enter the contract counter party · ระบุผู้ขายหรือคู่สัญญา</small>
                    </label>
                    <input type="hidden" name="workType" id="addWorkType" value="">
                  </div>""",
        """                  <div class="case-form-section-title"><span class="section-index">1</span>Contract classification flow <small>Contract Classification → Type of Contract → Sub Type → Contract Name → Vendor / Counter party</small></div>
                  <div class="linked-contract-flow full">
                    <label class="form-field full linked-flow-field classification-flow-field" data-flow-order="1">
                      <span class="linked-flow-head">
                        <span class="linked-flow-number">1</span>
                        <span class="bilingual-label"><span>Contract Classification<span class="field-required-mark" aria-hidden="true">*</span></span><span>กลุ่มสัญญา<span class="field-required-mark" aria-hidden="true">*</span></span></span>
                      </span>
	                      <input type="hidden" name="classification" id="addContractClassification" value="Day-to-day Work / งานดำเนินงานทั่วไป">
	                      <div class="classification-choice-group" role="group" aria-label="Contract Classification / กลุ่มสัญญา">
	                        <button class="classification-choice active" type="button" data-classification-value="Day-to-day Work">
	                          <strong>Day-to-day Work</strong>
	                          <span>งานดำเนินงานทั่วไป</span>
	                        </button>
	                        <button class="classification-choice" type="button" data-classification-value="Confidential">
	                          <strong>Confidential</strong>
	                          <span>สัญญาลับ</span>
	                        </button>
	                      </div>
	                      <input type="hidden" name="accessLevel" id="addAccessLevel" value="Normal">
                      <div class="inline-access-status" aria-label="Access Level / ระดับการเข้าถึง">
                        <small class="access-level-chip normal" id="addAccessLevelBadge">Normal</small>
                        <small class="linked-flow-status linked compact-access-status" id="linkedAccessStatus">Day-to-day Work · งานดำเนินงานทั่วไป</small>
                      </div>
                    </label>
                    <label class="form-field full linked-flow-field" data-flow-order="2">
                      <span class="linked-flow-head">
                        <span class="linked-flow-number">2</span>
                        <span class="bilingual-label"><span>Type of Contract<span class="field-required-mark" aria-hidden="true">*</span></span><span>ประเภทสัญญา<span class="field-required-mark" aria-hidden="true">*</span></span></span>
                      </span>
                      <input class="input" name="type" id="addContractType" list="contractTypeOptions" required autocomplete="off" placeholder="Select Type of Contract">
                      <small class="bilingual-field-help"><span>Select the main contract category to determine the SLA.</span><span>เลือกประเภทสัญญาหลักเพื่อกำหนด SLA</span></small>
                      <small class="linked-flow-status waiting" id="linkedTypeStatus">Select Type to calculate SLA · เลือกประเภทเพื่อคำนวณ SLA</small>
                    </label>
                    <label class="form-field full linked-flow-field" data-flow-order="3">
                      <span class="linked-flow-head">
                        <span class="linked-flow-number">3</span>
                        <span class="bilingual-label"><span>Sub Type of Contract</span><span>ประเภทย่อยของสัญญา</span></span>
                      </span>
                      <input class="input" name="subType" id="addContractSubType" list="contractSubTypeOptions" autocomplete="off" placeholder="Select Sub Type">
                      <small class="bilingual-field-help"><span>Select the specific Sub Type when available.</span><span>เลือกประเภทย่อยที่ตรงกับสัญญาเมื่อมีตัวเลือก</span></small>
                      <small class="linked-flow-status waiting" id="linkedSubTypeStatus">Optional when Type has no Sub Type · ไม่บังคับเมื่อไม่มีประเภทย่อย</small>
                      <div class="practical-example" id="addPracticalExample" aria-live="polite">
                        <strong><span>Practical Example</span><span>ตัวอย่างการใช้งาน</span></strong>
                        <span class="practical-example-line" id="addPracticalExampleEn">Select Type and Sub Type to view an example.</span>
                        <span class="practical-example-line thai" id="addPracticalExampleTh">เลือกประเภทและประเภทย่อยเพื่อดูตัวอย่าง</span>
                      </div>
                    </label>
                    <label class="form-field full linked-flow-field" data-flow-order="4">
                      <span class="linked-flow-head">
                        <span class="linked-flow-number">4</span>
                        <span class="bilingual-label"><span>Contract Name<span class="field-required-mark" aria-hidden="true">*</span></span><span>ชื่อสัญญา<span class="field-required-mark" aria-hidden="true">*</span></span></span>
                      </span>
                      <input class="input" name="name" id="addContractName" required autocomplete="off" placeholder="Enter Contract Name / ระบุชื่อสัญญา">
                      <small class="linked-flow-status waiting" id="linkedNameStatus">Enter Contract Name · ระบุชื่อสัญญา</small>
                    </label>
                    <label class="form-field full linked-flow-field" data-flow-order="5">
                      <span class="linked-flow-head">
                        <span class="linked-flow-number">5</span>
                        <span class="bilingual-label"><span>Vendor / Counter party</span><span>ผู้ขาย / คู่สัญญา</span></span>
                      </span>
                      <input class="input" name="vendor" id="addVendor" autocomplete="off" placeholder="Enter Vendor or Counter party / ระบุผู้ขายหรือคู่สัญญา">
                      <small class="linked-flow-status">Enter the contract counter party · ระบุผู้ขายหรือคู่สัญญา</small>
                    </label>
                    <input type="hidden" name="workType" id="addWorkType" value="">
                  </div>""",
    )
    html = html.replace(
        """                    <span class="bilingual-label"><span>SLA (System / Fix)</span><span>SLA ที่ระบบกำหนด</span></span>""",
        """                    <span class="bilingual-label"><span>Total SLA - Day</span><span>SLA รวม - วันทำการ</span></span>""",
    )
    html = html.replace(
        """            <div class="panel-header">
              <div>
                <h2>Contract Master</h2>
                <small>ติดตาม Contract Owner, cycle, return และสถานะล่าสุด</small>
              </div>
            </div>""",
        """            <div class="panel-header contract-status-header">
              <div>
                <h2>Contract Master</h2>
                <small>ติดตาม Contract Owner, cycle, return และสถานะล่าสุด</small>
              </div>
              <div class="contract-access-tabs" role="tablist" aria-label="Contract access view / มุมมองสิทธิ์สัญญา">
                <button class="contract-access-tab active" type="button" role="tab" aria-selected="true" data-contract-access-view="normal">
                  <span>Normal Contracts</span><small>สัญญาทั่วไป</small><b id="normalContractCount">0</b>
                </button>
                <button class="contract-access-tab confidential" type="button" role="tab" aria-selected="false" data-contract-access-view="confidential" data-confidential-access-tab>
                  <span>Confidential</span><small>สัญญาลับ</small><b id="confidentialContractCount">0</b>
                </button>
              </div>
            </div>""",
        1,
    )
    html = html.replace(
        """                    <small class="form-hint">Linked from Type of Contract · เชื่อมจากประเภทสัญญา</small>""",
        """                    <small class="form-hint" id="addSlaHint">Fixed by selected Type / Sub Type · กำหนดตามประเภทและประเภทย่อยที่เลือก</small>""",
    )
    html = html.replace(
        """                <datalist id="contractNameOptions"></datalist>""",
        """                <datalist id="contractClassificationOptions"></datalist>
                <datalist id="contractSubTypeOptions"></datalist>
                <datalist id="contractNameOptions"></datalist>""",
    )
    html = html.replace(
        """    .linked-flow-field .input {
      width: 100%;
      max-width: none;
    }""",
        """    .linked-flow-field .input {
      width: 100%;
      max-width: none;
    }

    .linked-flow-field .combo-toggle {
      display: none;
    }

    .linked-flow-field .editable-combobox .input {
      padding-right: 12px;
    }

    .linked-flow-field > .linked-flow-status {
      display: block;
      min-height: 16px;
      line-height: 1.3;
    }

    .bilingual-field-help {
      display: block;
      margin-top: 7px;
      color: var(--muted);
      line-height: 1.35;
    }

    .bilingual-field-help span,
    .practical-example strong span,
    .practical-example-line {
      display: block;
    }

    .bilingual-field-help span + span,
    .practical-example .thai {
      margin-top: 2px;
      color: var(--muted);
      font-size: 0.92em;
    }

    .practical-example {
      margin-top: 10px;
      padding-left: 10px;
      border-left: 3px solid var(--green);
      line-height: 1.4;
    }

    .practical-example strong {
      display: block;
      margin-bottom: 5px;
      color: var(--ink);
      font-size: 12px;
    }

    .contract-status-header {
      align-items: flex-end;
      gap: 16px;
    }

    .contract-access-tabs {
      display: inline-grid;
      grid-template-columns: repeat(2, minmax(138px, 1fr));
      gap: 4px;
      padding: 4px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface-soft);
    }

    .contract-access-tab {
      min-height: 44px;
      padding: 6px 10px;
      border: 1px solid transparent;
      border-radius: 6px;
      background: transparent;
      color: var(--muted);
      cursor: pointer;
      text-align: left;
    }

    .contract-access-tab span,
    .contract-access-tab small {
      display: block;
      line-height: 1.25;
    }

    .contract-access-tab small {
      font-size: 10px;
    }

    .contract-access-tab b {
      float: right;
      margin-top: -25px;
      font-size: 11px;
    }

    .contract-access-tab.active {
      border-color: var(--green);
      background: var(--surface);
      color: var(--ink);
    }

    .contract-access-tab.confidential.active {
      border-color: var(--red);
    }

    @media (max-width: 760px) {
      .contract-status-header {
        align-items: stretch;
      }

      .contract-access-tabs {
        width: 100%;
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }
    }

    .access-level-chip {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: fit-content;
      min-height: 22px;
      margin-top: 4px;
      padding: 3px 9px;
      border: 1px solid var(--line);
      border-radius: 999px;
      background: #f8faf3;
      color: var(--muted);
      font-size: 11px;
      font-weight: 800;
      line-height: 1;
    }

    .access-level-chip.normal {
      border-color: #cfe2a0;
      background: #f1f7d8;
      color: #60790f;
    }

    .access-level-chip.confidential {
      border-color: #f0c2a7;
      background: #fff0e6;
      color: #b94700;
    }

    .access-level-chip.waiting {
      border-style: dashed;
    }

    .inline-access-status {
      display: flex;
      align-items: center;
      flex-wrap: wrap;
      gap: 5px;
      margin-top: 5px;
    }

    .inline-access-status .access-level-chip {
      margin-top: 0;
    }

    .compact-access-status {
      margin-top: 0;
      font-size: 10px;
      line-height: 1.25;
    }

    .classification-flow-field {
      display: flex;
      flex-direction: column;
      gap: 8px;
    }

    .classification-choice-group {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 7px;
    }

    .classification-choice {
      min-width: 0;
      min-height: 54px;
      padding: 8px 9px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #ffffff;
      color: var(--ink);
      text-align: left;
      cursor: pointer;
      transition: border-color .15s ease, background .15s ease, box-shadow .15s ease;
    }

    .classification-choice strong,
    .classification-choice span {
      display: block;
      overflow-wrap: anywhere;
    }

    .classification-choice strong {
      font-size: 11px;
      line-height: 1.2;
    }

    .classification-choice span {
      margin-top: 3px;
      color: var(--muted);
      font-size: 10px;
      line-height: 1.25;
    }

    .classification-choice.active {
      border-color: var(--primary);
      background: #f3f8de;
      box-shadow: 0 0 0 2px rgba(184, 214, 61, .22);
    }

    .add-case-smart-summary {
      grid-column: 1 / -1;
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 8px;
      padding: 10px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fbfcf8;
    }

    .add-case-smart-summary .summary-cell {
      min-width: 0;
      padding: 8px 9px;
      border: 1px solid #e4e8d8;
      border-radius: 7px;
      background: #fff;
    }

    .add-case-smart-summary span {
      display: block;
      color: var(--muted);
      font-size: 10px;
      font-weight: 800;
      text-transform: uppercase;
    }

    .add-case-smart-summary strong {
      display: block;
      margin-top: 3px;
      color: var(--ink);
      font-size: 12px;
      line-height: 1.35;
      overflow-wrap: anywhere;
    }

    .add-case-smart-summary small {
      display: block;
      margin-top: 2px;
      color: var(--muted);
      font-size: 10px;
      line-height: 1.35;
      overflow-wrap: anywhere;
    }

    @media (max-width: 960px) {
      .add-case-smart-summary {
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }
    }

    @media (max-width: 560px) {
      .add-case-smart-summary {
        grid-template-columns: 1fr;
      }
    }""",
    )
    html = html.replace(
        """    function setLinkedFlowStatus(id, text, state = "waiting") {
      const node = document.querySelector(`#${id}`);
      if (!node) return;
      node.textContent = text;
      node.classList.remove("waiting", "linked");
      if (state) node.classList.add(state);
    }""",
        """    function setLinkedFlowStatus(id, text, state = "waiting") {
      const node = document.querySelector(`#${id}`);
      if (!node) return;
      node.textContent = text;
      node.classList.remove("waiting", "linked");
      if (state) node.classList.add(state);
    }

    function setAccessLevelBadge(value = "") {
      const badge = document.querySelector("#addAccessLevelBadge");
      if (!badge) return;
      const normalized = String(value || "").trim();
      badge.textContent = normalized || "Waiting";
      badge.classList.toggle("normal", normalized === "Normal");
      badge.classList.toggle("confidential", normalized === "Confidential");
      badge.classList.toggle("waiting", !normalized);
    }

    function contractTypeMasterV2Display(row, level = "sub") {
      if (!row) return "";
      const enKey = level === "type" ? "Type of Contract EN" : "Sub Type of Contract EN";
      const thKey = level === "type" ? "Type of Contract TH" : "Sub Type of Contract TH";
      const en = String(row[enKey] || row["Type of Contract EN"] || "").trim();
      const th = String(row[thKey] || row["Type of Contract TH"] || "").trim();
      return [en, th].filter(Boolean).join(" / ");
    }

    function contractTypeMasterV2Candidates(row) {
      if (!row) return [];
      return [
        contractTypeMasterV2Display(row, "sub"),
        contractTypeMasterV2Display(row, "type"),
        row["Sub Type of Contract EN"],
        row["Sub Type of Contract TH"],
        row["Type of Contract EN"],
        row["Type of Contract TH"]
      ].map(value => String(value || "").trim()).filter(Boolean);
    }

    function contractTypeMasterV2Match(typeValue = "", classification = "") {
      const lookupValue = /service provider agreement/i.test(String(typeValue || ""))
        ? String(typeValue || "").replace(/service provider agreement/ig, "Outsourcing Agreement")
        : typeValue;
      const normalized = normalizeDirectoryValue(lookupValue);
      if (!normalized) return null;
      const classificationEn = classification ? classificationEnFromValue(classification) : "";
      const parts = String(lookupValue || "")
        .split(/[|/·•,;()\\[\\]–—-]+/)
        .map(normalizeDirectoryValue)
        .filter(Boolean);
      const scoreRow = row => {
        const weightedCandidates = [
          [contractTypeMasterV2Display(row, "sub"), 120],
          [row["Sub Type of Contract EN"], 115],
          [contractTypeMasterV2Display(row, "type"), 105],
          [row["Type of Contract EN"], 100],
          [row["Sub Type of Contract TH"], 90],
          [row["Type of Contract TH"], 80]
        ].map(([value, weight]) => [normalizeDirectoryValue(value), weight])
          .filter(([value]) => value);
        return weightedCandidates.reduce((best, [candidate, weight]) => {
          if (candidate === normalized || parts.includes(candidate)) return Math.max(best, weight + 100);
          if (normalized.includes(candidate)) return Math.max(best, weight + 20);
          if (candidate.includes(normalized)) return Math.max(best, weight + 5);
          return best;
        }, 0);
      };
      return contractTypeMasterV2
        .filter(row => !classificationEn || row["Contract Classification EN"] === classificationEn)
        .map(row => ({ row, score: scoreRow(row) }))
        .filter(item => item.score > 0)
        .sort((a, b) => b.score - a.score)[0]?.row || null;
    }

    function renderAddCasePracticalExample(typeValue = "", classification = "", awaitingSubType = false) {
      const enNode = document.querySelector("#addPracticalExampleEn");
      const thNode = document.querySelector("#addPracticalExampleTh");
      if (!enNode || !thNode) return;
      const typeInfo = typeValue && !awaitingSubType ? contractTypeMasterV2Match(typeValue, classification) : null;
      const effectiveType = String(typeInfo?.["Sub Type of Contract EN"] || typeInfo?.["Type of Contract EN"] || "").trim();
      const key = effectiveType ? `${classificationEnFromValue(classification)}|${effectiveType}` : "";
      const example = key ? practicalExampleConfig[key] : null;
      enNode.textContent = example?.en || (typeValue
        ? "Select a Sub Type to view the matching practical example."
        : "Select Type and Sub Type to view an example.");
      thNode.textContent = example?.th || (typeValue
        ? "เลือกประเภทย่อยเพื่อดูตัวอย่างที่ตรงกับสัญญา"
        : "เลือกประเภทและประเภทย่อยเพื่อดูตัวอย่าง");
    }

    function setAddCaseSummaryText(id, value, fallback = "-") {
      const node = document.querySelector(`#${id}`);
      if (!node) return;
      node.textContent = value || fallback;
    }

    function renderAddCaseSmartSummary() {
      const form = document.querySelector("#addCaseForm");
      if (!form) return;
      const selectedTemplate = contractTemplateFor(String(form.elements.name?.value || "").trim());
      const type = String(effectiveAddCaseContractType() || selectedTemplate?.type || "").trim();
	      const selectedClassification = selectedContractClassification();
	      const typeInfo = contractTypeMasterV2Match(type, selectedClassification);
      const classificationEn = typeInfo?.["Contract Classification EN"] || selectedClassification || (selectedTemplate?.accessLevel === "Confidential" ? "Confidential" : "Day-to-day Work");
      const classificationTh = typeInfo?.["Contract Classification TH"] || classificationThaiFor(classificationEn);
      const typeDisplay = selectedContractTypeGroup() || (typeInfo ? contractTypeMasterV2Display(typeInfo, "type") : type);
      const subTypeDisplay = selectedContractSubType() || (typeInfo ? subTypeForRow(typeInfo) : "");
      const contractName = selectedTemplate?.name || String(form.elements.name?.value || "").trim();
      const sla = String(form.elements.sla?.value || "").trim();
      const due = String(form.elements.due?.value || "").trim();
      setAddCaseSummaryText("addSummaryClassification", classificationEn || "Waiting");
      setAddCaseSummaryText("addSummaryClassificationTh", classificationTh || "รอข้อมูลประเภทสัญญา");
      setAddCaseSummaryText("addSummaryType", typeDisplay || "-");
      setAddCaseSummaryText("addSummaryTypeTh", typeDisplay ? "ประเภทสัญญา" : "เลือกประเภทสัญญา");
      setAddCaseSummaryText("addSummarySubType", subTypeDisplay || "No Sub Type");
      setAddCaseSummaryText("addSummarySubTypeTh", subTypeDisplay ? "ประเภทย่อยของสัญญา" : "ไม่มีประเภทย่อย");
      setAddCaseSummaryText("addSummarySla", sla ? `${sla} Working Days / ${sla} วันทำการ` : "-");
      setAddCaseSummaryText("addSummaryDue", due ? `Due ${due}` : "รอการคำนวณวันครบกำหนด");
      setAddCaseSummaryText("addSummaryContractName", contractName || "-");
      setAddCaseSummaryText("addSummaryDepartment", String(form.elements.department?.value || selectedTemplate?.department || "").trim());
      setAddCaseSummaryText("addSummaryVendor", String(form.elements.vendor?.value || selectedTemplate?.vendor || "").trim());
      setAddCaseSummaryText("addSummaryOwner", String(form.elements.owner?.value || "").trim());
      setAddCaseSummaryText("addSummaryContractId", document.querySelector("#newCaseId")?.textContent || "-");
    }""",
    )
    html = html.replace(
        """      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 12px;
      padding: 0;""",
        """      grid-template-columns: repeat(5, minmax(0, 1fr));
      gap: 12px;
      padding: 0;""",
    )
    html = html.replace(
        """      [...contractInputCatalog, ...contracts.map(item => ({ name: item.name, type: item.type, workType: item.workType }))]
        .forEach(item => {
          const name = String(item?.name || "").trim();
          if (!name) return;
          if (!merged.has(name)) merged.set(name, { name, type: item.type || "", workType: item.workType || "Other" });
        });""",
        """      contractInputCatalog
        .forEach(item => {
          const name = String(item?.name || "").trim();
          if (!name) return;
          if (!merged.has(name)) merged.set(name, {
            name,
            type: item.type || "",
            workType: item.workType || "Other",
            contractId: item.contractId || "",
            accessLevel: item.accessLevel || "",
            category: item.category || ""
          });
        });""",
    )
    html = html.replace(
        """    .filter-row {
      display: flex;
      align-items: center;
      gap: 8px;
      flex-wrap: wrap;
    }

    .input,""",
        """    .filter-row {
      display: flex;
      align-items: center;
      gap: 8px;
      flex-wrap: wrap;
    }

    .filter-th {
      min-width: 118px;
      vertical-align: top;
    }

    .filter-th details {
      position: relative;
      display: inline-block;
      width: 100%;
    }

    .filter-th summary {
      display: inline-flex;
      align-items: center;
      gap: 4px;
      width: 100%;
      cursor: pointer;
      list-style: none;
      user-select: none;
    }

    .filter-th summary::-webkit-details-marker {
      display: none;
    }

    .filter-th summary::after {
      content: "▾";
      margin-left: auto;
      width: 18px;
      height: 18px;
      border: 1px solid var(--line);
      border-radius: 6px;
      display: inline-grid;
      place-items: center;
      background: #ffffff;
      color: var(--muted);
      font-size: 10px;
      line-height: 1;
    }

    .filter-th details[open] summary::after {
      color: var(--ink);
      border-color: var(--green);
      background: #f8faf3;
    }

    .th-filter-popover {
      position: fixed;
      z-index: 900;
      top: var(--filter-popover-top, 0);
      left: var(--filter-popover-left, 0);
      min-width: 185px;
      max-width: min(320px, calc(100vw - 24px));
      padding: 8px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #ffffff;
      box-shadow: 0 14px 32px rgba(15, 23, 42, .16);
    }

    .th-filter-popover .select,
    .th-filter-popover .input {
      width: 100%;
      min-height: 30px;
      font-size: 12px;
      padding: 6px 8px;
    }

    .th-filter-popover .secondary-button {
      width: 100%;
      min-height: 28px;
      margin-top: 6px;
      border-radius: 7px;
      font-size: 11px;
    }

    .due-date-range-popover {
      min-width: 172px;
      display: grid;
      gap: 7px;
    }

    .due-date-range-row {
      display: grid;
      gap: 3px;
      font-size: 10px;
      color: var(--muted);
      line-height: 1.2;
    }

    .due-date-range-row .input {
      margin: 0;
    }

    .database-sync-strip {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
      margin: 0 16px 12px;
      padding: 10px 12px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #f8faf3;
    }

    .database-sync-title {
      display: grid;
      gap: 2px;
      min-width: 190px;
    }

    .database-sync-title strong {
      color: var(--ink);
      font-size: 13px;
    }

    .database-sync-title span {
      color: var(--muted);
      font-size: 11px;
    }

    .database-sync-actions {
      display: flex;
      align-items: center;
      justify-content: flex-end;
      gap: 8px;
      flex-wrap: wrap;
    }

    .database-sync-actions .secondary-button,
    .database-import-label {
      min-height: 34px;
      border-radius: 8px;
      font-size: 12px;
      white-space: nowrap;
    }

    .database-import-label input {
      display: none;
    }

    @media (max-width: 760px) {
      .database-sync-strip {
        align-items: stretch;
        flex-direction: column;
      }

      .database-sync-actions {
        justify-content: stretch;
      }

      .database-sync-actions .secondary-button,
      .database-import-label {
        flex: 1 1 145px;
      }
    }

    .input,""",
    )
    html = html.replace(
        """    .combo-option-meta {
      color: var(--muted);
      font-size: 11px;
      line-height: 1.35;
      overflow-wrap: anywhere;
    }

    .combo-empty {""",
        """    .combo-option-meta {
      color: var(--muted);
      font-size: 11px;
      line-height: 1.35;
      overflow-wrap: anywhere;
    }

    .combo-option-description {
      display: none;
      margin-top: 3px;
      padding-top: 6px;
      border-top: 1px solid var(--line);
      color: var(--muted);
      font-size: 11px;
      line-height: 1.45;
      white-space: pre-line;
      overflow-wrap: anywhere;
    }

    .combo-option:hover .combo-option-description,
    .combo-option.active .combo-option-description {
      display: block;
    }

    .master-data-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 16px;
      align-items: start;
    }

    .master-data-panel {
      min-width: 0;
    }

    .master-data-panel-full {
      grid-column: 1 / -1;
    }

    .master-data-actions {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      justify-content: flex-end;
    }

    .master-table {
      min-width: 720px;
    }

    .master-table th,
    .master-table td {
      vertical-align: top;
    }

    .master-table .input,
    .master-table .select {
      min-height: 34px;
      padding: 7px 8px;
      font-size: 12px;
    }

    .master-table textarea.input {
      min-height: 58px;
      resize: vertical;
    }

    .master-table .icon-button {
      width: 30px;
      height: 30px;
    }

    @media (max-width: 1040px) {
      .master-data-grid {
        grid-template-columns: 1fr;
      }
    }

    .combo-empty {""",
    )
    html = html.replace(
        """      if (!state || !item) return;
      state.input.value = item.value;""",
        """      if (!state || !item) return;
      if (state.input.disabled) return;
      state.input.value = item.value;""",
    )
    html = html.replace(
        """      if (!state) return;
      closeAllEditableDropdowns(inputId);""",
        """      if (!state || state.input.disabled) return;
      closeAllEditableDropdowns(inputId);""",
    )
    html = html.replace(
        """          <button type="button" class="combo-option" role="option" data-index="${index}" aria-selected="false">
            <span class="combo-option-main">${escapeHtml(item.primary || item.value)}</span>
            ${item.secondary ? `<span class="combo-option-meta">${escapeHtml(item.secondary)}</span>` : ""}
          </button>`).join("")""",
        """          <button type="button" class="combo-option" role="option" data-index="${index}" aria-selected="false" title="${escapeHtml(item.description || item.secondary || item.primary || item.value)}">
            <span class="combo-option-main">${escapeHtml(item.primary || item.value)}</span>
            ${item.secondary ? `<span class="combo-option-meta">${escapeHtml(item.secondary)}</span>` : ""}
            ${item.description ? `<span class="combo-option-description">${escapeHtml(item.description)}</span>` : ""}
          </button>`).join("")""",
    )
    html = html.replace(
        """    function actionDropdownOptions() {
      return updateActionDefinitions.map(item => ({""",
        """    function isMasterActive(value) {
      return !["no", "false", "inactive", "0"].includes(normalizeDirectoryValue(value));
    }

    function activeMasterContractTypes() {
      return (masterData.contractTypes || []).filter(row => masterContractTypeValue(row) && isMasterActive(row.Active));
    }

    function masterContractTypeValue(row = {}) {
      return String(row["Type of Contract"] || row["Sub Type of Contract"] || "").trim();
    }

    function contractPrimaryTypeDisplay(typeValue = "") {
      const typeInfo = contractTypeMasterV2Match(typeValue);
      return typeInfo ? typeGroupForRow(typeInfo) : String(typeValue || "").trim();
    }

    function contractSubTypeDisplay(typeValue = "") {
      const typeInfo = contractTypeMasterV2Match(typeValue);
      return typeInfo ? subTypeForRow(typeInfo) : "";
    }

    function activeMasterContractTemplates() {
      return (masterData.contractTemplates || []).filter(item => item.name && isMasterActive(item.active));
    }

    function activeMasterPeople() {
      return (masterData.people || []).filter(item => item.name && isMasterActive(item.active));
    }

    function activeMasterDepartments() {
      return (masterData.departments || []).filter(item => item["Department / Restaurant"] && isMasterActive(item.Active));
    }

    function allPeopleDirectory() {
      const merged = new Map();
      [...employeeDirectory, ...activeMasterPeople()].forEach(item => {
        const email = String(item.email || "").trim().toLowerCase();
        const key = email || normalizeDirectoryValue(item.name);
        if (!key) return;
        merged.set(key, { ...item });
      });
      return [...merged.values()];
    }

    function contractTypeDescriptionFor(typeValue) {
      const normalized = normalizeDirectoryValue(typeValue);
      if (!normalized) return "";
      const typeInfo = contractTypeMasterV2Match(typeValue);
      if (typeInfo) {
        return [
          `${typeInfo["Contract Classification EN"] || ""} / ${typeInfo["Contract Classification TH"] || ""}`.trim(),
          contractTypeMasterV2Display(typeInfo, "type"),
          contractTypeMasterV2Display(typeInfo, "sub")
        ].filter(Boolean).join("\\n");
      }
      const rows = activeMasterContractTypes();
      const match = rows.find(row => normalizeDirectoryValue(masterContractTypeValue(row)) === normalized);
      return match?.["Description / คำอธิบาย"] || "";
    }

    function contractTypeCategoryFor(typeValue) {
      const normalized = normalizeDirectoryValue(typeValue);
      if (!normalized) return "";
      const typeInfo = contractTypeMasterV2Match(typeValue);
      if (typeInfo) return typeInfo["Contract Classification EN"] || "";
      const rows = activeMasterContractTypes();
      const match = rows.find(row => normalizeDirectoryValue(masterContractTypeValue(row)) === normalized);
      return match?.Category || "";
    }

    function accessLevelForContractType(typeValue) {
      const category = contractTypeCategoryFor(typeValue);
      return normalizeDirectoryValue(category) === "confidential" ? "Confidential" : "Normal";
    }

    function fixedSlaValue(value) {
      const match = String(value ?? "").trim().match(/\\d+(?:\\.\\d+)?/);
      const number = match ? Number(match[0]) : 0;
      return Number.isFinite(number) && number > 0 ? number : 0;
    }

    function standardSlaFromContractTypeMasterV2(typeValue = "", classification = "") {
      const typeInfo = contractTypeMasterV2Match(typeValue, classification);
      return fixedSlaValue(
        typeInfo?.["Standard SLA"] ||
        typeInfo?.["Standard SLA / SLA รวม"] ||
        typeInfo?.["Standard SLA (Working Days)"] ||
        typeInfo?.["Total SLA / SLA รวม"]
      );
    }

    function migrateMasterDataToCurrentSlaVersion(savedVersion = "") {
      let migrated = 0;
      (masterData.contractTypes || []).forEach(row => {
        const rowVersion = String(row["Standard SLA Version"] || savedVersion || "").trim();
        if (rowVersion === standardSlaDataVersion) return;
        const classification = String(row["Contract Classification"] || row.Category || "").trim();
        const typeValue = String(row["Sub Type of Contract"] || row["Type of Contract"] || "").trim();
        const typeInfo = contractTypeMasterV2Match(typeValue, classification);
        const currentSla = standardSlaFromContractTypeMasterV2(typeValue, classification);
        if (!typeInfo || !currentSla) return;
        const classificationDisplay = [typeInfo["Contract Classification EN"], typeInfo["Contract Classification TH"]].filter(Boolean).join(" / ");
        const typeDisplay = [typeInfo["Type of Contract EN"], typeInfo["Type of Contract TH"]].filter(Boolean).join(" / ");
        const subTypeDisplay = [typeInfo["Sub Type of Contract EN"], typeInfo["Sub Type of Contract TH"]].filter(Boolean).join(" / ");
        row["Contract Classification"] = classificationDisplay;
        row.Category = classificationDisplay;
        row["Type of Contract"] = typeDisplay;
        row["Sub Type of Contract"] = subTypeDisplay;
        row["Fixed SLA (Working Days)"] = String(currentSla);
        row["Standard SLA Version"] = standardSlaDataVersion;
        row["Description / คำอธิบาย"] = typeDisplay;
        migrated += 1;
      });
      (masterData.contractTemplates || []).forEach(row => {
        const rowVersion = String(row.slaVersion || savedVersion || "").trim();
        if (rowVersion === standardSlaDataVersion) return;
        const classification = String(row.classification || row.category || "").trim();
        const typeValue = String(row.subType || row.typeGroup || row.type || row.workType || "").trim();
        const currentSla = standardSlaFromContractTypeMasterV2(typeValue, classification);
        if (!currentSla) return;
        row.fixedSla = String(currentSla);
        row.slaVersion = standardSlaDataVersion;
        migrated += 1;
      });
      return migrated;
    }

    function migrateActionMasterToCurrentVersion(savedVersion = "") {
      const rows = masterData.actionSla || [];
      const allRowsCurrent = rows.length === updateActionList.length && rows.every(row =>
        String(row["Action Data Version"] || "").trim() === actionDataVersion
      );
      if (savedVersion === actionDataVersion || allRowsCurrent) return 0;

      const previousByAction = new Map(rows.map(row => [String(row.Action || "").trim(), row]));
      masterData.actionSla = updateActionList.map(action => {
        const previous = previousByAction.get(action) || {};
        const configured = actionDescriptionConfig[action] || {};
        return {
          ...previous,
          Action: action,
          "Description / รายละเอียด": configured.descriptionTh || "",
          "Fixed SLA (Working Days)": String(actionSlaConfig[action] || ""),
          "SLA Rule / วิธีนับ": configured.slaRuleTh || "",
          "Action Data Version": actionDataVersion,
          Active: previous.Active || "Yes"
        };
      });
      return masterData.actionSla.length;
    }

    function canonicalDepartmentName(value = "") {
      const name = String(value || "").trim();
      if (!name || name === "0") return "";
      if (departmentNameAliases[name]) return departmentNameAliases[name];
      const normalized = normalizeDirectoryValue(name);
      const alias = Object.keys(departmentNameAliases)
        .find(item => normalizeDirectoryValue(item) === normalized);
      return alias ? departmentNameAliases[alias] : name;
    }

    function migrateDepartmentMasterToCurrentVersion(savedVersion = "") {
      const currentRows = masterData.departments || [];
      const allRowsCurrent = currentRows.length > 0 && currentRows.every(row =>
        String(row["Department Data Version"] || "").trim() === departmentDataVersion
      );
      if (savedVersion === departmentDataVersion || allRowsCurrent) return 0;

      const previousCount = currentRows.length;
      const codeByDepartment = new Map();
      currentRows.forEach(row => {
        const name = canonicalDepartmentName(row["Department / Restaurant"]);
        const code = String(row["Department Code"] || "").trim().toUpperCase();
        if (name && code && code !== "0" && !codeByDepartment.has(name)) codeByDepartment.set(name, code);
      });

      (masterData.people || []).forEach(row => {
        row.department = canonicalDepartmentName(row.department);
      });
      const departmentsWithPeople = orderedUniqueList(
        (masterData.people || [])
          .filter(row => isMasterActive(row.active) && row.name)
          .map(row => canonicalDepartmentName(row.department))
          .filter(Boolean)
      );
      masterData.departments = departmentsWithPeople.map(name => ({
        "Department / Restaurant": name,
        "Department Code": codeByDepartment.get(name) || String(departmentCodeConfig[name] || departmentCodeSuggestion(name) || "").trim().toUpperCase(),
        "Department Data Version": departmentDataVersion,
        Active: "Yes"
      })).filter(row => row["Department Code"] && row["Department Code"] !== "0");

      (masterData.contractTemplates || []).forEach(row => {
        row.department = canonicalDepartmentName(row.department);
      });
      return Math.max(1, Math.abs(previousCount - masterData.departments.length));
    }

	    function fixedSlaFromMasterData(workType, contractName = "") {
	      const normalizedType = normalizeDirectoryValue(workType);
	      const normalizedName = normalizeDirectoryValue(contractName);
	      if (normalizedName) {
	        const template = contractTemplateFor(contractName) || activeMasterContractTemplates().find(item => normalizeDirectoryValue(item.name) === normalizedName);
        const templateSla = fixedSlaValue(template?.fixedSla || template?.["Fixed SLA (Working Days)"] || template?.SLA);
        if (templateSla > 0) return templateSla;
      }
      if (normalizedType) {
        const typeRow = activeMasterContractTypes().find(row => {
          const candidates = [
            masterContractTypeValue(row),
            row["Sub Type of Contract"],
            row["Description / คำอธิบาย"]
          ];
          return candidates.some(value => normalizeDirectoryValue(value) === normalizedType || contractTypeValuesMatch(value, workType));
        });
        const typeSla = fixedSlaValue(typeRow?.["Fixed SLA (Working Days)"] || typeRow?.FixedSLA || typeRow?.SLA);
        if (typeSla > 0) return typeSla;
        const standardSla = standardSlaFromContractTypeMasterV2(workType);
        if (standardSla > 0) return standardSla;
        if (typeRow || contractTypeMasterV2Match(workType)) return 0;
      }
      return null;
    }

    function accessLevelForAddCase(template, contractType) {
      const typeAccessLevel = accessLevelForContractType(contractType);
      if (typeAccessLevel === "Confidential" || template?.accessLevel === "Confidential") return "Confidential";
      return typeAccessLevel || template?.accessLevel || "Normal";
    }

    function contractTypeDropdownOptions() {
      const seen = new Set();
      const options = [];
      contractTypeMasterV2.forEach(row => {
        const value = contractTypeMasterV2Display(row, "sub") || contractTypeMasterV2Display(row, "type");
        const key = normalizeDirectoryValue(value);
        if (!value || seen.has(key)) return;
        seen.add(key);
        const classification = [row["Contract Classification EN"], row["Contract Classification TH"]].filter(Boolean).join(" / ");
        const typeGroup = contractTypeMasterV2Display(row, "type");
        options.push({
          value,
          primary: value,
          secondary: [classification, typeGroup && typeGroup !== value ? typeGroup : ""].filter(Boolean).join(" · "),
          description: [typeGroup, contractTypeMasterV2Display(row, "sub")].filter(Boolean).join("\\n")
        });
      });
      const rows = realWorkbookData?.sheets?.["Contract Type Master"] || [];
      rows
        .filter(row => row["Type of Contract"])
        .forEach(row => {
          const value = row["Type of Contract"];
          const key = normalizeDirectoryValue(value);
          if (seen.has(key)) return;
          seen.add(key);
          options.push({
            value,
            primary: value,
            secondary: row.Category || "Contract Type",
            description: row["Description / คำอธิบาย"] || ""
          });
        });
      return options;
    }

    function selectedContractGroup() {
      return String(document.querySelector("#addContractGroup")?.value || "").trim();
    }

    function groupContractDropdownOptions() {
      return orderedUniqueList(getContractFormCatalog().map(item => item.group || item.category).filter(Boolean))
        .map(group => {
          const count = getContractFormCatalog().filter(item => (item.group || item.category) === group).length;
          return {
            value: group,
            primary: group,
            secondary: `${count} Contract Name${count === 1 ? "" : "s"}`
          };
        });
    }

    function contractNameDropdownOptions() {
      const selectedGroup = selectedContractGroup();
      return getContractFormCatalog()
        .filter(item => !selectedGroup || (item.group || item.category) === selectedGroup)
        .map(item => {
          const access = item.accessLevel || accessLevelForContractType(item.type) || "Normal";
          return {
            value: item.selectionLabel || item.name,
            primary: item.name,
            secondary: [item.department, item.type, item.fixedSla ? `SLA ${item.fixedSla}` : "", access].filter(Boolean).join(" · "),
            description: [item.group || item.category, item.vendor ? `Vendor: ${item.vendor}` : "", item.contractId ? `Contract ID: ${item.contractId}` : ""].filter(Boolean).join("\\n")
          };
        })
        .sort((a, b) => {
          const rank = value => String(value.secondary || "").startsWith("Confidential") ? 1 : 0;
          return rank(a) - rank(b) || String(a.primary || "").localeCompare(String(b.primary || ""));
        });
    }

    function actionDropdownOptions() {
	      return activeUpdateActionDefinitions.map(item => ({""",
    )
    html = html.replace(
        """      const rows = realWorkbookData?.sheets?.["Contract Type Master"] || [];""",
        """      const rows = activeMasterContractTypes();""",
    )
    html = html.replace(
        """    function totalSlaFor(workType) {
      const normalized = String(workType || "").trim();
      if (Object.prototype.hasOwnProperty.call(fixedSlaConfig, normalized)) {
        return Number(fixedSlaConfig[normalized]);
      }
      return Number(fixedSlaConfig.Other);
    }""",
	        """    function totalSlaFor(workType, contractName = "", classification = "") {
	      const masterSla = fixedSlaFromMasterData(workType, contractName, classification);
	      if (masterSla !== null && masterSla !== undefined) return Number(masterSla) || 0;
	      const normalized = String(workType || "").trim();
	      if (Object.prototype.hasOwnProperty.call(fixedSlaConfig, normalized)) {
	        return Number(fixedSlaConfig[normalized]);
	      }
	      return Number(fixedSlaConfig.Other);
    }""",
    )
    html = html.replace(
        """    function actionSlaFor(action) {
      const normalizedAction = String(action || "").trim();
      if (Object.prototype.hasOwnProperty.call(actionSlaConfig, normalizedAction)) {
        return Number(actionSlaConfig[normalizedAction]);
      }
      return 0;
    }""",
        """    function activeMasterActionSlaRows() {
      return (masterData.actionSla || []).filter(row => row.Action && isMasterActive(row.Active));
    }

    function actionSlaFor(action) {
      const normalizedAction = String(action || "").trim();
      const masterRow = activeMasterActionSlaRows().find(row => String(row.Action || "").trim() === normalizedAction);
      const masterSla = fixedSlaValue(masterRow?.["Fixed SLA (Working Days)"] || masterRow?.sla);
      if (masterSla > 0) return masterSla;
      if (Object.prototype.hasOwnProperty.call(actionSlaConfig, normalizedAction)) {
        return Number(actionSlaConfig[normalizedAction]);
      }
      return 0;
    }""",
    )
    html = replace_between(
        html,
        """    function getContractFormCatalog() {""",
        """    function contractSummaryMarkup(contract, options = {}) {""",
        """    function getContractFormCatalog() {
      const rows = [];
      const seen = new Set();
      [...contractInputCatalog, ...activeMasterContractTemplates()]
        .forEach(item => {
          const name = String(item?.name || "").trim();
          if (!name) return;
          const selectionLabel = String(item.selectionLabel || item.displayName || name).trim();
          const classification = item.classification || item.category || item.accessLevel || "";
          const typeValue = item.subType || item.type || item.typeGroup || "";
          const typeInfo = contractTypeMasterV2Match(typeValue, classification);
          const typeGroup = item.typeGroup || (typeInfo ? typeGroupForRow(typeInfo) : contractPrimaryTypeDisplay(item.type));
          const subType = item.subType || (typeInfo ? subTypeForRow(typeInfo) : contractSubTypeDisplay(item.type));
          const fixedSla = standardSlaFromContractTypeMasterV2(subType || typeGroup || item.type, classification) || "";
          const key = normalizeDirectoryValue([
            selectionLabel,
            item.sourceRow,
            item.vendor,
            fixedSla,
            item.type
          ].filter(Boolean).join("|"));
          if (seen.has(key)) return;
          seen.add(key);
          rows.push({
            ...item,
            name,
            selectionLabel,
            type: subType || typeGroup || item.type || "",
            workType: subType || typeGroup || item.workType || item.type || "Other",
            contractId: item.contractId || "",
            accessLevel: item.accessLevel || "",
            category: item.category || "",
            department: item.department || "",
            vendor: item.vendor || "",
            group: typeGroup || item.group || item.category || "",
            typeGroup,
            subType,
            fixedSla,
            remark: item.remark || "",
            sourceRow: item.sourceRow || ""
          });
        });
      const nameCounts = rows.reduce((map, item) => {
        const key = normalizeDirectoryValue(item.name);
        map.set(key, (map.get(key) || 0) + 1);
        return map;
      }, new Map());
      rows.forEach(item => {
        const key = normalizeDirectoryValue(item.name);
        if (nameCounts.get(key) > 1 && normalizeDirectoryValue(item.selectionLabel) === key) {
          item.selectionLabel = [item.name, item.vendor, item.fixedSla ? `SLA ${item.fixedSla}` : ""].filter(Boolean).join(" — ");
        }
      });
      return rows.sort((a, b) => String(a.selectionLabel || a.name).localeCompare(String(b.selectionLabel || b.name)));
    }

    function contractTemplateFor(value) {
      const normalized = normalizeDirectoryValue(value);
      if (!normalized) return null;
      const catalog = getContractFormCatalog();
      return catalog.find(item => normalizeDirectoryValue(item.selectionLabel || item.name) === normalized)
        || catalog.find(item => normalizeDirectoryValue(item.name) === normalized)
        || null;
    }

""",
    )
    html = replace_between(
        html,
        """      // Type of Contract links to the hidden Work Type, which determines the fixed SLA.""",
        """      form.elements.sla.value = sla;""",
        """      // Fixed SLA follows the selected Type/Sub Type master; Contract Name only helps choose the template.
      const contractName = String(form.elements.name?.value || "").trim();
      const sla = totalSlaFor(workType, contractName, selectedContractClassification());
      const systemDue = addBusinessDays(lockedInDate, sla);
""",
    )
    html = replace_between(
        html,
        """	    function fixedSlaFromMasterData(workType, contractName = "") {""",
        """

    function accessLevelForAddCase(template, contractType) {""",
        """    function englishContractPart(value) {
      return String(value || "").split("|")[0].split("/")[0].trim();
    }

		    function masterContractTypeFor(typeValue, classification = "") {
		      const normalizedType = normalizeDirectoryValue(typeValue);
		      if (!normalizedType) return null;
		      const classificationEn = classification ? classificationEnFromValue(classification) : "";
		      const rows = activeMasterContractTypes().filter(row => {
		        if (!classificationEn) return true;
		        return classificationEnFromValue(row["Contract Classification"] || row.Category) === classificationEn;
		      });
	      const normalizedEnglish = normalizeDirectoryValue(englishContractPart(typeValue));
	      const exactSubType = rows.find(row => {
	        const subType = row["Sub Type of Contract"];
	        if (!subType) return false;
	        return [subType, englishContractPart(subType)]
	          .some(value => normalizeDirectoryValue(value) === normalizedType || normalizeDirectoryValue(value) === normalizedEnglish);
	      });
	      if (exactSubType) return exactSubType;
	      const exactTypeWithoutSubType = rows.find(row => {
	        if (String(row["Sub Type of Contract"] || "").trim()) return false;
	        const type = row["Type of Contract"];
	        return [type, englishContractPart(type)]
	          .some(value => normalizeDirectoryValue(value) === normalizedType || normalizeDirectoryValue(value) === normalizedEnglish);
	      });
	      if (exactTypeWithoutSubType) return exactTypeWithoutSubType;
	      const exact = rows.find(row => [masterContractTypeValue(row), row["Description / คำอธิบาย"]]
	        .some(value => normalizeDirectoryValue(value) === normalizedType || normalizeDirectoryValue(englishContractPart(value)) === normalizedEnglish));
	      if (exact) return exact;
		      const typeInfo = contractTypeMasterV2Match(typeValue, classification);
	      const englishCandidates = [
	        englishContractPart(typeValue),
	        typeInfo?.["Sub Type of Contract EN"],
	        typeInfo?.["Type of Contract EN"]
	      ].map(normalizeDirectoryValue).filter(Boolean);
	      const bySubType = rows.find(row => {
	        const rowEnglishValues = [row["Sub Type of Contract"]]
	          .map(englishContractPart)
	          .map(normalizeDirectoryValue)
	          .filter(Boolean);
	        return rowEnglishValues.some(value => englishCandidates.includes(value));
	      });
	      if (bySubType) return bySubType;
	      return rows.find(row => {
	        const rowEnglishValues = [masterContractTypeValue(row), row["Description / คำอธิบาย"]]
	          .map(englishContractPart)
	          .map(normalizeDirectoryValue)
	          .filter(Boolean);
	        return rowEnglishValues.some(value => englishCandidates.includes(value));
	      }) || null;
	    }

    function fixedSlaFromMasterData(workType, contractName = "", classification = "") {
      const normalizedType = normalizeDirectoryValue(workType);
      if (normalizedType) {
        const typeRow = masterContractTypeFor(workType, classification);
        const typeSla = fixedSlaValue(typeRow?.["Fixed SLA (Working Days)"] || typeRow?.FixedSLA || typeRow?.SLA);
        if (typeSla > 0) return typeSla;
        const standardSla = standardSlaFromContractTypeMasterV2(workType, classification);
        if (standardSla > 0) return standardSla;
        if (typeRow || contractTypeMasterV2Match(workType, classification)) return 0;
      }
      const normalizedName = normalizeDirectoryValue(contractName);
      if (normalizedName) {
        const template = contractTemplateFor(contractName) || activeMasterContractTemplates().find(item => normalizeDirectoryValue(item.name) === normalizedName);
        const templateSla = fixedSlaValue(template?.fixedSla || template?.["Fixed SLA (Working Days)"] || template?.SLA);
        if (templateSla > 0) return templateSla;
      }
      return null;
    }""",
    )
    html = replace_between(
        html,
        """    function accessLevelForAddCase(template, contractType) {""",
        """

    function actionDropdownOptions() {""",
        """    function classificationEnFromValue(classification = "") {
      const text = String(classification || "").trim();
      const english = text.split("|")[0].split("/")[0].trim();
      const normalized = normalizeDirectoryValue(english || text);
      return normalized === "confidential" ? "Confidential" : "Day-to-day Work";
    }

    function classificationDisplayValue(classification = "") {
      const en = classificationEnFromValue(classification);
      return `${en} / ${classificationThaiFor(en)}`;
    }

    function classificationAccessLevelFor(classification = "") {
      return classificationEnFromValue(classification) === "Confidential" ? "Confidential" : "Normal";
    }

    function classificationThaiFor(classification = "") {
      return classificationEnFromValue(classification) === "Confidential" ? "สัญญาลับ" : "งานดำเนินงานทั่วไป";
    }

    function selectedContractClassification() {
      return classificationEnFromValue(document.querySelector("#addContractClassification")?.value || "Day-to-day Work");
    }

    function contractClassificationDropdownOptions() {
      return [
        { value: "Day-to-day Work / งานดำเนินงานทั่วไป", primary: "Day-to-day Work", secondary: "งานดำเนินงานทั่วไป" },
        { value: "Confidential / สัญญาลับ", primary: "Confidential", secondary: "สัญญาลับ" }
      ];
    }

    function typeGroupForRow(row) {
      return contractTypeMasterV2Display(row, "type") || contractTypeMasterV2Display(row, "sub");
    }

    function typeGroupEnglishForRow(row) {
      return String(row?.["Type of Contract EN"] || row?.["Sub Type of Contract EN"] || "").trim();
    }

    function typeGroupThaiForRow(row) {
      return String(row?.["Type of Contract TH"] || row?.["Sub Type of Contract TH"] || "").trim();
    }

    function contractTypeShortDescriptionTh(typeValue = "") {
      const descriptions = {
        "Lease & Rental Agreement": "สัญญาเช่า ให้เช่า หรือเช่าช่วง"
      };
      return descriptions[englishContractPart(typeValue)] || "";
    }

    function subTypeForRow(row) {
      const subType = contractTypeMasterV2Display(row, "sub");
      const typeGroup = typeGroupForRow(row);
      return subType && normalizeDirectoryValue(subType) !== normalizeDirectoryValue(typeGroup) ? subType : "";
    }

    function subTypeEnglishForRow(row) {
      const subType = String(row?.["Sub Type of Contract EN"] || "").trim();
      const typeGroup = typeGroupEnglishForRow(row);
      return subType && normalizeDirectoryValue(subType) !== normalizeDirectoryValue(typeGroup) ? subType : "";
    }

    function subTypeThaiForRow(row) {
      const subType = String(row?.["Sub Type of Contract TH"] || "").trim();
      const typeGroup = typeGroupThaiForRow(row);
      return subType && normalizeDirectoryValue(subType) !== normalizeDirectoryValue(typeGroup) ? subType : "";
    }

    function contractTypeValuesMatch(left = "", right = "") {
      const leftNorm = normalizeDirectoryValue(left);
      const rightNorm = normalizeDirectoryValue(right);
      if (!leftNorm || !rightNorm) return false;
      if (leftNorm === rightNorm || leftNorm.includes(rightNorm) || rightNorm.includes(leftNorm)) return true;
      const leftEnglish = normalizeDirectoryValue(englishContractPart(left));
      const rightEnglish = normalizeDirectoryValue(englishContractPart(right));
      return Boolean(leftEnglish && rightEnglish && leftEnglish === rightEnglish);
    }

    function contractTypeRowsForSelection() {
      const classification = selectedContractClassification();
      const typeGroup = selectedContractTypeGroup();
      return contractTypeMasterV2.filter(row => {
        const rowClassification = String(row["Contract Classification EN"] || "").trim();
        if (classification && rowClassification !== classification) return false;
        if (typeGroup && !contractTypeValuesMatch(typeGroupForRow(row), typeGroup)) return false;
        return true;
      });
    }

    function setContractClassificationButtons(classification = selectedContractClassification()) {
      const normalized = classificationEnFromValue(classification || "Day-to-day Work");
      const input = document.querySelector("#addContractClassification");
      if (input) input.value = classificationDisplayValue(normalized);
      document.querySelectorAll("[data-classification-value]").forEach(button => {
        const active = button.dataset.classificationValue === normalized;
        button.classList.toggle("active", active);
        button.setAttribute("aria-pressed", String(active));
      });
      const accessLevel = classificationAccessLevelFor(normalized);
      const accessInput = document.querySelector("#addAccessLevel");
      if (accessInput) accessInput.value = accessLevel;
      setAccessLevelBadge(accessLevel);
      setLinkedFlowStatus("linkedAccessStatus", `${normalized} · ${classificationThaiFor(normalized)}`, accessLevel === "Confidential" ? "waiting" : "linked");
    }

    function accessLevelForAddCase(template, contractType) {
      const selectedLevel = classificationAccessLevelFor(selectedContractClassification());
      const typeAccessLevel = accessLevelForContractType(contractType);
      if (selectedLevel === "Confidential" || typeAccessLevel === "Confidential" || template?.accessLevel === "Confidential") return "Confidential";
      return selectedLevel || typeAccessLevel || template?.accessLevel || "Normal";
    }

    function contractTypeDropdownOptions() {
      const classification = selectedContractClassification();
      const seen = new Set();
      const options = [];
      contractTypeMasterV2
        .filter(row => !classification || row["Contract Classification EN"] === classification)
        .forEach(row => {
          const value = typeGroupEnglishForRow(row);
          const key = normalizeDirectoryValue(value);
          if (!value || seen.has(key)) return;
          seen.add(key);
          options.push({
            value,
            primary: value,
            secondary: "",
            description: "",
            translationTh: typeGroupThaiForRow(row)
          });
        });
      return options;
    }

    function selectedContractTypeGroup() {
      return String(document.querySelector("#addContractType")?.value || "").trim();
    }

    function selectedContractSubType() {
      return String(document.querySelector("#addContractSubType")?.value || "").trim();
    }

    function contractSubTypeDropdownOptions() {
      const rows = contractTypeRowsForSelection();
      const seen = new Set();
      return rows.map(row => {
        const value = subTypeEnglishForRow(row);
        if (!value) return null;
        const key = normalizeDirectoryValue(value);
        if (seen.has(key)) return null;
        seen.add(key);
        return {
          value,
          primary: value,
          secondary: "",
          description: "",
          translationTh: subTypeThaiForRow(row)
        };
      }).filter(Boolean);
    }

    function selectedContractGroup() {
      return "";
    }

    function groupContractDropdownOptions() {
      return [];
    }

    function effectiveAddCaseContractType() {
      return selectedContractSubType() || selectedContractTypeGroup();
    }

    function templateMatchesAddCaseFilters(item) {
      const classification = selectedContractClassification();
      const typeGroup = selectedContractTypeGroup();
      const subType = selectedContractSubType();
      const itemType = String(item?.type || "").trim();
      const itemInfo = contractTypeMasterV2Match(itemType);
      const itemClassification = itemInfo?.["Contract Classification EN"]
        || (item.accessLevel === "Confidential" ? "Confidential" : "Day-to-day Work");
      if (classification && itemClassification !== classification) return false;
      if (typeGroup) {
        const itemTypeGroup = itemInfo ? typeGroupForRow(itemInfo) : itemType;
        if (!contractTypeValuesMatch(itemTypeGroup, typeGroup) && !contractTypeValuesMatch(itemType, typeGroup)) return false;
      }
      if (subType) {
        const itemSubType = itemInfo ? (subTypeForRow(itemInfo) || typeGroupForRow(itemInfo)) : itemType;
        if (!contractTypeValuesMatch(itemSubType, subType) && !contractTypeValuesMatch(itemType, subType)) return false;
      }
      return true;
    }

    function contractNameDropdownOptions() {
      return getContractFormCatalog()
        .filter(templateMatchesAddCaseFilters)
        .map(item => {
          const access = item.accessLevel || accessLevelForContractType(item.type) || "Normal";
          return {
            value: item.selectionLabel || item.name,
            primary: item.name,
            secondary: [item.department, item.type, item.fixedSla ? `SLA ${item.fixedSla}` : "", access].filter(Boolean).join(" · "),
            description: [item.group || item.category, item.vendor ? `Vendor: ${item.vendor}` : "", item.contractId ? `Contract ID: ${item.contractId}` : ""].filter(Boolean).join("\\n")
          };
        })
        .sort((a, b) => String(a.primary || "").localeCompare(String(b.primary || "")));
    }""",
    )
    html = html.replace(
        """      const totalSla = totalSlaFor(workType);
      const lockedInDate = todayInputValue();""",
	        """      const totalSla = totalSlaFor(workType, String(form.get("name") || "").trim(), classificationEnFromValue(String(form.get("classification") || "")));
      const lockedInDate = todayInputValue();""",
    )
    html = html.replace(
        """      const owner = String(form.get("owner") || "").trim();
      const to = String(form.get("to") || "").trim();""",
        """      const owner = String(form.get("owner") || "").trim();
      const to = String(form.get("to") || "").trim();
      const selectedAddTemplate = contractTemplateFor(String(form.get("name") || "").trim());
      const contractName = selectedAddTemplate?.name || String(form.get("name") || "").trim();""",
    )
    html = html.replace(
        """      const contractType = String(form.get("type") || "").trim();
      const accessLevel = String(form.get("accessLevel") || accessLevelForContractType(contractType) || "Normal").trim();""",
        """      const contractType = String(form.get("type") || "").trim();
      const selectedAddTemplate = contractTemplateFor(String(form.get("name") || "").trim());
      const contractName = selectedAddTemplate?.name || String(form.get("name") || "").trim();
      const accessLevel = String(form.get("accessLevel") || selectedAddTemplate?.accessLevel || accessLevelForContractType(contractType) || "Normal").trim();""",
    )
    html = html.replace(
        """      const totalSla = totalSlaFor(workType, String(form.get("name") || "").trim());""",
	        """      const totalSla = totalSlaFor(workType, contractName, classificationEnFromValue(String(form.get("classification") || "")));""",
    )
    html = html.replace(
        """        name: form.get("name"),""",
        """        name: contractName,""",
    )
    html = html.replace(
        """        vendor: form.get("vendor"),""",
        """        vendor: String(form.get("vendor") || selectedAddTemplate?.vendor || "").trim(),""",
    )
    html = html.replace(
        """        station: `From ${owner} >> To ${to}`,
        due: finalDue,""",
        """        station: `From ${owner} >> To ${to}`,
        addCaseDate: lockedInDate,
        due: finalDue,""",
    )
    html = html.replace(
        """      attachEditableDropdown("addOwner", ownerDropdownOptions, syncAddCaseSystemFields);
      attachEditableDropdown("updateTo", () => directoryEmployeeOptions(), () => syncUpdateRecipientEmail(true));""",
        """      attachEditableDropdown("addOwner", ownerDropdownOptions, syncAddCaseSystemFields);
      attachEditableDropdown("addContractClassification", contractClassificationDropdownOptions, () => syncAddCaseLinkedFields("classification"), {
        ariaLabel: "Contract Classification / กลุ่มสัญญา",
        strict: true,
        defaultValue: "Day-to-day Work / งานดำเนินงานทั่วไป"
      });
      attachEditableDropdown("addContractType", contractTypeDropdownOptions, () => syncAddCaseLinkedFields("type"), {
        ariaLabel: "Type of Contract / ประเภทสัญญา"
      });
      attachEditableDropdown("addContractSubType", contractSubTypeDropdownOptions, () => syncAddCaseLinkedFields("subType"), {
        ariaLabel: "Sub Type of Contract / ประเภทย่อยของสัญญา"
      });
      attachEditableDropdown("updateTo", () => directoryEmployeeOptions(), () => syncUpdateRecipientEmail(true));""",
    )
    html = html.replace(
        """      const contractCatalog = getContractFormCatalog();
      const contractTypes = orderedUniqueList(contractCatalog.map(item => item.type).filter(Boolean));
      fillDatalist("contractNameOptions", contractCatalog.map(item => ({ value: item.name, label: `${item.type} · ${item.workType}` })));""",
        """      const contractCatalog = getContractFormCatalog();
      const contractTypes = orderedUniqueList(contractCatalog.map(item => item.type).filter(Boolean));
      fillDatalist("contractClassificationOptions", contractClassificationDropdownOptions().map(item => ({ value: item.value, label: item.secondary || "" })));
      fillDatalist("contractTypeOptions", contractTypeDropdownOptions().map(item => ({ value: item.value, label: item.secondary || "" })));
      fillDatalist("contractSubTypeOptions", contractSubTypeDropdownOptions().map(item => ({ value: item.value, label: item.secondary || "" })));
      fillDatalist("contractNameOptions", contractNameDropdownOptions().map(item => ({ value: item.value, label: item.secondary || "" })));""",
    )
    html = html.replace(
        """        <button class="nav-button" data-view="user" title="User Case Action">
          <span>✎</span><span class="nav-label">User Case Action</span><span class="nav-count">3</span>
        </button>""",
        """        <button class="nav-button confidential-nav" data-view="confidential" title="Confidential Contracts" data-confidential-nav>
          <span>◆</span><span class="nav-label">Confidential</span><span class="nav-count">0</span>
        </button>
        <button class="nav-button" data-view="user" title="User Case Action">
          <span>✎</span><span class="nav-label">User Case Action</span><span class="nav-count">3</span>
        </button>
        <button class="nav-button" data-view="master" title="Master Data">
          <span>▤</span><span class="nav-label">Master Data</span><span class="nav-count">5</span>
        </button>""",
    )
    html = html.replace(
        """        <section class="view" id="notifications" hidden>""",
        """        <section class="view" id="master">
          <section class="panel">
            <div class="panel-header">
              <div>
                <h2>Master Data</h2>
                <small>Edit dropdown data and save it back to Shared Drive</small>
              </div>
              <div class="master-data-actions">
                <button class="primary-button" type="button" id="saveMasterDataBtn">Save Master Data</button>
              </div>
            </div>
            <div class="master-data-grid">
              <section class="panel master-data-panel master-data-panel-full">
                <div class="panel-header">
                  <div>
                    <h2>Contract Records</h2>
                    <small>Edit or delete incorrect cases created from Add Case</small>
                  </div>
                </div>
                <div class="table-wrap">
                  <table class="master-table">
                    <thead><tr><th>Contract ID</th><th>Contract Name</th><th>Department / Restaurant</th><th>Contract Owner</th><th>Type of Contract</th><th>Stage</th><th>Status Update</th><th>Station Owner</th><th>Add Case Date</th><th>Due Date</th><th></th></tr></thead>
                    <tbody id="masterContractRows"></tbody>
                  </table>
                </div>
              </section>

              <section class="panel master-data-panel master-data-panel-full">
                <div class="panel-header">
                  <div>
                    <h2>Log Records</h2>
                    <small>Edit In / Out dates from Log View Detail. Other fields are read only.</small>
                  </div>
                </div>
                <div class="table-wrap">
                  <table class="master-table">
                    <thead><tr><th>Contract ID</th><th>Log No</th><th>Cycle</th><th>From</th><th>To</th><th>In</th><th>Out</th><th>SLA</th><th>Days on Hand</th><th>Alert</th><th>Action</th></tr></thead>
                    <tbody id="masterLogRows"></tbody>
                  </table>
                </div>
              </section>

              <section class="panel master-data-panel">
                <div class="panel-header">
                  <div>
                    <h2>Department Master</h2>
                    <small>Department / Restaurant + code for Contract ID</small>
                  </div>
                  <button class="secondary-button" type="button" data-add-master-row="departments">Add Row</button>
                </div>
                <div class="table-wrap">
                  <table class="master-table">
                    <thead><tr><th>Department / Restaurant</th><th>Code</th><th>Active</th><th></th></tr></thead>
                    <tbody id="masterDepartmentRows"></tbody>
                  </table>
                </div>
              </section>

              <section class="panel master-data-panel">
                <div class="panel-header">
                  <div>
                    <h2>People / Owner Master</h2>
                    <small>Used by Owner, To, CC and email lookup</small>
                  </div>
                  <button class="secondary-button" type="button" data-add-master-row="people">Add Row</button>
                </div>
                <div class="table-wrap">
                  <table class="master-table">
                    <thead><tr><th>Name</th><th>Department</th><th>Email</th><th>Company</th><th>Active</th><th></th></tr></thead>
                    <tbody id="masterPeopleRows"></tbody>
                  </table>
                </div>
              </section>

              <section class="panel master-data-panel">
                <div class="panel-header">
	                  <div>
                    <h2>Type of Contract Master</h2>
	                    <small>Used by Type/Sub Type dropdown, tooltip and Total SLA</small>
	                  </div>
                  <button class="secondary-button" type="button" data-add-master-row="contractTypes">Add Row</button>
                </div>
                <div class="table-wrap">
                  <table class="master-table">
                    <thead><tr><th>Contract Classification</th><th>Type of Contract</th><th>Sub Type of Contract</th><th>Total SLA - Day<br><small>SLA รวม - วันทำการ</small></th><th>Active</th><th></th></tr></thead>
                    <tbody id="masterContractTypeRows"></tbody>
                  </table>
	                </div>
	              </section>

	              <section class="panel master-data-panel">
	                <div class="panel-header">
	                  <div>
	                    <h2>Action SLA Master</h2>
	                    <small>Included in Total SLA · Action SLA รวมอยู่ใน Total SLA แล้ว ไม่บวกซ้ำ</small>
	                  </div>
	                  <button class="secondary-button" type="button" data-add-master-row="actionSla">Add Row</button>
	                </div>
	                <div class="table-wrap">
	                  <table class="master-table">
	                    <thead><tr><th>Action</th><th>Description<br><small>รายละเอียด</small></th><th>Fixed SLA<br><small>วันทำการ</small></th><th>Active</th><th></th></tr></thead>
	                    <tbody id="masterActionSlaRows"></tbody>
	                  </table>
	                </div>
	              </section>

	              <section class="panel master-data-panel master-data-panel-full">
	                <div class="panel-header">
	                  <div>
	                    <h2>Contract Name Template</h2>
	                    <small>Same flow as Add Case: Classification → Type → Sub Type → Contract Name → Vendor</small>
	                  </div>
	                  <button class="secondary-button" type="button" data-add-master-row="contractTemplates">Add Row</button>
	                </div>
	                <div class="table-wrap">
	                  <table class="master-table">
	                    <thead><tr><th>Contract Classification</th><th>Type of Contract</th><th>Sub Type of Contract</th><th>Contract Name</th><th>Vendor / Counter party</th><th>Department / Restaurant</th><th>Total SLA - Day<br><small>SLA รวม - วันทำการ</small></th><th>Access</th><th>Active</th><th></th></tr></thead>
	                    <tbody id="masterTemplateRows"></tbody>
	                  </table>
                </div>
              </section>
            </div>
          </section>
        </section>

        <section class="view" id="notifications" hidden>""",
    )
    html = html.replace(
        """      const contractType = String(typeInput.value || "").trim();
      if (!contractType) {""",
        """      const contractType = String(typeInput.value || "").trim();
      typeInput.title = contractTypeDescriptionFor(contractType);
      if (!contractType) {""",
    )
    html = html.replace(
        """      const linkedEmail = emailForPerson(toValue);""",
        """      const linkedEmail = updateEmailPattern(toValue) ? toValue.toLowerCase() : emailForPerson(toValue);""",
    )
    html = html.replace(
        """    function nextContractId() {
      const maxId = contracts.reduce((max, item) => {
        const number = Number(String(item.id).replace(/\\D/g, ""));
        return Number.isFinite(number) ? Math.max(max, number) : max;
      }, 100);
      return `C-${String(maxId + 1).padStart(3, "0")}`;
    }""",
        """    function currentAddCaseAccessLevel() {
      const form = document.querySelector("#addCaseForm");
      const explicitLevel = String(form?.elements.accessLevel?.value || "").trim();
      if (explicitLevel) return explicitLevel;
      const contractType = String(form?.elements.type?.value || "").trim();
      return accessLevelForContractType(contractType) || "Normal";
    }

    function nextContractId(accessLevel = "") {
      const level = String(accessLevel || currentAddCaseAccessLevel() || "Normal").trim();
      const classCode = level === "Confidential" ? "C" : "N";
      const yearCode = todayInputValue().slice(2, 4) || "26";
      const maxSequence = contracts.reduce((max, item) => {
        const parts = String(item.id || "").trim().split("-");
        if (parts.length !== 4 || parts[0] !== "CT" || parts[1] !== classCode || parts[2] !== yearCode) return max;
        const sequence = Number(parts[3]);
        return Number.isFinite(sequence) ? Math.max(max, sequence) : max;
      }, 0);
      return `CT-${classCode}-${yearCode}-${String(maxSequence + 1).padStart(3, "0")}`;
    }

    function refreshNewCaseIdPreview() {
      const badge = document.querySelector("#newCaseId");
      if (badge) badge.textContent = nextContractId();
    }""",
    )
    html = html.replace(
        """      document.querySelector("#newCaseId").textContent = nextContractId();""",
        """      refreshNewCaseIdPreview();""",
    )
    html = html.replace(
        """    function syncAddCaseSystemFields() {
      const form = document.querySelector("#addCaseForm");
      if (!form) return;""",
        """    function syncAddCaseSystemFields() {
      const form = document.querySelector("#addCaseForm");
      if (!form) return;
      refreshNewCaseIdPreview();""",
    )
    html = html.replace(
        """      if (!contractType || !workType) {""",
        """      const requiresSubType = Boolean(selectedContractTypeGroup() && contractSubTypeDropdownOptions().length);
      const hasSelectedSubType = Boolean(selectedContractSubType());
      if (!contractType || !workType || (requiresSubType && !hasSelectedSubType)) {""",
    )
    html = html.replace(
        """        if (preview) preview.textContent = "Waiting for Type of Contract · รอข้อมูลประเภทสัญญาเพื่อคำนวณ SLA และ Due Date";
        return;""",
        """        if (preview) preview.textContent = requiresSubType
          ? "Select Sub Type to calculate SLA and Due Date · เลือกประเภทย่อยเพื่อคำนวณ SLA และ Due Date"
          : "Select Type to calculate SLA and Due Date · เลือกประเภทเพื่อคำนวณ SLA และ Due Date";
        const slaHint = document.querySelector("#addSlaHint");
        if (slaHint) slaHint.textContent = requiresSubType
          ? "Select Sub Type · เลือกประเภทย่อย"
          : "Select Type · เลือกประเภท";
        renderAddCaseSmartSummary();
        return;""",
    )
    html = html.replace(
        """      if (preview) {
        const selectedDue = dueInput?.value || systemDue;
        const adjusted = Boolean(selectedDue && systemDue && selectedDue !== systemDue);
        preview.innerHTML = `Type of Contract: ${escapeHtml(contractType)} · SLA ${sla} Working Days · System Due ${escapeHtml(systemDue)}${adjusted ? ` · Initial Due Date adjusted to ${escapeHtml(selectedDue)}` : " · Initial Due Date can be edited before Add Case"}`;
      }
    }""",
        """      if (preview) {
        const selectedDue = dueInput?.value || systemDue;
        const adjusted = Boolean(selectedDue && systemDue && selectedDue !== systemDue);
        preview.innerHTML = `Type of Contract: ${escapeHtml(contractType)} · SLA ${sla} Working Days / ${sla} วันทำการ · System Due ${escapeHtml(systemDue)}${adjusted ? ` · Initial Due Date adjusted to ${escapeHtml(selectedDue)}` : " · Initial Due Date can be edited before Add Case"}`;
      }
      const slaHint = document.querySelector("#addSlaHint");
      if (slaHint) slaHint.textContent = `Fixed SLA: ${sla} Working Days / ${sla} วันทำการ`;
      renderAddCaseSmartSummary();
    }""",
    )
    html = replace_between(
        html,
        """    function syncAddCaseLinkedFields(source = "init") {""",
        """

    function syncAddCaseSystemFields() {""",
        """    function syncAddCaseGroupFilter() {
      syncAddCaseLinkedFields("classification");
    }

    function syncAddCaseLinkedFields(source = "init") {
      const form = document.querySelector("#addCaseForm");
      if (!form) return;
      const nameInput = document.querySelector("#addContractName");
      const typeInput = document.querySelector("#addContractType");
      const subTypeInput = document.querySelector("#addContractSubType");
      const workInput = document.querySelector("#addWorkType");
      const accessInput = document.querySelector("#addAccessLevel");
      const departmentInput = document.querySelector("#addDepartment");
      const vendorInput = document.querySelector("#addVendor");
      if (!nameInput || !typeInput || !subTypeInput || !workInput) return;

      setContractClassificationButtons(selectedContractClassification());

      if (source === "classification") {
        typeInput.value = "";
        subTypeInput.value = "";
        nameInput.value = "";
        if (vendorInput) vendorInput.value = "";
        workInput.value = "";
      }

      if (source === "type") {
        subTypeInput.value = "";
        nameInput.value = "";
        if (vendorInput) vendorInput.value = "";
        workInput.value = "";
      }

      if (source === "subType") {
        nameInput.value = "";
        if (vendorInput) vendorInput.value = "";
      }

      const selectedName = String(nameInput.value || "").trim();
      const template = contractTemplateFor(selectedName);
      if (source === "name" && template) {
        const templateInfo = contractTypeMasterV2Match(template.type);
        if (templateInfo) {
          const templateClassification = templateInfo["Contract Classification EN"] || selectedContractClassification();
          const classificationInput = document.querySelector("#addContractClassification");
          if (classificationInput) classificationInput.value = templateClassification;
          setContractClassificationButtons(templateClassification);
          typeInput.value = typeGroupEnglishForRow(templateInfo) || template.type || "";
          subTypeInput.value = subTypeEnglishForRow(templateInfo) || "";
        } else if (!typeInput.value) {
          typeInput.value = template.type || "";
          subTypeInput.value = "";
        }
        workInput.value = template.type || effectiveAddCaseContractType() || "Other";
        if (departmentInput && template.department) {
          departmentInput.value = template.department;
          syncDepartmentOwnerOptions(true);
          refreshEditableDropdown("addOwner");
        }
        if (vendorInput) vendorInput.value = template.vendor || "";
      }

      const typeGroup = selectedContractTypeGroup();
      const subOptions = contractSubTypeDropdownOptions();
      const selectedSubType = selectedContractSubType();
      const selectedClassification = selectedContractClassification();
      const selectedTypeRow = contractTypeMasterV2Match(typeGroup, selectedClassification);
      const selectedSubTypeRow = contractTypeMasterV2Match(selectedSubType, selectedClassification);
      const selectedTypeThai = typeGroupThaiForRow(selectedTypeRow);
      const selectedTypeDescriptionTh = contractTypeShortDescriptionTh(typeGroup);
      const selectedSubTypeThai = subTypeThaiForRow(selectedSubTypeRow);
      subTypeInput.disabled = !typeGroup || subOptions.length === 0;
      subTypeInput.required = Boolean(typeGroup && subOptions.length);
      subTypeInput.placeholder = !typeGroup
        ? "Waiting for Type of Contract"
        : subOptions.length
          ? "Select Sub Type"
          : "No Sub Type for this Type";
      if (!subOptions.length && source !== "name") subTypeInput.value = "";
      if (selectedSubType && !subOptions.some(item => contractTypeValuesMatch(item.value, selectedSubType))) {
        subTypeInput.value = "";
      }

      refreshEditableDropdown("addContractType");
      refreshEditableDropdown("addContractSubType");
      refreshEditableDropdown("addContractName");

      const effectiveType = effectiveAddCaseContractType();
      workInput.value = template?.type || effectiveType || "";

      if (!typeGroup) {
        if (accessInput) accessInput.value = classificationAccessLevelFor(selectedContractClassification());
        setLinkedFlowStatus("linkedTypeStatus", "Select Type · เลือกประเภท", "waiting");
        setLinkedFlowStatus("linkedSubTypeStatus", "Select Type first · เลือกประเภทก่อน", "waiting");
        renderAddCasePracticalExample();
        setLinkedFlowStatus("linkedNameStatus", "Enter Contract Name · ระบุชื่อสัญญา", "waiting");
        syncAddCaseSystemFields();
        return;
      }

      const accessLevel = accessLevelForAddCase(template, effectiveType || typeGroup);
      const selectedFixedSla = (!subOptions.length || selectedSubType)
        ? standardSlaFromContractTypeMasterV2(effectiveType || typeGroup, selectedContractClassification())
        : 0;
      if (accessInput) accessInput.value = accessLevel;
      setAccessLevelBadge(accessLevel);
      setLinkedFlowStatus("linkedAccessStatus", `${selectedContractClassification()} · ${classificationThaiFor(selectedContractClassification())}`, accessLevel === "Confidential" ? "waiting" : "linked");
      setLinkedFlowStatus(
        "linkedTypeStatus",
        subOptions.length
          ? `${selectedTypeDescriptionTh || selectedTypeThai || "ประเภทสัญญา"} · เลือก Sub Type`
          : `${selectedTypeThai || "ประเภทสัญญา"} · Fixed SLA ${selectedFixedSla} Working Days / ${selectedFixedSla} วันทำการ`,
        "linked"
      );
      setLinkedFlowStatus(
        "linkedSubTypeStatus",
        subOptions.length
          ? (selectedSubType && selectedFixedSla
              ? `${selectedSubTypeThai || "ประเภทย่อยของสัญญา"} · Fixed SLA ${selectedFixedSla} Working Days / ${selectedFixedSla} วันทำการ`
              : "Select Sub Type · เลือกประเภทย่อย")
          : "No Sub Type · ไม่มีประเภทย่อย",
        selectedSubType || !subOptions.length ? "linked" : "waiting"
      );
      renderAddCasePracticalExample(
        selectedSubType || typeGroup,
        selectedContractClassification(),
        Boolean(subOptions.length && !selectedSubType)
      );
      const nameCount = contractNameDropdownOptions().length;
      setLinkedFlowStatus(
        "linkedNameStatus",
        selectedName
          ? (template ? "Contract Name linked · เชื่อมข้อมูลชื่อสัญญาแล้ว" : "New Contract Name · ชื่อสัญญาใหม่")
          : `${nameCount} Contract Name option(s) · เลือกชื่อสัญญา`,
        selectedName ? "linked" : "waiting"
      );

      syncAddCaseSystemFields();
    }""",
    )
    html = html.replace(
        """      const workType = String(form.elements.workType?.value || "").trim();
      const contractType = String(form.elements.type?.value || "").trim();""",
        """      const workType = String(form.elements.workType?.value || effectiveAddCaseContractType() || "").trim();
      const contractType = String(effectiveAddCaseContractType() || form.elements.type?.value || "").trim();""",
    )
    html = html.replace(
        """      const workType = String(form.get("workType") || "Other").trim();
      const contractType = String(form.get("type") || "").trim();""",
        """      const workType = String(form.get("workType") || form.get("subType") || form.get("type") || "Other").trim();
      const contractType = String(form.get("subType") || form.get("type") || "").trim();""",
    )
    html = html.replace(
        """        type: form.get("type"),""",
        """        type: contractType,""",
    )
    html = html.replace(
        """      const accessLevel = String(form.get("accessLevel") || accessLevelForContractType(contractType) || "Normal").trim();
      const owner = String(form.get("owner") || "").trim();""",
        """      const accessLevel = String(form.get("accessLevel") || accessLevelForContractType(contractType) || "Normal").trim();
      const id = nextContractId(accessLevel);
      const owner = String(form.get("owner") || "").trim();""",
    )
    html = html.replace(
        """        email: toEmail || emailForPerson(to),""",
        """        email: toEmail || (updateEmailPattern(to) ? to.toLowerCase() : emailForPerson(to)),""",
    )
    html = html.replace(
        """      const id = nextContractId();
      const workType = String(form.get("workType") || "Other").trim();
      const contractType = String(form.get("type") || "").trim();""",
        """      const workType = String(form.get("workType") || "Other").trim();
      const contractType = String(form.get("type") || "").trim();""",
    )
    html = html.replace(
        """      const workType = String(form.get("workType") || "Other").trim();
      const contractType = String(form.get("type") || "").trim();
      const accessLevel = String(form.get("accessLevel") || accessLevelForContractType(contractType) || "Normal").trim();
      const owner = String(form.get("owner") || "").trim();""",
        """      const workType = String(form.get("workType") || "Other").trim();
      const contractType = String(form.get("type") || "").trim();
      const accessLevel = String(form.get("accessLevel") || accessLevelForContractType(contractType) || "Normal").trim();
      const id = nextContractId(accessLevel);
      const owner = String(form.get("owner") || "").trim();""",
    )
    html = html.replace(
        """      if (actionInfo) actionInfo.hidden = false;
      if (descriptionEn) descriptionEn.textContent = info.shortEn;
      if (descriptionTh) descriptionTh.textContent = info.shortTh;
      if (slaEn) slaEn.textContent = `SLA: ${info.sla} Working Day${info.sla === 1 ? "" : "s"}`;
      if (slaTh) slaTh.textContent = `/ ${info.sla} วันทำการ`;""",
        """      if (actionInfo) actionInfo.hidden = false;
      if (descriptionEn) descriptionEn.textContent = `Status: ${info.nameEn}`;
      if (descriptionTh) descriptionTh.textContent = `สถานะ: ${info.nameTh}`;
      if (slaEn) slaEn.textContent = `SLA: ${info.sla} Working Day${info.sla === 1 ? "" : "s"}`;
      if (slaTh) slaTh.textContent = `/ ${info.sla} วันทำการ`;""",
    )
    html = html.replace(
        """                    <div class="action-info-inline" id="updateActionDescription" aria-live="polite" hidden>
                      <span id="updateActionDescriptionEn" class="action-info-en"></span>
                      <span class="action-info-separator" aria-hidden="true">·</span>
                      <span id="updateActionDescriptionTh" class="action-info-th"></span>
                      <span class="action-info-separator" aria-hidden="true">·</span>
                      <strong id="updateSlaEn" class="action-sla-en"></strong>
                      <span id="updateSlaTh" class="action-sla-th"></span>
                    </div>""",
        """                    <div class="action-info-inline" id="updateActionDescription" aria-live="polite" hidden>
                      <span class="action-status-pill">
                        <span id="updateActionDescriptionEn" class="action-info-en"></span>
                        <span id="updateActionDescriptionTh" class="action-info-th"></span>
                      </span>
                      <span class="action-sla-pill">
                        <strong id="updateSlaEn" class="action-sla-en"></strong>
                        <span id="updateSlaTh" class="action-sla-th"></span>
                      </span>
                    </div>""",
    )
    html = html.replace(
        """    .action-select-field .action-sla-th {
      color: var(--muted);
      font-size: 10px;
      font-weight: 500;
      white-space: nowrap;
    }""",
        """    .action-select-field .action-sla-th {
      color: var(--muted);
      font-size: 10px;
      font-weight: 500;
      white-space: nowrap;
    }

    .action-select-field .action-sla-pill,
    .action-select-field .action-status-pill {
      display: inline-flex;
      align-items: center;
      flex-wrap: wrap;
      gap: 4px;
      min-height: 26px;
      padding: 4px 8px;
      border: 1px solid var(--line);
      border-radius: 7px;
      background: #ffffff;
    }

    .action-select-field .action-sla-pill {
      border-color: #c6d6ee;
      background: #f5f9ff;
    }

    .action-select-field .action-status-pill {
      border-color: #d8e6b8;
      background: #f8faf3;
    }""",
    )
    html = html.replace(
        """                    <select class="select" name="stage" required>
                      <option>Signed / Completed</option>
                      <option>Cancelled</option>
                    </select>
                  </label>
                  <input type="hidden" name="owner" id="closeOwner">
                  <input type="hidden" name="closeDate" id="closeDate">
                  <div class="runtime-field close-system-summary">
                    <strong>System fields</strong>
                    <span>Final Contract Owner และ Close Date จะบันทึกจากข้อมูลปัจจุบันโดยอัตโนมัติ</span>
                  </div>""",
        """                    <select class="select" name="stage" required>
                      <option>Signed / Completed</option>
                      <option>Cancelled</option>
                    </select>
                    <small class="form-hint">Status: Signed / Completed = case finished · Cancelled = case closed without completion</small>
                  </label>
                  <input type="hidden" name="owner" id="closeOwner">
                  <input type="hidden" name="closeDate" id="closeDate">
                  <div class="runtime-field close-system-summary">
                    <strong>System fields</strong>
                    <span>Auto: Final Contract Owner + Close Date</span>
                    <small>ระบบบันทึกเจ้าของสัญญาสุดท้ายและวันที่ปิดจากข้อมูลปัจจุบัน</small>
                  </div>""",
    )
    html = html.replace(
        """        messageEn = "The requested date has not been changed. Select a date different from the current Due Date.";
        messageTh = "ยังไม่ได้ปรับวันที่ กรุณาเลือกวันที่ใหม่ที่แตกต่างจากวันที่ใช้งานปัจจุบัน";""",
        """        messageEn = "Select a different date.";
        messageTh = "เลือกวันที่ใหม่ให้ต่างจากวันปัจจุบัน";""",
    )
    html = html.replace(
        """              <div class="filter-row">
                <select class="select" id="stageFilter">
                  <option value="all">All stages</option>
                  <option value="Draft Created">Draft Created</option>
                  <option value="Submit to Review">Submit to Review</option>
                  <option value="Return">Return</option>
                  <option value="Resubmit">Resubmit</option>
	                  <option value="Forward">Forward</option>
	                  <option value="Signed / Completed">Signed / Completed</option>
                  <option value="Cancelled">Cancelled</option>
                </select>
                <select class="select" id="departmentFilter">
                  <option value="all">All Department / Restaurant</option>
                  <option value="OP / 3 fl Capri">OP / 3 fl Capri</option>
                  <option value="Baan Turtle BKK">Baan Turtle BKK</option>
                  <option value="Turtle 10 / Lease Asset">Turtle 10 / Lease Asset</option>
                  <option value="IT">IT</option>
                  <option value="สาธุประดิษฐ์ - BK Salon">สาธุประดิษฐ์ - BK Salon</option>
                  <option value="Baan Turtle">Baan Turtle</option>
                  <option value="FSQ / Utility">FSQ / Utility</option>
                  <option value="Legal">Legal</option>
                  <option value="Marketing">Marketing</option>
                </select>
              </div>
""",
	        "",
	    )
    html = re.sub(
        r'\n\s*<div class="filter-row">\s*<select class="select" id="stageFilter">.*?</div>\n',
        "\n",
        html,
        count=1,
        flags=re.S,
    )
    html = html.replace(
        """                    <th>Contract ID</th>
                    <th>Contract Name</th>
                    <th>Department / Restaurant</th>
                    <th>Contract Owner</th>
                    <th>Type of Contract</th>
                    <th>Vendor / Counter party</th>
                    <th>Log View</th>
                    <th>Stage</th>
                    <th>Total No. of Cycle</th>
                    <th>Total No. of Return</th>
                    <th>Status Update</th>
                    <th>Station Owner</th>
                    <th>Due date</th>""",
        """                    <th class="filter-th">
                      <details data-filter-menu>
                        <summary>Contract ID</summary>
                        <div class="th-filter-popover">
                          <select class="select contract-master-filter" id="contractIdFilter" aria-label="Filter Contract ID"><option value="all">All Contract ID</option></select>
                        </div>
                      </details>
                    </th>
                    <th>Contract Name</th>
                    <th class="filter-th">
                      <details data-filter-menu>
                        <summary>Department / Restaurant</summary>
                        <div class="th-filter-popover">
                          <select class="select contract-master-filter" id="departmentFilter" aria-label="Filter Department / Restaurant"><option value="all">All Department / Restaurant</option></select>
                        </div>
                      </details>
                    </th>
                    <th class="filter-th">
                      <details data-filter-menu>
                        <summary>Contract Owner</summary>
                        <div class="th-filter-popover">
                          <select class="select contract-master-filter" id="ownerFilter" aria-label="Filter Contract Owner"><option value="all">All Contract Owner</option></select>
                        </div>
                      </details>
                    </th>
                    <th class="filter-th">
                      <details data-filter-menu>
                        <summary>Type of Contract</summary>
                        <div class="th-filter-popover">
                          <select class="select contract-master-filter" id="typeFilter" aria-label="Filter Type of Contract"><option value="all">All Type of Contract</option></select>
                        </div>
                      </details>
                    </th>
                    <th>Vendor / Counter party</th>
                    <th>Log View</th>
                    <th class="filter-th">
                      <details data-filter-menu>
                        <summary>Stage</summary>
                        <div class="th-filter-popover">
                          <select class="select contract-master-filter" id="stageFilter" aria-label="Filter Stage"><option value="all">All stages</option></select>
                        </div>
                      </details>
                    </th>
                    <th>Total No. of Cycle</th>
                    <th>Total No. of Return</th>
                    <th class="filter-th">
                      <details data-filter-menu>
                        <summary>Status Update</summary>
                        <div class="th-filter-popover">
                          <select class="select contract-master-filter" id="statusFilter" aria-label="Filter Status Update"><option value="all">All Status Update</option></select>
                        </div>
                      </details>
                    </th>
                    <th class="filter-th">
                      <details data-filter-menu>
                        <summary>Station Owner</summary>
                        <div class="th-filter-popover">
                          <select class="select contract-master-filter" id="stationOwnerFilter" aria-label="Filter Station Owner"><option value="all">All Station Owner</option></select>
                        </div>
                      </details>
                    </th>
                    <th class="filter-th">
                      <details data-filter-menu>
                        <summary>Due date</summary>
                        <div class="th-filter-popover due-date-range-popover">
                          <label class="due-date-range-row">
                            <span>From / จากวันที่</span>
                            <input class="input contract-master-filter" type="date" id="dueDateFromFilter" aria-label="Filter Due date from">
                          </label>
                          <label class="due-date-range-row">
                            <span>To / ถึงวันที่</span>
                            <input class="input contract-master-filter" type="date" id="dueDateToFilter" aria-label="Filter Due date to">
                          </label>
                          <button class="secondary-button" type="button" id="clearDueDateFilter">Clear</button>
                        </div>
                      </details>
                    </th>""",
    )
    html = html.replace(
        """            <div class="panel-header">
              <div>
                <h2>Contract Master</h2>
                <small>ติดตาม Contract Owner, cycle, return และสถานะล่าสุด</small>
              </div>
            </div>
            <div class="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th class="filter-th">""",
        """            <div class="panel-header">
              <div>
                <h2>Contract Master</h2>
                <small>ติดตาม Contract Owner, cycle, return และสถานะล่าสุด</small>
              </div>
            </div>
            <div class="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th class="filter-th">""",
    )
    html = html.replace(
        """          { labelEn: "CC", labelTh: "ผู้รับสำเนา", valueEn: updateCcRecipients.length ? updateCcRecipients.map(item => item.email).join(", ") : "-", valueTh: updateCcRecipients.length ? updateCcRecipients.map(item => item.email).join(", ") : "-", full: updateCcRecipients.length > 1 },
""",
        "",
    )
    old_log_headers = """                      <th>Alert</th>
                      <th>Action</th>
                      <th>Action Reason</th>
                      <th>Approval</th>
                      <th>Delay Reason</th>
                      <th>Corrective Action</th>
                      <th>Attachments<br><small>ไฟล์แนบ</small></th>
                      <th>CC Recipients<br><small>ผู้รับสำเนา</small></th>
                      <th>Updated By</th>
                      <th>Updated Date and Time</th>"""
    new_log_headers = """                      <th>Alert</th>
                      <th>Delay Reason</th>
                      <th>Action</th>"""
    html = html.replace(old_log_headers, new_log_headers)
    old_detail_headers = """                  <th>Alert</th>
                  <th>Action</th>
                  <th>Action Reason</th>
                  <th>Approval</th>
                  <th>Delay Reason</th>
                  <th>Corrective Action</th>
                  <th>Attachments<br><small>ไฟล์แนบ</small></th>
                  <th>CC Recipients<br><small>ผู้รับสำเนา</small></th>
                  <th>Updated By</th>
                  <th>Updated Date and Time</th>"""
    new_detail_headers = """                  <th>Alert</th>
                  <th>Delay Reason</th>
                  <th>Action</th>"""
    html = html.replace(old_detail_headers, new_detail_headers)
    old_log_cells = """            <td>${alertBadge(row[10])}</td>
            <td><strong>${escapeHtml(row[24] || actionDefinition(row[12]).nameTh)}</strong><small style="display:block;color:var(--muted)">${escapeHtml(row[25] || row[12])}</small></td>
            <td>${escapeHtml(logActionReasonText(row))}</td>
            <td>${escapeHtml(logApprovalText(row))}</td>
            <td>${escapeHtml(row[11] || "-")}</td>
            <td>${escapeHtml(row[20] || "-")}</td>
            <td>${renderLogAttachmentCount(row)}</td>
            <td>${renderLogCcCount(row)}</td>
            <td>${escapeHtml(row[21] || "-")}</td>
            <td>${escapeHtml(formatDateTime(row[22]))}</td>
          </tr>`).join("") : `<tr><td colspan="16">No log for selected Contract ID</td></tr>`;"""
    new_log_cells = """            <td>${alertBadge(row[10])}</td>
            <td>${escapeHtml(row[11] || "-")}</td>
            <td>${escapeHtml(row[12] || "-")}</td>
          </tr>`).join("") : `<tr><td colspan="11">No log for selected Contract ID</td></tr>`;"""
    html = html.replace(old_log_cells, new_log_cells)
    old_detail_cells = """                    <td>${alertBadge(row[10])}</td>
                    <td><strong>${escapeHtml(row[24] || actionDefinition(row[12]).nameTh)}</strong><small style="display:block;color:var(--muted)">${escapeHtml(row[25] || row[12])}</small></td>
                    <td>${escapeHtml(logActionReasonText(row))}</td>
                    <td>${escapeHtml(logApprovalText(row))}</td>
                    <td>${escapeHtml(row[11] || "-")}</td>
                    <td>${escapeHtml(row[20] || "-")}</td>
                    <td>${renderLogAttachmentCount(row)}</td>
                    <td>${renderLogCcCount(row)}</td>
                    <td>${escapeHtml(row[21] || "-")}</td>
                    <td>${escapeHtml(formatDateTime(row[22]))}</td>
                  </tr>`).join("") : `<tr><td colspan="16">No Log View Detail found.</td></tr>`}"""
    new_detail_cells = """                    <td>${alertBadge(row[10])}</td>
                    <td>${escapeHtml(row[11] || "-")}</td>
                    <td>${escapeHtml(row[12] || "-")}</td>
                  </tr>`).join("") : `<tr><td colspan="11">No Log View Detail found.</td></tr>`}"""
    html = html.replace(old_detail_cells, new_detail_cells)
    html = html.replace(
        """      const typeInput = document.querySelector("#addContractType");
      const workInput = document.querySelector("#addWorkType");
      if (!nameInput || !typeInput || !workInput) return;""",
        """      const typeInput = document.querySelector("#addContractType");
      const workInput = document.querySelector("#addWorkType");
      const accessInput = document.querySelector("#addAccessLevel");
      const groupInput = document.querySelector("#addContractGroup");
      const departmentInput = document.querySelector("#addDepartment");
      const vendorInput = document.querySelector("#addVendor");
      if (!nameInput || !typeInput || !workInput) return;""",
    )
    html = html.replace(
        """        typeInput.value = "";
        workInput.value = "";
        typeInput.disabled = true;""",
        """        typeInput.value = "";
        workInput.value = "";
        if (accessInput) accessInput.value = "";
        if (groupInput && source !== "group") groupInput.value = "";
        if (vendorInput) vendorInput.value = "";
        setAccessLevelBadge("");
        typeInput.disabled = true;""",
    )
    html = html.replace(
        """          typeInput.value = template.type || "";
          workInput.value = template.workType || "Other";
          typeInput.dataset.autoLinked = "true";""",
        """          typeInput.value = template.type || "";
          workInput.value = template.workType || "Other";
          if (groupInput && !groupInput.value && (template.group || template.category)) {
            groupInput.value = template.group || template.category;
          }
          if (departmentInput && template.department) {
            departmentInput.value = template.department;
            syncDepartmentOwnerOptions(true);
            refreshEditableDropdown("addOwner");
          }
          if (vendorInput) vendorInput.value = template.vendor || "";
          typeInput.dataset.autoLinked = "true";""",
    )
    html = html.replace(
        """          typeInput.value = "";
          workInput.value = "";
          delete typeInput.dataset.autoLinked;""",
        """          typeInput.value = "";
          workInput.value = "";
          if (vendorInput) vendorInput.value = "";
          delete typeInput.dataset.autoLinked;""",
    )
    html = html.replace(
        """        setLinkedFlowStatus("linkedNameStatus", "Start here · เริ่มกรอกชื่อสัญญา", "waiting");
        setLinkedFlowStatus("linkedTypeStatus", "Waiting for Contract Name · รอข้อมูล Contract Name", "waiting");
        syncAddCaseSystemFields();""",
        """        setLinkedFlowStatus("linkedNameStatus", "Select Contract Name · เลือกชื่อสัญญา", "waiting");
        setLinkedFlowStatus("linkedTypeStatus", "Waiting for Contract Name · รอชื่อสัญญา", "waiting");
        setLinkedFlowStatus("linkedAccessStatus", "Waiting for Type · รอประเภท", "waiting");
        syncAddCaseSystemFields();""",
    )
    html = html.replace(
        """    function syncAddCaseLinkedFields__legacy_group_insertion_disabled(source = "init") {""",
        """    function syncAddCaseGroupFilter() {
      const groupInput = document.querySelector("#addContractGroup");
      const nameInput = document.querySelector("#addContractName");
      const typeInput = document.querySelector("#addContractType");
      const vendorInput = document.querySelector("#addVendor");
      const workInput = document.querySelector("#addWorkType");
      const group = String(groupInput?.value || "").trim();
      const groupStatus = document.querySelector("#linkedGroupStatus");
      if (groupStatus) {
        const count = contractNameDropdownOptions().length;
        groupStatus.textContent = group ? `${count} Contract Name(s) in selected group · ${count} รายการในกลุ่มนี้` : "All groups · แสดงทุกกลุ่ม";
        groupStatus.classList.toggle("linked", Boolean(group));
        groupStatus.classList.toggle("waiting", !group);
      }
      const selectedTemplate = contractTemplateFor(String(nameInput?.value || "").trim());
      if (group && selectedTemplate && (selectedTemplate.group || selectedTemplate.category) !== group) {
        if (nameInput) nameInput.value = "";
        if (typeInput) {
          typeInput.value = "";
          typeInput.disabled = true;
        }
        if (vendorInput) vendorInput.value = "";
        if (workInput) workInput.value = "";
        setAccessLevelBadge("");
      }
      refreshEditableDropdown("addContractName");
      syncAddCaseLinkedFields("group");
    }

    function syncAddCaseLinkedFields(source = "init") {""",
    )
    html = html.replace(
        """      if (!contractType) {
        workInput.value = "";
        setLinkedFlowStatus("linkedTypeStatus", "Enter Type of Contract · กรอกประเภทสัญญา", "waiting");
        syncAddCaseSystemFields();
        return;
      }""",
        """      if (!contractType) {
        workInput.value = "";
        if (accessInput) accessInput.value = "";
        setAccessLevelBadge("");
        setLinkedFlowStatus("linkedTypeStatus", "Enter Type of Contract · กรอกประเภทสัญญา", "waiting");
        setLinkedFlowStatus("linkedAccessStatus", "Waiting for Type · รอประเภท", "waiting");
        syncAddCaseSystemFields();
        return;
      }""",
    )
    html = html.replace(
        """      setLinkedFlowStatus("linkedTypeStatus", template && template.type === contractType
        ? "Linked from Contract Name · เชื่อมจากชื่อสัญญา"
        : "Type of Contract received · รับข้อมูลประเภทสัญญาแล้ว", template && template.type === contractType ? "linked" : "");

      if (source === "type" || !workInput.value) {""",
        """      setLinkedFlowStatus("linkedTypeStatus", template && template.type === contractType
        ? "Linked from Contract Name · เชื่อมจากชื่อสัญญา"
        : "Type of Contract received · รับข้อมูลประเภทสัญญาแล้ว", template && template.type === contractType ? "linked" : "");

      const accessLevel = accessLevelForAddCase(template, contractType);
      if (accessInput) accessInput.value = accessLevel;
      setAccessLevelBadge(accessLevel);
      const category = contractTypeCategoryFor(contractType);
      const accessSource = template?.accessLevel
        ? `From Contract Register (${template.contractId || "existing contract"})`
        : `From Contract Type Master category: ${category || "Day-to-day"}`;
      setLinkedFlowStatus(
        "linkedAccessStatus",
        `${accessLevel} · ${accessSource} · ระบบจำแนกจากข้อมูล Excel`,
        accessLevel === "Confidential" ? "waiting" : "linked"
      );

      if (source === "type" || !workInput.value) {""",
    )
    html = html.replace(
        """      const workType = String(form.get("workType") || "Other").trim();
      const owner = String(form.get("owner") || "").trim();""",
        """      const workType = String(form.get("workType") || "Other").trim();
      const contractType = String(form.get("type") || "").trim();
      const accessLevel = String(form.get("accessLevel") || accessLevelForContractType(contractType) || "Normal").trim();
      const id = nextContractId(accessLevel);
      const owner = String(form.get("owner") || "").trim();""",
    )
    html = html.replace(
        """        owner,
        type: form.get("type"),
        vendor: form.get("vendor"),""",
        """        owner,
        type: contractType,
        vendor: form.get("vendor"),""",
    )
    html = html.replace(
        """        alert,
        remark: combinedRemark
      };""",
        """        alert,
        remark: combinedRemark,
        accessLevel,
        visibility: accessLevel === "Confidential" ? "Restricted access / จำกัดสิทธิ์" : "Standard access / สิทธิ์ทั่วไป",
        category: contractTypeCategoryFor(contractType)
      };""",
    )
    html = html.replace(
        """      const id = nextContractId();
      const workType = String(form.get("workType") || "Other").trim();""",
        """      const workType = String(form.get("workType") || "Other").trim();""",
    )
    html = html.replace(
        """    function populateDepartmentFilter() {
      const select = document.querySelector("#departmentFilter");
      if (!select) return;
      const current = select.value;
      const departments = [...new Set(contracts.map(item => item.department))];
      select.innerHTML = [
        `<option value="all">All Department / Restaurant</option>`,
        ...departments.map(department => `<option value="${escapeHtml(department)}">${escapeHtml(department)}</option>`)
      ].join("");
      select.value = departments.includes(current) ? current : "all";
    }""",
        """    function uniqueContractMasterValues(getter) {
      return [...new Set(contracts.map(getter).map(value => String(value || "").trim()).filter(Boolean))]
        .sort((a, b) => a.localeCompare(b));
    }

    function setContractMasterFilterOptions(id, label, values) {
      const select = document.querySelector(`#${id}`);
      if (!select) return;
      const current = select.value;
      select.innerHTML = [
        `<option value="all">All ${escapeHtml(label)}</option>`,
        ...values.map(value => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`)
      ].join("");
      select.value = values.includes(current) ? current : "all";
    }

    function populateContractMasterFilters() {
      setContractMasterFilterOptions("contractIdFilter", "Contract ID", uniqueContractMasterValues(item => item.id));
      setContractMasterFilterOptions("departmentFilter", "Department / Restaurant", uniqueContractMasterValues(item => item.department));
      setContractMasterFilterOptions("ownerFilter", "Contract Owner", uniqueContractMasterValues(item => item.owner));
      setContractMasterFilterOptions("typeFilter", "Type of Contract", uniqueContractMasterValues(item => contractPrimaryTypeDisplay(item.type)));
      setContractMasterFilterOptions("stageFilter", "Stage", uniqueContractMasterValues(item => item.stage));
      setContractMasterFilterOptions("statusFilter", "Status Update", uniqueContractMasterValues(item => item.status));
      setContractMasterFilterOptions("stationOwnerFilter", "Station Owner", uniqueContractMasterValues(item => {
        const latestLog = latestLogFor(item.id);
        return latestLog ? latestLog[5] : stationParts(item.station).to;
      }));
    }

    function contractMasterFilterValue(id) {
      const control = document.querySelector(`#${id}`);
      return control ? control.value : "all";
    }

    function matchesContractMasterFilter(value, selected) {
      return selected === "all" || String(value || "").trim() === selected;
    }""",
    )
    html = html.replace(
        """            <td>${item.type}</td>""",
        """            <td>${contractPrimaryTypeDisplay(item.type)}</td>""",
    )
    html = html.replace(
        """      const stage = document.querySelector("#stageFilter").value;
      const department = document.querySelector("#departmentFilter").value;
      const visibleContracts = contracts.filter(item => {
        return (stage === "all" || item.stage === stage) && (department === "all" || item.department === department);
      });
      const rows = visibleContracts.map(item => {
        const latestLog = latestLogFor(item.id);
        const latestStationOwner = latestLog ? latestLog[5] : stationParts(item.station).to;""",
        """      const filters = {
        contractId: contractMasterFilterValue("contractIdFilter"),
        department: contractMasterFilterValue("departmentFilter"),
        owner: contractMasterFilterValue("ownerFilter"),
        type: contractMasterFilterValue("typeFilter"),
        stage: contractMasterFilterValue("stageFilter"),
        status: contractMasterFilterValue("statusFilter"),
        stationOwner: contractMasterFilterValue("stationOwnerFilter"),
        dueDateFrom: document.querySelector("#dueDateFromFilter")?.value || "",
        dueDateTo: document.querySelector("#dueDateToFilter")?.value || ""
      };
      const visibleContracts = contracts.filter(item => {
        const latestLog = latestLogFor(item.id);
        const latestStationOwner = latestLog ? latestLog[5] : stationParts(item.station).to;
        return matchesContractMasterFilter(item.id, filters.contractId)
          && matchesContractMasterFilter(item.department, filters.department)
          && matchesContractMasterFilter(item.owner, filters.owner)
          && matchesContractMasterFilter(contractPrimaryTypeDisplay(item.type), filters.type)
          && matchesContractMasterFilter(item.stage, filters.stage)
          && matchesContractMasterFilter(item.status, filters.status)
          && matchesContractMasterFilter(latestStationOwner, filters.stationOwner)
          && (!filters.dueDateFrom || dateToInputValue(item.due) >= filters.dueDateFrom)
          && (!filters.dueDateTo || dateToInputValue(item.due) <= filters.dueDateTo);
      });
      const rows = visibleContracts.map(item => {
        const latestLog = latestLogFor(item.id);
        const latestStationOwner = latestLog ? latestLog[5] : stationParts(item.station).to;""",
    )
    html = html.replace(
        """      populateDepartmentFilter();
      renderContractsTable();""",
        """      populateContractMasterFilters();
      renderContractsTable();""",
    )
    html = html.replace(
        """    ["stageFilter", "departmentFilter", "queueFilter"].forEach(id => {
      const control = document.querySelector(`#${id}`);
      if (!control) return;
      control.addEventListener("input", renderAll);
      control.addEventListener("change", renderAll);
    });""",
        """    [
      "contractIdFilter",
      "departmentFilter",
      "ownerFilter",
      "typeFilter",
      "stageFilter",
      "statusFilter",
      "stationOwnerFilter",
      "dueDateFromFilter",
      "dueDateToFilter",
      "queueFilter"
    ].forEach(id => {
      const control = document.querySelector(`#${id}`);
      if (!control) return;
      control.addEventListener("input", renderAll);
      control.addEventListener("change", renderAll);
    });

    document.querySelector("#clearDueDateFilter")?.addEventListener("click", () => {
      const fromInput = document.querySelector("#dueDateFromFilter");
      const toInput = document.querySelector("#dueDateToFilter");
      if (fromInput) fromInput.value = "";
      if (toInput) toInput.value = "";
      renderAll();
    });

    document.addEventListener("click", event => {
      if (event.target.closest("[data-filter-menu]")) return;
      document.querySelectorAll("[data-filter-menu][open]").forEach(menu => menu.removeAttribute("open"));
    });

    function positionContractFilterMenu(menu) {
      const summary = menu?.querySelector("summary");
      const popover = menu?.querySelector(".th-filter-popover");
      if (!summary || !popover || !menu.open) return;
      const summaryRect = summary.getBoundingClientRect();
      const popoverWidth = Math.max(185, Math.min(popover.offsetWidth || 185, window.innerWidth - 24));
      const left = Math.min(Math.max(12, summaryRect.left), Math.max(12, window.innerWidth - popoverWidth - 12));
      const topBelow = summaryRect.bottom + 6;
      const estimatedHeight = Math.max(popover.offsetHeight || 48, 48);
      const top = topBelow + estimatedHeight + 12 > window.innerHeight
        ? Math.max(12, summaryRect.top - estimatedHeight - 6)
        : topBelow;
      popover.style.setProperty("--filter-popover-left", `${left}px`);
      popover.style.setProperty("--filter-popover-top", `${top}px`);
    }

    function positionOpenContractFilterMenus() {
      document.querySelectorAll("[data-filter-menu][open]").forEach(positionContractFilterMenu);
    }

    document.querySelectorAll("[data-filter-menu]").forEach(menu => {
      menu.addEventListener("toggle", () => {
        if (!menu.open) return;
        document.querySelectorAll("[data-filter-menu][open]").forEach(otherMenu => {
          if (otherMenu !== menu) otherMenu.removeAttribute("open");
        });
        requestAnimationFrame(() => positionContractFilterMenu(menu));
      });
      menu.querySelector("summary")?.addEventListener("click", () => {
        requestAnimationFrame(() => positionContractFilterMenu(menu));
      });
    });

    window.addEventListener("resize", positionOpenContractFilterMenus);
    document.querySelector("#contracts .table-wrap")?.addEventListener("scroll", positionOpenContractFilterMenus);

    document.querySelectorAll("[data-filter-menu] .select, [data-filter-menu] .input").forEach(control => {
      control.addEventListener("click", event => event.stopPropagation());
    });""",
    )
    html = html.replace(
        """          <button class="ghost-button" type="button" id="copyStatusEmailBtn">Copy Email</button>
          <button class="ghost-button" type="button" id="skipStatusEmailBtn">Close</button>
          <button class="primary-button" type="button" id="openMailClientBtn"><span>✉</span> Open Email &amp; Send</button>""",
        """          <button class="ghost-button" type="button" id="copyStatusEmailBtn">Copy Email</button>
          <button class="ghost-button" type="button" id="skipStatusEmailBtn">Close</button>
          <button class="primary-button" type="button" id="openMailClientBtn"><span>✉</span> Send Email</button>""",
    )
    html = html.replace(
        """          <label class="form-field full" id="statusEmailCcField" hidden>
            <span class="bilingual-stack"><span class="en">CC</span><span class="th">ผู้รับสำเนา</span></span>
            <input class="input" type="text" id="statusEmailCc" placeholder="cc1@turtle23.com, cc2@turtle23.com">
          </label>""",
        """          <label class="form-field full status-email-cc-field" id="statusEmailCcField">
            <span class="bilingual-stack"><span class="en">CC</span><span class="th">สำเนาอีเมล</span></span>
            <div class="cc-tag-input" id="statusEmailCcTagInput">
              <div class="cc-chip-list" id="statusEmailCcChips"></div>
              <input class="cc-entry-input" id="statusEmailCc" list="statusEmailCcOptions" autocomplete="off"
                placeholder="Search name or enter email address&#10;ค้นหาชื่อหรือกรอกอีเมล"
                aria-label="CC / สำเนาอีเมล">
            </div>
            <datalist id="statusEmailCcOptions"></datalist>
            <div class="compact-help bilingual-stack">
              <span class="en">Press Enter or comma to add more than one recipient.</span>
              <span class="th">กด Enter หรือเครื่องหมายจุลภาคเพื่อเพิ่มผู้รับหลายคน</span>
            </div>
            <div class="field-validation-error bilingual-error" id="statusEmailCcError"></div>
          </label>""",
    )
    html = html.replace(
        """          <label class="form-field full">
            <span>Email Body</span>
            <textarea id="statusEmailBody"></textarea>
          </label>""",
        """          <label class="form-field full">
            <span>Message</span>
            <textarea id="statusEmailBody"></textarea>
          </label>
          <div class="form-field full" id="statusEmailAttachmentPreview">
            <span>Attachments</span>
            <strong>-</strong>
          </div>""",
    )
    html = html.replace(
        """                      <button class="ghost-button compact-attach-button" type="button" id="updateAttachFilesBtn">
                        <span class="button-bilingual"><strong>Attach Files</strong><small>แนบไฟล์</small></span>
                      </button>""",
	        """                      <button class="ghost-button compact-attach-button" type="button" id="updateAttachFilesBtn">
	                        <span class="button-bilingual"><strong>Attach Files</strong><small>แนบไฟล์</small></span>
	                      </button>""",
    )
    html = html.replace(
        """                      <span class="en">Up to 10 files · 20 MB per file · PDF, Office, JPG or PNG</span>
                      <span class="th">สูงสุด 10 ไฟล์ · ไฟล์ละไม่เกิน 20 MB · PDF, Office, JPG หรือ PNG</span>""",
	        """                      <span class="en">Up to 10 files · 20 MB per file · system uploads attachments before sending email</span>
	                      <span class="th">สูงสุด 10 ไฟล์ · ไฟล์ละไม่เกิน 20 MB · ระบบอัปโหลดไฟล์ก่อนส่งอีเมล</span>""",
    )
    html = html.replace(
        """                    <div class="compact-help bilingual-stack">
                      <span class="en">Up to 10 files · 20 MB per file · system uploads attachments before sending email</span>
\t                      <span class="th">สูงสุด 10 ไฟล์ · ไฟล์ละไม่เกิน 20 MB · ระบบอัปโหลดไฟล์ก่อนส่งอีเมล</span>
                    </div>""",
        """                    <div class="compact-help bilingual-stack" id="updateAttachmentHelp">
                      <span class="en" id="updateAttachmentHelpEn">Up to 10 files · 20 MB per file · system uploads attachments before sending email</span>
\t                      <span class="th" id="updateAttachmentHelpTh">สูงสุด 10 ไฟล์ · ไฟล์ละไม่เกิน 20 MB · ระบบอัปโหลดไฟล์ก่อนส่งอีเมล</span>
                    </div>""",
    )
    html = html.replace(
        """    function validateUpdateAttachmentsForAction(action) {
      if (updateAttachmentQueue.length > UPDATE_MAX_FILES) {
        setAttachmentValidation("You can attach up to 10 files per update.", "สามารถแนบไฟล์ได้สูงสุด 10 ไฟล์ต่อการอัปเดตหนึ่งครั้ง");
        document.querySelector("#updateAttachmentSection")?.focus();
        return false;
      }
      if (action === "Resubmit" && updateAttachmentQueue.length === 0) {
        setAttachmentValidation("Please attach at least one revised document.", "กรุณาแนบเอกสารฉบับแก้ไขอย่างน้อย 1 ไฟล์");
        document.querySelector("#updateAttachmentSection")?.scrollIntoView({ behavior: "smooth", block: "center" });
        setTimeout(() => document.querySelector("#updateAttachmentSection")?.focus(), 180);
        return false;
      }
      setAttachmentValidation();
      return true;
    }

    function updateAttachmentRequirement(action) {
      const section = document.querySelector("#updateAttachmentSection");
      const labelEn = document.querySelector("#updateAttachmentLabelEn");
      const labelTh = document.querySelector("#updateAttachmentLabelTh");
      if (!section) return;
      const required = action === "Resubmit";
      section.classList.toggle("is-required", required);
      if (labelEn) labelEn.textContent = required ? "Attachments *" : "Attachments";
      if (labelTh) labelTh.textContent = required ? "ไฟล์แนบ *" : "ไฟล์แนบ";
      if (!required) setAttachmentValidation();
    }""",
        """    function updateActionRequiresAttachment(action) {
      // Resubmit is intentionally optional: corrected details may be submitted without a file.
      return false;
    }

    function validateUpdateAttachmentsForAction(action) {
      if (updateAttachmentQueue.length > UPDATE_MAX_FILES) {
        setAttachmentValidation("You can attach up to 10 files per update.", "สามารถแนบไฟล์ได้สูงสุด 10 ไฟล์ต่อการอัปเดตหนึ่งครั้ง");
        document.querySelector("#updateAttachmentSection")?.focus();
        return false;
      }
      if (updateActionRequiresAttachment(action) && updateAttachmentQueue.length === 0) {
        setAttachmentValidation("Please attach at least one required document.", "กรุณาแนบเอกสารที่กำหนดอย่างน้อย 1 ไฟล์");
        document.querySelector("#updateAttachmentSection")?.scrollIntoView({ behavior: "smooth", block: "center" });
        setTimeout(() => document.querySelector("#updateAttachmentSection")?.focus(), 180);
        return false;
      }
      setAttachmentValidation();
      return true;
    }

    function updateAttachmentRequirement(action) {
      const section = document.querySelector("#updateAttachmentSection");
      const labelEn = document.querySelector("#updateAttachmentLabelEn");
      const labelTh = document.querySelector("#updateAttachmentLabelTh");
      const helpEn = document.querySelector("#updateAttachmentHelpEn");
      const helpTh = document.querySelector("#updateAttachmentHelpTh");
      const fileInput = document.querySelector("#updateAttachmentInput");
      if (!section) return;
      const required = updateActionRequiresAttachment(action);
      section.classList.toggle("is-required", required);
      if (labelEn) labelEn.textContent = required ? "Attachments *" : "Attachments";
      if (labelTh) labelTh.textContent = required ? "ไฟล์แนบ *" : "ไฟล์แนบ";
      if (fileInput) {
        fileInput.required = required;
        if (!required) fileInput.removeAttribute("required");
        fileInput.setCustomValidity("");
      }
      if (action === "Resubmit") {
        if (helpEn) helpEn.textContent = "Optional for Resubmit.";
        if (helpTh) helpTh.textContent = "ไม่บังคับแนบไฟล์สำหรับการส่งกลับเข้าตรวจ";
      } else {
        if (helpEn) helpEn.textContent = "Up to 10 files · 20 MB per file · system uploads attachments before sending email";
        if (helpTh) helpTh.textContent = "สูงสุด 10 ไฟล์ · ไฟล์ละไม่เกิน 20 MB · ระบบอัปโหลดไฟล์ก่อนส่งอีเมล";
      }
      if (!required) setAttachmentValidation();
    }""",
    )
    html = html.replace(
        """    function prepareUploadedAttachments({ contractId, logId, action, updatedBy, uploadedAt }) {
      return updateAttachmentQueue.map((item, index) => ({
        fileId: `FILE-${Date.now()}-${String(index + 1).padStart(2, "0")}`,
        fileName: item.fileName,
        originalFileName: item.originalFileName,
        fileSize: item.fileSize,
        mimeType: item.mimeType,
        contractId,
        logId,
        action,
        uploadedBy: updatedBy,
        uploadedAt,
        status: "Uploaded",
        url: `https://contract-tracking.local/files/${encodeURIComponent(item.fileName)}`
      }));
    }""",
        """    function prepareUploadedAttachments({ contractId, logId, action, updatedBy, uploadedAt }) {
      return updateAttachmentQueue.map((item, index) => ({
        fileId: `FILE-${Date.now()}-${String(index + 1).padStart(2, "0")}`,
        fileName: item.fileName,
        originalFileName: item.originalFileName,
        fileSize: item.fileSize,
        mimeType: item.mimeType,
        file: item.file,
        contractId,
        logId,
        action,
        uploadedBy: updatedBy,
        uploadedAt,
        status: "Ready to send",
        url: item.cloudUrl || "",
        downloadUrl: item.downloadUrl || "",
        cloudFolderUrl: attachmentCloudConfig.folderUrl
      }));
    }""",
    )
    html = html.replace(
        """    function buildStatusEmailDraft({ contract, action, from, to, email, ccRecipients = [], attachments = [], alert, actionReasonTypeTh, actionReasonTypeEn, actionReason, days, sla }) {""",
        """    let activeStatusEmailDraft = null;
    let statusEmailCcRecipients = [];

    function statusEmailCcValues() {
      return statusEmailCcRecipients.map(item => item.email).filter(Boolean);
    }

    function statusEmailCcText() {
      return statusEmailCcValues().join(", ");
    }

    function statusEmailCcItemsFrom(value) {
      const rawItems = Array.isArray(value) ? value : String(value || "").split(/[;,\\n]+/);
      const seen = new Set();
      return rawItems
        .map(item => String(typeof item === "object" ? item.email || item.value || "" : item).trim())
        .filter(Boolean)
        .map(raw => {
          const directoryItem = employeeForValue(raw) || employeeDirectory.find(item => normalizeDirectoryValue(item.email) === normalizeDirectoryValue(raw));
          const email = String(directoryItem?.email || raw).trim().toLowerCase();
          return { name: directoryItem?.name || email, email };
        })
        .filter(item => {
          if (!updateEmailPattern(item.email) || seen.has(item.email)) return false;
          seen.add(item.email);
          return true;
        });
    }

    function setStatusEmailCcValidation(messageEn = "", messageTh = "") {
      const field = document.querySelector("#statusEmailCcField");
      const input = document.querySelector("#statusEmailCc");
      const error = document.querySelector("#statusEmailCcError");
      const hasError = Boolean(messageEn || messageTh);
      field?.classList.toggle("validation-error", hasError);
      input?.setAttribute("aria-invalid", String(hasError));
      if (error) error.innerHTML = hasError ? `<span class="en">${escapeHtml(messageEn)}</span><span class="th">${escapeHtml(messageTh)}</span>` : "";
    }

    function renderStatusEmailCcRecipients() {
      const chips = document.querySelector("#statusEmailCcChips");
      if (!chips) return;
      chips.innerHTML = statusEmailCcRecipients.map(item => `
        <span class="cc-chip" data-status-email-cc="${escapeHtml(item.email)}">
          <span class="cc-chip-copy"><strong>${escapeHtml(item.name || item.email)}</strong><small>${escapeHtml(item.email)}</small></span>
          <button class="cc-chip-remove" type="button" title="Remove" aria-label="Remove ${escapeHtml(item.email)}" data-remove-status-email-cc="${escapeHtml(item.email)}">×</button>
        </span>`).join("");
    }

    function addStatusEmailCcRecipient(rawValue, { silentEmpty = false } = {}) {
      const input = document.querySelector("#statusEmailCc");
      const rawInput = String(rawValue || "").trim();
      const parts = rawInput.split(/[;,]+/).map(value => value.trim()).filter(Boolean);
      if (parts.length > 1) {
        for (const part of parts) {
          if (!addStatusEmailCcRecipient(part, { silentEmpty: true })) return false;
        }
        if (input) input.value = "";
        return true;
      }
      const raw = rawInput.replace(/[;,]+$/, "").trim();
      if (!raw) {
        if (!silentEmpty) setStatusEmailCcValidation();
        return true;
      }
      const emailMatch = raw.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,}/i);
      const directoryItem = employeeForValue(raw) || employeeDirectory.find(item => normalizeDirectoryValue(item.email) === normalizeDirectoryValue(emailMatch?.[0] || raw));
      const email = String(directoryItem?.email || emailMatch?.[0] || raw).trim().toLowerCase();
      if (!updateEmailPattern(email)) {
        setStatusEmailCcValidation("Enter a valid email address.", "กรุณากรอกอีเมลให้ถูกต้อง");
        input?.focus();
        return false;
      }
      if (statusEmailCcRecipients.some(item => item.email.toLowerCase() === email)) {
        if (input) input.value = "";
        setStatusEmailCcValidation();
        return true;
      }
      statusEmailCcRecipients.push({ name: directoryItem?.name || email, email });
      if (input) input.value = "";
      setStatusEmailCcValidation();
      renderStatusEmailCcRecipients();
      return true;
    }

    function commitPendingStatusEmailCc() {
      const input = document.querySelector("#statusEmailCc");
      const raw = String(input?.value || "").trim();
      return raw ? addStatusEmailCcRecipient(raw) : true;
    }

    function initializeStatusEmailCcControl() {
      const datalist = document.querySelector("#statusEmailCcOptions");
      if (datalist) {
        datalist.innerHTML = employeeDirectory
          .filter(item => item.email)
          .map(item => `<option value="${escapeHtml(item.email)}" label="${escapeHtml([item.name, item.department].filter(Boolean).join(" — "))}"></option>`)
          .join("");
      }
      const input = document.querySelector("#statusEmailCc");
      input?.addEventListener("keydown", event => {
        if (["Enter", ",", ";"].includes(event.key)) {
          event.preventDefault();
          addStatusEmailCcRecipient(event.currentTarget.value);
        }
        if (event.key === "Backspace" && !event.currentTarget.value && statusEmailCcRecipients.length) {
          statusEmailCcRecipients.pop();
          renderStatusEmailCcRecipients();
        }
      });
      input?.addEventListener("change", event => {
        if (event.currentTarget.value.trim()) addStatusEmailCcRecipient(event.currentTarget.value);
      });
      input?.addEventListener("input", () => setStatusEmailCcValidation());
      document.addEventListener("click", event => {
        const removeCc = event.target.closest("[data-remove-status-email-cc]");
        if (!removeCc) return;
        event.preventDefault();
        statusEmailCcRecipients = statusEmailCcRecipients.filter(item => item.email !== removeCc.dataset.removeStatusEmailCc);
        setStatusEmailCcValidation();
        renderStatusEmailCcRecipients();
      });
    }

    function fileToBase64(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.addEventListener("load", () => {
          const result = String(reader.result || "");
          resolve(result.includes(",") ? result.split(",").pop() : result);
        });
        reader.addEventListener("error", () => reject(reader.error || new Error("Cannot read attachment")));
        reader.readAsDataURL(file);
      });
    }

    function configuredAttachmentUploadEndpoint() {
      const endpoint = String(attachmentCloudConfig.uploadEndpoint || "").trim();
      if (endpoint && !endpoint.includes("PASTE_")) return endpoint;
      return String(localStorage.getItem("contractTrackingAttachmentUploadEndpoint") || "").trim();
    }

    function isAttachmentCloudUploadConfigured() {
      return Boolean(configuredAttachmentUploadEndpoint());
    }

    function ensureAttachmentUploadEndpoint() {
      const saved = configuredAttachmentUploadEndpoint();
      if (saved) return saved;
      const endpoint = window.prompt("Paste Apps Script Web App URL for sending email and uploading attachments:");
      const trimmed = String(endpoint || "").trim();
      if (trimmed) localStorage.setItem("contractTrackingAttachmentUploadEndpoint", trimmed);
      return trimmed;
    }

    async function uploadQueuedAttachmentsToCloud() {
      if (!updateAttachmentQueue.length) return true;
      updateAttachmentQueue.forEach(item => {
        if (!item.cloudUrl) item.status = "Ready to send";
      });
      renderUpdateAttachmentQueue();
      setAttachmentValidation();
      return true;
    }

    async function buildEmailAttachmentPayload(item) {
      const attachment = {
        fileName: item.fileName || item.originalFileName || item.file?.name || "attachment",
        originalFileName: item.originalFileName || item.fileName || item.file?.name || "attachment",
        fileSize: item.fileSize || item.file?.size || 0,
        mimeType: item.mimeType || item.file?.type || "application/octet-stream",
        cloudFileId: item.cloudFileId || "",
        cloudUrl: item.cloudUrl || "",
        downloadUrl: item.downloadUrl || ""
      };
      if (!attachment.cloudUrl && item.file instanceof File) attachment.base64 = await fileToBase64(item.file);
      return attachment;
    }

    async function sendStatusEmailViaEndpoint(draft) {
      const endpoint = ensureAttachmentUploadEndpoint();
      if (!endpoint) {
        throw new Error("Missing Apps Script Web App URL. Deploy the Apps Script and paste the URL into ATTACHMENT_UPLOAD_ENDPOINT.");
      }
      updateAttachmentQueue.forEach(item => {
        item.status = "Sending email";
      });
      renderUpdateAttachmentQueue();
      const response = await fetch(endpoint, {
        method: "POST",
        mode: "no-cors",
        redirect: "follow",
        headers: { "Content-Type": "text/plain" },
        body: JSON.stringify({
          mode: "sendStatusEmail",
          folderId: attachmentCloudConfig.folderId,
          folderUrl: attachmentCloudConfig.folderUrl,
          to: draft.to,
          ...(draft.ccList?.length ? { cc: draft.ccList } : {}),
          ...(draft.cc ? { ccText: draft.cc } : {}),
          subject: draft.subject,
          body: draft.body,
          attachments: await Promise.all((draft.attachments || []).map(buildEmailAttachmentPayload)),
          contractId: draft.contractId || "",
          action: draft.action || "",
          sentFrom: "T23 Contract Tracking"
        })
      });
      if (response.type === "opaque") {
        updateAttachmentQueue.forEach(item => {
          item.status = "Email sent";
        });
        renderUpdateAttachmentQueue();
        return {
          success: true,
          sent: true,
          files: (draft.attachments || []).map(item => ({
            fileName: item.fileName,
            originalFileName: item.originalFileName
          }))
        };
      }
      const result = await response.json().catch(() => ({}));
      if (!response.ok || !result.success) throw new Error(result.error || "Send email failed");
      (result.files || []).forEach(file => {
        const match = updateAttachmentQueue.find(item =>
          item.fileName === file.fileName || item.originalFileName === file.originalFileName
        );
        if (!match) return;
        match.cloudFileId = file.id || match.cloudFileId || "";
        match.cloudUrl = file.url || match.cloudUrl || "";
        match.downloadUrl = file.downloadUrl || match.downloadUrl || "";
        match.status = "Email sent";
      });
      renderUpdateAttachmentQueue();
      return result;
    }

    function buildStatusEmailDraft({ contract, action, from, to, email, ccRecipients = [], attachments = [], alert, actionReasonTypeTh, actionReasonTypeEn, actionReason, days, sla }) {""",
    )
    html = html.replace(
        """      if (attachments.length) {
        details.push(
          "",
          `Attachments: ${attachments.length} file${attachments.length === 1 ? "" : "s"}`,
          `ไฟล์แนบ: ${attachments.length} ไฟล์`,
          "",
          ...attachments.flatMap((file, index) => [
            `${index + 1}. ${file.originalFileName || file.fileName}`,
            `   View Attachment: ${file.url || "-"}`
          ]),
          "",
          "View Attachments",
          "เปิดดูไฟล์แนบ"
        );
      }""",
        """      // Attachments are uploaded, linked, and attached by the backend after sending.""",
    )
    html = html.replace(
        """      document.querySelector("#statusEmailBody").value = draft.body || "";
      document.querySelector("#statusEmailSubtitle").textContent = draft.to
        ? `Recipient linked automatically: ${draft.to}${draft.cc ? ` · CC: ${draft.cc}` : ""}${draft.attachments?.length ? ` · ${draft.attachments.length} attachment link(s)` : ""}`
        : "ไม่พบอีเมลจากช่อง To กรุณากรอกอีเมลก่อนส่ง";""",
        """      activeStatusEmailDraft = draft;
      document.querySelector("#statusEmailBody").value = draft.body || "";
      document.querySelector("#statusEmailSubtitle").textContent = draft.to
        ? `Recipient linked automatically: ${draft.to}${draft.cc ? ` · CC: ${draft.cc}` : ""}${draft.attachments?.length ? ` · ${draft.attachments.length} cloud attachment(s)` : ""}`
        : "ไม่พบอีเมลจากช่อง To กรุณากรอกอีเมลก่อนส่ง";""",
    )
    html = html.replace(
        """    function currentStatusEmailDraft() {
      return {
        to: document.querySelector("#statusEmailTo")?.value.trim() || "",
        cc: document.querySelector("#statusEmailCc")?.value.trim() || "",
        subject: document.querySelector("#statusEmailSubject")?.value || "",
        body: document.querySelector("#statusEmailBody")?.value || ""
      };
    }""",
        """    function currentStatusEmailDraft() {
      return {
        ...(activeStatusEmailDraft || {}),
        to: document.querySelector("#statusEmailTo")?.value.trim() || "",
        cc: document.querySelector("#statusEmailCc")?.value.trim() || "",
        subject: document.querySelector("#statusEmailSubject")?.value || "",
        body: document.querySelector("#statusEmailBody")?.value || ""
      };
    }""",
    )
    html = html.replace(
        """    function openStatusEmailPopup(draft) {
      const modal = document.querySelector("#statusEmailModal");
      if (!modal) return;
      document.querySelector("#statusEmailTo").value = draft.to || "";
      const ccField = document.querySelector("#statusEmailCcField");
      const ccInput = document.querySelector("#statusEmailCc");
      if (ccInput) ccInput.value = draft.cc || "";
      if (ccField) ccField.hidden = !draft.cc;
      document.querySelector("#statusEmailSubject").value = draft.subject || "";
      activeStatusEmailDraft = draft;
      document.querySelector("#statusEmailBody").value = draft.body || "";
      document.querySelector("#statusEmailSubtitle").textContent = draft.to
        ? `Recipient linked automatically: ${draft.to}${draft.cc ? ` · CC: ${draft.cc}` : ""}${draft.attachments?.length ? ` · ${draft.attachments.length} cloud attachment(s)` : ""}`
        : "ไม่พบอีเมลจากช่อง To กรุณากรอกอีเมลก่อนส่ง";
      modal.classList.add("show");
      modal.setAttribute("aria-hidden", "false");
      setTimeout(() => document.querySelector(draft.to ? "#openMailClientBtn" : "#statusEmailTo")?.focus(), 0);
    }""",
        """    function openStatusEmailPopup(draft) {
      const modal = document.querySelector("#statusEmailModal");
      if (!modal) return;
      document.querySelector("#statusEmailTo").value = draft.to || "";
      statusEmailCcRecipients = statusEmailCcItemsFrom(draft.ccList || draft.cc || "");
      document.querySelector("#statusEmailSubject").value = draft.subject || "";
      activeStatusEmailDraft = { ...draft, cc: statusEmailCcText(), ccList: statusEmailCcValues() };
      document.querySelector("#statusEmailBody").value = draft.body || "";
      const attachmentSummary = document.querySelector("#statusEmailAttachmentPreview strong");
      if (attachmentSummary) {
        const count = draft.attachments?.length || 0;
        attachmentSummary.textContent = count ? `${count} file${count === 1 ? "" : "s"} / ${count} ไฟล์` : "-";
      }
      renderStatusEmailCcRecipients();
      setStatusEmailCcValidation();
      document.querySelector("#statusEmailSubtitle").textContent = draft.to
        ? `Recipient linked automatically: ${draft.to}${statusEmailCcText() ? ` · CC: ${statusEmailCcText()}` : ""}${draft.attachments?.length ? ` · ${draft.attachments.length} cloud attachment(s)` : ""}`
        : "ไม่พบอีเมลจากช่อง To กรุณากรอกอีเมลก่อนส่ง";
      modal.classList.add("show");
      modal.setAttribute("aria-hidden", "false");
      setTimeout(() => document.querySelector(draft.to ? "#openMailClientBtn" : "#statusEmailTo")?.focus(), 0);
    }""",
    )
    html = html.replace(
        """    function currentStatusEmailDraft() {
      return {
        ...(activeStatusEmailDraft || {}),
        to: document.querySelector("#statusEmailTo")?.value.trim() || "",
        cc: document.querySelector("#statusEmailCc")?.value.trim() || "",
        subject: document.querySelector("#statusEmailSubject")?.value || "",
        body: document.querySelector("#statusEmailBody")?.value || ""
      };
    }""",
        """    function currentStatusEmailDraft() {
      if (!commitPendingStatusEmailCc()) return null;
      const ccList = statusEmailCcValues();
      return {
        ...(activeStatusEmailDraft || {}),
        to: document.querySelector("#statusEmailTo")?.value.trim() || "",
        cc: ccList.join(", "),
        ccList,
        subject: document.querySelector("#statusEmailSubject")?.value || "",
        body: document.querySelector("#statusEmailBody")?.value || ""
      };
    }""",
    )
    html = html.replace(
        """      return { to: email || "", cc: ccRecipients.map(item => item.email).join(", "), subject, body, attachments };""",
        """      return { to: email || "", cc: ccRecipients.map(item => item.email).join(", "), subject, body, attachments, contractId: contract.id, action: info.nameEn };""",
    )
    html = replace_between(
        html,
        """    function buildStatusEmailDraft({ contract, action, from, to, email, ccRecipients = [], attachments = [], alert, actionReasonTypeTh, actionReasonTypeEn, actionReason, days, sla }) {""",
        """

    function openStatusEmailPopup(draft) {""",
        """    function buildStatusEmailDraft({ contract, action, from, to, email, ccRecipients = [], attachments = [], alert, actionReasonTypeTh, actionReasonTypeEn, actionReason, days, sla }) {
      const info = actionDefinition(action);
      const safeText = value => String(value ?? "").trim();
      const contractId = safeText(contract?.id);
      const contractName = safeText(contract?.name);
      const recipientName = safeText(to) || safeText(email) || "Team";
      const fromName = safeText(from);
      const toName = safeText(to) || safeText(email);
      const dueDate = safeText(formatDateForEmail(contract?.due)).replace(/^[-–]$/, "");
      const systemUrl = "https://turtletwentythree.github.io/Project-Contract-tracking/";
      const slaNumber = Number(sla) || 0;
      const copyByAction = {
        "Submit to Review": {
          actionEn: "Submit to Review",
          actionTh: "ส่งตรวจสอบ",
          descriptionEn: "Submit the contract to the review process.",
          descriptionTh: "ส่งสัญญาเข้าสู่กระบวนการตรวจสอบ",
          reasonEn: "Submission Remark:",
          reasonTh: "หมายเหตุการส่งตรวจ:"
        },
        Return: {
          actionEn: "Return",
          actionTh: "ส่งกลับเพื่อแก้ไข",
          descriptionEn: "Return the contract for correction or additional information.",
          descriptionTh: "ส่งสัญญากลับเพื่อแก้ไขหรือเพิ่มเติมข้อมูล",
          reasonEn: "Return Reason:",
          reasonTh: "เหตุผลที่ส่งกลับ:"
        },
        Resubmit: {
          actionEn: "Resubmit",
          actionTh: "ส่งกลับเข้าตรวจอีกครั้ง",
          descriptionEn: "Resubmit the corrected contract to the review process.",
          descriptionTh: "ส่งสัญญากลับเข้าสู่กระบวนการตรวจสอบอีกครั้งหลังแก้ไข",
          reasonEn: "Resubmission Remark:",
          reasonTh: "หมายเหตุการส่งกลับเข้าตรวจ:"
        },
        Forward: {
          actionEn: "Forward",
          actionTh: "ส่งต่อ",
          descriptionEn: "Forward the contract to the next department or responsible person.",
          descriptionTh: "ส่งต่อสัญญาไปยังหน่วยงานหรือผู้รับผิดชอบลำดับถัดไป",
          reasonEn: "Forward Remark:",
          reasonTh: "หมายเหตุการส่งต่อ:"
        },
        Approved: {
          actionEn: "Approved",
          actionTh: "อนุมัติ",
          descriptionEn: "Confirm that the contract has been approved.",
          descriptionTh: "ยืนยันว่าสัญญาได้รับการอนุมัติแล้ว",
          reasonEn: "Approval Remark:",
          reasonTh: "หมายเหตุการอนุมัติ:"
        }
      };
      const copy = copyByAction[info.nameEn] || {
        actionEn: safeText(info.nameEn) || safeText(action),
        actionTh: safeText(info.nameTh) || safeText(action),
        descriptionEn: safeText(info.shortEn),
        descriptionTh: safeText(info.shortTh),
        reasonEn: actionReasonTypeEn ? `${actionReasonTypeEn}:` : "Reason / Remark:",
        reasonTh: actionReasonTypeTh ? `${actionReasonTypeTh}:` : "เหตุผล / หมายเหตุ:"
      };
      const subject = `[Contract Tracking] ${info.nameEn}: ${contractId} - ${contractName}`;
      const reasonText = safeText(actionReason);
      const englishLines = [
        `Dear ${recipientName},`,
        "",
        `Contract ID: ${contractId}`,
        `Contract Name: ${contractName}`,
        "",
        `Action: ${copy.actionEn}`,
        `Description: ${copy.descriptionEn}`,
        `From: ${fromName}`,
        `To: ${toName}`,
        `Action SLA: ${slaNumber} Working Day${slaNumber === 1 ? "" : "s"}`,
        `Due Date: ${dueDate}`,
        `System Link: ${systemUrl}`
      ];
      if (reasonText) englishLines.push("", copy.reasonEn, reasonText);
      englishLines.push("", "Please review and proceed with the contract accordingly.");

      const thaiLines = [
        `เรียน ${recipientName},`,
        "",
        `รหัสสัญญา: ${contractId}`,
        `ชื่อสัญญา: ${contractName}`,
        "",
        `การดำเนินการ: ${copy.actionTh}`,
        `คำอธิบาย: ${copy.descriptionTh}`,
        `ส่งจาก: ${fromName}`,
        `ส่งถึง: ${toName}`,
        `SLA ของขั้นตอน: ${slaNumber} วันทำการ`,
        `วันครบกำหนด: ${dueDate}`
      ];
      if (reasonText) thaiLines.push("", copy.reasonTh, reasonText);
      thaiLines.push("", "กรุณาตรวจสอบและดำเนินการตามขั้นตอนที่เกี่ยวข้อง", "", "Contract Tracking System");

      thaiLines.splice(Math.max(thaiLines.length - 1, 0), 0, "", `System Link: ${systemUrl}`);
      const body = [...englishLines, "", "", ...thaiLines].join("\\n");
      return {
        to: email || "",
        cc: ccRecipients.map(item => item.email).join(", "),
        ccList: ccRecipients.map(item => item.email),
        subject,
        body,
        attachments,
        contractId,
        action: info.nameEn
      };
    }""",
    )
    html = replace_between(
        html,
        """    function currentAddCaseAccessLevel() {""",
        """

    function uniqueDepartments() {""",
        f"""    const departmentCodeConfig = Object.freeze({js(DEPARTMENT_CODE_CONFIG)});

	    function departmentCodeFor(department) {{
	      const raw = String(department || "").trim();
	      if (!raw) return "";
	      const masterMatch = activeMasterDepartments().find(item => normalizeDirectoryValue(item["Department / Restaurant"]) === normalizeDirectoryValue(raw));
	      if (masterMatch?.["Department Code"]) return String(masterMatch["Department Code"]).trim().toUpperCase();
      if (departmentCodeConfig[raw]) return departmentCodeConfig[raw];
      const normalized = normalizeDirectoryValue(raw);
	      const match = Object.keys(departmentCodeConfig).find(key => normalizeDirectoryValue(key) === normalized);
	      return match ? departmentCodeConfig[match] : "";
	    }}

	    function departmentCodeSuggestion(department) {{
	      const configured = departmentCodeFor(department);
	      if (configured) return configured;
	      const raw = String(department || "").trim();
	      if (!raw) return "";
	      const letters = raw
	        .replace(/[^A-Za-z0-9\\s]/g, " ")
	        .split(/\\s+/)
	        .filter(Boolean);
	      const code = letters.length > 1
	        ? letters.map(part => part[0]).join("")
	        : raw.replace(/[^A-Za-z0-9]/g, "").slice(0, 6);
	      return String(code || "").toUpperCase();
	    }}

	    function accessLevelCodeFor(accessLevel) {{
      return String(accessLevel || "").trim() === "Confidential" ? "C" : "N";
    }}

    function currentAddCaseAccessLevel() {{
      const form = document.querySelector("#addCaseForm");
      const explicitLevel = String(form?.elements.accessLevel?.value || "").trim();
      if (explicitLevel) return explicitLevel;
      const contractType = String(form?.elements.type?.value || "").trim();
      return contractType ? accessLevelForContractType(contractType) : "";
    }}

    function currentAddCaseDepartment() {{
      const form = document.querySelector("#addCaseForm");
      return String(form?.elements.department?.value || "").trim();
    }}

    function parseContractIdParts(contractId) {{
      const match = String(contractId || "").trim().match(/^CT-([NC])-([A-Z0-9]+)-(\\d+)$/i);
      if (!match) return null;
      return {{
        classCode: match[1].toUpperCase(),
        departmentCode: match[2].toUpperCase(),
        sequence: Number(match[3])
      }};
    }}

    function nextContractId(accessLevelOrOptions = "", departmentValue = "") {{
      const options = typeof accessLevelOrOptions === "object" && accessLevelOrOptions !== null
        ? accessLevelOrOptions
        : {{ accessLevel: accessLevelOrOptions, department: departmentValue }};
      const level = String(options.accessLevel || currentAddCaseAccessLevel()).trim();
      const department = String(options.department || currentAddCaseDepartment()).trim();
      const departmentCode = departmentCodeFor(department);
      if (!level || !departmentCode) return "";
      const classCode = accessLevelCodeFor(level);
      const maxSequence = contracts.reduce((max, item) => {{
        const parts = parseContractIdParts(item.id);
        if (!parts || parts.classCode !== classCode || parts.departmentCode !== departmentCode) return max;
        return Number.isFinite(parts.sequence) ? Math.max(max, parts.sequence) : max;
      }}, 0);
      return `CT-${{classCode}}-${{departmentCode}}-${{String(maxSequence + 1).padStart(3, "0")}}`;
    }}

    function setAddDepartmentCodeValidity(messageEn = "", messageTh = "") {{
      const input = document.querySelector("#addDepartment");
      if (!input) return;
      input.setCustomValidity(messageEn || "");
      input.setAttribute("aria-invalid", messageEn ? "true" : "false");
      if (messageEn || messageTh) showToast(`${{messageEn}} / ${{messageTh}}`);
    }}

    function validateAddCaseContractId(options = {{}}) {{
      const form = document.querySelector("#addCaseForm");
      const accessLevel = String(options.accessLevel || currentAddCaseAccessLevel()).trim();
      const department = String(options.department || currentAddCaseDepartment()).trim();
      setAddDepartmentCodeValidity();
      document.querySelector("#addContractType")?.setCustomValidity("");
      if (!department) {{
        setAddDepartmentCodeValidity("Please select Department / Restaurant.", "กรุณาเลือกแผนกหรือร้านอาหาร");
        document.querySelector("#addDepartment")?.reportValidity();
        return false;
      }}
      if (!departmentCodeFor(department)) {{
        setAddDepartmentCodeValidity("Department Code is not configured.", "ยังไม่ได้กำหนดรหัสย่อของแผนกหรือร้านอาหาร");
        document.querySelector("#addDepartment")?.reportValidity();
        return false;
      }}
      if (!accessLevel) {{
        document.querySelector("#addContractType")?.setCustomValidity("Please select Type of Contract.");
        document.querySelector("#addContractType")?.reportValidity();
        showToast("Please select Type of Contract. / กรุณาเลือกประเภทสัญญา");
        return false;
      }}
      const id = nextContractId({{ accessLevel, department }});
      if (!id) {{
        showToast("Cannot generate Contract ID. / ไม่สามารถสร้างรหัสสัญญาได้");
        return false;
      }}
      if (contracts.some(item => String(item.id || "").trim() === id)) {{
        showToast("This Contract ID already exists. / มีรหัสสัญญานี้อยู่ในระบบแล้ว");
        return false;
      }}
      return Boolean(form);
    }}

    function refreshNewCaseIdPreview() {{
      const badge = document.querySelector("#newCaseId");
      if (!badge) return;
      setAddDepartmentCodeValidity();
      const department = currentAddCaseDepartment();
      const accessLevel = currentAddCaseAccessLevel();
      const departmentCode = departmentCodeFor(department);
      const nextId = nextContractId({{ accessLevel, department }});
      badge.textContent = nextId || (department && accessLevel ? "Department Code missing" : "Select Department / Access Level");
      badge.title = departmentCode ? `Department Code: ${{departmentCode}}` : "Department Code is not configured.";
      badge.classList.toggle("amber", Boolean(department && !departmentCode));
    }}""",
    )
    html = html.replace(
        """      attachEditableDropdown("addDepartment", directoryDepartmentOptions, () => {
        syncDepartmentOwnerOptions(true);
        refreshEditableDropdown("addOwner");
      });""",
        """      attachEditableDropdown("addDepartment", directoryDepartmentOptions, () => {
        syncDepartmentOwnerOptions(true);
        refreshEditableDropdown("addOwner");
        syncAddCaseSystemFields();
      });""",
    )
    html = html.replace(
        """    document.querySelector("#addDepartment")?.addEventListener("input", () => {
      syncDepartmentOwnerOptions(false);
      refreshEditableDropdown("addOwner");
    });
    document.querySelector("#addDepartment")?.addEventListener("change", () => {
      syncDepartmentOwnerOptions(true);
      refreshEditableDropdown("addOwner");
    });""",
        """    document.querySelector("#addDepartment")?.addEventListener("input", () => {
      syncDepartmentOwnerOptions(false);
      refreshEditableDropdown("addOwner");
      syncAddCaseSystemFields();
    });
    document.querySelector("#addDepartment")?.addEventListener("change", () => {
      syncDepartmentOwnerOptions(true);
      refreshEditableDropdown("addOwner");
      syncAddCaseSystemFields();
    });""",
    )
    html = html.replace(
        """    document.querySelector("#addContractName")?.addEventListener("input", () => {""",
        """    document.querySelector("#addContractClassification")?.addEventListener("change", () => {
      resetInitialDueDateOverride();
      syncAddCaseLinkedFields("classification");
	    });
	    document.querySelectorAll("[data-classification-value]").forEach(button => {
	      button.addEventListener("click", () => {
	        const input = document.querySelector("#addContractClassification");
	        if (!input) return;
	        input.value = classificationDisplayValue(button.dataset.classificationValue || "Day-to-day Work");
	        resetInitialDueDateOverride();
	        syncAddCaseLinkedFields("classification");
	      });
	    });
	    document.querySelector("#addContractSubType")?.addEventListener("input", () => {
      resetInitialDueDateOverride();
      syncAddCaseLinkedFields("subType");
    });
    document.querySelector("#addContractSubType")?.addEventListener("change", () => {
      resetInitialDueDateOverride();
      syncAddCaseLinkedFields("subType");
    });
    document.querySelector("#addContractName")?.addEventListener("input", () => {""",
    )
    html = html.replace(
        """      const accessLevel = String(form.get("accessLevel") || accessLevelForContractType(contractType) || "Normal").trim();
      const id = nextContractId(accessLevel);
      const owner = String(form.get("owner") || "").trim();""",
        """      const accessLevel = String(form.get("accessLevel") || accessLevelForContractType(contractType) || "Normal").trim();
      const department = String(form.get("department") || "").trim();
      if (!validateAddCaseContractId({ accessLevel, department })) return;
      const id = nextContractId({ accessLevel, department });
      const owner = String(form.get("owner") || "").trim();""",
    )
    html = html.replace(
        """        department: form.get("department"),""",
        """        department,""",
    )
    html = html.replace(
        """      if (form.id === "updateStatusForm" && !validateUpdateReasonBeforeAction()) return;
      openActionConfirmModal(form);""",
        """      if (form.id === "addCaseForm" && !validateAddCaseContractId()) return;
      if (form.id === "updateStatusForm" && !validateUpdateReasonBeforeAction()) return;
      openActionConfirmModal(form);""",
    )
    html = html.replace(
        """          ["Contract Owner", value("owner")],
          ["Type of Contract", value("type")],
          ["SLA", `${value("sla", "0")} Working Days`],""",
        """          ["Contract Owner", value("owner")],
          ["Type of Contract", value("type")],
          ["Sub Type of Contract", value("subType"), true],
          ["SLA", `${value("sla", "0")} Working Days`],""",
    )
    html = html.replace(
        """    function openStatusEmailClient() {
      const draft = currentStatusEmailDraft();
      if (!draft.to) {
        document.querySelector("#statusEmailTo")?.focus();
        showToast("Please enter the recipient email.");
        return;
      }
      const mailto = `mailto:${encodeURIComponent(draft.to)}?${draft.cc ? `cc=${encodeURIComponent(draft.cc)}&` : ""}subject=${encodeURIComponent(draft.subject)}&body=${encodeURIComponent(draft.body)}`;
      window.location.href = mailto;
      closeStatusEmailPopup();
    }""",
        """    async function openStatusEmailClient() {
      const draft = currentStatusEmailDraft();
      if (!draft.to) {
        document.querySelector("#statusEmailTo")?.focus();
        showToast("Please enter the recipient email.");
        return;
      }
      if (!updateEmailPattern(draft.to)) {
        document.querySelector("#statusEmailTo")?.focus();
        showToast("Enter a valid recipient email. / กรุณากรอกอีเมลผู้รับให้ถูกต้อง");
        return;
      }
      const button = document.querySelector("#openMailClientBtn");
      try {
        if (button) {
          button.disabled = true;
          button.textContent = "Sending...";
        }
        const result = await sendStatusEmailViaEndpoint(draft);
        showToast(`Status update email sent${result.files?.length ? ` with ${result.files.length} cloud attachment link(s)` : ""}.`);
        closeStatusEmailPopup();
      } catch (error) {
        const message = error?.message || "Cannot send email.";
        document.querySelector("#statusEmailSubtitle").textContent = message;
        showToast(message);
      } finally {
        if (button) {
          button.disabled = false;
          button.innerHTML = "<span>✉</span> Send Email";
        }
      }
    }""",
    )
    html = html.replace(
        """    async function openStatusEmailClient() {
      const draft = currentStatusEmailDraft();
      if (!draft.to) {
        document.querySelector("#statusEmailTo")?.focus();
        showToast("Please enter the recipient email.");
        return;
      }""",
        """    async function openStatusEmailClient() {
      const draft = currentStatusEmailDraft();
      if (!draft) {
        document.querySelector("#statusEmailCc")?.focus();
        return;
      }
      if (!draft.to) {
        document.querySelector("#statusEmailTo")?.focus();
        showToast("Please enter the recipient email.");
        return;
      }""",
    )
    html = html.replace(
        """    async function copyStatusEmailDraft() {
      const draft = currentStatusEmailDraft();
      const text = `To: ${draft.to}${draft.cc ? `\\nCC: ${draft.cc}` : ""}\\nSubject: ${draft.subject}\\n\\n${draft.body}`;""",
        """    async function copyStatusEmailDraft() {
      const draft = currentStatusEmailDraft();
      if (!draft) return;
      const attachmentCount = draft.attachments?.length || 0;
      const text = `To: ${draft.to}\\nCC: ${draft.cc || "-"}\\nSubject: ${draft.subject}\\nMessage:\\n${draft.body}\\n\\nAttachments: ${attachmentCount ? `${attachmentCount} file${attachmentCount === 1 ? "" : "s"}` : "-"}`;""",
    )
    html = html.replace(
        """    document.querySelector("#updateStatusForm").addEventListener("submit", event => {""",
        """    document.querySelector("#updateStatusForm").addEventListener("submit", async event => {""",
    )
    html = html.replace(
        """      const updatedAt = localIsoDateTime();
      const nextLogNo = logRecords.filter(row => row[0] === contractId).length + 1;""",
        """      if (!await uploadQueuedAttachmentsToCloud()) return;

      const updatedAt = localIsoDateTime();
      const nextLogNo = logRecords.filter(row => row[0] === contractId).length + 1;""",
    )
    html = html.replace(
        """      const fileInput = document.querySelector("#updateAttachmentInput");
      const attachButton = document.querySelector("#updateAttachFilesBtn");
      const dropzone = document.querySelector("#updateAttachmentDropzone");
      attachButton?.addEventListener("click", () => fileInput?.click());
      fileInput?.addEventListener("change", event => {
        addFilesToUpdateQueue(event.currentTarget.files);
        event.currentTarget.value = "";
      });""",
	        """      const fileInput = document.querySelector("#updateAttachmentInput");
	      const attachButton = document.querySelector("#updateAttachFilesBtn");
	      const dropzone = document.querySelector("#updateAttachmentDropzone");
	      attachButton?.addEventListener("click", () => fileInput?.click());
	      fileInput?.addEventListener("change", async event => {
	        addFilesToUpdateQueue(event.currentTarget.files);
	        event.currentTarget.value = "";
        if (isAttachmentCloudUploadConfigured()) await uploadQueuedAttachmentsToCloud();
      });""",
    )
    html = html.replace(
        """      dropzone?.addEventListener("drop", event => addFilesToUpdateQueue(event.dataTransfer?.files));""",
        """      dropzone?.addEventListener("drop", async event => {
        addFilesToUpdateQueue(event.dataTransfer?.files);
        if (isAttachmentCloudUploadConfigured()) await uploadQueuedAttachmentsToCloud();
      });""",
    )
    html = html.replace(
        """    document.querySelector("#copyStatusEmailBtn")?.addEventListener("click", copyStatusEmailDraft);""",
        """    document.querySelector("#copyStatusEmailBtn")?.addEventListener("click", copyStatusEmailDraft);
    initializeStatusEmailCcControl();""",
    )
    html = html.replace(
        """    function latestLogFor(contractId) {
      return logRecords.filter(row => row[0] === contractId).slice(-1)[0] || null;
    }""",
        """    function latestLogFor(contractId) {
      return logRecords.filter(row => row[0] === contractId).slice(-1)[0] || null;
    }

    function firstLogFor(contractId) {
      return logRecords.find(row => row[0] === contractId) || null;
    }

    function addCaseStartDateForContract(contract) {
      if (!contract) return "";
      return dateToInputValue(contract.addCaseDate || firstLogFor(contract.id)?.[6] || "");
    }

    function recalculateLogTiming(row) {
      if (!Array.isArray(row)) return row;
      row[9] = calculateDaysOnHand(row[6], row[7]);
      row[10] = calculateAlert(row[12], row[9], row[8]);
      return row;
    }

    function syncContractTimingFromLogs(contract) {
      if (!contract) return contract;
      const rows = logRecords.filter(row => row[0] === contract.id);
      rows.forEach(recalculateLogTiming);
      const latest = rows.slice(-1)[0] || null;
      const totalSla = Number(contract.totalSla || totalSlaFor(contract.workType || contract.type)) || 0;
      const used = rows.reduce((sum, row) => sum + (Number(row[9]) || 0), 0);
      const startDate = addCaseStartDateForContract(contract);
      contract.addCaseDate = startDate;
      contract.totalSla = totalSla;
      contract.used = used;
      contract.days = latest ? Number(latest[9]) || 0 : used;
      contract.balance = totalSla - used;
      contract.systemDue = startDate ? addBusinessDays(startDate, totalSla) : (contract.systemDue || "");
      if (!contract.due) contract.due = contract.systemDue;
      contract.alert = latest?.[10] || calculateAlert(contract.stage, contract.days, totalSla);
      contract.status = calculateContractStatus(contract.stage, used, totalSla);
      return contract;
    }""",
        1,
    )
    html = html.replace(
        """    function renderAll() {
      rebuildNotificationQueue();""",
        """    const contractDatabaseHeaders = [
      "Contract ID",
      "Contract Name",
      "Department / Restaurant",
      "Contract Owner",
      "Type of Contract",
      "Vendor / Counter party",
      "Stage",
      "Cycle",
      "Returns",
      "Status Update",
      "Station",
      "Station Owner",
      "Add Case Date",
      "Due Date",
      "System Due Date",
      "Work Type",
      "Total SLA",
      "Days Used",
      "Days on Hand",
      "Balance",
      "Alert",
      "Remark",
      "Access Level",
      "Visibility",
      "Category"
    ];

    const logDatabaseHeaders = [
      "Contract ID",
      "Log No",
      "Cycle",
      "Log View",
      "From",
      "To",
      "In",
      "Out",
      "SLA",
      "Days on Hand",
      "Alert",
      "Delay Reason",
      "Action",
      "Action Reason",
      "Approval",
      "Corrective Action",
      "Action Reason Type",
      "Action Reason Detail",
      "Approval Type",
      "Approval Conditions",
      "Corrective Action Detail",
      "Updated By",
      "Updated Date and Time",
      "Action Code",
      "Action Name TH",
      "Action Name EN",
      "Action Description TH",
      "Action Description EN",
      "Action SLA",
      "Action Reason Type TH",
      "Action Reason Type EN",
      "Attachments",
      "CC Recipients"
    ];

    const localDatabaseKey = `trackingContracts.csvDatabase.${driveDatabaseConfig.folderId}.v1`;
    let driveDatabaseSaveTimer = 0;
    let isApplyingDriveDatabaseLoad = false;

    function csvCell(value) {
      let text = value == null ? "" : value;
      if (Array.isArray(text) || typeof text === "object") text = JSON.stringify(text);
      text = String(text);
      if (/[",\\n\\r]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
      return text;
    }

    function objectsToCsv(headers, rows) {
      return [
        headers.map(csvCell).join(","),
        ...rows.map(row => headers.map(header => csvCell(row[header])).join(","))
      ].join("\\r\\n");
    }

    function parseCsvRows(text) {
      const rows = [];
      let row = [];
      let cell = "";
      let quoted = false;
      const source = String(text || "").replace(/^\\uFEFF/, "");
      for (let index = 0; index < source.length; index += 1) {
        const char = source[index];
        const next = source[index + 1];
        if (quoted) {
          if (char === '"' && next === '"') {
            cell += '"';
            index += 1;
          } else if (char === '"') {
            quoted = false;
          } else {
            cell += char;
          }
          continue;
        }
        if (char === '"') {
          quoted = true;
        } else if (char === ",") {
          row.push(cell);
          cell = "";
        } else if (char === "\\n") {
          row.push(cell);
          rows.push(row);
          row = [];
          cell = "";
        } else if (char !== "\\r") {
          cell += char;
        }
      }
      row.push(cell);
      rows.push(row);
      return rows.filter(values => values.some(value => String(value || "").trim()));
    }

    function csvToObjects(text) {
      const rows = parseCsvRows(text);
      if (!rows.length) return [];
      const headers = rows.shift().map(header => String(header || "").trim());
      return rows.map(values => Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""])));
    }

    function contractDbRow(item) {
      const station = stationParts(item.station);
      return {
        "Contract ID": item.id,
        "Contract Name": item.name,
        "Department / Restaurant": item.department,
        "Contract Owner": item.owner,
        "Type of Contract": contractPrimaryTypeDisplay(item.type),
        "Vendor / Counter party": item.vendor,
        "Stage": item.stage,
        "Cycle": item.cycle,
        "Returns": item.returns,
        "Status Update": item.status,
        "Station": item.station,
        "Station Owner": station.to,
        "Add Case Date": item.addCaseDate || firstLogFor(item.id)?.[6] || "",
        "Due Date": item.due,
        "System Due Date": item.systemDue || "",
        "Work Type": item.workType,
        "Total SLA": item.totalSla,
        "Days Used": item.used,
        "Days on Hand": item.days || item.used || 0,
        "Balance": item.balance,
        "Alert": item.alert,
        "Remark": item.remark,
        "Access Level": item.accessLevel || accessLevelForContractType(item.type),
        "Visibility": item.visibility || "",
        "Category": item.category || contractTypeCategoryFor(item.type)
      };
    }

    function contractFromDbRow(row) {
      const id = String(row["Contract ID"] || "").trim();
      const rawType = String(row["Type of Contract"] || row["Work Type"] || "Other").trim();
      const type = contractPrimaryTypeDisplay(rawType) || rawType || "Other";
      const owner = String(row["Contract Owner"] || row["Station Owner"] || "").trim();
      const stationOwner = String(row["Station Owner"] || "Legal").trim();
      const totalSla = Number(row["Total SLA"] || totalSlaFor(type)) || 0;
      const used = Number(row["Days Used"] || row["Days on Hand"] || 0) || 0;
      const accessLevel = String(row["Access Level"] || accessLevelForContractType(type) || "Normal").trim();
      return {
        id,
        name: String(row["Contract Name"] || "").trim(),
        department: String(row["Department / Restaurant"] || "").trim(),
        owner,
        type,
        vendor: String(row["Vendor / Counter party"] || "").trim(),
        stage: String(row["Stage"] || "Draft Created").trim(),
        cycle: Number(row["Cycle"] || 1) || 1,
        returns: Number(row["Returns"] || 0) || 0,
        status: String(row["Status Update"] || row["Alert"] || "Green >>G=On Track").trim(),
        station: String(row["Station"] || `From ${owner || "Owner"} >> To ${stationOwner}`).trim(),
        addCaseDate: String(row["Add Case Date"] || "").trim(),
        due: String(row["Due Date"] || "").trim(),
        systemDue: String(row["System Due Date"] || "").trim(),
        workType: String(row["Work Type"] || type || "Other").trim(),
        totalSla,
        used,
        days: Number(row["Days on Hand"] || used) || 0,
        balance: Number(row["Balance"] || (totalSla - used)) || 0,
        alert: String(row["Alert"] || row["Status Update"] || "Green >>G=On Track").trim(),
        remark: String(row["Remark"] || "").trim(),
        accessLevel,
        visibility: String(row["Visibility"] || (accessLevel === "Confidential" ? "Restricted access / จำกัดสิทธิ์" : "Standard access / สิทธิ์ทั่วไป")).trim(),
        category: String(row["Category"] || contractTypeCategoryFor(type)).trim()
      };
    }

    function logDbRow(row) {
      return Object.fromEntries(logDatabaseHeaders.map((header, index) => [header, row[index] ?? ""]));
    }

    function logFromDbRow(row) {
      return logDatabaseHeaders.map(header => {
        const value = row[header] ?? "";
        if (header === "Attachments" || header === "CC Recipients") {
          if (!value) return [];
          try {
            const parsed = JSON.parse(value);
            return Array.isArray(parsed) ? parsed : [];
          } catch (error) {
            return [];
          }
        }
        return value;
      });
    }

    function downloadTextFile(filename, text, mimeType = "text/csv;charset=utf-8") {
      const blob = new Blob(["\\uFEFF", text], { type: mimeType });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
    }

    function refreshDashboardDataFromContracts() {
      dashboardData.splice(0, dashboardData.length);
      contracts.forEach(contract => upsertDashboardData(contract));
      selectedContractIndex = Math.min(selectedContractIndex, Math.max(contracts.length - 1, 0));
      if (!contracts.some(item => item.id === selectedLogContractId)) selectedLogContractId = "all";
    }

    function updateDatabaseSyncStatus(message = "") {
      const status = document.querySelector("#databaseSyncStatus");
      if (!status) return;
      status.textContent = message || `${contracts.length} contracts · ${driveDatabaseConfig.contractsCsv}`;
    }

    function migrateContractIdsToDepartmentFormat(contractRows = [], logRows = []) {
      const oldIdPattern = /^CT-([NC])-26-(\\d+)$/i;
      const idMap = new Map();
      const usedIds = new Set(contractRows.map(item => String(item.id || "").trim()).filter(Boolean));
      const maxByGroup = new Map();

      contractRows.forEach(contract => {
        const parts = parseContractIdParts(contract.id);
        if (!parts) return;
        const key = `${parts.classCode}|${parts.departmentCode}`;
        maxByGroup.set(key, Math.max(maxByGroup.get(key) || 0, parts.sequence || 0));
      });

      contractRows.forEach(contract => {
        const oldId = String(contract.id || "").trim();
        const oldMatch = oldId.match(oldIdPattern);
        if (!oldMatch) return;
        const accessLevel = String(contract.accessLevel || (oldMatch[1].toUpperCase() === "C" ? "Confidential" : "Normal")).trim();
        const departmentCode = departmentCodeFor(contract.department);
        if (!accessLevel || !departmentCode) return;
        const classCode = accessLevelCodeFor(accessLevel);
        const key = `${classCode}|${departmentCode}`;
        let nextSequence = (maxByGroup.get(key) || 0) + 1;
        let newId = `CT-${classCode}-${departmentCode}-${String(nextSequence).padStart(3, "0")}`;
        while (usedIds.has(newId) && newId !== oldId) {
          nextSequence += 1;
          newId = `CT-${classCode}-${departmentCode}-${String(nextSequence).padStart(3, "0")}`;
        }
        maxByGroup.set(key, nextSequence);
        usedIds.delete(oldId);
        usedIds.add(newId);
        idMap.set(oldId, newId);
        contract.id = newId;
      });

      if (idMap.size && Array.isArray(logRows)) {
        logRows.forEach(row => {
          if (Array.isArray(row) && idMap.has(row[0])) row[0] = idMap.get(row[0]);
        });
      }
      return idMap.size;
    }

    function saveContractsDatabase() {
      try {
        localStorage.setItem(localDatabaseKey, JSON.stringify({
          savedAt: localIsoDateTime(),
          standardSlaDataVersion,
          departmentDataVersion,
          actionDataVersion,
          contracts,
          logRecords,
          dueDateAdjustmentRequests,
          masterData
        }));
        updateDatabaseSyncStatus(`${contracts.length} contracts saved locally`);
        scheduleDriveDatabaseSave();
      } catch (error) {
        updateDatabaseSyncStatus("Local database not saved");
      }
    }

    function loadContractsDatabase() {
      try {
        const raw = localStorage.getItem(localDatabaseKey);
        if (!raw) {
          updateDatabaseSyncStatus();
          return;
        }
        const parsed = JSON.parse(raw);
        if (!Array.isArray(parsed.contracts)) {
          updateDatabaseSyncStatus();
          return;
        }
        const parsedContracts = parsed.contracts;
        const parsedLogs = Array.isArray(parsed.logRecords) ? parsed.logRecords : [];
        if (Array.isArray(parsed.dueDateAdjustmentRequests)) {
          dueDateAdjustmentRequests.splice(0, dueDateAdjustmentRequests.length, ...parsed.dueDateAdjustmentRequests);
        }
        if (parsed.masterData && typeof parsed.masterData === "object") {
          if (Array.isArray(parsed.masterData.departments)) masterData.departments = parsed.masterData.departments;
	          if (Array.isArray(parsed.masterData.people)) masterData.people = parsed.masterData.people;
	          if (Array.isArray(parsed.masterData.contractTypes)) masterData.contractTypes = parsed.masterData.contractTypes;
	          if (Array.isArray(parsed.masterData.contractTemplates)) masterData.contractTemplates = parsed.masterData.contractTemplates;
	          if (Array.isArray(parsed.masterData.actionSla)) masterData.actionSla = parsed.masterData.actionSla;
	        }
        const migratedSlaCount = migrateMasterDataToCurrentSlaVersion(parsed.standardSlaDataVersion || "");
        const migratedDepartmentCount = migrateDepartmentMasterToCurrentVersion(parsed.departmentDataVersion || "");
        const migratedActionCount = migrateActionMasterToCurrentVersion(parsed.actionDataVersion || "");
        const migratedCount = migrateContractIdsToDepartmentFormat(parsedContracts, parsedLogs);
        parsedContracts.forEach(item => {
          item.department = canonicalDepartmentName(item.department);
          const normalizedType = contractPrimaryTypeDisplay(item.type);
          if (normalizedType) {
            item.type = normalizedType;
            item.category = contractTypeCategoryFor(normalizedType) || item.category || "";
          }
        });
        contracts.splice(0, contracts.length, ...parsedContracts);
        if (Array.isArray(parsed.logRecords)) logRecords.splice(0, logRecords.length, ...parsedLogs);
        if (migratedCount || migratedSlaCount || migratedDepartmentCount || migratedActionCount) saveContractsDatabase();
        refreshDashboardDataFromContracts();
        updateDatabaseSyncStatus(`${contracts.length} contracts loaded from local DB${migratedCount ? ` · ${migratedCount} Contract ID migrated` : ""}${migratedSlaCount ? ` · ${migratedSlaCount} SLA values updated` : ""}${migratedDepartmentCount ? ` · Department Master cleaned` : ""}${migratedActionCount ? ` · Action SLA updated` : ""}`);
      } catch (error) {
        updateDatabaseSyncStatus("Local database could not load");
      }
    }

    function driveDatabaseEndpoint() {
      return configuredAttachmentUploadEndpoint();
    }

    function typeMasterCsvText() {
      const rows = masterData.contractTypes || [];
      const headers = ["Contract Classification", "Type of Contract", "Sub Type of Contract", "Fixed SLA (Working Days)", "Standard SLA Version", "Active", "Category", "Description / คำอธิบาย"];
      return objectsToCsv(headers, rows);
    }

    function departmentMasterCsvText() {
      return objectsToCsv(["Department / Restaurant", "Department Code", "Department Data Version", "Active"], masterData.departments || []);
    }

    function peopleMasterCsvText() {
      return objectsToCsv(["company", "department", "name", "email", "active"], masterData.people || []);
    }

	    function contractTemplateCsvText() {
	      return objectsToCsv(["classification", "typeGroup", "subType", "name", "selectionLabel", "sourceRow", "type", "workType", "contractId", "accessLevel", "category", "department", "vendor", "group", "fixedSla", "slaVersion", "remark", "active"], masterData.contractTemplates || []);
	    }

	    function actionSlaCsvText() {
	      return objectsToCsv(["Action", "Description / รายละเอียด", "Fixed SLA (Working Days)", "SLA Rule / วิธีนับ", "Action Data Version", "Active"], masterData.actionSla || []);
	    }

	    function driveDatabaseCsvPayload() {
      return {
        mode: "saveDriveDatabase",
        standardSlaDataVersion,
        departmentDataVersion,
        actionDataVersion,
        folderId: driveDatabaseConfig.folderId,
        contractsCsv: driveDatabaseConfig.contractsCsv,
        logsCsv: driveDatabaseConfig.logsCsv,
        typeMasterCsv: driveDatabaseConfig.typeMasterCsv,
        departmentMasterCsv: driveDatabaseConfig.departmentMasterCsv,
	        peopleMasterCsv: driveDatabaseConfig.peopleMasterCsv,
	        contractTemplateCsv: driveDatabaseConfig.contractTemplateCsv,
	        actionSlaCsv: driveDatabaseConfig.actionSlaCsv,
	        contractsCsvText: objectsToCsv(contractDatabaseHeaders, contracts.map(contractDbRow)),
	        logsCsvText: objectsToCsv(logDatabaseHeaders, logRecords.map(logDbRow)),
	        typeMasterCsvText: typeMasterCsvText(),
	        departmentMasterCsvText: departmentMasterCsvText(),
	        peopleMasterCsvText: peopleMasterCsvText(),
	        contractTemplateCsvText: contractTemplateCsvText(),
	        actionSlaCsvText: actionSlaCsvText()
	      };
	    }

    function applyDriveDatabasePayload(payload) {
      if (!payload || payload.success === false) return false;
      const contractText = payload.contractsCsvText || payload.files?.contracts?.text || "";
      const logText = payload.logsCsvText || payload.files?.logs?.text || "";
      const typeMasterText = payload.typeMasterCsvText || payload.files?.typeMaster?.text || "";
      const departmentMasterText = payload.departmentMasterCsvText || payload.files?.departments?.text || "";
	      const peopleMasterText = payload.peopleMasterCsvText || payload.files?.people?.text || "";
	      const contractTemplateText = payload.contractTemplateCsvText || payload.files?.contractTemplates?.text || "";
	      const actionSlaText = payload.actionSlaCsvText || payload.files?.actionSla?.text || "";
	      if (typeMasterText) masterData.contractTypes = csvToObjects(typeMasterText);
	      if (departmentMasterText) masterData.departments = csvToObjects(departmentMasterText);
	      if (peopleMasterText) masterData.people = csvToObjects(peopleMasterText);
	      if (contractTemplateText) masterData.contractTemplates = csvToObjects(contractTemplateText);
	      if (actionSlaText) masterData.actionSla = csvToObjects(actionSlaText);
      const migratedSlaCount = migrateMasterDataToCurrentSlaVersion(payload.standardSlaDataVersion || "");
      const migratedDepartmentCount = migrateDepartmentMasterToCurrentVersion(payload.departmentDataVersion || "");
      const migratedActionCount = migrateActionMasterToCurrentVersion(payload.actionDataVersion || "");
      if (!contractText) return false;
      const cloudContracts = csvToObjects(contractText).map(contractFromDbRow).filter(item => item.id && item.name);
      cloudContracts.forEach(item => { item.department = canonicalDepartmentName(item.department); });
      const cloudLogs = logText ? csvToObjects(logText).map(logFromDbRow).filter(row => String(row?.[0] || "").trim()) : [];
      const migratedCount = cloudContracts.length ? migrateContractIdsToDepartmentFormat(cloudContracts, cloudLogs) : 0;
      isApplyingDriveDatabaseLoad = true;
      contracts.splice(0, contracts.length, ...cloudContracts);
      logRecords.splice(0, logRecords.length, ...cloudLogs);
	      localStorage.setItem(localDatabaseKey, JSON.stringify({
	        savedAt: localIsoDateTime(),
	        standardSlaDataVersion,
	        departmentDataVersion,
	        actionDataVersion,
	        contracts,
	        logRecords,
	        masterData
	      }));
      isApplyingDriveDatabaseLoad = false;
      if (migratedSlaCount || migratedDepartmentCount || migratedActionCount) scheduleDriveDatabaseSave();
      refreshDashboardDataFromContracts();
      renderMasterData();
      renderAll();
      updateDatabaseSyncStatus(`${contracts.length} contracts loaded from Shared Drive${migratedCount ? ` · ${migratedCount} Contract ID migrated` : ""}${migratedSlaCount ? ` · ${migratedSlaCount} SLA values updated` : ""}${migratedDepartmentCount ? ` · Department Master cleaned` : ""}${migratedActionCount ? ` · Action SLA updated` : ""}`);
      return true;
    }

    function loadDriveDatabaseFromCloud() {
      const endpoint = driveDatabaseEndpoint();
      if (!endpoint) return;
      const callbackName = `t23DriveDb_${Date.now()}_${Math.random().toString(36).slice(2)}`;
      const params = new URLSearchParams({
        mode: "loadDriveDatabase",
        folderId: driveDatabaseConfig.folderId,
        contractsCsv: driveDatabaseConfig.contractsCsv,
        logsCsv: driveDatabaseConfig.logsCsv,
        typeMasterCsv: driveDatabaseConfig.typeMasterCsv,
        departmentMasterCsv: driveDatabaseConfig.departmentMasterCsv,
	        peopleMasterCsv: driveDatabaseConfig.peopleMasterCsv,
	        contractTemplateCsv: driveDatabaseConfig.contractTemplateCsv,
	        actionSlaCsv: driveDatabaseConfig.actionSlaCsv,
	        callback: callbackName
	      });
      const script = document.createElement("script");
      let done = false;
      window[callbackName] = payload => {
        done = true;
        try {
          if (!applyDriveDatabasePayload(payload)) updateDatabaseSyncStatus("Shared Drive database is empty");
        } catch (error) {
          updateDatabaseSyncStatus("Shared Drive database could not load");
        } finally {
          delete window[callbackName];
          script.remove();
        }
      };
      script.onerror = () => {
        if (!done) updateDatabaseSyncStatus("Shared Drive database could not load");
        delete window[callbackName];
        script.remove();
      };
      script.src = `${endpoint}${endpoint.includes("?") ? "&" : "?"}${params.toString()}`;
      document.head.appendChild(script);
    }

    async function saveDriveDatabaseToCloud() {
      const endpoint = driveDatabaseEndpoint();
      if (!endpoint || isApplyingDriveDatabaseLoad) return;
      try {
        updateDatabaseSyncStatus(`${contracts.length} contracts saved locally · syncing Shared Drive`);
        await fetch(endpoint, {
          method: "POST",
          mode: "no-cors",
          redirect: "follow",
          headers: { "Content-Type": "text/plain" },
          body: JSON.stringify(driveDatabaseCsvPayload())
        });
        updateDatabaseSyncStatus(`${contracts.length} contracts saved to Shared Drive`);
      } catch (error) {
        updateDatabaseSyncStatus(`${contracts.length} contracts saved locally · Shared Drive sync failed`);
      }
    }

    function scheduleDriveDatabaseSave() {
      if (isApplyingDriveDatabaseLoad) return;
      clearTimeout(driveDatabaseSaveTimer);
      driveDatabaseSaveTimer = window.setTimeout(saveDriveDatabaseToCloud, 800);
    }

    function exportContractsCsv() {
      downloadTextFile(driveDatabaseConfig.contractsCsv, objectsToCsv(contractDatabaseHeaders, contracts.map(contractDbRow)));
      showToast("Contracts CSV exported");
    }

    function exportLogsCsv() {
      downloadTextFile(driveDatabaseConfig.logsCsv, objectsToCsv(logDatabaseHeaders, logRecords.map(logDbRow)));
      showToast("Log CSV exported");
    }

    function importContractsCsv(file) {
      if (!requireSystemAdministrator()) return;
      if (!file) return;
      const reader = new FileReader();
      reader.addEventListener("load", () => {
        const rows = csvToObjects(reader.result);
        const importedContracts = rows.map(contractFromDbRow).filter(item => item.id && item.name);
        if (!importedContracts.length) {
          showToast("No valid contracts in CSV");
          return;
        }
        migrateContractIdsToDepartmentFormat(importedContracts, logRecords);
        const importedIds = new Set(importedContracts.map(item => item.id));
        logRecords.splice(0, logRecords.length, ...logRecords.filter(row => importedIds.has(row[0])));
        importedContracts.forEach(contract => {
          if (!latestLogFor(contract.id)) {
            addLogRecord({
              contractId: contract.id,
              cycle: contract.cycle || 1,
              station: contract.station,
              from: stationParts(contract.station).from || contract.owner,
              to: stationParts(contract.station).to || "Legal",
              inDate: contract.addCaseDate || todayInputValue(),
              outDate: "",
              sla: contract.totalSla || totalSlaFor(contract.workType),
              delayReason: "",
              action: contract.stage || "Draft Created",
              updatedBy: "CSV Import"
            });
          }
        });
        contracts.splice(0, contracts.length, ...importedContracts);
        refreshDashboardDataFromContracts();
        saveContractsDatabase();
        renderAll();
        showToast(`${importedContracts.length} contracts imported`);
      });
      reader.readAsText(file, "utf-8");
    }

    function masterInput(field, value = "", options = {}) {
      const tag = options.multiline ? "textarea" : "input";
      const classes = options.select ? "select" : "input";
      if (options.select) {
        return `<select class="${classes}" data-master-field="${escapeHtml(field)}">
          ${(options.choices || []).map(choice => `<option value="${escapeHtml(choice)}" ${String(value || "") === choice ? "selected" : ""}>${escapeHtml(choice)}</option>`).join("")}
        </select>`;
      }
      if (tag === "textarea") return `<textarea class="${classes}" data-master-field="${escapeHtml(field)}">${escapeHtml(value)}</textarea>`;
      return `<input class="${classes}" type="${escapeHtml(options.type || "text")}" data-master-field="${escapeHtml(field)}" value="${escapeHtml(value)}">`;
    }

    function masterActiveSelect(field, value = "Yes") {
      return masterInput(field, value || "Yes", { select: true, choices: ["Yes", "No"] });
    }

    function masterReadOnly(value = "") {
      return `<span class="master-readonly">${escapeHtml(value === "" || value == null ? "-" : value)}</span>`;
    }

    function masterDeleteButton() {
      return `<button class="icon-button" type="button" data-remove-master-row title="Remove row">×</button>`;
    }

    function stationOwnerForMasterContract(contract) {
      return stationParts(contract.station).to || contract.stationOwner || contract.owner || "";
    }

    function renderMasterData() {
      const contractBody = document.querySelector("#masterContractRows");
      if (contractBody) {
        contractBody.innerHTML = contracts.map(contract => `
          <tr>
            <td><input type="hidden" data-master-field="_originalId" value="${escapeHtml(contract.id)}">${masterInput("id", contract.id)}</td>
            <td>${masterInput("name", contract.name)}</td>
            <td>${masterInput("department", contract.department)}</td>
            <td>${masterInput("owner", contract.owner)}</td>
            <td>${masterInput("type", contractPrimaryTypeDisplay(contract.type))}</td>
            <td>${masterInput("stage", contract.stage)}</td>
            <td>${masterInput("status", contract.status)}</td>
            <td>${masterInput("stationOwner", stationOwnerForMasterContract(contract))}</td>
            <td>${masterInput("addCaseDate", addCaseStartDateForContract(contract), { type: "date" })}</td>
            <td>${masterInput("due", contract.due, { type: "date" })}</td>
            <td>${masterDeleteButton()}</td>
          </tr>`).join("");
      }

      const logBody = document.querySelector("#masterLogRows");
      if (logBody) {
        logBody.innerHTML = logRecords.map(row => `
          <tr>
            <td><input type="hidden" data-master-field="_originalContractId" value="${escapeHtml(row[0])}">${masterReadOnly(row[0])}</td>
            <td><input type="hidden" data-master-field="_originalLogNo" value="${escapeHtml(row[1])}">${masterReadOnly(row[1])}</td>
            <td>${masterReadOnly(row[2])}</td>
            <td>${masterReadOnly(row[4])}</td>
            <td>${masterReadOnly(row[5])}</td>
            <td>${masterInput("inDate", dateToInputValue(row[6]), { type: "date" })}</td>
            <td>${masterInput("outDate", dateToInputValue(row[7]), { type: "date" })}</td>
            <td>${masterReadOnly(row[8])}</td>
            <td>${masterReadOnly(row[9])}</td>
            <td>${masterReadOnly(row[10])}</td>
            <td>${masterReadOnly(row[12])}</td>
          </tr>`).join("");
      }

      const deptBody = document.querySelector("#masterDepartmentRows");
      if (deptBody) {
        deptBody.innerHTML = (masterData.departments || []).map(row => `
          <tr>
            <td>${masterInput("Department / Restaurant", row["Department / Restaurant"])}</td>
            <td>${masterInput("Department Code", row["Department Code"])}</td>
            <td>${masterActiveSelect("Active", row.Active)}</td>
            <td>${masterDeleteButton()}</td>
          </tr>`).join("");
      }

      const peopleBody = document.querySelector("#masterPeopleRows");
      if (peopleBody) {
        peopleBody.innerHTML = (masterData.people || []).map(row => `
          <tr>
            <td>${masterInput("name", row.name)}</td>
            <td>${masterInput("department", row.department)}</td>
            <td>${masterInput("email", row.email)}</td>
            <td>${masterInput("company", row.company)}</td>
            <td>${masterActiveSelect("active", row.active)}</td>
            <td>${masterDeleteButton()}</td>
          </tr>`).join("");
      }

	      const typeBody = document.querySelector("#masterContractTypeRows");
	      if (typeBody) {
	        typeBody.innerHTML = (masterData.contractTypes || []).map(row => `
	          <tr>
	            <td>${masterInput("Contract Classification", row["Contract Classification"] || row.Category)}</td>
            <td>${masterInput("Type of Contract", contractPrimaryTypeDisplay(row["Type of Contract"]))}</td>
            <td>${masterInput("Sub Type of Contract", row["Sub Type of Contract"] || contractSubTypeDisplay(row["Type of Contract"]))}</td>
	            <td>${masterInput("Fixed SLA (Working Days)", row["Fixed SLA (Working Days)"] || standardSlaFromContractTypeMasterV2(row["Sub Type of Contract"] || row["Type of Contract"]) || "")}</td>
            <td>${masterActiveSelect("Active", row.Active)}</td>
	            <td>${masterDeleteButton()}</td>
	          </tr>`).join("");
	      }

	      const actionBody = document.querySelector("#masterActionSlaRows");
	      if (actionBody) {
	        actionBody.innerHTML = (masterData.actionSla || []).map(row => `
	          <tr>
	            <td>${masterInput("Action", row.Action, { select: true, choices: updateActionList })}</td>
	            <td>${masterInput("Description / รายละเอียด", row["Description / รายละเอียด"] || actionDescriptionConfig[row.Action]?.descriptionTh || "")}</td>
	            <td>${masterInput("Fixed SLA (Working Days)", row["Fixed SLA (Working Days)"] || row.sla)}</td>
	            <td>${masterActiveSelect("Active", row.Active)}</td>
	            <td>${masterDeleteButton()}</td>
	          </tr>`).join("");
	      }

		      const templateBody = document.querySelector("#masterTemplateRows");
		      if (templateBody) {
		        templateBody.innerHTML = (masterData.contractTemplates || []).map(row => {
		          const classificationHint = row.classification || row.category || row.accessLevel;
		          const typeInfo = contractTypeMasterV2Match(row.subType || row.type || row.typeGroup, classificationHint);
		          const classification = row.classification || (typeInfo ? classificationDisplayValue(typeInfo["Contract Classification EN"]) : classificationDisplayValue(row.accessLevel === "Confidential" ? "Confidential" : "Day-to-day Work"));
		          const typeGroup = row.typeGroup || (typeInfo ? typeGroupForRow(typeInfo) : contractPrimaryTypeDisplay(row.type));
		          const subType = row.subType || (typeInfo ? subTypeForRow(typeInfo) : contractSubTypeDisplay(row.type));
		          const fixedSla = standardSlaFromContractTypeMasterV2(subType || typeGroup || row.type, classification) || "";
		          const access = row.accessLevel || classificationAccessLevelFor(classification);
		          return `
	          <tr>
	            <td><input type="hidden" data-master-field="selectionLabel" value="${escapeHtml(row.selectionLabel || "")}"><input type="hidden" data-master-field="sourceRow" value="${escapeHtml(row.sourceRow || "")}">${masterInput("classification", classification)}</td>
	            <td>${masterInput("typeGroup", typeGroup)}</td>
	            <td>${masterInput("subType", subType)}</td>
	            <td>${masterInput("name", row.name)}</td>
	            <td>${masterInput("vendor", row.vendor)}</td>
	            <td>${masterInput("department", row.department)}</td>
		            <td>${masterInput("fixedSla", fixedSla)}</td>
	            <td>${masterInput("accessLevel", access, { select: true, choices: ["Normal", "Confidential"] })}</td>
	            <td>${masterActiveSelect("active", row.active)}</td>
	            <td>${masterDeleteButton()}</td>
	          </tr>`;
	        }).join("");
	      }
	    }

    function readMasterRows(selector, fields, requiredField) {
      return [...document.querySelectorAll(`${selector} tr`)].map(row => {
        const item = {};
        fields.forEach(field => {
          const input = [...row.querySelectorAll("[data-master-field]")].find(control => control.dataset.masterField === field);
          item[field] = String(input?.value || "").trim();
        });
        return item;
      }).filter(item => String(item[requiredField] || "").trim());
    }

    function normalizeMasterDataFromUi() {
      if (!requireSystemAdministrator()) return false;
      const originalContracts = new Map(contracts.map(contract => [String(contract.id || "").trim(), contract]));
      const contractRows = readMasterRows("#masterContractRows", ["_originalId", "id", "name", "department", "owner", "type", "stage", "status", "stationOwner", "addCaseDate", "due"], "_originalId");
      const logRows = readMasterRows("#masterLogRows", ["_originalContractId", "_originalLogNo", "inDate", "outDate"], "_originalContractId");
      const nextContracts = [];
      const keptIds = new Set();
      const idMap = new Map();
      const addCaseDateMap = new Map();
      const logDateEdits = new Map();
      let hasContractError = false;
      let hasLogDateError = false;

      contractRows.forEach(row => {
        const originalId = String(row._originalId || "").trim();
        const id = String(row.id || "").trim();
        const name = String(row.name || "").trim();
        if (!id || !name) {
          hasContractError = true;
          return;
        }
        if (keptIds.has(id)) {
          hasContractError = true;
          return;
        }
        keptIds.add(id);
        const original = originalContracts.get(originalId) || {};
        const type = contractPrimaryTypeDisplay(row.type || original.type || "Other") || "Other";
        const owner = String(row.owner || original.owner || "").trim();
        const stationOwner = String(row.stationOwner || stationOwnerForMasterContract(original) || "Legal").trim();
        const totalSla = Number(original.totalSla || totalSlaFor(type)) || 0;
        const used = Number(original.used || original.days || 0) || 0;
        const accessLevel = String(original.accessLevel || accessLevelForContractType(type) || "Normal").trim();
        const addCaseDate = dateToInputValue(row.addCaseDate || original.addCaseDate || firstLogFor(originalId)?.[6] || firstLogFor(id)?.[6] || "");
        addCaseDateMap.set(id, addCaseDate);
        nextContracts.push({
          ...original,
          id,
          name,
          department: String(row.department || original.department || "").trim(),
          owner,
          type,
          vendor: String(original.vendor || "").trim(),
          stage: String(row.stage || original.stage || "Draft Created").trim(),
          status: String(row.status || original.status || original.alert || "Green >>G=On Track").trim(),
          station: `From ${owner || stationParts(original.station || "").from || "Owner"} >> To ${stationOwner}`,
          addCaseDate,
          due: String(row.due || original.due || "").trim(),
          workType: String(original.workType || type || "Other").trim(),
          totalSla,
          used,
          days: Number(original.days || used) || 0,
          balance: Number(original.balance || (totalSla - used)) || 0,
          alert: String(row.status || original.alert || original.status || "Green >>G=On Track").trim(),
          accessLevel,
          visibility: String(original.visibility || (accessLevel === "Confidential" ? "Restricted access / จำกัดสิทธิ์" : "Standard access / สิทธิ์ทั่วไป")).trim(),
          category: String(original.category || contractTypeCategoryFor(type)).trim()
        });
        if (originalId && originalId !== id) idMap.set(originalId, id);
      });

      logRows.forEach(row => {
        const originalContractId = String(row._originalContractId || "").trim();
        const contractId = idMap.get(originalContractId) || originalContractId;
        const logNo = String(row._originalLogNo || "").trim();
        if (!contractId || !logNo) return;
        const originalLog = logRecords.find(item => String(item?.[0] || "") === originalContractId && String(item?.[1] || "") === logNo) || [];
        const nextIn = dateToInputValue(row.inDate || "");
        const nextOut = dateToInputValue(row.outDate || "");
        if (!nextIn || (nextOut && normalizeDateOnly(nextOut) < normalizeDateOnly(nextIn))) {
          hasLogDateError = true;
          return;
        }
        const originalIn = dateToInputValue(originalLog[6] || "");
        const originalOut = dateToInputValue(originalLog[7] || "");
        if (nextIn !== originalIn || nextOut !== originalOut) {
          logDateEdits.set(`${contractId}::${logNo}`, { inDate: nextIn, outDate: nextOut });
        }
      });

      if (hasContractError) {
        showToast("Contract ID and Contract Name are required, and Contract ID must not duplicate");
        return false;
      }
      if (hasLogDateError) {
        showToast("Log Records require In date, and Out date must not be before In date");
        return false;
      }

      contracts.splice(0, contracts.length, ...nextContracts);
      logRecords.splice(0, logRecords.length, ...logRecords
        .map(row => {
          if (Array.isArray(row) && idMap.has(row[0])) row[0] = idMap.get(row[0]);
          return row;
        })
        .filter(row => keptIds.has(row?.[0])));
      logRecords.forEach(row => {
        const addCaseDate = addCaseDateMap.get(row?.[0]);
        if (addCaseDate && Number(row?.[1] || 0) === 1) row[6] = addCaseDate;
        const edit = logDateEdits.get(`${row?.[0]}::${row?.[1]}`);
        if (edit) {
          row[6] = edit.inDate || row[6];
          row[7] = edit.outDate || "";
        }
        recalculateLogTiming(row);
      });
      contracts.forEach(syncContractTimingFromLogs);
      refreshDashboardDataFromContracts();

	      masterData.departments = readMasterRows("#masterDepartmentRows", ["Department / Restaurant", "Department Code", "Active"], "Department / Restaurant")
	        .map(row => {
	          const department = canonicalDepartmentName(row["Department / Restaurant"]);
	          const code = String(row["Department Code"] || departmentCodeSuggestion(department) || "").trim().toUpperCase();
	          return { ...row, "Department / Restaurant": department, "Department Code": code, "Department Data Version": departmentDataVersion, Active: row.Active || "Yes" };
	        })
	        .filter(row => row["Department / Restaurant"] && row["Department / Restaurant"] !== "0" && row["Department Code"] && row["Department Code"] !== "0");
      masterData.people = readMasterRows("#masterPeopleRows", ["company", "department", "name", "email", "active"], "name")
        .map(row => ({ ...row, email: String(row.email || "").trim().toLowerCase(), active: row.active || "Yes" }));
	      masterData.contractTypes = readMasterRows("#masterContractTypeRows", ["Contract Classification", "Type of Contract", "Sub Type of Contract", "Fixed SLA (Working Days)", "Active"], "Contract Classification")
	        .filter(row => String(row["Type of Contract"] || row["Sub Type of Contract"] || "").trim())
	        .map(row => ({
	          ...row,
	          Category: row["Contract Classification"] || "",
	          "Description / คำอธิบาย": row["Type of Contract"] || "",
	          "Fixed SLA (Working Days)": String(row["Fixed SLA (Working Days)"] || "").trim(),
	          "Standard SLA Version": standardSlaDataVersion,
	          Active: row.Active || "Yes"
	        }));
	      masterData.actionSla = readMasterRows("#masterActionSlaRows", ["Action", "Description / รายละเอียด", "Fixed SLA (Working Days)", "Active"], "Action")
	        .filter(row => updateActionList.includes(row.Action))
	        .map(row => ({
	          ...row,
	          "Description / รายละเอียด": String(row["Description / รายละเอียด"] || actionDescriptionConfig[row.Action]?.descriptionTh || "").trim(),
	          "Fixed SLA (Working Days)": String(row["Fixed SLA (Working Days)"] || "").trim(),
	          "SLA Rule / วิธีนับ": actionDescriptionConfig[row.Action]?.slaRuleTh || "",
	          "Action Data Version": actionDataVersion,
	          Active: row.Active || "Yes"
	        }));
		      masterData.contractTemplates = readMasterRows("#masterTemplateRows", ["selectionLabel", "sourceRow", "classification", "typeGroup", "subType", "name", "vendor", "department", "fixedSla", "accessLevel", "active"], "name")
		        .map(row => {
		          const classification = row.classification || classificationDisplayValue(row.accessLevel === "Confidential" ? "Confidential" : "Day-to-day Work");
		          const type = String(row.subType || row.typeGroup || "").trim();
		          const fixedSla = standardSlaFromContractTypeMasterV2(type, classification) || "";
		          return {
		            ...row,
		            type,
		            workType: type || "Other",
		            fixedSla: String(fixedSla || "").trim(),
		            slaVersion: standardSlaDataVersion,
	            category: classification,
	            group: row.typeGroup || "",
	            contractId: "",
	            accessLevel: row.accessLevel || classificationAccessLevelFor(classification),
	            active: row.active || "Yes"
	          };
	        });
	      return true;
	    }

    function addMasterRow(kind) {
      if (!normalizeMasterDataFromUi()) return;
      if (kind === "departments") masterData.departments.push({ "Department / Restaurant": "", "Department Code": "", "Department Data Version": departmentDataVersion, Active: "Yes" });
	      if (kind === "people") masterData.people.push({ company: "Turtle 23", department: "", name: "", email: "", active: "Yes" });
	      if (kind === "contractTypes") masterData.contractTypes.push({ "Contract Classification": "Day-to-day Work / งานดำเนินงานทั่วไป", Category: "Day-to-day Work / งานดำเนินงานทั่วไป", "Type of Contract": "", "Sub Type of Contract": "", "Fixed SLA (Working Days)": "", "Standard SLA Version": standardSlaDataVersion, "Description / คำอธิบาย": "", Active: "Yes" });
	      if (kind === "actionSla") masterData.actionSla.push({ Action: "Submit to Review", "Description / รายละเอียด": actionDescriptionConfig["Submit to Review"].descriptionTh, "Fixed SLA (Working Days)": "", "SLA Rule / วิธีนับ": actionDescriptionConfig["Submit to Review"].slaRuleTh, "Action Data Version": actionDataVersion, Active: "Yes" });
	      if (kind === "contractTemplates") masterData.contractTemplates.push({ classification: "Day-to-day Work / งานดำเนินงานทั่วไป", typeGroup: "", subType: "", name: "", selectionLabel: "", sourceRow: "", type: "", workType: "", contractId: "", accessLevel: "Normal", category: "Day-to-day Work / งานดำเนินงานทั่วไป", department: "", vendor: "", group: "", fixedSla: "", slaVersion: standardSlaDataVersion, active: "Yes" });
      renderMasterData();
    }

    async function saveMasterDataFromUi() {
      if (!normalizeMasterDataFromUi()) return;
      renderMasterData();
      saveContractsDatabase();
      await saveDriveDatabaseToCloud();
      populateUserControls();
      renderUserCasePreview();
      showToast("Master Data saved to Shared Drive");
    }

    function setupMasterDataControls() {
      document.querySelector("#saveMasterDataBtn")?.addEventListener("click", saveMasterDataFromUi);
      document.querySelectorAll("[data-add-master-row]").forEach(button => {
        button.addEventListener("click", () => addMasterRow(button.dataset.addMasterRow));
      });
      document.querySelector("#master")?.addEventListener("click", event => {
        const removeButton = event.target.closest("[data-remove-master-row]");
        if (!removeButton) return;
        removeButton.closest("tr")?.remove();
      });
      renderMasterData();
    }

    function setupCsvDatabaseControls() {
      document.querySelector("#openDriveDatabase")?.setAttribute("href", driveDatabaseConfig.folderUrl);
      document.querySelector("#exportContractsCsv")?.addEventListener("click", exportContractsCsv);
      document.querySelector("#exportLogsCsv")?.addEventListener("click", exportLogsCsv);
      document.querySelector("#importContractsCsv")?.addEventListener("change", event => {
        importContractsCsv(event.target.files?.[0]);
        event.target.value = "";
      });
      document.querySelector("#resetLocalDatabase")?.addEventListener("click", () => {
        localStorage.removeItem(localDatabaseKey);
        window.location.reload();
      });
      updateDatabaseSyncStatus();
    }

    function renderAll() {
      rebuildNotificationQueue();""",
    )
    html = html.replace(
        """      renderAll();
      setView("user");""",
        """      saveContractsDatabase();
      renderAll();
      setView("user");""",
    )
    html = html.replace(
        """      get realWorkbookData() { return realWorkbookData; }""",
        """      get realWorkbookData() { return realWorkbookData; },
      get driveDatabaseConfig() { return driveDatabaseConfig; },
      get attachmentCloudConfig() { return attachmentCloudConfig; }""",
    )
    html = html.replace(
        """    applyRoleAccess();
    renderAll();""",
        """    applyRoleAccess();
    setupCsvDatabaseControls();
    loadContractsDatabase();
    renderAll();
    loadDriveDatabaseFromCloud();""",
    )
    html = html.replace("Mockup Data · จำนวน Delayed / Overdue แตกต่างกันตาม Contract Owner และ Department", "Excel Data · ข้อมูล Contract ID / SLA จาก workbook จริง")
    html = html.replace("applyDiverseMockupData();\n    normalizeExistingData();", "    // Real workbook data is embedded above. Keep the mock generator available for reference, but do not run it.\n    normalizeExistingData();")
    html = html.replace("    // Fixed SLA values imported\n    });\n\n    // Fixed SLA values imported from", "    // Fixed SLA values imported from")
    html = html.replace("    // Contract examples imported\n    });\n\n    // Contract examples imported from", "    // Contract examples imported from")
    html = html.replace("});\n    const requestedRole\n    ];\n    const requestedRole =", "});\n    const requestedRole =")
    html = html.replace("    const requestedRole\n    ];\n", "")
    html = html.replace("get directorySummary() { return { people: employeeDirectory.length, departments: directoryDepartments().length }; }", "get directorySummary() { return { people: employeeDirectory.length, departments: directoryDepartments().length }; },\n      get realWorkbookData() { return realWorkbookData; }")
    html = html.replace(
        "      contractInputCatalog\n        .forEach(item => {",
        "      [...contractInputCatalog, ...activeMasterContractTemplates()]\n        .forEach(item => {",
    )
    html = html.replace(
        """    function directoryDepartments() {
      return orderedUniqueList(employeeDirectory.map(item => item.department).filter(Boolean));
    }""",
        """    function directoryDepartments() {
      return orderedUniqueList([
        ...activeMasterDepartments().map(item => item["Department / Restaurant"]),
        ...allPeopleDirectory().map(item => item.department),
        ...contracts.map(item => item.department)
      ].filter(Boolean));
    }""",
    )
    html = html.replace(
        "    function directoryPeopleOptions(items = employeeDirectory) {",
        "    function directoryPeopleOptions(items = allPeopleDirectory()) {",
    )
    html = html.replace(
        "      return employeeDirectory.find(item => normalizeDirectoryValue(item.name) === needle || normalizeDirectoryValue(item.email) === needle)",
        "      return allPeopleDirectory().find(item => normalizeDirectoryValue(item.name) === needle || normalizeDirectoryValue(item.email) === needle)",
    )
    html = html.replace(
        "      return employeeDirectory.filter(item => normalizeDirectoryValue(item.department) === needle);",
        "      return allPeopleDirectory().filter(item => normalizeDirectoryValue(item.department) === needle);",
    )
    html = html.replace(
        "    function directoryEmployeeOptions(items = employeeDirectory) {",
        "    function directoryEmployeeOptions(items = allPeopleDirectory()) {",
    )
    html = html.replace(
        "      return directoryEmployeeOptions(matches.length ? matches : employeeDirectory);",
        "      return directoryEmployeeOptions(matches.length ? matches : allPeopleDirectory());",
    )
    html = html.replace(
        "employeeDirectory.find(item => normalizeDirectoryValue(item.email)",
        "allPeopleDirectory().find(item => normalizeDirectoryValue(item.email)",
    )
    html = html.replace(
        """      user: ["User Case Action", "เพิ่มเคส อัปเดทสถานะ และปิดเคสจาก Contract Status / Log View"],
      notifications: ["Notification Queue", "NotificationQueueTable"],""",
        """      user: ["User Case Action", "เพิ่มเคส อัปเดทสถานะ และปิดเคสจาก Contract Status / Log View"],
      master: ["Master Data", "แก้ไขข้อมูล dropdown และบันทึกกลับ Shared Drive"],
      confidential: ["Confidential Contract Status", "แสดงเฉพาะสัญญาลับสำหรับผู้มีสิทธิ์เข้าถึง"],
      notifications: ["Notification Queue", "NotificationQueueTable"],""",
    )
    html = html.replace(
        """    function canAccessView(viewName) {
      if (viewName === "user") return isAdmin();
      return true;
    }""",
        """    function hasAccessLevel(minimumLevel) {
      return Number(currentUser?.level || 1) >= Number(minimumLevel || 1);
    }

    function canUseUserCaseAction() {
      return hasAccessLevel(2);
    }

    function canManageMasterData() {
      return hasAccessLevel(4);
    }

    function canAccessView(viewName) {
      if (viewName === "user") return canUseUserCaseAction();
      if (viewName === "confidential") return canViewConfidentialContracts();
      if (viewName === "master") return canManageMasterData();
      return true;
    }

    let contractStatusAccessView = "normal";

    function isConfidentialContract(contract) {
      const access = String(contract?.accessLevel || "").trim().toLowerCase();
      const classification = String(contract?.classification || contract?.category || "").toLowerCase();
      const visibility = String(contract?.visibility || "").toLowerCase();
      const id = String(contract?.id || "").trim().toUpperCase();
      return access === "confidential"
        || classification.includes("confidential")
        || classification.includes("สัญญาลับ")
        || visibility.includes("restricted")
        || id.startsWith("CT-C-");
    }

    function canViewConfidentialContracts() {
      return hasAccessLevel(3);
    }

    function contractsVisibleToCurrentUser() {
      return canViewConfidentialContracts()
        ? contracts.slice()
        : contracts.filter(contract => !isConfidentialContract(contract));
    }

    function contractsForContractStatus() {
      const visible = contractsVisibleToCurrentUser();
      return visible.filter(contract => contractStatusAccessView === "confidential"
        ? isConfidentialContract(contract)
        : !isConfidentialContract(contract));
    }

    function logsVisibleToCurrentUser(rows = logRecords) {
      const visibleIds = new Set(contractsVisibleToCurrentUser().map(contract => contract.id));
      return rows.filter(row => visibleIds.has(row[0]));
    }""",
    )
    html = html.replace(
        """      const userNav = document.querySelector('.nav-button[data-view="user"]');
      const newContractBtn = document.querySelector("#newContractBtn");""",
        """      const userNav = document.querySelector('.nav-button[data-view="user"]');
      const masterNav = document.querySelector('.nav-button[data-view="master"]');
      const confidentialNav = document.querySelector("[data-confidential-nav]");
      const confidentialClassificationChoice = document.querySelector('[data-classification-value="Confidential"]');
      const newContractBtn = document.querySelector("#newContractBtn");""",
    )
    html = html.replace(
        """        if (userNav) userNav.hidden = true;
        if (newContractBtn) newContractBtn.hidden = true;""",
        """        if (userNav) userNav.hidden = !canUseUserCaseAction();
        if (masterNav) masterNav.hidden = !canManageMasterData();
        if (confidentialNav) confidentialNav.hidden = !canViewConfidentialContracts();
        if (confidentialClassificationChoice) confidentialClassificationChoice.hidden = !canViewConfidentialContracts();
        if (!canViewConfidentialContracts()) contractStatusAccessView = "normal";
        if (newContractBtn) newContractBtn.hidden = !canUseUserCaseAction();""",
    )
    html = html.replace(
        """      document.body.dataset.role = currentUser.role;""",
        """      document.body.dataset.role = currentUser.role;
      document.body.dataset.accessLevel = currentUser.accessLevel;
      document.body.dataset.roleName = currentUser.roleNameEn;""",
    )
    html = html.replace(
        """      const contractOptions = contracts.map(item => ({""",
        """      const accessibleContracts = contractsVisibleToCurrentUser();
      const contractOptions = accessibleContracts.map(item => ({""",
        1,
    )
    html = html.replace(
        """      fillSelect("updateContract", contractOptions, lastUserContractId || contracts[0]?.id);
      fillSelect("closeContract", contractOptions, lastUserContractId || contracts[0]?.id);
      fillSelect("adjustDueContract", contractOptions, lastUserContractId || contracts[0]?.id);""",
        """      fillSelect("updateContract", contractOptions, lastUserContractId || accessibleContracts[0]?.id);
      fillSelect("closeContract", contractOptions, lastUserContractId || accessibleContracts[0]?.id);
      fillSelect("adjustDueContract", contractOptions, lastUserContractId || accessibleContracts[0]?.id);""",
    )
    html = html.replace(
        """          const contract = contracts.find(item => item.id === option.value);""",
        """          const contract = accessibleContracts.find(item => item.id === option.value);""",
        1,
    )
    html = html.replace(
        """      if (lastUserContractId && contracts.some(item => item.id === lastUserContractId)) {""",
        """      if (lastUserContractId && accessibleContracts.some(item => item.id === lastUserContractId)) {""",
        1,
    )
    html = html.replace(
        """      const updateContract = contracts.find(item => item.id === updateId);
      const closeContract = contracts.find(item => item.id === closeId);""",
        """      const accessibleContracts = contractsVisibleToCurrentUser();
      const updateContract = accessibleContracts.find(item => item.id === updateId);
      const closeContract = accessibleContracts.find(item => item.id === closeId);""",
    )
    html = html.replace(
        """      const selectedId = lastUserContractId || document.querySelector("#updateContract")?.value || contracts[0]?.id;
      const item = contracts.find(contract => contract.id === selectedId) || contracts[0];""",
        """      const accessibleContracts = contractsVisibleToCurrentUser();
      const selectedId = lastUserContractId || document.querySelector("#updateContract")?.value || accessibleContracts[0]?.id;
      const item = accessibleContracts.find(contract => contract.id === selectedId) || accessibleContracts[0];""",
    )
    html = html.replace(
        """      const rows = contracts.slice(-6).reverse().map(contract => {""",
        """      const rows = accessibleContracts.slice(-6).reverse().map(contract => {""",
        1,
    )
    html = html.replace(
        """      const contract = contracts.find(item => item.id === form.elements.contractId?.value) || contracts[0];""",
        """      const accessibleContracts = contractsVisibleToCurrentUser();
      const contract = accessibleContracts.find(item => item.id === form.elements.contractId?.value) || accessibleContracts[0];""",
        1,
    )
    html = html.replace(
        """        const contract = contracts.find(item => item.id === updateForm.elements.contractId?.value);""",
        """        const contract = contractsVisibleToCurrentUser().find(item => item.id === updateForm.elements.contractId?.value);""",
        1,
    )
    html = html.replace(
        """      const accessLevel = String(form.get("accessLevel") || accessLevelForContractType(contractType) || "Normal").trim();
      const department = String(form.get("department") || "").trim();""",
        """      const accessLevel = String(form.get("accessLevel") || accessLevelForContractType(contractType) || "Normal").trim();
      if (accessLevel === "Confidential" && !canViewConfidentialContracts()) {
        showToast("Level 3 or Level 4 access is required for confidential contracts.");
        return;
      }
      const department = String(form.get("department") || "").trim();""",
        1,
    )
    html = html.replace(
        """      const contract = contracts.find(item => item.id === form.get("contractId"));""",
        """      const contract = contractsVisibleToCurrentUser().find(item => item.id === form.get("contractId"));""",
        1,
    )
    html = html.replace(
        """      const contract = contracts.find(item => item.id === contractId);""",
        """      const contract = contractsVisibleToCurrentUser().find(item => item.id === contractId);""",
        1,
    )
    html = html.replace(
        """      const contract = contracts.find(item => item.id === id);""",
        """      const contract = contractsVisibleToCurrentUser().find(item => item.id === id);""",
        1,
    )
    html = html.replace(
        """        const contract = contracts.find(item => item.id === field("contractId")?.value);""",
        """        const contract = contractsVisibleToCurrentUser().find(item => item.id === field("contractId")?.value);""",
        1,
    )
    html = html.replace(
        """      if (!isAdmin()) {
        showToast("Only admin can create or update contracts.");
        return;
      }""",
        """      if (!canUseUserCaseAction()) {
        showToast("Level 2, Level 3, or Level 4 access is required to create or update contracts.");
        return;
      }""",
        1,
    )
    html = html.replace(
        """    function setView(viewName) {
      if (viewName === "workflow") viewName = "contracts";
      if (!pageLabels[viewName]) viewName = "dashboard";
      if (!canAccessView(viewName)) {
        showToast("You do not have permission to access User Case Action.");
        viewName = "dashboard";
      }
      document.querySelectorAll(".nav-button").forEach(button => {
        button.classList.toggle("active", button.dataset.view === viewName);
      });
      document.querySelectorAll(".view").forEach(view => {
        view.classList.toggle("active", view.id === viewName);
      });
      document.querySelector("#pageHeading").textContent = pageLabels[viewName][0];
      document.querySelector("#pageSubheading").textContent = pageLabels[viewName][1];
    }""",
        """    function setView(viewName) {
      if (viewName === "workflow") viewName = "contracts";
      if (!pageLabels[viewName]) viewName = "dashboard";
      if (!canAccessView(viewName)) {
        showToast("You do not have permission to access this page.");
        viewName = "dashboard";
      }
      if (viewName === "contracts" || viewName === "confidential") {
        contractStatusAccessView = viewName === "confidential" ? "confidential" : "normal";
        selectedLogContractId = "all";
        populateContractMasterFilters();
        renderContractsTable();
        renderKanban();
        renderLogView();
      }
      const targetView = viewName === "confidential" ? "contracts" : viewName;
      document.querySelectorAll(".nav-button").forEach(button => {
        button.classList.toggle("active", button.dataset.view === viewName);
      });
      document.querySelectorAll(".view").forEach(view => {
        view.classList.toggle("active", view.id === targetView);
      });
      document.querySelector("#pageHeading").textContent = pageLabels[viewName][0];
      document.querySelector("#pageSubheading").textContent = pageLabels[viewName][1];
    }""",
    )
    html = html.replace(
        """    function filteredDashboardContracts() {
      const search = document.querySelector("#dashboardSearch").value.toLowerCase();
      const alert = document.querySelector("#dashboardAlert").value;
      return contracts.filter(item => {""",
        """    function filteredDashboardContracts() {
      const search = document.querySelector("#dashboardSearch").value.toLowerCase();
      const alert = document.querySelector("#dashboardAlert").value;
      return contractsVisibleToCurrentUser().filter(item => {""",
    )
    html = html.replace(
        """    function ownerRelatedContracts(owner) {
      return contracts.filter(item => {""",
        """    function ownerRelatedContracts(owner) {
      return contractsVisibleToCurrentUser().filter(item => {""",
    )
    html = html.replace(
        """    function ownerRelatedLogs(owner) {
      return logRecords.filter(row => row[4] === owner || row[5] === owner).slice(0, 8);
    }""",
        """    function ownerRelatedLogs(owner) {
      return logsVisibleToCurrentUser().filter(row => row[4] === owner || row[5] === owner).slice(0, 8);
    }""",
    )
    html = html.replace(
        """    function renderContractDetail() {
      const item = contracts[selectedContractIndex];
      const logs = logItems[item.id] || [""",
        """    function renderContractDetail() {
      const visibleContracts = contractsVisibleToCurrentUser();
      let item = contracts[selectedContractIndex];
      if (!item || !visibleContracts.some(contract => contract.id === item.id)) {
        item = visibleContracts[0];
        selectedContractIndex = item ? contracts.findIndex(contract => contract.id === item.id) : 0;
      }
      if (!item) {
        document.querySelector("#selectedStamp").textContent = "No accessible contract";
        document.querySelector("#contractDetail").innerHTML = `<div class="empty-state">No contract available for your access level</div>`;
        return;
      }
      const logs = logItems[item.id] || [""",
    )
    html = html.replace(
        """        const cards = contracts.filter(item => item.stage === stage).map(item => `""",
        """        const cards = contractsForContractStatus().filter(item => item.stage === stage).map(item => `""",
        1,
    )
    html = html.replace(
        """    function uniqueContractMasterValues(getter) {
      return [...new Set(contracts.map(getter).map(value => String(value || "").trim()).filter(Boolean))]""",
        """    function uniqueContractMasterValues(getter) {
      return [...new Set(contractsForContractStatus().map(getter).map(value => String(value || "").trim()).filter(Boolean))]""",
    )
    html = html.replace(
        """      const visibleContracts = contracts.filter(item => {
        const latestLog = latestLogFor(item.id);""",
        """      const visibleContracts = contractsForContractStatus().filter(item => {
        const latestLog = latestLogFor(item.id);""",
        1,
    )
    html = html.replace(
        """      const currentContract = contracts.find(item => item.id === contractId) || contracts.find(item => item.station === station);
      const logViewRows = currentContract ? logRecords.filter(row => row[0] === currentContract.id).slice().reverse() : [];""",
        """      const visibleContracts = contractsVisibleToCurrentUser();
      const currentContract = visibleContracts.find(item => item.id === contractId) || visibleContracts.find(item => item.station === station);
      if (!currentContract) {
        showToast("You do not have permission to view this confidential contract.");
        return;
      }
      const logViewRows = logsVisibleToCurrentUser().filter(row => row[0] === currentContract.id).slice().reverse();""",
    )
    html = html.replace(
        """    function renderLogView() {
      const select = document.querySelector("#logViewContract");
      const selected = selectedLogContractId || select.value || "all";
      const options = [
        `<option value="all" ${selected === "all" ? "selected" : ""}>All Contract ID</option>`,
        ...contracts.map(item => item.id).map(id => `""",
        """    function renderLogView() {
      const select = document.querySelector("#logViewContract");
      const statusContracts = contractsForContractStatus();
      const statusIds = new Set(statusContracts.map(item => item.id));
      let selected = selectedLogContractId || select.value || "all";
      if (selected !== "all" && !statusIds.has(selected)) selected = "all";
      const options = [
        `<option value="all" ${selected === "all" ? "selected" : ""}>All Contract ID</option>`,
        ...statusContracts.map(item => item.id).map(id => `""",
    )
    html = html.replace(
        """      const rows = selected === "all" ? logRecords : logRecords.filter(row => row[0] === selected);""",
        """      const statusLogs = logsVisibleToCurrentUser().filter(row => statusIds.has(row[0]));
      const rows = selected === "all" ? statusLogs : statusLogs.filter(row => row[0] === selected);""",
        1,
    )
    html = html.replace(
        """    function syncNavCounts() {
      const counts = {
        contracts: contracts.length,
        workflow: logRecords.length,""",
        """    function syncNavCounts() {
      const visibleContracts = contractsVisibleToCurrentUser();
      const visibleIds = new Set(visibleContracts.map(contract => contract.id));
      const counts = {
        contracts: contracts.filter(contract => !isConfidentialContract(contract)).length,
        confidential: canViewConfidentialContracts() ? contracts.filter(isConfidentialContract).length : 0,
        workflow: logRecords.filter(row => visibleIds.has(row[0])).length,""",
    )
    html = html.replace(
        """    function renderAll() {
      rebuildNotificationQueue();
      renderDashboardSummary();
      populateContractMasterFilters();""",
        """    function renderAll() {
      rebuildNotificationQueue();
      renderDashboardSummary();
      populateContractMasterFilters();""",
    )
    html = html.replace(
        """    document.querySelectorAll(".nav-button").forEach(button => {
      button.addEventListener("click", () => {
        if (button.hidden || button.dataset.view === "workflow") return;
        setView(button.dataset.view);
      });
    });""",
        """    document.querySelectorAll(".nav-button").forEach(button => {
      button.addEventListener("click", () => {
        if (button.hidden || button.dataset.view === "workflow") return;
        setView(button.dataset.view);
      });
    });""",
    )
    html = html.replace(
        """      renderUserCasePreview();
      syncNavCounts();""",
        """      renderUserCasePreview();
      renderMasterData();
      syncNavCounts();""",
    )
    html = html.replace(
        """    setupCsvDatabaseControls();
    loadContractsDatabase();""",
        """    setupCsvDatabaseControls();
    setupMasterDataControls();
    loadContractsDatabase();""",
    )
    html = html.replace(
        """            <div class="panel-header">
              <div>
                <h2>Contract Master</h2>
                <small>ติดตาม Contract Owner, cycle, return และสถานะล่าสุด</small>
              </div>
            </div>""",
        """            <div class="panel-header">
              <div>
                <h2>Contract Status</h2>
                <small>ติดตาม Contract Owner, cycle, return และสถานะล่าสุด</small>
              </div>
            </div>""",
        1,
    )
    html = html.replace("<title>Tracking Contracts — User Status & Email</title>", "<title>Tracking Contracts — Real Excel Dropdowns</title>")

    # Dashboard summaries must use the same role-based contract scope as the
    # dashboard table. Level 1/2 therefore never aggregate Confidential cases.
    html = html.replace(
        '''    function renderDashboardSummary() {
      const pending = sumBy("isPending");
      const completed = sumBy("isCompletedLast30Days");
      const avgComplete = average(dashboardData.filter(item => item.isCompletedLast30Days === 1), "totalSpendingDays");
      const overdue = sumBy("isOverdue");''',
        '''    function renderDashboardSummary() {
      const visibleContractIds = new Set(contractsVisibleToCurrentUser().map(item => item.id));
      const visibleDashboardData = dashboardData.filter(item => visibleContractIds.has(item.id));
      const visibleSumBy = field => visibleDashboardData.reduce((total, item) => total + (Number(item[field]) || 0), 0);
      const pending = visibleSumBy("isPending");
      const completed = visibleSumBy("isCompletedLast30Days");
      const avgComplete = average(visibleDashboardData.filter(item => item.isCompletedLast30Days === 1), "totalSpendingDays");
      const overdue = visibleSumBy("isOverdue");''',
        1,
    )
    html = html.replace(
        'const departments = orderedUniqueValues(dashboardData, "department", departmentOrder);',
        'const departments = orderedUniqueValues(visibleDashboardData, "department", departmentOrder);',
        1,
    )
    html = html.replace(
        'const deptBuckets = orderedUniqueValues(dashboardData, "deptBucket", deptBucketOrder);',
        'const deptBuckets = orderedUniqueValues(visibleDashboardData, "deptBucket", deptBucketOrder);',
        1,
    )
    html = html.replace(
        'const contractStatusData = dashboardData.filter(item => item.isPending === 1);',
        'const contractStatusData = visibleDashboardData.filter(item => item.isPending === 1);',
        1,
    )
    html = html.replace(
        'value: average(dashboardData.filter(item => item.department === department), "totalSpendingDays")',
        'value: average(visibleDashboardData.filter(item => item.department === department), "totalSpendingDays")',
        1,
    )
    html = html.replace(
        'const longestPendingOnHand = dashboardData',
        'const longestPendingOnHand = visibleDashboardData',
        1,
    )

    # Login Gateway is injected after all existing dashboard transformations so the
    # original views and business logic remain intact behind the authentication gate.
    auth_css = r'''
    body.auth-locked {
      overflow: hidden;
      background: #f3f6f2;
    }

    body.auth-locked .app-shell,
    body.auth-locked .access-denied-screen {
      display: none !important;
    }

    body.auth-ready .login-gateway {
      display: none;
    }

    .login-gateway {
      position: fixed;
      inset: 0;
      z-index: 10000;
      display: grid;
      place-items: center;
      overflow: auto;
      padding: 28px 18px;
      background:
        linear-gradient(135deg, rgba(184, 214, 61, .12), transparent 44%),
        #f3f6f2;
    }

    .login-gateway-inner {
      width: min(100%, 460px);
      margin: auto;
    }

    .login-card {
      padding: 30px;
      border: 1px solid #dce5d2;
      border-top: 4px solid var(--t23-green);
      border-radius: 8px;
      background: #fff;
      box-shadow: 0 18px 48px rgba(18, 31, 19, .12);
    }

    .login-brand-mark {
      width: 48px;
      height: 48px;
      display: grid;
      place-items: center;
      margin-bottom: 18px;
      border-radius: 8px;
      background: var(--t23-green);
      color: #111;
      font-size: 23px;
      font-weight: 900;
    }

    .login-card h1 {
      margin: 0;
      color: #172119;
      font-size: 27px;
      line-height: 1.2;
      letter-spacing: 0;
    }

    .login-system-name {
      margin: 7px 0 0;
      color: #49614b;
      font-size: 14px;
      font-weight: 750;
    }

    .login-intro {
      margin: 18px 0 22px;
      color: #657067;
      font-size: 13px;
      line-height: 1.5;
    }

    .login-form {
      display: grid;
      gap: 15px;
    }

    .login-field {
      display: grid;
      gap: 6px;
    }

    .login-field label {
      color: #29372b;
      font-size: 13px;
      font-weight: 800;
    }

    .login-input-shell {
      min-height: 44px;
      display: grid;
      grid-template-columns: 38px minmax(0, 1fr) auto;
      align-items: center;
      overflow: hidden;
      border: 1px solid #cbd7c3;
      border-radius: 7px;
      background: #fff;
    }

    .login-input-shell:focus-within {
      border-color: #769313;
      box-shadow: 0 0 0 3px rgba(184, 214, 61, .2);
    }

    .login-input-shell.is-invalid {
      border-color: #c93d35;
      box-shadow: 0 0 0 3px rgba(201, 61, 53, .1);
    }

    .login-input-icon {
      color: #647164;
      text-align: center;
      font-size: 16px;
    }

    .login-input-shell input {
      min-width: 0;
      height: 42px;
      padding: 0 8px 0 0;
      border: 0;
      outline: 0;
      color: #182019;
      background: transparent;
      font: inherit;
      font-size: 14px;
    }

    .password-toggle {
      min-width: 54px;
      height: 34px;
      padding: 0 10px;
      border-left: 1px solid #dde4d8;
      color: #526253;
      background: transparent;
      font-size: 12px;
      font-weight: 800;
    }

    .login-validation {
      min-height: 0;
      color: #b3261e;
      font-size: 11px;
      line-height: 1.35;
    }

    .login-validation:empty {
      display: none;
    }

    .login-error {
      display: none;
      gap: 9px;
      padding: 11px 12px;
      border: 1px solid #efc7c3;
      border-radius: 7px;
      color: #9d251e;
      background: #fff4f3;
      font-size: 12px;
      line-height: 1.45;
    }

    .login-error.show {
      display: grid;
      grid-template-columns: auto 1fr;
    }

    .remember-row {
      display: flex;
      align-items: center;
      gap: 8px;
      color: #4e5d4f;
      font-size: 12px;
      cursor: pointer;
    }

    .remember-row input {
      width: 16px;
      height: 16px;
      accent-color: #688300;
    }

    .login-actions {
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 9px;
    }

    .login-submit,
    .login-clear {
      min-height: 43px;
      border-radius: 7px;
      font-weight: 850;
    }

    .login-submit {
      color: #111;
      background: var(--t23-green);
    }

    .login-submit:hover {
      background: #a8c72e;
    }

    .login-clear {
      padding: 0 18px;
      border: 1px solid #ccd6c7;
      color: #3e4d40;
      background: #fff;
    }

    .login-form[aria-busy="true"] input,
    .login-form[aria-busy="true"] button,
    .login-form[aria-busy="true"] .remember-row {
      pointer-events: none;
      opacity: .72;
    }

    .login-spinner {
      width: 14px;
      height: 14px;
      display: none;
      border: 2px solid rgba(0, 0, 0, .22);
      border-top-color: #111;
      border-radius: 50%;
      animation: login-spin .7s linear infinite;
    }

    .login-form[aria-busy="true"] .login-spinner {
      display: inline-block;
    }

    @keyframes login-spin { to { transform: rotate(360deg); } }

    .demo-account-section {
      margin-top: 22px;
      padding-top: 18px;
      border-top: 1px solid #e4e9e1;
    }

    .demo-account-section h2 {
      margin: 0 0 10px;
      color: #526153;
      font-size: 12px;
      letter-spacing: 0;
      text-transform: uppercase;
    }

    .demo-account-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 7px;
    }

    .demo-account-button {
      min-width: 0;
      min-height: 48px;
      padding: 8px 10px;
      border: 1px solid #d5decf;
      border-radius: 7px;
      color: #324333;
      background: #f8faf6;
      text-align: left;
    }

    .demo-account-button:hover {
      border-color: #9ab32d;
      background: #f2f7db;
    }

    .demo-account-button strong,
    .demo-account-button span {
      display: block;
    }

    .demo-account-button strong { font-size: 12px; }
    .demo-account-button span { margin-top: 2px; color: #6c766d; font-size: 10px; }

    .login-footer {
      margin-top: 16px;
      color: #6c756d;
      text-align: center;
      font-size: 10px;
      line-height: 1.5;
    }

    .profile-menu {
      position: relative;
    }

    .profile-trigger {
      min-height: 44px;
      display: grid;
      grid-template-columns: 32px minmax(0, auto) 16px;
      gap: 8px;
      align-items: center;
      padding: 5px 9px 5px 6px;
      border: 1px solid #d5ddd1;
      border-radius: 7px;
      color: #243126;
      background: #fff;
      text-align: left;
    }

    .profile-avatar {
      width: 32px;
      height: 32px;
      display: grid;
      place-items: center;
      border-radius: 50%;
      color: #111;
      background: var(--t23-green);
      font-size: 13px;
      font-weight: 900;
    }

    .profile-copy strong,
    .profile-copy small {
      display: block;
      max-width: 190px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    .profile-copy strong { font-size: 12px; }
    .profile-copy small { margin-top: 1px; color: #68736a; font-size: 9px; }

    .profile-dropdown {
      position: absolute;
      top: calc(100% + 7px);
      right: 0;
      z-index: 30;
      width: 205px;
      display: none;
      padding: 6px;
      border: 1px solid #d8dfd4;
      border-radius: 7px;
      background: #fff;
      box-shadow: 0 14px 32px rgba(18, 31, 19, .14);
    }

    .profile-menu.open .profile-dropdown { display: grid; }

    .profile-dropdown button {
      min-height: 36px;
      padding: 8px 10px;
      border-radius: 5px;
      color: #354337;
      background: transparent;
      text-align: left;
      font-size: 12px;
    }

    .profile-dropdown button:hover:not(:disabled) { background: #f0f5e5; }
    .profile-dropdown button:disabled { color: #9aa19b; cursor: not-allowed; }
    .profile-dropdown .logout-button { color: #a32c24; border-top: 1px solid #e5e9e2; border-radius: 0 0 5px 5px; }

    .access-denied-screen {
      position: fixed;
      inset: 0 0 0 256px;
      z-index: 20;
      display: none;
      place-items: center;
      padding: 28px;
      background: #f4f7f3;
    }

    .access-denied-screen.show { display: grid; }

    .access-denied-card {
      width: min(100%, 440px);
      padding: 30px;
      border: 1px solid #d9e1d5;
      border-top: 4px solid #d97828;
      border-radius: 8px;
      background: #fff;
      box-shadow: 0 16px 40px rgba(18, 31, 19, .1);
      text-align: center;
    }

    .access-denied-icon {
      width: 46px;
      height: 46px;
      display: grid;
      place-items: center;
      margin: 0 auto 14px;
      border-radius: 50%;
      color: #9b3f12;
      background: #fff0e6;
      font-size: 22px;
      font-weight: 900;
    }

    .access-denied-card h2 { margin: 0; font-size: 24px; letter-spacing: 0; }
    .access-denied-card p { margin: 10px 0 0; color: #667067; font-size: 13px; line-height: 1.55; }
    .access-denied-card button { margin-top: 20px; min-height: 41px; padding: 0 18px; border-radius: 7px; color: #111; background: var(--t23-green); font-weight: 850; }

    @media (max-width: 1120px) {
      .access-denied-screen { left: 78px; }
      .profile-copy { display: none; }
      .profile-trigger { grid-template-columns: 32px 16px; }
    }

    @media (max-width: 720px) {
      .login-gateway { padding: 16px 12px; place-items: start center; }
      .login-card { padding: 22px 18px; }
      .login-card h1 { font-size: 24px; }
      .login-actions { grid-template-columns: 1fr; }
      .login-clear { width: 100%; }
      .demo-account-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .access-denied-screen { left: 0; bottom: 67px; padding: 16px; }
      .access-denied-card { padding: 24px 18px; }
      .profile-menu { margin-left: auto; }
    }

    @media (max-width: 390px) {
      .demo-account-grid { grid-template-columns: 1fr; }
    }
'''

    login_html = r'''
  <section class="login-gateway" id="loginGateway" aria-labelledby="loginTitle">
    <div class="login-gateway-inner">
      <div class="login-card">
        <div class="login-brand-mark" aria-hidden="true">▣</div>
        <h1 id="loginTitle">Welcome Back</h1>
        <p class="login-system-name">Contract Tracking System</p>
        <p class="login-intro">Sign in to access the Contract Tracking System<br>Track, Manage and Monitor Contract Progress</p>
        <form class="login-form" id="loginForm" novalidate aria-busy="false">
          <div class="login-error" id="loginError" role="alert" aria-live="polite"><strong aria-hidden="true">!</strong><span id="loginErrorText"></span></div>
          <div class="login-field">
            <label for="loginUsername">Username</label>
            <div class="login-input-shell" id="loginUsernameShell">
              <span class="login-input-icon" aria-hidden="true">●</span>
              <input id="loginUsername" name="username" type="text" placeholder="Enter your username" autocomplete="username" autocapitalize="none" spellcheck="false">
              <span></span>
            </div>
            <div class="login-validation" id="loginUsernameError"></div>
          </div>
          <div class="login-field">
            <label for="loginPassword">Password</label>
            <div class="login-input-shell" id="loginPasswordShell">
              <span class="login-input-icon" aria-hidden="true">■</span>
              <input id="loginPassword" name="password" type="password" placeholder="Enter your password" autocomplete="current-password">
              <button class="password-toggle" id="passwordToggle" type="button" aria-label="Show password">Show</button>
            </div>
            <div class="login-validation" id="loginPasswordError"></div>
          </div>
          <label class="remember-row"><input id="rememberMe" type="checkbox"> <span>Remember Me</span></label>
          <div class="login-actions">
            <button class="login-submit" id="loginSubmit" type="submit"><span class="login-spinner" aria-hidden="true"></span><span id="loginSubmitText">Sign In</span></button>
            <button class="login-clear" id="loginClear" type="button">Clear</button>
          </div>
        </form>
        <section class="demo-account-section" id="demoAccountSection" aria-labelledby="demoAccountsTitle">
          <h2 id="demoAccountsTitle">Accounts</h2>
          <div class="demo-account-grid">
            <button class="demo-account-button" type="button" data-demo-account="viewer"><strong>Viewer</strong><span>Contract Viewer</span></button>
            <button class="demo-account-button" type="button" data-demo-account="user"><strong>User</strong><span>Contract User</span></button>
            <button class="demo-account-button" type="button" data-demo-account="confidential"><strong>Confidential</strong><span>Confidential User</span></button>
            <button class="demo-account-button" type="button" data-demo-account="admin"><strong>Admin</strong><span>System Administrator</span></button>
          </div>
        </section>
      </div>
      <footer class="login-footer">Contract Tracking System · Version 1.0.0<br>© 2026 All rights reserved.</footer>
    </div>
  </section>
'''

    profile_html = r'''
          <div class="profile-menu" id="profileMenu">
            <button class="profile-trigger" id="profileTrigger" type="button" aria-haspopup="menu" aria-expanded="false">
              <span class="profile-avatar" id="profileAvatar">U</span>
              <span class="profile-copy"><strong id="profileUsername">user</strong><small id="profileRole">Contract User</small></span>
              <span aria-hidden="true">▾</span>
            </button>
            <div class="profile-dropdown" role="menu" aria-label="User profile">
              <button type="button" role="menuitem" disabled>My Profile</button>
              <button type="button" role="menuitem" disabled>Change Password</button>
              <button class="logout-button" id="logoutBtn" type="button" role="menuitem">Logout</button>
            </div>
          </div>
'''

    denied_html = r'''
  <section class="access-denied-screen" id="accessDeniedScreen" aria-labelledby="accessDeniedTitle">
    <div class="access-denied-card">
      <div class="access-denied-icon" aria-hidden="true">!</div>
      <h2 id="accessDeniedTitle">Access Denied</h2>
      <p>You do not have permission to access this page.<br>คุณไม่มีสิทธิ์เข้าใช้งานหน้านี้</p>
      <button id="backToDashboardBtn" type="button">Back to Dashboard</button>
    </div>
  </section>
'''

    html = html.replace("  </style>", auth_css + "\n  </style>", 1)
    html = html.replace("<body>", '<body class="auth-locked">\n' + login_html, 1)
    html = html.replace(
        '          <button class="icon-button" id="refreshBtn" title="Refresh">↻</button>',
        '          <button class="icon-button" id="refreshBtn" title="Refresh">↻</button>\n' + profile_html,
        1,
    )
    toast_anchor = '  <div class="toast" id="toast" role="status" aria-live="polite"></div>'
    if toast_anchor not in html:
        raise RuntimeError("Could not find toast anchor for Access Denied page")
    html = html.replace(toast_anchor, denied_html + "\n" + toast_anchor, 1)

    old_role_block = r'''    const requestedRole = new URLSearchParams(window.location.search).get("role");
    // Prototype-only auth state. Real deployments must enforce admin permissions on the backend.
    const accessRoleConfig = Object.freeze({
      viewer: { level: 1, code: "level1", nameEn: "Contract Viewer", nameTh: "ผู้ดูข้อมูลสัญญา", userName: "Viewer" },
      user: { level: 2, code: "level2", nameEn: "Contract User", nameTh: "ผู้ดำเนินการสัญญา", userName: "Contract User" },
      confidential: { level: 3, code: "level3", nameEn: "Confidential Contract User", nameTh: "ผู้ดำเนินการสัญญาลับ", userName: "Confidential User" },
      admin: { level: 4, code: "level4", nameEn: "System Administrator", nameTh: "ผู้ดูแลระบบ", userName: "Kira" }
    });
    const requestedRoleKey = ({
      level1: "viewer",
      "contract-viewer": "viewer",
      level2: "user",
      "contract-user": "user",
      level3: "confidential",
      "confidential-user": "confidential",
      level4: "admin",
      "system-administrator": "admin"
    })[String(requestedRole || "").trim().toLowerCase()] || String(requestedRole || "").trim().toLowerCase();
    const activeRoleKey = accessRoleConfig[requestedRoleKey] ? requestedRoleKey : "viewer";
    const activeRole = accessRoleConfig[activeRoleKey];
    const currentUser = {
      name: activeRole.userName,
      role: activeRoleKey,
      level: activeRole.level,
      accessLevel: activeRole.code,
      roleNameEn: activeRole.nameEn,
      roleNameTh: activeRole.nameTh
    };'''

    auth_config_js = r'''    // Demo authentication gateway. Production should replace this client-side adapter
    // with Microsoft Entra ID or another server-validated identity provider.
    const authConfig = Object.freeze({
      showDemoAccounts: true,
      sessionTimeoutMs: 30 * 60 * 1000,
      storageKey: "t23-contract-auth-session-v1"
    });
    const accessRoleConfig = Object.freeze({
      viewer: { level: 1, code: "level1", nameEn: "Contract Viewer", nameTh: "ผู้ดูข้อมูลสัญญา" },
      user: { level: 2, code: "level2", nameEn: "Contract User", nameTh: "ผู้ดำเนินการสัญญา" },
      confidential: { level: 3, code: "level3", nameEn: "Confidential Contract User", nameTh: "ผู้ดำเนินการสัญญาลับ" },
      admin: { level: 4, code: "level4", nameEn: "System Administrator", nameTh: "ผู้ดูแลระบบ" }
    });
    const demoAccountConfig = Object.freeze({
      viewer: { role: "viewer", passwordHash: "9f9f20ec6958d3c176c17ea3eb9731b053b6fc753367e3c25e3fb2578ec0f4e6" },
      user: { role: "user", passwordHash: "cc66576333f36fb04050e842039951ddbdb54b5dcca031a5eb5861fa1e1a1f09" },
      confidential: { role: "confidential", passwordHash: "7e7f06d4dccf2376e8774a707806817c9fe7079979ddc6e74c088d7a57afa08b" },
      admin: { role: "admin", passwordHash: "fb07d63fb5ae1ee60f5f96592cb25a99ed421f4e32969d347af615bf62c5e5b3" }
    });
    let currentUser = null;
    let authenticatedApplicationInitialized = false;
    let sessionActivityTimer = 0;
    let sessionExpiryTimer = 0;

    async function sha256(value) {
      if (!window.crypto?.subtle) return "";
      const bytes = new TextEncoder().encode(String(value || ""));
      const digest = await window.crypto.subtle.digest("SHA-256", bytes);
      return Array.from(new Uint8Array(digest)).map(byte => byte.toString(16).padStart(2, "0")).join("");
    }

    function buildCurrentUser(username, roleKey) {
      const role = accessRoleConfig[roleKey];
      if (!role) return null;
      return {
        name: username,
        username,
        role: roleKey,
        level: role.level,
        accessLevel: role.code,
        roleNameEn: role.nameEn,
        roleNameTh: role.nameTh
      };
    }

    function readStoredSession() {
      const candidates = [localStorage, sessionStorage];
      for (const storage of candidates) {
        try {
          const raw = storage.getItem(authConfig.storageKey);
          if (!raw) continue;
          const session = JSON.parse(raw);
          const account = demoAccountConfig[session.username];
          const valid = account
            && account.role === session.role
            && Number(session.expiresAt || 0) > Date.now()
            && Number(session.lastActivity || 0) <= Date.now();
          if (valid) return { ...session, storage };
          storage.removeItem(authConfig.storageKey);
        } catch (error) {
          storage.removeItem(authConfig.storageKey);
        }
      }
      return null;
    }

    function persistSession(session, remember) {
      localStorage.removeItem(authConfig.storageKey);
      sessionStorage.removeItem(authConfig.storageKey);
      const storage = remember ? localStorage : sessionStorage;
      storage.setItem(authConfig.storageKey, JSON.stringify(session));
      return storage;
    }

    function clearStoredSession() {
      localStorage.removeItem(authConfig.storageKey);
      sessionStorage.removeItem(authConfig.storageKey);
    }

    function setLoginError(message) {
      const box = document.querySelector("#loginError");
      const text = document.querySelector("#loginErrorText");
      if (text) text.innerHTML = message || "";
      box?.classList.toggle("show", Boolean(message));
    }

    function setFieldValidation(fieldName, message) {
      document.querySelector(`#login${fieldName}Error`).innerHTML = message || "";
      document.querySelector(`#login${fieldName}Shell`)?.classList.toggle("is-invalid", Boolean(message));
    }

    function setLoginLoading(loading) {
      const form = document.querySelector("#loginForm");
      form?.setAttribute("aria-busy", String(loading));
      ["#loginUsername", "#loginPassword", "#rememberMe", "#loginSubmit", "#loginClear"].forEach(selector => {
        const control = document.querySelector(selector);
        if (control) control.disabled = loading;
      });
      const label = document.querySelector("#loginSubmitText");
      if (label) label.textContent = loading ? "Signing in..." : "Sign In";
    }

    function showLoginGateway(message = "") {
      currentUser = null;
      document.body.classList.add("auth-locked");
      document.body.classList.remove("auth-ready");
      document.querySelector("#accessDeniedScreen")?.classList.remove("show");
      document.querySelector("#profileMenu")?.classList.remove("open");
      setLoginError(message);
      window.setTimeout(() => document.querySelector("#loginUsername")?.focus(), 0);
    }

    function updateProfileDisplay() {
      if (!currentUser) return;
      document.querySelector("#profileAvatar").textContent = currentUser.username.charAt(0).toUpperCase();
      document.querySelector("#profileUsername").textContent = currentUser.username;
      document.querySelector("#profileRole").textContent = currentUser.roleNameEn;
    }

    function currentRouteView() {
      const hashRoute = window.location.hash.replace(/^#\/?/, "").split(/[?&]/)[0];
      const pathRoute = window.location.pathname.split("/").filter(Boolean).pop() || "";
      const routes = {
        dashboard: "dashboard",
        "contract-status": "contracts",
        "user-case-action": "user",
        "confidential-contract-status": "confidential",
        "master-data": "master"
      };
      return routes[hashRoute] || routes[pathRoute] || "dashboard";
    }

    function routeForView(viewName) {
      return ({
        dashboard: "dashboard",
        contracts: "contract-status",
        user: "user-case-action",
        confidential: "confidential-contract-status",
        master: "master-data"
      })[viewName] || "dashboard";
    }

    function updateSessionActivity(force = false) {
      if (!currentUser) return;
      const now = Date.now();
      if (!force && now - sessionActivityTimer < 30000) return;
      sessionActivityTimer = now;
      const stored = readStoredSession();
      if (!stored || stored.username !== currentUser.username || stored.role !== currentUser.role) {
        expireSession();
        return;
      }
      const refreshed = { ...stored, lastActivity: now, expiresAt: now + authConfig.sessionTimeoutMs };
      delete refreshed.storage;
      stored.storage.setItem(authConfig.storageKey, JSON.stringify(refreshed));
    }

    function scheduleSessionExpiry() {
      window.clearInterval(sessionExpiryTimer);
      sessionExpiryTimer = window.setInterval(() => {
        if (currentUser && !readStoredSession()) expireSession();
      }, 30000);
    }

    function expireSession() {
      clearStoredSession();
      showLoginGateway("Your session has expired. Please sign in again.<br>เซสชันของคุณหมดอายุ กรุณาเข้าสู่ระบบอีกครั้ง");
    }

    function showAccessDenied() {
      document.querySelectorAll(".view").forEach(view => view.classList.remove("active"));
      document.querySelectorAll(".nav-button").forEach(button => button.classList.remove("active"));
      document.querySelector("#accessDeniedScreen")?.classList.add("show");
      document.querySelector("#pageHeading").textContent = "Access Denied";
      document.querySelector("#pageSubheading").textContent = "Permission required / ต้องมีสิทธิ์เข้าถึง";
    }

    function logoutUser() {
      clearStoredSession();
      window.clearInterval(sessionExpiryTimer);
      if (window.history?.replaceState) {
        window.history.replaceState(null, "", window.location.pathname + window.location.search);
      } else {
        window.location.hash = "";
      }
      document.querySelector("#loginPassword").value = "";
      showLoginGateway("");
    }

    function initializeAuthenticatedApplication() {
      if (authenticatedApplicationInitialized) return;
      authenticatedApplicationInitialized = true;
      initializeEditableDropdowns();
      setupUpdateAttachmentAndCc();
      setCaseActionStep("add");
      setupCsvDatabaseControls();
      setupMasterDataControls();
      loadContractsDatabase();
      renderAll();
      loadDriveDatabaseFromCloud();
      syncAddCaseLinkedFields("init");
    }

    function enterAuthenticatedApplication(session, restoreRoute = false) {
      currentUser = buildCurrentUser(session.username, session.role);
      if (!currentUser) {
        clearStoredSession();
        showLoginGateway("");
        return;
      }
      document.body.classList.remove("auth-locked");
      document.body.classList.add("auth-ready");
      updateProfileDisplay();
      applyRoleAccess();
      initializeAuthenticatedApplication();
      applyRoleAccess();
      renderAll();
      scheduleSessionExpiry();
      updateSessionActivity(true);
      setView(restoreRoute ? currentRouteView() : "dashboard", { replaceRoute: true });
    }

    async function authenticateDemoAccount(username, password) {
      const account = demoAccountConfig[username];
      if (!account) return null;
      const passwordHash = await sha256(password);
      if (!passwordHash || passwordHash !== account.passwordHash) return null;
      return account;
    }
'''
    if old_role_block not in html:
        raise RuntimeError("Could not find the role block for Login Gateway injection")
    html = html.replace(old_role_block, auth_config_js, 1)

    html = html.replace(
        '''    function canAccessView(viewName) {
      if (viewName === "user") return canUseUserCaseAction();''',
        '''    function canAccessView(viewName) {
      if (!currentUser) return false;
      if (viewName === "user") return canUseUserCaseAction();''',
        1,
    )
    html = html.replace(
        '''      if (!isAdmin()) {
        if (userNav) userNav.hidden = !canUseUserCaseAction();
        if (masterNav) masterNav.hidden = !canManageMasterData();
        if (confidentialNav) confidentialNav.hidden = !canViewConfidentialContracts();
        if (confidentialClassificationChoice) confidentialClassificationChoice.hidden = !canViewConfidentialContracts();
        if (!canViewConfidentialContracts()) contractStatusAccessView = "normal";
        if (newContractBtn) newContractBtn.hidden = !canUseUserCaseAction();
        document.querySelector("#user")?.classList.remove("active");
      }''',
        '''      if (userNav) userNav.hidden = !canUseUserCaseAction();
      if (masterNav) masterNav.hidden = !canManageMasterData();
      if (confidentialNav) confidentialNav.hidden = !canViewConfidentialContracts();
      if (confidentialClassificationChoice) confidentialClassificationChoice.hidden = !canViewConfidentialContracts();
      if (!canViewConfidentialContracts()) contractStatusAccessView = "normal";
      if (newContractBtn) newContractBtn.hidden = !canUseUserCaseAction();
      if (!canUseUserCaseAction()) document.querySelector("#user")?.classList.remove("active");''',
        1,
    )

    old_set_view = r'''    function setView(viewName) {
      if (viewName === "workflow") viewName = "contracts";
      if (!pageLabels[viewName]) viewName = "dashboard";
      if (!canAccessView(viewName)) {
        showToast("You do not have permission to access this page.");
        viewName = "dashboard";
      }
      if (viewName === "contracts" || viewName === "confidential") {
        contractStatusAccessView = viewName === "confidential" ? "confidential" : "normal";
        selectedLogContractId = "all";
        populateContractMasterFilters();
        renderContractsTable();
        renderKanban();
        renderLogView();
      }
      const targetView = viewName === "confidential" ? "contracts" : viewName;
      document.querySelectorAll(".nav-button").forEach(button => {
        button.classList.toggle("active", button.dataset.view === viewName);
      });
      document.querySelectorAll(".view").forEach(view => {
        view.classList.toggle("active", view.id === targetView);
      });
      document.querySelector("#pageHeading").textContent = pageLabels[viewName][0];
      document.querySelector("#pageSubheading").textContent = pageLabels[viewName][1];
    }'''
    new_set_view = r'''    function setView(viewName, options = {}) {
      if (!currentUser || !readStoredSession()) {
        showLoginGateway("");
        return;
      }
      if (viewName === "workflow") viewName = "contracts";
      if (!pageLabels[viewName]) viewName = "dashboard";
      if (!canAccessView(viewName)) {
        showAccessDenied();
        if (!options.skipRoute) {
          const route = routeForView(viewName);
          window.history.replaceState(null, "", `#/${route}`);
        }
        return;
      }
      document.querySelector("#accessDeniedScreen")?.classList.remove("show");
      if (viewName === "contracts" || viewName === "confidential") {
        contractStatusAccessView = viewName === "confidential" ? "confidential" : "normal";
        selectedLogContractId = "all";
        populateContractMasterFilters();
        renderContractsTable();
        renderKanban();
        renderLogView();
      }
      const targetView = viewName === "confidential" ? "contracts" : viewName;
      document.querySelectorAll(".nav-button").forEach(button => {
        button.classList.toggle("active", button.dataset.view === viewName);
      });
      document.querySelectorAll(".view").forEach(view => {
        view.classList.toggle("active", view.id === targetView);
      });
      document.querySelector("#pageHeading").textContent = pageLabels[viewName][0];
      document.querySelector("#pageSubheading").textContent = pageLabels[viewName][1];
      if (!options.skipRoute) {
        const route = `#/${routeForView(viewName)}`;
        if (options.replaceRoute) window.history.replaceState(null, "", route);
        else if (window.location.hash !== route) window.location.hash = route;
      }
      updateSessionActivity();
    }'''
    if old_set_view not in html:
        raise RuntimeError("Could not find setView for route protection")
    html = html.replace(old_set_view, new_set_view, 1)

    old_startup = r'''    // Initialize visible editable dropdowns with the embedded Turtle 23 directory.
    // Each field supports selecting from the list and typing a new value.
    initializeEditableDropdowns();
    setupUpdateAttachmentAndCc();
    setCaseActionStep("add");
    applyRoleAccess();
    setupCsvDatabaseControls();
    setupMasterDataControls();
    loadContractsDatabase();
    renderAll();
    loadDriveDatabaseFromCloud();
    syncAddCaseLinkedFields("init");
    const initialView = window.location.hash.replace("#", "");
    if (initialView) setView(initialView);
'''
    new_startup = r'''    function initializeLoginGateway() {
      const demoSection = document.querySelector("#demoAccountSection");
      if (demoSection) demoSection.hidden = !authConfig.showDemoAccounts;

      document.querySelector("#passwordToggle")?.addEventListener("click", event => {
        const input = document.querySelector("#loginPassword");
        const show = input.type === "password";
        input.type = show ? "text" : "password";
        event.currentTarget.textContent = show ? "Hide" : "Show";
        event.currentTarget.setAttribute("aria-label", show ? "Hide password" : "Show password");
      });

      document.querySelector("#loginClear")?.addEventListener("click", () => {
        document.querySelector("#loginForm")?.reset();
        setLoginError("");
        setFieldValidation("Username", "");
        setFieldValidation("Password", "");
        document.querySelector("#loginUsername")?.focus();
      });

      document.querySelectorAll("[data-demo-account]").forEach(button => {
        button.addEventListener("click", () => {
          const username = button.dataset.demoAccount;
          document.querySelector("#loginUsername").value = username;
          document.querySelector("#loginPassword").value = "";
          setLoginError("");
          setFieldValidation("Username", "");
          setFieldValidation("Password", "");
          document.querySelector("#loginPassword")?.focus();
        });
      });

      document.querySelector("#loginForm")?.addEventListener("submit", async event => {
        event.preventDefault();
        const usernameInput = document.querySelector("#loginUsername");
        const passwordInput = document.querySelector("#loginPassword");
        const username = usernameInput.value.trim();
        const password = passwordInput.value;
        const usernameMessage = username ? "" : "Username is required.<br>กรุณากรอกชื่อผู้ใช้งาน";
        const passwordMessage = password ? "" : "Password is required.<br>กรุณากรอกรหัสผ่าน";
        setFieldValidation("Username", usernameMessage);
        setFieldValidation("Password", passwordMessage);
        setLoginError("");
        if (usernameMessage || passwordMessage) {
          (usernameMessage ? usernameInput : passwordInput).focus();
          return;
        }
        setLoginLoading(true);
        try {
          const account = await authenticateDemoAccount(username, password);
          await new Promise(resolve => window.setTimeout(resolve, 280));
          if (!account) {
            passwordInput.value = "";
            setLoginError("Incorrect username or password. Please try again.<br>ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง กรุณาลองใหม่อีกครั้ง");
            passwordInput.focus();
            return;
          }
          const now = Date.now();
          const remember = document.querySelector("#rememberMe").checked;
          const session = {
            username,
            role: account.role,
            loginTime: now,
            lastActivity: now,
            expiresAt: now + authConfig.sessionTimeoutMs,
            remember
          };
          persistSession(session, remember);
          passwordInput.value = "";
          enterAuthenticatedApplication(session, false);
        } finally {
          setLoginLoading(false);
        }
      });

      document.querySelector("#profileTrigger")?.addEventListener("click", event => {
        const menu = document.querySelector("#profileMenu");
        const open = !menu.classList.contains("open");
        menu.classList.toggle("open", open);
        event.currentTarget.setAttribute("aria-expanded", String(open));
      });
      document.querySelector("#logoutBtn")?.addEventListener("click", logoutUser);
      document.querySelector("#backToDashboardBtn")?.addEventListener("click", () => setView("dashboard"));
      document.addEventListener("click", event => {
        if (!event.target.closest("#profileMenu")) {
          document.querySelector("#profileMenu")?.classList.remove("open");
          document.querySelector("#profileTrigger")?.setAttribute("aria-expanded", "false");
        }
      });
      ["pointerdown", "keydown", "scroll", "touchstart"].forEach(eventName => {
        document.addEventListener(eventName, () => updateSessionActivity(), { passive: true });
      });
      window.addEventListener("hashchange", () => {
        if (currentUser) setView(currentRouteView(), { skipRoute: true });
      });

      const storedSession = readStoredSession();
      if (storedSession) enterAuthenticatedApplication(storedSession, true);
      else showLoginGateway("");
    }

    initializeLoginGateway();
'''
    if old_startup not in html:
        raise RuntimeError("Could not find startup block for authentication gate")
    html = html.replace(old_startup, new_startup, 1)

    admin_workflow_css = r'''
    .admin-only-panel[hidden] { display: none !important; }
    .admin-only-badge { margin-left: 6px; }
    .admin-workflow-grid { display: grid; grid-template-columns: repeat(12, minmax(0, 1fr)); gap: 12px; }
    .admin-workflow-grid > .form-field { grid-column: span 4; }
    .admin-workflow-grid > .span-6 { grid-column: span 6; }
    .admin-workflow-grid > .full { grid-column: 1 / -1; }
    .admin-summary-band { grid-column: 1 / -1; display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 10px; }
    .admin-summary-section { min-width: 0; padding: 12px; border: 1px solid var(--line); border-radius: 6px; background: #fafbf8; }
    .admin-summary-section h3 { margin: 0 0 8px; font-size: 13px; }
    .admin-summary-section dl { display: grid; gap: 6px; margin: 0; }
    .admin-summary-section dl div { display: flex; justify-content: space-between; gap: 10px; font-size: 11px; }
    .admin-summary-section dt { color: var(--muted); }
    .admin-summary-section dd { margin: 0; text-align: right; font-weight: 750; overflow-wrap: anywhere; }
    .due-approver-select-field, .due-approver-detail-field { display: none !important; }
    .admin-rights-table td:first-child, .admin-rights-table th:first-child { white-space: nowrap; }
    .password-validation-message { min-height: 18px; color: #b42318; font-size: 11px; }
    @media (max-width: 820px) {
      .admin-workflow-grid > .form-field, .admin-workflow-grid > .span-6 { grid-column: 1 / -1; }
      .admin-summary-band { grid-template-columns: 1fr; }
    }
'''
    html = html.replace("  </style>", admin_workflow_css + "\n  </style>", 1)

    due_decision_email_modal = r'''
  <div class="modal-backdrop" id="dueDecisionEmailModal" aria-hidden="true">
    <section class="modal email-compose-modal" role="dialog" aria-modal="true" aria-labelledby="dueDecisionEmailTitle">
      <header class="modal-header">
        <div>
          <h2 id="dueDecisionEmailTitle">Due Date Decision Notification</h2>
          <p id="dueDecisionEmailSubtitle">แจ้งผลการพิจารณากลับไปยังผู้ขอ</p>
        </div>
        <button class="icon-button" id="closeDueDecisionEmailModal" type="button" title="Close">×</button>
      </header>
      <div class="modal-body">
        <div class="email-compose-grid">
          <label class="form-field full">
            <span class="bilingual-label"><span>To</span><span>อีเมลผู้ขอ</span></span>
            <input class="input" type="email" id="dueDecisionEmailTo" placeholder="requester@turtle23.com">
          </label>
          <label class="form-field full">
            <span>Subject</span>
            <input class="input" id="dueDecisionEmailSubject">
          </label>
          <label class="form-field full">
            <span class="bilingual-label"><span>Message</span><span>ข้อความ</span></span>
            <textarea id="dueDecisionEmailBody"></textarea>
          </label>
        </div>
        <div class="email-modal-actions">
          <button class="ghost-button" type="button" id="copyDueDecisionEmailBtn">Copy Email</button>
          <button class="ghost-button" type="button" id="closeDueDecisionEmailBtn">Close</button>
          <button class="primary-button" type="button" id="sendDueDecisionEmailBtn"><span>✉</span> Send Email</button>
        </div>
      </div>
    </section>
  </div>
'''
    html = html.replace(
        '  <div class="modal-backdrop" id="actionConfirmModal" aria-hidden="true">',
        due_decision_email_modal + '\n  <div class="modal-backdrop" id="actionConfirmModal" aria-hidden="true">',
        1,
    )

    html = html.replace(
        '''        <button class="nav-button" data-view="master" title="Master Data">
          <span>▤</span><span class="nav-label">Master Data</span><span class="nav-count">5</span>
        </button>''',
        '''        <button class="nav-button" data-view="master" title="Master Data">
          <span>▤</span><span class="nav-label">Master Data</span><span class="nav-count">5</span>
        </button>
        <button class="nav-button admin-tools-nav" data-view="admin" title="Admin Tools" data-admin-only hidden>
          <span>⚙</span><span class="nav-label">Admin Tools</span><span class="nav-count">2</span>
        </button>''',
        1,
    )

    html = html.replace(
        '<button type="button" role="menuitem" disabled>Change Password</button>',
        '<button type="button" role="menuitem" id="adminPasswordMenu" hidden>Change Password</button>',
        1,
    )
    html = html.replace(
        '<h2 class="reference-page-title"><span>Adjust Due Date</span><small>ปรับวันครบกำหนด</small></h2>',
        '<h2 class="reference-page-title"><span>Request Due Date Adjustment</span><small>ขอปรับวันครบกำหนด</small></h2>',
        1,
    )
    html = html.replace('<strong>Adjust Due Date</strong>', '<strong>Request Due Date</strong>', 1)
    html = html.replace(
        'Add Case → Update Status → Close Case → Adjust Due Date',
        'Add Case → Update Status → Close Case → Request Due Date',
        1,
    )
    html = html.replace(
        '<span class="badge amber" id="dueRequestStatus">Pending Approval</span>',
        '<span class="badge amber" id="dueRequestStatus">Pending Admin Review</span>',
        1,
    )
    html = html.replace(
        '''          <label class="form-field full">
            <span>Subject</span>
            <input class="input" id="dueApprovalEmailSubject">
          </label>''',
        '''          <label class="form-field full">
            <span class="bilingual-label"><span>CC</span><span>สำเนาอีเมล</span></span>
            <input class="input" type="text" id="dueApprovalEmailCc" placeholder="email1@turtle23.com, email2@turtle23.com" aria-describedby="dueApprovalEmailCcHint">
            <small class="form-hint" id="dueApprovalEmailCcHint">Optional · Separate multiple emails with comma or semicolon / ไม่บังคับ · คั่นหลายอีเมลด้วย comma หรือ semicolon</small>
          </label>
          <label class="form-field full">
            <span>Subject</span>
            <input class="input" id="dueApprovalEmailSubject">
          </label>''',
        1,
    )
    html = html.replace('name="approver" id="dueApprovalTo" multiple size="2" required', 'name="approver" id="dueApprovalTo" multiple size="2"', 1)
    html = html.replace('name="approverEmail" id="dueApprovalEmail" required readonly', 'name="approverEmail" id="dueApprovalEmail" readonly', 1)
    html = html.replace(
        '''                  <div class="case-form-section-title"><span class="section-index">4</span>Reason and submission <small>ต้องระบุเหตุผลทุกครั้ง</small></div>
                  <label class="reason-status-panel alert-yellow" id="dueAdjustmentReasonField">''',
        '''                  <div class="case-form-section-title"><span class="section-index">4</span>Request details <small>รายละเอียดคำขอและไฟล์ประกอบ</small></div>
                  <label class="form-field span-4">
                    <span class="bilingual-label"><span>Request Type<span class="field-required-mark" aria-hidden="true">*</span></span><span>ประเภทคำขอ<span class="field-required-mark" aria-hidden="true">*</span></span></span>
                    <select class="select" name="requestType" id="dueRequestType" required>
                      <option value="Extension">Extension / ขอขยายเวลา</option>
                      <option value="Correction">Correction / แก้ไขวันที่</option>
                      <option value="Other">Other / อื่น ๆ</option>
                    </select>
                  </label>
                  <label class="form-field span-6">
                    <span class="bilingual-label"><span>Attachment</span><span>ไฟล์แนบ</span></span>
                    <input class="input" type="file" name="attachments" id="dueRequestAttachments" multiple>
                    <small class="form-hint">Optional supporting files / แนบไฟล์ประกอบได้ตามความจำเป็น</small>
                  </label>
                  <label class="reason-status-panel alert-yellow" id="dueAdjustmentReasonField">''',
        1,
    )

    admin_master_panels = r'''
              <section class="panel master-data-panel master-data-panel-full admin-only-panel" id="dueDateApprovalPanel" data-admin-only hidden>
                <div class="panel-header">
                  <div>
                    <h2>Due Date Approval <span class="badge black admin-only-badge">Admin Only</span></h2>
                    <small>อนุมัติการปรับวันครบกำหนด</small>
                  </div>
                </div>
                <div class="table-wrap">
                  <table class="master-table">
                    <thead><tr><th>Request ID</th><th>Contract</th><th>Requested Due Date</th><th>Requested By</th><th>Status</th><th></th></tr></thead>
                    <tbody id="dueApprovalQueueRows"></tbody>
                  </table>
                </div>
                <form class="admin-workflow-grid" id="dueApprovalDecisionForm" hidden>
                  <input type="hidden" name="requestId" id="adminDueRequestId">
                  <div class="admin-summary-band" id="adminDueRequestSummary"></div>
                  <label class="form-field">
                    <span class="bilingual-label"><span>Decision</span><span>ผลการพิจารณา</span></span>
                    <select class="select" name="decision" id="adminDueDecision" required>
                      <option value="Approve and Apply">Approve and Apply</option>
                      <option value="Reject Request">Reject Request</option>
                      <option value="Request More Information">Request More Information</option>
                    </select>
                  </label>
                  <label class="form-field">
                    <span class="bilingual-label"><span>Final Due Date</span><span>วันครบกำหนดสุดท้าย</span></span>
                    <input class="input" type="date" name="finalDue" id="adminFinalDue">
                  </label>
                  <label class="form-field">
                    <span class="bilingual-label"><span>Admin Password Confirmation</span><span>ยืนยันรหัสผ่านผู้ดูแล</span></span>
                    <input class="input" type="password" name="adminPassword" required autocomplete="current-password">
                  </label>
                  <label class="form-field span-6">
                    <span class="bilingual-label"><span>Admin Decision Reason</span><span>เหตุผลการพิจารณา</span></span>
                    <textarea name="decisionReason" required></textarea>
                  </label>
                  <label class="form-field span-6">
                    <span class="bilingual-label"><span>Reason for Different Final Date</span><span>เหตุผลเมื่อกำหนดวันที่ต่างจากคำขอ</span></span>
                    <textarea name="differentDateReason" id="adminDifferentDateReason"></textarea>
                  </label>
                  <label class="form-field full">
                    <span class="bilingual-label"><span>Admin Note</span><span>หมายเหตุผู้ดูแล</span></span>
                    <textarea name="adminNote"></textarea>
                  </label>
                  <div class="user-form-actions full">
                    <button class="primary-button" type="submit">Confirm and Apply / ยืนยันผล</button>
                    <button class="secondary-button" type="button" id="cancelDueAdminDecision">Cancel</button>
                  </div>
                </form>
                <div class="table-wrap">
                  <table class="master-table">
                    <thead><tr><th>Request ID</th><th>Decision</th><th>Final Due Date</th><th>Admin</th><th>Date</th></tr></thead>
                    <tbody id="dueAdjustmentHistoryRows"></tbody>
                  </table>
                </div>
              </section>

              <section class="panel master-data-panel master-data-panel-full admin-only-panel" id="passwordManagementPanel" data-admin-only hidden>
                <div class="panel-header">
                  <div>
                    <h2>Password Management <span class="badge black admin-only-badge">Admin Only</span></h2>
                    <small>จัดการรหัสผ่าน</small>
                  </div>
                </div>
                <form class="admin-workflow-grid" id="passwordManagementForm">
                  <label class="form-field">
                    <span class="bilingual-label"><span>Username</span><span>ชื่อผู้ใช้งาน</span></span>
                    <select class="select" name="username" id="passwordAccountSelect" required>
                      <option value="viewer">viewer</option><option value="user">user</option><option value="confidential">confidential</option><option value="admin">admin</option>
                    </select>
                  </label>
                  <label class="form-field"><span>Level</span><input class="input readonly-input" id="passwordAccountLevel" readonly></label>
                  <label class="form-field"><span>Role Name</span><input class="input readonly-input" id="passwordAccountRole" readonly></label>
                  <label class="form-field"><span>New Password</span><input class="input" type="password" name="newPassword" autocomplete="new-password"></label>
                  <label class="form-field"><span>Confirm New Password</span><input class="input" type="password" name="confirmPassword" autocomplete="new-password"></label>
                  <label class="form-field"><span>Admin Current Password</span><input class="input" type="password" name="adminCurrentPassword" required autocomplete="current-password"></label>
                  <label class="form-field full"><span>Remark</span><textarea name="remark"></textarea></label>
                  <div class="password-validation-message full" id="passwordManagementMessage" aria-live="polite"></div>
                  <div class="user-form-actions full">
                    <button class="primary-button" type="submit">Change Password</button>
                    <button class="secondary-button" type="button" id="resetPasswordBtn">Reset to Default Password</button>
                    <button class="secondary-button" type="button" id="unlockAccountBtn">Unlock Account</button>
                  </div>
                </form>
                <div class="table-wrap">
                  <table class="master-table admin-rights-table">
                    <thead><tr><th>Username</th><th>Level</th><th>Role</th><th>Access Rights</th><th>Account Status</th></tr></thead>
                    <tbody id="passwordAccessRightsRows"></tbody>
                  </table>
                </div>
                <div class="table-wrap">
                  <table class="master-table">
                    <thead><tr><th>Date</th><th>Admin</th><th>Username</th><th>Action</th><th>Remark</th></tr></thead>
                    <tbody id="passwordHistoryRows"></tbody>
                  </table>
                </div>
              </section>
'''
    master_end_anchor = '''              </section>
            </div>
          </section>
        </section>

        <section class="view" id="notifications" hidden>'''
    if master_end_anchor not in html:
        raise RuntimeError("Could not find Master Data closing anchor")
    html = html.replace(
        master_end_anchor,
        '''              </section>
            </div>
          </section>
        </section>

        <section class="view" id="admin">
          <section class="panel">
            <div class="panel-header">
              <div>
                <h2>Admin Tools <span class="badge black admin-only-badge">Admin Only</span></h2>
                <small>เครื่องมือสำหรับผู้ดูแลระบบ</small>
              </div>
            </div>
            <div class="master-data-grid">
''' + admin_master_panels + '''            </div>
          </section>
        </section>

        <section class="view" id="notifications" hidden>''',
        1,
    )

    html = html.replace(
        '''      master: ["Master Data", "แก้ไขข้อมูล dropdown และบันทึกกลับ Shared Drive"],
      confidential:''',
        '''      master: ["Master Data", "แก้ไขข้อมูล dropdown และบันทึกกลับ Shared Drive"],
      admin: ["Admin Tools", "อนุมัติ Due Date และจัดการรหัสผ่าน"],
      confidential:''',
        1,
    )
    html = html.replace(
        '''        "master-data": "master"
      };''',
        '''        "master-data": "master",
        "admin-tools": "admin"
      };''',
        1,
    )
    html = html.replace(
        '''        confidential: "confidential-contract-status",
        master: "master-data"
      })[viewName]''',
        '''        confidential: "confidential-contract-status",
        master: "master-data",
        admin: "admin-tools"
      })[viewName]''',
        1,
    )
    html = html.replace(
        '''      if (viewName === "master") return canManageMasterData();
      return true;''',
        '''      if (viewName === "master" || viewName === "admin") return canManageMasterData();
      return true;''',
        1,
    )
    html = html.replace(
        '''      const masterNav = document.querySelector('.nav-button[data-view="master"]');
      const confidentialNav''',
        '''      const masterNav = document.querySelector('.nav-button[data-view="master"]');
      const adminNav = document.querySelector('.nav-button[data-view="admin"]');
      const confidentialNav''',
        1,
    )
    html = html.replace(
        '''        if (masterNav) masterNav.hidden = !canManageMasterData();
        if (confidentialNav)''',
        '''        if (masterNav) masterNav.hidden = !canManageMasterData();
        if (adminNav) adminNav.hidden = !canManageMasterData();
        if (confidentialNav)''',
        1,
    )

    html = html.replace(
        '''    async function authenticateDemoAccount(username, password) {
      const account = demoAccountConfig[username];
      if (!account) return null;
      const passwordHash = await sha256(password);
      if (!passwordHash || passwordHash !== account.passwordHash) return null;
      return account;
    }''',
        '''    async function authenticateDemoAccount(username, password) {
      const account = demoAccountConfig[username];
      const credential = credentialRecord(username);
      if (!account || !credential || credential.locked) return null;
      const passwordHash = await sha256(password);
      if (!passwordHash || passwordHash !== credential.passwordHash) {
        credential.failedAttempts = Number(credential.failedAttempts || 0) + 1;
        if (credential.failedAttempts >= 5) credential.locked = true;
        persistAdministrativeSecurityState();
        return null;
      }
      credential.failedAttempts = 0;
      persistAdministrativeSecurityState();
      return account;
    }''',
        1,
    )
    html = html.replace(
        'const pending = dueDateAdjustmentRequests.slice().reverse().find(item => item.contractId === contract.id && item.status === "Pending Approval");',
        'const latestDueRequest = dueDateAdjustmentRequests.slice().reverse().find(item => item.contractId === contract.id);',
        1,
    )
    html = html.replace(
        'statusBadge.textContent = pending ? `Pending Approval · ${pending.requestedDue}` : "Ready to Request";',
        'statusBadge.textContent = latestDueRequest ? `${latestDueRequest.status} · ${latestDueRequest.requestedDue}` : "Ready to Request";',
        1,
    )
    html = html.replace(
        'statusBadge.className = `badge ${pending ? "amber" : "blue"}`;',
        'statusBadge.className = `badge ${latestDueRequest?.status === "Approved and Applied" ? "green" : latestDueRequest ? "amber" : "blue"}`;',
        1,
    )
    html = html.replace(
        '''    document.querySelector("#dueDateAdjustmentForm")?.addEventListener("submit", event => {
      event.preventDefault();
      const formElement = event.currentTarget;''',
        '''    document.querySelector("#dueDateAdjustmentForm")?.addEventListener("submit", async event => {
      event.preventDefault();
      if (!canSubmitDueDateRequest()) {
        showToast("Access Denied: Due date requests require Contract User access.");
        return;
      }
      const formElement = event.currentTarget;''',
        1,
    )
    html = html.replace(
        '''      const requestedBy = String(form.get("requestedBy") || contract.owner || "User").trim();
      const selectedApproverEmails = selectedDueApproverEmails();''',
        '''      const requestedBy = currentUser?.username || String(form.get("requestedBy") || contract.owner || "User").trim();
      const requestType = String(form.get("requestType") || "Extension").trim();
      const attachmentFiles = [...(document.querySelector("#dueRequestAttachments")?.files || [])];
      const attachments = await Promise.all(attachmentFiles.map(async file => ({
        fileName: file.name,
        originalFileName: file.name,
        mimeType: file.type || "application/octet-stream",
        fileSize: file.size || 0,
        base64: await fileToBase64(file)
      })));
      const selectedApproverEmails = selectedDueApproverEmails();''',
        1,
    )
    html = html.replace(
        '''      const approver = selectedApproverEmails.join(", ");
      const approverEmail = normalizeEmailList(selectedApproverEmails.join(", "));''',
        '''      const approverEmail = normalizeEmailList(selectedApproverEmails.join(", ")) || "admin@turtle23.com";''',
        1,
    )
    html = html.replace(
        '''      if (!approver || !approverEmail) {
        document.querySelector(approver ? "#dueApprovalEmail" : "#dueApprovalTo")?.focus();
        showToast("Please select at least one approver email.");
        return;
      }
      if (!isValidEmailList(approverEmail)) {
        document.querySelector("#dueApprovalEmail")?.focus();
        showToast("Please check the approver email list. Separate multiple emails with comma or semicolon.");
        return;
      }

''',
        "",
        1,
    )
    html = html.replace(
        '''        requestedDue,
        reason,
        requestedBy,
        approver,
        approverEmail,
        requestedAt: todayInputValue(),
        status: "Pending Approval"
      };''',
        '''        requestedDue,
        requestType,
        reason,
        attachments,
        requestedBy,
        requestedByEmail: emailForPerson(contract.owner) || emailForPerson(requestedBy) || "",
        approver: "System Administrator",
        approverEmail,
        requestedAt: localIsoDateTime(),
        currentBalance: Number(contract.balance || 0),
        currentAlert: contract.alert || contract.status || "",
        status: "Pending Admin Review"
      };''',
        1,
    )
    html = html.replace(
        'station: `From ${request.requestedBy} >> To ${approver}`,\n        from: request.requestedBy,\n        to: approver,',
        'station: `From ${request.requestedBy} >> To admin`,\n        from: request.requestedBy,\n        to: "admin",',
        1,
    )
    html = html.replace(
        '''      dueDateAdjustmentRequests.push(request);
      contract.pendingDueDateRequest = { ...request };
      addLogRecord({''',
        '''      dueDateAdjustmentRequests.push(request);
      contract.pendingDueDateRequest = { ...request };
      notificationQueue.unshift({
        contractId: contract.id,
        name: contract.name,
        department: contract.department || "-",
        owner: request.requestedBy,
        currentStage: contract.stage,
        slaAlert: contract.alert || contract.status,
        dueDate: contract.due,
        slaBalance: contract.balance,
        station: `From ${request.requestedBy} >> To admin`,
        suggestedRuleId: request.requestId,
        triggerType: "Due Date Request",
        toGroup: "System Administrator",
        toEmails: approverEmail,
        ccGroup: "",
        ccEmails: "",
        templateId: "TMP-DUE-ADMIN-REVIEW",
        sendNow: "Yes",
        frequency: "Once",
        subject: `[Contract Tracking] Admin Review Required: ${request.requestId} - ${contract.id}`,
        body: `Requested Due Date ${requestedDue}: ${reason}`
      });
      addLogRecord({''',
        1,
    )
    html = html.replace(
        'showToast(`${request.requestId} created · Due Date remains unchanged pending approval`);',
        'showToast(`${request.requestId} created · Pending Admin Review · Due Date unchanged`);',
        1,
    )
    html = html.replace("`Request Status: Pending Approval`", "`Request Status: Pending Admin Review`", 1)
    html = html.replace(
        '''      return {
        to: request.approverEmail || "",
        subject: `[Contract Tracking] Approval Required: Due Date Adjustment - ${contract.id} - ${contract.name}`,
        body
      };''',
        '''      return {
        to: request.approverEmail || "",
        cc: normalizeEmailList(request.cc || ""),
        subject: `[Contract Tracking] Approval Required: Due Date Adjustment - ${contract.id} - ${contract.name}`,
        body
      };''',
        1,
    )
    html = html.replace(
        '''      document.querySelector("#dueApprovalEmailTo").value = draft.to || "";
      document.querySelector("#dueApprovalEmailSubject").value = draft.subject || "";''',
        '''      document.querySelector("#dueApprovalEmailTo").value = draft.to || "";
      document.querySelector("#dueApprovalEmailCc").value = draft.cc || "";
      document.querySelector("#dueApprovalEmailCc").setCustomValidity("");
      document.querySelector("#dueApprovalEmailSubject").value = draft.subject || "";''',
        1,
    )
    html = html.replace(
        '''      return {
        to: normalizeEmailList(document.querySelector("#dueApprovalEmailTo")?.value || ""),
        subject: document.querySelector("#dueApprovalEmailSubject")?.value || "",''',
        '''      return {
        to: normalizeEmailList(document.querySelector("#dueApprovalEmailTo")?.value || ""),
        cc: normalizeEmailList(document.querySelector("#dueApprovalEmailCc")?.value || ""),
        subject: document.querySelector("#dueApprovalEmailSubject")?.value || "",''',
        1,
    )
    html = html.replace(
        '''      const recipients = normalizeEmailList(draft.to).replace(/,\\s*/g, ",");
      window.location.href = `mailto:${encodeURIComponent(recipients)}?subject=${encodeURIComponent(draft.subject)}&body=${encodeURIComponent(draft.body)}`;''',
        '''      if (draft.cc && !isValidEmailList(draft.cc)) {
        const ccInput = document.querySelector("#dueApprovalEmailCc");
        ccInput?.setCustomValidity("Enter a valid email address. / กรุณากรอกอีเมลให้ถูกต้อง");
        ccInput?.focus();
        ccInput?.reportValidity();
        return;
      }
      const recipients = normalizeEmailList(draft.to).replace(/,\\s*/g, ",");
      const ccRecipients = normalizeEmailList(draft.cc).replace(/,\\s*/g, ",");
      const ccParameter = ccRecipients ? `cc=${encodeURIComponent(ccRecipients)}&` : "";
      window.location.href = `mailto:${encodeURIComponent(recipients)}?${ccParameter}subject=${encodeURIComponent(draft.subject)}&body=${encodeURIComponent(draft.body)}`;''',
        1,
    )
    html = html.replace(
        '''      const text = `To: ${draft.to}\\nSubject: ${draft.subject}\\n\\n${draft.body}`;''',
        '''      const text = `To: ${draft.to}\\nCC: ${draft.cc || "-"}\\nSubject: ${draft.subject}\\n\\n${draft.body}`;''',
        1,
    )
    html = html.replace(
        '''    document.querySelector("#closeDueApprovalEmailModal")?.addEventListener("click", closeDueApprovalEmailPopup);''',
        '''    document.querySelector("#dueApprovalEmailCc")?.addEventListener("input", event => event.currentTarget.setCustomValidity(""));
    document.querySelector("#dueApprovalEmailCc")?.addEventListener("blur", event => {
      const value = normalizeEmailList(event.currentTarget.value || "");
      event.currentTarget.value = value;
      event.currentTarget.setCustomValidity(value && !isValidEmailList(value) ? "Enter a valid email address. / กรุณากรอกอีเมลให้ถูกต้อง" : "");
    });
    document.querySelector("#closeDueApprovalEmailModal")?.addEventListener("click", closeDueApprovalEmailPopup);''',
        1,
    )

    html = html.replace(
        '''          contracts,
          logRecords,
          masterData''',
        '''          contracts,
          logRecords,
          dueDateAdjustmentRequests,
          masterData''',
        1,
    )
    html = html.replace(
        '''        const parsedContracts = parsed.contracts;
        const parsedLogs = Array.isArray(parsed.logRecords) ? parsed.logRecords : [];''',
        '''        const parsedContracts = parsed.contracts;
        const parsedLogs = Array.isArray(parsed.logRecords) ? parsed.logRecords : [];
        if (Array.isArray(parsed.dueDateAdjustmentRequests)) {
          dueDateAdjustmentRequests.splice(0, dueDateAdjustmentRequests.length, ...parsed.dueDateAdjustmentRequests);
        }''',
        1,
    )

    admin_security_js = r'''
    const administrativeSecurityStorageKey = "t23-contract-administrative-security-v1";
    const defaultPasswordHashes = Object.freeze(Object.fromEntries(
      Object.entries(demoAccountConfig).map(([username, account]) => [username, account.passwordHash])
    ));
    const deprecatedConfidentialPasswordHashes = Object.freeze([
      "eb44db23e31fa6df080d8cb73efb730478abafd2724781d1199d2fb090f3b293"
    ]);

    function defaultAdministrativeSecurityState() {
      return {
        accounts: Object.fromEntries(Object.keys(demoAccountConfig).map(username => [username, {
          passwordHash: defaultPasswordHashes[username], passwordHashHistory: [defaultPasswordHashes[username]], failedAttempts: 0, locked: false
        }])),
        passwordHistory: [],
        dueDateHistory: []
      };
    }

    function loadAdministrativeSecurityState() {
      const fallback = defaultAdministrativeSecurityState();
      try {
        const parsed = JSON.parse(localStorage.getItem(administrativeSecurityStorageKey) || "null");
        if (!parsed || typeof parsed !== "object") return fallback;
        Object.keys(fallback.accounts).forEach(username => {
          const saved = parsed.accounts?.[username] || {};
          const savedPasswordHash = String(saved.passwordHash || defaultPasswordHashes[username]);
          const migratedPasswordHash = username === "confidential" && deprecatedConfidentialPasswordHashes.includes(savedPasswordHash)
            ? defaultPasswordHashes[username]
            : savedPasswordHash;
          fallback.accounts[username] = {
            passwordHash: migratedPasswordHash,
            passwordHashHistory: Array.isArray(saved.passwordHashHistory)
              ? [migratedPasswordHash, ...saved.passwordHashHistory.map(String).filter(Boolean)].filter((value, index, rows) => rows.indexOf(value) === index).slice(0, 5)
              : [migratedPasswordHash],
            failedAttempts: Number(saved.failedAttempts || 0),
            locked: Boolean(saved.locked)
          };
        });
        fallback.passwordHistory = Array.isArray(parsed.passwordHistory) ? parsed.passwordHistory : [];
        fallback.dueDateHistory = Array.isArray(parsed.dueDateHistory) ? parsed.dueDateHistory : [];
      } catch (error) {
        return fallback;
      }
      return fallback;
    }

    const administrativeSecurityState = loadAdministrativeSecurityState();

    function persistAdministrativeSecurityState() {
      localStorage.setItem(administrativeSecurityStorageKey, JSON.stringify(administrativeSecurityState));
    }

    function credentialRecord(username) {
      return administrativeSecurityState.accounts[String(username || "").trim()] || null;
    }

    function isSystemAdministrator() {
      return currentUser?.username === "admin" && currentUser?.role === "admin" && Number(currentUser?.level) === 4;
    }

    function requireSystemAdministrator(message = "Access Denied: Password management is available only to the System Administrator.") {
      if (isSystemAdministrator()) return true;
      showToast(message);
      return false;
    }

    function canSubmitDueDateRequest() {
      return Boolean(currentUser && ["user", "confidential", "admin"].includes(currentUser.role) && Number(currentUser.level) >= 2);
    }

    async function verifyAdminCurrentPassword(password) {
      const hash = await sha256(password);
      return Boolean(hash && hash === credentialRecord("admin")?.passwordHash);
    }

    function passwordRoleSummary(username) {
      const role = accessRoleConfig[demoAccountConfig[username]?.role];
      return role ? { level: `Level ${role.level}`, role: role.nameEn } : { level: "-", role: "-" };
    }

    function passwordAccessRights(username) {
      return ({
        viewer: "Dashboard, Contract Status",
        user: "Dashboard, Contract Status, User Case Action, Due Date Request",
        confidential: "Dashboard, Contract Status, Confidential Status, User Case Action, Due Date Request",
        admin: "All menus, Master Data, Due Date Approval, Password Management"
      })[username] || "-";
    }

    function appendPasswordAudit(username, action, remark = "") {
      administrativeSecurityState.passwordHistory.unshift({
        date: localIsoDateTime(), admin: "admin", username, action, remark: String(remark || "").trim()
      });
      administrativeSecurityState.passwordHistory = administrativeSecurityState.passwordHistory.slice(0, 100);
      persistAdministrativeSecurityState();
    }

    function setPasswordManagementMessage(message = "", success = false) {
      const node = document.querySelector("#passwordManagementMessage");
      if (!node) return;
      node.textContent = message;
      node.style.color = success ? "#4f6f00" : "#b42318";
    }

    function renderPasswordManagement() {
      if (!isSystemAdministrator()) {
        requireSystemAdministrator();
        return;
      }
      const username = document.querySelector("#passwordAccountSelect")?.value || "viewer";
      const role = passwordRoleSummary(username);
      if (document.querySelector("#passwordAccountLevel")) document.querySelector("#passwordAccountLevel").value = role.level;
      if (document.querySelector("#passwordAccountRole")) document.querySelector("#passwordAccountRole").value = role.role;
      const rights = document.querySelector("#passwordAccessRightsRows");
      if (rights) rights.innerHTML = Object.keys(demoAccountConfig).map(accountName => {
        const accountRole = passwordRoleSummary(accountName);
        const account = credentialRecord(accountName);
        return `<tr><td>${escapeHtml(accountName)}</td><td>${escapeHtml(accountRole.level)}</td><td>${escapeHtml(accountRole.role)}</td><td>${escapeHtml(passwordAccessRights(accountName))}</td><td>${account?.locked ? '<span class="badge red">Locked</span>' : '<span class="badge green">Active</span>'}</td></tr>`;
      }).join("");
      const history = document.querySelector("#passwordHistoryRows");
      if (history) history.innerHTML = administrativeSecurityState.passwordHistory.length
        ? administrativeSecurityState.passwordHistory.map(item => `<tr><td>${escapeHtml(formatDateTime(item.date))}</td><td>${escapeHtml(item.admin)}</td><td>${escapeHtml(item.username)}</td><td>${escapeHtml(item.action)}</td><td>${escapeHtml(item.remark || "-")}</td></tr>`).join("")
        : '<tr><td colspan="5">No password change history</td></tr>';
    }

    async function changeManagedPassword(formElement) {
      if (!requireSystemAdministrator()) return false;
      const form = new FormData(formElement);
      const username = String(form.get("username") || "");
      const newPassword = String(form.get("newPassword") || "");
      const confirmPassword = String(form.get("confirmPassword") || "");
      const adminPassword = String(form.get("adminCurrentPassword") || "");
      const remark = String(form.get("remark") || "");
      let error = "";
      if (newPassword.length < 8) error = "New Password must contain at least 8 characters.";
      else if (!/[A-Za-z]/.test(newPassword)) error = "New Password must contain an English letter.";
      else if (!/[0-9]/.test(newPassword)) error = "New Password must contain a number.";
      else if (!/[^A-Za-z0-9]/.test(newPassword)) error = "New Password must contain a special character.";
      else if (newPassword !== confirmPassword) error = "New Password and Confirm New Password must match.";
      else if (!await verifyAdminCurrentPassword(adminPassword)) error = "Admin Current Password is incorrect.";
      const newHash = error ? "" : await sha256(newPassword);
      if (!error && (credentialRecord(username)?.passwordHashHistory || []).includes(newHash)) error = "The new password must not match a recently used password.";
      if (error) {
        setPasswordManagementMessage(error);
        return false;
      }
      administrativeSecurityState.accounts[username].passwordHash = newHash;
      administrativeSecurityState.accounts[username].passwordHashHistory = [newHash, ...(administrativeSecurityState.accounts[username].passwordHashHistory || [])].slice(0, 5);
      administrativeSecurityState.accounts[username].failedAttempts = 0;
      administrativeSecurityState.accounts[username].locked = false;
      appendPasswordAudit(username, "Change Password", remark);
      formElement.reset();
      document.querySelector("#passwordAccountSelect").value = username;
      setPasswordManagementMessage("Password changed successfully. / เปลี่ยนรหัสผ่านเรียบร้อย", true);
      renderPasswordManagement();
      if (username === "admin") logoutUser();
      return true;
    }

    async function resetManagedPassword() {
      if (!requireSystemAdministrator()) return false;
      const formElement = document.querySelector("#passwordManagementForm");
      const form = new FormData(formElement);
      const username = String(form.get("username") || "");
      if (!await verifyAdminCurrentPassword(String(form.get("adminCurrentPassword") || ""))) {
        setPasswordManagementMessage("Admin Current Password is incorrect.");
        return false;
      }
      const previousHashes = credentialRecord(username)?.passwordHashHistory || [];
      administrativeSecurityState.accounts[username] = { passwordHash: defaultPasswordHashes[username], passwordHashHistory: [defaultPasswordHashes[username], ...previousHashes].filter((value, index, rows) => rows.indexOf(value) === index).slice(0, 5), failedAttempts: 0, locked: false };
      appendPasswordAudit(username, "Reset to Default Password", form.get("remark"));
      setPasswordManagementMessage("Default password restored. / คืนค่ารหัสผ่านเริ่มต้นแล้ว", true);
      renderPasswordManagement();
      if (username === "admin") logoutUser();
      return true;
    }

    async function unlockManagedAccount() {
      if (!requireSystemAdministrator()) return false;
      const formElement = document.querySelector("#passwordManagementForm");
      const form = new FormData(formElement);
      const username = String(form.get("username") || "");
      if (!await verifyAdminCurrentPassword(String(form.get("adminCurrentPassword") || ""))) {
        setPasswordManagementMessage("Admin Current Password is incorrect.");
        return false;
      }
      const account = credentialRecord(username);
      account.locked = false;
      account.failedAttempts = 0;
      appendPasswordAudit(username, "Unlock Account", form.get("remark"));
      setPasswordManagementMessage("Account unlocked. / ปลดล็อกบัญชีแล้ว", true);
      renderPasswordManagement();
      return true;
    }

    function requestAttachmentNames(request) {
      return (request.attachments || []).map(item => item.originalFileName || item.fileName).filter(Boolean).join(", ") || "-";
    }

    function renderDueDateApprovalQueue() {
      if (!isSystemAdministrator()) {
        requireSystemAdministrator("Access Denied: Only the System Administrator can approve and apply due date adjustments.");
        return;
      }
      const queue = document.querySelector("#dueApprovalQueueRows");
      if (queue) queue.innerHTML = dueDateAdjustmentRequests.length ? dueDateAdjustmentRequests.slice().reverse().map(request => `
        <tr><td>${escapeHtml(request.requestId)}</td><td>${escapeHtml(request.contractId)}</td><td>${escapeHtml(request.requestedDue)}</td><td>${escapeHtml(request.requestedBy)}</td><td>${escapeHtml(request.status)}</td><td><button class="secondary-button" type="button" data-review-due-request="${escapeHtml(request.requestId)}">Review</button></td></tr>`).join("") : '<tr><td colspan="6">No Due Date requests</td></tr>';
      const history = document.querySelector("#dueAdjustmentHistoryRows");
      if (history) history.innerHTML = administrativeSecurityState.dueDateHistory.length
        ? administrativeSecurityState.dueDateHistory.map(item => `<tr><td>${escapeHtml(item.requestId)}</td><td>${escapeHtml(item.decision)}</td><td>${escapeHtml(item.finalDue || "-")}</td><td>${escapeHtml(item.admin)}</td><td>${escapeHtml(formatDateTime(item.decidedAt))}</td></tr>`).join("")
        : '<tr><td colspan="5">No Due Date adjustment history</td></tr>';
    }

    function openDueDateApprovalRequest(requestId) {
      if (!requireSystemAdministrator("Access Denied: Only the System Administrator can approve and apply due date adjustments.")) return false;
      const request = dueDateAdjustmentRequests.find(item => item.requestId === requestId);
      const contract = contracts.find(item => item.id === request?.contractId);
      if (!request || !contract) return false;
      document.querySelector("#adminDueRequestId").value = request.requestId;
      document.querySelector("#adminFinalDue").value = request.requestedDue || "";
      document.querySelector("#adminDueRequestSummary").innerHTML = `
        <section class="admin-summary-section"><h3>Current Due Date</h3><dl><div><dt>Original Due Date</dt><dd>${escapeHtml(contract.originalDue || request.currentDue || "-")}</dd></div><div><dt>Current Due Date</dt><dd>${escapeHtml(contract.due || request.currentDue || "-")}</dd></div><div><dt>Current SLA Balance</dt><dd>${escapeHtml(String(contract.balance ?? request.currentBalance ?? 0))}</dd></div><div><dt>Current Alert</dt><dd>${escapeHtml(formatContractStatusLabel(contract.status || contract.alert || ""))}</dd></div></dl></section>
        <section class="admin-summary-section"><h3>User Request</h3><dl><div><dt>Requested Due Date</dt><dd>${escapeHtml(request.requestedDue)}</dd></div><div><dt>Request Type</dt><dd>${escapeHtml(request.requestType || "Extension")}</dd></div><div><dt>Request Reason</dt><dd>${escapeHtml(request.reason)}</dd></div><div><dt>Requested By</dt><dd>${escapeHtml(request.requestedBy)}</dd></div><div><dt>Requested Date</dt><dd>${escapeHtml(formatDateTime(request.requestedAt))}</dd></div><div><dt>Attachment</dt><dd>${escapeHtml(requestAttachmentNames(request))}</dd></div></dl></section>
        <section class="admin-summary-section"><h3>Admin Decision</h3><dl><div><dt>Status</dt><dd>${escapeHtml(request.status)}</dd></div><div><dt>Decision</dt><dd>Awaiting admin</dd></div></dl></section>`;
      document.querySelector("#dueApprovalDecisionForm").hidden = false;
      return true;
    }

    function dueDateStatusFor(finalDue) {
      const today = todayInputValue();
      if (finalDue < today) return "Red >>R=Overdue";
      if (finalDue === today) return "Yellow >>Y=Delayed";
      return "Green >>G=On Track";
    }

    function buildDueDecisionEmailDraft(request, contract) {
      const resultCopy = {
        "Approved and Applied": { en: "Approved", th: "อนุมัติ", subject: "Approved" },
        "Rejected": { en: "Not Approved", th: "ไม่อนุมัติ", subject: "Not Approved" },
        "More Information Requested": { en: "More Information Required", th: "ขอข้อมูลเพิ่มเติม", subject: "More Information Required" }
      }[request.status] || { en: request.status, th: request.status, subject: request.status };
      const finalDue = request.finalDue || contract.due || request.currentDue || "-";
      const body = [
        `Dear ${contract.owner || request.requestedBy || "Requester"},`,
        "",
        "Your Due Date adjustment request has been reviewed.",
        `Request ID: ${request.requestId}`,
        `Contract ID: ${contract.id}`,
        `Contract Name: ${contract.name}`,
        `Requested Due Date: ${request.requestedDue}`,
        `Decision: ${resultCopy.en}`,
        `Final Due Date: ${finalDue}`,
        `Decision Reason: ${request.decisionReason || "-"}`,
        ...(request.differentDateReason ? [`Reason for Different Final Date: ${request.differentDateReason}`] : []),
        ...(request.adminNote ? [`Admin Note: ${request.adminNote}`] : []),
        "",
        `เรียน ${contract.owner || request.requestedBy || "ผู้ขอ"}`,
        "",
        "คำขอปรับวันครบกำหนดได้รับการพิจารณาแล้ว",
        `เลขที่คำขอ: ${request.requestId}`,
        `รหัสสัญญา: ${contract.id}`,
        `ชื่อสัญญา: ${contract.name}`,
        `วันที่ขอปรับ: ${request.requestedDue}`,
        `ผลการพิจารณา: ${resultCopy.th}`,
        `วันครบกำหนดสุดท้าย: ${finalDue}`,
        `เหตุผลการพิจารณา: ${request.decisionReason || "-"}`,
        ...(request.differentDateReason ? [`เหตุผลที่กำหนดวันแตกต่าง: ${request.differentDateReason}`] : []),
        ...(request.adminNote ? [`หมายเหตุจากผู้ดูแลระบบ: ${request.adminNote}`] : []),
        "",
        "Contract Tracking System"
      ].join("\n");
      return {
        to: request.requestedByEmail || emailForPerson(contract.owner) || emailForPerson(request.requestedBy) || "",
        subject: `[Contract Tracking] Due Date Request ${resultCopy.subject}: ${request.requestId} - ${contract.id}`,
        body,
        attachments: [],
        contractId: contract.id,
        action: `Due Date ${request.status}`
      };
    }

    function openDueDecisionEmailPopup(draft) {
      const modal = document.querySelector("#dueDecisionEmailModal");
      if (!modal) return;
      document.querySelector("#dueDecisionEmailTo").value = draft.to || "";
      document.querySelector("#dueDecisionEmailSubject").value = draft.subject || "";
      document.querySelector("#dueDecisionEmailBody").value = draft.body || "";
      modal.dataset.contractId = draft.contractId || "";
      modal.dataset.action = draft.action || "";
      document.querySelector("#dueDecisionEmailSubtitle").textContent = draft.to
        ? `Requester email linked automatically: ${draft.to}`
        : "Requester email not found. Enter an email before sending. / ไม่พบอีเมลผู้ขอ กรุณากรอกก่อนส่ง";
      modal.classList.add("show");
      modal.setAttribute("aria-hidden", "false");
      setTimeout(() => document.querySelector(draft.to ? "#sendDueDecisionEmailBtn" : "#dueDecisionEmailTo")?.focus(), 0);
    }

    function closeDueDecisionEmailPopup() {
      const modal = document.querySelector("#dueDecisionEmailModal");
      modal?.classList.remove("show");
      modal?.setAttribute("aria-hidden", "true");
    }

    function currentDueDecisionEmailDraft() {
      const modal = document.querySelector("#dueDecisionEmailModal");
      return {
        to: String(document.querySelector("#dueDecisionEmailTo")?.value || "").trim().toLowerCase(),
        subject: document.querySelector("#dueDecisionEmailSubject")?.value || "",
        body: document.querySelector("#dueDecisionEmailBody")?.value || "",
        attachments: [],
        contractId: modal?.dataset.contractId || "",
        action: modal?.dataset.action || ""
      };
    }

    async function sendDueDecisionEmail() {
      const draft = currentDueDecisionEmailDraft();
      if (!updateEmailPattern(draft.to)) {
        document.querySelector("#dueDecisionEmailTo")?.focus();
        showToast("Enter a valid requester email. / กรุณากรอกอีเมลผู้ขอให้ถูกต้อง");
        return false;
      }
      const button = document.querySelector("#sendDueDecisionEmailBtn");
      try {
        if (button) { button.disabled = true; button.textContent = "Sending..."; }
        await sendStatusEmailViaEndpoint(draft);
        showToast("Due Date decision email sent. / ส่งอีเมลแจ้งผลแล้ว");
        closeDueDecisionEmailPopup();
        return true;
      } catch (error) {
        const message = error?.message || "Cannot send email.";
        document.querySelector("#dueDecisionEmailSubtitle").textContent = message;
        showToast(message);
        return false;
      } finally {
        if (button) { button.disabled = false; button.innerHTML = "<span>✉</span> Send Email"; }
      }
    }

    async function copyDueDecisionEmailDraft() {
      const draft = currentDueDecisionEmailDraft();
      const text = `To: ${draft.to}\nSubject: ${draft.subject}\n\n${draft.body}`;
      try {
        await navigator.clipboard.writeText(text);
      } catch (error) {
        const helper = document.createElement("textarea");
        helper.value = text;
        helper.style.position = "fixed";
        helper.style.opacity = "0";
        document.body.appendChild(helper);
        helper.select();
        document.execCommand("copy");
        helper.remove();
      }
      showToast("Email copied");
    }

    async function applyDueDateDecision(formElement) {
      if (!requireSystemAdministrator("Access Denied: Only the System Administrator can approve and apply due date adjustments.")) return false;
      const form = new FormData(formElement);
      const request = dueDateAdjustmentRequests.find(item => item.requestId === form.get("requestId"));
      const contract = contracts.find(item => item.id === request?.contractId);
      if (!request || !contract || request.status === "Approved and Applied") return false;
      const decision = String(form.get("decision") || "");
      const finalDue = String(form.get("finalDue") || "");
      const decisionReason = String(form.get("decisionReason") || "").trim();
      const differentDateReason = String(form.get("differentDateReason") || "").trim();
      if (!await verifyAdminCurrentPassword(String(form.get("adminPassword") || ""))) {
        showToast("Admin password confirmation is incorrect.");
        return false;
      }
      if (!decisionReason) { showToast("Admin Decision Reason is required."); return false; }
      if (decision === "Approve and Apply" && !finalDue) { showToast("Final Due Date is required."); return false; }
      if (decision === "Approve and Apply" && finalDue !== request.requestedDue && !differentDateReason) {
        showToast("Reason for Different Final Date is required.");
        return false;
      }
      if (decision === "Approve and Apply") {
        contract.originalDue = contract.originalDue || request.currentDue || contract.due;
        contract.due = finalDue;
        contract.balance = finalDue >= todayInputValue() ? calculateDaysOnHand(todayInputValue(), finalDue) : 0;
        contract.alert = dueDateStatusFor(finalDue);
        contract.status = contract.alert;
        contract.lastModifiedBy = "admin";
        contract.lastModifiedAt = localIsoDateTime();
        request.status = "Approved and Applied";
        addLogRecord({ contractId: contract.id, cycle: contract.cycle || 1, station: "From admin >> To " + (request.requestedBy || contract.owner), from: "admin", to: request.requestedBy || contract.owner, inDate: todayInputValue(), outDate: todayInputValue(), sla: 0, delayReason: decisionReason, action: "Due Date Approved and Applied", approvalType: decision, approvalConditions: differentDateReason, correctiveAction: `Final Due Date ${finalDue}` });
      } else if (decision === "Reject Request") {
        request.status = "Rejected";
      } else {
        request.status = "More Information Requested";
      }
      request.decision = decision;
      request.finalDue = decision === "Approve and Apply" ? finalDue : "";
      request.decisionReason = decisionReason;
      request.differentDateReason = differentDateReason;
      request.adminNote = String(form.get("adminNote") || "").trim();
      request.decidedBy = "admin";
      request.decidedAt = localIsoDateTime();
      request.requesterNotification = `${request.requestId}: ${request.status}`;
      const requesterEmailDraft = buildDueDecisionEmailDraft(request, contract);
      administrativeSecurityState.dueDateHistory.unshift({ requestId: request.requestId, contractId: request.contractId, decision, finalDue: request.finalDue, decisionReason, differentDateReason, admin: "admin", decidedAt: request.decidedAt });
      persistAdministrativeSecurityState();
      delete contract.pendingDueDateRequest;
      refreshDashboardDataFromContracts();
      saveContractsDatabase();
      renderAll();
      formElement.reset();
      formElement.hidden = true;
      showToast(`${request.requestId} · ${request.status}`);
      openDueDecisionEmailPopup(requesterEmailDraft);
      return true;
    }

    function renderAdministrativeControls() {
      const admin = isSystemAdministrator();
      document.querySelectorAll("[data-admin-only]").forEach(node => { node.hidden = !admin; });
      const passwordMenu = document.querySelector("#adminPasswordMenu");
      if (passwordMenu) passwordMenu.hidden = !admin;
      if (!admin) return;
      renderPasswordManagement();
      renderDueDateApprovalQueue();
    }

    function setupAdministrativeControls() {
      document.querySelector("#passwordAccountSelect")?.addEventListener("change", renderPasswordManagement);
      document.querySelector("#passwordManagementForm")?.addEventListener("submit", event => { event.preventDefault(); changeManagedPassword(event.currentTarget); });
      document.querySelector("#resetPasswordBtn")?.addEventListener("click", resetManagedPassword);
      document.querySelector("#unlockAccountBtn")?.addEventListener("click", unlockManagedAccount);
      document.querySelector("#dueApprovalQueueRows")?.addEventListener("click", event => {
        const button = event.target.closest("[data-review-due-request]");
        if (button) openDueDateApprovalRequest(button.dataset.reviewDueRequest);
      });
      document.querySelector("#dueApprovalDecisionForm")?.addEventListener("submit", event => { event.preventDefault(); applyDueDateDecision(event.currentTarget); });
      document.querySelector("#cancelDueAdminDecision")?.addEventListener("click", () => { document.querySelector("#dueApprovalDecisionForm").hidden = true; });
      document.querySelector("#sendDueDecisionEmailBtn")?.addEventListener("click", sendDueDecisionEmail);
      document.querySelector("#copyDueDecisionEmailBtn")?.addEventListener("click", copyDueDecisionEmailDraft);
      document.querySelector("#closeDueDecisionEmailBtn")?.addEventListener("click", closeDueDecisionEmailPopup);
      document.querySelector("#closeDueDecisionEmailModal")?.addEventListener("click", closeDueDecisionEmailPopup);
      document.querySelector("#adminPasswordMenu")?.addEventListener("click", () => {
        if (!requireSystemAdministrator()) return;
        setView("admin");
        document.querySelector("#passwordManagementPanel")?.scrollIntoView({ behavior: "smooth", block: "start" });
      });
      renderAdministrativeControls();
    }
'''
    html = html.replace("    window.appState = {", admin_security_js + "\n    window.appState = {", 1)

    html = html.replace(
        '''      setupCsvDatabaseControls();
      setupMasterDataControls();
      loadContractsDatabase();''',
        '''      setupCsvDatabaseControls();
      setupMasterDataControls();
      setupAdministrativeControls();
      loadContractsDatabase();''',
        1,
    )
    html = html.replace(
        '''      renderMasterData();
      syncNavCounts();''',
        '''      renderMasterData();
      renderAdministrativeControls();
      syncNavCounts();''',
        1,
    )
    html = html.replace(
        '''      if (newContractBtn) newContractBtn.hidden = !canUseUserCaseAction();
      if (!canUseUserCaseAction()) document.querySelector("#user")?.classList.remove("active");''',
        '''      if (newContractBtn) newContractBtn.hidden = !canUseUserCaseAction();
      if (!canUseUserCaseAction()) document.querySelector("#user")?.classList.remove("active");
      renderAdministrativeControls();''',
        1,
    )
    html = html.replace(
        '''    function isAdmin() {
      return currentUser?.role === "admin";
    }''',
        '''    function isAdmin() {
      return currentUser?.username === "admin" && currentUser?.role === "admin" && Number(currentUser?.level) === 4;
    }''',
        1,
    )
    html = html.replace(
        '''    function canManageMasterData() {
      return hasAccessLevel(4);
    }''',
        '''    function canManageMasterData() {
      return currentUser?.username === "admin" && currentUser?.role === "admin" && Number(currentUser?.level) === 4;
    }''',
        1,
    )
    html = html.replace(
        '''        requestId: `DDR-${String(dueDateAdjustmentRequests.length + 1).padStart(3, "0")}`,''',
        '''        requestId: `DDR-${String(Math.max(0, ...dueDateAdjustmentRequests.map(item => Number(String(item.requestId || "").replace(/\\D/g, "")) || 0)) + 1).padStart(3, "0")}`,''',
        1,
    )
    html = html.replace(
        '''    async function saveMasterDataFromUi() {
      if (!normalizeMasterDataFromUi()) return;''',
        '''    async function saveMasterDataFromUi() {
      if (!requireSystemAdministrator()) return;
      if (!normalizeMasterDataFromUi()) return;''',
        1,
    )
    html = html.replace(
        '''    function addMasterRow(kind) {
      if (!normalizeMasterDataFromUi()) return;''',
        '''    function addMasterRow(kind) {
      if (!requireSystemAdministrator()) return;
      if (!normalizeMasterDataFromUi()) return;''',
        1,
    )
    html = html.replace(
        '''      get contracts() { return contracts; },
      get logRecords() { return logRecords; },''',
        '''      get contracts() { return contracts.map(item => ({ ...item })); },
      get logRecords() { return logRecords.map(row => Array.isArray(row) ? [...row] : row); },''',
        1,
    )
    html = html.replace(
        'get dueDateAdjustmentRequests() { return dueDateAdjustmentRequests; },',
        'get dueDateAdjustmentRequests() { return dueDateAdjustmentRequests.map(item => ({ ...item, attachments: (item.attachments || []).map(file => ({ ...file, base64: undefined })) })); },',
        1,
    )

    contract_headers = [
        "Contract ID",
        "Contract Name",
        "Department / Restaurant",
        "Contract Owner",
        "Type of Contract",
        "Vendor / Counter party",
        "Stage",
        "Cycle",
        "Returns",
        "Status Update",
        "Station",
        "Station Owner",
        "Due Date",
        "System Due Date",
        "Work Type",
        "Total SLA",
        "Days Used",
        "Days on Hand",
        "Balance",
        "Alert",
        "Remark",
        "Access Level",
        "Visibility",
        "Category",
    ]
    contract_csv_rows = []
    for contract in contracts:
        station_to = contract.get("station", "").split(">> To ")[-1].strip() if ">> To " in contract.get("station", "") else ""
        contract_csv_rows.append({
            "Contract ID": contract.get("id", ""),
            "Contract Name": contract.get("name", ""),
            "Department / Restaurant": contract.get("department", ""),
            "Contract Owner": contract.get("owner", ""),
            "Type of Contract": contract.get("type", ""),
            "Vendor / Counter party": contract.get("vendor", ""),
            "Stage": contract.get("stage", ""),
            "Cycle": contract.get("cycle", ""),
            "Returns": contract.get("returns", ""),
            "Status Update": contract.get("status", ""),
            "Station": contract.get("station", ""),
            "Station Owner": station_to,
            "Due Date": contract.get("due", ""),
            "System Due Date": contract.get("systemDue", ""),
            "Work Type": contract.get("workType", ""),
            "Total SLA": contract.get("totalSla", ""),
            "Days Used": contract.get("used", ""),
            "Days on Hand": contract.get("days", contract.get("used", "")),
            "Balance": contract.get("balance", ""),
            "Alert": contract.get("alert", ""),
            "Remark": contract.get("remark", ""),
            "Access Level": contract.get("accessLevel", ""),
            "Visibility": contract.get("visibility", ""),
            "Category": contract.get("category", ""),
        })

    log_headers = [
        "Contract ID",
        "Log No",
        "Cycle",
        "Log View",
        "From",
        "To",
        "In",
        "Out",
        "SLA",
        "Days on Hand",
        "Alert",
        "Delay Reason",
        "Action",
        "Action Reason",
        "Approval",
        "Corrective Action",
        "Action Reason Type",
        "Action Reason Detail",
        "Approval Type",
        "Approval Conditions",
        "Corrective Action Detail",
        "Updated By",
        "Updated Date and Time",
        "Action Code",
        "Action Name TH",
        "Action Name EN",
        "Action Description TH",
        "Action Description EN",
        "Action SLA",
        "Action Reason Type TH",
        "Action Reason Type EN",
        "Attachments",
        "CC Recipients",
    ]
    log_csv_rows = [
        {header: row[index] if index < len(row) else "" for index, header in enumerate(log_headers)}
        for row in log_records
    ]

    type_headers = [
        "Contract Classification",
        "Type of Contract",
        "Sub Type of Contract",
        "Fixed SLA (Working Days)",
        "Standard SLA Version",
        "Active",
        "Category",
        "Description / คำอธิบาย",
    ]

    OUTPUT_HTML.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_HTML.write_text(html, encoding="utf-8")
    write_csv(OUTPUT_CONTRACTS_CSV, contract_csv_rows, contract_headers)
    write_csv(OUTPUT_LOGS_CSV, log_csv_rows, log_headers)
    write_csv(OUTPUT_TYPE_MASTER_CSV, type_rows, type_headers)
    write_csv(OUTPUT_DEPARTMENT_MASTER_CSV, department_master_rows, ["Department / Restaurant", "Department Code", "Department Data Version", "Active"])
    write_csv(OUTPUT_PEOPLE_MASTER_CSV, people_master_rows, ["company", "department", "name", "email", "active"])
    write_csv(OUTPUT_CONTRACT_TEMPLATE_CSV, contract_catalog, ["classification", "typeGroup", "subType", "name", "selectionLabel", "sourceRow", "type", "workType", "contractId", "accessLevel", "category", "department", "vendor", "group", "fixedSla", "slaVersion", "remark", "active"])
    write_csv(
        OUTPUT_ACTION_SLA_MASTER_CSV,
        action_sla_master_rows,
        ["Action", "Description / รายละเอียด", "Fixed SLA (Working Days)", "SLA Rule / วิธีนับ", "Action Data Version", "Active"],
    )
    write_csv(
        OUTPUT_CONTRACT_TYPE_MASTER_V2_CSV,
        contract_type_master_v2_rows,
        [
            "Contract Classification EN",
            "Contract Classification TH",
            "Type of Contract EN",
            "Type of Contract TH",
            "Sub Type of Contract EN",
	            "Sub Type of Contract TH",
	            "Type Sort Order",
	            "Sub Type Sort Order",
	            "Standard SLA",
	        ],
	    )
    OUTPUT_ATTACHMENT_APPS_SCRIPT.write_text(
        f"""const DEFAULT_FOLDER_ID = "{ATTACHMENT_CLOUD_FOLDER_ID}";
const DEFAULT_DATABASE_FOLDER_ID = "{DRIVE_FOLDER_ID}";
const DEFAULT_BACKUP_FOLDER_ID = "{DRIVE_FOLDER_ID}";
const EMAIL_SENDER_NAME = "T23 Contract Tracking";

function doPost(e) {{
  try {{
    const payload = JSON.parse((e.postData && e.postData.contents) || "{{}}");
    const mode = payload.mode || (payload.to ? "sendStatusEmail" : "uploadAttachment");
    if (mode === "sendStatusEmail") return sendStatusEmail_(payload);
    if (mode === "saveDriveDatabase") return saveDriveDatabase_(payload);
    if (mode === "backupDriveDatabase") return backupDriveDatabase_(payload);
    if (mode === "installDailyBackup") return installDailyBackupTrigger_(payload);
    return jsonResponse({{ success: true, files: [saveAttachment_(payload)] }});
  }} catch (error) {{
    return jsonResponse({{ success: false, error: errorMessage_(error) }});
  }}
}}

function doGet(e) {{
  const params = (e && e.parameter) || {{}};
  const callback = String(params.callback || "").trim();
  try {{
    if (params.mode === "loadDriveDatabase") return jsonpResponse(loadDriveDatabase_(params), callback);
    return jsonpResponse({{
      success: true,
      message: "T23 attachment upload, status email, and Drive database endpoint is running.",
      folderId: DEFAULT_FOLDER_ID,
      databaseFolderId: DEFAULT_DATABASE_FOLDER_ID,
      backupFolderId: DEFAULT_BACKUP_FOLDER_ID
    }}, callback);
  }} catch (error) {{
    return jsonpResponse({{ success: false, error: errorMessage_(error) }}, callback);
  }}
}}

function loadDriveDatabase_(params) {{
  const folder = DriveApp.getFolderById(params.folderId || DEFAULT_FOLDER_ID);
  return {{
    success: true,
    folderId: folder.getId(),
    loadedAt: new Date().toISOString(),
    contractsCsvText: readTextFileByName_(folder, params.contractsCsv || "tracking_contracts_contracts_db.csv"),
    logsCsvText: readTextFileByName_(folder, params.logsCsv || "tracking_contracts_log_db.csv"),
    typeMasterCsvText: readTextFileByName_(folder, params.typeMasterCsv || "tracking_contracts_type_master_db.csv"),
	    departmentMasterCsvText: readTextFileByName_(folder, params.departmentMasterCsv || "tracking_contracts_department_master_db.csv"),
	    peopleMasterCsvText: readTextFileByName_(folder, params.peopleMasterCsv || "tracking_contracts_people_master_db.csv"),
	    contractTemplateCsvText: readTextFileByName_(folder, params.contractTemplateCsv || "tracking_contracts_contract_template_master_db.csv"),
	    actionSlaCsvText: readTextFileByName_(folder, params.actionSlaCsv || "tracking_contracts_action_sla_master_db.csv")
  }};
}}

function saveDriveDatabase_(payload) {{
  const folder = DriveApp.getFolderById(payload.folderId || DEFAULT_FOLDER_ID);
  const files = {{
    contracts: upsertTextFileByName_(folder, payload.contractsCsv || "tracking_contracts_contracts_db.csv", payload.contractsCsvText || ""),
    logs: upsertTextFileByName_(folder, payload.logsCsv || "tracking_contracts_log_db.csv", payload.logsCsvText || ""),
    typeMaster: upsertTextFileByName_(folder, payload.typeMasterCsv || "tracking_contracts_type_master_db.csv", payload.typeMasterCsvText || ""),
	    departments: upsertTextFileByName_(folder, payload.departmentMasterCsv || "tracking_contracts_department_master_db.csv", payload.departmentMasterCsvText || ""),
	    people: upsertTextFileByName_(folder, payload.peopleMasterCsv || "tracking_contracts_people_master_db.csv", payload.peopleMasterCsvText || ""),
	    contractTemplates: upsertTextFileByName_(folder, payload.contractTemplateCsv || "tracking_contracts_contract_template_master_db.csv", payload.contractTemplateCsvText || ""),
	    actionSla: upsertTextFileByName_(folder, payload.actionSlaCsv || "tracking_contracts_action_sla_master_db.csv", payload.actionSlaCsvText || "")
  }};
  return jsonResponse({{
    success: true,
    saved: true,
    savedAt: new Date().toISOString(),
    folderId: folder.getId(),
    files: files
  }});
}}

function backupDriveDatabase_(payload) {{
  const sourceFolder = DriveApp.getFolderById(payload.folderId || DEFAULT_DATABASE_FOLDER_ID);
  const backupFolder = DriveApp.getFolderById(payload.backupFolderId || DEFAULT_BACKUP_FOLDER_ID);
  const stamp = Utilities.formatDate(new Date(), "Etc/UTC", "yyyyMMdd-HHmmss");
  const prefix = payload.prefix || "daily_backup";
  const files = {{
    contracts: backupTextFileByName_(sourceFolder, backupFolder, payload.contractsCsv || "tracking_contracts_contracts_db.csv", stamp, prefix),
    logs: backupTextFileByName_(sourceFolder, backupFolder, payload.logsCsv || "tracking_contracts_log_db.csv", stamp, prefix),
    typeMaster: backupTextFileByName_(sourceFolder, backupFolder, payload.typeMasterCsv || "tracking_contracts_type_master_db.csv", stamp, prefix),
    departments: backupTextFileByName_(sourceFolder, backupFolder, payload.departmentMasterCsv || "tracking_contracts_department_master_db.csv", stamp, prefix),
    people: backupTextFileByName_(sourceFolder, backupFolder, payload.peopleMasterCsv || "tracking_contracts_people_master_db.csv", stamp, prefix),
    contractTemplates: backupTextFileByName_(sourceFolder, backupFolder, payload.contractTemplateCsv || "tracking_contracts_contract_template_master_db.csv", stamp, prefix),
    actionSla: backupTextFileByName_(sourceFolder, backupFolder, payload.actionSlaCsv || "tracking_contracts_action_sla_master_db.csv", stamp, prefix)
  }};
  return jsonResponse({{
    success: true,
    backedUp: true,
    backedUpAt: new Date().toISOString(),
    sourceFolderId: sourceFolder.getId(),
    backupFolderId: backupFolder.getId(),
    files: files
  }});
}}

function runDailyBackup() {{
  return backupDriveDatabase_({{
    folderId: DEFAULT_DATABASE_FOLDER_ID,
    backupFolderId: DEFAULT_BACKUP_FOLDER_ID,
    prefix: "daily_backup"
  }});
}}

function installDailyBackupTrigger_() {{
  ScriptApp.getProjectTriggers().forEach(function(trigger) {{
    if (trigger.getHandlerFunction && trigger.getHandlerFunction() === "runDailyBackup") {{
      ScriptApp.deleteTrigger(trigger);
    }}
  }});
  const trigger = ScriptApp.newTrigger("runDailyBackup")
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
  return jsonResponse({{
    success: true,
    installed: true,
    handlerFunction: trigger.getHandlerFunction(),
    backupFolderId: DEFAULT_BACKUP_FOLDER_ID
  }});
}}

function readTextFileByName_(folder, fileName) {{
  const files = folder.getFilesByName(fileName);
  if (!files.hasNext()) return "";
  return files.next().getBlob().getDataAsString("UTF-8").replace(/^\\uFEFF/, "");
}}

function backupTextFileByName_(sourceFolder, backupFolder, fileName, stamp, prefix) {{
  const name = cleanFileName_(fileName || "database.csv");
  const files = sourceFolder.getFilesByName(name);
  if (!files.hasNext()) {{
    return {{
      sourceName: name,
      skipped: true,
      reason: "Source file not found"
    }};
  }}
  const sourceFile = files.next();
  const backupName = cleanFileName_([prefix, stamp, name].join("_"));
  const backupFile = backupFolder.createFile(backupName, sourceFile.getBlob().getDataAsString("UTF-8"), MimeType.CSV);
  backupFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {{
    id: backupFile.getId(),
    sourceName: name,
    name: backupFile.getName(),
    url: backupFile.getUrl(),
    downloadUrl: "https://drive.google.com/uc?export=download&id=" + backupFile.getId()
  }};
}}

function upsertTextFileByName_(folder, fileName, text) {{
  const name = cleanFileName_(fileName || "database.csv");
  const content = String(text || "");
  const files = folder.getFilesByName(name);
  const file = files.hasNext()
    ? files.next().setContent(content)
    : folder.createFile(name, content, MimeType.CSV);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {{
    id: file.getId(),
    name: file.getName(),
    url: file.getUrl(),
    downloadUrl: "https://drive.google.com/uc?export=download&id=" + file.getId()
  }};
}}

function sendStatusEmail_(payload) {{
  const to = String(payload.to || "").trim();
  if (!to) throw new Error("Missing recipient email.");
  if (!isValidEmailList_(to)) throw new Error("Invalid recipient email.");

  const files = (payload.attachments || []).map(function(attachment) {{
    return saveAttachment_(Object.assign({{ folderId: payload.folderId }}, attachment));
  }});
  const mailAttachments = buildMailAttachmentBlobs_(payload.attachments || []);
  const body = buildEmailBody_(payload.body || "", files, payload.folderUrl || "");
  const cc = normalizeCc_(payload.cc || payload.ccText || "");
  if (cc && !isValidEmailList_(cc)) throw new Error("Invalid CC email.");
  const options = {{
    to: to,
    subject: payload.subject || "Contract Status Update",
    body: body,
    name: EMAIL_SENDER_NAME,
    htmlBody: buildHtmlBody_(body)
  }};
  if (cc) options.cc = cc;
  if (mailAttachments.length) options.attachments = mailAttachments;

  MailApp.sendEmail(options);
  return jsonResponse({{
    success: true,
    sent: true,
    sentAt: new Date().toISOString(),
    to: to,
    cc: cc,
    files: files,
    attachedFiles: mailAttachments.map(function(blob) {{ return blob.getName(); }})
  }});
}}

function normalizeCc_(value) {{
  const list = Array.isArray(value) ? value : String(value || "").split(/[;,\\n]+/);
  const seen = {{}};
  return list
    .map(function(item) {{
      return String(item && item.email ? item.email : item || "").trim();
    }})
    .filter(function(email) {{
      if (!email || seen[email.toLowerCase()]) return false;
      seen[email.toLowerCase()] = true;
      return true;
    }})
    .join(", ");
}}

function isValidEmailList_(value) {{
  const emails = String(value || "").split(/[;,\\n]+/).map(function(item) {{
    return item.trim();
  }}).filter(Boolean);
  if (!emails.length) return false;
  return emails.every(function(email) {{
    return /^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/.test(email);
  }});
}}

function buildMailAttachmentBlobs_(attachments) {{
  return attachments
    .filter(function(attachment) {{ return attachment && attachment.base64; }})
    .map(function(attachment) {{
      const fileName = cleanFileName_(attachment.originalFileName || attachment.fileName || "attachment");
      const mimeType = attachment.mimeType || "application/octet-stream";
      return Utilities.newBlob(Utilities.base64Decode(attachment.base64), mimeType, fileName);
    }});
}}

function saveAttachment_(payload) {{
  const existingUrl = reusableDriveFileUrl_(payload.cloudUrl || payload.url || "") || "";
  if (existingUrl) {{
    return {{
      id: payload.cloudFileId || "",
      fileName: payload.fileName || payload.originalFileName || "attachment",
      originalFileName: payload.originalFileName || payload.fileName || "attachment",
      mimeType: payload.mimeType || "",
      fileSize: payload.fileSize || 0,
      url: existingUrl,
      downloadUrl: payload.downloadUrl || existingUrl,
      reused: true
    }};
  }}

  const base64 = payload.base64 || "";
  if (!base64) {{
    return {{
      id: "",
      fileName: payload.fileName || payload.originalFileName || "attachment",
      originalFileName: payload.originalFileName || payload.fileName || "attachment",
      mimeType: payload.mimeType || "",
      fileSize: payload.fileSize || 0,
      url: "",
      downloadUrl: "",
      skipped: true
    }};
  }}

  const folder = DriveApp.getFolderById(payload.folderId || DEFAULT_FOLDER_ID);
  const originalFileName = payload.originalFileName || payload.fileName || "attachment";
  const fileName = cleanFileName_(payload.fileName || originalFileName);
  const mimeType = payload.mimeType || "application/octet-stream";
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {{
    id: file.getId(),
    fileName: file.getName(),
    originalFileName: originalFileName,
    mimeType: mimeType,
    fileSize: payload.fileSize || blob.getBytes().length,
    url: file.getUrl(),
    downloadUrl: "https://drive.google.com/uc?export=download&id=" + file.getId(),
    reused: false
  }};
}}

function buildEmailBody_(baseBody, files, folderUrl) {{
  const usableFiles = files.filter(function(file) {{ return file.url; }});
  const lines = [String(baseBody || "")];
  if (usableFiles.length) {{
    lines.push("", "Cloud Attachments / ไฟล์แนบบน Cloud");
    usableFiles.forEach(function(file, index) {{
      lines.push(
        String(index + 1) + ". " + (file.originalFileName || file.fileName),
        "   Open: " + file.url,
        "   Download: " + (file.downloadUrl || file.url)
      );
    }});
  }}
  return lines.join("\\n");
}}

function buildHtmlBody_(body) {{
  return String(body || "")
    .split("\\n")
    .map(function(line) {{
      return escapeHtml_(line).replace(/(https:\\/\\/[^\\s<>"']+)/g, function(url) {{
        return '<a href="' + url + '" target="_blank" style="color:#1155cc;text-decoration:underline;">' + url + '</a>';
      }});
    }})
    .join("<br>");
}}

function escapeHtml_(value) {{
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}}

function cleanFileName_(name) {{
  const cleaned = String(name || "attachment").replace(/[\\\\/:*?"<>|]+/g, "_").slice(0, 180);
  return cleaned || "attachment";
}}

function reusableDriveFileUrl_(url) {{
  const text = String(url || "").trim();
  if (!text) return "";
  if (/^https:\\/\\/drive\\.google\\.com\\/file\\/d\\//.test(text)) return text;
  if (/^https:\\/\\/drive\\.google\\.com\\/uc\\?/.test(text)) return text;
  if (/^https:\\/\\/docs\\.google\\.com\\//.test(text)) return text;
  return "";
}}

function errorMessage_(error) {{
  return String((error && error.message) || error || "Unknown error");
}}

function jsonResponse(data) {{
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}}

function jsonpResponse(data, callback) {{
  const safeCallback = String(callback || "").trim();
  if (safeCallback && /^[A-Za-z_$][\\w$]*(\\.[A-Za-z_$][\\w$]*)*$/.test(safeCallback)) {{
    return ContentService
      .createTextOutput(safeCallback + "(" + JSON.stringify(data) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }}
  return jsonResponse(data);
}}
""",
        encoding="utf-8",
    )
    OUTPUT_CODE.write_text(Path(__file__).read_text(encoding="utf-8"), encoding="utf-8")
    OUTPUT_README.write_text(
        "\n".join([
            "T23 Tracking Contracts CSV Database",
            f"Google Drive folder: {DRIVE_FOLDER_URL}",
            f"HTML: {OUTPUT_HTML.name}",
            f"Codex generator code: {OUTPUT_CODE.name}",
            f"Contracts CSV database: {OUTPUT_CONTRACTS_CSV.name}",
            f"Log CSV database: {OUTPUT_LOGS_CSV.name}",
            f"Contract type master CSV: {OUTPUT_TYPE_MASTER_CSV.name}",
            f"Department master CSV: {OUTPUT_DEPARTMENT_MASTER_CSV.name}",
            f"People master CSV: {OUTPUT_PEOPLE_MASTER_CSV.name}",
            f"Contract name template CSV: {OUTPUT_CONTRACT_TEMPLATE_CSV.name}",
            f"Attachment upload Apps Script: {OUTPUT_ATTACHMENT_APPS_SCRIPT.name}",
            f"Attachment Cloud folder: {ATTACHMENT_CLOUD_FOLDER_URL}",
            "",
            "The dashboard loads the CSV database from the shared Google Drive folder through the Apps Script Web App.",
            "Add Case / Update Status / Close Case save locally first, then sync the CSV database back to the shared Drive folder in the background.",
            "To enable real email sending, direct attachment upload, and Drive CSV sync, deploy tracking_contracts_attachment_upload_apps_script.js as a Google Apps Script Web App.",
            "Then paste the Web App URL into ATTACHMENT_UPLOAD_ENDPOINT in tracking_contracts_dashboard_codex_code.py and regenerate the HTML.",
        ]),
        encoding="utf-8",
    )
    print(OUTPUT_HTML)
    print(f"contracts={len(contracts)} action_sla={len(action_sla)} type_master={len(type_rows)}")


if __name__ == "__main__":
    main()
