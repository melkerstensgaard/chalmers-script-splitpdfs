# Detta är ett skript för skanningsarbetet som Melker och Theodor utförde sommaren 2026 för att digitalisera examensbevis.
# Det är konfigurerat på så vis att det bara fungerar för den serie det är avsett att bearbeta. 
# För att ta reda på hur det fungerar finns det separat dokumentation.

import os
import re
import sys
from collections import Counter
from concurrent.futures import ProcessPoolExecutor, as_completed
from tqdm import tqdm
from pypdf import PdfReader, PdfWriter
from openpyxl import Workbook

# Log > GUI
def emit_log(message, log_callback=None):
    if log_callback:
        log_callback(message)
    else:
        print(message)



# Start-page detection is deliberately score-based rather than a single exact match.
# The ordinary degree certificate first page is a sparse ceremonial cover page.
# Later pages are administrative/table/diploma-supplement pages and should be rejected.

START_PAGE_STRONG_AWARD_MARKERS = [
    "HAR AVLAGT",
    "HAS BEEN AWARDED THE DEGREE OF",
    "HAR GENOMGÅTT",
]

# Degree names/titles that commonly appear prominently on first pages.
# Keep these broad enough to support OCR-text variants and new programmes.
START_PAGE_DEGREE_MARKERS = [
    "CIVILINGENJÖRSEXAMEN",
    "ARKITEKTEXAMEN",
    "MASTER OF ARCHITECTURE",
    "MASTER OF SCIENCE",
    "BACHELOR OF SCIENCE",
    "HÖGSKOLEINGENJÖRSEXAMEN",
    "TEKNOLOGIE MASTEREXAMEN",
    "TEKNOLOGIE KANDIDATEXAMEN",
    "FILOSOFIE MASTEREXAMEN",
    "FILOSOFIE KANDIDATEXAMEN",
    "SJÖKAPTENSEXAMEN",
    "SJÖINGENJÖRSEXAMEN",
    "MASTEREXAMEN",
    "KANDIDATEXAMEN",
    "HÖGSKOLEEXAMEN",

    # Engineer / maritime programmes from F1A

    # Bygg
    "BYGGINGENJÖR",
    "BYGGINGENJÖRSEXAMEN",
    "BYGGINGENJÖRSLINJEN",
    "BYGGNADSINGENJÖR",
    "BYGGNADSINGENJÖRSEXAMEN",
    "BYGGNADSINGENJÖRSLINJEN",

    # Elektro
    "ELEKTROINGENJÖR",
    "ELEKTROINGENJÖRSEXAMEN",
    "ELEKTROINGENJÖRSLINJEN",

    # Kemi / Transport
    "KEMIINGENJÖR",
    "KEMIINGENJÖRSEXAMEN",
    "KEMIINGENJÖRSLINJEN",
    "TRANSPORTINGENJÖR",
    "TRANSPORTINGENJÖRSEXAMEN",
        "TRANSPORTINGENJÖRSLINJEN",

    # Maskin
    "MASKININGENJÖR",
    "MASKININGENJÖRSEXAMEN",
    "MASKININGENJÖRSLINJEN",

    # Drift
    "DRIFTINGENJÖR",
    "DRIFTINGENJÖRSEXAMEN",
    "DRIFTINGENJÖRSLINJEN",
    "DRIFTTEKNIKER",
    "DRIFTTEKNIKEREXAMEN",
    "DRIFTTEKNIKERLINJEN",

    # Sjöfart
    "SJÖKAPTEN",
    "SJÖKAPTENSEXAMEN",
    "SJÖKAPTENSLINJEN",

    "SJÖINGENJÖR",
    "SJÖINGENJÖRSEXAMEN",
    "SJÖINGENJÖRSLINJEN",

    "STYRMAN",
    "STYRMANSEXAMEN",
    "STYRMANSLINJEN",

    "SKEPPARE",
    "SKEPPAREXAMEN",
    "SKEPPARELINJEN",

    # YTH
    "YTH",
    "YTH-BYGGNADSTEKNIK",
    "BYGGNADSTEKNIK",
    "BYGGNADSTEKNIKEXAMEN",
    "BYGGNADSTEKNIKLINJEN",

    # Tekniker
    "MASKINTEKNIKER",
    "MASKINTEKNIKEREXAMEN",
    "MASKINTEKNIKERLINJEN",

    # Generic fallbacks
    "INGENJÖRSEXAMEN",
    "INGENJÖRSLINJEN",
    "UNIVERSITY CERTIFICATE IN CIVIL ENGINEERING",
]

# Weak/secondary indicators. These are not enough on their own because they can
# also occur on later pages. They only add support when award/degree evidence exists.
START_PAGE_WEAK_MARKERS = [
    "CHALMERS TEKNISKA HÖGSKOLA",
    "CHALMERS UNIVERSITY OF TECHNOLOGY GOTHENBURG SWEDEN",
    "CHALMERS UNIVERSITY OF TECHNOLOGY",
    "EXAMENSHANDLÄGGARE OFFICER OF DEGREE",
    "OFFICER OF DEGREE",
    "ON BEHALF OF THE PRESIDENT",
    "PÅ REKTORS VÄGNAR",
]

SPECIAL_START_PAGE_IDENTIFIERS = [
    "INTYG OM FULLGJORDA STUDIER",
    "TRANSPORT STYRELSEN",
]

# These strongly indicate that a page is NOT a first page.
# The patterns are intentionally focused on headings/phrases rather than single
# common words, so that dates on the cover page do not cause false rejection.
NON_START_PAGE_PATTERNS = [
    "DIPLOMA SUPPLEMENT",
    "THIS DIPLOMA SUPPLEMENT FOLLOWS THE MODEL",
    "EXAMENSBEVIS FÖR",
    "DIPLOMA FOR",
    "OBLIGATORISKA KURSER",
    "VALFRIA KURSER",
    "COMPULSORY COURSES",
    "ELECTIVE COURSES",
    "COURSE CODE",
    "KURSKOD",
    "IDENTIFICATION NUMBER",
    "STUDENT IDENTIFICATION NUMBER",
    "EXAMINATOR",
    "THESIS TITLE"
]

EXCEL_FILENAME = "index.xlsx"
VALIDATION_FILENAME = "validation_report.txt"
UNREADABLE_LOG = "unreadable_volumes.log"

### Detta avsnitt är det som importerar configurationer som man fastställer i gränssnittet

def get_default_config():
    return {
        "start_page_strong_award_markers": START_PAGE_STRONG_AWARD_MARKERS,
        "start_page_degree_markers": START_PAGE_DEGREE_MARKERS,
        "start_page_weak_markers": START_PAGE_WEAK_MARKERS,
        "special_start_page_identifiers": SPECIAL_START_PAGE_IDENTIFIERS,
        "non_start_page_patterns": NON_START_PAGE_PATTERNS,

        # GUI/configurable behavior flags
        "skip_unreadable_volumes": True,
        "require_second_page_confirmation": True,
        "allow_personnummer_on_start_page": False,

    }


# -------------------------
# DETECTION
# -------------------------

def normalize_ocr_text(text):
    """
    Normalizes OCR text to make matching less sensitive to:
    - line breaks
    - repeated spaces
    - OCR spacing issues
    - Swedish characters being read inconsistently
    - punctuation differences around the Chalmers heading
    """
    if not text:
        return ""

    t = text.upper()
    t = t.replace("\r", " ").replace("\n", " ")
    t = t.replace("Å", "Å").replace("Ä", "Ä").replace("Ö", "Ö")

    # OCR often inserts/removes punctuation in the university line.
    t = re.sub(r"[;:,.]+", " ", t)
    t = re.sub(r"\s+", " ", t)
    return t.strip()


def contains_personnummer(text):
    """
    Detects Swedish personnummer / identifier formats that should NOT appear
    on an ordinary certificate start page.

    Covers examples such as:
    - 19991012-1234
    - 991012-1234
    - 1999-10-12
    - 1999-1012
    - 890747-T123
    - 990101T123
    - 990101-T123
    """
    if not text:
        return False

    t = text.upper()

    # Här är det personnummerformat som skriptet accepterar

    patterns = [
        r"\b\d{8}[-+]\d{4}\b",          # xxxxxxxx-xxxx
        r"\b\d{6}[-+]\d{4}\b",          # xxxxxx-xxxx
       # r"\b\d{4}-\d{2}-\d{2}\b",       # yyyy-mm-dd
       # r"\b\d{4}-\d{4}\b",             # yyyy-mmdd
        r"\b\d{6}-?T\d{3}\b",           # xxxxxx-Txxx or xxxxxxTxxx
        r"\b\d{8}-?T\d{3}\b",           # xxxxxxxx-Txxx or xxxxxxxxTxxx
        r"\b\d{6}-?R\d{3}\b",           # xxxxxx-Rxxx   
        r"\b\d{8}-?R\d{3}\b",           # xxxxxxxx-Rxxx
    ]

    return any(re.search(pattern, t) for pattern in patterns)


def fuzzy_identifier_present(identifier, text):
    """
    OCR-tolerant identifier check.
    First checks exact normalized match. Then checks whether most words in the
    identifier are present.
    """
    identifier_norm = normalize_ocr_text(identifier)
    text_norm = normalize_ocr_text(text)

    if identifier_norm in text_norm:
        return True

    words = identifier_norm.split()
    if not words:
        return False

    matched_words = sum(1 for word in words if word in text_norm)

    if len(words) <= 3:
        return matched_words == len(words)
    if len(words) <= 5:
        return matched_words >= len(words) - 1
    return matched_words >= max(3, int(len(words) * 0.70))


def count_fuzzy_matches(markers, text):
    return sum(1 for marker in markers if fuzzy_identifier_present(marker, text))


def compact_ocr_text(text):
    """
    Removes all non-alphanumeric characters. This catches OCR/PDF extraction cases
    where headings are glued together, e.g. 'Examensbevisför' or 'DatumDate'.
    """
    t = normalize_ocr_text(text)
    return re.sub(r"[^A-ZÅÄÖ0-9]+", "", t)

def plain_text_to_fuzzy_regex(marker):
    """
    Converts

    DIPLOMA SUPPLEMENT

    into a regex that matches:

    DIPLOMA SUPPLEMENT
    DIPLOMASUPPLEMENT
    DIPLOMA   SUPPLEMENT
    DIPLOMA-SUPPLEMENT
    """

    marker = normalize_ocr_text(marker)

    words = re.findall(
        r"[A-ZÅÄÖ0-9]+",
        marker
    )

    if not words:
        return None

    return (
        r"\b"
        + r"[\W_]*".join(
            re.escape(word)
            for word in words
        )
        + r"\b"
    )

def has_non_start_page_indicators(text, config):

    compact = compact_ocr_text(text)

    for marker in config["non_start_page_patterns"]:

        if fuzzy_identifier_present(
            marker,
            text
        ):
            return True

        if compact_ocr_text(marker) in compact:
            return True

    return False


def has_page_one_footer(text):
    """
    Detects page-one footers with OCR noise.
    Handles:
    - Sid 1 av 4
    - Sid I av 4
    - Sid I av Page I of
    - Page 1 of 4
    - Page I of
    - Page L of
    """
    if not text:
        return False

    t = text.upper()

    t = t.replace("Ⅰ", "I")
    t = re.sub(r"\s+", " ", t)

    footer_patterns = [
        r"(?:^|\s)SID\s*[1IL]\s*(?:AV|/)?\s*\d*",
        r"(?:^|\s)SIDA\s*[1IL]\s*(?:AV|/)?\s*\d*",
        r"(?:^|\s)PAGE\s*[1IL]\s*(?:OF|/)?\s*\d*",
        r"(?:^|\s)[1IL]\s*(?:AV|OF|/)\s*\d*",
        r"(?:^|\s)SIDA\s*/\s*PAGE\s*[1IL]\s*\(\s*\d+\s*\)",
    ]

    return any(re.search(pattern, t) for pattern in footer_patterns)



def is_special_certificate(text, config):
    return any(fuzzy_identifier_present(marker, text)
               for marker in config["special_start_page_identifiers"])


def is_start_page(text, config):
    """
    Determines whether a page is the start page of a degree certificate.

    Ordinary degree-certificate start pages are accepted when they look like the
    ceremonial cover page:
    - strong award wording and/or a prominent degree title
    - optional page-1 footer bonus
    - optional weak Chalmers/officer support
    - no personnummer
    - no course-table / diploma-supplement / administrative-page indicators

    Special 'INTYG OM FULLGJORDA STUDIER' certificates are still accepted by
    their own marker because they intentionally follow a different format.
    """

    #Kontrollerar om det finns personummer och utestluter då att det är första sidan
    t_raw = text.upper()

    # HARD FAIL 1: administrative certificate pages
    if (
        "EXAMENSBEVIS" in t_raw
        or "DIPLOMA FOR" in t_raw
    ):
        return False

    # Preserve the existing exception workflow for rare special certificates.
    if is_special_certificate(text, config):
        return True

    # HARD FAIL 2: contains personnummer
    if (
            contains_personnummer(text)
            and
            not config.get(
                "allow_personnummer_on_start_page",
                False
            )
    ):
        return False

    if not text:
        return False


    # Later pages contain tables, bilingual administrative fields, or diploma supplement text.
    if has_non_start_page_indicators(text, config):
        return False

    # HARD FAIL: any structured academic page
    if (
        "EXAMENSBEVIS" in text.upper()
        or "DEGREE CERTIFICATE FOR" in text.upper()
    ):
        return False


   # HARD FAIL: course/table layout indicators
   # Här kan det lätt uppstå problem om det finns undantag där följande ord förekommer i titeln. 
   # Tidigare var ordet "credits" här vilket ställde till problem då vissa program har (120 credits) med i titeln för examen.
    # Only reject if CLEAR table structure is present
    table_indicators = [
        "KURS",
        "BETYG", "GRADE",
        "DATUM", "DATE"
    ]

    table_hits = sum(
        1 for word in table_indicators
        if word in text.upper()
    )

    if table_hits >= 2:
        return False


    t = normalize_ocr_text(text)

    award_hits = count_fuzzy_matches(config["start_page_strong_award_markers"], t)
    degree_hits = count_fuzzy_matches(config["start_page_degree_markers"], t)
    weak_hits = count_fuzzy_matches(config["start_page_weak_markers"], t)
    page_one = has_page_one_footer(text)

    # A later administrative page can repeat the awarded-degree wording and degree title.
    # Therefore, award+degree alone is NOT enough. Require one cover-page-only support:
    # either a page-1 footer or the officer/signature wording from the ceremonial page.
    has_officer_or_signature = any(
        fuzzy_identifier_present(marker, t)
        for marker in [
            "EXAMENSHANDLÄGGARE OFFICER OF DEGREE",
            "OFFICER OF DEGREE",
            "ON BEHALF OF THE PRESIDENT",
            "PÅ REKTORS VÄGNAR",

            "REKTOR",
            "ADM ASSISTENT",
            "ADM ASS",
        ]
    )

    if award_hits >= 1 and degree_hits >= 1 and (page_one or has_officer_or_signature):
        return True

    if (
            degree_hits >= 1
            and award_hits >= 1
            and has_officer_or_signature
    ):
        return True

    # OCR fallback: if the degree title is damaged, require award wording + page-1 footer
    # + at least one cover-page support marker. This prevents page 2 from being split.
    if award_hits >= 1 and page_one and (weak_hits >= 1 or has_officer_or_signature):
        return True

    # OCR fallback: if award wording is damaged, require degree title + page-1 footer
    # + officer/signature wording. This is intentionally strict.
    if degree_hits >= 1 and page_one and has_officer_or_signature:
        return True

    return False

# -------------------------
# OCR CHECK
# -------------------------
def is_pdf_readable(reader):
    try:
        for i in range(min(3, len(reader.pages))):
            text = reader.pages[i].extract_text()
            if text and len(text.strip()) > 20:
                return True
    except:
        return False
    return False

def has_certificates(reader, config):
    try:
        for page in reader.pages:
            text = page.extract_text() or ""
            if is_start_page(text, config):
                return True
    except:
        return False
    return False

# -------------------------
# ✅ PRIORITY-BASED PERSONNUMMER EXTRACTION
# -------------------------
def extract_personnummer_candidates(text):
    high = []
    medium = []
    low = []

    normalized = text.upper()
    normalized = normalized.replace("–", "-").replace("—", "-")
    normalized = re.sub(r"\s+", "", normalized)

    # -------------------------
    # HIGH PRIORITY
    # -------------------------

    # YYYY-MM-DD-XXXX
    for m in re.findall(
            r"\d{4}-\d{2}-\d{2}-\d{4}",
            normalized
    ):
        y, mo, d, tail = m.split("-")
        high.append(f"{y[2:]}{mo}{d}-{tail}")

    high.extend(re.findall(r"\d{6}-\d{4}", normalized))
    high.extend(re.findall(r"\d{6}-T\d{3}", normalized))

    for m in re.findall(r"\d{10}", normalized):
        high.append(f"{m[:6]}-{m[6:]}")

    # -------------------------
    # MEDIUM PRIORITY
    # -------------------------
    for m in re.findall(r"\d{4}-\d{4}", normalized):
        y, md = m.split("-")
        medium.append(y[2:] + md)

    # -------------------------
    # LOW PRIORITY
    # -------------------------
    for m in re.findall(r"\d{4}-\d{2}-\d{2}", normalized):
        y, mo, d = m.split("-")
        low.append(y[2:] + mo + d)

    for m in re.findall(r"\d{8}", normalized):
        low.append(m[2:])

    if high:
        return high
    if medium:
        return medium
    return low

# -------------------------
# ✅ SPECIAL CERTIFICATE PERSONNUMMER
# -------------------------
def extract_pnr_special(text):
    lines = text.split("\n")

    for line in lines:
        if "PERSONNUMMER" in line.upper():
            l = line.upper()
            l = l.replace("–", "-").replace("—", "-")
            l = re.sub(r"\s+", "", l)

            match = re.search(
                r"\d{6}-\d{4}|\d{6}-T\d{3}|\d{10}|\d{8}|\d{6}",
                l
            )

            if match:
                val = match.group(0)

                if re.match(r"\d{10}", val):
                    return f"{val[:6]}-{val[6:]}"
                if re.match(r"\d{8}", val):
                    return val[2:]

                return val

    return ""

def consensus(values):
    if not values:
        return ""
    return Counter(values).most_common(1)[0][0]

# -------------------------
# PROCESS ONE PDF
# -------------------------
def process_pdf(args):
    pdf_path, input_root, output_root, config = args

    rows = []
    validation = []
# Denna funktion kollar så att PDF:en inte är korrumperad, om någonting är fel med den borde detta rapporteras i "validation_report".
    try:
        reader = PdfReader(pdf_path)
    except Exception as e:
        return [], [f"❌ Cannot open: {pdf_path} | {type(e).__name__}: {e}"]

    start_pages = []

    for i, page in enumerate(reader.pages):
        text = page.extract_text() or ""

        if is_start_page(text, config):
            start_pages.append(i)

            validation.append(
                f"START PAGE detected on page {i + 1} in {pdf_path}"
            )

    if not start_pages:
        return [], []

    start_pages.append(len(reader.pages))

    if len(start_pages) > 2:
        validation.append(
            f"Split into {len(start_pages) - 1} certificates: {pdf_path}"
        )

    rel_dir = os.path.relpath(os.path.dirname(pdf_path), input_root)
    output_dir = os.path.join(output_root, rel_dir)
    os.makedirs(output_dir, exist_ok=True)

    base = os.path.splitext(os.path.basename(pdf_path))[0]
    volym = rel_dir.split(os.sep)[0]

    for i in range(len(start_pages)-1):
        start = start_pages[i]
        end = start_pages[i+1]


        writer = PdfWriter()
        for p in range(start, end):
            writer.add_page(reader.pages[p])

        out_name = f"{base}_{i+1:03d}.pdf"
        out_path = os.path.join(output_dir, out_name)

        try:
            with open(out_path, "wb") as f:
                writer.write(f)
        except:
            validation.append(f"❌ Write failed: {out_path}")
            continue

        all_pnrs = []

        first_text = reader.pages[start].extract_text() or ""

        # ✅ Special certificate priority
        if is_special_certificate(first_text, config):
            pnr_special = extract_pnr_special(first_text)
            if pnr_special:
                all_pnrs.append(pnr_special)

        for p in range(start, end):
            try:
                text = reader.pages[p].extract_text() or ""
            except:
                continue

            all_pnrs.extend(extract_personnummer_candidates(text))

        pnr = consensus(all_pnrs)

        if not pnr:
            validation.append(f"⚠️ Missing personnummer: {out_path}")

        rows.append([out_name, volym, pnr])

        # ✅ VALIDATION: exactly one start page per split
        start_count = 0

        for p in range(start, end):
            try:
                text = reader.pages[p].extract_text() or ""
            except:
                continue

            if is_start_page(text, config):
                start_count += 1

        if start_count != 1:
            validation.append(
                f"⚠️ Suspicious split (start pages={start_count}): {out_path}"
            )

    return rows, validation


# -------------------------
# MAIN
# -------------------------
def process_all(input_root, output_root, config, log_callback=None, progress_callback=None):

    # Följande är debug kod man kan aktivera för att se vilka "config keys" som matas vidare till gränssnittet.
    #print("CONFIG KEYS RECEIVED:")
    #for key in config.keys():
    #    print(" -", key)

    volumes = {}

    for root, _, files in os.walk(input_root):
        pdfs = [os.path.join(root, f) for f in files if f.lower().endswith(".pdf")]
        if pdfs:
            volumes[root] = pdfs

    print(f"📦 Found {len(volumes)} volumes")

    valid_pdfs = []
    skipped_volumes = []

    for volume, pdf_list in volumes.items():
        volume_ok = True

        for pdf in pdf_list:
            try:
                reader = PdfReader(pdf)
            except:
                volume_ok = False
                break

            if not is_pdf_readable(reader):
                volume_ok = False
                break

            if not has_certificates(reader, config):
                volume_ok = False
                break

        if volume_ok:
            valid_pdfs.extend(pdf_list)
        else:
            skipped_volumes.append(volume)

    print(f"✅ Valid PDFs: {len(valid_pdfs)}")
    print(f"⛔ Skipped volumes: {len(skipped_volumes)}")
## Progress bar
    total_pdfs = len(valid_pdfs)
    completed_pdfs = 0

    emit_log(
        f"Found {total_pdfs} valid PDF files.",
        log_callback
    )

    if progress_callback:
        progress_callback(
            0,
            total_pdfs
        )
###

    all_rows = []
    validation = []
    all_futures = []

    batch_size = 5

    max_workers = 2  # Om processen crashar beror det på att det är för mycket som laddas till minnet samtidigt. Sänk då detta värde.

    with ProcessPoolExecutor(max_workers=max_workers) as executor:

        batch = []

        for pdf_path in valid_pdfs:

            batch.append(
                (pdf_path, input_root, output_root, config)
            )

            if len(batch) == batch_size:
                futures = [
                    executor.submit(process_pdf, arg)
                    for arg in batch
                ]

                all_futures.extend(futures)
                batch = []

        # process remaining
        if batch:
            futures = [executor.submit(process_pdf, arg) for arg in batch]
            all_futures.extend(futures)

        for f in tqdm(
                as_completed(all_futures),
                total=len(all_futures),
                desc="Processing"
        ):
            rows, val = f.result()

            all_rows.extend(rows)
            validation.extend(val)

            completed_pdfs += 1

            if progress_callback:
                progress_callback(
                    completed_pdfs,
                    total_pdfs
                )

            emit_log(
                f"Processed {completed_pdfs}/{total_pdfs}",
                log_callback
            )

    os.makedirs(output_root, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.append(["file name", "volym", "personnummer"])

    for r in all_rows:
        ws.append(r)

    wb.save(os.path.join(output_root, EXCEL_FILENAME))

    with open(os.path.join(output_root, VALIDATION_FILENAME), "w", encoding="utf-8") as f:
        for v in validation:
            f.write(v + "\n")

    with open(os.path.join(output_root, UNREADABLE_LOG), "w", encoding="utf-8") as f:
        for v in skipped_volumes:
            f.write(v + "\n")

    print(f"\n📊 Excel: {EXCEL_FILENAME}")
    print(f"📋 Validation: {VALIDATION_FILENAME}")
    print(f"🚫 Skipped volumes: {UNREADABLE_LOG}")
    print("🎉 Done!")

# -------------------------
# ENTRY
# -------------------------
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python split_certificates.py <input_root> <output_root>")
        sys.exit(1)

    config = get_default_config()
    process_all(sys.argv[1], sys.argv[2], config)

