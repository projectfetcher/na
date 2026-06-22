import os
import re
import sys
import time
import json
import base64
import hashlib
import logging
from datetime import datetime
from urllib.parse import urljoin, urlparse, urlsplit, urlunsplit, parse_qsl, urlencode

import requests
from bs4 import BeautifulSoup

# Optional: load secrets from a local .env file if python-dotenv is installed.
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# Optional heavy deps used for Excel export / duplicate tracking.
try:
    import pandas as pd
    import openpyxl
    _XLSX_AVAILABLE = True
except ImportError:
    _XLSX_AVAILABLE = False

# Optional heavy deps used for paraphrase quality gating.
try:
    import language_tool_python
    from sentence_transformers import SentenceTransformer, util as st_util
    _NLP_AVAILABLE = True
except ImportError:
    _NLP_AVAILABLE = False

# =============================================================================
#  CONFIG
# =============================================================================

BASE_URL = "https://www.jobsnamibia.net"

# jobsnamibia.net's main feed is the paginated "Latest Vacancies" listing.
# Page 1 is the bare path, subsequent pages use ?page=N (confirmed by the
# "Page 1 [2] [Next Page >>]" pager seen on the live site).
LISTING_PATH = "/latest_jobs_in_namibia"

REQUEST_DELAY = float(os.environ.get("REQUEST_DELAY", "1.0"))  # polite delay between requests, seconds
MAX_JOBS = int(os.environ.get("MAX_JOBS", "0"))                # 0 = no cap, otherwise stop after N new jobs

# How many listing pages to crawl. 0 or unset means "crawl page 1, then keep
# going until a page returns no job cards" (auto-detect end of pagination).
# A positive value caps the crawl at that many pages.
_scrape_pages_raw = int(os.environ.get("SCRAPE_PAGES", "0"))
SCRAPE_PAGES = _scrape_pages_raw if _scrape_pages_raw > 0 else None

OUTPUT_FILE = "jobsnamibia_jobs.xlsx"
PROCESSED_IDS_FILE = "jobsnamibia_processed.csv"

# ── WordPress (secrets via environment variables — see header docstring) ────
WP_URL      = os.environ.get("WP_BASE_URL", "")
WP_USER     = os.environ.get("WP_USERNAME", "")
WP_PASSWORD = os.environ.get("WP_APP_PASSWORD", "")
WP_BASE      = WP_URL.rstrip("/")
WP_JOBS_URL  = f"{WP_BASE}/job-listings"
WP_MEDIA_URL = f"{WP_BASE}/media"

# ── Mistral (secret via environment variable — see header docstring) ────────
MISTRAL_API_KEY = os.environ.get("MISTRAL_API_KEY", "")
MISTRAL_MODEL   = "mistral-small-latest"
MISTRAL_URL     = "https://api.mistral.ai/v1/chat/completions"

ENABLE_PARAPHRASE = True   # set False to skip paraphrasing entirely

# ── Startup checks: warn (don't crash) if secrets are missing ───────────────
for _var, _val, _feature in [
    ("MISTRAL_API_KEY", MISTRAL_API_KEY, "paraphrasing"),
    ("WP_USERNAME",     WP_USER,         "WordPress posting"),
    ("WP_APP_PASSWORD", WP_PASSWORD,     "WordPress posting"),
]:
    if not _val:
        logging.getLogger(__name__).warning(
            f"Environment variable {_var} is not set — {_feature} will be disabled/skipped."
        )

JOB_TYPE_MAPPING = {
    "full-time": "full-time", "full time": "full-time",
    "part-time": "part-time", "part time": "part-time",
    "contract":  "contract",  "temporary": "temporary",
    "internship":"internship","freelance": "freelance",
    "volunteer": "volunteer",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml",
    "Accept-Charset": "utf-8",
}

REQUEST_TIMEOUT = 25

# Reuse one TCP/TLS connection where possible for every request this run makes.
SESSION = requests.Session()
SESSION.headers.update(HEADERS)

# =============================================================================
#  LOGGING / COLOUR
# =============================================================================

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log_ = logging.getLogger(__name__)   # logger instance (.info/.warning/.error)

_USE_COLOUR = sys.stdout.isatty()

def _c(code, text):
    return f"\033[{code}m{text}\033[0m" if _USE_COLOUR else text

C_HEADER  = lambda t: _c("1;36",  t)
C_LABEL   = lambda t: _c("1;33",  t)
C_VALUE   = lambda t: _c("97",    t)
C_DIM     = lambda t: _c("2",     t)
C_GREEN   = lambda t: _c("1;32",  t)
C_RED     = lambda t: _c("1;31",  t)
C_BLUE    = lambda t: _c("1;34",  t)
C_DIVIDER = lambda: _c("2", "─" * 80)

def log(msg):
    """Plain console print (kept distinct from the log_ logger instance)."""
    print(msg, flush=True)

# Matches a plain email address inside free text job descriptions / company details.
EMAIL_PATTERN = re.compile(r"[A-Za-z0-9.+_-]+@[A-Za-z0-9-]+\.[A-Za-z0-9.-]+")

# UI boilerplate phrases that leak into job-details text and aren't part of
# the actual job content. Stripped before printing.
BOILERPLATE_PATTERNS = [
    re.compile(r"Need Help drafting up your CV.*$", re.I | re.S),
    re.compile(r"Our Services and their Price Lists.*$", re.I | re.S),
    re.compile(r"Get Our CV Package.*$", re.I | re.S),
    re.compile(r"Contact Details:.*$", re.I | re.S),
    re.compile(r"INTERVIEW TIPS:.*$", re.I | re.S),
    re.compile(r"Checkout our.*CV Layout.*$", re.I),
]

# Section headers used on jobsnamibia.net job-detail pages. Several of these
# show up as their own h3/h4 blocks above a short value (Experience, Job Type,
# Closing Date) and are extracted as key/value pairs rather than free text.
KEY_VALUE_HEADERS = ("experience", "job type", "closing date", "salary", "location")

# =============================================================================
#  TEXT CLEANUP / SANITIZATION
# =============================================================================

_MOJIBAKE = [
    ("Â", ""), ("â€™", "'"), ("â€œ", '"'), ("â€\x9d", '"'), ("â€", '"'),
    ("â€¢", "•"), ("â„¢", "™"), ("\u00a0", " "), ("\u200b", ""), ("\ufeff", ""),
]

def _fix_mojibake(text: str) -> str:
    for pattern, replacement in _MOJIBAKE:
        text = text.replace(pattern, replacement)
    text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", text)
    return text

def sanitize_text(text, is_url=False) -> str:
    """Light cleanup pass used right before sending a field to WordPress."""
    if not isinstance(text, str):
        text = str(text) if (text is not None and str(text) not in ("nan", "None", "NaN")) else ""
    text = text.strip()
    if text in ("nan", "None", "NaN", "", "N/A", "n/a", "NA", "na"):
        return ""
    text = _fix_mojibake(text)
    if is_url:
        return re.sub(r"[ \t\r\n\f\v]+", " ", text).strip()
    text = re.sub(r"#+\s*", "", text)
    text = re.sub(r"\*\*", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()

def clean_description(text):
    if not text:
        return text
    for pattern in BOILERPLATE_PATTERNS:
        text = pattern.sub("", text)
    return re.sub(r"\s+", " ", text).strip()

def clean_text(el):
    if el is None:
        return ""
    return re.sub(r"\s+", " ", el.get_text(" ", strip=True)).strip()

# =============================================================================
#  BASIC HTTP / PARSING HELPERS
# =============================================================================

def get_soup(url):
    resp = SESSION.get(url, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    resp.encoding = resp.encoding or "utf-8"
    return BeautifulSoup(resp.text, "lxml")

def parse_posted_date(date_str):
    """Parses strings like '30 June 2026' -> datetime, or None on failure."""
    if not date_str:
        return None
    date_str = date_str.strip()
    for fmt in ("%d %B %Y", "%d %b %Y", "%B %d, %Y", "%b %d, %Y"):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None

def absolute_url(href):
    if not href:
        return ""
    if href.startswith("http"):
        return href
    return urljoin(BASE_URL + "/", href.lstrip("/"))

def extract_email(text):
    """Returns the first email address found in the given text, or ''."""
    if not text:
        return ""
    m = EMAIL_PATTERN.search(text)
    return m.group(0) if m else ""

def region_from_url(job_url):
    """jobsnamibia.net job URLs look like /windhoek/Job-Slug or
    /walvis-bay/Job-Slug — the first path segment is a usable region hint."""
    try:
        path = urlparse(job_url).path.strip("/")
        seg = path.split("/")[0] if path else ""
        return seg.replace("-", " ").replace("_", " ").title()
    except Exception:
        return ""

# =============================================================================
#  COMPANY LOGO EXTRACTION
# =============================================================================

LOGO_KEYWORDS_RE = re.compile(r"logo", re.I)
PLACEHOLDER_LOGO_RE = re.compile(r"default|placeholder|avatar|no-?image|blank|generic|logo2023", re.I)

def clean_logo_url(raw: str) -> str:
    if not raw:
        return ""
    raw = raw.strip()
    if not raw.startswith("http"):
        raw = absolute_url(raw)
    return re.sub(r"[\"')\s]+$", "", raw)

def is_placeholder_logo(url: str) -> bool:
    if not url:
        return True
    return bool(PLACEHOLDER_LOGO_RE.search(url))

def extract_company_logo(soup: BeautifulSoup) -> str:
    """
    Best-effort company logo lookup for a jobsnamibia.net job-detail page.
    The site embeds a per-job company logo image right above/near the job
    title (alt text observed as "This is the company Logo"), distinct from
    the global site logo ("logo2023.webp") which we explicitly exclude.
    Priority: og:image meta tag > <img alt*="company logo"> > any <img>
    whose alt/src mentions "logo" and isn't the site logo placeholder.
    """
    og = soup.find("meta", property="og:image") or soup.find("meta", attrs={"name": "og:image"})
    if og:
        content = og.get("content", "")
        if content:
            cand = clean_logo_url(content)
            if cand and not is_placeholder_logo(cand):
                return cand

    for img in soup.find_all("img"):
        alt = img.get("alt", "") or ""
        if "company logo" in alt.lower():
            cand = clean_logo_url(img.get("src") or img.get("data-src") or "")
            if cand and not is_placeholder_logo(cand):
                return cand

    for img in soup.find_all("img"):
        blob = " ".join(filter(None, [
            " ".join(img.get("class", []) or []),
            img.get("id", ""), img.get("alt", ""), img.get("src", ""),
        ]))
        if LOGO_KEYWORDS_RE.search(blob):
            cand = clean_logo_url(img.get("src") or img.get("data-src") or "")
            if cand and not is_placeholder_logo(cand):
                return cand

    return ""

# =============================================================================
#  NLP TOOLS (lazy init, optional)
# =============================================================================

_grammar_tool = None
_sim_model    = None

def _get_grammar_tool():
    global _grammar_tool
    if _grammar_tool is None and _NLP_AVAILABLE:
        try:
            _grammar_tool = language_tool_python.LanguageTool(
                "en-US", remote_server="https://api.languagetool.org")
        except Exception as e:
            log_.warning(f"LanguageTool init failed: {e}")
    return _grammar_tool

def _get_sim_model():
    global _sim_model
    if _sim_model is None and _NLP_AVAILABLE:
        try:
            _sim_model = SentenceTransformer("all-MiniLM-L6-v2", device="cpu")
        except Exception as e:
            log_.warning(f"SentenceTransformer init failed: {e}")
    return _sim_model

def grammar_correct(text: str) -> str:
    tool = _get_grammar_tool()
    if tool:
        try:
            return language_tool_python.utils.correct(text, tool.check(text))
        except Exception:
            pass
    return text

def similarity_score(a: str, b: str) -> float:
    model = _get_sim_model()
    if model:
        try:
            emb = model.encode([a, b], convert_to_tensor=True)
            return float(st_util.pytorch_cos_sim(emb[0], emb[1]))
        except Exception:
            pass
    def tokens(s):
        return set(re.sub(r"[^a-z0-9 ]", " ", s.lower()).split())
    ta, tb = tokens(a), tokens(b)
    if not ta or not tb: return 0.0
    return len(ta & tb) / max(len(ta), len(tb))

def clean_output(text: str) -> str:
    text = _fix_mojibake(text)
    for pat in [r"\[/?INST\]", r"</?s>",
                r"(?i)(rewritten?|rephrased?|output|paraphrase[d]?)[:\s]+",
                r"\*\*", r"###", r"---"]:
        text = re.sub(pat, "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return grammar_correct(text.strip())

# =============================================================================
#  MISTRAL API
# =============================================================================

def mistral_generate(prompt: str, max_tokens: int = 400, temperature: float = 0.7) -> str:
    if not MISTRAL_API_KEY:
        log_.warning("MISTRAL_API_KEY not set — skipping paraphrase")
        return ""
    try:
        response = requests.post(
            MISTRAL_URL,
            headers={
                "Authorization": f"Bearer {MISTRAL_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": MISTRAL_MODEL,
                "messages": [{"role": "user", "content": prompt}],
                "max_tokens": max_tokens,
                "temperature": temperature,
            },
            timeout=30,
        )
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"].strip()
    except Exception as e:
        log_.error(f"Mistral API error: {e}")
        return ""

# =============================================================================
#  PARAPHRASE FUNCTIONS
# =============================================================================

def _print_wrapped(text: str, prefix: str = "   ", width: int = 100):
    words = text.split()
    line  = []
    for w in words:
        line.append(w)
        if len(" ".join(line)) >= width:
            print(f"{prefix}{' '.join(line)}")
            line = []
    if line:
        print(f"{prefix}{' '.join(line)}")

def paraphrase_title(title: str) -> str:
    if not ENABLE_PARAPHRASE:
        return title
    clean = sanitize_text(title)
    if not clean:
        return title

    print(f"\n ┌─ TITLE PARAPHRASE {'─'*45}")
    print(f" │ Original : \"{clean}\"")
    print(f" │ {'─'*60}")

    best_result = None
    best_sim    = 0.0

    for attempt in range(4):
        temp = round(0.68 + attempt * 0.06, 2)
        print(f" │ Attempt {attempt+1} (temp={temp}):")

        prompt = (
            f"Rewrite this job title professionally using different words. "
            f"Output ONLY the rewritten title, nothing else. "
            f"Keep it between 4 and 12 words.\n\nJob title: {clean}"
        )

        raw    = mistral_generate(prompt, max_tokens=50, temperature=temp)
        result = clean_output(raw).split("\n")[0].strip().strip('"').strip("'")

        wc     = len(result.split()) if result else 0
        sim    = similarity_score(clean, result) if result else 0.0
        is_dup = result.lower().strip() == clean.lower().strip()

        print(f" │    Output  : \"{result}\"")
        print(f" │    Words   : {wc} | Similarity: {sim:.3f} | Duplicate: {'Yes ⚠️' if is_dup else 'No'}")

        valid = bool(result) and 4 <= wc <= 14 and sim >= 0.55 and not is_dup

        if not valid:
            reasons = []
            if not result:  reasons.append("empty output")
            if wc < 4:      reasons.append(f"too short ({wc} words, min=4)")
            if wc > 14:     reasons.append(f"too long ({wc} words, max=14)")
            if sim < 0.55:  reasons.append(f"sim={sim:.3f} < 0.55")
            if is_dup:      reasons.append("identical to original")
            print(f" │    → ❌ REJECTED — {', '.join(reasons)}")
        else:
            if sim > best_sim:
                best_sim    = sim
                best_result = result
                print(f" │    → ✅ ACCEPTED — new best candidate (sim={sim:.3f})")
            else:
                print(f" │    → ✅ VALID but not better than current best (best sim={best_sim:.3f})")

        print(f" │ {'─'*60}")
        time.sleep(1)

    if best_result:
        print(f" │ 🏆 FINAL SELECTED : \"{best_result}\"")
        print(f" │    Similarity     : {best_sim:.3f}")
        print(f" └{'─'*65}")
        return best_result
    else:
        print(f" │ ⚠️  No valid paraphrase found → Keeping original: \"{clean}\"")
        print(f" └{'─'*65}")
        return clean

def paraphrase_description(text: str) -> str:
    if not ENABLE_PARAPHRASE:
        return text
    clean = sanitize_text(text)
    if not clean:
        return text

    paragraphs  = [p.strip() for p in re.split(r"\n+", clean) if p.strip()]
    if not paragraphs:
        paragraphs = [clean]
    rewritten   = []
    success_count = 0

    print(f"\n ┌─ DESCRIPTION PARAPHRASE ({len(paragraphs)} paragraph(s)) {'─'*15}")

    for i, para in enumerate(paragraphs):
        orig_wc = len(para.split())

        print(f"\n │ ┌─ Paragraph {i+1}/{len(paragraphs)} {'─'*50}")
        print(f" │ │ ORIGINAL ({orig_wc} words):")
        _print_wrapped(para, prefix=" │ │    ")
        print(f" │ │ {'─'*60}")

        prompt = (
            f"Rewrite this job description paragraph professionally. "
            f"Keep ALL facts, requirements, and responsibilities. "
            f"Use different sentence structure and vocabulary. "
            f"Output ONLY the rewritten paragraph — no labels, no explanation.\n\n"
            f"Original:\n{para}"
        )

        best_result = None
        best_sim    = 0.0
        accepted_text = None

        for attempt in range(3):
            temp = round(0.65 + attempt * 0.08, 2)
            print(f" │ │ Attempt {attempt+1}/3 (temp={temp}):")

            raw    = mistral_generate(prompt, max_tokens=500, temperature=temp)
            result = clean_output(raw).strip()

            rw  = len(result.split()) if result else 0
            sim = similarity_score(para, result) if result and rw >= 5 else 0.0

            if result:
                print(f" │ │    Paraphrased ({rw} words, sim={sim:.3f}):")
                _print_wrapped(result, prefix=" │ │       ")
            else:
                print(f" │ │    Paraphrased : (no output from model)")

            valid = bool(result) and rw >= 8 and sim >= 0.48

            if not valid:
                reasons = []
                if not result: reasons.append("empty output")
                if rw < 8:     reasons.append(f"too short ({rw} words, min=8)")
                if sim < 0.48: reasons.append(f"sim={sim:.3f} < 0.48")
                print(f" │ │    → ❌ REJECTED — {', '.join(reasons)}")
                if result and sim > best_sim:
                    best_sim    = sim
                    best_result = result
                    print(f" │ │       (stored as best fallback, sim={sim:.3f})")
            else:
                print(f" │ │    → ✅ ACCEPTED on attempt {attempt+1}")
                rewritten.append(result)
                success_count += 1
                accepted_text = result
                break

            print(f" │ │ {'─'*60}")
            time.sleep(1)

        if accepted_text is None:
            print(f" │ │ {'─'*60}")
            if best_result and best_sim >= 0.40:
                print(f" │ │ 🔁 FALLBACK — Using best attempt (sim={best_sim:.3f}):")
                _print_wrapped(best_result, prefix=" │ │    ")
                rewritten.append(best_result)
                success_count += 1
            else:
                print(f" │ │ ⚠️  KEPT ORIGINAL — no acceptable paraphrase (best sim={best_sim:.3f})")
                rewritten.append(para)

        print(f" │ └{'─'*62}")

    print(f"\n │ SUMMARY: {success_count}/{len(paragraphs)} paragraphs successfully paraphrased")
    print(f" └{'─'*80}\n")

    return "\n\n".join(rewritten)

def paraphrase_company(text: str) -> str:
    if not ENABLE_PARAPHRASE:
        return text
    clean = sanitize_text(text)
    if not clean:
        return text

    print(f"\n ┌─ COMPANY BLURB PARAPHRASE {'─'*37}")
    orig_wc = len(clean.split())
    print(f" │ Original ({orig_wc} words):")
    _print_wrapped(clean, prefix=" │    ")
    print(f" │ {'─'*60}")

    prompt = (
        f"Rewrite this company description professionally. "
        f"Preserve all facts. Use different wording. "
        f"Output ONLY the rewritten description.\n\nOriginal:\n{clean}"
    )

    raw    = mistral_generate(prompt, max_tokens=600, temperature=0.68)
    result = clean_output(raw)
    rw     = len(result.split()) if result else 0
    sim    = similarity_score(clean, result) if result and rw >= 10 else 0.0

    if result and rw >= 10:
        print(f" │ Paraphrased ({rw} words, sim={sim:.3f}):")
        _print_wrapped(result, prefix=" │    ")
        print(f" │ → ✅ ACCEPTED")
        print(f" └{'─'*65}")
        time.sleep(1)
        return result
    else:
        reasons = []
        if not result: reasons.append("empty output")
        if rw < 10:    reasons.append(f"too short ({rw} words, min=10)")
        print(f" │ → ❌ REJECTED — {', '.join(reasons)} — keeping original")
        print(f" └{'─'*65}")
        time.sleep(1)
        return clean

# =============================================================================
#  DUPLICATE TRACKER (persists across runs)
# =============================================================================

def _init_tracker():
    if not _XLSX_AVAILABLE:
        return
    if not os.path.exists(PROCESSED_IDS_FILE):
        pd.DataFrame(columns=[
            "Job ID", "Job URL", "Job Title", "Company Name",
            "Status", "Timestamp", "WP ID",
        ]).to_csv(PROCESSED_IDS_FILE, index=False)

def load_processed_ids() -> tuple:
    if not _XLSX_AVAILABLE:
        log_.warning("pandas not installed — duplicate tracking is in-run only, not persisted")
        return set(), set()
    _init_tracker()
    df = pd.read_csv(PROCESSED_IDS_FILE)
    return (
        set(df["Job ID"].fillna("").astype(str)),
        set(df.get("Job URL", pd.Series()).fillna("").astype(str)),
    )

def _upsert_row(job_id: str, updates: dict):
    if not _XLSX_AVAILABLE:
        return
    _init_tracker()
    df   = pd.read_csv(PROCESSED_IDS_FILE)
    mask = df["Job ID"].astype(str) == str(job_id)
    if mask.any():
        for col, val in updates.items():
            if col in df.columns:
                df.loc[mask, col] = val
        df.loc[mask, "Timestamp"] = datetime.now().isoformat()
    else:
        row = {"Job ID": job_id, "Timestamp": datetime.now().isoformat()}
        row.update(updates)
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    df.to_csv(PROCESSED_IDS_FILE, index=False)

def make_job_id(job_url: str, title: str = "", company: str = "") -> str:
    # jobsnamibia.net has no real per-job URL (multiple job cards share the
    # same listing page URL), so title+company is the primary key here.
    # job_url is kept as a last-resort fallback only.
    if title or company:
        seed = f"{title}|{company}"
        return hashlib.md5(seed.encode()).hexdigest()[:16]
    if job_url:
        return hashlib.md5(job_url.encode()).hexdigest()[:16]
    return hashlib.md5(b"unknown").hexdigest()[:16]

def mark_scraped(job_id, job_url, title, company):
    _upsert_row(job_id, {"Job URL": job_url, "Job Title": title,
                          "Company Name": company, "Status": "scraped"})

def mark_paraphrased(job_id):
    _upsert_row(job_id, {"Status": "paraphrased"})

def mark_posted(job_id, wp_id, wp_url):
    _upsert_row(job_id, {"Status": "posted", "WP ID": wp_id})

def mark_failed(job_id, reason):
    _upsert_row(job_id, {"Status": f"failed|{reason}"})

# =============================================================================
#  WORDPRESS POSTING
# =============================================================================

def _wp_auth_headers() -> dict:
    token = base64.b64encode(f"{WP_USER}:{WP_PASSWORD}".encode()).decode()
    return {"Authorization": f"Basic {token}", "Content-Type": "application/json"}

def get_or_create_term(taxonomy_url: str, name: str):
    if not name or not name.strip():
        return None
    slug = re.sub(r"[^a-z0-9-]", "-", name.lower().strip())
    h = _wp_auth_headers()
    try:
        r = requests.get(f"{taxonomy_url}?slug={slug}", headers=h, timeout=10, verify=False)
        terms = r.json()
        if isinstance(terms, list) and terms:
            return terms[0]["id"]
    except Exception:
        pass
    try:
        r = requests.post(taxonomy_url, json={"name": name, "slug": slug},
                          headers=h, auth=(WP_USER, WP_PASSWORD), timeout=10, verify=False)
        return r.json().get("id")
    except Exception as e:
        log_.error(f"Term create error '{name}': {e}")
        return None

def post_job_to_wordpress(job: dict) -> tuple:
    if not WP_USER or not WP_PASSWORD:
        log_.warning("WP_USERNAME / WP_APP_PASSWORD not set — skipping WordPress post")
        return None, None

    h = _wp_auth_headers()

    title       = sanitize_text(job.get("jobTitle", ""))
    description = sanitize_text(job.get("jobDescription", ""))
    if not title or not description:
        return None, None

    slug = re.sub(r"[^a-z0-9-]", "-", title.lower())[:80]
    try:
        r = requests.get(f"{WP_JOBS_URL}?slug={slug}", headers=h, timeout=10, verify=False)
        posts = r.json()
        if isinstance(posts, list) and posts:
            log_.info(f"⏭ Job already on WP: {title}")
            return posts[0]["id"], posts[0].get("link")
    except Exception:
        pass

    logo_url    = sanitize_text(job.get("companyLogo", ""), is_url=True)
    location    = sanitize_text(job.get("jobLocation", ""))
    raw_type    = sanitize_text(job.get("jobType", "")) or "Full-time"
    job_type_s  = JOB_TYPE_MAPPING.get(raw_type.lower().strip(), "full-time")
    company     = sanitize_text(job.get("companyName", ""))
    application = sanitize_text(job.get("application", ""), is_url=True)
    company_url = sanitize_text(job.get("companyUrl", ""), is_url=True)
    deadline    = sanitize_text(job.get("deadline", ""))
    co_website  = sanitize_text(job.get("companyWebsite", ""), is_url=True)
    qualif      = sanitize_text(job.get("jobQualifications", ""))
    experience  = sanitize_text(job.get("jobExperience", ""))
    co_address  = sanitize_text(job.get("companyAddress", ""))
    job_field   = sanitize_text(job.get("jobField", ""))
    salary      = sanitize_text(job.get("salaryRange", ""))
    about       = sanitize_text(job.get("companyDetails", ""))

    is_email = bool(re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", application))
    is_url_v = bool(re.match(r"^https?://[^\s]+$", application))
    if not (is_email or is_url_v):
        application = ""

    # Upload logo
    attachment_id = None
    if logo_url:
        try:
            img_r = requests.get(logo_url, headers={"User-Agent": HEADERS["User-Agent"]}, timeout=15)
            if img_r.status_code == 200:
                ct  = img_r.headers.get("Content-Type", "image/jpeg")
                ext = "png" if "png" in ct else "jpg"
                fn  = re.sub(r"[^a-z0-9]", "-", company.lower()) + "-logo." + ext
                up_h = dict(_wp_auth_headers())
                up_h["Content-Disposition"] = f"attachment; filename={fn}"
                up_h["Content-Type"] = ct
                up_r = requests.post(WP_MEDIA_URL, headers=up_h, data=img_r.content,
                                     auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False)
                if up_r.status_code in (200, 201):
                    attachment_id = up_r.json().get("id")
        except Exception as e:
            log_.warning(f"Logo upload failed: {e}")

    region_term_id   = get_or_create_term(f"{WP_BASE}/job_listing_region", location)
    job_type_term_id = get_or_create_term(f"{WP_BASE}/job_listing_type",
                                           job_type_s.replace("-", " ").title())

    payload = {
        "title":          title,
        "content":        description,
        "status":         "publish",
        "featured_media": attachment_id or 0,
        "meta": {
            "_job_title":          title,
            "_job_location":       location,
            "_job_type":           job_type_s,
            "_job_description":    description,
            "_application":        application,
            "_company_url":        company_url,
            "_job_expires":        deadline,
            "_company_name":       company,
            "_company_website":    co_website,
            "_company_logo":       str(attachment_id) if attachment_id else "",
            "_company_address":    co_address,
            "_company_details":    about,
            "_job_qualifications": qualif,
            "_job_experiences":    experience,
            "_job_field":          job_field,
            "_job_salary":         salary,
        },
    }
    if region_term_id:   payload["job_listing_region"] = [region_term_id]
    if job_type_term_id: payload["job_listing_type"]   = [job_type_term_id]

    for attempt in range(3):
        try:
            r = requests.post(WP_JOBS_URL, json=payload, headers=h,
                              auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False)
            r.raise_for_status()
            post = r.json()
            log_.info(f"✅ Job posted: '{title}' → WP ID {post.get('id')}")
            return post.get("id"), post.get("link")
        except Exception as e:
            log_.error(f"Job post attempt {attempt+1} failed: {e}")
            if attempt < 2:
                time.sleep(2 ** attempt)
    return None, None

# =============================================================================
#  STEP 1 — COLLECT LISTING PAGE URLS
# =============================================================================
#
# IMPORTANT: jobsnamibia.net's "Latest Vacancies" feed is NOT a two-step
# crawl (listing page -> separate job-detail page) the way MyJobMag is.
# Each job card in the feed (div.index_latest_jobs) already contains the
# full job content — title, company, location, experience, job type,
# closing date, company details, and the full description. The card's own
# title link is a dead `href="#"` anchor, so there is no usable per-job
# detail URL to follow. We therefore parse jobs directly out of each
# listing page (see STEP 2) instead of fetching a second page per job.

def _listing_page_url(page_num: int) -> str:
    if page_num <= 1:
        return BASE_URL + LISTING_PATH
    return f"{BASE_URL}{LISTING_PATH}?page={page_num}"

def collect_listing_page_urls(pages=SCRAPE_PAGES):
    """Builds the list of listing-page URLs to crawl. If `pages` is None,
    this just seeds page 1 — collect_and_parse_jobs() below will keep
    requesting subsequent pages on its own until a page yields zero job
    cards (auto-detecting the end of pagination)."""
    if pages:
        return [_listing_page_url(i) for i in range(1, pages + 1)]
    return [_listing_page_url(1)]

# =============================================================================
#  STEP 2 — PARSE JOB CARDS OUT OF A LISTING PAGE
# =============================================================================

def _card_kv_fields(card):
    """Reads the Experience / Job Type / Closing Date / Salary mini-fields
    rendered as <div class="index_jobs_details"><h4>Label</h4><hr><h5>Value</h5></div>
    blocks inside a job card's .index_more_details container."""
    fields = {}
    for kv in card.select(".index_more_details .index_jobs_details"):
        h4 = kv.find("h4")
        h5 = kv.find("h5")
        label = clean_text(h4).lower().rstrip(":").strip() if h4 else ""
        value = clean_text(h5) if h5 else ""
        if label:
            fields[label] = value
    return fields

def _card_description(card):
    """Pulls the full free-text job body from a card's .job_description
    block, stripping the boilerplate phrases that sometimes leak in
    (e.g. 'Go to method of application »')."""
    desc_el = card.select_one(".job_description")
    if desc_el is None:
        return ""
    return clean_description(clean_text(desc_el))

def _card_company_and_location(card):
    """The .latest_content block holds, in order: an <h2><a>Title</a></h2>,
    then one or more <p> tags — the first is the company name, the next is
    usually the location (often prefixed by a map-marker icon)."""
    ps = card.select(".latest_content p")
    company = clean_text(ps[0]) if ps else ""
    location = clean_text(ps[1]) if len(ps) > 1 else ""
    return company, location

def parse_listing_page(url):
    """Returns the list of job dicts found on one jobsnamibia.net listing
    page by reading each div.index_latest_jobs card directly (see the note
    at the top of STEP 1 for why there's no separate detail-page fetch)."""

    soup = get_soup(url)
    cards = soup.select("div.index_latest_jobs")

    jobs = []
    for card in cards:
        title_a = card.select_one(".latest_content h2 a") or card.select_one(".latest_content h2")
        title = clean_text(title_a)
        if not title:
            continue

        company_name, location = _card_company_and_location(card)

        kv = _card_kv_fields(card)
        job_type   = kv.get("job type", "")
        experience = kv.get("experience", "")
        deadline   = kv.get("closing date", "")
        salary     = kv.get("salary", "")

        description = _card_description(card)

        # Applications on jobsnamibia.net are email-based; look in the
        # company-details block first (most reliable), then fall back to
        # scanning the description.
        company_details_el = card.select_one(".latest_company_details")
        company_details_text = clean_text(company_details_el) if company_details_el else ""
        apply_email = extract_email(company_details_text) or extract_email(description)

        company_logo = clean_logo_url(
            (card.select_one("img[alt='This is the company Logo']") or {}).get("src", "")
            if card.select_one("img[alt='This is the company Logo']") else ""
        )
        if is_placeholder_logo(company_logo):
            company_logo = ""

        job = {
            "title": title,
            "job_url": url,           # no real per-job URL on this site; use the listing page
            "job_type": job_type,
            "qualification": "",       # not exposed as its own field on this site
            "experience": experience,
            "location": location,
            "city": location,
            "field": "",                # category taxonomy lives on separate /category/ pages, not per-job
            "posted_date": "",          # jobsnamibia.net doesn't expose a separate "posted" date
            "deadline": deadline,
            "description": description,
            "apply_url": "",            # applications are email-based, not redirect links
            "apply_email": apply_email,
            "apply_raw": "",
            "company_name": company_name,
            "company_url": "",          # no dedicated company profile page per job on this site
            "company_blurb": "",        # no separate company-blurb section distinct from the job body
            "company_logo": company_logo,
            "salary": salary,
            "source_page": url,
        }
        jobs.append(job)

    return jobs

def collect_and_parse_jobs(pages=SCRAPE_PAGES):
    """Crawls listing page(s) and returns (all_raw_jobs, pages_visited).
    If `pages` is falsy, keeps requesting page 2, 3, ... until a page
    yields zero job cards, then stops (auto-detecting pagination end)."""
    all_jobs = []
    visited = []
    i = 1
    while True:
        page_url = _listing_page_url(i)
        log(f"\n{'=' * 80}\nFETCHING LISTING PAGE {i}: {page_url}\n{'=' * 80}")

        try:
            page_jobs = parse_listing_page(page_url)
        except Exception as e:
            log(f"  ERROR fetching/parsing listing page {i}: {e}")
            break

        visited.append(page_url)
        log(f"  Found {len(page_jobs)} job card(s) on this page")
        all_jobs.extend(page_jobs)

        if pages:
            if i >= pages:
                break
        else:
            if len(page_jobs) == 0:
                log("  No job cards found — assuming end of pagination, stopping.")
                break

        i += 1
        time.sleep(REQUEST_DELAY)

    log(f"\nTotal job cards collected across {len(visited)} page(s): {len(all_jobs)}")
    return all_jobs, visited

# =============================================================================
#  STEP 3 — DEDUPLICATE + PARAPHRASE  (turns a raw scraped job into the
#           standardized, WordPress-ready / Excel-ready job dict)
# =============================================================================

def process_job(raw_job: dict, processed_ids: set, processed_urls: set, seen_content: set):
    """
    Applies persistent + in-run duplicate detection, then paraphrases the
    title/description/company blurb via Mistral, and returns the
    standardized job dict ready for WordPress posting / Excel export.
    Returns None if the job was a duplicate (and should be skipped).
    """
    job_url = raw_job.get("job_url", "")
    title   = raw_job.get("title", "")
    company = raw_job.get("company_name", "")
    location = raw_job.get("location") or raw_job.get("city", "")

    job_id = make_job_id(job_url, title, company)

    if job_id in processed_ids:
        log(C_DIM(f"  ⧳ Already processed (tracker) — skipped: {title} @ {company}"))
        return None

    fingerprint = (title.lower().strip(), company.lower().strip(), location.lower().strip())
    if fingerprint in seen_content:
        log(C_DIM(f"  ⧳ Duplicate content this run — skipped: {title}"))
        return None
    seen_content.add(fingerprint)

    mark_scraped(job_id, job_url, title, company)
    processed_ids.add(job_id)
    processed_urls.add(job_url)

    description = raw_job.get("description", "")
    blurb       = raw_job.get("company_blurb", "")

    paraphrased_title = title
    paraphrased_desc  = description
    paraphrased_blurb = blurb

    if ENABLE_PARAPHRASE and MISTRAL_API_KEY:
        print(C_BLUE(f"\n  ✍️  Paraphrasing '{title}' ..."))
        paraphrased_title = paraphrase_title(title)
        paraphrased_desc  = paraphrase_description(description)
        if blurb:
            paraphrased_blurb = paraphrase_company(blurb)
        mark_paraphrased(job_id)
    else:
        print(C_DIM("  ⚠️  Paraphrasing skipped (ENABLE_PARAPHRASE=False or MISTRAL_API_KEY not set)"))

    apply_url   = raw_job.get("apply_url", "")
    apply_email = raw_job.get("apply_email", "")
    application = apply_url or apply_email

    company_website = ""  # jobsnamibia.net applications are email-based; no employer domain to derive

    apply_method = "resolved_redirect" if apply_url else ("description_email" if apply_email else "not_found")

    return {
        # Paraphrased fields
        "jobTitle":          paraphrased_title,
        "jobDescription":    paraphrased_desc,
        "companyDetails":    paraphrased_blurb,
        # Original fields (audit / duplicate detection)
        "originalTitle":     title,
        "originalDesc":      description,
        # Structured fields
        "jobType":           raw_job.get("job_type", ""),
        "jobQualifications": raw_job.get("qualification", ""),
        "jobExperience":     raw_job.get("experience", ""),
        "jobLocation":       location,
        "jobField":          raw_job.get("field", ""),
        "datePosted":        raw_job.get("posted_date", ""),
        "deadline":          raw_job.get("deadline", ""),
        "application":       application,
        "companyUrl":        raw_job.get("company_url", ""),
        "companyName":       company,
        "companyLogo":       raw_job.get("company_logo", ""),
        "companyWebsite":    company_website,
        "companyAddress":    raw_job.get("city", ""),
        "jobUrl":            job_url,
        "salaryRange":       raw_job.get("salary", ""),
        "_jobId":            job_id,
        "_apply_method":     apply_method,
        "_apply_raw":        raw_job.get("apply_raw", ""),
    }

# =============================================================================
#  VERBOSE PRINTER
# =============================================================================

def print_job_verbose(index, job):
    desc = job.get("jobDescription", "")
    desc_preview = (desc[:400] + " [...]") if len(desc) > 400 else desc

    print()
    print(C_DIVIDER())
    print(C_HEADER(f"  JOB #{index}"))
    print(C_DIVIDER())
    print(f"  {C_LABEL('Title (original)')}    : {C_VALUE(job.get('originalTitle',''))}")
    print(f"  {C_LABEL('Title (paraphrased)')} : {C_GREEN(job.get('jobTitle',''))}")
    print(f"  {C_LABEL('Job Type')}             : {job.get('jobType','') or C_DIM('—')}")
    print(f"  {C_LABEL('Qualification')}        : {job.get('jobQualifications','') or C_DIM('—')}")
    print(f"  {C_LABEL('Experience')}           : {job.get('jobExperience','') or C_DIM('—')}")
    print(f"  {C_LABEL('Location')}             : {job.get('jobLocation','') or C_DIM('—')}")
    print(f"  {C_LABEL('Field')}                : {job.get('jobField','') or C_DIM('—')}")
    print(f"  {C_LABEL('Salary')}               : {job.get('salaryRange','') or C_DIM('—')}")
    print(f"  {C_LABEL('Posted')}               : {job.get('datePosted','') or C_DIM('—')}")
    print(f"  {C_LABEL('Deadline')}             : {job.get('deadline','') or C_DIM('—')}")

    application = job.get("application", "")
    print(f"  {C_LABEL('Apply')}                : {C_GREEN(application) if application else C_DIM('— not found —')}")
    print(f"  {C_LABEL('Apply Method')}         : {C_DIM(job.get('_apply_method',''))}")
    if job.get("_apply_raw"):
        print(f"  {C_LABEL('  (tracking link)')}   : {C_DIM(job['_apply_raw'])}")

    print()
    print(f"  {C_BLUE('── COMPANY ──────────────────────────────────────────')}")
    print(f"  {C_LABEL('Name')}      : {C_VALUE(job.get('companyName','') or C_DIM('—'))}")
    print(f"  {C_LABEL('Page')}      : {job.get('companyUrl','') or C_DIM('—')}")
    print(f"  {C_LABEL('Website')}   : {job.get('companyWebsite','') or C_DIM('—')}")
    print(f"  {C_LABEL('Logo')}      : {job.get('companyLogo','') or C_DIM('— none —')}")
    about = job.get("companyDetails", "")
    if about:
        preview = (about[:200] + " [...]") if len(about) > 200 else about
        print(f"  {C_LABEL('About')}     : {preview}")

    print()
    print(f"  {C_BLUE('── DESCRIPTION PREVIEW ─────────────────────────────')}")
    print(desc_preview if desc_preview else C_DIM("   — no description —"))
    print(f"  {C_LABEL('Job URL')}   : {job.get('jobUrl','')}")
    print(C_DIVIDER())

# =============================================================================
#  EXCEL SAVE (standardized column order)
# =============================================================================

EXCEL_HEADERS = [
    "Job Title", "Job Type", "Job Qualifications", "Job Experience",
    "Job Location", "Job Field", "Date Posted", "Deadline",
    "Job Description", "Application", "Company URL", "Company Name",
    "Company Logo", "Company Website", "Company Address",
    "Company Details", "Job URL", "Salary Range",
]

def _save_excel(jobs: list):
    if not _XLSX_AVAILABLE:
        log_.warning("pandas/openpyxl not installed — skipping Excel export")
        return
    if not jobs:
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(EXCEL_HEADERS)
    for job in jobs:
        ws.append([
            job["jobTitle"], job["jobType"], job["jobQualifications"], job["jobExperience"],
            job["jobLocation"], job["jobField"], job["datePosted"], job["deadline"],
            job["jobDescription"], job["application"], job["companyUrl"], job["companyName"],
            job["companyLogo"], job["companyWebsite"], job["companyAddress"],
            job["companyDetails"], job["jobUrl"], job["salaryRange"],
        ])
    wb.save(OUTPUT_FILE)
    log_.info(f"Saved {len(jobs)} rows → {OUTPUT_FILE}")

# =============================================================================
#  MAIN
# =============================================================================

def main():
    start_time = datetime.now()

    print()
    print(C_HEADER("=" * 80))
    print(C_HEADER("  JOBSNAMIBIA.NET SCRAPER + MISTRAL PARAPHRASE + WORDPRESS POSTING"))
    print(C_HEADER("=" * 80))
    print(f"  Scrape pages    : {SCRAPE_PAGES}")
    print(f"  Request delay   : {REQUEST_DELAY}s")
    print(f"  Max new jobs    : {'unlimited' if not MAX_JOBS else MAX_JOBS}")
    print(f"  Paraphrase      : {'✅ enabled' if (ENABLE_PARAPHRASE and MISTRAL_API_KEY) else '❌ disabled'}")
    print(f"  WordPress post  : {'✅ enabled' if (WP_USER and WP_PASSWORD) else '❌ disabled'}")
    print(f"  Excel export    : {'✅ enabled' if _XLSX_AVAILABLE else '❌ disabled (pip install pandas openpyxl)'}")
    print(f"  NLP gating      : {'✅' if _NLP_AVAILABLE else '⚠️  no sentence-transformers / language-tool'}")
    print(f"  Started         : {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(C_HEADER("=" * 80))

    _init_tracker()
    processed_ids, processed_urls = load_processed_ids()
    print(f"  Tracker loaded: {len(processed_ids)} previously processed job IDs\n")

    raw_jobs, page_urls = collect_and_parse_jobs(SCRAPE_PAGES)

    jobs_out = []
    seen_content = set()
    total_raw_jobs = 0
    posted_count = 0
    errors = 0

    for raw_job in raw_jobs:
        total_raw_jobs += 1
        try:
            job = process_job(raw_job, processed_ids, processed_urls, seen_content)
        except Exception as e:
            errors += 1
            log(C_RED(f"  ✗ ERROR processing job: {e}"))
            continue

        if job is None:
            continue

        jobs_out.append(job)
        print_job_verbose(len(jobs_out), job)

        print(C_BLUE("\n  📤 Posting to WordPress …"))
        wp_id, wp_url = post_job_to_wordpress(job)
        if wp_id:
            mark_posted(job["_jobId"], wp_id, wp_url or "")
            posted_count += 1
            print(C_GREEN(f"  ✅ WP ID={wp_id}  🔗 {wp_url}"))
        else:
            mark_failed(job["_jobId"], "wp_post_failed_or_skipped")
            print(C_RED("  ❌ WordPress post failed / skipped"))

        if len(jobs_out) % 25 == 0:
            _save_excel(jobs_out)

        if MAX_JOBS and len(jobs_out) >= MAX_JOBS:
            log(f"\nMAX_JOBS limit ({MAX_JOBS}) reached, stopping.")
            break

    _save_excel(jobs_out)

    end_time = datetime.now()
    duration = (end_time - start_time).total_seconds() / 60.0
    print()
    print(C_HEADER("=" * 80))
    print(C_HEADER("  SCRAPE COMPLETE"))
    print(C_HEADER("=" * 80))
    print(f"  {C_LABEL('Job pages visited')}          : {len(page_urls)}")
    print(f"  {C_LABEL('Raw jobs found')}             : {total_raw_jobs}")
    print(f"  {C_LABEL('New jobs processed')}         : {C_GREEN(str(len(jobs_out)))}")
    print(f"  {C_LABEL('Posted to WordPress')}        : {C_GREEN(str(posted_count))}")
    print(f"  {C_LABEL('Errors')}                     : {C_RED(str(errors)) if errors else '0'}")
    print(f"  {C_LABEL('Duration')}                   : ~{duration:.1f} min")
    print(f"  {C_LABEL('Output file')}                : {OUTPUT_FILE}")
    print(f"  {C_LABEL('Tracker file')}                : {PROCESSED_IDS_FILE}")

    if jobs_out:
        with_apply = sum(1 for j in jobs_out if j.get("application"))
        with_email = sum(1 for j in jobs_out if "@" in (j.get("application") or ""))
        with_url   = with_apply - with_email
        no_apply   = len(jobs_out) - with_apply
        print(f"\n  {C_LABEL('Application links:')}")
        print(f"    URL found    : {with_url}")
        print(f"    Email found  : {with_email}")
        print(f"    Not found    : {no_apply}")

        para_count = sum(1 for j in jobs_out if j.get("jobTitle") != j.get("originalTitle"))
        print(f"\n  {C_LABEL('Paraphrased titles')} : {para_count}/{len(jobs_out)}")

        with_logo = sum(1 for j in jobs_out if j.get("companyLogo"))
        print(f"  {C_LABEL('Logos found')}        : {with_logo}/{len(jobs_out)}")

    print(C_HEADER("=" * 80))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(1)
