"""
GAL scraper: extracción completa del directorio OWA vía API FindPeople.

Usa paginación con Offset + MaxEntriesReturned para iterar sobre toda la
GAL (Global Address List). Es reanudable: guarda progreso en JSON y permite
continuar desde el último offset si la sesión expira.

Estrategia:
- Peticiones lentas (delay configurable, default 8s)
- Progreso guardado tras cada página (offset + personas acumuladas)
- Exportación final a JSON + CSV
- Detección de sesión expirada (HTTP 307) similar a api_extractor.py
"""

import csv
import json
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any, Optional
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError

from verificacion_correo.utils.logging import get_logger
from verificacion_correo.core.api_extractor import (
    _build_cookie_header,
    _get_canary,
    _build_headers,
    _get_persona,
    validate_session_api,
    SessionExpiredError,
    OWA_BASE,
    SESSION_HEALTH_CHECK_INTERVAL,
    SESSION_ESTIMATED_LIMIT,
)

logger = get_logger(__name__)


REQUEST_TIMEOUT = 120
DEFAULT_BATCH_SIZE = 100
DEFAULT_DELAY = 8.0
PROGRESS_FILENAME = "gal_progress.json"
OUTPUT_JSON = "directorio_completo.json"
OUTPUT_CSV = "directorio_completo.csv"
OUTPUT_CSV_FILTERED = "directorio_filtrado.csv"
COMPANIES_CACHE_FILE = "gal_companies.json"


def _build_find_people_payload(
    offset: int,
    max_entries: int,
    address_list_id: str,
    query_string: Optional[str] = None,
) -> Dict[str, Any]:
    """Build the FindPeople request payload for a given offset.

    Args:
        offset: Pagination offset.
        max_entries: Max results per page.
        address_list_id: GAL address list GUID.
        query_string: Optional AQS query string for server-side filtering
            (e.g. company name). When set, the server only returns contacts
            matching this query instead of the full GAL.
    """
    body: Dict[str, Any] = {
        "__type": "FindPeopleRequest:#Exchange",
        "IndexedPageItemView": {
            "__type": "IndexedPageView:#Exchange",
            "BasePoint": "Beginning",
            "Offset": offset,
            "MaxEntriesReturned": max_entries,
        },
        "ParentFolderId": {
            "__type": "TargetFolderId:#Exchange",
            "BaseFolderId": {
                "__type": "AddressListId:#Exchange",
                "Id": address_list_id,
            },
        },
        "PersonaShape": {
            "__type": "PersonaResponseShape:#Exchange",
            "BaseShape": "Default",
        },
    }
    if query_string:
        body["QueryString"] = query_string

    return {
        "__type": "FindPeopleJsonRequest:#Exchange",
        "Header": {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": "Exchange2013",
        },
        "Body": body,
    }


def fetch_company_list(
    session_file: str,
    address_list_id: str = "fed75805-8ba2-4323-9f6d-80be7e3abc6a",
    sample_size: int = 500,
) -> List[str]:
    """Fetch unique company names from GAL via a quick API scan.

    Makes a single FindPeople request to get a sample of contacts,
    then extracts unique CompanyName values.

    Args:
        session_file: Path to session file with cookies.
        address_list_id: GAL address list ID.
        sample_size: Number of contacts to sample for company names.

    Returns:
        Sorted list of unique company names.
    """
    cookie_str = _build_cookie_header(session_file)
    canary = _get_canary(session_file)
    if not canary:
        raise ValueError("X-OWA-CANARY not found in session file")

    payload = _build_find_people_payload(0, sample_size, address_list_id)
    body_bytes = json.dumps(payload).encode("utf-8")
    req = Request(
        f"{OWA_BASE}/owa/service.svc?action=FindPeople",
        data=body_bytes,
        headers=_build_headers(canary, cookie_str, "FindPeople"),
    )

    try:
        resp = urlopen(req, timeout=REQUEST_TIMEOUT)
        data = json.loads(resp.read().decode("utf-8"))
        body = data.get("Body", {})
        people = body.get("People") or body.get("ResultSet") or []

        companies = set()
        for person in people:
            company = (person.get("CompanyName") or "").strip()
            if company:
                companies.add(company)

        return sorted(companies)

    except HTTPError as e:
        if e.code == 307:
            raise SessionExpiredError("Session expired during company list fetch")
        raise
    except Exception as e:
        logger.error(f"Error fetching company list: {e}")
        raise


def save_companies_cache(companies: List[str], output_dir: Path):
    """Save company list to cache file."""
    cache_path = output_dir / COMPANIES_CACHE_FILE
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(companies, f, ensure_ascii=False, indent=2)
    logger.info(f"Companies cache saved: {cache_path} ({len(companies)} companies)")


def load_companies_cache(output_dir: Path) -> Optional[List[str]]:
    """Load company list from cache file."""
    cache_path = output_dir / COMPANIES_CACHE_FILE
    if cache_path.exists():
        with open(cache_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return None


def _flatten_persona(persona: dict) -> dict:
    """Flatten raw persona dict into a flat dict for CSV export.

    Extracts available fields from FindPeople response (nombre, email, sip, empresa)
    and enriched fields from GetPersona (telefono, departamento, oficina, direccion).
    """
    name = persona.get("DisplayName") or ""
    email_obj = persona.get("EmailAddress") or {}
    email = ""
    if isinstance(email_obj, dict):
        email = email_obj.get("EmailAddress") or ""
    elif isinstance(email_obj, str):
        email = email_obj

    phone = ""
    phone_items = persona.get("BusinessPhoneNumbersArray") or []
    if phone_items and isinstance(phone_items[0], dict):
        val = phone_items[0].get("Value") or {}
        if isinstance(val, dict):
            phone = val.get("Number") or val.get("NormalizedNumber") or ""
        elif isinstance(val, str):
            phone = val

    sip = persona.get("ImAddress") or ""
    company = persona.get("CompanyName") or ""
    department = persona.get("Department") or ""
    office = persona.get("OfficeLocation") or ""

    address = ""
    addr_items = persona.get("BusinessAddressesArray") or []
    if addr_items and isinstance(addr_items[0], dict):
        val = addr_items[0].get("Value") or {}
        if isinstance(val, dict):
            parts = [
                val.get("Street") or "",
                val.get("City") or "",
                val.get("State") or "",
                val.get("PostalCode") or "",
                val.get("Country") or "",
            ]
            parts = [p for p in parts if p.strip()]
            address = ", ".join(parts) if parts else ""
        elif isinstance(val, str):
            address = val

    return {
        "nombre": name,
        "email": email,
        "telefono": phone,
        "sip": sip,
        "empresa": company,
        "departamento": department,
        "oficina": office,
        "direccion": address,
    }


def _enrich_persona(persona: dict, cookie_str: str, canary: str) -> dict:
    """Enrich a FindPeople persona with full GetPersona details.

    FindPeople only returns basic fields (name, email, sip, company).
    GetPersona returns the full contact card (phone, department, address, office).
    """
    persona_id_obj = persona.get("PersonaId")
    if not persona_id_obj or not isinstance(persona_id_obj, dict):
        return persona

    persona_id = persona_id_obj.get("Id")
    if not persona_id:
        return persona

    try:
        enriched = _get_persona(persona_id, cookie_str, canary)
        if enriched:
            persona["BusinessPhoneNumbersArray"] = [{"Value": {"Number": enriched.phone}}] if enriched.phone else []
            persona["Department"] = enriched.department or ""
            persona["OfficeLocation"] = enriched.office_location or ""
            persona["BusinessAddressesArray"] = [{"Value": enriched.address}] if enriched.address else []
            if enriched.name and not persona.get("DisplayName"):
                persona["DisplayName"] = enriched.name
    except Exception as e:
        logger.debug(f"GetPersona enrichment failed for {persona_id}: {e}")

    return persona


def _csv_safe(text: Any) -> str:
    """Convert value to CSV-safe string."""
    if text is None:
        return ""
    s = str(text)
    return s.replace("\r", " ").replace("\n", " ").replace(";", ",")


def save_to_csv(people: List[dict], path: Path):
    """Save flattened contact list to CSV."""
    fieldnames = ["nombre", "email", "telefono", "sip", "empresa", "departamento", "oficina", "direccion"]
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=";")
        writer.writeheader()
        for p in people:
            writer.writerow({k: _csv_safe(v) for k, v in p.items()})
    logger.info(f"CSV exported: {path} ({len(people)} contacts)")


def save_to_json(people: List[dict], path: Path):
    """Save raw people data to JSON."""
    with open(path, "w", encoding="utf-8") as f:
        json.dump(people, f, ensure_ascii=False, indent=2)
    logger.info(f"JSON exported: {path} ({len(people)} contacts)")


class ProgressFile:
    """Manages resumable progress for GAL scraping.

    Supports two modes:
    - Unfiltered (browse all): tracks offset + accumulated people.
    - Filtered (company-by-company): tracks completed companies,
      current company + its offset, and accumulated people.
    """

    def __init__(self, directory: Path):
        self.directory = Path(directory)
        self.directory.mkdir(parents=True, exist_ok=True)
        self.path = self.directory / PROGRESS_FILENAME

    def load(self) -> Dict[str, Any]:
        if self.path.exists():
            with open(self.path, "r", encoding="utf-8") as f:
                data = json.load(f)
            data.setdefault("people", [])
            data.setdefault("completed_companies", [])
            return data
        return {"offset": 0, "people": [], "completed_companies": []}

    def save(
        self,
        offset: int,
        people: List[dict],
        completed_companies: Optional[List[str]] = None,
        current_company: Optional[str] = None,
        company_offset: int = 0,
    ):
        data = {
            "offset": offset,
            "count": len(people),
            "last_update": datetime.now().isoformat(),
            "completed_companies": completed_companies or [],
        }
        if current_company is not None:
            data["current_company"] = current_company
            data["company_offset"] = company_offset
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def clear(self):
        if self.path.exists():
            self.path.unlink()

    @property
    def exists(self) -> bool:
        return self.path.exists()


def scrape_gal(
    session_file: str,
    output_dir: str = "data",
    max_contacts: int = 0,
    batch_size: int = DEFAULT_BATCH_SIZE,
    request_delay: float = DEFAULT_DELAY,
    address_list_id: str = "fed75805-8ba2-4323-9f6d-80be7e3abc6a",
    force_restart: bool = False,
    progress_callback=None,
    session_health_callback=None,
    stop_flag: Optional[dict] = None,
    company_filter: Optional[List[str]] = None,
    enrich_contacts: bool = False,
) -> Dict[str, Any]:
    """Scrape contacts from the GAL using paginated FindPeople.

    When company_filter is provided, uses server-side AQS QueryString
    filtering — iterating company by company so only matching contacts
    are downloaded. Without a filter, browses the entire GAL.

    Args:
        session_file: Path to session file with cookies.
        output_dir: Directory for progress + output files.
        max_contacts: Max contacts to fetch (0 = unlimited).
        batch_size: Max entries per page.
        request_delay: Seconds between pagination requests.
        address_list_id: GAL address list ID.
        force_restart: Ignore existing progress and start fresh.
        progress_callback: Optional fn(count, total) for UI updates.
        session_health_callback: Optional fn(health_info) for session health updates.
        stop_flag: Optional dict with {"stop": bool} to signal stop.
        company_filter: Optional list of company names to filter by.
            When provided, the full GAL is scanned and contacts are filtered
            client-side by CompanyName (server-side Restriction returns HTTP 500
            on this Exchange server). Enrichment is applied only to filtered contacts.
        enrich_contacts: If True, fetch full details via GetPersona for each contact
            (phone, department, address, office). Slower but complete.

    Returns:
        Dict with stats (total, offset_end, expired, files, etc.).
    """
    if not session_file or not Path(session_file).exists():
        raise FileNotFoundError(f"Session file not found: {session_file}")
    if not output_dir:
        raise ValueError("output_dir cannot be empty")
    if max_contacts < 0:
        raise ValueError(f"max_contacts must be >= 0, got {max_contacts}")
    if batch_size <= 0:
        raise ValueError(f"batch_size must be > 0, got {batch_size}")
    if request_delay < 0:
        raise ValueError(f"request_delay must be >= 0, got {request_delay}")

    start = time.time()
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    progress = ProgressFile(output_path)
    cookie_str = _build_cookie_header(session_file)
    canary = _get_canary(session_file)

    if not canary:
        raise ValueError("X-OWA-CANARY not found in session file")

    # Initial session validation
    if session_health_callback:
        initial_health = validate_session_api(session_file)
        initial_health['calls_used'] = 0
        initial_health['estimated_limit'] = SESSION_ESTIMATED_LIMIT
        session_health_callback(initial_health)

        if not initial_health['valid']:
            raise SessionExpiredError(f"Session validation failed: {initial_health['message']}")

    # Load or reset progress
    if force_restart or not progress.exists:
        progress.clear()
        state = {"offset": 0, "people": [], "completed_companies": []}
    else:
        state = progress.load()

    all_people = state.get("people", [])
    completed_companies: List[str] = state.get("completed_companies", [])
    session_expired = False
    total_fetched = len(all_people)
    total_scanned = state.get("offset", 0)
    consecutive_failures = 0
    max_consecutive_failures = 5
    api_calls = 0

    def _check_session_health():
        nonlocal session_expired, api_calls
        if session_health_callback and api_calls % SESSION_HEALTH_CHECK_INTERVAL == 0:
            health = validate_session_api(session_file)
            health['calls_used'] = api_calls
            health['estimated_limit'] = SESSION_ESTIMATED_LIMIT
            session_health_callback(health)
            if not health['valid']:
                logger.warning(f"Session health check failed: {health['message']}")
                session_expired = True
                return True
        return False

    def _fetch一页(offset, current_batch, query_string=None):
        """Fetch a single page from FindPeople. Returns (people_list, raw_response)."""
        payload = _build_find_people_payload(offset, current_batch, address_list_id, query_string)
        body_bytes = json.dumps(payload).encode("utf-8")
        req = Request(
            f"{OWA_BASE}/owa/service.svc?action=FindPeople",
            data=body_bytes,
            headers=_build_headers(canary, cookie_str, "FindPeople"),
        )
        resp = urlopen(req, timeout=REQUEST_TIMEOUT)
        data = json.loads(resp.read().decode("utf-8"))
        body = data.get("Body", {})
        return body.get("People") or body.get("ResultSet") or []

    def _enrich_batch(people_list):
        """Enrich contacts with full GetPersona details if requested."""
        if not enrich_contacts or not people_list:
            return
        for i, person in enumerate(people_list):
            people_list[i] = _enrich_persona(person, cookie_str, canary)
            if i < len(people_list) - 1:
                time.sleep(0.5)

    # ── Filtered mode: scan all GAL + filter by CompanyName ─────────
    # NOTE: Exchange GAL FindPeople does NOT support server-side filtering
    # by CompanyName (Restriction returns HTTP 500). We fetch all contacts
    # in large batches, then filter client-side by CompanyName.
    # Strategy:
    #   1. Scan all GAL (no enrichment) — ~10 API calls for 10k contacts
    #   2. Filter by CompanyName — instant, client-side
    #   3. Enrich ONLY the filtered contacts (if enrich_contacts=True)
    # This keeps API calls well under the ~40-call session limit.
    if company_filter:
        company_set = set(company_filter)
        effective_batch = max(batch_size, 1000)
        logger.info(
            f"Filtered mode: fetching all GAL with batch={effective_batch}, "
            f"filtering client-side by {company_filter}"
        )

        # all_people = ALL scanned contacts (for resume/progress)
        # filtered_people = only contacts matching company_filter
        filtered_people: List[Dict[str, Any]] = []
        scanned_all: List[Dict[str, Any]] = state.get("people", [])
        scanned_offset = state.get("offset", 0)
        total_scanned = scanned_offset
        if scanned_all:
            logger.info(f"Resuming: {len(scanned_all)} contacts already scanned at offset {scanned_offset}")
            for p in scanned_all:
                if p.get("CompanyName") in company_set:
                    filtered_people.append(p)

        while True:
            if session_expired or (stop_flag and stop_flag.get("stop")):
                break
            if max_contacts > 0 and len(filtered_people) >= max_contacts:
                break

            current_batch = effective_batch
            remaining = max_contacts - len(filtered_people) if max_contacts > 0 else 0
            if remaining > 0:
                current_batch = min(effective_batch, remaining)

            try:
                people = _fetch一页(scanned_offset, current_batch)

                if not people:
                    logger.info("No more results from GAL")
                    break

                consecutive_failures = 0
                api_calls += 1
                total_scanned += len(people)

                # NO enrichment here — we enrich only filtered contacts later
                scanned_all.extend(people)
                for person in people:
                    if person.get("CompanyName") in company_set:
                        filtered_people.append(person)

                scanned_offset += len(people)

                if _check_session_health():
                    progress.save(scanned_offset, scanned_all)
                    break

                progress.save(scanned_offset, scanned_all)

                if progress_callback:
                    progress_callback(len(filtered_people), max_contacts)

                if stop_flag and stop_flag.get("stop"):
                    break

                logger.info(
                    f"Scanned {total_scanned} total, "
                    f"{len(filtered_people)} matched from {company_filter}"
                )
                time.sleep(request_delay)

            except HTTPError as e:
                if e.code == 307:
                    logger.warning(f"Session expired at offset {scanned_offset}")
                    session_expired = True
                    progress.save(scanned_offset, scanned_all)
                    break
                consecutive_failures += 1
                logger.error(f"HTTP {e.code} at offset {scanned_offset}: {e.reason} (fail {consecutive_failures}/{max_consecutive_failures})")
                if consecutive_failures >= max_consecutive_failures:
                    logger.error("Too many consecutive failures, aborting")
                    break
                time.sleep(request_delay * 2)

            except URLError as e:
                consecutive_failures += 1
                logger.error(f"Connection error at offset {scanned_offset}: {e.reason} (fail {consecutive_failures}/{max_consecutive_failures})")
                if consecutive_failures >= max_consecutive_failures:
                    logger.error("Too many consecutive failures, aborting")
                    break
                time.sleep(request_delay * 2)

            except Exception as e:
                consecutive_failures += 1
                logger.error(f"Unexpected error at offset {scanned_offset}: {e} (fail {consecutive_failures}/{max_consecutive_failures})")
                if consecutive_failures >= max_consecutive_failures:
                    logger.error("Too many consecutive failures, aborting")
                    break
                time.sleep(request_delay * 2)

        # ── Enrichment: only for filtered contacts ───────────────────
        if enrich_contacts and filtered_people and not session_expired:
            logger.info(f"Enriching {len(filtered_people)} filtered contacts via GetPersona...")
            for i, person in enumerate(filtered_people):
                if stop_flag and stop_flag.get("stop"):
                    break
                filtered_people[i] = _enrich_persona(person, cookie_str, canary)
                if i < len(filtered_people) - 1:
                    time.sleep(0.5)
                if (i + 1) % 50 == 0:
                    logger.info(f"  Enriched {i+1}/{len(filtered_people)}...")

        # Use filtered_people as the final result
        all_people = filtered_people
        total_scanned = scanned_offset

    # ── Unfiltered mode: browse entire GAL ───────────────────────────
    else:
        offset = state.get("offset", 0)

        while True:
            if stop_flag and stop_flag.get("stop"):
                break
            if max_contacts > 0 and total_fetched >= max_contacts:
                break

            current_batch = batch_size
            if max_contacts > 0:
                remaining = max_contacts - total_fetched
                if remaining <= 0:
                    break
                current_batch = min(batch_size, remaining)

            logger.info(f"Fetching offset {offset} (batch size {current_batch})...")

            try:
                people = _fetch一页(offset, current_batch)

                if not people:
                    logger.info("No more results — GAL fully fetched")
                    break

                consecutive_failures = 0
                total_scanned += len(people)

                _enrich_batch(people)

                all_people.extend(people)
                total_fetched += len(people)
                offset += len(people)
                api_calls += 1

                if _check_session_health():
                    progress.save(offset, all_people)
                    break

                progress.save(offset, all_people)

                if progress_callback:
                    progress_callback(total_fetched, max_contacts)

                if stop_flag and stop_flag.get("stop"):
                    break

                time.sleep(request_delay)

            except HTTPError as e:
                if e.code == 307:
                    logger.warning(f"Session expired at offset {offset}")
                    session_expired = True
                    progress.save(offset, all_people)
                    break
                consecutive_failures += 1
                logger.error(f"HTTP {e.code} at offset {offset}: {e.reason} (fail {consecutive_failures}/{max_consecutive_failures})")
                if consecutive_failures >= max_consecutive_failures:
                    logger.error("Too many consecutive failures, aborting")
                    break
                time.sleep(request_delay * 2)

            except URLError as e:
                consecutive_failures += 1
                logger.error(f"Connection error at offset {offset}: {e.reason} (fail {consecutive_failures}/{max_consecutive_failures})")
                if consecutive_failures >= max_consecutive_failures:
                    logger.error("Too many consecutive failures, aborting")
                    break
                time.sleep(request_delay * 2)

            except Exception as e:
                consecutive_failures += 1
                logger.error(f"Unexpected error at offset {offset}: {e} (fail {consecutive_failures}/{max_consecutive_failures})")
                if consecutive_failures >= max_consecutive_failures:
                    logger.error("Too many consecutive failures, aborting")
                    break
                time.sleep(request_delay * 2)

    # Final session health update
    if session_health_callback:
        final_health = validate_session_api(session_file) if not session_expired else {'valid': False, 'health': 'expired', 'message': 'Sesión expirada'}
        final_health['calls_used'] = api_calls
        final_health['estimated_limit'] = SESSION_ESTIMATED_LIMIT
        session_health_callback(final_health)

    # Export results
    flattened = [_flatten_persona(p) for p in all_people]

    json_path = output_path / OUTPUT_JSON
    csv_path = output_path / OUTPUT_CSV

    save_to_json(all_people, json_path)
    save_to_csv(flattened, csv_path)

    duration = time.time() - start
    stats = {
        "total": len(all_people),
        "total_scanned": total_scanned,
        "offset_end": total_scanned,
        "expired": session_expired,
        "stopped": stop_flag and stop_flag.get("stop"),
        "files": {
            "json": str(json_path),
            "csv": str(csv_path),
            "progress": str(progress.path),
        },
        "duration": duration,
        "api_calls": api_calls,
    }

    if company_filter:
        stats["filtered_companies"] = company_filter
        stats["completed_companies"] = completed_companies

    if session_expired:
        logger.warning(f"Scraper stopped early: session expired ({len(all_people)} contacts)")
    elif stop_flag and stop_flag.get("stop"):
        logger.info(f"Scraper stopped by user: {len(all_people)} contacts")
    else:
        logger.info(f"GAL scraping complete: {len(all_people)} contacts in {duration:.1f}s")

    return stats
