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
    SessionExpiredError,
    OWA_BASE,
)

logger = get_logger(__name__)


REQUEST_TIMEOUT = 120
DEFAULT_BATCH_SIZE = 100
DEFAULT_DELAY = 8.0
PROGRESS_FILENAME = "gal_progress.json"
OUTPUT_JSON = "directorio_completo.json"
OUTPUT_CSV = "directorio_completo.csv"


def _build_find_people_payload(
    offset: int,
    max_entries: int,
    address_list_id: str,
) -> Dict[str, Any]:
    """Build the FindPeople request payload for a given offset."""
    return {
        "__type": "FindPeopleJsonRequest:#Exchange",
        "Header": {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": "Exchange2013",
        },
        "Body": {
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
        },
    }


def _flatten_persona(persona: dict) -> dict:
    """Flatten raw persona dict into a flat dict for CSV export."""
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
    """Manages resumable progress for GAL scraping."""

    def __init__(self, directory: Path):
        self.directory = Path(directory)
        self.directory.mkdir(parents=True, exist_ok=True)
        self.path = self.directory / PROGRESS_FILENAME

    def load(self) -> Dict[str, Any]:
        if self.path.exists():
            with open(self.path, "r", encoding="utf-8") as f:
                return json.load(f)
        return {"offset": 0, "people": []}

    def save(self, offset: int, people: List[dict]):
        data = {
            "offset": offset,
            "count": len(people),
            "last_update": datetime.now().isoformat(),
        }
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
    stop_flag: Optional[dict] = None,
) -> Dict[str, Any]:
    """Scrape all contacts from the GAL using paginated FindPeople.

    Args:
        session_file: Path to session file with cookies.
        output_dir: Directory for progress + output files.
        max_contacts: Max contacts to fetch (0 = unlimited).
        batch_size: Max entries per page.
        request_delay: Seconds between pagination requests.
        address_list_id: GAL address list ID.
        force_restart: Ignore existing progress and start fresh.
        progress_callback: Optional fn(count, total) for UI updates.
        stop_flag: Optional dict with {"stop": bool} to signal stop.

    Returns:
        Dict with stats (total, offset_end, expired, files, etc.).
    """
    start = time.time()
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    progress = ProgressFile(output_path)
    cookie_str = _build_cookie_header(session_file)
    canary = _get_canary(session_file)

    if not canary:
        raise ValueError("X-OWA-CANARY not found in session file")

    # Load or reset progress
    if force_restart or not progress.exists:
        progress.clear()
        state = {"offset": 0, "people": []}
    else:
        state = progress.load()
        logger.info(f"Resuming from offset {state['offset']} ({len(state['people'])} contacts already fetched)")

    offset = state["offset"]
    all_people = state["people"]
    session_expired = False
    total_fetched = len(all_people)

    while True:
        if stop_flag and stop_flag.get("stop"):
            logger.info("Scraper stopped by user")
            break

        if max_contacts > 0 and total_fetched >= max_contacts:
            logger.info(f"Reached max contacts: {max_contacts}")
            break

        current_batch = min(batch_size, max_contacts - total_fetched) if max_contacts > 0 else batch_size

        logger.info(f"Fetching offset {offset} (batch size {current_batch})...")

        try:
            payload = _build_find_people_payload(offset, current_batch, address_list_id)
            body_bytes = json.dumps(payload).encode("utf-8")
            req = Request(
                f"{OWA_BASE}/owa/service.svc?action=FindPeople",
                data=body_bytes,
                headers=_build_headers(canary, cookie_str, "FindPeople"),
            )

            resp = urlopen(req, timeout=REQUEST_TIMEOUT)
            data = json.loads(resp.read().decode("utf-8"))

            body = data.get("Body", {})
            people = body.get("People") or body.get("ResultSet") or []

            if not people:
                logger.info("No more results — GAL fully fetched")
                break

            all_people.extend(people)
            total_fetched += len(people)
            offset += len(people)

            # Save progress
            progress.save(offset, all_people)

            if progress_callback:
                progress_callback(total_fetched, max_contacts)

            if stop_flag and stop_flag.get("stop"):
                break

            # Delay before next request
            time.sleep(request_delay)

        except HTTPError as e:
            if e.code == 307:
                logger.warning(f"Session expired at offset {offset}")
                session_expired = True
                progress.save(offset, all_people)
                break
            logger.error(f"HTTP {e.code} at offset {offset}: {e.reason}")
            time.sleep(request_delay * 2)

        except URLError as e:
            logger.error(f"Connection error at offset {offset}: {e.reason}")
            time.sleep(request_delay * 2)

        except Exception as e:
            logger.error(f"Unexpected error at offset {offset}: {e}")
            time.sleep(request_delay * 2)

    # Export results
    flattened = [_flatten_persona(p) for p in all_people]

    json_path = output_path / OUTPUT_JSON
    csv_path = output_path / OUTPUT_CSV

    save_to_json(all_people, json_path)
    save_to_csv(flattened, csv_path)

    duration = time.time() - start
    stats = {
        "total": len(all_people),
        "offset_end": offset,
        "expired": session_expired,
        "stopped": stop_flag and stop_flag.get("stop"),
        "files": {
            "json": str(json_path),
            "csv": str(csv_path),
            "progress": str(progress.path),
        },
        "duration": duration,
    }

    if session_expired:
        logger.warning(f"Scraper stopped early: session expired at offset {offset} ({len(all_people)} contacts)")
    elif stop_flag and stop_flag.get("stop"):
        logger.info(f"Scraper stopped by user: {len(all_people)} contacts")
    else:
        logger.info(f"GAL scraping complete: {len(all_people)} contacts in {duration:.1f}s")

    return stats
