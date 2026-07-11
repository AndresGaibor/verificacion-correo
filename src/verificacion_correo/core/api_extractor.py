"""
API-based contact extraction from OWA FindPeople/GetPersona services.

Two-step approach:
1. FindPeople — search directory by email, get PersonaId
2. GetPersona — fetch full contact details (phone, address, department, etc.)
"""

import json
import time
from collections.abc import Callable
from pathlib import Path
from typing import Dict, List, Any, Optional
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError

from verificacion_correo.core.extractor import ContactInfo
from verificacion_correo.core.excel import ExcelReader, ExcelWriter, EmailRecord, ProcessingStatus
from verificacion_correo.utils.logging import get_logger


logger = get_logger(__name__)


OWA_BASE = "https://correoweb.madrid.org"
REQUEST_DELAY = 3.0

SESSION_HEALTH_CHECK_INTERVAL = 5  # Validate session every N API calls
SESSION_ESTIMATED_LIMIT = 40  # Estimated calls before expiry


class SessionExpiredError(Exception):
    """Raised when OWA returns 307, meaning the session needs re-authentication."""
    pass


def _build_cookie_header(session_file: str) -> str:
    path = Path(session_file)
    if not path.exists():
        raise FileNotFoundError(f"Session file not found: {session_file}")
    with open(path) as f:
        data = json.load(f)
    cookies = data.get("cookies", [])
    parts = [f"{c['name']}={c['value']}" for c in cookies if c.get("domain", "").endswith("madrid.org")]
    return "; ".join(parts)


def _get_canary(session_file: str) -> str:
    path = Path(session_file)
    if not path.exists():
        return ""
    with open(path) as f:
        data = json.load(f)
    for c in data.get("cookies", []):
        if c.get("name") == "X-OWA-CANARY":
            return c.get("value", "")
    for origin in data.get("origins", []):
        for item in origin.get("localStorage", []):
            if isinstance(item, dict):
                for k, v in item.items():
                    if "canary" in k.lower():
                        return v
    return ""


def validate_session_api(session_file: str) -> Dict[str, Any]:
    """Lightweight session validation using API (no Playwright).

    Makes a minimal FindPeople request to check if session is still valid.
    Returns a dict with validation status and health info.

    Returns:
        Dict with keys: valid (bool), message (str), calls_used (int),
        estimated_limit (int), health (str: 'ok'/'warning'/'danger'/'expired')
    """
    result = {
        'valid': False,
        'message': '',
        'calls_used': 0,
        'estimated_limit': SESSION_ESTIMATED_LIMIT,
        'health': 'unknown'
    }

    try:
        cookie_str = _build_cookie_header(session_file)
        canary = _get_canary(session_file)

        if not canary:
            result['message'] = 'X-OWA-CANARY no encontrado en la sesión'
            result['health'] = 'danger'
            return result

        # Make minimal FindPeople request (1 result only)
        payload = {
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
                    "Offset": 0,
                    "MaxEntriesReturned": 1,
                },
                "QueryString": "*",
                "ParentFolderId": {
                    "__type": "TargetFolderId:#Exchange",
                    "BaseFolderId": {
                        "__type": "AddressListId:#Exchange",
                        "Id": "fed75805-8ba2-4323-9f6d-80be7e3abc6a",
                    },
                },
                "PersonaShape": {
                    "__type": "PersonaResponseShape:#Exchange",
                    "BaseShape": "Default",
                },
            },
        }

        body_bytes = json.dumps(payload).encode("utf-8")
        req = Request(
            f"{OWA_BASE}/owa/service.svc?action=FindPeople",
            data=body_bytes,
            headers=_build_headers(canary, cookie_str, "FindPeople"),
        )

        resp = urlopen(req, timeout=30)
        data = json.loads(resp.read().decode("utf-8"))
        body = data.get("Body", {})

        if body.get("ResponseCode") == "NoError":
            result['valid'] = True
            result['message'] = 'Sesión válida'
            result['health'] = 'ok'
        else:
            result['message'] = f"Respuesta inesperada: {body.get('ResponseCode', 'desconocido')}"
            result['health'] = 'danger'

    except HTTPError as e:
        if e.code == 307:
            result['message'] = 'Sesión expirada (HTTP 307)'
            result['health'] = 'expired'
        else:
            result['message'] = f'Error HTTP: {e.code} - {e.reason}'
            result['health'] = 'danger'
    except URLError as e:
        result['message'] = f'Error de conexión: {e.reason}'
        result['health'] = 'danger'
    except FileNotFoundError:
        result['message'] = 'Archivo de sesión no encontrado'
        result['health'] = 'danger'
    except Exception as e:
        result['message'] = f'Error inesperado: {type(e).__name__}: {e}'
        result['health'] = 'danger'

    return result


def get_people_filters(session_file: str) -> List[Dict[str, Any]]:
    """Get all available address lists from OWA via GetPeopleFilters.

    Returns a list of dicts, each with keys:
        - DisplayName: human-readable name (e.g. "Default Global Address List")
        - FolderId: dict with "Id" (the AddressListId GUID)

    Raises SessionExpiredError on HTTP 307.
    """
    cookie_str = _build_cookie_header(session_file)
    canary = _get_canary(session_file)
    if not canary:
        raise ValueError("X-OWA-CANARY not found in session file")

    req = Request(
        f"{OWA_BASE}/owa/service.svc?action=GetPeopleFilters",
        data=b"{}",
        headers=_build_headers(canary, cookie_str, "GetPeopleFilters"),
    )

    try:
        resp = urlopen(req, timeout=30)
        data = json.loads(resp.read().decode("utf-8"))

        filters = []
        for item in data:
            display_name = item.get("DisplayName", "")
            folder_id = item.get("FolderId", {})
            list_id = folder_id.get("Id", "")
            if display_name and list_id:
                filters.append({
                    "DisplayName": display_name,
                    "FolderId": {"Id": list_id},
                })
        return filters

    except HTTPError as e:
        if e.code == 307:
            raise SessionExpiredError("Session expired during GetPeopleFilters")
        raise
    except Exception as e:
        logger.error(f"GetPeopleFilters failed: {type(e).__name__}: {e}")
        raise


def _build_headers(canary: str, cookie_str: str, action: str) -> Dict[str, str]:
    return {
        "Accept": "*/*",
        "Content-Type": "application/json; charset=UTF-8",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Action": action,
        "X-OWA-ActionName": "BrowseInDirectory",
        "X-OWA-CANARY": canary,
        "X-Requested-With": "XMLHttpRequest",
        "Cookie": cookie_str,
    }


def _find_persona_id(
    email: str,
    cookie_str: str,
    canary: str,
    address_list_id: str,
) -> Optional[str]:
    """Step 1: Search directory and return PersonaId."""
    payload = {
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
                "Offset": 0,
                "MaxEntriesReturned": 50,
            },
            "QueryString": email,
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

    body_bytes = json.dumps(payload).encode("utf-8")
    req = Request(
        f"{OWA_BASE}/owa/service.svc?action=FindPeople",
        data=body_bytes,
        headers=_build_headers(canary, cookie_str, "FindPeople"),
    )

    try:
        resp = urlopen(req, timeout=60)
        data = json.loads(resp.read().decode("utf-8"))
        body = data.get("Body", {})
        if body.get("ResponseCode") != "NoError":
            return None
        result_set = body.get("ResultSet") or body.get("People")
        if not result_set:
            return None
        persona_id = result_set[0].get("PersonaId", {}).get("Id")
        return persona_id
    except HTTPError as e:
        if e.code == 307:
            raise SessionExpiredError(f"Session expired (307) during FindPeople for {email}")
        logger.debug(f"FindPeople HTTP {e.code} for {email}: {e.reason}")
        return None
    except URLError as e:
        if isinstance(e.reason, TimeoutError):
            logger.debug(f"FindPeople timeout for {email}")
        else:
            logger.debug(f"FindPeople connection error for {email}: {e.reason}")
        return None
    except Exception as e:
        logger.debug(f"FindPeople failed for {email}: {type(e).__name__}: {e}")
        return None


def _get_persona(
    persona_id: str,
    cookie_str: str,
    canary: str,
) -> Optional[ContactInfo]:
    """Step 2: Fetch full persona details."""
    payload = {
        "__type": "GetPersonaJsonRequest:#Exchange",
        "Header": {
            "__type": "JsonRequestHeaders:#Exchange",
            "RequestServerVersion": "Exchange2013",
        },
        "Body": {
            "__type": "GetPersonaRequest:#Exchange",
            "PersonaId": {
                "__type": "ItemId:#Exchange",
                "Id": persona_id,
            },
            "PersonaShape": {
                "__type": "PersonaResponseShape:#Exchange",
                "BaseShape": "Default",
            },
        },
    }

    body_bytes = json.dumps(payload).encode("utf-8")
    req = Request(
        f"{OWA_BASE}/owa/service.svc?action=GetPersona",
        data=body_bytes,
        headers=_build_headers(canary, cookie_str, "GetPersona"),
    )

    try:
        resp = urlopen(req, timeout=60)
        data = json.loads(resp.read().decode("utf-8"))
        persona = data.get("Body", {}).get("Persona")
        if not persona:
            return None
        return _parse_persona(persona)
    except HTTPError as e:
        if e.code == 307:
            raise SessionExpiredError(f"Session expired (307) during GetPersona for {persona_id}")
        logger.debug(f"GetPersona HTTP {e.code}: {e.reason}")
        return None
    except URLError as e:
        if isinstance(e.reason, TimeoutError):
            logger.debug(f"GetPersona timeout")
        else:
            logger.debug(f"GetPersona connection error: {e.reason}")
        return None
    except Exception as e:
        logger.debug(f"GetPersona failed: {type(e).__name__}: {e}")
        return None


def _extract_first_array_value(persona: dict, array_key: str, value_key: str = "Value") -> Optional[str]:
    """Extract first value from a Persona array field (e.g. BusinessPhoneNumbersArray)."""
    items = persona.get(array_key)
    if not items or not isinstance(items, list):
        return None
    first = items[0]
    if not isinstance(first, dict):
        return str(first) if first else None
    val = first.get(value_key)
    return str(val) if val else None


def _parse_persona(persona: dict) -> Optional[ContactInfo]:
    """Parse GetPersona response into ContactInfo."""
    name = persona.get("DisplayName") or None

    email_obj = persona.get("EmailAddress")
    email = None
    if isinstance(email_obj, dict):
        email = email_obj.get("EmailAddress")
    elif isinstance(email_obj, str):
        email = email_obj

    sip = persona.get("ImAddress") or None
    company = persona.get("CompanyName") or None
    department = persona.get("Department") or None

    office = (
        _extract_first_array_value(persona, "OfficeLocationsArray")
        or persona.get("OfficeLocation")
        or None
    )

    phone = None
    phone_items = persona.get("BusinessPhoneNumbersArray")
    if phone_items and isinstance(phone_items, list) and len(phone_items) > 0:
        phone_val = phone_items[0].get("Value") if isinstance(phone_items[0], dict) else None
        if isinstance(phone_val, dict):
            phone = phone_val.get("Number") or phone_val.get("NormalizedNumber")
        elif isinstance(phone_val, str):
            phone = phone_val

    address = None
    addr_item = persona.get("BusinessAddressesArray")
    if addr_item and isinstance(addr_item, list) and len(addr_item) > 0:
        addr_val = addr_item[0].get("Value") if isinstance(addr_item[0], dict) else None
        if isinstance(addr_val, dict):
            parts = [
                addr_val.get("Street"),
                addr_val.get("City"),
                addr_val.get("State"),
                addr_val.get("PostalCode"),
                addr_val.get("Country"),
            ]
            parts = [p for p in parts if p]
            if parts:
                address = ", ".join(parts)
        elif isinstance(addr_val, str):
            address = addr_val

    contact = ContactInfo(
        name=name,
        email=email,
        phone=phone,
        sip=sip,
        address=address,
        department=department,
        company=company,
        office_location=office,
    )

    return contact if contact.is_valid() else None


def find_people(
    email: str,
    session_file: str,
    address_list_id: str = "fed75805-8ba2-4323-9f6d-80be7e3abc6a",
) -> Optional[ContactInfo]:
    """Full 2-step search: FindPeople → GetPersona."""
    cookie_str = _build_cookie_header(session_file)
    canary = _get_canary(session_file)
    if not canary:
        logger.error("X-OWA-CANARY not found in session")
        return None

    persona_id = _find_persona_id(email, cookie_str, canary, address_list_id)
    if not persona_id:
        logger.debug(f"No persona found for {email}")
        return None

    contact = _get_persona(persona_id, cookie_str, canary)
    if not contact:
        logger.debug(f"GetPersona returned no data for {email}")
        return None

    logger.debug(f"API found {email}: {contact.name or contact.email}")
    return contact


def process_emails_via_api(
    excel_path: str,
    session_file: str,
    batch_size: int = 10,
    address_list_id: str = "fed75805-8ba2-4323-9f6d-80be7e3abc6a",
    request_delay: float = REQUEST_DELAY,
    progress_callback: Optional[Callable[[int, int], None]] = None,
    session_health_callback: Optional[Callable[[Dict[str, Any]], None]] = None,
) -> Dict[str, Any]:
    """Process all pending emails using FindPeople + GetPersona API."""
    start = time.time()

    reader = ExcelReader(excel_path)
    summary = reader.read_pending_emails(batch_size=batch_size)

    if not summary.batches:
        logger.info("No pending emails to process via API")
        return {"total": 0, "success": 0, "not_found": 0, "errors": 0, "expired": False, "duration": 0}

    # Initial session validation
    if session_health_callback:
        initial_health = validate_session_api(session_file)
        initial_health['calls_used'] = 0
        initial_health['estimated_limit'] = SESSION_ESTIMATED_LIMIT
        session_health_callback(initial_health)

    writer = ExcelWriter(excel_path)
    total_success = total_not_found = total_errors = total_processed = 0
    session_expired = False
    remaining_skipped = 0
    api_calls = 0

    for batch_num, batch in enumerate(summary.batches, 1):
        logger.info(f"API batch {batch_num}/{len(summary.batches)}: {len(batch)} emails")

        idx = -1
        for idx, record in enumerate(batch):
            if idx > 0:
                time.sleep(request_delay)

            logger.debug(f"API searching: {record.email}")

            try:
                contact = find_people(record.email, session_file, address_list_id)
                api_calls += 1

                # Check session health periodically
                if session_health_callback and api_calls % SESSION_HEALTH_CHECK_INTERVAL == 0:
                    health = validate_session_api(session_file)
                    health['calls_used'] = api_calls
                    health['estimated_limit'] = SESSION_ESTIMATED_LIMIT
                    session_health_callback(health)

                    if not health['valid']:
                        logger.warning(f"Session health check failed: {health['message']}")
                        session_expired = True
                        record.status = ProcessingStatus.ERROR
                        record.data = {"error": f"Sesión expirada: {health['message']}"}
                        writer.write_result(record)
                        total_errors += 1
                        total_processed += 1
                        break

            except SessionExpiredError as e:
                logger.warning(f"Session expired: {e}")
                session_expired = True
                record.status = ProcessingStatus.ERROR
                record.data = {"error": "Sesión expirada — necesita reautenticación"}
                writer.write_result(record)
                total_errors += 1
                total_processed += 1
                break

            if contact:
                record.data = contact.to_dict()
                record.status = ProcessingStatus.SUCCESS
                total_success += 1
            else:
                record.status = ProcessingStatus.NOT_FOUND
                total_not_found += 1

            writer.write_result(record)
            total_processed += 1

            if progress_callback:
                progress_callback(total_processed, summary.pending_count)

        if session_expired:
            remaining_skipped = sum(len(b) for b in summary.batches[batch_num:]) + len(batch) - idx - 1
            if remaining_skipped > 0:
                logger.warning(f"Session expired — {remaining_skipped} emails skipped (will retry in next run)")
            break

        logger.info(f"API batch {batch_num} done")

    # Final session health update
    if session_health_callback:
        final_health = validate_session_api(session_file) if not session_expired else {'valid': False, 'health': 'expired', 'message': 'Sesión expirada'}
        final_health['calls_used'] = api_calls
        final_health['estimated_limit'] = SESSION_ESTIMATED_LIMIT
        session_health_callback(final_health)

    duration = time.time() - start
    if session_expired:
        logger.warning(f"API processing stopped early: session expired after {total_processed} emails")
    else:
        logger.info(f"API processing complete: {total_success}/{total_processed} found in {duration:.1f}s")

    return {
        "total": total_processed,
        "success": total_success,
        "not_found": total_not_found,
        "errors": total_errors,
        "expired": session_expired,
        "remaining": remaining_skipped,
        "duration": duration,
        "api_calls": api_calls,
    }
