"""
Contact information extraction from OWA popup cards.

This module provides robust extraction of contact information from OWA popup
elements, combining DOM-based extraction with regex fallbacks for maximum
reliability.
"""

import re
from typing import Dict, Any, Optional, List
from dataclasses import dataclass
from playwright.sync_api import Page, Locator, TimeoutError as PWTimeout

from .config import Config
from ..utils.logging import get_logger


logger = get_logger(__name__)


@dataclass
class ContactInfo:
    """
    Structured contact information extracted from OWA popup.

    Note: Due to Microsoft OWA anti-scraping protection, the 'name' field
    may not be available for certain accounts when automation is detected.
    """
    name: Optional[str] = None
    email: Optional[str] = None
    phone: Optional[str] = None
    sip: Optional[str] = None
    address: Optional[str] = None
    department: Optional[str] = None
    company: Optional[str] = None
    office_location: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary format for Excel writing."""
        return {
            'name': self.name,
            'email': self.email,
            'phone': self.phone,
            'sip': self.sip,
            'address': self.address,
            'department': self.department,
            'company': self.company,
            'office_location': self.office_location,
        }

    def is_valid(self) -> bool:
        """Check if the contact information contains meaningful data."""
        # Consider valid if at least email or phone is present
        return bool(self.email or self.phone)

    def __repr__(self) -> str:
        """String representation of contact info."""
        fields = [f"{k}={v!r}" for k, v in self.to_dict().items() if v]
        return f"ContactInfo({', '.join(fields)})"


class ContactExtractor:
    """
    Extracts contact information from OWA popup cards.

    Uses a multi-layered approach:
    1. DOM-based extraction using specific selectors
    2. Text-based extraction with regex patterns
    3. Heuristics for data validation and cleanup
    """

    def __init__(self, config: Config):
        """
        Initialize contact extractor.

        Args:
            config: Application configuration
        """
        self.config = config
        self.patterns = config.patterns
        self.selectors = config.selectors
        self.wait_times = config.wait_times

    def extract_from_popup(self, page: Page) -> Optional[ContactInfo]:
        """
        Extract contact information from OWA popup card.

        Args:
            page: Playwright Page object with visible popup

        Returns:
            ContactInfo object or None if extraction fails
        """
        try:
            # Wait for popup to be visible
            popup = self._wait_for_popup(page)
            if not popup:
                return None

            # Extract popup text for regex fallback
            popup_text = self._get_popup_text(popup)

            # Try DOM-based extraction first
            contact_info = self._extract_dom_based(popup)
            if contact_info and contact_info.is_valid():
                logger.debug("Successfully extracted using DOM-based method")
                return contact_info

            # Fall back to text-based extraction
            contact_info = self._extract_text_based(popup_text)
            if contact_info and contact_info.is_valid():
                logger.debug("Successfully extracted using text-based method")
                return contact_info

            logger.warning("Extraction failed - no valid contact data found")
            return None

        except Exception as e:
            logger.error(f"Error during contact extraction: {e}")
            return None

    def _wait_for_popup(self, page: Page) -> Optional[Locator]:
        """
        Wait for popup card to become visible.

        Args:
            page: Playwright Page object

        Returns:
            Popup locator or None if timeout
        """
        try:
            popup = page.locator(self.selectors.popup).first
            popup.wait_for(
                state="visible",
                timeout=self.wait_times.popup_visible
            )
            return popup
        except PWTimeout:
            logger.warning("Popup card not found or not visible")
            return None

    def _get_popup_text(self, popup: Locator) -> str:
        """
        Extract all visible text from popup.

        Args:
            popup: Popup locator

        Returns:
            Complete text content of popup
        """
        try:
            return popup.inner_text()
        except Exception as e:
            logger.error(f"Error extracting popup text: {e}")
            return ""

    def _extract_dom_based(self, popup: Locator) -> Optional[ContactInfo]:
        """
        Extract contact information using DOM selectors.

        This method tries to find specific DOM elements that contain
        contact information. It's more reliable than text-based extraction
        when the structure is consistent.

        Args:
            popup: Popup locator

        Returns:
            ContactInfo object or None if extraction fails
        """
        try:
            # Try specific OWA selectors (these may vary based on OWA version)
            dom_data = {}

            # Common OWA popup selectors (examples based on research)
            dom_selectors = {
                'name': [
                    'span._pe_c1._pe_t1',  # OWA 2019+
                    'div.ms-ContactCard-primaryText',
                    '[data-automation-id="contactCardName"]',
                    '.ms-PersonaCard-primaryText',
                ],
                'email': [
                    'a[href^="mailto:"]',
                    'span[title*="@"]',
                    '[data-automation-id="contactCardEmail"]',
                ],
                'phone': [
                    'a[href^="tel:"]',
                    '[data-automation-id="contactCardPhone"]',
                    '.ms-ContactCard-phoneText',
                ],
                'department': [
                    '[data-automation-id="contactCardTitle"]',
                    '.ms-ContactCard-jobTitle',
                ],
                'company': [
                    '[data-automation-id="contactCardCompany"]',
                    '.ms-ContactCard-company',
                ],
                'office_location': [
                    '[data-automation-id="contactCardOffice"]',
                    '.ms-ContactCard-office',
                ]
            }

            for field, selectors in dom_selectors.items():
                for selector in selectors:
                    try:
                        element = popup.locator(selector).first
                        if element.count() > 0:
                            text = element.inner_text(timeout=1000).strip()
                            if text and text.lower() != field.lower():  # Avoid field name placeholders
                                # Clean up common artifacts
                                if field == 'email' and text.startswith('mailto:'):
                                    text = text.replace('mailto:', '')
                                elif field == 'phone' and text.startswith('tel:'):
                                    text = text.replace('tel:', '')

                                dom_data[field] = text
                                break
                    except:
                        continue  # Try next selector

            # If we got meaningful data from DOM, create ContactInfo
            if dom_data:
                logger.debug(f"DOM extraction found: {list(dom_data.keys())}")
                return ContactInfo(**dom_data)

            return None

        except Exception as e:
            logger.debug(f"DOM-based extraction failed: {e}")
            return None

    def _extract_text_based(self, text: str) -> Optional[ContactInfo]:
        """
        Extract contact information from text using regex patterns.

        This is the fallback method when DOM extraction fails.
        It uses regex patterns to find contact information in the raw text.

        Args:
            text: Complete text content from popup

        Returns:
            ContactInfo object or None if no valid data found
        """
        try:
            if not text.strip():
                return None

            contact_data = {}

            # Extract email (prefer non-generic emails)
            contact_data['email'] = self._extract_specific_email(text)

            # Extract name (note: this may be blocked by OWA anti-scraping)
            contact_data['name'] = self._extract_name(text)

            # Extract phone number
            contact_data['phone'] = self._extract_phone(text)

            # Extract SIP address
            contact_data['sip'] = self._extract_sip(text)

            # Extract address
            contact_data['address'] = self._extract_address(text)

            # Extract structured work information
            self._extract_work_info(text, contact_data)

            contact_info = ContactInfo(**contact_data)

            if contact_info.is_valid():
                logger.debug(f"Text extraction found: {list(contact_data.keys())}")
                return contact_info

            return None

        except Exception as e:
            logger.debug(f"Text-based extraction failed: {e}")
            return None

    def _extract_specific_email(self, text: str) -> Optional[str]:
        """
        Extract specific (non-generic) email addresses.

        Filters out generic emails like ASP123@madrid.org and AGM456@madrid.org
        in favor of personal emails.
        """
        all_emails = self.patterns.EMAIL.findall(text)
        for email in all_emails:
            # Filter out generic institutional emails
            if not re.match(r'^(ASP|AGM|AEM|ADM)\d+@', email, re.I):
                return email.strip()

        # Return first email if no specific one found
        return all_emails[0].strip() if all_emails else None

    def _extract_name(self, text: str) -> Optional[str]:
        """
        Extract name in "SURNAME, NAME" format.

        Note: This may be blocked by Microsoft OWA anti-scraping protection.
        """
        lines = text.split('\n')

        # Try regex first
        for line in lines:
            match = self.patterns.NAME.search(line)
            if match:
                return match.group(1).strip()

        # Try heuristic approach
        for line in lines[:10]:  # Check first 10 lines
            line = line.strip()
            if (',' in line and len(line) > 5 and
                line not in ['CONTACTO', 'NOTAS', 'ORGANIZACIÓN'] and
                not line.startswith('C/') and
                re.match(r'^[A-ZÁÉÍÓÚÑ\s]+,\s*[A-ZÁÉÍÓÚÑ\s]+$', line, re.I)):
                return line

        return None

    def _extract_phone(self, text: str) -> Optional[str]:
        """
        Extract phone number with various heuristics.
        """
        lines = text.split('\n')

        # Prefer 9-digit numbers (Spanish format)
        for line in lines:
            if 'sip:' not in line.lower():
                # Look for 9-digit phone numbers
                match = re.search(r'\b\d{9}\b', line)
                if match:
                    return match.group(0)

        # Fall back to 6-8 digit numbers in work context
        if 'Trabajo' in text:
            for line in lines:
                if 'sip:' not in line.lower() and not re.search(r'\d{5}\s+[A-Z]', line):
                    match = re.search(r'\b\d{6,8}\b', line)
                    if match:
                        return match.group(0)

        return None

    def _extract_sip(self, text: str) -> Optional[str]:
        """Extract SIP address with validation."""
        match = self.patterns.SIP.search(text)
        if match:
            sip = match.group(0).strip()
            # Validate SIP format
            if re.match(r'^sip:[\w.+-]+@[\w.-]+\.[a-z]{2,}$', sip, re.I):
                return sip
        return None

    def _extract_address(self, text: str) -> Optional[str]:
        """Extract postal address with Spanish format validation."""
        # Try specific Spanish address format first
        match = re.search(
            r'C/\s*[A-ZÁÉÍÓÚÑ\s,]+\d+\s+\d{5}\s+[A-ZÁÉÍÓÚÑ\-\s]+',
            text, re.I
        )
        if match:
            return match.group(0).strip()

        # Fall back to postal code + city pattern
        match = self.patterns.POSTAL_ADDR.search(text)
        if match:
            return match.group(0).strip()

        return None

    def _extract_work_info(self, text: str, contact_data: Dict[str, str]):
        """
        Extract structured work information (department, company, office).

        Args:
            text: Popup text content
            contact_data: Dictionary to update with found information
        """
        # Look for labeled fields
        patterns = {
            'department': [
                r'Departamento:\s*([^\n]+)',
                r'Título:\s*([^\n]+)',
                r'Puesto:\s*([^\n]+)',
            ],
            'company': [
                r'Compañía:\s*([^\n]+)',
                r'Empresa:\s*([^\n]+)',
                r'Organización:\s*([^\n]+)',
            ],
            'office_location': [
                r'Oficina:\s*([^\n]+)',
                r'Ubicación:\s*([^\n]+)',
                r'Localización:\s*([^\n]+)',
            ]
        }

        for field, field_patterns in patterns.items():
            for pattern in field_patterns:
                match = re.search(pattern, text, re.I)
                if match:
                    value = match.group(1).strip()
                    if value and len(value) > 2:  # Avoid empty placeholders
                        contact_data[field] = value
                        break

        # Heuristic for office/job title (all-caps lines)
        if not contact_data.get('department'):
            lines = text.split('\n')
            for line in lines:
                line = line.strip()
                if (line and line.isupper() and len(line) > 3 and
                    line not in ['CONTACTO', 'NOTAS', 'ORGANIZACIÓN']):
                    contact_data['department'] = line
                    break