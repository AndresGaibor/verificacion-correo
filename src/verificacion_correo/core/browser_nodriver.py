"""
NoDriver-based browser automation for verificacion-correo.

This module reimplements the browser automation logic using NoDriver
(undetected Chrome) with integrated anti-detection techniques including:
- Mouse movement emulation
- Human typing patterns
- Random delays
- User-Agent rotation

This provides significantly better evasion of Microsoft OWA's anti-scraping
protections, particularly for extracting the contact name field.
"""

import re
import asyncio
from typing import Dict, List, Any, Optional
from dataclasses import dataclass

try:
    import nodriver as uc
except ImportError:
    uc = None

from verificacion_correo.core.config import Config
from verificacion_correo.core.extractor import ContactExtractor
from verificacion_correo.core.excel import ExcelReader, ExcelWriter, EmailRecord, ProcessingStatus
from verificacion_correo.core.antidetection import (
    NoDriverManager,
    MouseEmulator,
    TypingSimulator,
    DelayManager,
    UserAgentRotator,
    DelayConfig,
    MouseConfig,
    TypingConfig,
    UserAgentConfig,
)
from verificacion_correo.core.antidetection.nodriver_manager import NoDriverConfig
from verificacion_correo.utils.logging import get_logger


logger = get_logger(__name__)


@dataclass
class BatchResult:
    """Results from processing a batch of emails."""
    batch_number: int
    total_emails: int
    successful: int
    not_found: int
    errors: int
    records: List[EmailRecord]


@dataclass
class ProcessingStats:
    """Overall processing statistics."""
    total_batches: int
    total_emails: int
    successful: int
    not_found: int
    errors: int
    duration_seconds: float


class BrowserAutomationNoDriver:
    """
    NoDriver-based browser automation for OWA email processing.

    Uses advanced anti-detection techniques to evade Microsoft OWA's
    anti-scraping protections.
    """

    def __init__(self, config: Config):
        """
        Initialize NoDriver browser automation.

        Args:
            config: Application configuration
        """
        if uc is None:
            raise ImportError(
                "NoDriver not installed. Install with: pip install nodriver"
            )

        self.config = config
        self.contact_extractor = ContactExtractor(config)
        self.excel_writer = ExcelWriter(config.get_excel_file_path())

        # Initialize anti-detection components
        self._init_antidetection()

    def _init_antidetection(self):
        """Initialize anti-detection components based on configuration."""
        ad_cfg = self.config.antidetection

        # NoDriver manager
        nodriver_cfg = NoDriverConfig(
            headless=self.config.browser.headless,
            sandbox=True,
            lang="es-ES"
        )
        ua_cfg = UserAgentConfig(
            rotate=ad_cfg.ua_rotate,
            pool_size=ad_cfg.ua_pool_size,
            prefer_platform=ad_cfg.ua_prefer_platform
        )
        self.nodriver_manager = NoDriverManager(nodriver_cfg, ua_cfg)

        # Mouse emulator
        if ad_cfg.mouse_emulation:
            mouse_cfg = MouseConfig(
                bezier_curves=ad_cfg.mouse_bezier_curves,
                random_offset_px=ad_cfg.mouse_random_offset_px,
                move_duration_ms=(
                    ad_cfg.mouse_move_duration_ms_min,
                    ad_cfg.mouse_move_duration_ms_max
                )
            )
            self.mouse_emulator = MouseEmulator(mouse_cfg)
        else:
            self.mouse_emulator = None

        # Typing simulator
        if ad_cfg.human_typing:
            typing_cfg = TypingConfig(
                chars_per_second=(
                    ad_cfg.typing_chars_per_second_min,
                    ad_cfg.typing_chars_per_second_max
                ),
                mistake_probability=ad_cfg.typing_mistake_probability
            )
            self.typing_simulator = TypingSimulator(typing_cfg)
        else:
            self.typing_simulator = None

        # Delay manager
        if ad_cfg.random_delays:
            delay_cfg = DelayConfig(
                between_actions=(
                    ad_cfg.delay_between_actions_min,
                    ad_cfg.delay_between_actions_max
                ),
                between_emails=(
                    ad_cfg.delay_between_emails_min,
                    ad_cfg.delay_between_emails_max
                )
            )
            self.delay_manager = DelayManager(delay_cfg)
        else:
            self.delay_manager = DelayManager()  # Use defaults

    async def process_emails(self, excel_file: Optional[str] = None) -> ProcessingStats:
        """
        Process all pending emails from Excel file using NoDriver.

        Args:
            excel_file: Path to Excel file (optional, uses config default)

        Returns:
            ProcessingStats with comprehensive results
        """
        import time
        start_time = time.time()

        # Setup Excel reader
        excel_path = excel_file or self.config.get_excel_file_path()
        excel_reader = ExcelReader(
            excel_path,
            start_row=self.config.excel.start_row,
            email_column=self.config.excel.email_column
        )

        # Read pending emails
        logger.info(f"Reading pending emails from {excel_path}")
        summary = excel_reader.read_pending_emails(batch_size=self.config.processing.batch_size)

        if not summary.batches:
            logger.info("No pending emails to process")
            return ProcessingStats(0, 0, 0, 0, 0, 0.0)

        logger.info(
            f"Found {summary.pending_count} pending emails, "
            f"organized into {len(summary.batches)} batches"
        )

        # Process all batches
        total_stats = ProcessingStats(0, 0, 0, 0, 0, 0.0)
        batch_results = []

        # Start NoDriver browser
        try:
            await self.nodriver_manager.start(
                session_file=self.config.get_session_file_path()
            )
            page = await self.nodriver_manager.get_page()

            logger.info("NoDriver browser started with anti-detection enabled")

            # Process each batch
            for batch_num, batch_records in enumerate(summary.batches, 1):
                result = await self._process_batch(page, batch_records, batch_num)
                batch_results.append(result)

                # Update running totals
                total_stats.total_batches = batch_num
                total_stats.total_emails += result.total_emails
                total_stats.successful += result.successful
                total_stats.not_found += result.not_found
                total_stats.errors += result.errors

                # Write batch results to Excel
                self.excel_writer.write_batch_results(result.records)

                logger.info(f"Batch {batch_num}/{len(summary.batches)} completed")

                # Delay between batches
                if batch_num < len(summary.batches):
                    delay = self.delay_manager.between_emails()
                    logger.debug(f"Waiting {delay:.2f}s before next batch")
                    await self.delay_manager.sleep_async(delay)

        finally:
            await self.nodriver_manager.close()

        total_stats.duration_seconds = time.time() - start_time
        self._log_final_results(total_stats)

        return total_stats

    async def _process_batch(
        self,
        page,
        email_records: List[EmailRecord],
        batch_number: int
    ) -> BatchResult:
        """
        Process a single batch of email records.

        Args:
            page: NoDriver page object
            email_records: List of EmailRecord objects to process
            batch_number: Current batch number for logging

        Returns:
            BatchResult with processing outcomes
        """
        logger.info(f"Processing batch {batch_number}: {len(email_records)} emails")

        # Initialize batch statistics
        batch_result = BatchResult(
            batch_number=batch_number,
            total_emails=len(email_records),
            successful=0,
            not_found=0,
            errors=0,
            records=[]
        )

        try:
            # Navigate to OWA and open new message
            await self._navigate_and_open_message(page)

            # Add all emails to the "To" field
            await self._add_emails_to_field(page, [r.email for r in email_records])

            # Process each email individually
            for i, record in enumerate(email_records):
                await self._process_single_email(page, record, batch_result)

                # Delay between emails within batch
                if i < len(email_records) - 1:
                    delay = self.delay_manager.between_actions()
                    await self.delay_manager.sleep_async(delay)

            # Close the message without saving
            await self._close_message(page)

        except Exception as e:
            logger.error(f"Error processing batch {batch_number}: {e}")
            # Mark remaining unprocessed records as error
            for record in email_records:
                if record.status == ProcessingStatus.PENDING:
                    record.status = ProcessingStatus.ERROR
                    batch_result.errors += 1
                    batch_result.records.append(record)

        return batch_result

    async def _navigate_and_open_message(self, page):
        """Navigate to OWA and open new message."""
        # Navigate if not already at OWA
        current_url = page.url
        if (current_url != self.config.page_url and
                not current_url.startswith(self.config.page_url.split('#')[0])):
            await page.get(self.config.page_url)
            await asyncio.sleep(2)  # Wait for page load

        # Click new message button (with mouse emulation if enabled)
        try:
            new_msg_btn = await page.select(self.config.selectors.new_message_btn)

            if self.mouse_emulator and new_msg_btn:
                await self.mouse_emulator.click_element_async(page, new_msg_btn)
            else:
                await new_msg_btn.click()

            await asyncio.sleep(self.config.wait_times.after_new_message / 1000.0)

        except Exception as e:
            logger.error(f"Error opening new message: {e}")
            raise

    async def _add_emails_to_field(self, page, emails: List[str]):
        """Add email addresses to the 'To' field."""
        emails_str = ";".join(emails)

        try:
            # Find To field
            input_boxes = await page.select_all('[role="textbox"]')
            to_field = None

            for box in input_boxes:
                name = await box.get_attribute('name')
                if name and self.config.selectors.to_field_name.lower() in name.lower():
                    to_field = box
                    break

            if not to_field:
                logger.error("Could not find 'To' field")
                raise ValueError("To field not found")

            # Type emails (with human typing if enabled)
            if self.typing_simulator:
                await self.typing_simulator.fill_with_typing_async(to_field, emails_str)
            else:
                await to_field.send_keys(emails_str)

            await asyncio.sleep(self.config.wait_times.after_fill_to / 1000.0)

            # Blur field
            await page.evaluate("document.activeElement.blur()")
            await asyncio.sleep(self.config.wait_times.after_blur / 1000.0)

        except Exception as e:
            logger.error(f"Error adding emails to field: {e}")
            raise

    async def _process_single_email(
        self,
        page,
        record: EmailRecord,
        batch_result: BatchResult
    ):
        """Process a single email record."""
        try:
            logger.debug(f"Processing email: {record.email}")

            # Find the email token
            email_span = await self._find_email_token(page, record.email)
            if not email_span:
                record.status = ProcessingStatus.ERROR
                batch_result.errors += 1
                logger.debug(f"Token not found for: {record.email}")
                batch_result.records.append(record)
                return

            # Click on the token (with mouse emulation if enabled)
            if self.mouse_emulator:
                await self.mouse_emulator.click_element_async(page, email_span)
            else:
                await email_span.click()

            await asyncio.sleep(self.config.wait_times.after_click_token / 1000.0)

            # Extract contact information
            # Note: For NoDriver, we might need to adapt the extractor
            # For now, we'll convert to a compatible format
            contact_info = await self._extract_contact_info_async(page)

            # Determine processing status
            if contact_info and self._is_valid_contact(contact_info):
                record.data = contact_info.to_dict()
                record.status = ProcessingStatus.SUCCESS
                batch_result.successful += 1
                logger.debug(f"Successfully extracted info for: {record.email}")
            else:
                record.status = ProcessingStatus.NOT_FOUND
                batch_result.not_found += 1
                logger.debug(f"No valid info found for: {record.email}")

            # Close popup (ESC key)
            await page.send_keys('\x1b')  # ESC key
            await asyncio.sleep(self.config.wait_times.after_close_popup / 1000.0)

        except Exception as e:
            record.status = ProcessingStatus.ERROR
            batch_result.errors += 1
            logger.error(f"Error processing {record.email}: {e}")

        finally:
            batch_result.records.append(record)

    async def _find_email_token(self, page, email: str):
        """Find the email token in the interface."""
        try:
            # Find all spans
            all_spans = await page.select_all('span')

            # Filter by text content
            for span in all_spans:
                try:
                    text = await span.text_content()
                    if text and text.strip().lower() == email.lower():
                        return span
                except:
                    continue

            return None

        except Exception as e:
            logger.debug(f"Error finding email token for {email}: {e}")
            return None

    async def _extract_contact_info_async(self, page):
        """
        Extract contact information from popup (async version).

        Adapts the sync ContactExtractor to work with NoDriver's async API.
        """
        try:
            # Wait for popup to be visible
            popup_selector = self.config.selectors.popup
            popup = await page.select(popup_selector)

            if not popup:
                await asyncio.sleep(1)
                popup = await page.select(popup_selector)

            if not popup:
                logger.warning("Popup not found")
                return None

            # Get popup text content
            popup_text = await popup.text_content()

            # Use text-based extraction (most reliable for NoDriver)
            from .extractor import ContactInfo
            contact_info = self.contact_extractor._extract_text_based(popup_text)

            return contact_info

        except Exception as e:
            logger.error(f"Error extracting contact info: {e}")
            return None

    def _is_valid_contact(self, contact_info) -> bool:
        """
        Determine if extracted contact information is valid.

        Uses SIP presence as primary validation criterion.
        """
        # Primary validation: SIP address
        if contact_info.sip and contact_info.sip.strip():
            return True

        # Secondary validation: personal email
        if (contact_info.email and contact_info.email.strip() and
                not re.match(r'^(ASP|AGM|AEM|ADM)\d+@', contact_info.email, re.I)):
            return True

        # Tertiary validation: phone number
        if contact_info.phone and contact_info.phone.strip():
            return True

        return False

    async def _close_message(self, page):
        """Close the message without saving."""
        try:
            await asyncio.sleep(self.config.wait_times.before_discard / 1000.0)

            # Find and click discard button
            discard_btn = await page.select(self.config.selectors.discard_btn)
            if discard_btn:
                await discard_btn.click()
                await asyncio.sleep(1)

                # Confirm discard (second click if dialog appears)
                discard_btn2 = await page.select(self.config.selectors.discard_btn)
                if discard_btn2:
                    await discard_btn2.click()

        except Exception as e:
            logger.debug(f"Error closing message: {e}")

    def _log_final_results(self, stats: ProcessingStats):
        """Log final processing results."""
        logger.info("=" * 60)
        logger.info("PROCESSING COMPLETED (NoDriver Anti-Detection)")
        logger.info("=" * 60)
        logger.info(f"Total Batches: {stats.total_batches}")
        logger.info(f"Total Emails: {stats.total_emails}")
        logger.info(f"Successful: {stats.successful} ({stats.successful/stats.total_emails*100:.1f}%)")
        logger.info(f"Not Found: {stats.not_found} ({stats.not_found/stats.total_emails*100:.1f}%)")
        logger.info(f"Errors: {stats.errors} ({stats.errors/stats.total_emails*100:.1f}%)")
        logger.info(f"Duration: {stats.duration_seconds:.1f} seconds")
        logger.info("=" * 60)


# Async entry point
async def process_emails_async(config: Optional[Config] = None) -> ProcessingStats:
    """
    Process emails with NoDriver (async entry point).

    Args:
        config: Application configuration (optional)

    Returns:
        ProcessingStats with results
    """
    if config is None:
        from .config import get_config
        config = get_config()

    automation = BrowserAutomationNoDriver(config)
    return await automation.process_emails()


# Sync wrapper for compatibility
def process_emails_nodriver(config: Optional[Config] = None) -> ProcessingStats:
    """
    Process emails with NoDriver (sync wrapper).

    Args:
        config: Application configuration (optional)

    Returns:
        ProcessingStats with results
    """
    return asyncio.run(process_emails_async(config))
