"""
Browser automation module for verificacion-correo.

This module provides the main automation logic for processing email addresses
through the OWA interface, with batch processing and error handling.
"""

import re
from typing import Dict, List, Any, Optional
from dataclasses import dataclass
from playwright.sync_api import sync_playwright, Page, TimeoutError as PWTimeout

from .config import Config
from .extractor import ContactExtractor
from .session import SessionManager
from .excel import ExcelReader, ExcelWriter, EmailRecord, ProcessingStatus
from ..utils.logging import get_logger


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


class BrowserAutomation:
    """
    Main automation class for OWA email processing.

    Orchestrates browser automation, contact extraction, and result storage
    with proper error handling and progress tracking.
    """

    def __init__(self, config: Config):
        """
        Initialize browser automation.

        Args:
            config: Application configuration
        """
        self.config = config
        self.session_manager = SessionManager(config)
        self.contact_extractor = ContactExtractor(config)
        self.excel_writer = ExcelWriter(config.get_excel_file_path())

    def process_emails(self, excel_file: Optional[str] = None) -> ProcessingStats:
        """
        Process all pending emails from Excel file.

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
            f"Found {summary.pending_count} pending emails in {summary.total_emails} total, "
            f"organized into {len(summary.batches)} batches"
        )

        # Process all batches
        total_stats = ProcessingStats(0, 0, 0, 0, 0, 0.0)
        batch_results = []

        with sync_playwright() as p:
            # Setup browser context
            context = self._create_automation_context(p)
            if not context:
                raise RuntimeError("Failed to create browser context")

            try:
                # Process each batch
                for batch_num, batch_records in enumerate(summary.batches, 1):
                    result = self._process_batch(context, batch_records, batch_num)
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

            finally:
                context.close()

        total_stats.duration_seconds = time.time() - start_time
        self._log_final_results(total_stats)

        return total_stats

    def _create_automation_context(self, playwright):
        """Create browser context with session for automation."""
        context = self.session_manager.create_automation_context()
        if not context:
            raise RuntimeError("Failed to create automation context")
        return context

    def _process_batch(self, context, email_records: List[EmailRecord], batch_number: int) -> BatchResult:
        """
        Process a single batch of email records.

        Args:
            context: Browser context
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

        # Get or create page
        page = self._get_or_create_page(context)
        if not page:
            # Mark all as error if page creation failed
            for record in email_records:
                record.status = ProcessingStatus.ERROR
                batch_result.errors += 1
                batch_result.records.append(record)
            return batch_result

        try:
            # Navigate to OWA and open new message
            self._navigate_and_open_message(page)

            # Add all emails to the "To" field
            self._add_emails_to_field(page, [r.email for r in email_records])

            # Process each email individually
            for record in email_records:
                self._process_single_email(page, record, batch_result)

            # Close the message without saving
            self._close_message(page)

        except Exception as e:
            logger.error(f"Error processing batch {batch_number}: {e}")
            # Mark remaining unprocessed records as error
            for record in email_records:
                if record.status == ProcessingStatus.PENDING:
                    record.status = ProcessingStatus.ERROR
                    batch_result.errors += 1
                    batch_result.records.append(record)

        return batch_result

    def _get_or_create_page(self, context) -> Optional[Page]:
        """Get existing page or create new one."""
        try:
            if context.pages:
                return context.pages[0]
            else:
                return context.new_page()
        except Exception as e:
            logger.error(f"Error creating page: {e}")
            return None

    def _navigate_and_open_message(self, page: Page):
        """Navigate to OWA and open new message."""
        # Navigate if not already at OWA
        if (page.url != self.config.page_url and
                not page.url.startswith(self.config.page_url.split('#')[0])):
            page.goto(self.config.page_url)
            page.wait_for_load_state("networkidle")

        # Click new message button
        page.click(self.config.selectors.new_message_btn)
        page.wait_for_timeout(self.config.wait_times.after_new_message)

    def _add_emails_to_field(self, page: Page, emails: List[str]):
        """Add email addresses to the 'To' field."""
        emails_str = ";".join(emails)
        input_box = page.get_by_role(
            self.config.selectors.to_field_role,
            name=self.config.selectors.to_field_name
        )
        input_box.fill(emails_str)
        page.wait_for_timeout(self.config.wait_times.after_fill_to)
        input_box.blur()
        page.wait_for_timeout(self.config.wait_times.after_blur)

    def _process_single_email(self, page: Page, record: EmailRecord, batch_result: BatchResult):
        """Process a single email record."""
        try:
            logger.debug(f"Processing email: {record.email}")

            # Find the email token
            email_span = self._find_email_token(page, record.email)
            if not email_span:
                record.status = ProcessingStatus.ERROR
                batch_result.errors += 1
                logger.debug(f"Token not found for: {record.email}")
                return

            # Click on the token
            email_span.click(timeout=3000)
            page.wait_for_timeout(self.config.wait_times.after_click_token)

            # Extract contact information
            contact_info = self.contact_extractor.extract_from_popup(page)

            # Determine processing status based on extracted information
            if contact_info and self._is_valid_contact(contact_info):
                record.data = contact_info.to_dict()
                record.status = ProcessingStatus.SUCCESS
                batch_result.successful += 1
                logger.debug(f"Successfully extracted info for: {record.email}")
            else:
                record.status = ProcessingStatus.NOT_FOUND
                batch_result.not_found += 1
                logger.debug(f"No valid info found for: {record.email}")

            # Close popup
            page.keyboard.press("Escape")
            page.wait_for_timeout(self.config.wait_times.after_close_popup)

        except Exception as e:
            record.status = ProcessingStatus.ERROR
            batch_result.errors += 1
            logger.error(f"Error processing {record.email}: {e}")

        finally:
            batch_result.records.append(record)

    def _find_email_token(self, page: Page, email: str):
        """Find the email token in the interface."""
        try:
            # Create a regex pattern for exact email match
            pattern = re.compile(f'^{re.escape(email)}$', re.I)

            # Find spans that contain the email text
            email_span = page.locator(f'span:has-text("{email}")').filter(
                has_text=pattern
            )

            if email_span.count() > 0:
                return email_span.first
            else:
                return None

        except Exception as e:
            logger.debug(f"Error finding email token for {email}: {e}")
            return None

    def _is_valid_contact(self, contact_info) -> bool:
        """
        Determine if extracted contact information is valid.

        Uses SIP presence as primary validation criterion since it's the most
        reliable indicator of a valid contact in OWA.
        """
        # Primary validation: SIP address indicates valid contact
        if contact_info.sip and contact_info.sip.strip():
            return True

        # Secondary validation: personal email indicates valid contact
        if (contact_info.email and contact_info.email.strip() and
                not re.match(r'^(ASP|AGM|AEM|ADM)\d+@', contact_info.email, re.I)):
            return True

        # Tertiary validation: phone number indicates valid contact
        if contact_info.phone and contact_info.phone.strip():
            return True

        return False

    def _close_message(self, page: Page):
        """Close the message without saving."""
        try:
            page.wait_for_timeout(self.config.wait_times.before_discard)
            page.click(self.config.selectors.discard_btn, timeout=2000)
            page.wait_for_timeout(1000)
            # Confirm discard (second click if needed)
            page.click(self.config.selectors.discard_btn, timeout=2000)
        except PWTimeout:
            logger.debug("Timeout while closing message")
        except Exception as e:
            logger.debug(f"Error closing message: {e}")

    def _log_final_results(self, stats: ProcessingStats):
        """Log final processing results."""
        logger.info("=" * 60)
        logger.info("PROCESSING COMPLETED")
        logger.info("=" * 60)
        logger.info(f"Total Batches: {stats.total_batches}")
        logger.info(f"Total Emails: {stats.total_emails}")
        logger.info(f"Successful: {stats.successful} ({stats.successful/stats.total_emails*100:.1f}%)")
        logger.info(f"Not Found: {stats.not_found} ({stats.not_found/stats.total_emails*100:.1f}%)")
        logger.info(f"Errors: {stats.errors} ({stats.errors/stats.total_emails*100:.1f}%)")
        logger.info(f"Duration: {stats.duration_seconds:.1f} seconds")
        logger.info("=" * 60)

    def validate_setup(self) -> Dict[str, Any]:
        """
        Validate the complete setup before processing.

        Returns:
            Dictionary with validation results
        """
        validation_results = {
            'config_valid': len(self.config.validate()) == 0,
            'session_valid': False,
            'excel_file_exists': False,
            'excel_file_readable': False,
            'pending_emails': 0,
            'issues': []
        }

        # Validate configuration
        config_issues = self.config.validate()
        if config_issues:
            validation_results['issues'].extend([f"Config: {issue}" for issue in config_issues])

        # Validate session
        validation_results['session_valid'] = self.session_manager.validate_session()
        if not validation_results['session_valid']:
            validation_results['issues'].append("Browser session invalid or expired")

        # Validate Excel file
        excel_path = self.config.get_excel_file_path()
        validation_results['excel_file_exists'] = Path(excel_path).exists()
        if validation_results['excel_file_exists']:
            try:
                excel_reader = ExcelReader(excel_path)
                summary = excel_reader.read_pending_emails()
                validation_results['pending_emails'] = summary.pending_count
                validation_results['excel_file_readable'] = True
            except Exception as e:
                validation_results['issues'].append(f"Excel file error: {e}")
        else:
            validation_results['issues'].append(f"Excel file not found: {excel_path}")

        return validation_results


# Convenience function for simple usage
def process_emails(config: Optional[Config] = None) -> ProcessingStats:
    """
    Convenience function to process emails with default configuration.

    Args:
        config: Application configuration (optional)

    Returns:
        ProcessingStats with results
    """
    if config is None:
        from .config import get_config
        config = get_config()

    automation = BrowserAutomation(config)
    return automation.process_emails()