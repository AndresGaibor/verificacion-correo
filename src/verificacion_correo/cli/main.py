"""
Command Line Interface for verificacion-correo.

This module provides a comprehensive CLI for email verification and contact
extraction from OWA, with multiple operation modes and options.
"""

import argparse
import sys
import json
from pathlib import Path
from typing import Optional

from verificacion_correo.core.config import get_config, Config
from verificacion_correo.core.first_run import check_and_run_first_time_setup
from verificacion_correo.core.browser import BrowserAutomation, process_emails
from verificacion_correo.core.session import setup_session_interactive, validate_saved_session, get_session_status
from verificacion_correo.core.excel import ExcelReader
from verificacion_correo.core.gal_scraper import scrape_gal
from verificacion_correo.utils.logging import setup_logging, get_logger


logger = get_logger(__name__)


class VerificacionCorreoCLI:
    """Main CLI class for verificacion-correo."""

    def __init__(self):
        """Initialize CLI with default configuration."""
        self.config: Optional[Config] = None
        self.is_first_run: bool = False

    def run(self, args=None):
        """
        Run the CLI with provided arguments.

        Args:
            args: Command line arguments (optional, uses sys.argv if None)
        """
        parser = self._create_parser()
        parsed_args = parser.parse_args(args)

        # Setup logging level
        log_level = "DEBUG" if parsed_args.verbose else "INFO"
        setup_logging(level=log_level, log_file=parsed_args.log_file)

        # Load configuration with first-run setup
        try:
            self.config = check_and_run_first_time_setup(parsed_args.config)
            if parsed_args.config and not self.is_first_run:
                # Override config file if specified
                self.config = Config(parsed_args.config)
        except Exception as e:
            print(f"Error loading configuration: {e}")
            return 1

        # Execute the requested command
        try:
            return parsed_args.func(parsed_args)
        except KeyboardInterrupt:
            print("\nOperation cancelled by user")
            return 130
        except Exception as e:
            logger.error(f"Error: {e}")
            if parsed_args.verbose:
                import traceback
                traceback.print_exc()
            return 1

    def _create_parser(self) -> argparse.ArgumentParser:
        """Create the argument parser."""
        parser = argparse.ArgumentParser(
            prog="verificacion-correo",
            description="Verificación de correos - Herramienta de extracción de contactos OWA",
            formatter_class=argparse.RawDescriptionHelpFormatter,
            epilog="""
Examples:
  %(prog)s                          # Process pending emails with default settings
  %(prog)s --excel-file emails.xlsx  # Use specific Excel file
  %(prog)s --batch-size 5           # Process in smaller batches
  %(prog)s --dry-run               # Show what would be processed without executing
  %(prog)s setup                   # Set up browser session
  %(prog)s validate                # Validate setup and show status
  %(prog)s status                  # Show current session and file status
            """
        )

        # Global options
        parser.add_argument(
            "-v", "--verbose",
            action="store_true",
            help="Enable verbose output"
        )
        parser.add_argument(
            "--log-file",
            type=str,
            help="Log file path (default: console only)"
        )
        parser.add_argument(
            "--config",
            type=str,
            help="Configuration file path (default: config/default.yaml)"
        )

        # Add process options to main parser for default command
        self._add_process_options(parser)

        # Create subparsers for commands
        subparsers = parser.add_subparsers(
            title="commands",
            description="Available commands",
            help="Use '%(prog)s <command> --help' for command-specific help",
            dest="command"
        )

        # Process command (default)
        process_parser = subparsers.add_parser(
            "process",
            help="Process pending emails (default command)"
        )
        self._add_process_options(process_parser)
        process_parser.set_defaults(func=self._cmd_process)

        # Setup command
        setup_parser = subparsers.add_parser(
            "setup",
            help="Set up browser session for automation"
        )
        setup_parser.set_defaults(func=self._cmd_setup)

        # Validate command
        validate_parser = subparsers.add_parser(
            "validate",
            help="Validate configuration and setup"
        )
        validate_parser.set_defaults(func=self._cmd_validate)

        # Status command
        status_parser = subparsers.add_parser(
            "status",
            help="Show current status information"
        )
        status_parser.set_defaults(func=self._cmd_status)

        # Scrape gallery command
        gallery_parser = subparsers.add_parser(
            "scrape-gallery",
            help="Extract full GAL directory via OWA FindPeople API"
        )
        gallery_parser.add_argument(
            "--output-dir",
            type=str,
            default="data/gal",
            help="Output directory (default: data/gal)"
        )
        gallery_parser.add_argument(
            "--max-contacts",
            type=int,
            default=0,
            help="Max contacts to extract (0 = unlimited)"
        )
        gallery_parser.add_argument(
            "--batch-size",
            type=int,
            default=100,
            help="Entries per API call (default: 100)"
        )
        gallery_parser.add_argument(
            "--delay",
            type=float,
            default=8.0,
            help="Seconds between requests (default: 8.0)"
        )
        gallery_parser.add_argument(
            "--force-restart",
            action="store_true",
            help="Ignore saved progress and start from scratch"
        )
        gallery_parser.set_defaults(func=self._cmd_scrape_gallery)

        # Default command is 'process'
        parser.set_defaults(func=lambda args: self._cmd_process(args))

        return parser

    def _add_process_options(self, parser):
        """Add options for the process command."""
        parser.add_argument(
            "--excel-file",
            type=str,
            help="Excel file path (default: from config)"
        )
        parser.add_argument(
            "--batch-size",
            type=int,
            help="Batch size for processing (default: from config)"
        )
        parser.add_argument(
            "--dry-run",
            action="store_true",
            help="Show what would be processed without executing"
        )
        parser.add_argument(
            "--force",
            action="store_true",
            help="Force processing even if session validation fails"
        )
        parser.add_argument(
            "--keep-draft",
            action="store_true",
            help="Keep draft message instead of discarding it (overrides config)"
        )

    def _cmd_process(self, args):
        """Process pending emails command."""
        print("=" * 70)
        print("VERIFICACIÓN DE CORREOS - SISTEMA INCREMENTAL")
        print("=" * 70)

        # Override config with command line options
        if hasattr(args, 'excel_file') and args.excel_file:
            self.config.excel.default_file = Path(args.excel_file).absolute()
        if hasattr(args, 'batch_size') and args.batch_size:
            self.config.processing.batch_size = args.batch_size
        if hasattr(args, 'keep_draft') and args.keep_draft:
            self.config.processing.discard_draft = False

        # Validate setup
        validation = self._validate_setup()
        if not validation['session_valid'] and not getattr(args, 'force', False):
            print("❌ Browser session invalid or expired")
            print("   Run 'verificacion-correo setup' to create a new session")
            return 1

        # Check Excel file
        if not validation['excel_file_exists']:
            print(f"❌ Excel file not found: {self.config.get_excel_file_path()}")
            print("   Create an Excel file with emails in column A")
            return 1

        # Read pending emails
        excel_reader = ExcelReader(
            self.config.get_excel_file_path(),
            start_row=self.config.excel.start_row,
            email_column=self.config.excel.email_column
        )

        summary = excel_reader.read_pending_emails(batch_size=self.config.processing.batch_size)

        # Show summary
        print(f"\n📁 Excel file: {self.config.get_excel_file_path()}")
        print(f"📊 Total emails: {summary.total_emails}")
        print(f"✅ Already processed: {summary.processed_count}")
        print(f"⏳ Pending: {summary.pending_count}")

        if summary.pending_count == 0:
            print("\n✅ No pending emails to process")
            print("   To re-process an email, clear its 'Status' column in Excel")
            return 0

        print(f"📦 Organized in {len(summary.batches)} batches of {self.config.processing.batch_size}")

        if getattr(args, 'dry_run', False):
            print("\n🔍 DRY RUN MODE - No processing will be performed")
            print("   First 5 pending emails:")
            for i, batch in enumerate(summary.batches[:1]):
                for j, record in enumerate(batch.records[:5]):
                    print(f"   {j+1}. {record.email}")
            return 0

        # Confirm processing
        try:
            input(f"\nPress ENTER to start processing (or Ctrl+C to cancel)...")
        except KeyboardInterrupt:
            print("\nProcessing cancelled by user")
            return 130

        # Process emails
        print("\n🚀 Starting processing...")
        automation = BrowserAutomation(self.config)
        stats = automation.process_emails(self.config.get_excel_file_path())

        # Show final results
        self._show_final_results(stats)
        return 0

    def _cmd_setup(self, args):
        """Set up browser session command."""
        print("=" * 50)
        print("BROWSER SESSION SETUP")
        print("=" * 50)

        print(f"🌐 Target URL: {self.config.page_url}")
        print(f"💾 Session file: {self.config.get_session_file_path()}")

        success = setup_session_interactive(self.config)
        if success:
            print("\n✅ Session setup completed successfully")
            print("   You can now run 'verificacion-correo' to process emails")
            return 0
        else:
            print("\n❌ Session setup failed")
            return 1

    def _cmd_validate(self, args):
        """Validate configuration and setup command."""
        print("=" * 50)
        print("VALIDATION RESULTS")
        print("=" * 50)

        validation = self._validate_setup()

        print(f"📄 Configuration: {'✅' if validation['config_valid'] else '❌'}")
        print(f"🌐 Browser session: {'✅' if validation['session_valid'] else '❌'}")
        print(f"📁 Excel file: {'✅' if validation['excel_file_exists'] else '❌'}")
        print(f"📊 Excel readable: {'✅' if validation['excel_file_readable'] else '❌'}")
        print(f"📧 Pending emails: {validation['pending_emails']}")

        if validation['issues']:
            print("\n⚠️ Issues found:")
            for issue in validation['issues']:
                print(f"   - {issue}")

        if all([
            validation['config_valid'],
            validation['session_valid'],
            validation['excel_file_exists'],
            validation['excel_file_readable']
        ]):
            print("\n✅ Everything looks good!")
            return 0
        else:
            print("\n❌ Some issues need to be resolved")
            return 1

    def _cmd_status(self, args):
        """Show status information command."""
        print("=" * 50)
        print("STATUS INFORMATION")
        print("=" * 50)

        # Session status
        session_info = get_session_status(self.config)
        print(f"🌐 Session file: {session_info.get('file_path', 'N/A')}")

        if session_info.get('exists'):
            print(f"   File size: {session_info.get('file_size', 0)} bytes")
            print(f"   Cookies: {session_info.get('cookies_count', 0)}")
            print(f"   Origins: {session_info.get('origins_count', 0)}")
            print(f"   Valid: {'✅' if session_info.get('is_valid') else '❌'}")
        else:
            print("   Status: Not found")

        # Excel file status
        excel_path = self.config.get_excel_file_path()
        excel_file = Path(excel_path)
        print(f"\n📁 Excel file: {excel_path}")

        if excel_file.exists():
            stat = excel_file.stat()
            print(f"   File size: {stat.st_size} bytes")
            print(f"   Modified: {stat.st_mtime}")

            try:
                excel_reader = ExcelReader(excel_path)
                summary = excel_reader.read_pending_emails()
                print(f"   Total emails: {summary.total_emails}")
                print(f"   Pending: {summary.pending_emails}")
            except Exception as e:
                print(f"   Error reading: {e}")
        else:
            print("   Status: Not found")

        return 0

    def _cmd_scrape_gallery(self, args):
        """Scrape full GAL directory command."""
        print("=" * 70)
        print("EXTRACCIÓN DEL DIRECTORIO (GAL) VÍA API")
        print("=" * 70)

        session_file = self.config.get_session_file_path()
        if not Path(session_file).exists():
            print("❌ No hay sesión guardada")
            print("   Ejecute 'verificacion-correo setup' primero")
            return 1

        output_dir = args.output_dir
        max_contacts = args.max_contacts
        force_restart = args.force_restart

        # Check progress file for resume
        progress_file = Path(output_dir) / "gal_progress.json"
        if progress_file.exists() and not force_restart:
            with open(progress_file) as f:
                prev = json.load(f)
            print(f"\n⏳ Progreso anterior encontrado:")
            print(f"   Offset: {prev.get('offset', 0)}")
            print(f"   Contactos: {prev.get('count', 0)}")
            print(f"   Última actualización: {prev.get('last_update', 'N/A')}")
            try:
                resp = input("\n¿Reanudar? [S/n]: ").strip().lower()
                if resp in ("n", "no"):
                    force_restart = True
            except (EOFError, KeyboardInterrupt):
                pass
        else:
            print(f"\n🆕 Nueva extracción desde cero")

        print(f"\n📁 Directorio salida: {output_dir}")
        print(f"📦 Batch size: {args.batch_size}")
        print(f"⏱️  Delay: {args.delay}s")
        if max_contacts > 0:
            print(f"🎯 Máx contactos: {max_contacts}")
        else:
            print(f"🎯 Máx contactos: ilimitado")
        if force_restart:
            print("🔄 Forzando reinicio (ignorando progreso)")

        try:
            input("\nPresione ENTER para comenzar (Ctrl+C para cancelar)...")
        except KeyboardInterrupt:
            print("\nCancelado")
            return 130

        print("\n🚀 Iniciando extracción...\n")

        stats = scrape_gal(
            session_file=session_file,
            output_dir=output_dir,
            max_contacts=max_contacts,
            batch_size=args.batch_size,
            request_delay=args.delay,
            force_restart=force_restart,
        )

        print("\n" + "=" * 70)
        print("EXTRACCIÓN COMPLETADA")
        print("=" * 70)
        print(f"📧 Total contactos: {stats['total']}")
        print(f"⏱️  Duración: {stats['duration']:.1f}s")

        if stats['expired']:
            print("⚠️  Sesión expirada — progreso guardado para reanudar")
        elif stats['stopped']:
            print("⏹️  Detenido por usuario")

        files = stats.get('files', {})
        print(f"📄 JSON: {files.get('json', 'N/A')}")
        print(f"📄 CSV: {files.get('csv', 'N/A')}")
        print(f"💾 Progreso: {files.get('progress', 'N/A')}")
        print("=" * 70)

        return 0

    def _validate_setup(self):
        """Validate the current setup."""
        automation = BrowserAutomation(self.config)
        return automation.validate_setup()

    def _show_final_results(self, stats):
        """Display final processing results."""
        print("\n" + "=" * 70)
        print("PROCESSING COMPLETED")
        print("=" * 70)
        print(f"📧 Total emails processed: {stats.total_emails}")
        print(f"✅ Successful: {stats.successful} ({stats.successful/stats.total_emails*100:.1f}%)")
        print(f"❌ Not found: {stats.not_found} ({stats.not_found/stats.total_emails*100:.1f}%)")
        print(f"⚠️ Errors: {stats.errors} ({stats.errors/stats.total_emails*100:.1f}%)")
        print(f"⏱️ Duration: {stats.duration_seconds:.1f} seconds")
        print(f"💾 Results saved to: {self.config.get_excel_file_path()}")
        print("=" * 70)


def main(args=None):
    """
    Main entry point for the CLI.

    Args:
        args: Command line arguments (optional, uses sys.argv if None)

    Returns:
        Exit code (0 for success, non-zero for error)
    """
    cli = VerificacionCorreoCLI()
    return cli.run(args)


if __name__ == "__main__":
    sys.exit(main())