"""
Main GUI module for verificacion-correo.

This module provides a modern Tkinter-based graphical user interface for
email verification and contact extraction from OWA.
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
import queue
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, Any, Callable
import yaml
import asyncio
import sys
import json

from verificacion_correo.core.config import Config
from verificacion_correo.core.browser import BrowserAutomation
from verificacion_correo.core.session import SessionManager, get_session_status
from verificacion_correo.core.excel import ExcelReader, ExcelWriter
from verificacion_correo.core.api_extractor import process_emails_via_api
from verificacion_correo.core.gal_scraper import scrape_gal
from verificacion_correo.utils.logging import setup_logging, get_logger


logger = get_logger(__name__)


class GUIService:
    """Background service for GUI operations."""

    def __init__(self, config: Config):
        """Initialize GUI service."""
        self.config = config
        self.session_manager = SessionManager(config)
        self.progress_queue = queue.Queue()
        self.current_thread: Optional[threading.Thread] = None
        self.is_processing = False
        self.should_stop = False
        self._gal_stop_flag: dict = {'stop': False}

    def validate_session(self) -> Dict[str, Any]:
        """Validate browser session."""
        return get_session_status(self.config)

    def setup_session(self, progress_callback: Callable = None) -> bool:
        """Set up browser session interactively."""
        try:
            return self.session_manager.setup_interactive_session()
        except Exception as e:
            logger.error(f"Session setup error: {e}")
            return False

    def get_excel_summary(self, excel_path: str) -> Dict[str, Any]:
        """Get Excel file summary."""
        try:
            reader = ExcelReader(excel_path)
            summary = reader.read_pending_emails(batch_size=self.config.processing.batch_size)
            return {
                'total_emails': summary.total_emails,
                'pending_count': summary.pending_count,
                'processed_count': summary.processed_count,
                'batch_count': len(summary.batches)
            }
        except Exception as e:
            logger.error(f"Excel summary error: {e}")
            return {'error': str(e)}

    def start_processing(self, excel_path: str) -> None:
        """Start email processing in background thread."""
        if self.is_processing:
            raise RuntimeError("Processing already active")

        self.should_stop = False
        self.is_processing = True

        def processing_thread():
            try:
                logger.info("Using Playwright automation engine")
                automation = BrowserAutomation(self.config)
                stats = automation.process_emails(excel_path, progress_callback=self._handle_progress)
                self.progress_queue.put(('complete', stats))
            except Exception as e:
                logger.error(f"Processing error: {e}")
                self.progress_queue.put(('error', str(e)))
            finally:
                self.is_processing = False

        self.current_thread = threading.Thread(target=processing_thread, daemon=True)
        self.current_thread.start()

    def _handle_progress(self, current: int, total: int):
        """Handle progress updates from background threads."""
        self.progress_queue.put(('progress', {'current': current, 'total': total}))

    def stop_processing(self):
        """Stop current processing."""
        self.should_stop = True
        self.is_processing = False

    def start_api_processing(self, excel_path: str) -> None:
        """Start API-based contact search in background thread."""
        if self.is_processing:
            raise RuntimeError("Processing already active")

        self.should_stop = False
        self.is_processing = True

        def api_thread():
            try:
                session_file = self.config.get_session_file_path()
                logger.info(f"Starting API search with session: {session_file}")
                result = process_emails_via_api(excel_path, session_file, progress_callback=self._handle_progress)
                self.progress_queue.put(('api_complete', result))
            except Exception as e:
                logger.error(f"API search error: {e}")
                self.progress_queue.put(('api_error', str(e)))
            finally:
                self.is_processing = False

        self.current_thread = threading.Thread(target=api_thread, daemon=True)
        self.current_thread.start()

    def start_gal_scraping(
        self,
        output_dir: str,
        max_contacts: int = 0,
        force_restart: bool = False,
    ) -> None:
        """Start GAL directory scraping in background thread."""
        if self.is_processing:
            raise RuntimeError("Processing already active")

        self.should_stop = False
        self.is_processing = True

        def gal_thread():
            try:
                session_file = self.config.get_session_file_path()
                logger.info(f"Starting GAL scraping with session: {session_file}")

                stop_flag = {'stop': False}
                self._gal_stop_flag = stop_flag

                def progress_callback(count, total):
                    self.progress_queue.put(('gal_progress', {'count': count, 'total': total}))

                result = scrape_gal(
                    session_file=session_file,
                    output_dir=output_dir,
                    max_contacts=max_contacts,
                    force_restart=force_restart,
                    progress_callback=progress_callback,
                    stop_flag=stop_flag,
                )
                self.progress_queue.put(('gal_complete', result))
            except Exception as e:
                logger.error(f"GAL scraping error: {e}")
                self.progress_queue.put(('gal_error', str(e)))
            finally:
                self.is_processing = False

        self.current_thread = threading.Thread(target=gal_thread, daemon=True)
        self.current_thread.start()

    def stop_gal_scraping(self):
        """Signal the GAL scraper to stop."""
        if hasattr(self, '_gal_stop_flag'):
            self._gal_stop_flag['stop'] = True

    def check_queue(self):
        """Check for progress updates."""
        try:
            while True:
                item = self.progress_queue.get_nowait()
                yield item
        except queue.Empty:
            pass


class VerificacionCorreosGUI:
    """Main GUI application class."""

    def __init__(self, root: tk.Tk):
        """Initialize GUI."""
        self.root = root
        self.root.title("Verificación de Correos OWA v2.0")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)

        # Load configuration
        try:
            self.config = Config()
        except Exception as e:
            messagebox.showerror("Configuration Error", f"Failed to load config: {e}")
            self.root.quit()
            return

        # Initialize service
        self.service = GUIService(self.config)

        # Setup logging for GUI - use DEBUG to see detailed asyncio/playwright logs
        setup_logging(level="DEBUG")
        self.log_messages = []
        
        self.scraper_output_dir = tk.StringVar(value=str(Path.cwd() / "data"))
        self.scraper_max_contacts = tk.IntVar(value=100)
        self.scraper_extracted_count = tk.IntVar(value=0)
        self.scraper_active = False
        self.scraper_log_messages = []

        # Create interface
        self._create_widgets()
        self._setup_status_check()

    def _create_widgets(self):
        """Create all GUI widgets."""
        # Create main container with padding
        main_container = ttk.Frame(self.root)
        main_container.pack(fill='both', expand=True, padx=10, pady=10)

        # Create notebook for tabs
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill='both', expand=True)

        # Create tabs
        self._create_processing_tab()
        self._create_session_tab()
        self._create_config_tab()
        self._create_scraper_tab()

        # Create status bar
        self._create_status_bar(main_container)

    def _create_processing_tab(self):
        """Create processing tab."""
        self.processing_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.processing_frame, text="📧 Procesamiento")

        # Main container
        main_frame = ttk.Frame(self.processing_frame)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="Archivo de Excel", padding=10)
        file_frame.pack(fill='x', pady=(0, 10))

        self.excel_path_var = tk.StringVar(value=str(self.config.get_excel_file_path()))
        ttk.Entry(file_frame, textvariable=self.excel_path_var, width=60).pack(side='left', padx=(0, 5))
        ttk.Button(file_frame, text="Seleccionar", command=self._select_excel_file).pack(side='left')
        ttk.Button(file_frame, text="Refrescar", command=self._refresh_excel_info).pack(side='left', padx=(5, 0))

        # Summary section
        summary_frame = ttk.LabelFrame(main_frame, text="Resumen", padding=10)
        summary_frame.pack(fill='x', pady=(0, 10))

        self.summary_text = tk.StringVar(value="Cargando información...")
        ttk.Label(summary_frame, textvariable=self.summary_text, wraplength=600).pack()

        # Automation engine indicator
        engine_info = self._get_automation_engine_info()
        self.engine_label = ttk.Label(
            summary_frame,
            text=engine_info,
            foreground='gray'
        )
        self.engine_label.pack(pady=(5, 0))

        # Control buttons
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill='x', pady=(0, 10))

        self.start_btn = ttk.Button(
            control_frame,
            text="🚀 Iniciar Procesamiento",
            command=self._start_processing,
            style='Accent.TButton'
        )
        self.start_btn.pack(side='left', padx=(0, 5))

        self.stop_btn = ttk.Button(
            control_frame,
            text="⏹ Detener",
            command=self._stop_processing,
            state='disabled'
        )
        self.stop_btn.pack(side='left', padx=5)

        ttk.Button(
            control_frame,
            text="📊 Ver Resultados",
            command=self._open_excel_file
        ).pack(side='left', padx=(5, 0))

        self.api_btn = ttk.Button(
            control_frame,
            text="🔍 Buscar por API",
            command=self._start_api_search
        )
        self.api_btn.pack(side='left', padx=(5, 0))

        # Progress section
        progress_frame = ttk.LabelFrame(main_frame, text="Progreso", padding=10)
        progress_frame.pack(fill='x', pady=(0, 10))

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.pack(fill='x', pady=(0, 5))

        self.progress_text = tk.StringVar(value="Listo para procesar")
        ttk.Label(progress_frame, textvariable=self.progress_text).pack()

        # Results section
        results_frame = ttk.LabelFrame(main_frame, text="Resultados", padding=10)
        results_frame.pack(fill='x', pady=(0, 10))

        columns = ('correo', 'status', 'nombre', 'email_personal', 'telefono')
        self.results_tree = ttk.Treeview(results_frame, columns=columns, show='headings', height=5)
        self.results_tree.heading('correo', text='Correo')
        self.results_tree.heading('status', text='Estado')
        self.results_tree.heading('nombre', text='Nombre')
        self.results_tree.heading('email_personal', text='Email Personal')
        self.results_tree.heading('telefono', text='Teléfono')
        self.results_tree.column('correo', width=200)
        self.results_tree.column('status', width=80)
        self.results_tree.column('nombre', width=150)
        self.results_tree.column('email_personal', width=180)
        self.results_tree.column('telefono', width=120)
        self.results_tree.pack(fill='x')

        results_scroll = ttk.Scrollbar(results_frame, orient='vertical', command=self.results_tree.yview)
        self.results_tree.configure(yscrollcommand=results_scroll.set)
        results_scroll.pack(side='right', fill='y')

        # Log section
        log_frame = ttk.LabelFrame(main_frame, text="Registro de Eventos", padding=10)
        log_frame.pack(fill='both', expand=True)

        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=15,
            wrap=tk.WORD,
            font=('Consolas', 9)
        )
        self.log_text.pack(fill='both', expand=True)

        # Log control buttons
        log_control_frame = ttk.Frame(log_frame)
        log_control_frame.pack(fill='x', pady=(5, 0))

        ttk.Button(log_control_frame, text="Limpiar", command=self._clear_log).pack(side='left')
        ttk.Button(log_control_frame, text="Guardar Log", command=self._save_log).pack(side='left', padx=(5, 0))

    def _create_session_tab(self):
        """Create session management tab."""
        self.session_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.session_frame, text="🔐 Sesión del Navegador")

        main_frame = ttk.Frame(self.session_frame)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Session status
        status_frame = ttk.LabelFrame(main_frame, text="Estado de la Sesión", padding=10)
        status_frame.pack(fill='x', pady=(0, 10))

        self.session_status_text = tk.StringVar(value="Verificando...")
        ttk.Label(status_frame, textvariable=self.session_status_text, wraplength=600).pack()

        # Session actions
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill='x', pady=(0, 10))

        ttk.Button(
            action_frame,
            text="🔄 Verificar Sesión",
            command=self._check_session_status
        ).pack(side='left', padx=(0, 5))

        ttk.Button(
            action_frame,
            text="🔧 Configurar Sesión",
            command=self._setup_session
        ).pack(side='left', padx=5)

        ttk.Button(
            action_frame,
            text="🗑️ Eliminar Sesión",
            command=self._delete_session
        ).pack(side='left', padx=(5, 0))

        # Session info
        info_frame = ttk.LabelFrame(main_frame, text="Información de la Sesión", padding=10)
        info_frame.pack(fill='both', expand=True)

        self.session_info_text = scrolledtext.ScrolledText(
            info_frame,
            height=15,
            wrap=tk.WORD,
            font=('Consolas', 9),
            state='disabled'
        )
        self.session_info_text.pack(fill='both', expand=True)

    def _create_config_tab(self):
        """Create configuration tab."""
        self.config_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.config_frame, text="⚙️ Configuración")

        main_frame = ttk.Frame(self.config_frame)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Configuration editor
        editor_frame = ttk.LabelFrame(main_frame, text="Editor de Configuración", padding=10)
        editor_frame.pack(fill='both', expand=True, pady=(0, 10))

        # Create scrollable configuration editor
        self._create_config_editor(editor_frame)

        # Quick actions
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill='x', pady=(0, 10))

        ttk.Button(
            action_frame,
            text="💾 Guardar Configuración",
            command=self._save_config
        ).pack(side='left', padx=(0, 5))

        ttk.Button(
            action_frame,
            text="🔄 Recargar Configuración",
            command=self._reload_config
        ).pack(side='left', padx=(5, 0))

        ttk.Button(
            action_frame,
            text="📁 Abrir Carpeta de Datos",
            command=self._open_data_folder
        ).pack(side='left', padx=(5, 0))

        ttk.Button(
            action_frame,
            text="🔧 Asistente de Configuración",
            command=self._run_config_wizard
        ).pack(side='left', padx=(5, 0))

    def _create_scraper_tab(self):
        """Create scraper tab for extracting Outlook contacts."""
        self.scraper_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.scraper_frame, text="🔍 Scraper de Contactos")

        main_frame = ttk.Frame(self.scraper_frame)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Configuration section
        config_frame = ttk.LabelFrame(main_frame, text="Configuración de Scraping", padding=10)
        config_frame.pack(fill='x', pady=(0, 10))

        # Output directory
        dir_label_frame = ttk.Frame(config_frame)
        dir_label_frame.pack(fill='x', pady=(0, 5))
        ttk.Label(dir_label_frame, text="Directorio de salida:").pack(anchor='w')

        dir_input_frame = ttk.Frame(config_frame)
        dir_input_frame.pack(fill='x', pady=(0, 10))

        ttk.Entry(
            dir_input_frame,
            textvariable=self.scraper_output_dir,
            state='readonly',
            width=60
        ).pack(side='left', fill='x', expand=True, padx=(0, 5))

        ttk.Button(
            dir_input_frame,
            text="📁 Seleccionar",
            command=self._select_scraper_directory
        ).pack(side='left')

        # Contact quantity
        quantity_label_frame = ttk.Frame(config_frame)
        quantity_label_frame.pack(fill='x', pady=(0, 5))
        ttk.Label(quantity_label_frame, text="Cantidad de contactos a extraer:").pack(anchor='w')

        quantity_frame = ttk.Frame(config_frame)
        quantity_frame.pack(fill='x')

        ttk.Spinbox(
            quantity_frame,
            from_=1,
            to=10000,
            textvariable=self.scraper_max_contacts,
            width=20
        ).pack(anchor='w')

        # Control section
        control_frame = ttk.LabelFrame(main_frame, text="Control de Extracción", padding=10)
        control_frame.pack(fill='x', pady=(0, 10))

        button_frame = ttk.Frame(control_frame)
        button_frame.pack(fill='x')

        self.scraper_start_btn = ttk.Button(
            button_frame,
            text="▶️ Iniciar Extracción",
            command=self._start_scraper,
            style='Accent.TButton'
        )
        self.scraper_start_btn.pack(side='left', padx=(0, 5), fill='x', expand=True)

        self.scraper_stop_btn = ttk.Button(
            button_frame,
            text="⏹️ Detener",
            command=self._stop_scraper,
            state='disabled'
        )
        self.scraper_stop_btn.pack(side='left', fill='x', expand=True)

        # Progress section
        progress_frame = ttk.LabelFrame(main_frame, text="Progreso de Extracción", padding=10)
        progress_frame.pack(fill='x', pady=(0, 10))

        # Counter
        counter_frame = ttk.Frame(progress_frame)
        counter_frame.pack(fill='x', pady=(0, 5))

        ttk.Label(counter_frame, text="Contactos extraídos:").pack(side='left')

        self.scraper_counter_label = ttk.Label(
            counter_frame,
            textvariable=self.scraper_extracted_count,
            font=('TkDefaultFont', 14, 'bold'),
            foreground='#27ae60'
        )
        self.scraper_counter_label.pack(side='left', padx=5)

        self.scraper_total_label = ttk.Label(
            counter_frame,
            text="/ 0",
            font=('TkDefaultFont', 12),
            foreground='gray'
        )
        self.scraper_total_label.pack(side='left')

        # Progress bar
        self.scraper_progress_bar = ttk.Progressbar(
            progress_frame,
            mode='determinate',
            length=300
        )
        self.scraper_progress_bar.pack(fill='x', pady=(0, 5))

        # Status
        self.scraper_status_label = ttk.Label(
            progress_frame,
            text="⚪ Esperando...",
            font=('TkDefaultFont', 10, 'italic'),
            foreground='gray'
        )
        self.scraper_status_label.pack(anchor='w')

        # Log section
        log_frame = ttk.LabelFrame(main_frame, text="Log de Scraping", padding=10)
        log_frame.pack(fill='both', expand=True)

        self.scraper_log_text = scrolledtext.ScrolledText(
            log_frame,
            height=12,
            wrap=tk.WORD,
            font=('Consolas', 9)
        )
        self.scraper_log_text.pack(fill='both', expand=True, pady=(0, 5))

        # Log control buttons
        log_control_frame = ttk.Frame(log_frame)
        log_control_frame.pack(fill='x')

        ttk.Button(
            log_control_frame,
            text="Limpiar",
            command=self._clear_scraper_log
        ).pack(side='left', padx=(0, 5))

        ttk.Button(
            log_control_frame,
            text="Guardar Log",
            command=self._save_scraper_log
        ).pack(side='left')

        # Initial log message
        self._add_scraper_log("✅ Interfaz de scraper iniciada correctamente")
        self._add_scraper_log(f"📁 Directorio de salida: {self.scraper_output_dir.get()}")
        self._update_scraper_total_label()

    def _select_scraper_directory(self):
        """Select output directory for scraper."""
        directory = filedialog.askdirectory(
            title="Seleccionar directorio de salida",
            initialdir=self.scraper_output_dir.get()
        )
        if directory:
            self.scraper_output_dir.set(directory)
            self._add_scraper_log(f"📁 Directorio cambi ado a: {directory}")

    def _update_scraper_total_label(self):
        """Update the total label with current max contacts."""
        total = self.scraper_max_contacts.get()
        self.scraper_total_label.config(text=f"/ {total}")

    def _add_scraper_log(self, message):
        """Add message to scraper log."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        self.scraper_log_messages.append(log_entry)

        self.scraper_log_text.insert(tk.END, log_entry)
        self.scraper_log_text.see(tk.END)

    def _clear_scraper_log(self):
        """Clear scraper log messages."""
        self.scraper_log_text.delete('1.0', tk.END)
        self.scraper_log_messages.clear()

    def _save_scraper_log(self):
        """Save scraper log messages to file."""
        file_path = filedialog.asksaveasfilename(
            title="Guardar Log",
            defaultextension=".log",
            filetypes=[("Log files", "*.log"), ("Text files", "*.txt"), ("All files", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.writelines(self.scraper_log_messages)
                messagebox.showinfo("Log Guardado", f"Log guardado en:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Error al guardar log: {e}")

    def _update_scraper_progress(self, count):
        """Update scraper progress counter and bar."""
        self.scraper_extracted_count.set(count)
        total = self.scraper_max_contacts.get()

        # Update progress bar
        if total > 0:
            progress = (count / total) * 100
            self.scraper_progress_bar['value'] = progress

    def _update_scraper_status(self, message, color='gray'):
        """Update scraper status label."""
        self.scraper_status_label.config(text=message, foreground=color)

    def _start_scraper(self):
        """Start the GAL scraper via OWA API."""
        if self.scraper_active:
            self._add_scraper_log("⚠️ El scraper ya está activo")
            return

        # Validate quantity
        if self.scraper_max_contacts.get() <= 0:
            messagebox.showwarning("Cantidad Inválida", "La cantidad debe ser mayor a 0")
            return

        session_file = self.config.get_session_file_path()
        if not Path(session_file).exists():
            messagebox.showwarning(
                "Sesión Requerida",
                "No hay sesión guardada.\nUse el comando 'verificacion-correo setup' o el botón 'Configurar Sesión' en la pestaña de Sesión."
            )
            return

        # Confirm if resuming
        output_dir = Path(self.scraper_output_dir.get()) / "gal"
        progress_file = output_dir / "gal_progress.json"
        if progress_file.exists():
            with open(progress_file) as f:
                prev = json.load(f)
            resume = messagebox.askyesno(
                "Reanudar Extracción",
                f"Se encontró un progreso anterior:\n"
                f"  Offset: {prev.get('offset', 0)}\n"
                f"  Contactos: {prev.get('count', 0)}\n\n"
                f"¿Desea reanudar desde donde se quedó?\n\n"
                f"Seleccione 'No' para empezar desde cero."
            )
            force_restart = not resume
        else:
            force_restart = True

        # Reset variables
        self.scraper_extracted_count.set(0)
        self._update_scraper_progress(0)
        self._update_scraper_total_label()

        # Update UI
        self.scraper_start_btn.config(state='disabled')
        self.scraper_stop_btn.config(state='normal')
        self.scraper_active = True

        max_contacts = self.scraper_max_contacts.get()
        self._add_scraper_log(f"🚀 Iniciando scraper (máx {max_contacts} contactos)...")
        self._update_scraper_status("🔄 Ejecutando...", "#3498db")

        try:
            self.service.start_gal_scraping(
                output_dir=str(output_dir),
                max_contacts=max_contacts,
                force_restart=force_restart,
            )
        except Exception as e:
            self._add_scraper_log(f"❌ Error al iniciar: {e}")
            self.scraper_start_btn.config(state='normal')
            self.scraper_stop_btn.config(state='disabled')
            self.scraper_active = False

    def _stop_scraper(self):
        """Stop the scraper gracefully."""
        if not self.scraper_active:
            return

        self._add_scraper_log("⚠️ Deteniendo scraper...")
        self._update_scraper_status("⏸️ Deteniendo...", "#e67e22")
        self.service.stop_gal_scraping()
        self.scraper_stop_btn.config(state='disabled')

    # _run_scraper_thread and _run_scraper_async removed — replaced by service.start_gal_scraping()

    def _create_config_editor(self, parent):
        """Create configuration editor interface."""
        # Create notebook for different config sections
        config_notebook = ttk.Notebook(parent)
        config_notebook.pack(fill='both', expand=True)

        # Basic settings
        basic_frame = ttk.Frame(config_notebook)
        config_notebook.add(basic_frame, text="⚡ Básico")

        # URL
        ttk.Label(basic_frame, text="URL de OWA:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.owa_url_var = tk.StringVar(value=self.config.page_url)
        ttk.Entry(basic_frame, textvariable=self.owa_url_var, width=60).grid(row=0, column=1, padx=5, pady=5, sticky='ew')

        # Batch size
        ttk.Label(basic_frame, text="Tamaño de lote:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.batch_size_var = tk.IntVar(value=self.config.processing.batch_size)
        ttk.Entry(basic_frame, textvariable=self.batch_size_var, width=10).grid(row=1, column=1, padx=5, pady=5, sticky='w')

        # Excel file
        ttk.Label(basic_frame, text="Archivo Excel:").grid(row=2, column=0, sticky='w', padx=5, pady=5)
        excel_frame = ttk.Frame(basic_frame)
        excel_frame.grid(row=2, column=1, padx=5, pady=5, sticky='ew')

        self.excel_file_var = tk.StringVar(value=self.config.get_excel_file_path())
        ttk.Entry(excel_frame, textvariable=self.excel_file_var, width=50).pack(side='left', padx=(0, 5))
        ttk.Button(excel_frame, text="Seleccionar", command=self._select_excel_config).pack(side='left')

        # Browser settings
        browser_frame = ttk.Frame(config_notebook)
        config_notebook.add(browser_frame, text="🌐 Navegador")

        # Headless mode
        self.headless_var = tk.BooleanVar(value=self.config.browser.headless)
        ttk.Checkbutton(browser_frame, text="Modo sin ventana (headless)", variable=self.headless_var).pack(anchor='w', padx=5, pady=5)

        # Discard draft option
        self.discard_draft_var = tk.BooleanVar(value=self.config.processing.discard_draft)
        ttk.Checkbutton(browser_frame, text="Descartar borrador automáticamente", variable=self.discard_draft_var).pack(anchor='w', padx=5, pady=5)

        # Session file
        ttk.Label(browser_frame, text="Archivo de sesión:").pack(anchor='w', padx=5, pady=(10, 0))
        session_frame = ttk.Frame(browser_frame)
        session_frame.pack(fill='x', padx=5, pady=5)

        self.session_file_var = tk.StringVar(value=self.config.get_session_file_path())
        ttk.Entry(session_frame, textvariable=self.session_file_var, width=50).pack(side='left', padx=(0, 5))
        ttk.Button(session_frame, text="Seleccionar", command=self._select_session_file).pack(side='left')

        # Default emails
        emails_frame = ttk.Frame(config_notebook)
        config_notebook.add(emails_frame, text="📧 Correos por Defecto")

        ttk.Label(emails_frame, text="Correos electrónicos de respaldo:").pack(anchor='w', padx=5, pady=5)

        # Text widget for emails
        self.emails_text = scrolledtext.ScrolledText(emails_frame, height=10, width=60)
        self.emails_text.pack(fill='both', expand=True, padx=5, pady=5)

        # Load current default emails
        default_emails_text = '\n'.join(self.config.default_emails)
        self.emails_text.insert('1.0', default_emails_text)

        # Configure grid weights
        basic_frame.columnconfigure(1, weight=1)

    def _select_excel_config(self):
        """Select Excel file for configuration."""
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo de Excel",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_file_var.set(file_path)

    def _select_session_file(self):
        """Select session file for configuration."""
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo de sesión",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if file_path:
            self.session_file_var.set(file_path)

    def _save_config(self):
        """Save configuration changes."""
        try:
            # Update configuration object
            self.config.page_url = self.owa_url_var.get()
            self.config.processing.batch_size = self.batch_size_var.get()
            self.config.processing.discard_draft = self.discard_draft_var.get()
            self.config.excel.default_file = self.excel_file_var.get()
            self.config.browser.headless = self.headless_var.get()
            self.config.browser.session_file = self.session_file_var.get()

            # Update default emails
            emails_text = self.emails_text.get('1.0', tk.END).strip()
            self.config.default_emails = [email.strip() for email in emails_text.split('\n') if email.strip()]

            # Save to file
            self._save_config_to_file()

            messagebox.showinfo("Configuración", "Configuración guardada exitosamente")
            self._refresh_excel_info()

        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar configuración: {e}")

    @staticmethod
    def _build_config_yaml(config: Config) -> dict:
        """Build YAML-serializable dict from Config object."""
        return {
            'page_url': config.page_url,
            'default_emails': config.default_emails,
            'browser': {
                'headless': config.browser.headless,
                'session_file': config.browser.session_file
            },
            'excel': {
                'default_file': config.excel.default_file,
                'start_row': config.excel.start_row,
                'email_column': config.excel.email_column
            },
            'processing': {
                'batch_size': config.processing.batch_size,
                'discard_draft': config.processing.discard_draft
            },
            'selectors': {
                'new_message_btn': config.selectors.new_message_btn,
                'to_field_role': config.selectors.to_field_role,
                'to_field_name': config.selectors.to_field_name,
                'popup': config.selectors.popup,
                'discard_btn': config.selectors.discard_btn
            },
            'wait_times': {
                'after_new_message': config.wait_times.after_new_message,
                'after_fill_to': config.wait_times.after_fill_to,
                'after_blur': config.wait_times.after_blur,
                'popup_visible': config.wait_times.popup_visible,
                'after_click_token': config.wait_times.after_click_token,
                'popup_load_data': config.wait_times.popup_load_data,
                'after_close_popup': config.wait_times.after_close_popup,
                'before_discard': config.wait_times.before_discard
            }
        }

    def _save_config_to_file(self):
        """Save configuration to YAML file."""
        config_data = self._build_config_yaml(self.config)
        with open(self.config._config_path, 'w', encoding='utf-8') as f:
            yaml.dump(config_data, f, default_flow_style=False, allow_unicode=True)

    def _run_config_wizard(self):
        """Run configuration wizard for first-time setup."""
        wizard = ConfigWizard(self.root, self.config)
        if wizard.result:
            self._reload_config()

    def _create_status_bar(self, parent):
        """Create status bar."""
        status_frame = ttk.Frame(parent)
        status_frame.pack(fill='x', pady=(10, 0))

        # Status label
        self.status_label = ttk.Label(status_frame, text="Listo")
        self.status_label.pack(side='left')

        # Clock label
        self.clock_label = ttk.Label(status_frame)
        self.clock_label.pack(side='right')
        self._update_clock()

    def _setup_status_check(self):
        """Setup periodic status checks."""
        self._check_session_status()
        self._refresh_excel_info()
        self._check_progress()

    def _update_clock(self):
        """Update clock in status bar."""
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.clock_label.config(text=now)
        self.root.after(1000, self._update_clock)

    def _check_progress(self):
        """Check for progress updates."""
        for item_type, data in self.service.check_queue():
            if item_type == 'complete':
                self._processing_complete(data)
            elif item_type == 'error':
                self._processing_error(data)
            elif item_type == 'api_complete':
                self._handle_api_complete(data)
            elif item_type == 'api_error':
                self._handle_api_error(data)
            elif item_type == 'gal_progress':
                self._update_gal_progress(data)
            elif item_type == 'gal_complete':
                self._handle_gal_complete(data)
            elif item_type == 'gal_error':
                self._handle_gal_error(data)
            elif item_type == 'progress':
                self._update_progress(data)

        if not self.service.is_processing:
            self.progress_bar.stop()
            if self.progress_bar['mode'] != 'determinate':
                self.progress_bar.config(mode='determinate')

        # Schedule next check
        self.root.after(100, self._check_progress)

    def _get_automation_engine_info(self) -> str:
        """Get information about the automation engine being used."""
        return "🤖 Motor: Playwright"

    def _select_excel_file(self):
        """Select Excel file dialog."""
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo de Excel",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_path_var.set(file_path)
            self._refresh_excel_info()

    def _refresh_excel_info(self):
        """Refresh Excel file information."""
        excel_path = self.excel_path_var.get()
        if not excel_path:
            return

        try:
            summary = self.service.get_excel_summary(excel_path)
            if 'error' in summary:
                self.summary_text.set(f"Error: {summary['error']}")
                self.start_btn.config(state='disabled')
            else:
                text = (f"Total de correos: {summary['total_emails']}\n"
                       f"Procesados: {summary['processed_count']}\n"
                       f"Pendientes: {summary['pending_count']}\n"
                       f"Lotes: {summary['batch_count']}")
                self.summary_text.set(text)

                # Enable/disable start button based on pending emails
                self.start_btn.config(
                    state='normal' if summary['pending_count'] > 0 else 'disabled'
                )

        except Exception as e:
            self.summary_text.set(f"Error: {e}")
            self.start_btn.config(state='disabled')

    def _start_processing(self):
        """Start email processing."""
        excel_path = self.excel_path_var.get()
        if not excel_path:
            messagebox.showwarning("Archivo Requerido", "Por favor seleccione un archivo de Excel")
            return

        # Validate session first
        session_status = self.service.validate_session()
        if not session_status.get('is_valid'):
            if not messagebox.askyesno(
                "Sesión Inválida",
                "La sesión del navegador no es válida o ha expirado.\n¿Desea configurar una nueva sesión?"
            ):
                return
            self._setup_session()
            return

        # Confirm processing
        if not messagebox.askyesno(
            "Confirmar Procesamiento",
            "¿Desea iniciar el procesamiento de correos pendientes?"
        ):
            return

        try:
            self.service.start_processing(excel_path)
            self.start_btn.config(state='disabled')
            self.stop_btn.config(state='normal')
            self.progress_text.set("Iniciando procesamiento...")
            self._add_log("🚀 Iniciando procesamiento de correos")
            self.status_label.config(text="Procesando...")

        except Exception as e:
            messagebox.showerror("Error", f"Error al iniciar procesamiento: {e}")

    def _stop_processing(self):
        """Stop email processing."""
        try:
            self.service.stop_processing()
            self.start_btn.config(state='normal')
            self.stop_btn.config(state='disabled')
            self.progress_text.set("Procesamiento detenido")
            self._add_log("⏹ Procesamiento detenido por el usuario")
            self.status_label.config(text="Detenido")

        except Exception as e:
            messagebox.showerror("Error", f"Error al detener procesamiento: {e}")

    def _processing_complete(self, stats):
        """Handle processing completion."""
        self.root.after(0, self._handle_processing_complete, stats)

    def _handle_processing_complete(self, stats):
        """Handle processing completion in GUI thread."""
        self.start_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        self.progress_text.set("Procesamiento completado")
        self.status_label.config(text="Completado")

        # Show results
        total = stats.total_emails
        success = stats.successful
        not_found = stats.not_found
        errors = stats.errors

        message = f"""Procesamiento completado:

📧 Total procesados: {total}
✅ Exitosos: {success} ({success/total*100:.1f}%)
❌ No encontrados: {not_found} ({not_found/total*100:.1f}%)
⚠️ Errores: {errors} ({errors/total*100:.1f}%)

Duración: {stats.duration_seconds:.1f} segundos
Resultados guardados en: {self.excel_path_var.get()}"""

        messagebox.showinfo("Procesamiento Completado", message)
        self._add_log(f"✅ Procesamiento completado: {success} exitosos, {not_found} no encontrados, {errors} errores")
        self._refresh_excel_info()
        self._refresh_results_tree()

    def _processing_error(self, error_msg):
        """Handle processing error."""
        self.root.after(0, self._handle_processing_error, error_msg)

    def _handle_processing_error(self, error_msg):
        """Handle processing error in GUI thread."""
        self.start_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        self.progress_text.set("Error en procesamiento")
        self.status_label.config(text="Error")

        messagebox.showerror("Error de Procesamiento", f"Ocurrió un error:\n{error_msg}")
        self._add_log(f"❌ Error de procesamiento: {error_msg}")

    def _start_api_search(self):
        """Start API-based contact search."""
        excel_path = self.excel_path_var.get()
        if not excel_path:
            messagebox.showwarning("Archivo Requerido", "Por favor seleccione un archivo de Excel")
            return

        session_file = self.config.get_session_file_path()
        if not Path(session_file).exists():
            if not messagebox.askyesno(
                "Sesión Requerida",
                "El archivo de sesión no existe.\n¿Desea configurar una nueva sesión primero?"
            ):
                return
            self._setup_session()
            return

        # Confirm
        if not messagebox.askyesno(
            "Confirmar Búsqueda por API",
            "¿Desea buscar contactos en el directorio de OWA vía API?\n\n"
            "Esto consultará directamente el servicio FindPeople de Exchange\n"
            "usando las cookies de la sesión guardada."
        ):
            return

        try:
            self.service.start_api_processing(excel_path)
            self.start_btn.config(state='disabled')
            self.api_btn.config(state='disabled')
            self.stop_btn.config(state='normal')
            self.progress_text.set("Buscando contactos por API...")
            self._add_log("🔍 Iniciando búsqueda de contactos vía API de OWA")
            self.status_label.config(text="Buscando...")

        except Exception as e:
            messagebox.showerror("Error", f"Error al iniciar búsqueda API: {e}")

    def _handle_api_complete(self, result):
        """Handle API search completion in GUI thread."""
        self.start_btn.config(state='normal')
        self.api_btn.config(state='normal')
        self.stop_btn.config(state='disabled')

        total = result.get("total", 0)
        success = result.get("success", 0)
        not_found = result.get("not_found", 0)
        errors = result.get("errors", 0)
        duration = result.get("duration", 0)
        session_expired = result.get("expired", False)
        found_pct = (success / total * 100) if total > 0 else 0

        if session_expired:
            remaining = result.get("remaining", 0)
            self.progress_text.set(f"⚠️ Sesión expirada tras {total} consultas")
            self.status_label.config(text="Sesión expirada")
            message = f"""Búsqueda por API DETENIDA - Sesión expirada

⚠️ La sesión de OWA ha expirado después de {total} consultas.
   Los emails restantes quedarán pendientes para el próximo ciclo.

📧 Procesados: {total}
✅ Encontrados: {success} ({found_pct:.1f}%)
❌ No encontrados: {not_found}
⚠️ Errores (sesión): {errors}

            📌 Para continuar:
                1. Cierra sesión y vuelve a iniciar en OWA
                2. Usa 'Configurar Sesión' en la pestaña de Sesión
                3. Vuelve a ejecutar la búsqueda

Duración: {duration:.1f} segundos"""

            messagebox.showwarning("Sesión Expirada", message)
            self._add_log(f"⚠️ API search: sesión expirada tras {total} consultas ({success} encontrados, {not_found} no encontrados)")
        else:
            self.progress_text.set("Búsqueda API completada")
            self.status_label.config(text="Completado")
            message = f"""Búsqueda por API completada:

📧 Total procesados: {total}
✅ Encontrados: {success} ({found_pct:.1f}%)
❌ No encontrados: {not_found}
⚠️ Errores: {errors}

Duración: {duration:.1f} segundos
Resultados guardados en: {self.excel_path_var.get()}"""

            messagebox.showinfo("Búsqueda por API Completada", message)
            self._add_log(f"✅ API search: {success} encontrados, {not_found} no encontrados, {errors} errores")

        self._refresh_excel_info()
        self._refresh_results_tree()

    def _handle_api_error(self, error_msg):
        """Handle API search error in GUI thread."""
        self.start_btn.config(state='normal')
        self.api_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        self.progress_text.set("Error en búsqueda API")
        self.status_label.config(text="Error")

        messagebox.showerror("Error de Búsqueda API", f"Ocurrió un error:\n{error_msg}")
        self._add_log(f"❌ Error en búsqueda API: {error_msg}")

    def _update_gal_progress(self, data):
        """Update GAL scraper progress from background thread."""
        count = data.get("count", 0)
        total = data.get("total", 0) or 1
        self.scraper_extracted_count.set(count)
        self._update_scraper_progress(count)
        if total > 0:
            pct = (count / total) * 100
            self.scraper_progress_bar["value"] = pct

    def _handle_gal_complete(self, result):
        """Handle GAL scraper completion."""
        self.scraper_start_btn.config(state="normal")
        self.scraper_stop_btn.config(state="disabled")
        self.scraper_active = False

        total = result.get("total", 0)
        duration = result.get("duration", 0)
        expired = result.get("expired", False)
        stopped = result.get("stopped", False)
        files = result.get("files", {})

        self.scraper_extracted_count.set(total)
        self._update_scraper_progress(total)

        if expired:
            self._update_scraper_status("⚠️ Sesión expirada", "#e67e22")
            self._add_scraper_log(f"⚠️ Sesión expirada tras {total} contactos")
            messagebox.showwarning(
                "Sesión Expirada",
                f"La sesión expiró después de {total} contactos.\n\n"
                f"El progreso se ha guardado. Para continuar:\n"
                f"1. Vuelva a iniciar sesión en OWA\n"
                f"2. Use 'Configurar Sesión' en la pestaña de Sesión\n"
                f"3. Inicie la extracción nuevamente (se reanudará)\n\n"
                f"📁 {files.get('json', '')}"
            )
        elif stopped:
            self._update_scraper_status("⏹️ Detenido", "#e67e22")
            self._add_scraper_log(f"⏹️ Scraper detenido ({total} contactos)")
            messagebox.showinfo(
                "Extracción Detenida",
                f"Scraper detenido con {total} contactos.\n\n"
                f"Resultados guardados en:\n"
                f"  {files.get('json', '')}\n"
                f"  {files.get('csv', '')}"
            )
        else:
            self._update_scraper_status("✅ Completado", "#27ae60")
            self._add_scraper_log(f"✅ Extracción completada: {total} contactos en {duration:.1f}s")
            messagebox.showinfo(
                "Extracción Completada",
                f"Directorio completo extraído: {total} contactos\n\n"
                f"Duración: {duration:.1f}s\n\n"
                f"Resultados guardados en:\n"
                f"  {files.get('json', '')}\n"
                f"  {files.get('csv', '')}"
            )

    def _handle_gal_error(self, error_msg):
        """Handle GAL scraper error."""
        self.scraper_start_btn.config(state="normal")
        self.scraper_stop_btn.config(state="disabled")
        self.scraper_active = False
        self._update_scraper_status("❌ Error", "#e74c3c")
        self._add_scraper_log(f"❌ Error: {error_msg}")
        messagebox.showerror("Error de Extracción", f"Ocurrió un error:\n{error_msg}")

    MAX_LOG_MESSAGES = 1000

    def _add_log(self, message):
        """Add message to log."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        self.log_messages.append(log_entry)
        if len(self.log_messages) > self.MAX_LOG_MESSAGES:
            self.log_messages.pop(0)
            self.log_text.delete('1.0', '2.0')

        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)

    def _clear_log(self):
        """Clear log messages."""
        self.log_text.delete('1.0', tk.END)
        self.log_messages.clear()

    def _save_log(self):
        """Save log messages to file."""
        file_path = filedialog.asksaveasfilename(
            title="Guardar Log",
            defaultextension=".log",
            filetypes=[("Log files", "*.log"), ("Text files", "*.txt"), ("All files", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.writelines(self.log_messages)
                messagebox.showinfo("Log Guardado", f"Log guardado en:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Error al guardar log: {e}")

    def _check_session_status(self):
        """Check and display session status."""
        try:
            from pathlib import Path
            import json

            # Check Playwright session
            session_file = Path(self.config.get_session_file_path())
            session_type = "Playwright (Chromium)"

            status_text = f"Tipo: {session_type}\n"
            status_text += f"Archivo: {session_file}\n"
            status_text += f"Existe: {'Sí' if session_file.exists() else 'No'}\n"

            if session_file.exists():
                try:
                    with open(session_file, 'r') as f:
                        session_data = json.load(f)

                    cookies_count = len(session_data.get('cookies', []))
                    status_text += f"Válida: Sí\n"
                    status_text += f"Cookies: {cookies_count}\n"

                    if 'origins' in session_data:
                        status_text += f"Orígenes: {len(session_data.get('origins', []))}"
                except (json.JSONDecodeError, KeyError):
                    status_text += f"Válida: Error al leer"
            else:
                status_text += f"Válida: No\n"
                status_text += f"\n⚠️ Sesión no encontrada\n"
                status_text += f"Usa 'Configurar Sesión' para crear una"

            self.session_status_text.set(status_text)

            # Update session info text
            self.session_info_text.config(state='normal')
            self.session_info_text.delete('1.0', tk.END)
            self.session_info_text.insert('1.0', f"Información de la sesión:\n\n{status_text}")
            self.session_info_text.config(state='disabled')

        except Exception as e:
            self.session_status_text.set(f"Error: {e}")

    def _setup_session(self):
        """Set up browser session."""
        browser_type = "Chromium con Playwright"
        session_file = "state.json"

        if not messagebox.askyesno(
            "Configurar Sesión",
            f"Se abrirá {browser_type} para que inicie sesión manualmente.\n"
            f"La sesión se guardará en: {session_file}\n\n"
            "Después de iniciar sesión, vuelve a esta ventana.\n\n"
            "¿Desea continuar?"
        ):
            return

        try:
            # Use Playwright session setup
            self._setup_playwright_session()

        except Exception as e:
            messagebox.showerror("Error", f"Error al iniciar configuración de sesión: {e}")

    def _setup_playwright_session(self):
        """Set up Playwright (Chromium) session."""
        def setup_in_background():
            try:
                success = self.service.setup_session()
                if success:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Éxito",
                        "Sesión de Playwright configurada correctamente.\n\n"
                        "La sesión ha sido guardada en state.json"
                    ))
                else:
                    self.root.after(0, lambda: messagebox.showerror(
                        "Error",
                        "No se pudo configurar la sesión.\n\n"
                        "Asegúrate de iniciar sesión correctamente antes de guardar."
                    ))

                self.root.after(0, self._check_session_status)

            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror(
                    "Error",
                    f"Error durante la configuración de sesión: {e}"
                ))

        setup_thread = threading.Thread(target=setup_in_background, daemon=True)
        setup_thread.start()

        messagebox.showinfo(
            "Configuración en Progreso",
            "Se está abriendo Chromium...\n\n"
            "1. Inicia sesión con tus credenciales\n"
            "2. Navega a tu bandeja de entrada\n"
            "3. Vuelve a esta ventana y presiona ENTER en la terminal\n"
            "4. La sesión se guardará automáticamente"
        )

    def _delete_session(self):
        """Delete browser session."""
        if not messagebox.askyesno(
            "Eliminar Sesión",
            "¿Está seguro de que desea eliminar la sesión guardada?\n"
            "Necesitará configurar una nueva sesión para continuar."
        ):
            return

        try:
            session_file = Path(self.config.get_session_file_path())
            if session_file.exists():
                session_file.unlink()
                self._add_log("🗑️ Sesión eliminada correctamente")
                messagebox.showinfo(
                    "Sesión Eliminada",
                    f"La sesión ha sido eliminada:\n{session_file}"
                )
            else:
                messagebox.showinfo(
                    "Sin Sesión",
                    "No hay archivo de sesión para eliminar."
                )
            self._check_session_status()

        except Exception as e:
            messagebox.showerror("Error", f"Error al eliminar sesión: {e}")

    def _open_excel_file(self):
        """Open Excel file with system default application."""
        excel_path = self.excel_path_var.get()
        if not excel_path or not Path(excel_path).exists():
            messagebox.showwarning("Archivo no encontrado", "El archivo de Excel no existe")
            return

        try:
            import subprocess
            import platform
            if platform.system() == 'Windows':
                subprocess.startfile(excel_path)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.run(['open', excel_path])
            else:  # Linux
                subprocess.run(['xdg-open', excel_path])
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo: {e}")

    def _open_data_folder(self):
        """Open data folder in system file explorer."""
        data_path = Path(self.config.get_excel_file_path()).parent
        try:
            import subprocess
            import platform
            if platform.system() == 'Windows':
                subprocess.startfile(str(data_path))
            elif platform.system() == 'Darwin':  # macOS
                subprocess.run(['open', str(data_path)])
            else:  # Linux
                subprocess.run(['xdg-open', str(data_path)])
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la carpeta: {e}")

    def _refresh_results_tree(self):
        """Refresh results treeview from Excel file."""
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        try:
            excel_path = self.excel_path_var.get()
            if not excel_path or not Path(excel_path).exists():
                return
            from verificacion_correo.core.excel import ExcelReader, ProcessingStatus
            reader = ExcelReader(excel_path)
            records = reader.read_all_records()
            for rec in records[-50:]:
                status_char = '✅' if rec.status == ProcessingStatus.SUCCESS else ('❌' if rec.status == ProcessingStatus.ERROR else '⬜')
                status_label = rec.status.value if hasattr(rec.status, 'value') else str(rec.status)
                nombre = (rec.data or {}).get('name', '')
                email_personal = (rec.data or {}).get('email', '')
                telefono = (rec.data or {}).get('phone', '')
                self.results_tree.insert('', 'end', values=(
                    rec.email, f"{status_char} {status_label}",
                    nombre[:40], email_personal[:40], telefono[:30]
                ))
        except Exception as e:
            self._add_log(f"⚠️ No se pudieron cargar resultados: {e}")

    def _reload_config(self):
        """Reload configuration."""
        try:
            from ..core.config import reload_config
            reload_config()
            self.config = Config()
            messagebox.showinfo("Configuración", "Configuración recargada exitosamente")
            self._refresh_excel_info()
            self._check_session_status()
        except Exception as e:
            messagebox.showerror("Error", f"Error al recargar configuración: {e}")

    def _update_progress(self, progress_data):
        """Update progress bar with current/total info."""
        current = progress_data.get('current', 0)
        total = progress_data.get('total', 0) or 1
        pct = (current / total) * 100
        self.progress_var.set(pct)
        self.progress_bar.config(mode='determinate')
        self.progress_bar['value'] = pct
        self.progress_text.set(f"Procesando: {current}/{total} ({pct:.0f}%)")


class ConfigWizard:
    """Configuration wizard for first-time setup."""

    def __init__(self, parent, config: Config):
        """Initialize configuration wizard."""
        self.config = config
        self.result = False
        self.current_step = 0
        self.total_steps = 4

        # Create wizard window
        self.wizard = tk.Toplevel(parent)
        self.wizard.title("Asistente de Configuración Inicial")
        self.wizard.geometry("600x500")
        self.wizard.resizable(False, False)
        self.wizard.transient(parent)
        self.wizard.grab_set()

        # Center the wizard
        self.wizard.update_idletasks()
        x = (self.wizard.winfo_screenwidth() // 2) - (600 // 2)
        y = (self.wizard.winfo_screenheight() // 2) - (500 // 2)
        self.wizard.geometry(f"600x500+{x}+{y}")

        # Create wizard interface
        self._create_wizard_interface()
        self._show_step(0)

    def _create_wizard_interface(self):
        """Create wizard interface."""
        # Title
        title_frame = ttk.Frame(self.wizard)
        title_frame.pack(fill='x', padx=20, pady=(20, 10))

        self.title_label = ttk.Label(title_frame, text="", font=('TkDefaultFont', 12, 'bold'))
        self.title_label.pack()

        # Progress bar
        progress_frame = ttk.Frame(self.wizard)
        progress_frame.pack(fill='x', padx=20, pady=(0, 20))

        self.progress_var = tk.IntVar(value=0)
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=self.total_steps)
        self.progress_bar.pack(fill='x')

        # Content area
        self.content_frame = ttk.Frame(self.wizard)
        self.content_frame.pack(fill='both', expand=True, padx=20, pady=10)

        # Navigation buttons
        nav_frame = ttk.Frame(self.wizard)
        nav_frame.pack(fill='x', padx=20, pady=(0, 20))

        self.back_btn = ttk.Button(nav_frame, text="← Anterior", command=self._back_step, state='disabled')
        self.back_btn.pack(side='left')

        self.next_btn = ttk.Button(nav_frame, text="Siguiente →", command=self._next_step)
        self.next_btn.pack(side='right', padx=(0, 10))

        self.cancel_btn = ttk.Button(nav_frame, text="Cancelar", command=self._cancel_wizard)
        self.cancel_btn.pack(side='right')

        self.finish_btn = ttk.Button(nav_frame, text="Finalizar", command=self._finish_wizard, style='Accent.TButton')
        # finish_btn will be shown on last step

    def _show_step(self, step: int):
        """Show specific wizard step."""
        # Clear content frame
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        self.current_step = step
        self.progress_var.set(step + 1)

        if step == 0:
            self._show_welcome_step()
        elif step == 1:
            self._show_basic_config_step()
        elif step == 2:
            self._show_session_setup_step()
        elif step == 3:
            self._show_final_step()

        # Update navigation buttons
        self.back_btn.config(state='normal' if step > 0 else 'disabled')

        if step == self.total_steps - 1:
            self.next_btn.pack_forget()
            self.finish_btn.pack(side='right', padx=(0, 10))
        else:
            self.finish_btn.pack_forget()
            self.next_btn.pack(side='right', padx=(0, 10))

    def _show_welcome_step(self):
        """Show welcome step."""
        self.title_label.config(text="Bienvenido al Asistente de Configuración")

        welcome_text = """
Este asistente te guiará en la configuración inicial de la aplicación
de Verificación de Correos OWA.

Configuraremos los siguientes aspectos:
• Configuración básica de la aplicación
• Archivo de Excel con correos a procesar
• Sesión del navegador para acceso a OWA

Al finalizar, tendrás la aplicación lista para usar.

¿Deseas continuar con la configuración?
        """

        ttk.Label(self.content_frame, text=welcome_text, wraplength=550, justify='left').pack(pady=20)

    def _show_basic_config_step(self):
        """Show basic configuration step."""
        self.title_label.config(text="Configuración Básica")

        # URL configuration
        url_frame = ttk.LabelFrame(self.content_frame, text="URL de OWA", padding=10)
        url_frame.pack(fill='x', pady=(0, 10))

        self.url_var = tk.StringVar(value=self.config.page_url)
        ttk.Label(url_frame, text="URL del webmail:").pack(anchor='w')
        ttk.Entry(url_frame, textvariable=self.url_var, width=70).pack(fill='x', pady=(5, 0))

        # Excel file configuration
        excel_frame = ttk.LabelFrame(self.content_frame, text="Archivo de Excel", padding=10)
        excel_frame.pack(fill='x', pady=(0, 10))

        self.excel_path_var = tk.StringVar(value=self.config.get_excel_file_path())
        excel_entry_frame = ttk.Frame(excel_frame)
        excel_entry_frame.pack(fill='x')
        ttk.Entry(excel_entry_frame, textvariable=self.excel_path_var).pack(side='left', fill='x', expand=True, padx=(0, 5))
        ttk.Button(excel_entry_frame, text="Examinar", command=self._browse_excel_file).pack(side='right')

        ttk.Label(excel_frame, text="El archivo debe tener correos en la columna A, starting from row 2",
                 font=('TkDefaultFont', 9)).pack(anchor='w', pady=(5, 0))

        # Batch size
        batch_frame = ttk.LabelFrame(self.content_frame, text="Procesamiento", padding=10)
        batch_frame.pack(fill='x')

        self.batch_size_var = tk.IntVar(value=self.config.processing.batch_size)
        ttk.Label(batch_frame, text="Tamaño de lote (correos por lote):").pack(anchor='w')
        ttk.Entry(batch_frame, textvariable=self.batch_size_var, width=10).pack(anchor='w', pady=(5, 0))

    def _show_session_setup_step(self):
        """Show session setup step."""
        self.title_label.config(text="Configuración de Sesión")

        session_info = """
Para acceder al webmail de OWA, necesitas configurar una sesión de navegador.
La aplicación guardará tu sesión para que no tengas que iniciar sesión
cada vez que proceses correos.

Pasos para configurar la sesión:
1. Se abrirá una ventana del navegador
2. Inicia sesión manualmente en OWA
3. Cierra el navegador cuando hayas iniciado sesión
4. La aplicación guardará la sesión automáticamente

¿Listo para configurar la sesión?
        """

        ttk.Label(self.content_frame, text=session_info, wraplength=550, justify='left').pack(pady=20)

        session_frame = ttk.LabelFrame(self.content_frame, text="Configuración de Sesión", padding=10)
        session_frame.pack(fill='x', pady=(20, 0))

        self.session_file_var = tk.StringVar(value=self.config.get_session_file_path())
        ttk.Label(session_frame, text="Archivo de sesión:").pack(anchor='w')
        session_entry_frame = ttk.Frame(session_frame)
        session_entry_frame.pack(fill='x', pady=(5, 0))
        ttk.Entry(session_entry_frame, textvariable=self.session_file_var).pack(side='left', fill='x', expand=True, padx=(0, 5))
        ttk.Button(session_entry_frame, text="Examinar", command=self._browse_session_file).pack(side='right')

        # Setup session button
        ttk.Button(self.content_frame, text="🔧 Configurar Sesión Ahora",
                  command=self._setup_session_wizard, style='Accent.TButton').pack(pady=(20, 0))

    def _show_final_step(self):
        """Show final step."""
        self.title_label.config(text="Configuración Completa")

        final_text = """
¡Excelencial Has completado la configuración inicial.

Resumen de tu configuración:
• URL de OWA: {}
• Archivo de Excel: {}
• Archivo de sesión: {}
• Tamaño de lote: {}

La aplicación está lista para usar. Puedes:
1. Ir a la pestaña "Procesamiento" para empezar a verificar correos
2. Configurar la sesión del navegador si aún no lo has hecho
3. Ajustar cualquier configuración en la pestaña "Configuración"

¿Deseas finalizar el asistente?
        """.format(
            self.url_var.get() if hasattr(self, 'url_var') else self.config.page_url,
            self.excel_path_var.get() if hasattr(self, 'excel_path_var') else self.config.get_excel_file_path(),
            self.session_file_var.get() if hasattr(self, 'session_file_var') else self.config.get_session_file_path(),
            self.batch_size_var.get() if hasattr(self, 'batch_size_var') else self.config.processing.batch_size
        )

        ttk.Label(self.content_frame, text=final_text, wraplength=550, justify='left').pack(pady=20)

    def _browse_excel_file(self):
        """Browse for Excel file."""
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo de Excel",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_path_var.set(file_path)

    def _browse_session_file(self):
        """Browse for session file."""
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo de sesión",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if file_path:
            self.session_file_var.set(file_path)

    def _setup_session_wizard(self):
        """Setup browser session during wizard."""
        try:
            messagebox.showinfo(
                "Configuración de Sesión",
                "Se abrirá una ventana del navegador. Inicia sesión en OWA y luego cierra la ventana."
            )

            session_manager = SessionManager(self.config)
            if session_manager.setup_interactive_session():
                messagebox.showinfo("Éxito", "Sesión configurada correctamente")
            else:
                messagebox.showwarning("Advertencia", "No se pudo configurar la sesión automáticamente. Puedes hacerlo más tarde desde la aplicación.")
        except Exception as e:
            messagebox.showerror("Error", f"Error al configurar sesión: {e}")

    def _back_step(self):
        """Go to previous step."""
        if self.current_step > 0:
            self._show_step(self.current_step - 1)

    def _next_step(self):
        """Go to next step."""
        if self.current_step < self.total_steps - 1:
            # Validate current step
            if self._validate_current_step():
                self._show_step(self.current_step + 1)

    def _validate_current_step(self) -> bool:
        """Validate current wizard step."""
        if self.current_step == 1:  # Basic config step
            if not self.url_var.get().strip():
                messagebox.showerror("Error", "La URL de OWA es requerida")
                return False
            if not self.excel_path_var.get().strip():
                messagebox.showerror("Error", "La ruta del archivo Excel es requerida")
                return False
            if self.batch_size_var.get() <= 0:
                messagebox.showerror("Error", "El tamaño de lote debe ser mayor que 0")
                return False
        return True

    def _cancel_wizard(self):
        """Cancel wizard."""
        if messagebox.askyesno("Cancelar", "¿Estás seguro de que deseas cancelar el asistente de configuración?"):
            self.wizard.destroy()

    def _finish_wizard(self):
        """Finish wizard and save configuration."""
        try:
            # Update configuration with wizard values
            if hasattr(self, 'url_var'):
                self.config.page_url = self.url_var.get()
            if hasattr(self, 'excel_path_var'):
                self.config.excel.default_file = self.excel_path_var.get()
            if hasattr(self, 'session_file_var'):
                self.config.browser.session_file = self.session_file_var.get()
            if hasattr(self, 'batch_size_var'):
                self.config.processing.batch_size = self.batch_size_var.get()

            config_data = VerificacionCorreosGUI._build_config_yaml(self.config)
            with open(self.config._config_path, 'w', encoding='utf-8') as f:
                yaml.dump(config_data, f, default_flow_style=False, allow_unicode=True)

            self.result = True
            messagebox.showinfo("Éxito", "Configuración guardada correctamente")
            self.wizard.destroy()

        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar configuración: {e}")


def main():
    """Main entry point for GUI."""
    root = tk.Tk()

    # Set up modern styling if available
    try:
        style = ttk.Style()
        if 'clam' in style.theme_names():
            style.theme_use('clam')
        style.configure('Accent.TButton', font=('TkDefaultFont', 10, 'bold'))
        style.configure('Treeview', rowheight=28)
    except tk.TclError:
        pass

    app = VerificacionCorreosGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()