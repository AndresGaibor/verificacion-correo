#!/usr/bin/env python3
"""
Interfaz gráfica principal para la verificación de correos OWA.
Implementa una GUI con pestañas para procesamiento, configuración y gestión de sesiones.
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
from datetime import datetime
from gui_config_manager import GUIConfigManager
from gui_runner import GUIRunner
from gui_session_manager import GUISessionManager


class VerificacionCorreosGUI:
    """Interfaz gráfica principal para verificación de correos."""

    def __init__(self, root):
        self.root = root
        self.root.title("Verificación de Correos OWA")
        self.root.geometry("900x650")
        self.root.resizable(True, True)

        # Inicializar componentes
        self.config_manager = GUIConfigManager()
        self.runner = GUIRunner()
        self.session_manager = GUISessionManager(
            log_callback=self.add_log,
            status_callback=self.update_session_status
        )

        # Configurar callbacks del runner
        self.runner.set_callbacks(
            progress_callback=self.update_progress,
            log_callback=self.add_log,
            complete_callback=self.processing_complete,
            error_callback=self.processing_error
        )

        # Variables de estado
        self.current_progress = 0
        self.total_progress = 100
        self.processing_active = False
        self.session_status = tk.StringVar(value="Verificando sesión...")

        # Crear interfaz
        self._create_widgets()
        self._load_config_to_gui()
        self._check_session_status()

    def _create_widgets(self):
        """Crea todos los widgets de la interfaz."""
        # Notebook (pestañas)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # Pestaña de procesamiento
        self._create_processing_tab()

        # Pestaña de gestión de sesiones
        self._create_session_tab()

        # Pestaña de configuración
        self._create_config_tab()

        # Barra de estado
        self._create_status_bar()

    def _create_processing_tab(self):
        """Crea la pestaña de procesamiento."""
        self.processing_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.processing_frame, text="Procesamiento")

        # Frame principal
        main_frame = ttk.Frame(self.processing_frame)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Botones de control
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill='x', pady=(0, 10))

        self.start_btn = ttk.Button(
            buttons_frame,
            text="🚀 Iniciar Verificación",
            command=self.start_processing,
            style='Accent.TButton'
        )
        self.start_btn.pack(side='left', padx=(0, 5))

        self.pause_btn = ttk.Button(
            buttons_frame,
            text="⏸ Pausar",
            command=self.pause_processing,
            state='disabled'
        )
        self.pause_btn.pack(side='left', padx=5)

        self.stop_btn = ttk.Button(
            buttons_frame,
            text="🛑 Detener",
            command=self.stop_processing,
            state='disabled'
        )
        self.stop_btn.pack(side='left', padx=5)

        # Frame de estado de sesión
        session_frame = ttk.LabelFrame(main_frame, text="Estado de Sesión OWA")
        session_frame.pack(fill='x', pady=(0, 10))

        session_inner = ttk.Frame(session_frame)
        session_inner.pack(fill='x', padx=10, pady=5)

        # Indicador de estado
        self.session_status_label = ttk.Label(
            session_inner,
            textvariable=self.session_status,
            font=('Arial', 9)
        )
        self.session_status_label.pack(side='left', padx=(0, 10))

        # Botón para gestionar sesión
        self.manage_session_btn = ttk.Button(
            session_inner,
            text="🔐 Gestionar Sesión",
            command=self.open_session_tab
        )
        self.manage_session_btn.pack(side='right')

        # Frame de progreso
        progress_frame = ttk.LabelFrame(main_frame, text="Progreso")
        progress_frame.pack(fill='x', pady=(0, 10))

        # Barra de progreso
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.pack(fill='x', padx=10, pady=5)

        # Etiqueta de estado
        self.status_label = ttk.Label(
            progress_frame,
            text="Listo para iniciar",
            font=('Arial', 9)
        )
        self.status_label.pack(pady=(0, 5))

        # Frame de estadísticas
        stats_frame = ttk.LabelFrame(main_frame, text="Estadísticas")
        stats_frame.pack(fill='x', pady=(0, 10))

        stats_inner = ttk.Frame(stats_frame)
        stats_inner.pack(fill='x', padx=10, pady=5)

        self.ok_label = ttk.Label(stats_inner, text="✅ OK: 0", foreground='green')
        self.ok_label.pack(side='left', padx=(0, 20))

        self.no_existe_label = ttk.Label(stats_inner, text="❌ No Existe: 0", foreground='red')
        self.no_existe_label.pack(side='left', padx=(0, 20))

        self.error_label = ttk.Label(stats_inner, text="⚠️ Error: 0", foreground='orange')
        self.error_label.pack(side='left')

        # Frame de logs
        log_frame = ttk.LabelFrame(main_frame, text="Log de Actividad")
        log_frame.pack(fill='both', expand=True)

        # Área de texto con scroll
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=15,
            wrap=tk.WORD,
            font=('Consolas', 9)
        )
        self.log_text.pack(fill='both', expand=True, padx=5, pady=5)

        # Botón para limpiar log
        clear_log_btn = ttk.Button(
            log_frame,
            text="🗑️ Limpiar Log",
            command=self.clear_log
        )
        clear_log_btn.pack(pady=(5, 0))

    def _create_session_tab(self):
        """Crea la pestaña de gestión de sesiones."""
        self.session_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.session_frame, text="🔐 Sesión")

        # Frame principal
        main_frame = ttk.Frame(self.session_frame)
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)

        # Información de sesión actual
        info_frame = ttk.LabelFrame(main_frame, text="Estado Actual")
        info_frame.pack(fill='x', pady=(0, 20))

        info_inner = ttk.Frame(info_frame)
        info_inner.pack(fill='x', padx=10, pady=10)

        # Archivo de sesión
        ttk.Label(info_inner, text="Archivo de Sesión:").grid(row=0, column=0, sticky='w', pady=2)
        self.current_session_file_var = tk.StringVar(value="state.json")
        self.session_file_label = ttk.Label(info_inner, textvariable=self.current_session_file_var)
        self.session_file_label.grid(row=0, column=1, sticky='w', padx=(10, 0), pady=2)

        # Estado detallado
        ttk.Label(info_inner, text="Estado:").grid(row=1, column=0, sticky='w', pady=2)
        self.detailed_session_status_var = tk.StringVar(value="Verificando...")
        self.detailed_session_status_label = ttk.Label(
            info_inner,
            textvariable=self.detailed_session_status_var,
            font=('Arial', 9, 'bold')
        )
        self.detailed_session_status_label.grid(row=1, column=1, sticky='w', padx=(10, 0), pady=2)

        # Última verificación
        ttk.Label(info_inner, text="Última verificación:").grid(row=2, column=0, sticky='w', pady=2)
        self.last_check_var = tk.StringVar(value="Nunca")
        self.last_check_label = ttk.Label(info_inner, textvariable=self.last_check_var)
        self.last_check_label.grid(row=2, column=1, sticky='w', padx=(10, 0), pady=2)

        # Acciones
        actions_frame = ttk.LabelFrame(main_frame, text="Acciones")
        actions_frame.pack(fill='x', pady=(0, 20))

        actions_inner = ttk.Frame(actions_frame)
        actions_inner.pack(fill='x', padx=10, pady=10)

        # Botones principales
        buttons_row1 = ttk.Frame(actions_inner)
        buttons_row1.pack(fill='x', pady=(0, 10))

        self.create_session_btn = ttk.Button(
            buttons_row1,
            text="🚀 Crear Nueva Sesión",
            command=self.create_new_session,
            style='Accent.TButton'
        )
        self.create_session_btn.pack(side='left', padx=(0, 10))

        self.validate_session_btn = ttk.Button(
            buttons_row1,
            text="🔍 Validar Sesión Actual",
            command=self.validate_current_session
        )
        self.validate_session_btn.pack(side='left', padx=(0, 10))

        self.refresh_status_btn = ttk.Button(
            buttons_row1,
            text="🔄 Refrescar Estado",
            command=self.refresh_session_status
        )
        self.refresh_status_btn.pack(side='left')

        # Estado de creación
        creation_frame = ttk.LabelFrame(main_frame, text="Creación de Sesión")
        creation_frame.pack(fill='x', pady=(0, 20))

        creation_inner = ttk.Frame(creation_frame)
        creation_inner.pack(fill='x', padx=10, pady=10)

        # Instrucciones
        self.creation_instructions = tk.Text(
            creation_inner,
            height=8,
            wrap=tk.WORD,
            font=('Arial', 9),
            state='disabled'
        )
        self.creation_instructions.pack(fill='x', pady=(0, 10))

        # Botones de control de creación
        creation_buttons = ttk.Frame(creation_inner)
        creation_buttons.pack(fill='x')

        self.save_session_btn = ttk.Button(
            creation_buttons,
            text="💾 Guardar Sesión",
            command=self.save_current_session,
            state='disabled'
        )
        self.save_session_btn.pack(side='left', padx=(0, 10))

        self.cancel_session_btn = ttk.Button(
            creation_buttons,
            text="❌ Cancelar Creación",
            command=self.cancel_session_creation,
            state='disabled'
        )
        self.cancel_session_btn.pack(side='left')

        # Logs de sesión
        log_frame = ttk.LabelFrame(main_frame, text="Log de Sesión")
        log_frame.pack(fill='both', expand=True)

        # Área de texto con scroll
        self.session_log_text = scrolledtext.ScrolledText(
            log_frame,
            height=10,
            wrap=tk.WORD,
            font=('Consolas', 9)
        )
        self.session_log_text.pack(fill='both', expand=True, padx=5, pady=5)

        # Botón para limpiar log de sesión
        clear_session_log_btn = ttk.Button(
            log_frame,
            text="🗑️ Limpiar Log",
            command=self.clear_session_log
        )
        clear_session_log_btn.pack(pady=(5, 0))

        # Inicializar estado
        self._update_session_ui_state()

    def _create_config_tab(self):
        """Crea la pestaña de configuración."""
        self.config_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.config_frame, text="Configuración")

        # Scrollable frame
        canvas = tk.Canvas(self.config_frame)
        scrollbar = ttk.Scrollbar(self.config_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Contenido de configuración
        main_frame = ttk.Frame(scrollable_frame)
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)

        # Configuración de OWA
        owa_frame = ttk.LabelFrame(main_frame, text="Configuración OWA")
        owa_frame.pack(fill='x', pady=(0, 15))

        ttk.Label(owa_frame, text="URL de OWA:").grid(row=0, column=0, sticky='w', padx=10, pady=5)
        self.url_var = tk.StringVar()
        self.url_entry = ttk.Entry(owa_frame, textvariable=self.url_var, width=50)
        self.url_entry.grid(row=0, column=1, padx=(0, 10), pady=5, sticky='ew')
        owa_frame.columnconfigure(1, weight=1)

        # Configuración de procesamiento
        proc_frame = ttk.LabelFrame(main_frame, text="Configuración de Procesamiento")
        proc_frame.pack(fill='x', pady=(0, 15))

        ttk.Label(proc_frame, text="Tamaño de Lote:").grid(row=0, column=0, sticky='w', padx=10, pady=5)
        self.batch_size_var = tk.StringVar(value="10")
        self.batch_size_spinbox = ttk.Spinbox(
            proc_frame,
            from_=1,
            to=50,
            textvariable=self.batch_size_var,
            width=10
        )
        self.batch_size_spinbox.grid(row=0, column=1, sticky='w', padx=(0, 10), pady=5)

        ttk.Label(proc_frame, text="emails por lote").grid(row=0, column=2, sticky='w', pady=5)

        # Configuración del navegador
        browser_frame = ttk.LabelFrame(main_frame, text="Configuración del Navegador")
        browser_frame.pack(fill='x', pady=(0, 15))

        self.headless_var = tk.BooleanVar(value=False)
        self.headless_check = ttk.Checkbutton(
            browser_frame,
            text="Modo Headless (sin ventana visible)",
            variable=self.headless_var
        )
        self.headless_check.pack(anchor='w', padx=10, pady=5)

        # Configuración de archivos
        files_frame = ttk.LabelFrame(main_frame, text="Configuración de Archivos")
        files_frame.pack(fill='x', pady=(0, 15))

        ttk.Label(files_frame, text="Archivo Excel:").grid(row=0, column=0, sticky='w', padx=10, pady=5)
        self.excel_file_var = tk.StringVar()
        self.excel_file_entry = ttk.Entry(files_frame, textvariable=self.excel_file_var, width=40)
        self.excel_file_entry.grid(row=0, column=1, padx=(0, 5), pady=5, sticky='ew')

        browse_btn = ttk.Button(
            files_frame,
            text="📁 Explorar",
            command=self.browse_excel_file
        )
        browse_btn.grid(row=0, column=2, padx=(0, 10), pady=5)
        files_frame.columnconfigure(1, weight=1)

        # Archivo de sesión
        ttk.Label(files_frame, text="Archivo Sesión:").grid(row=1, column=0, sticky='w', padx=10, pady=5)
        self.session_file_var = tk.StringVar(value="state.json")
        self.session_file_entry = ttk.Entry(files_frame, textvariable=self.session_file_var, width=40)
        self.session_file_entry.grid(row=1, column=1, padx=(0, 5), pady=5, sticky='ew')

        browse_session_btn = ttk.Button(
            files_frame,
            text="📁 Explorar",
            command=self.browse_session_file
        )
        browse_session_btn.grid(row=1, column=2, padx=(0, 10), pady=5)

        # Botones de acción
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill='x', pady=(20, 0))

        save_btn = ttk.Button(
            action_frame,
            text="💾 Guardar Configuración",
            command=self.save_config,
            style='Accent.TButton'
        )
        save_btn.pack(side='left', padx=(0, 10))

        reload_btn = ttk.Button(
            action_frame,
            text="🔄 Recargar Configuración",
            command=self.reload_config
        )
        reload_btn.pack(side='left')

        # Empaquetar canvas
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def _create_status_bar(self):
        """Crea la barra de estado."""
        self.status_bar = ttk.Frame(self.root)
        self.status_bar.pack(fill='x', side='bottom')

        self.status_text = tk.StringVar(value="Listo")
        status_label = ttk.Label(
            self.status_bar,
            textvariable=self.status_text,
            relief='sunken',
            anchor='w'
        )
        status_label.pack(side='left', fill='x', expand=True, padx=2, pady=2)

        # Hora actual
        self.time_var = tk.StringVar()
        self.update_time()
        time_label = ttk.Label(
            self.status_bar,
            textvariable=self.time_var,
            relief='sunken',
            width=20
        )
        time_label.pack(side='right', padx=2, pady=2)

    def update_time(self):
        """Actualiza la hora en la barra de estado."""
        self.time_var.set(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.root.after(1000, self.update_time)

    def _load_config_to_gui(self):
        """Carga la configuración actual en la interfaz."""
        settings = self.config_manager.get_current_settings()

        self.url_var.set(settings.get('page_url', ''))
        self.batch_size_var.set(str(settings.get('batch_size', 10)))
        self.headless_var.set(settings.get('headless', False))
        self.excel_file_var.set(settings.get('excel_file', ''))
        self.session_file_var.set('state.json')  # Siempre state.json

    def start_processing(self):
        """Inicia el procesamiento de correos."""
        # Validar prerrequisitos
        valid, error_msg = self.runner.validate_prerequisites()
        if not valid:
            messagebox.showerror("Error de Prerrequisitos", error_msg)
            return

        # Guardar configuración actual
        if not self.save_config(silent=True):
            return

        # Actualizar estado de botones
        self._update_button_states(can_start=False, can_stop=True, can_pause=True)

        # Iniciar procesamiento
        self.processing_active = True
        self.add_log("🚀 Iniciando verificación de correos...")
        self.runner.start_processing()

    def pause_processing(self):
        """Pausa el procesamiento."""
        self.runner.pause_processing()
        self._update_button_states(can_start=False, can_stop=True, can_pause=False, can_resume=True)

    def resume_processing(self):
        """Reanuda el procesamiento."""
        self.runner.resume_processing()
        self._update_button_states(can_start=False, can_stop=True, can_pause=True, can_resume=False)

    def stop_processing(self):
        """Detiene el procesamiento."""
        self.runner.stop_processing()
        self._update_button_states(can_start=True, can_stop=False, can_pause=False)

    def _update_button_states(self, can_start=True, can_stop=False, can_pause=False, can_resume=False):
        """Actualiza el estado de los botones."""
        self.start_btn.config(state='normal' if can_start else 'disabled')
        self.stop_btn.config(state='normal' if can_stop else 'disabled')
        self.pause_btn.config(state='normal' if can_pause else 'disabled')

        # Actualizar texto del botón de pausa/reanudar
        if can_resume:
            self.pause_btn.config(text="▶ Reanudar")
        else:
            self.pause_btn.config(text="⏸ Pausar")

    def update_progress(self, current: int, total: int, message: str):
        """Actualiza la barra de progreso y el mensaje de estado."""
        # Actualizar variables
        self.current_progress = current
        self.total_progress = total if total > 0 else 1

        # Calcular porcentaje
        percentage = (current / self.total_progress) * 100

        # Actualizar GUI en thread principal
        self.root.after(0, self._update_progress_ui, percentage, message)

    def _update_progress_ui(self, percentage: float, message: str):
        """Actualiza los elementos UI de progreso."""
        self.progress_var.set(percentage)
        self.status_label.config(text=message)
        self.status_text.set(message)

    def add_log(self, message: str):
        """Añade un mensaje al log."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"

        # Actualizar en thread principal
        self.root.after(0, self._add_log_to_ui, log_entry)

    def _add_log_to_ui(self, log_entry: str):
        """Añade entrada al log en la UI."""
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)

    def processing_complete(self, stats: dict):
        """Maneja la finalización del procesamiento."""
        self.processing_active = False
        self._update_button_states(can_start=True, can_stop=False, can_pause=False)

        # Actualizar estadísticas
        self.ok_label.config(text=f"✅ OK: {stats['ok']}")
        self.no_existe_label.config(text=f"❌ No Existe: {stats['no_existe']}")
        self.error_label.config(text=f"⚠️ Error: {stats['error']}")

        # Mostrar resumen
        total = stats['ok'] + stats['no_existe'] + stats['error']
        self.add_log("="*50)
        self.add_log("📊 PROCESAMIENTO COMPLETADO")
        self.add_log("="*50)
        self.add_log(f"✅ Contactos encontrados: {stats['ok']}")
        self.add_log(f"❌ No existen: {stats['no_existe']}")
        self.add_log(f"⚠️ Errores: {stats['error']}")
        self.add_log(f"📈 Total procesado: {total}")

        # Mostrar mensaje de completado
        messagebox.showinfo(
            "Procesamiento Completado",
            f"Procesamiento finalizado:\n\n"
            f"✅ OK: {stats['ok']}\n"
            f"❌ No Existe: {stats['no_existe']}\n"
            f"⚠️ Error: {stats['error']}\n"
            f"📈 Total: {total}"
        )

    def processing_error(self, error_message: str):
        """Maneja errores durante el procesamiento."""
        self.processing_active = False
        self._update_button_states(can_start=True, can_stop=False, can_pause=False)

        self.add_log(f"❌ ERROR: {error_message}")
        messagebox.showerror("Error de Procesamiento", error_message)

    def clear_log(self):
        """Limpia el área de log."""
        self.log_text.delete(1.0, tk.END)

    def browse_excel_file(self):
        """Abre diálogo para seleccionar archivo Excel."""
        filename = filedialog.askopenfilename(
            title="Seleccionar Archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
        )
        if filename:
            self.excel_file_var.set(filename)

    def browse_session_file(self):
        """Abre diálogo para seleccionar archivo de sesión."""
        filename = filedialog.askopenfilename(
            title="Seleccionar Archivo de Sesión",
            filetypes=[("Archivos JSON", "*.json"), ("Todos los archivos", "*.*")]
        )
        if filename:
            self.session_file_var.set(filename)

    def save_config(self, silent=False):
        """Guarda la configuración actual."""
        settings = {
            'page_url': self.url_var.get().strip(),
            'batch_size': self.batch_size_var.get(),
            'headless': self.headless_var.get(),
            'excel_file': self.excel_file_var.get().strip()
        }

        # Validar y guardar
        success, error_msg = self.config_manager.update_settings(settings)

        if success:
            success = self.config_manager.save_config()
            if success:
                if not silent:
                    messagebox.showinfo("Configuración Guardada", "✅ Configuración guardada exitosamente")
                self.add_log("💾 Configuración guardada")
            else:
                if not silent:
                    messagebox.showerror("Error", "❌ No se pudo guardar la configuración")
        else:
            if not silent:
                messagebox.showerror("Error de Validación", f"❌ {error_msg}")

        return success

    def reload_config(self):
        """Recarga la configuración desde el archivo."""
        self.config_manager.reload_config()
        self._load_config_to_gui()
        self.add_log("🔄 Configuración recargada")
        messagebox.showinfo("Configuración Recargada", "✅ Configuración recargada exitosamente")

    # === Métodos de Gestión de Sesiones ===

    def _check_session_status(self):
        """Verifica el estado actual de la sesión al iniciar."""
        self.refresh_session_status()

    def refresh_session_status(self):
        """Refresca el estado de la sesión actual."""
        session_file = self.current_session_file_var.get()
        status_info = self.session_manager.get_session_status(session_file)

        if status_info["exists"]:
            if status_info.get("age_hours", 0) > 24:
                status_text = f"⚠️ Antigua ({status_info.get('age_str', 'Desconocido')})"
                self.detailed_session_status_var.set(status_text)
                self.session_status.set("Sesión antigua - requiere validación")
            else:
                status_text = f"✅ {status_info.get('age_str', 'Reciente')}"
                self.detailed_session_status_var.set(status_text)
                self.session_status.set("Sesión disponible")
        else:
            self.detailed_session_status_var.set("❌ No existe")
            self.session_status.set("No hay sesión - créela primero")

        # Actualizar última verificación
        self.last_check_var.set(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    def open_session_tab(self):
        """Abre la pestaña de gestión de sesiones."""
        self.notebook.select(self.session_frame)

    def create_new_session(self):
        """Inicia la creación de una nueva sesión."""
        if self.session_manager.is_creating_session:
            messagebox.showwarning("Sesión en Progreso", "Ya hay una sesión en creación.")
            return

        # Actualizar estado UI
        self._update_session_ui_state(creating=True)

        # Mostrar instrucciones
        instructions = """📋 Instrucciones para crear una nueva sesión:

1. Se abrirá una nueva ventana del navegador
2. Inicia sesión en OWA con tus credenciales
3. Espera a que la página principal cargue completamente
4. Una vez autenticado, vuelve a esta ventana
5. Haz clic en '💾 Guardar Sesión'

⚠️ Importante:
- Cierra todas las otras ventanas del navegador antes de empezar
- Asegúrate de estar en la página principal de OWA antes de guardar
- La sesión se guardará en el archivo: state.json"""

        self._update_instructions(instructions)

        # Iniciar creación de sesión en background
        session_file = self.current_session_file_var.get()
        self.session_manager.create_session_interactive(session_file)

    def validate_current_session(self):
        """Valida la sesión actual."""
        self._add_session_log("🔍 Iniciando validación de sesión...")

        # Ejecutar validación en thread separado
        def validate_thread():
            session_file = self.current_session_file_var.get()
            is_valid, message = self.session_manager.validate_session(session_file)

            # Actualizar UI en thread principal
            self.root.after(0, self._handle_validation_result, is_valid, message)

        threading.Thread(target=validate_thread, daemon=True).start()

    def _handle_validation_result(self, is_valid, message):
        """Maneja el resultado de la validación de sesión."""
        if is_valid:
            self._add_session_log("✅ Sesión válida y activa")
            self.detailed_session_status_var.set("✅ Válida")
            self.session_status.set("Sesión válida")
            messagebox.showinfo("Validación Exitosa", "✅ La sesión actual es válida y está activa.")
        else:
            self._add_session_log(f"❌ Sesión inválida: {message}")
            self.detailed_session_status_var.set("❌ Inválida")
            self.session_status.set("Sesión inválida - requiere nueva sesión")
            messagebox.showwarning("Sesión Inválida", f"❌ La sesión actual no es válida:\n{message}")

    def save_current_session(self):
        """Guarda la sesión actual."""
        if not self.session_manager.is_creating_session:
            messagebox.showwarning("Sin Sesión Activa", "No hay ninguna sesión en creación para guardar.")
            return

        session_file = self.current_session_file_var.get()
        success = self.session_manager.save_session_now(session_file)

        if success:
            self._add_session_log(f"✅ Sesión guardada en {session_file}")
            self._update_session_ui_state(creating=False)
            self.refresh_session_status()
            messagebox.showinfo("Sesión Guardada", f"✅ Sesión guardada exitosamente en:\n{session_file}")

    def cancel_session_creation(self):
        """Cancela la creación de sesión."""
        self.session_manager.cancel_session_creation()
        self._add_session_log("🛑 Creación de sesión cancelada")
        self._update_session_ui_state(creating=False)
        self._update_instructions("")

    def clear_session_log(self):
        """Limpia el log de sesión."""
        self.session_log_text.delete(1.0, tk.END)

    def _add_session_log(self, message):
        """Añade un mensaje al log de sesión."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        self.session_log_text.insert(tk.END, log_entry)
        self.session_log_text.see(tk.END)

    def update_session_status(self, status):
        """Actualiza el estado de sesión (callback del session manager)."""
        self.root.after(0, self.session_status.set, status)

    def _update_instructions(self, text):
        """Actualiza el texto de instrucciones."""
        self.creation_instructions.config(state='normal')
        self.creation_instructions.delete(1.0, tk.END)
        self.creation_instructions.insert(tk.END, text)
        self.creation_instructions.config(state='disabled')

    def _update_session_ui_state(self, creating=False):
        """Actualiza el estado de los botones según el contexto."""
        if creating:
            # Estado: creando sesión
            self.create_session_btn.config(state='disabled')
            self.validate_session_btn.config(state='disabled')
            self.save_session_btn.config(state='normal')
            self.cancel_session_btn.config(state='normal')
        else:
            # Estado: normal
            self.create_session_btn.config(state='normal')
            self.validate_session_btn.config(state='normal')
            self.save_session_btn.config(state='disabled')
            self.cancel_session_btn.config(state='disabled')


def main():
    """Función principal para ejecutar la GUI."""
    root = tk.Tk()
    app = VerificacionCorreosGUI(root)

    # Configurar estilo para botones de acento
    style = ttk.Style()
    style.configure('Accent.TButton', font=('Arial', 10, 'bold'))

    try:
        root.mainloop()
    except KeyboardInterrupt:
        print("\nAplicación cerrada por el usuario")


if __name__ == "__main__":
    main()