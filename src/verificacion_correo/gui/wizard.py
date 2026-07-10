"""
Configuration wizard for first-time setup.

This module provides the ConfigWizard class that guides users through
initial configuration of the application.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from verificacion_correo.core.config import Config
from verificacion_correo.core.session import SessionManager


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

            self.config.save()

            self.result = True
            messagebox.showinfo("Éxito", "Configuración guardada correctamente")
            self.wizard.destroy()

        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar configuración: {e}")
