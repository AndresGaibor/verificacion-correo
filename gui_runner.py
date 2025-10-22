"""
Ejecutor de procesamiento en segundo plano para la interfaz gráfica.
Permite ejecutar la verificación de correos sin bloquear la GUI.
"""

import threading
import queue
import time
from typing import Callable, Dict, Any, Optional
from excel_reader import leer_correos_pendientes
from browser_automation import procesar_todos_los_lotes
from gui_config_manager import GUIConfigManager


class GUIRunner:
    """Ejecuta el procesamiento de correos en segundo plano."""

    def __init__(self):
        self.config_manager = GUIConfigManager()
        self.is_running = False
        self.is_paused = False
        self.should_stop = False
        self.progress_queue = queue.Queue()
        self.current_thread = None

        # Callbacks para actualizar la GUI
        self.progress_callback: Optional[Callable[[int, int, str], None]] = None
        self.log_callback: Optional[Callable[[str], None]] = None
        self.complete_callback: Optional[Callable[[Dict[str, int]], None]] = None
        self.error_callback: Optional[Callable[[str], None]] = None

    def set_callbacks(self,
                     progress_callback: Callable[[int, int, str], None] = None,
                     log_callback: Callable[[str], None] = None,
                     complete_callback: Callable[[Dict[str, int]], None] = None,
                     error_callback: Callable[[str], None] = None):
        """
        Establece los callbacks para comunicación con la GUI.

        Args:
            progress_callback: Función llamada con (current, total, message)
            log_callback: Función llamada con mensaje de log
            complete_callback: Función llamada con estadísticas finales
            error_callback: Función llamada con mensaje de error
        """
        self.progress_callback = progress_callback
        self.log_callback = log_callback
        self.complete_callback = complete_callback
        self.error_callback = error_callback

    def start_processing(self, excel_file: str = None):
        """
        Inicia el procesamiento de correos en segundo plano.

        Args:
            excel_file: Ruta al archivo Excel (opcional, usa config si no se especifica)
        """
        if self.is_running:
            self._safe_log("❌ Ya hay un proceso en ejecución")
            return

        # Determinar archivo Excel
        if not excel_file:
            excel_file = self.config_manager.get_config_value('excel.default_file')

        # Recargar configuración
        self.config_manager.reload_config()

        # Iniciar thread
        self.should_stop = False
        self.is_paused = False
        self.current_thread = threading.Thread(
            target=self._process_emails,
            args=(excel_file,),
            daemon=True
        )
        self.current_thread.start()
        self.is_running = True

        # Iniciar monitor de cola
        self._start_queue_monitor()

    def stop_processing(self):
        """Detiene el procesamiento actual."""
        if not self.is_running:
            return

        self.should_stop = True
        self._safe_log("🛑 Deteniendo procesamiento...")

        # Esperar a que el thread termine (con timeout)
        if self.current_thread and self.current_thread.is_alive():
            self.current_thread.join(timeout=5)

    def pause_processing(self):
        """Pausa el procesamiento actual."""
        if not self.is_running:
            return

        self.is_paused = True
        self._safe_log("⏸ Procesamiento pausado")

    def resume_processing(self):
        """Reanuda el procesamiento pausado."""
        if not self.is_running or not self.is_paused:
            return

        self.is_paused = False
        self._safe_log("▶ Procesamiento reanudado")

    def _process_emails(self, excel_file: str):
        """
        Función principal de procesamiento ejecutada en thread separado.

        Args:
            excel_file: Ruta al archivo Excel
        """
        try:
            self._safe_log("🔍 Iniciando verificación de correos...")
            self._safe_log(f"📁 Archivo Excel: {excel_file}")

            # 1. Leer emails pendientes
            self._safe_progress(0, 100, "Leyendo correos desde Excel...")
            lotes_info = leer_correos_pendientes(excel_file)

            # Validar que haya correos para procesar
            if lotes_info['total_pendientes'] == 0:
                self._safe_log("✅ No hay correos pendientes para procesar")
                stats = {'ok': 0, 'no_existe': 0, 'error': 0}
                self._queue_put(('complete', stats))
                return

            # 2. Mostrar resumen inicial
            total_emails = lotes_info['total_pendientes']
            total_lotes = len(lotes_info['lotes'])

            self._safe_log(f"📊 Resumen:")
            self._safe_log(f"   • Total emails: {lotes_info['total_emails']}")
            self._safe_log(f"   • Ya procesados: {lotes_info['total_procesados']}")
            self._safe_log(f"   • Pendientes: {total_emails}")
            self._safe_log(f"   • En {total_lotes} lotes")

            # 3. Procesar lotes
            self._safe_progress(0, 100, "Iniciando procesamiento de lotes...")
            stats = procesar_todos_los_lotes(lotes_info, excel_file)

            # 4. Enviar resultado final
            self._queue_put(('complete', stats))

        except Exception as e:
            self._queue_put(('error', f"Error en procesamiento: {str(e)}"))
        finally:
            self.is_running = False

    def _start_queue_monitor(self):
        """Inicia el monitor de la cola de mensajes en el thread principal."""
        def monitor():
            while self.is_running or not self.progress_queue.empty():
                try:
                    # Obtener mensaje con timeout corto
                    msg_type, data = self.progress_queue.get(timeout=0.1)

                    # Procesar mensaje según tipo
                    if msg_type == 'progress':
                        current, total, message = data
                        if self.progress_callback:
                            self.progress_callback(current, total, message)
                    elif msg_type == 'log':
                        if self.log_callback:
                            self.log_callback(data)
                    elif msg_type == 'complete':
                        if self.complete_callback:
                            self.complete_callback(data)
                        break
                    elif msg_type == 'error':
                        if self.error_callback:
                            self.error_callback(data)
                        break

                except queue.Empty:
                    continue
                except Exception as e:
                    print(f"Error en monitor de cola: {e}")

        # Ejecutar monitor en thread principal si estamos en GUI
        if threading.current_thread() == threading.main_thread():
            monitor()
        else:
            # Si no estamos en main thread, crear nuevo thread para monitor
            monitor_thread = threading.Thread(target=monitor, daemon=True)
            monitor_thread.start()

    def _queue_put(self, message):
        """Añade un mensaje a la cola de forma segura."""
        try:
            self.progress_queue.put_nowait(message)
        except queue.Full:
            pass  # Ignorar si la cola está llena

    def _safe_progress(self, current: int, total: int, message: str):
        """Envía actualización de progreso de forma segura."""
        self._queue_put(('progress', (current, total, message)))

    def _safe_log(self, message: str):
        """Envía mensaje de log de forma segura."""
        self._queue_put(('log', message))

    def get_processing_stats(self) -> Dict[str, Any]:
        """
        Retorna el estado actual del procesamiento.

        Returns:
            Diccionario con estado actual
        """
        return {
            'is_running': self.is_running,
            'is_paused': self.is_paused,
            'can_start': not self.is_running,
            'can_stop': self.is_running,
            'can_pause': self.is_running and not self.is_paused,
            'can_resume': self.is_running and self.is_paused
        }

    def validate_prerequisites(self) -> tuple[bool, str]:
        """
        Valida que los prerrequisitos para el procesamiento estén cumplidos.

        Returns:
            tuple: (is_valid, error_message)
        """
        # Verificar archivo de sesión
        session_file = "state.json"  # Archivo fijo de sesión
        if not session_file or not os.path.exists(session_file):
            return False, "No se encontró el archivo de sesión. Use la pestaña '🔐 Sesión' para crear una."

        # Verificar archivo Excel
        excel_file = self.config_manager.get_config_value('excel.default_file')
        if not excel_file or not os.path.exists(excel_file):
            return False, f"No se encontró el archivo Excel: {excel_file}"

        # Verificar URL
        page_url = self.config_manager.get_config_value('page_url')
        if not page_url:
            return False, "La URL de la página no está configurada."

        return True, ""


# Importar aquí para evitar importación circular
import os