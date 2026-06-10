"""
GUI para el scraper de contactos de Outlook
Proporciona una interfaz gráfica para configurar y ejecutar el scraper
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from windows_compat import setup_console_encoding
setup_console_encoding()

import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import asyncio
import threading
from pathlib import Path
from datetime import datetime

# Importar el scraper
from debug_scraper import DebugScraper, scrape_outlook_contacts, logger


class ScraperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Scraper de Contactos Outlook")
        self.root.geometry("700x600")
        self.root.resizable(False, False)
        
        # Variables de control
        self.directorio_salida = tk.StringVar(value=str(Path.cwd() / "data"))
        self.cantidad_contactos = tk.IntVar(value=100)
        self.contactos_extraidos = tk.IntVar(value=0)
        self.scraper_activo = False
        self.scraper_thread = None
        self.loop = None
        self.scraper = None
        self.detener_flag = False
        
        # Crear interfaz
        self._crear_interfaz()
        
    def _crear_interfaz(self):
        """Crea todos los elementos de la interfaz"""
        
        # ===== HEADER =====
        header_frame = tk.Frame(self.root, bg="#2c3e50", height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame,
            text="🔍 Scraper de Contactos Outlook",
            font=("Helvetica", 20, "bold"),
            bg="#2c3e50",
            fg="white"
        )
        title_label.pack(pady=20)
        
        # ===== CONFIGURACIÓN =====
        config_frame = tk.LabelFrame(
            self.root,
            text="⚙️ Configuración",
            font=("Helvetica", 12, "bold"),
            padx=20,
            pady=15
        )
        config_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Directorio de salida
        dir_frame = tk.Frame(config_frame)
        dir_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(dir_frame, text="Directorio de salida:", font=("Helvetica", 10)).pack(anchor=tk.W)
        
        dir_input_frame = tk.Frame(dir_frame)
        dir_input_frame.pack(fill=tk.X, pady=5)
        
        dir_entry = tk.Entry(
            dir_input_frame,
            textvariable=self.directorio_salida,
            font=("Helvetica", 10),
            state="readonly"
        )
        dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        dir_button = tk.Button(
            dir_input_frame,
            text="📁 Seleccionar",
            command=self._seleccionar_directorio,
            bg="#3498db",
            fg="white",
            font=("Helvetica", 9, "bold"),
            cursor="hand2",
            relief=tk.FLAT,
            padx=10
        )
        dir_button.pack(side=tk.LEFT, padx=(10, 0))
        
        # Cantidad de contactos
        cantidad_frame = tk.Frame(config_frame)
        cantidad_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(
            cantidad_frame,
            text="Cantidad de contactos a extraer:",
            font=("Helvetica", 10)
        ).pack(anchor=tk.W)
        
        cantidad_spinbox = tk.Spinbox(
            cantidad_frame,
            from_=1,
            to=10000,
            textvariable=self.cantidad_contactos,
            font=("Helvetica", 10),
            width=20
        )
        cantidad_spinbox.pack(anchor=tk.W, pady=5)
        
        # ===== CONTROL =====
        control_frame = tk.LabelFrame(
            self.root,
            text="🎮 Control",
            font=("Helvetica", 12, "bold"),
            padx=20,
            pady=15
        )
        control_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Botones de control
        button_frame = tk.Frame(control_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        self.btn_iniciar = tk.Button(
            button_frame,
            text="▶️ Iniciar Extracción",
            command=self._iniciar_scraper,
            bg="#27ae60",
            fg="white",
            font=("Helvetica", 11, "bold"),
            cursor="hand2",
            relief=tk.FLAT,
            padx=20,
            pady=10
        )
        self.btn_iniciar.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        self.btn_detener = tk.Button(
            button_frame,
            text="⏹️ Detener",
            command=self._detener_scraper,
            bg="#e74c3c",
            fg="white",
            font=("Helvetica", 11, "bold"),
            cursor="hand2",
            relief=tk.FLAT,
            padx=20,
            pady=10,
            state=tk.DISABLED
        )
        self.btn_detener.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # ===== PROGRESO =====
        progress_frame = tk.LabelFrame(
            self.root,
            text="📊 Progreso",
            font=("Helvetica", 12, "bold"),
            padx=20,
            pady=15
        )
        progress_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Contador
        contador_frame = tk.Frame(progress_frame)
        contador_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(
            contador_frame,
            text="Contactos extraídos:",
            font=("Helvetica", 10)
        ).pack(side=tk.LEFT)
        
        self.label_contador = tk.Label(
            contador_frame,
            textvariable=self.contactos_extraidos,
            font=("Helvetica", 16, "bold"),
            fg="#27ae60"
        )
        self.label_contador.pack(side=tk.LEFT, padx=10)
        
        self.label_total = tk.Label(
            contador_frame,
            text="/ 0",
            font=("Helvetica", 12),
            fg="#7f8c8d"
        )
        self.label_total.pack(side=tk.LEFT)
        
        # Barra de progreso
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            mode='determinate',
            length=300
        )
        self.progress_bar.pack(fill=tk.X, pady=10)
        
        # Estado
        self.label_estado = tk.Label(
            progress_frame,
            text="⚪ Esperando...",
            font=("Helvetica", 10, "italic"),
            fg="#7f8c8d"
        )
        self.label_estado.pack(anchor=tk.W)
        
        # ===== LOG =====
        log_frame = tk.LabelFrame(
            self.root,
            text="📝 Log",
            font=("Helvetica", 12, "bold"),
            padx=10,
            pady=10
        )
        log_frame.pack(fill=tk.BOTH, padx=20, pady=10, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=8,
            font=("Courier", 9),
            bg="#ecf0f1",
            fg="#2c3e50",
            wrap=tk.WORD
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Log inicial
        self._agregar_log("✅ Interfaz iniciada correctamente")
        self._agregar_log(f"📁 Directorio de salida: {self.directorio_salida.get()}")
    
    def _seleccionar_directorio(self):
        """Abre un diálogo para seleccionar el directorio de salida"""
        directorio = filedialog.askdirectory(
            title="Seleccionar directorio de salida",
            initialdir=self.directorio_salida.get()
        )
        if directorio:
            self.directorio_salida.set(directorio)
            self._agregar_log(f"📁 Directorio cambiado a: {directorio}")
    
    def _agregar_log(self, mensaje):
        """Agrega un mensaje al log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {mensaje}\n")
        self.log_text.see(tk.END)
        self.log_text.update()
    
    def _actualizar_estado(self, mensaje, color="#7f8c8d"):
        """Actualiza el label de estado"""
        self.label_estado.config(text=mensaje, fg=color)
        self.root.update()
    
    def _actualizar_contador(self, valor):
        """Actualiza el contador de contactos extraídos"""
        self.contactos_extraidos.set(valor)
        total = self.cantidad_contactos.get()
        self.label_total.config(text=f"/ {total}")
        
        # Actualizar barra de progreso
        if total > 0:
            progreso = (valor / total) * 100
            self.progress_bar['value'] = progreso
        
        self.root.update()
    
    def _iniciar_scraper(self):
        """Inicia el scraper en un thread separado"""
        if self.scraper_activo:
            self._agregar_log("⚠️ El scraper ya está activo")
            return
        
        # Validar cantidad
        if self.cantidad_contactos.get() <= 0:
            self._agregar_log("❌ La cantidad debe ser mayor a 0")
            return
        
        # Resetear variables
        self.detener_flag = False
        self.contactos_extraidos.set(0)
        self._actualizar_contador(0)
        
        # Actualizar UI
        self.btn_iniciar.config(state=tk.DISABLED)
        self.btn_detener.config(state=tk.NORMAL)
        self.scraper_activo = True
        
        self._agregar_log("🚀 Iniciando scraper...")
        self._actualizar_estado("🔄 Ejecutando...", "#3498db")
        
        # Ejecutar en thread separado
        self.scraper_thread = threading.Thread(target=self._ejecutar_scraper_thread, daemon=True)
        self.scraper_thread.start()
    
    def _detener_scraper(self):
        """Detiene el scraper"""
        if not self.scraper_activo:
            return
        
        self._agregar_log("⚠️ Deteniendo scraper...")
        self._actualizar_estado("⏸️ Deteniendo...", "#e67e22")
        self.detener_flag = True
        self.btn_detener.config(state=tk.DISABLED)
    
    def _ejecutar_scraper_thread(self):
        """Ejecuta el scraper en un thread separado"""
        try:
            # Crear un nuevo loop para este thread
            self.loop = asyncio.new_event_loop()
            asyncio.set_event_loop(self.loop)
            
            # Ejecutar el scraper
            self.loop.run_until_complete(self._ejecutar_scraper_async())
            
        except Exception as e:
            self._agregar_log(f"❌ Error: {str(e)}")
            logger.error(f"Error en scraper thread: {e}", exc_info=True)
        finally:
            # Limpiar
            if self.loop:
                self.loop.close()
            
            # Actualizar UI
            self.scraper_activo = False
            self.btn_iniciar.config(state=tk.NORMAL)
            self.btn_detener.config(state=tk.DISABLED)
            
            if self.detener_flag:
                self._actualizar_estado("⏹️ Detenido", "#e67e22")
                self._agregar_log("⏹️ Scraper detenido por el usuario")
            else:
                self._actualizar_estado("✅ Completado", "#27ae60")
                self._agregar_log("✅ Extracción completada")
    
    async def _ejecutar_scraper_async(self):
        """Función asíncrona que ejecuta el scraper"""
        self.scraper = DebugScraper(headless=False)
        
        try:
            # 1. Iniciar sesión
            self._agregar_log("🔑 Iniciando sesión...")
            await self.scraper.iniciar_sesion("https://correoweb.madrid.org/owa/#path=/people")
            
            page = self.scraper.page
            await page.wait_for_load_state("load", timeout=120000)
            
            # 2. Navegar al directorio
            self._agregar_log("📂 Navegando al directorio...")
            directorio = page.get_by_label("Directorio", exact=True).locator("div").filter(has_text="Directorio").nth(1)
            await directorio.click(timeout=120000)
            
            # Guardar sesión inmediatamente después de login exitoso
            await self.scraper.guardar_sesion()
            self._agregar_log("💾 Sesión guardada correctamente")
            
            # 3. Extraer contactos con callback para actualizar contador
            self._agregar_log(f"🔍 Extrayendo {self.cantidad_contactos.get()} contactos...")
            
            contactos = await self._scrape_con_actualizacion(
                page,
                self.cantidad_contactos.get()
            )
            
            # 4. Guardar sesión
            if not self.detener_flag:
                await self.scraper.guardar_sesion()
                self._agregar_log(f"💾 {len(contactos)} contactos guardados en {self.directorio_salida.get()}")
            
        except Exception as e:
            self._agregar_log(f"❌ Error: {str(e)}")
            logger.error(f"Error ejecutando scraper: {e}", exc_info=True)
            raise
        finally:
            await self.scraper.cerrar()
    
    async def _scrape_con_actualizacion(self, page, max_contacts):
        """
        Versión modificada de scrape_outlook_contacts que actualiza el contador en tiempo real
        """
        import pandas as pd
        import signal
        import json
        
        row_selector = 'div[role="heading"]'
        
        # Buscar Excel más reciente en el directorio configurado
        data_dir = Path(self.directorio_salida.get())
        data_dir.mkdir(exist_ok=True)
        
        metadata_file = data_dir / "scraping_metadata.json"
        
        contactos_previos = []
        ultimo_nombre = None
        excel_path = None
        scroll_count_guardado = 0
        
        # Buscar archivo Excel más reciente
        excel_files = list(data_dir.glob("contactos_organos_judiciales_*.xlsx"))
        if excel_files:
            excel_path = max(excel_files, key=lambda p: p.stat().st_mtime)
            self._agregar_log(f"📂 Excel encontrado: {excel_path.name}")
            
            # Cargar metadata
            if metadata_file.exists():
                try:
                    with open(metadata_file, 'r', encoding='utf-8') as f:
                        metadata = json.load(f)
                    scroll_count_guardado = metadata.get('scroll_count', 0)
                    self._agregar_log(f"📜 Scrolls previos: {scroll_count_guardado}")
                except Exception as e:
                    self._agregar_log(f"⚠️ Error leyendo metadata: {e}")
            
            try:
                df_previo = pd.read_excel(excel_path)
                contactos_previos = df_previo.to_dict('records')
                self._agregar_log(f"✅ {len(contactos_previos)} contactos previos cargados")
                
                # Actualizar contador con contactos previos
                self._actualizar_contador(len(contactos_previos))
                
                if len(contactos_previos) > 0:
                    ultimo_nombre = contactos_previos[-1]['nombre']
            except Exception as e:
                self._agregar_log(f"⚠️ Error leyendo Excel: {e}")
                contactos_previos = []
        
        contactos_faltantes = max(0, max_contacts - len(contactos_previos))
        if contactos_faltantes == 0:
            self._agregar_log(f"✅ Ya se alcanzó el límite de {max_contacts} contactos")
            return contactos_previos
        
        self._agregar_log(f"🎯 Extrayendo {contactos_faltantes} contactos adicionales...")
        
        # Esperar a que cargue la lista
        try:
            await page.wait_for_selector(row_selector, state="visible", timeout=30000)
            self._agregar_log("✅ Lista detectada")
        except Exception as e:
            self._agregar_log(f"❌ Error: Lista no cargó - {e}")
            return contactos_previos
        
        # Importar función de espera
        from debug_scraper import esperar_carga_completa, extraer_detalles_contacto, guardar_excel, guardar_metadata
        
        # Si hay scrolls guardados, ir a esa posición
        scroll_count_actual = 0
        
        if scroll_count_guardado > 0:
            self._agregar_log(f"⚡ Saltando a scroll #{scroll_count_guardado}...")
            for i in range(scroll_count_guardado):
                if self.detener_flag:
                    return contactos_previos
                
                await page.keyboard.press("PageDown")
                await asyncio.sleep(0.3)
                
                if i % 10 == 0 and i > 0:
                    self._agregar_log(f"   Scrolleando... {i}/{scroll_count_guardado}")
            
            scroll_count_actual = scroll_count_guardado
            await asyncio.sleep(2)
        
        # Variables de scraping
        contactos_nuevos = []
        contactos_procesados = set(c['nombre'] for c in contactos_previos)
        last_item_count = 0
        retries = 0
        max_retries = 15
        
        # Loop principal de extracción
        while not self.detener_flag:
            filas = await page.locator(row_selector).all()
            
            if not filas:
                break
            
            for fila in filas:
                if self.detener_flag:
                    break
                
                try:
                    raw_name = await fila.get_attribute("aria-label")
                    nombre = raw_name.strip() if raw_name else "Desconocido"
                    
                    # Filtro: mayúsculas + coma
                    if "," not in nombre or (not nombre.isupper()):
                        continue
                    
                    # Evitar duplicados
                    if nombre in contactos_procesados:
                        continue
                    
                    contactos_procesados.add(nombre)
                    
                    total_actual = len(contactos_previos) + len(contactos_nuevos)
                    
                    # Extraer detalles
                    detalles = await extraer_detalles_contacto(page, fila, nombre)
                    
                    # Filtro: Solo ORGANOS JUDICIALES
                    if detalles['compania'] and 'ORGANOS JUDICIALES' in detalles['compania'].upper():
                        contactos_nuevos.append(detalles)
                        
                        # ACTUALIZAR CONTADOR EN TIEMPO REAL
                        nuevo_total = len(contactos_previos) + len(contactos_nuevos)
                        self._actualizar_contador(nuevo_total)
                        self._agregar_log(f"✅ ({nuevo_total}/{max_contacts}) - {nombre[:40]}")
                    else:
                        contactos_procesados.remove(nombre)
                        continue
                    
                    # Verificar límite
                    if len(contactos_previos) + len(contactos_nuevos) >= max_contacts:
                        self._agregar_log(f"🎯 Límite alcanzado: {max_contacts} contactos")
                        break
                
                except Exception as e:
                    self._agregar_log(f"❌ Error procesando: {e}")
            
            # Salir si alcanzamos el límite o se detuvo
            if len(contactos_previos) + len(contactos_nuevos) >= max_contacts or self.detener_flag:
                break
            
            # Scroll
            try:
                if filas:
                    await filas[-1].scroll_into_view_if_needed()
                    await page.keyboard.press("PageDown")
                    scroll_count_actual += 1
                    await asyncio.sleep(0.5)
                    await page.keyboard.press("PageDown")
                    scroll_count_actual += 1
            except Exception as e:
                pass
            
            await esperar_carga_completa(page)
            
            # Verificación de fin
            current_count = len(contactos_nuevos)
            if current_count == last_item_count:
                retries += 1
                
                await page.keyboard.press("End")
                await asyncio.sleep(0.5)
                await page.keyboard.press("PageDown")
                scroll_count_actual += 1
                await esperar_carga_completa(page)
                
                if retries >= max_retries:
                    self._agregar_log(f"⚠️ Terminado después de {retries} intentos")
                    break
            else:
                retries = 0
            
            last_item_count = current_count
        
        # Guardar resultado final
        todos_contactos = contactos_previos + contactos_nuevos
        if todos_contactos and not self.detener_flag:
            guardar_excel(todos_contactos, data_dir, excel_path)
            guardar_metadata(metadata_file, todos_contactos, scroll_count_actual)
        
        return todos_contactos


def main():
    """Función principal para iniciar la GUI"""
    root = tk.Tk()
    app = ScraperGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
