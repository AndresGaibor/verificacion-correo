# Interfaz Gráfica para Verificación de Correos OWA

Esta guía explica cómo usar la nueva interfaz gráfica (GUI) para la verificación de correos OWA.

## 🚀 Inicio Rápido

### Requisitos Previos
1. **Sesión activa**: Ejecuta `python copiar_sesion.py` primero para crear `state.json`
2. **Archivo Excel**: Asegúrate de tener `data/correos.xlsx` con los correos a verificar
3. **Configuración**: Revisa que `config.yaml` tenga la URL correcta de OWA

### Ejecutar la GUI
```bash
python gui.py
```

## 📋 Componentes de la Interfaz

### Pestaña "Procesamiento"

#### Botones de Control
- **🚀 Iniciar Verificación**: Comienza el proceso de verificación de correos
- **⏸ Pausar**: Pausa temporalmente el procesamiento
- **🛑 Detener**: Detiene completamente el proceso

#### Barra de Progreso
Muestra el progreso actual del procesamiento con:
- Barra visual con porcentaje
- Mensaje de estado detallado
- Estadísticas en tiempo real

#### Estadísticas
- **✅ OK**: Correos con información encontrada
- **❌ No Existe**: Correos que no existen en el sistema
- **⚠️ Error**: Correos con errores durante el procesamiento

#### Log de Actividad
Área de texto que muestra:
- Timestamp de cada acción
- Detalles del procesamiento
- Errores y advertencias
- Resumen final

### Pestaña "Configuración"

#### Configuración OWA
- **URL de OWA**: Dirección del servidor de correo web
  - Formato: `https://servidor.empresa.com/owa/#path=/mail`

#### Configuración de Procesamiento
- **Tamaño de Lote**: Número de correos procesados por lote (1-50)
  - Valor recomendado: 10
  - Lotes más grandes pueden ser más rápidos pero usan más memoria

#### Configuración del Navegador
- **Modo Headless**: Oculta la ventana del navegador
  - Desmarcado: Muestra ventana (recomendado para debugging)
  - Marcado: Ejecución sin interfaz visual

#### Configuración de Archivos
- **Archivo Excel**: Ruta al archivo con correos a procesar
- **Archivo Sesión**: Ruta al archivo de sesión (generalmente `state.json`)

## 🔄 Flujo de Trabajo

### 1. Configuración Inicial
1. Abre la aplicación: `python gui.py`
2. Ve a la pestaña "Configuración"
3. Verifica que la URL de OWA sea correcta
4. Ajusta el tamaño de lote si es necesario
5. Confirma las rutas de archivos
6. Click en "💾 Guardar Configuración"

### 2. Procesamiento
1. Ve a la pestaña "Procesamiento"
2. Click en "🚀 Iniciar Verificación"
3. Observa el progreso en tiempo real:
   - Barra de progreso con porcentaje
   - Log detallado de actividades
   - Estadísticas actualizadas
4. El proceso se detendrá automáticamente al finalizar

### 3. Resultados
- Al finalizar, verás un resumen con estadísticas finales
- Los resultados se guardan automáticamente en el archivo Excel
- Cada correo procesado tendrá su estado en la columna "Status"

## ⚠️ Solución de Problemas

### Errores Comunes

#### "No se encontró el archivo de sesión"
- **Solución**: Ejecuta `python copiar_sesion.py` primero
- **Causa**: El archivo `state.json` no existe o está dañado

#### "No se encontró el archivo Excel"
- **Solución**: Crea `data/correos.xlsx` con correos en columna A
- **Causa**: El archivo especificado no existe

#### "La URL no está configurada"
- **Solución**: Configura la URL correcta en la pestaña "Configuración"
- **Causa**: El campo URL está vacío o tiene formato incorrecto

### Consejos de Uso

#### Para Mejorar Rendimiento
- Usa lotes de 10-20 correos
- Mantén el navegador visible (headless=False) para debugging
- Cierra otras aplicaciones que usen muchos recursos

#### Para Debugging
- Desactiva el modo headless para ver lo que hace el navegador
- Observa el log de actividad para identificar errores específicos
- Usa lotes pequeños (1-5) para aislar problemas

#### Para Producción
- Activa el modo headless para ejecución automática
- Usa lotes más grandes (10-20) para mayor eficiencia
- Programa ejecuciones periódicas si es necesario

## 📊 Formato de Archivos

### Archivo Excel (correos.xlsx)
```
| A                 | B        | C      | D             | ... |
|-------------------|----------|--------|---------------|-----|
| Correo            | Status   | Nombre | Email Personal | ... |
| user1@empresa.com |          |        |               | ... |
| user2@empresa.com |          |        |               | ... |
```

- **Columna A**: Correos electrónicos a verificar (obligatorio)
- **Columna B**: Status (se llena automáticamente)
- **Columnas C-J**: Datos extraídos (se llenan automáticamente)

### Archivo de Configuración (config.yaml)
```yaml
page_url: "https://tu-servidor-owa.com/owa/#path=/mail"
processing:
  batch_size: 10
browser:
  headless: false
  session_file: "state.json"
excel:
  default_file: "data/correos.xlsx"
```

## 🛠️ Comandos Útiles

### Ejecutar GUI
```bash
python gui.py
```

### Crear Sesión (antes de usar GUI)
```bash
python copiar_sesion.py
```

### Probar Funcionamiento
```bash
python test_gui.py
```

### Ver Ayuda
```bash
python gui.py --help  # (si está implementado)
```

## 📝 Notas Técnicas

- **Threading**: La GUI usa threading para no bloquearse durante el procesamiento
- **Configuración**: Los cambios se guardan inmediatamente en `config.yaml`
- **Logs**: Se muestran en tiempo real en la interfaz
- **Errores**: Los errores se muestran tanto en el log como en ventanas emergentes
- **Progreso**: Se actualiza dinámicamente durante el procesamiento

## 🔄 Compatibilidad

La interfaz gráfica es **100% compatible** con los scripts existentes:
- `app.py`: Script de línea de comandos
- `browser_automation.py`: Automatización del navegador
- `excel_reader.py` y `excel_writer.py`: Manejo de archivos Excel
- `contact_extractor.py`: Extracción de información

Puedes usar tanto la GUI como los scripts de línea de comandos indistintamente.