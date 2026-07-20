# Extracción GAL con Excel + Enriquecimiento por Compañía

**Fecha:** 2026-07-19
**Estado:** Aprobado

## Problema

El scraper actual:
- Exporta a CSV (no Excel)
- No permite seleccionar qué compañías enriquecer
- No guarda caché local para enrichment offline
- No tiene mecanismo de reanudación

## Solución

Sistema de 2 hojas en Excel + enrichment selectivo por compañía.

## Arquitectura de Excel

### Sheet 1 — "Contactos"
| Col | Campo | Fuente |
|-----|-------|--------|
| A | Nombre | FindPeople |
| B | Email | FindPeople |
| C | Empresa | FindPeople (CompanyName) |
| D | Teléfono | GetPersona (si enrich) |
| E | Departamento | GetPersona (si enrich) |
| F | Oficina | GetPersona (si enrich) |
| G | Dirección | GetPersona (si enrich) |

### Sheet 2 — "Compañías"
| Col | Campo |
|-----|-------|
| A | Compañía |
| B | Enrich (X para seleccionar) |

## Extracción (GUI)

### Botón: "▶️ Extraer GAL"
1. Descarga GAL completo (sin enrichment para ser rápido)
2. Por cada contacto → extraer `CompanyName` → agregar a Sheet 2 si es nuevo
3. Guardar en Excel: `data/gal/gal_export.xlsx`
4. Guardar cache JSON: `data/gal/cache.json`

### Botón: "🔄 Enriquecer"
1. Leer Sheet 2 → filtrar filas con "X" en columna Enrich
2. Para cada compañía marcada:
   - Buscar en Sheet 1 contactos donde Empresa == Compañía
   - Para cada match → lookup en cache.json → solo llenar campos vacíos
3. Guardar progreso temporal: `data/gal/enrich_progress.json` cada 50 contactos
4. Actualizar Excel al completar
5. Si falla → puede reanudar desde progress

### Match Exacto
- CompanyName debe ser **exactamente igual** para hacer match
- Las compañías se extraen dinámicamente del GAL durante extracción

## Archivos

| Archivo | Propósito |
|---------|-----------|
| `data/gal/gal_export.xlsx` | Excel con 2 hojas |
| `data/gal/cache.json` | GAL completo en JSON (para enrichment offline) |
| `data/gal/enrich_progress.json` | Estado de enrichment (offset + companies completadas) |

## Campos de Contacto

```python
CONTACT_FIELDS = [
    'nombre',      # DisplayName
    'email',      # Email1
    'empresa',    # CompanyName
    'telefono',   # BusinessPhone
    'departamento', # Department
    'oficina',    # OfficeLocation
    'direccion',  # BusinessAddress
]
```

## Flujo de Enrichment

```
1. Leer Sheet2.filter(Enrich == "X")
2. Para cada compañía:
   a. Leer Sheet1.rows donde Empresa == compañía
   b. Para cada fila:
      - Buscar en cache.json por email/CompanyName
      - Solo llenar campos vacíos (no sobreescribir)
      - Si N contactos procesados -> guardar progress
   c. Marcar compañía como completada
3. Guardar Excel actualizado
```

## UI Cambios

- Checkbox "Enriquecer contactos" → **Botón** "🔄 Enriquecer"
- Sheet 2 restaurado con columnas [Compañía, Enrich]
- Nuevo frame para mostrar estado de enrichment
