#!/usr/bin/env python3
"""
Script principal para verificación de correos en OWA.
Lee emails desde Excel, procesa solo los pendientes en lotes,
y guarda resultados incrementalmente.
"""

from excel_reader import leer_correos_pendientes
from browser_automation import procesar_todos_los_lotes
from config import EXCEL_CONFIG


def main():
    """
    Función principal que orquesta el proceso de verificación incremental.
    """
    archivo_excel = EXCEL_CONFIG['default_file']

    print("="*70)
    print("VERIFICACIÓN DE CORREOS - SISTEMA INCREMENTAL")
    print("="*70)

    # 1. Leer emails pendientes desde Excel
    print(f"\nLeyendo correos desde {archivo_excel}...")
    lotes_info = leer_correos_pendientes(archivo_excel)

    # 2. Mostrar resumen
    print(f"\nTotal emails en archivo: {lotes_info['total_emails']}")
    print(f"Ya procesados: {lotes_info['total_procesados']}")
    print(f"Pendientes: {lotes_info['total_pendientes']}")

    if lotes_info['total_pendientes'] == 0:
        print("\n✓ No hay emails pendientes para procesar.")
        print("  Para re-procesar un email, borra su columna 'Status' en el Excel.")
        return

    print(f"Dividido en {len(lotes_info['lotes'])} lotes de 10 emails")

    # 3. Confirmar antes de procesar
    try:
        input("\nPresiona ENTER para iniciar el procesamiento (o Ctrl+C para cancelar)...")
    except KeyboardInterrupt:
        print("\n\nProcesamiento cancelado por el usuario.")
        return

    # 4. Procesar todos los lotes
    print("\nIniciando procesamiento...")
    stats = procesar_todos_los_lotes(lotes_info, archivo_excel)

    # 5. Mostrar resumen final
    print("\n" + "="*70)
    print("RESUMEN FINAL")
    print("="*70)
    print(f"✓ {stats['ok']} contactos encontrados y guardados")
    print(f"✗ {stats['no_existe']} no existen")
    print(f"⚠ {stats['error']} errores")
    print(f"\nTotal procesado: {stats['ok'] + stats['no_existe'] + stats['error']}")
    print(f"\nResultados guardados en: {archivo_excel}")
    print("="*70)

    return stats


if __name__ == "__main__":
    main()
