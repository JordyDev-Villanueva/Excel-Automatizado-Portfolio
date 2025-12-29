"""
Shared Utilities - Módulo de Utilidades Compartidas
====================================================
Paquete con funciones reutilizables para automatización de Excel.
"""

from .excel_helper import (
    setup_logger,
    leer_archivos_excel,
    crear_grafico_barras,
    crear_grafico_circular,
    crear_grafico_linea,
    aplicar_formato_profesional,
    insertar_imagen_en_excel,
    crear_tabla_excel,
    formatear_moneda_columna,
    formatear_porcentaje_columna,
    formatear_numero_columna,
    COLORES
)

__all__ = [
    'setup_logger',
    'leer_archivos_excel',
    'crear_grafico_barras',
    'crear_grafico_circular',
    'crear_grafico_linea',
    'aplicar_formato_profesional',
    'insertar_imagen_en_excel',
    'crear_tabla_excel',
    'formatear_moneda_columna',
    'formatear_porcentaje_columna',
    'formatear_numero_columna',
    'COLORES'
]
