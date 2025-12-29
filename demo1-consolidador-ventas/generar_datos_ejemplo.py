"""
Generador de Datos de Ejemplo para Demo 1
==========================================
Script que genera archivos Excel de ejemplo con datos realistas
de ventas de 3 sucursales para demostrar el consolidador.

Autor: Excel Automatizado
Fecha: Diciembre 2024
"""

import sys
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime, timedelta
import random

# Configurar encoding para Windows
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')

# ═══════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ═══════════════════════════════════════════════════════════════

# Rutas
BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "input"

# Crear carpeta si no existe
INPUT_DIR.mkdir(exist_ok=True)

# ═══════════════════════════════════════════════════════════════
# DATOS MAESTROS
# ═══════════════════════════════════════════════════════════════

SUCURSALES = {
    'Centro': {
        'vendedores': ['Juan Pérez', 'Ana López', 'Carlos Méndez', 'María García'],
        'filas': 150
    },
    'Norte': {
        'vendedores': ['Pedro Ramírez', 'Lucía Torres', 'Roberto Silva'],
        'filas': 120
    },
    'Sur': {
        'vendedores': ['Carmen Ruiz', 'Diego Morales', 'Patricia Herrera'],
        'filas': 130
    }
}

PRODUCTOS_CATEGORIAS = {
    'Electrónica': [
        ('Laptop Dell XPS', 850.00, 1200.00),
        ('Laptop HP Pavilion', 650.00, 900.00),
        ('Monitor Samsung 27"', 250.00, 400.00),
        ('Monitor LG 24"', 180.00, 300.00),
        ('Tablet Samsung', 300.00, 450.00),
        ('iPad Air', 550.00, 750.00),
        ('Impresora HP LaserJet', 200.00, 350.00),
        ('Impresora Epson Multifuncional', 150.00, 250.00),
        ('Disco Duro Externo 1TB', 50.00, 80.00),
        ('SSD 500GB', 60.00, 100.00),
    ],
    'Accesorios': [
        ('Mouse Logitech', 15.00, 35.00),
        ('Teclado Mecánico', 50.00, 100.00),
        ('Teclado Inalámbrico', 25.00, 45.00),
        ('Webcam HD', 40.00, 70.00),
        ('Audífonos Bluetooth', 35.00, 80.00),
        ('Audífonos Gaming', 60.00, 120.00),
        ('Cable HDMI 2m', 8.00, 15.00),
        ('Cable USB-C', 10.00, 20.00),
        ('Hub USB 4 puertos', 15.00, 30.00),
        ('Mousepad Gaming', 12.00, 25.00),
    ],
    'Software': [
        ('Licencia Office 365', 70.00, 100.00),
        ('Antivirus Norton', 40.00, 60.00),
        ('Windows 11 Pro', 150.00, 200.00),
        ('Adobe Creative Cloud', 50.00, 80.00),
        ('AutoCAD Licencia', 200.00, 300.00),
    ],
    'Componentes': [
        ('Memoria RAM 8GB', 35.00, 60.00),
        ('Memoria RAM 16GB', 70.00, 110.00),
        ('Procesador Intel i5', 180.00, 250.00),
        ('Procesador AMD Ryzen 5', 170.00, 240.00),
        ('Tarjeta Gráfica GTX', 250.00, 400.00),
        ('Placa Madre ASUS', 120.00, 180.00),
        ('Fuente de Poder 600W', 60.00, 90.00),
    ]
}

# ═══════════════════════════════════════════════════════════════
# FUNCIONES AUXILIARES
# ═══════════════════════════════════════════════════════════════

def generar_fecha_random(dias_atras: int = 30) -> datetime:
    """Genera una fecha aleatoria en los últimos N días"""
    hoy = datetime.now()
    fecha_inicio = hoy - timedelta(days=dias_atras)
    dias_random = random.randint(0, dias_atras)
    return fecha_inicio + timedelta(days=dias_random)


def generar_cantidad_random(producto: str) -> int:
    """
    Genera cantidad aleatoria basada en el tipo de producto.
    Productos caros: 1-3 unidades
    Productos baratos: 1-10 unidades
    """
    if any(keyword in producto for keyword in ['Laptop', 'Monitor', 'Tablet', 'iPad', 'Procesador']):
        return random.randint(1, 3)
    elif any(keyword in producto for keyword in ['Cable', 'Mouse', 'Mousepad']):
        return random.randint(1, 10)
    else:
        return random.randint(1, 5)


def seleccionar_precio_random(precio_min: float, precio_max: float) -> float:
    """Selecciona un precio aleatorio dentro del rango con distribución realista"""
    # Usar distribución normal centrada en el punto medio
    precio_medio = (precio_min + precio_max) / 2
    desviacion = (precio_max - precio_min) / 4
    precio = np.random.normal(precio_medio, desviacion)

    # Asegurar que esté dentro del rango
    precio = max(precio_min, min(precio_max, precio))

    # Redondear a 2 decimales
    return round(precio, 2)


# ═══════════════════════════════════════════════════════════════
# GENERACIÓN DE DATOS
# ═══════════════════════════════════════════════════════════════

def generar_ventas_sucursal(nombre_sucursal: str, config: dict) -> pd.DataFrame:
    """
    Genera un DataFrame con ventas aleatorias para una sucursal.

    Args:
        nombre_sucursal: Nombre de la sucursal
        config: Diccionario con configuración (vendedores, cantidad de filas)

    Returns:
        pd.DataFrame con las ventas generadas
    """
    vendedores = config['vendedores']
    num_filas = config['filas']

    datos = []

    for _ in range(num_filas):
        # Seleccionar categoría aleatoria
        categoria = random.choice(list(PRODUCTOS_CATEGORIAS.keys()))

        # Seleccionar producto aleatorio de esa categoría
        productos = PRODUCTOS_CATEGORIAS[categoria]
        producto, precio_min, precio_max = random.choice(productos)

        # Generar datos de la venta
        fila = {
            'Fecha': generar_fecha_random(30),
            'Producto': producto,
            'Categoría': categoria,
            'Cantidad': generar_cantidad_random(producto),
            'Precio_Unitario': seleccionar_precio_random(precio_min, precio_max),
            'Vendedor': random.choice(vendedores),
            'Sucursal': nombre_sucursal
        }

        datos.append(fila)

    # Crear DataFrame
    df = pd.DataFrame(datos)

    # Ordenar por fecha
    df = df.sort_values('Fecha').reset_index(drop=True)

    # Formatear fecha (solo fecha, sin hora)
    df['Fecha'] = df['Fecha'].dt.date

    return df


# ═══════════════════════════════════════════════════════════════
# SCRIPT PRINCIPAL
# ═══════════════════════════════════════════════════════════════

def main():
    """Genera los 3 archivos Excel de ejemplo"""
    print("=" * 60)
    print("GENERADOR DE DATOS DE EJEMPLO - DEMO 1")
    print("=" * 60)
    print()

    # Establecer semilla para reproducibilidad (opcional)
    random.seed(42)
    np.random.seed(42)

    # Generar archivo para cada sucursal
    for sucursal, config in SUCURSALES.items():
        print(f"Generando datos para: {sucursal}")

        # Generar DataFrame
        df = generar_ventas_sucursal(sucursal, config)

        # Nombre del archivo
        archivo = INPUT_DIR / f"ventas_sucursal_{sucursal.lower()}.xlsx"

        # Guardar Excel
        df.to_excel(archivo, index=False, sheet_name='Ventas')

        # Estadísticas
        total_ventas = (df['Cantidad'] * df['Precio_Unitario']).sum()
        print(f"  ✓ {len(df)} registros generados")
        print(f"  ✓ Total ventas: ${total_ventas:,.2f}")
        print(f"  ✓ Archivo: {archivo.name}")
        print()

    print("=" * 60)
    print("✓ ¡Archivos de ejemplo generados exitosamente!")
    print(f"📁 Ubicación: {INPUT_DIR}")
    print("=" * 60)


if __name__ == "__main__":
    main()
