# 📊 Consolidador de Ventas - Demo 1

> **Automatización Profesional de Excel con Python**
>
> Script que consolida múltiples archivos Excel de ventas en un reporte ejecutivo único con análisis avanzados y visualizaciones de alta calidad.

---

## 🎯 Problema que Resuelve

**Situación común en empresas:**
- Múltiples sucursales envían sus reportes de ventas en Excel separados
- El gerente debe consolidar manualmente toda la información
- Proceso tedioso que toma 3-4 horas cada mes
- Alto riesgo de errores humanos en cálculos y copias

**Solución automatizada:**
- ✅ Consolida automáticamente todos los archivos en **30 segundos**
- ✅ Genera análisis profesionales con cero errores
- ✅ Crea gráficos de alta calidad listos para presentaciones
- ✅ Formato corporativo profesional en el output
- ✅ Ahorro de **98% del tiempo** invertido

---

## 📋 Características

### Procesamiento de Datos
- Lectura automática de múltiples archivos Excel (.xlsx)
- Validación de estructura de datos
- Consolidación inteligente con verificación de duplicados
- Cálculos automáticos de métricas clave

### Análisis Incluidos
1. **Ventas por Sucursal** - Total y participación porcentual
2. **Top 10 Productos** - Por cantidad vendida y por monto
3. **Desempeño de Vendedores** - Ventas totales, transacciones, ticket promedio
4. **Análisis por Categoría** - Distribución de ventas
5. **Tendencia Temporal** - Evolución diaria de ventas

### Visualizaciones
- 📊 Gráfico de barras: Ventas por sucursal
- 🥧 Gráfico circular: Distribución por categoría
- 📈 Gráfico de línea: Tendencia temporal

### Output Excel Profesional
Archivo multi-hoja con:
- **Dashboard**: KPIs principales + gráficos insertados
- **Datos Consolidados**: Todos los registros en formato tabla
- **Top Productos**: Rankings de los más vendidos
- **Análisis Vendedores**: Métricas de desempeño
- **Resumen Sucursales**: Comparativa entre ubicaciones

---

## 🛠️ Tecnologías

- **Python 3.8+**
- **pandas** - Procesamiento y análisis de datos
- **openpyxl** - Manipulación avanzada de Excel
- **matplotlib** - Generación de gráficos
- **seaborn** - Visualizaciones profesionales

---

## 📦 Instalación

### 1. Clonar o descargar este proyecto

```bash
cd demo1-consolidador-ventas
```

### 2. Instalar dependencias

```bash
pip install -r requirements.txt
```

**requirements.txt incluye:**
```
pandas==2.1.4
openpyxl==3.1.2
matplotlib==3.8.2
seaborn==0.13.0
numpy==1.26.2
```

---

## 🚀 Uso

### Paso 1: Preparar archivos de entrada

Coloca tus archivos Excel de ventas en la carpeta `input/`

**Estructura requerida de cada Excel:**

| Fecha      | Producto    | Categoría   | Cantidad | Precio_Unitario | Vendedor    | Sucursal |
|------------|-------------|-------------|----------|-----------------|-------------|----------|
| 2025-01-05 | Laptop Dell | Electrónica | 2        | 850.00          | Juan Pérez  | Centro   |
| 2025-01-05 | Mouse USB   | Accesorios  | 5        | 25.00           | Ana López   | Centro   |

**Columnas obligatorias:**
- `Fecha` - Fecha de la venta
- `Producto` - Nombre del producto
- `Categoría` - Categoría del producto
- `Cantidad` - Unidades vendidas
- `Precio_Unitario` - Precio por unidad
- `Vendedor` - Nombre del vendedor
- `Sucursal` - Nombre de la sucursal

### Paso 2: Generar datos de ejemplo (opcional)

Si deseas probar el script con datos de ejemplo:

```bash
python generar_datos_ejemplo.py
```

Esto creará 3 archivos Excel de ejemplo en `input/`:
- `ventas_sucursal_centro.xlsx` (150 registros)
- `ventas_sucursal_norte.xlsx` (120 registros)
- `ventas_sucursal_sur.xlsx` (130 registros)

### Paso 3: Ejecutar el consolidador

```bash
python consolidador.py
```

### Paso 4: Revisar el resultado

El reporte consolidado se genera en: `output/reporte_consolidado.xlsx`

---

## 📁 Estructura del Proyecto

```
demo1-consolidador-ventas/
│
├── input/                          # Carpeta con archivos Excel de entrada
│   ├── ventas_sucursal_centro.xlsx
│   ├── ventas_sucursal_norte.xlsx
│   └── ventas_sucursal_sur.xlsx
│
├── output/                         # Carpeta con resultados generados
│   ├── reporte_consolidado.xlsx    # ← ARCHIVO FINAL
│   └── graficos_temp/              # Gráficos PNG temporales
│
├── consolidador.py                 # Script principal
├── generar_datos_ejemplo.py        # Generador de datos de prueba
├── requirements.txt                # Dependencias Python
└── README.md                       # Este archivo
```

---

## 📊 Ejemplo de Output

### Dashboard con KPIs

```
┌─────────────────────────────────────────────────┐
│  📊 REPORTE CONSOLIDADO DE VENTAS              │
│                                                 │
│  Total Ventas:           $113,220.50           │
│  Total Transacciones:    400                   │
│  Ticket Promedio:        $283.05               │
│  Sucursales:             3                     │
│  Vendedores:             10                    │
│  Productos Únicos:       27                    │
│                                                 │
│  [Gráfico: Ventas por Sucursal]               │
│  [Gráfico: Distribución por Categoría]        │
│  [Gráfico: Tendencia Temporal]                │
└─────────────────────────────────────────────────┘
```

### Hoja "Datos_Consolidados"
Tabla formateada con todos los registros consolidados, incluyendo columna calculada `Total_Venta`.

### Hoja "Top_Productos"
Rankings lado a lado:
- Más vendidos por cantidad
- Más rentables por monto

### Hoja "Analisis_Vendedores"
Tabla con métricas de cada vendedor:
- Total de ventas
- Número de transacciones
- Ticket promedio

### Hoja "Resumen_Sucursales"
Comparativa entre sucursales con participación porcentual.

---

## 🎨 Características de Diseño

### Formato Profesional
- ✅ Colores corporativos consistentes
- ✅ Headers con fondo azul y texto blanco
- ✅ Tablas formateadas tipo Excel nativo
- ✅ Anchos de columna ajustados automáticamente
- ✅ Bordes sutiles y alineación perfecta

### Gráficos de Alta Calidad
- ✅ Resolución 300 DPI (calidad impresión)
- ✅ Tamaño compacto uniforme (3x2.5 pulgadas)
- ✅ Layout horizontal para visualización completa
- ✅ Estilo profesional con seaborn
- ✅ Colores armoniosos
- ✅ Títulos y labels claros

### Formatos Numéricos
- 💰 Moneda: `$12,345.67`
- 📊 Números: `1,234`
- 📈 Porcentajes: `25.50%`

---

## 🔧 Personalización

### Cambiar cantidad de productos en el Top
En `consolidador.py`, línea ~120:

```python
top_cantidad, top_monto = analizar_top_productos(df_consolidado, top_n=10)  # Cambiar 10 por el número deseado
```

### Modificar colores corporativos
En `shared_utils/excel_helper.py`:

```python
COLORES = {
    'azul_oscuro': '1F4788',
    'azul_claro': 'D6E4F5',
    'verde': '70AD47',
    # ... modificar según preferencia
}
```

### Agregar nuevos análisis
Crea una función en la sección "FUNCIONES DE ANÁLISIS" de `consolidador.py`:

```python
def analizar_mi_metrica(df: pd.DataFrame) -> pd.DataFrame:
    """Tu análisis personalizado"""
    resultado = df.groupby('TuColumna').agg({'OtraColumna': 'sum'})
    return resultado
```

---

## ⚠️ Requisitos de los Archivos de Entrada

**✅ Los archivos deben:**
- Estar en formato `.xlsx` (Excel)
- Tener las 7 columnas obligatorias con nombres exactos
- Contener al menos 1 fila de datos (además del header)
- Estar ubicados en la carpeta `input/`

**❌ Errores comunes:**
- ✗ Nombres de columnas con espacios extra o acentos diferentes
- ✗ Columnas faltantes
- ✗ Archivos corruptos
- ✗ Formato `.xls` (antiguo, no compatible)

---

## 🐛 Troubleshooting

### Error: "La carpeta no existe"
**Solución:** Crear la carpeta `input/` en el mismo directorio del script.

### Error: "No se encontraron archivos"
**Solución:** Verificar que los archivos estén en `input/` y tengan extensión `.xlsx`.

### Error: "El archivo no tiene las columnas: {columnas}"
**Solución:** Verificar que los archivos tengan exactamente los nombres de columnas requeridos.

### Los gráficos no se ven en el Excel
**Solución:** Asegurarse de tener instaladas las librerías `matplotlib` y `seaborn`.

---

## 📝 Logging

El script genera logs detallados en consola:

```
2025-12-29 10:30:15 | INFO     | ============================================================
2025-12-29 10:30:15 | INFO     | CONSOLIDADOR DE VENTAS - DEMO 1
2025-12-29 10:30:15 | INFO     | ============================================================
2025-12-29 10:30:15 | INFO     |
2025-12-29 10:30:15 | INFO     | Paso 1: Leyendo archivos Excel...
2025-12-29 10:30:15 | INFO     | Encontrados 3 archivo(s) para procesar
2025-12-29 10:30:15 | INFO     |   → Leyendo: ventas_sucursal_centro.xlsx
2025-12-29 10:30:15 | INFO     |     ✓ 150 registros cargados
...
```

---

## 💼 Casos de Uso Reales

Este script es ideal para:

1. **Cadenas de retail** - Consolidar ventas de múltiples tiendas
2. **Equipos de ventas distribuidos** - Unificar reportes de diferentes regiones
3. **Franquicias** - Análisis centralizado de todas las ubicaciones
4. **Empresas con múltiples vendedores** - Seguimiento de desempeño
5. **Reportes ejecutivos mensuales** - Automatizar la generación de reportes

---

## 🎓 Notas Técnicas

### Principios de Código
- ✅ Sigue estrictamente PEP 8
- ✅ Type hints en todas las funciones
- ✅ Docstrings detallados
- ✅ Manejo robusto de errores con try-except
- ✅ Logging informativo en cada paso
- ✅ Variables con nombres descriptivos en español

### Rendimiento
- Procesa ~1000 registros en < 5 segundos
- Genera gráficos en < 3 segundos
- Memoria eficiente con pandas
- Sin dependencias pesadas

---

## 📞 Contacto y Soporte

**Autor:** Excel Automatizado
**Proyecto:** Demo 1 - Consolidador de Ventas
**Fecha:** Diciembre 2025

---

## 📄 Licencia

Este es un proyecto de demostración para portafolio. Libre para uso personal y educativo.

---

## 🚀 Próximos Pasos

1. Ejecuta `python generar_datos_ejemplo.py` para crear datos de prueba
2. Ejecuta `python consolidador.py` para generar el reporte
3. Abre `output/reporte_consolidado.xlsx` y ¡sorpréndete con el resultado!
4. Adapta el script para tus propios datos y necesidades

---

**¿Necesitas automatización personalizada de Excel?**
Este demo muestra solo una fracción de lo que es posible. ¡Contáctame para proyectos a medida!
