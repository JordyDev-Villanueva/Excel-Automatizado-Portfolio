# 📊 Excel Automatizado - Portafolio

> Soluciones prácticas de automatización de Excel con Python que resuelven problemas empresariales reales.

---

## 👋 Bienvenido

Soy un desarrollador especializado en automatización de procesos con Excel y Python. Este repositorio muestra proyectos reales que he desarrollado para optimizar tareas repetitivas, analizar datos y generar reportes ejecutivos de forma automática.

Cada demo aquí presentado es funcional, está documentado y resuelve un caso de uso específico que encontrarás en el día a día de muchas empresas.

### 💡 ¿Qué encontrarás aquí?

- **Scripts listos para usar** - Código limpio y bien estructurado
- **Documentación detallada** - Instrucciones paso a paso para cada demo
- **Ejemplos reales** - Datos de muestra para probar los scripts
- **Código reutilizable** - Funciones que puedes adaptar a tus necesidades

---

## 🎯 Proyectos Disponibles

### 1️⃣ [Consolidador de Ventas](demo1-consolidador-ventas/)

**El problema:**
Imagina que eres gerente de una empresa con varias sucursales. Cada mes, cada sucursal te envía su Excel de ventas. Tú necesitas consolidar todo, hacer análisis, crear gráficos y presentar un reporte ejecutivo. Manualmente, esto te puede tomar entre 3 a 4 horas.

**La solución:**
Este script hace todo el trabajo en menos de 30 segundos. Lee automáticamente todos los archivos, los consolida, calcula métricas, genera gráficos profesionales y crea un reporte ejecutivo listo para presentar.

**Lo que hace:**
- Lee y combina múltiples archivos Excel automáticamente
- Calcula totales, promedios y participaciones
- Genera análisis por sucursal, producto, vendedor y categoría
- Crea gráficos de alta calidad (barras, circular, línea de tendencia)
- Produce un Excel profesional con 5 hojas: Dashboard, Datos, Top Productos, Vendedores y Resumen

**Tecnologías:** Python, pandas, openpyxl, matplotlib, seaborn

**[📖 Ver documentación completa del proyecto →](demo1-consolidador-ventas/README.md)**

---

### 2️⃣ [Limpiador y Validador de Datos](demo2-limpiador-datos/) *(En desarrollo)*

**El problema:**
Recibes archivos Excel con errores: fechas mal formateadas, duplicados, espacios extra, valores faltantes, columnas inconsistentes. Limpiarlos manualmente es tedioso y propenso a errores.

**La solución:**
Un script que detecta y corrige automáticamente los errores más comunes, normaliza formatos y genera un reporte de calidad de datos.

**Lo que hará:**
- Detección automática de errores comunes
- Limpieza de espacios, caracteres especiales y duplicados
- Normalización de fechas, números y textos
- Validación de datos según reglas personalizables
- Reporte detallado de calidad con estadísticas

_Este proyecto estará disponible próximamente._

---

## 🚀 Inicio Rápido

### Requisitos Previos
- Python 3.8+
- pip

### Instalación

1. **Clonar el repositorio**
```bash
git clone https://github.com/TU_USUARIO/01-Excel-Automatizado.git
cd 01-Excel-Automatizado
```

2. **Elegir un demo** (ejemplo: Demo 1)
```bash
cd demo1-consolidador-ventas
```

3. **Instalar dependencias**
```bash
pip install -r requirements.txt
```

4. **Ejecutar el demo**
```bash
# Generar datos de ejemplo (opcional)
python generar_datos_ejemplo.py

# Ejecutar el script principal
python consolidador.py

# El resultado estará en: output/reporte_consolidado.xlsx
```

---

## 📁 Estructura del Repositorio

```
01-Excel-Automatizado/
│
├── shared_utils/                    # Código reutilizable entre demos
│   ├── __init__.py
│   └── excel_helper.py             # Funciones compartidas
│
├── demo1-consolidador-ventas/      # Demo 1: Consolidador
│   ├── consolidador.py             # Script principal
│   ├── generar_datos_ejemplo.py    # Generador de datos
│   ├── requirements.txt
│   ├── README.md                   # Documentación detallada
│   ├── input/                      # Archivos de entrada
│   └── output/                     # Resultados generados
│
├── demo2-limpiador-datos/          # Demo 2: Limpiador (próximamente)
│   └── ...
│
└── README.md                       # Este archivo
```

---

## 💼 ¿En qué puedo ayudarte?

Si tienes procesos repetitivos con Excel que te consumen tiempo, puedo ayudarte a automatizarlos. Algunos ejemplos:

- **Consolidación de reportes** - Combinar archivos de diferentes fuentes
- **Generación automática de dashboards** - KPIs actualizados sin intervención manual
- **Limpieza de datos** - Normalizar y validar información
- **Reportes ejecutivos** - Gráficos y análisis listos para presentar
- **Integración con otras herramientas** - Conectar Excel con bases de datos, APIs o sistemas empresariales

Cada solución se desarrolla según tus necesidades específicas, con código limpio, documentado y fácil de mantener.

---

## 🛠️ Tecnologías

| Categoría | Herramientas |
|-----------|--------------|
| **Lenguaje** | Python 3.8+ |
| **Datos** | pandas, numpy |
| **Excel** | openpyxl, xlsxwriter |
| **Visualización** | matplotlib, seaborn, plotly |
| **Otros** | logging, pathlib, datetime |

---

## 🎓 Sobre el código

Todos los scripts en este repositorio están desarrollados siguiendo buenas prácticas:

- **Código limpio** - Fácil de leer y entender
- **Bien documentado** - Comentarios claros explicando la lógica
- **Manejo de errores** - Validaciones para evitar fallos
- **Modular** - Funciones reutilizables que puedes adaptar
- **Probado** - Incluye datos de ejemplo para testing

No solo funciona, sino que está hecho pensando en que alguien más pueda entenderlo, modificarlo y mantenerlo.

---

## 📞 Contacto

Si necesitas ayuda con automatización de Excel, análisis de datos o tienes un proyecto en mente, puedes contactarme a través de:

- 💼 **GitHub:** [@JordyDev-Villanueva](https://github.com/JordyDev-Villanueva)
- 💼 **Fiverr:** _[Próximamente]_
- 💼 **Upwork:** _[Próximamente]_

---

## 📄 Licencia

Este repositorio es un portafolio personal que muestra proyectos de demostración. El código está disponible para consulta, aprendizaje y referencia.

---

## ⭐ Agradecimientos

Si este repositorio te resulta útil o te inspira para automatizar tus propios procesos:
- Dale una estrella ⭐ al repo
- Compártelo con otros que puedan beneficiarse
- Déjame saber si tienes sugerencias de mejora

---

**Última actualización:** Diciembre 2024
**Estado:** ✅ Activo - Demo 1 disponible
