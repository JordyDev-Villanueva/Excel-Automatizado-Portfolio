# 🤖 Excel Automatizado - Portafolio de Automatización

> **Soluciones profesionales de automatización de Excel con Python**
>
> Demos funcionales que muestran capacidades de automatización, análisis de datos y generación de reportes ejecutivos.

---

## 👋 Sobre Este Repositorio

Este repositorio contiene **demos completos y funcionales** de automatización de Excel usando Python. Cada demo resuelve un problema empresarial real y demuestra habilidades profesionales en:

- 📊 **Análisis de datos** con pandas
- 🎨 **Visualización profesional** con matplotlib/seaborn
- 📝 **Manipulación avanzada de Excel** con openpyxl
- 🔄 **Automatización de procesos** repetitivos
- 💼 **Soluciones empresariales** listas para producción

---

## 🎯 Demos Disponibles

### 1️⃣ [Consolidador de Ventas](demo1-consolidador-ventas/)
**Problema:** Consolidar reportes de múltiples sucursales manualmente toma 3-4 horas
**Solución:** Script que automatiza todo en 30 segundos

**Características:**
- ✅ Consolida múltiples archivos Excel automáticamente
- ✅ Genera 5 análisis diferentes (sucursales, productos, vendedores, etc.)
- ✅ Crea 3 gráficos profesionales de alta calidad
- ✅ Output Excel multi-hoja con formato corporativo
- ✅ Dashboard ejecutivo con KPIs

**Tecnologías:** pandas, openpyxl, matplotlib, seaborn

**[📖 Ver documentación completa →](demo1-consolidador-ventas/README.md)**

---

### 2️⃣ [Limpiador de Datos](demo2-limpiador-datos/) *(Próximamente)*
**Problema:** Datos sucios y errores comunes en archivos Excel
**Solución:** Limpieza automática con reporte de calidad

**Características:**
- ✅ Detección de errores comunes
- ✅ Normalización de formatos
- ✅ Validación de datos
- ✅ Reporte de calidad

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

## 💼 Servicios Ofrecidos

Basándome en estos demos, ofrezco:

### 🔹 Automatización de Excel
- Consolidación de múltiples archivos
- Generación automática de reportes
- Actualización de dashboards
- Procesamiento masivo de datos

### 🔹 Análisis de Datos
- Limpieza y normalización
- Análisis exploratorio
- Cálculos y métricas personalizadas
- Detección de patrones

### 🔹 Visualización
- Gráficos profesionales para presentaciones
- Dashboards ejecutivos
- Reportes con formato corporativo
- Inserción de visualizaciones en Excel

### 🔹 Integración
- APIs y bases de datos
- Sistemas ERP/CRM
- Google Sheets
- Automatización de workflows

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

## 📊 Casos de Uso Reales

Estos scripts son ideales para:

✅ **Empresas con múltiples sucursales** - Consolidar reportes
✅ **Equipos de ventas** - Análisis de desempeño
✅ **Departamentos financieros** - Reportes mensuales
✅ **Gerencias** - Dashboards ejecutivos
✅ **Analistas de datos** - Automatizar tareas repetitivas

---

## 🎓 Características del Código

- ✅ **PEP 8 compliant** - Código limpio y profesional
- ✅ **Documentación completa** - Docstrings en todas las funciones
- ✅ **Type hints** - Parámetros tipados
- ✅ **Manejo de errores** - Try-except robusto
- ✅ **Logging detallado** - Trazabilidad completa
- ✅ **Modular y reutilizable** - Fácil de adaptar
- ✅ **README detallados** - Instrucciones paso a paso

---

## 📞 Contacto

¿Necesitas automatización personalizada de Excel o análisis de datos?

- 💼 **Fiverr:** [Tu perfil]
- 💼 **Upwork:** [Tu perfil]
- 📧 **Email:** tu@email.com
- 💻 **GitHub:** [@TuUsuario](https://github.com/TuUsuario)

---

## 📄 Licencia

Este repositorio es un portafolio de demostración. Los scripts son libres para uso personal y educativo.

---

## ⭐ ¿Te gustó?

Si encuentras útiles estos demos:
- ⭐ Dale una estrella al repositorio
- 🔄 Comparte con otros
- 💬 Deja comentarios o sugerencias
- 📧 Contáctame para proyectos personalizados

---

**Última actualización:** Diciembre 2024
**Versión:** 1.0 - Demo 1 completo
**Estado:** ✅ Producción
