# 📊 Generador de Reporte de Ventas

Aplicación de escritorio desarrollada en Python para generar reportes de ventas en Excel de forma automática, con métricas calculadas, formato profesional y gráficos incluidos.

El sistema permite filtrar por rango de fechas, calcular indicadores clave y exportar un archivo Excel listo para uso empresarial.

---

## 🚀 Características

- 📂 Selección de archivo Excel (.xlsx)
- 📅 Filtro por rango de fechas
- 📈 Cálculo automático de métricas:
  - Total General
  - Producto más vendido
  - Día con mayor venta
  - Monto del día con mayor venta
- 📊 Gráfico automático de ventas por día
- 🗂 Carpeta automática `Reportes`
- 💰 Formato moneda aplicado automáticamente
- 📏 Ajuste automático de ancho de columnas
- 🔒 Encabezados congelados
- 🎨 Interfaz moderna con CustomTkinter
- 📦 Versión ejecutable (.exe) generada con PyInstaller

---

## 🧠 Tecnologías Utilizadas

- Python 3.x
- Pandas
- OpenPyXL
- CustomTkinter
- TkCalendar
- Pillow
- PyInstaller

---

## 📁 Estructura del Proyecto

```
Proyecto_Excel/
│
├── app.py              # Interfaz gráfica
├── reporte.py          # Lógica de procesamiento y generación Excel
├── logo.png            # Logo de la aplicación
├── icono.ico           # Icono del ejecutable
├── ventas_ejemplo.xlsx # Archivo de prueba
└── dist/
    └── app.exe         # Ejecutable generado
```

---

## 📥 Formato requerido del Excel

El archivo de entrada debe contener las siguientes columnas:

- `fecha`
- `vendedor`
- `producto`
- `cantidad`
- `precio`

Ejemplo:

| fecha       | vendedor | producto | cantidad | precio |
|------------|----------|----------|----------|--------|
| 2026-02-01 | Ana      | Mouse    | 2        | 10000  |

---

## ▶ Cómo ejecutar

### Ejecutar desde Python

```
python app.py
```

### Generar ejecutable

```
pyinstaller --onefile --windowed --add-data "logo.png;." --icon=icono.ico app.py
```

El ejecutable se generará en la carpeta `dist`.

---

## 📊 Funcionalidades del reporte generado

El archivo Excel incluye:

- Hoja **Ventas Detalladas**
- Hoja **Resumen**
- Hoja **Por Vendedor**
- Hoja **Ventas por Día**
- Gráfico automático de barras
- Formato moneda aplicado
- Ajuste automático de columnas
- Encabezados en negrita y congelados

---

## 🎯 Objetivo del Proyecto

Este proyecto fue desarrollado como parte del proceso de formación en desarrollo Python, con enfoque en:

- Arquitectura modular
- Manejo estructurado de errores
- Experiencia de usuario básica
- Automatización de reportes empresariales
- Empaquetado profesional de aplicaciones

---

## 📌 Posibles mejoras futuras (v1.1)

- Cierre automático diario
- Separación por método de pago (efectivo / tarjeta)
- Exportación a PDF
- Historial interno de reportes
- Panel administrativo

---

## 👨‍💻 Autor

Rene Lavanchy  
Desarrollador Python en formación