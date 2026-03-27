# 📄 Cargador de Facturas - Monotributo Argentina

[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Aplicación de escritorio diseñada para profesionales independientes y contadores en Argentina. Automatiza la extracción de datos desde PDFs de AFIP (Facturas y Notas de Crédito C) y los consolida en una planilla de control.

## ✨ Características

- **Extracción Inteligente:** Captura nombres largos (ej. Obras Sociales), importes, CAE y métodos de pago.
- **Detección de Duplicados:** Evita cargar dos veces el mismo comprobante comparando el número de CAE en tu Excel.
- **Soporte para Notas de Crédito:** Identifica automáticamente comprobantes asociados y vincula el número de factura original.
- **Interfaz Responsiva:** UI moderna con barra de progreso y procesamiento en segundo plano (threading) para evitar bloqueos.
- **Seguridad:** Los datos se procesan localmente; nada se sube a la nube.

## 📸 Vista Previa
![Vista Previa de la App](./screenshots/app_v4.png)

## 📋 Requisitos

Para instalar todo usa el comando: pip install -r requirements.txt

- Python 3.8 o superior
- Librerías necesarias:
  - `pdfplumber` (Extracción de texto precisa)
  - `pandas` & `openpyxl` (Gestión de Excel)
  - `tkinter` (Interfaz gráfica)

## 📊 Alcance de la Extracción

Para mantener la simpleza y eficiencia del reporte contable, el sistema captura únicamente los datos esenciales:

### ✅ Datos que SÍ se extraen:
- **Fecha de Emisión:** Formato DD/MM/AAAA.
- **Número de Comprobante:** Los 8 dígitos principales (ej: 00000056).
- **Razón Social del Cliente:** Nombre completo (limpiando etiquetas de AFIP).
- **Tipo de Comprobante:** Diferenciación entre "Factura C" y "Nota de Crédito C".
- **Importe Total:** Valor numérico final del comprobante.
- **CAE:** Código de Autorización Electrónico (usado para control de duplicados).
- **Método de Pago:** Extraído de la condición de venta (ej: Transferencia, Efectivo).
- **Comprobante Asociado:** Solo en Notas de Crédito (el nro. de factura que anula).

### ❌ Datos que NO se extraen (Ignorados por diseño):
- **Detalle de productos/servicios:** No se extrae la lista de ítems, cantidades o precios unitarios.
- **CUIT del Receptor:** Se prioriza la Razón Social para el listado de clientes.
- **Domicilio del Cliente:** No se considera relevante para el resumen de facturación.
- **Fechas de vencimiento o períodos:** Solo se toma la fecha de emisión del documento.
- **Logos o códigos de barras:** El sistema es exclusivamente de extracción de texto.

## ⚠️ Limitaciones Importantes
- **Documentos Originales:** Solo funciona con PDFs "nacidos digitales" de ARCA/AFIP. No procesa escaneos ni fotos (sin OCR).
- **Reglas de Negocio:** Optimizado para el diseño de facturas vigente a 2026. Cambios estructurales en los PDFs de AFIP/ARCA podrían requerir actualizaciones en los patrones de RegEx.
- **Exclusividad:** Soporta únicamente el modelo de Factura C y Nota de Crédito C (Monotributo).

## 🚀 Instalación

1. **Clonar y entrar al directorio:**
   ```bash
   git clone [https://github.com/NicoZdev/AutomatizacionTest.git](https://github.com/NicoZdev/AutomatizacionTest.git)
   cd AutomatizacionTest/contabilidad_app

## 📖 Cómo usar

- Selecciona tu archivo Excel de control.

- Arrastra o selecciona las facturas PDF (puedes seleccionar muchas a la vez).

- Haz clic en INICIAR PROCESAMIENTO.

- La aplicación omitirá automáticamente los que ya existan y te dará un reporte final.

## ⚖️ Licencia
Este proyecto está bajo la Licencia MIT.

## ❓ Preguntas Frecuentes (FAQ)

**¿Por qué Windows detecta el .exe como virus?** Al ser un ejecutable creado con Python y no estar firmado con un certificado de desarrollador pago, algunos antivirus pueden dar un "falso positivo". Es seguro añadirlo a exclusiones.

**¿Qué pasa si el programa no reconoce el comando 'pyinstaller' para generar un .exe?** Asegúrate de tener Python en el PATH de Windows o utiliza el comando:  
`python -m PyInstaller --noconsole --onefile main.py`