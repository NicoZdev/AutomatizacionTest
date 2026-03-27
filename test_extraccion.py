"""
Script de prueba para verificar la extracción de datos de facturas PDF
"""

import PyPDF2
import re
import os
from datetime import datetime

class ExtractorFacturas:
    """Clase para extraer datos de facturas C y notas de crédito C"""
    
    def __init__(self):
        self.patrones = {
            'fecha': r'(\d{2}/\d{2}/\d{4})',
            'cuit': r'CUIT[:\s]*(\d{2}-?\d{8}-?\d{1})',
            'importe_total': r'(?:TOTAL|Importe Total)[:\s]*\$?\s*([\d.,]+)',
            'neto_gravado': r'(?:Neto Gravado|Subtotal)[:\s]*\$?\s*([\d.,]+)',
            'iva': r'IVA(?:\s+21%)?[:\s]*\$?\s*([\d.,]+)',
            'punto_venta': r'Pto\.?\s*Venta[:\s]*(\d+)',
            'nro_comprobante': r'Nro(?:\s+)?(?:de\s+)?(?:Comp\.?|Comprobante)[:\s]*(\d+)',
            'tipo_comprobante': r'(FACTURA C|NOTA DE CREDITO C|FACTURA "C"|NOTA DE CREDITO "C")',
            'condicion_iva': r'(Monotributo|Consumidor Final|Responsable Inscripto|Exento)'
        }
    
    def limpiar_numero(self, texto):
        """Limpia un string para convertirlo a número"""
        if not texto:
            return 0.0
        texto = texto.replace('$', '').replace('.', '').replace(',', '.')
        try:
            return float(texto)
        except:
            return 0.0
    
    def extraer_datos(self, ruta_pdf):
        """Extrae datos de un PDF de factura"""
        datos = {
            'archivo': os.path.basename(ruta_pdf),
            'tipo_comprobante': '',
            'fecha_emision': '',
            'cuit_emisor': '',
            'cuit_cliente': '',
            'punto_venta': '',
            'numero_comprobante': '',
            'importe_total': 0.0,
            'neto_gravado': 0.0,
            'iva': 0.0,
            'condicion_iva': '',
            'estado': 'OK'
        }
        
        try:
            with open(ruta_pdf, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                texto_completo = ''
                
                for page in reader.pages:
                    texto_completo += page.extract_text() + '\n'
                
                texto_completo = texto_completo.upper()
                
                # Mostrar texto extraído para depuración
                print(f"\n{'='*80}")
                print(f"TEXTO EXTRAÍDO DE: {datos['archivo']}")
                print(f"{'='*80}")
                print(texto_completo[:1000])  # Primeros 1000 caracteres
                print(f"{'='*80}\n")
                
                # Extraer tipo de comprobante
                match = re.search(self.patrones['tipo_comprobante'], texto_completo)
                if match:
                    datos['tipo_comprobante'] = match.group(1).replace('"', '')
                
                # Extraer fecha
                match = re.search(self.patrones['fecha'], texto_completo)
                if match:
                    datos['fecha_emision'] = match.group(1)
                
                # Extraer CUITs
                cuits = re.findall(self.patrones['cuit'], texto_completo)
                if len(cuits) >= 1:
                    datos['cuit_emisor'] = cuits[0].replace('-', '')
                if len(cuits) >= 2:
                    datos['cuit_cliente'] = cuits[1].replace('-', '')
                
                # Extraer punto de venta
                match = re.search(self.patrones['punto_venta'], texto_completo)
                if match:
                    datos['punto_venta'] = match.group(1)
                
                # Extraer número de comprobante
                match = re.search(self.patrones['nro_comprobante'], texto_completo)
                if match:
                    datos['numero_comprobante'] = match.group(1)
                
                # Extraer importes
                match = re.search(self.patrones['importe_total'], texto_completo)
                if match:
                    datos['importe_total'] = self.limpiar_numero(match.group(1))
                
                match = re.search(self.patrones['neto_gravado'], texto_completo)
                if match:
                    datos['neto_gravado'] = self.limpiar_numero(match.group(1))
                
                match = re.search(self.patrones['iva'], texto_completo)
                if match:
                    datos['iva'] = self.limpiar_numero(match.group(1))
                
                # Extraer condición IVA
                match = re.search(self.patrones['condicion_iva'], texto_completo)
                if match:
                    datos['condicion_iva'] = match.group(1)
                
        except Exception as e:
            datos['estado'] = f'ERROR: {str(e)}'
            print(f"Error: {str(e)}")
        
        return datos


def probar_extraccion(ruta_pdf):
    """Función principal de prueba"""
    print("\n" + "="*80)
    print("PRUEBA DE EXTRACCIÓN DE FACTURAS")
    print("="*80)
    print(f"Archivo: {ruta_pdf}")
    print(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*80)
    
    extractor = ExtractorFacturas()
    datos = extractor.extraer_datos(ruta_pdf)
    
    print("\nRESULTADOS DE LA EXTRACCIÓN:")
    print("-"*80)
    for clave, valor in datos.items():
        if clave != 'archivo':
            print(f"{clave.replace('_', ' ').title()}: {valor}")
    
    print("-"*80)
    
    return datos


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        ruta_archivo = sys.argv[1]
    else:
        ruta_archivo = input("\nIngrese la ruta del archivo PDF a probar: ").strip()
    
    if os.path.exists(ruta_archivo):
        probar_extraccion(ruta_archivo)
    else:
        print(f"❌ El archivo no existe: {ruta_archivo}")
        print("\nUso: python test_extraccion.py <ruta_del_pdf>")
