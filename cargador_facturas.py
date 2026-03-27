import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import PyPDF2
import re
from datetime import datetime
import pandas as pd
import os
from pathlib import Path
import logging

# Configuración de logs
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('carga_facturas.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

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
                
                # Extraer tipo de comprobante
                match = re.search(self.patrones['tipo_comprobante'], texto_completo)
                if match:
                    datos['tipo_comprobante'] = match.group(1).replace('"', '')
                
                # Extraer fecha
                match = re.search(self.patrones['fecha'], texto_completo)
                if match:
                    datos['fecha_emision'] = match.group(1)
                
                # Extraer CUITs (generalmente hay 2: emisor y cliente)
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
                
                logger.info(f"Datos extraídos de {datos['archivo']}: {datos}")
                
        except Exception as e:
            datos['estado'] = f'ERROR: {str(e)}'
            logger.error(f"Error extrayendo {ruta_pdf}: {str(e)}")
        
        return datos


class AplicacionCargador:
    """Aplicación gráfica para cargar facturas"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Cargador de Facturas - Monotributo Argentina")
        self.root.geometry("900x700")
        self.root.configure(bg='#f0f0f0')
        
        self.pdfs_seleccionados = []
        self.archivo_excel = ""
        self.extractor = ExtractorFacturas()
        
        self.crear_interfaz()
    
    def crear_interfaz(self):
        """Crea la interfaz gráfica"""
        # Frame superior
        frame_superior = tk.Frame(self.root, bg='#2c3e50', height=80)
        frame_superior.pack(fill=tk.X, padx=0, pady=0)
        
        titulo = tk.Label(
            frame_superior, 
            text="📄 CARGADOR DE FACTURAS C",
            font=('Arial', 18, 'bold'),
            bg='#2c3e50',
            fg='white'
        )
        titulo.pack(pady=20)
        
        subtítulo = tk.Label(
            frame_superior,
            text="Para Monotributistas en Argentina",
            font=('Arial', 10),
            bg='#2c3e50',
            fg='#bdc3c7'
        )
        subtítulo.pack()
        
        # Frame principal
        frame_principal = tk.Frame(self.root, bg='#f0f0f0')
        frame_principal.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Sección de selección de PDFs
        frame_pdfs = tk.LabelFrame(
            frame_principal, 
            text="📁 Archivos PDF a procesar",
            font=('Arial', 12, 'bold'),
            bg='#f0f0f0',
            padx=10,
            pady=10
        )
        frame_pdfs.pack(fill=tk.X, pady=(0, 10))
        
        btn_seleccionar_pdfs = tk.Button(
            frame_pdfs,
            text="Seleccionar PDFs",
            command=self.seleccionar_pdfs,
            bg='#3498db',
            fg='white',
            font=('Arial', 11, 'bold'),
            cursor='hand2',
            width=20
        )
        btn_seleccionar_pdfs.pack(side=tk.LEFT, padx=(0, 10))
        
        self.lbl_cantidad_pdfs = tk.Label(
            frame_pdfs,
            text="0 archivos seleccionados",
            font=('Arial', 10),
            bg='#f0f0f0'
        )
        self.lbl_cantidad_pdfs.pack(side=tk.LEFT)
        
        btn_limpiar_pdfs = tk.Button(
            frame_pdfs,
            text="Limpiar",
            command=self.limpiar_pdfs,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 10),
            cursor='hand2',
            width=10
        )
        btn_limpiar_pdfs.pack(side=tk.RIGHT)
        
        # Lista de PDFs
        frame_lista = tk.Frame(frame_principal, bg='white', relief=tk.SUNKEN, bd=1)
        frame_lista.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.lista_pdfs = tk.Listbox(
            frame_lista,
            font=('Courier', 9),
            selectbackground='#3498db',
            selectforeground='white'
        )
        scrollbar = tk.Scrollbar(frame_lista, orient=tk.VERTICAL, command=self.lista_pdfs.yview)
        self.lista_pdfs.configure(yscrollcommand=scrollbar.set)
        
        self.lista_pdfs.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Sección de selección de Excel
        frame_excel = tk.LabelFrame(
            frame_principal,
            text="📊 Archivo Excel de destino",
            font=('Arial', 12, 'bold'),
            bg='#f0f0f0',
            padx=10,
            pady=10
        )
        frame_excel.pack(fill=tk.X, pady=(0, 10))
        
        btn_seleccionar_excel = tk.Button(
            frame_excel,
            text="Seleccionar Excel",
            command=self.seleccionar_excel,
            bg='#2ecc71',
            fg='white',
            font=('Arial', 11, 'bold'),
            cursor='hand2',
            width=20
        )
        btn_seleccionar_excel.pack(side=tk.LEFT, padx=(0, 10))
        
        self.lbl_archivo_excel = tk.Label(
            frame_excel,
            text="Ningún archivo seleccionado",
            font=('Arial', 10),
            bg='#f0f0f0',
            wraplength=500
        )
        self.lbl_archivo_excel.pack(side=tk.LEFT)
        
        # Barra de progreso
        frame_progreso = tk.Frame(frame_principal, bg='#f0f0f0')
        frame_progreso.pack(fill=tk.X, pady=(0, 10))
        
        self.barra_progreso = ttk.Progressbar(
            frame_progreso,
            mode='determinate',
            length=600
        )
        self.barra_progreso.pack(fill=tk.X)
        
        self.lbl_progreso = tk.Label(
            frame_progreso,
            text="",
            font=('Arial', 9),
            bg='#f0f0f0'
        )
        self.lbl_progreso.pack()
        
        # Botón de carga
        btn_cargar = tk.Button(
            frame_principal,
            text="⚡ CARGAR DATOS ⚡",
            command=self.cargar_datos,
            bg='#e67e22',
            fg='white',
            font=('Arial', 14, 'bold'),
            cursor='hand2',
            width=30,
            height=2
        )
        btn_cargar.pack(pady=10)
        
        # Área de logs
        frame_logs = tk.LabelFrame(
            frame_principal,
            text="📝 Registro de operaciones",
            font=('Arial', 10, 'bold'),
            bg='#f0f0f0',
            padx=5,
            pady=5
        )
        frame_logs.pack(fill=tk.BOTH, expand=True)
        
        self.txt_logs = scrolledtext.ScrolledText(
            frame_logs,
            height=8,
            font=('Courier', 8),
            bg='white',
            relief=tk.SUNKEN,
            bd=1
        )
        self.txt_logs.pack(fill=tk.BOTH, expand=True)
        
        # Inicializar logs
        self.log("Aplicación iniciada. Seleccione PDFs para comenzar.")
    
    def log(self, mensaje):
        """Agrega un mensaje al área de logs"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.txt_logs.insert(tk.END, f"[{timestamp}] {mensaje}\n")
        self.txt_logs.see(tk.END)
        self.root.update_idletasks()
    
    def seleccionar_pdfs(self):
        """Abre diálogo para seleccionar múltiples PDFs"""
        archivos = filedialog.askopenfilenames(
            title="Seleccionar Facturas PDF",
            filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")]
        )
        
        if archivos:
            self.pdfs_seleccionados.extend(archivos)
            self.actualizar_lista_pdfs()
            self.log(f"{len(archivos)} PDF(s) agregado(s). Total: {len(self.pdfs_seleccionados)}")
    
    def actualizar_lista_pdfs(self):
        """Actualiza la lista visual de PDFs"""
        self.lista_pdfs.delete(0, tk.END)
        self.lbl_cantidad_pdfs.config(text=f"{len(self.pdfs_seleccionados)} archivos seleccionados")
        
        for pdf in self.pdfs_seleccionados:
            self.lista_pdfs.insert(tk.END, os.path.basename(pdf))
    
    def limpiar_pdfs(self):
        """Limpia la selección de PDFs"""
        self.pdfs_seleccionados = []
        self.lista_pdfs.delete(0, tk.END)
        self.lbl_cantidad_pdfs.config(text="0 archivos seleccionados")
        self.log("Lista de PDFs limpiada")
    
    def seleccionar_excel(self):
        """Abre diálogo para seleccionar archivo Excel"""
        archivo = filedialog.asksaveasfilename(
            title="Seleccionar/Crear archivo Excel",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel Workbook", "*.xlsx"),
                ("Excel 97-2003", "*.xls"),
                ("Todos los archivos", "*.*")
            ]
        )
        
        if archivo:
            self.archivo_excel = archivo
            self.lbl_archivo_excel.config(text=os.path.basename(archivo))
            self.log(f"Excel seleccionado: {os.path.basename(archivo)}")
    
    def cargar_datos(self):
        """Procesa los PDFs y carga los datos en Excel"""
        if not self.pdfs_seleccionados:
            messagebox.showwarning("Advertencia", "No hay PDFs seleccionados")
            return
        
        if not self.archivo_excel:
            messagebox.showwarning("Advertencia", "No se ha seleccionado un archivo Excel")
            return
        
        self.log(f"Iniciando procesamiento de {len(self.pdfs_seleccionados)} PDF(s)...")
        
        total = len(self.pdfs_seleccionados)
        datos_extraidos = []
        
        for i, pdf in enumerate(self.pdfs_seleccionados, 1):
            self.log(f"Procesando ({i}/{total}): {os.path.basename(pdf)}")
            
            datos = self.extractor.extraer_datos(pdf)
            datos_extraidos.append(datos)
            
            # Actualizar barra de progreso
            progreso = (i / total) * 100
            self.barra_progreso['value'] = progreso
            self.lbl_progreso.config(text=f"Procesando: {i} de {total}")
            self.root.update_idletasks()
        
        # Guardar en Excel
        try:
            self.guardar_en_excel(datos_extraidos)
            self.log("✅ ¡Datos cargados exitosamente!")
            messagebox.showinfo("Éxito", f"Se procesaron {total} facturas correctamente\nArchivo: {os.path.basename(self.archivo_excel)}")
        except Exception as e:
            self.log(f"❌ Error guardando en Excel: {str(e)}")
            messagebox.showerror("Error", f"No se pudo guardar en Excel: {str(e)}")
        
        # Resetear barra
        self.barra_progreso['value'] = 0
        self.lbl_progreso.config(text="")
    
    def guardar_en_excel(self, datos):
        """Guarda los datos extraídos en un archivo Excel"""
        df = pd.DataFrame(datos)
        
        # Reordenar columnas
        columnas_orden = [
            'archivo', 'tipo_comprobante', 'fecha_emision', 'cuit_emisor',
            'cuit_cliente', 'punto_venta', 'numero_comprobante',
            'importe_total', 'neto_gravado', 'iva', 'condicion_iva', 'estado'
        ]
        df = df[columnas_orden]
        
        # Si el archivo ya existe, intentar agregar datos
        if os.path.exists(self.archivo_excel):
            try:
                df_existente = pd.read_excel(self.archivo_excel)
                df = pd.concat([df_existente, df], ignore_index=True)
            except:
                pass  # Si falla, simplemente sobrescribe
        
        # Guardar
        df.to_excel(self.archivo_excel, index=False, sheet_name='Facturas')
        self.log(f"Datos guardados en: {self.archivo_excel}")


def main():
    root = tk.Tk()
    app = AplicacionCargador(root)
    root.mainloop()


if __name__ == "__main__":
    main()
