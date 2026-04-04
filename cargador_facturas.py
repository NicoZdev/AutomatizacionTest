import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pdfplumber
import re
import os
import threading
import gspread # Necesitarás instalarlo: pip install gspread google-auth
from google.oauth2.service_account import Credentials

class ExtractorFacturas:
    def __init__(self):
        self.patrones = {
            'fecha': r'Fecha\s*de\s*Emisi[oó]n[:\s]*(\d{2}/\d{2}/\d{4})',
            'nro': r'Comp\.\s*Nro[:\s]*(\d{1,8})',
            'cae': r'CAE\s*N[°º]?[:\s]*(\d{14})',
            'importe': r'Importe\s*Total[:\s]*\$?\s*([\d\.]+,\d{2})',
            'condicion_venta': r'Condici[oó]n\s*de\s*venta[:\s]*(.*?)(?:\n|$)',
        }

    def limpiar_importe(self, texto):
        if not texto: return 0.0
        try:
            limpio = texto.replace('$', '').replace('.', '').replace(',', '.').strip()
            return float(limpio)
        except: return 0.0

    def extraer_datos(self, ruta_pdf):
        datos = {
            'fecha': '', 'comp_nro': '', 'cliente': 'Consumidor Final',
            'tipo': 'Factura C', 'servicio': 'Servicio prestado', 'importe': 0.0,
            'cae': '', 'estado': '', 'condicion_venta': '',
            'archivo_pdf': os.path.basename(ruta_pdf),
            'notas': ''
        }
        try:
            with pdfplumber.open(ruta_pdf) as pdf:
                texto = pdf.pages[0].extract_text()
                
                if "NOTA DE CRÉDITO" in texto.upper():
                    datos['tipo'] = "Nota de Crédito C"
                    asociado = re.search(r'Fac\.\s*C[:\s]*\d{5}-(\d{8})', texto)
                    if asociado: datos['notas'] = asociado.group(1)
                else:
                    datos['tipo'] = "Factura C"

                m_fecha = re.search(self.patrones['fecha'], texto)
                if m_fecha: datos['fecha'] = m_fecha.group(1)

                m_nro = re.search(self.patrones['nro'], texto)
                if m_nro: datos['comp_nro'] = m_nro.group(1).zfill(8)

                bloque_cliente = re.search(r'Apellido y Nombre / Razón Social:(.*?)Domicilio:', texto, re.S)
                if bloque_cliente:
                    nombre = bloque_cliente.group(1).replace('\n', ' ').strip()
                    nombre = re.split(r'Condici[oó]n|CUIT', nombre, flags=re.I)[0].strip()
                    if len(nombre) > 2: datos['cliente'] = " ".join(nombre.split()).upper()

                m_imp = re.search(self.patrones['importe'], texto)
                if m_imp: datos['importe'] = self.limpiar_importe(m_imp.group(1))

                m_cae = re.search(self.patrones['cae'], texto)
                if m_cae: datos['cae'] = m_cae.group(1)

                m_cond = re.search(self.patrones['condicion_venta'], texto)
                if m_cond:
                    metodo = m_cond.group(1).strip()
                    datos['condicion_venta'] = re.split(r'Fac\.\s*C', metodo, flags=re.I)[0].strip()

        except Exception as e:
            datos['cliente'] = f"Error: {str(e)}"
        return datos

class AplicacionCargador:
    def __init__(self, root):
        self.root = root
        self.root.title("Cargador Sheets Pro - v4.0")
        self.root.geometry("900x700")
        self.root.configure(bg='#f8f9fa')
        
        self.pdfs_seleccionados = []
        self.nombre_sheet = "Facturacion" # Nombre del archivo en Google Drive
        self.extractor = ExtractorFacturas()
        
        self.crear_interfaz()

    def crear_interfaz(self):
        self.main = tk.Frame(self.root, bg='#f8f9fa', padx=30, pady=30)
        self.main.grid(sticky="nsew")
        self.main.columnconfigure(0, weight=1)

        lbl_titulo = tk.Label(self.main, text="Carga Directa a Google Sheets", 
                              font=("Segoe UI", 18, "bold"), bg='#f8f9fa', fg='#2c3e50')
        lbl_titulo.grid(row=0, column=0, pady=(0, 20))

        # --- CONFIGURACIÓN SHEET ---
        frame_conf = tk.LabelFrame(self.main, text=" 1. Configuración Google Sheets ", font=("Segoe UI", 10, "bold"), bg='#f8f9fa', padx=15, pady=15)
        frame_conf.grid(row=1, column=0, sticky="ew", pady=10)
        
        tk.Label(frame_conf, text="Nombre del archivo Sheets:", bg='#f8f9fa').grid(row=0, column=0, sticky="w")
        self.ent_sheet = tk.Entry(frame_conf, width=40)
        self.ent_sheet.insert(0, "Mi Facturacion")
        self.ent_sheet.grid(row=0, column=1, padx=10)

        # --- SECCIÓN PDF ---
        frame_pdf = tk.LabelFrame(self.main, text=" 2. Comprobantes PDF ", font=("Segoe UI", 10, "bold"), bg='#f8f9fa', padx=15, pady=15)
        frame_pdf.grid(row=2, column=0, sticky="nsew", pady=10)
        frame_pdf.columnconfigure(0, weight=1)

        self.btn_pdf = tk.Button(frame_pdf, text="+ Agregar Facturas", command=self.seleccionar_pdfs,
                                 bg='#27ae60', fg='white', font=("Segoe UI", 10), relief='flat', cursor='hand2', padx=20)
        self.btn_pdf.grid(row=0, column=0, pady=(0,10))
        
        self.listbox = tk.Listbox(frame_pdf, font=("Consolas", 10), height=6)
        self.listbox.grid(row=1, column=0, sticky="nsew", pady=5)

        self.progress = ttk.Progressbar(self.main, orient="horizontal", mode="determinate")
        self.progress.grid(row=3, column=0, sticky="ew", pady=10)

        self.btn_run = tk.Button(self.main, text="INICIAR CARGA A LA NUBE", command=self.iniciar_hilo, 
                                bg='#4a90e2', fg='white', font=("Segoe UI", 12, "bold"), relief='flat', cursor='hand2', pady=12)
        self.btn_run.grid(row=4, column=0, sticky="ew", pady=10)

        self.txt_log = scrolledtext.ScrolledText(self.main, height=8, font=("Consolas", 9))
        self.txt_log.grid(row=5, column=0, sticky="ew", pady=5)

    def log(self, msj):
        self.txt_log.insert(tk.END, f"> {msj}\n")
        self.txt_log.see(tk.END)

    def seleccionar_pdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        for f in files:
            if f not in self.pdfs_seleccionados:
                self.pdfs_seleccionados.append(f)
                self.listbox.insert(tk.END, f" 📄 {os.path.basename(f)}")

    def conectar_sheets(self):
        # Asegúrate de que credentials.json esté en la carpeta
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_file("credentials.json", scopes=scope)
        client = gspread.authorize(creds)
        return client.open(self.ent_sheet.get()).get_worksheet(0)

    def iniciar_hilo(self):
        if not self.pdfs_seleccionados:
            messagebox.showwarning("Atención", "Selecciona los PDF primero.")
            return
        self.btn_run.config(state='disabled', text="CONECTANDO A GOOGLE...")
        threading.Thread(target=self.cargar_datos, daemon=True).start()

    def mapear_metodo_pago(self, texto):
        t = texto.upper()
        if "TRANSFERENCIA" in t: return "Transferencia"
        if "CONTADO" in t or "EFECTIVO" in t: return "Efectivo" # Mapeo a Efectivo
        if "DEBITO" in t or "DÉBITO" in t: return "Débito"
        if "CREDITO" in t or "CRÉDITO" in t: return "Crédito"
        return "Efectivo"

    def cargar_datos(self):
        try:
            sheet = self.conectar_sheets()
            # Obtenemos CAEs para evitar duplicados
            existentes = set(sheet.col_values(7)) 
            
            # Buscamos la fila de inicio basada en la columna B
            col_b_values = sheet.col_values(2)
            current_row = len(col_b_values) + 1
            
            for i, p in enumerate(self.pdfs_seleccionados):
                res = self.extractor.extraer_datos(p)
                
                if str(res['cae']) in existentes:
                    self.log(f"Duplicado omitido: {res['comp_nro']}")
                else:
                    self.log(f"Procesando: {res['comp_nro']}")
                    
                    # Formateo de datos según tus reglas
                    tipo_corto = "Nota de Credito C" if "NOTA" in res['tipo'].upper() else "C"
                    pago_final = self.mapear_metodo_pago(res['condicion_venta'])
                    
                    nombre_archivo = os.path.basename(p).strip()
                    # Fórmula de búsqueda en Drive
                    formula_pdf = f'=HYPERLINK("https://www.google.com/drive/s?q={nombre_archivo}"; "{nombre_archivo}")'

                    # --- CARGA DE BLOQUES ---
                    
                    # Bloque Izquierdo (B a G)
                    # Usamos value_input_option='USER_ENTERED' para que reconozca números y formatos
                    sheet.update(
                        range_name=f"B{current_row}:G{current_row}", 
                        values=[[res['comp_nro'], res['cliente'], tipo_corto, res['servicio'], res['importe'], res['cae']]],
                        value_input_option='USER_ENTERED'
                    )
                    
                    # Bloque Derecho (I a K)
                    # Enviamos "" en la última posición para asegurar que Notas (K) quede vacía
                    sheet.update(
                        range_name=f"I{current_row}:K{current_row}", 
                        values=[[pago_final, formula_pdf, ""]], 
                        value_input_option='USER_ENTERED' # CRITICO: Esto quita el apóstrofe (') de la fórmula
                    )
                    
                    current_row += 1
                
                self.progress["value"] = i + 1
                self.root.update_idletasks()

            self.log("--- CARGA FINALIZADA ---")
            messagebox.showinfo("Éxito", "Carga completada. Fórmulas activas y notas limpias.")

        except Exception as e:
            self.log(f"ERROR: {e}")
            messagebox.showerror("Error", f"Error en la carga: {e}")

        # Reset interfaz
        self.btn_run.config(state='normal', text="INICIAR CARGA A LA NUBE")
        self.progress["value"] = 0
        self.pdfs_seleccionados = []
        self.listbox.delete(0, tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = AplicacionCargador(root)
    root.mainloop()