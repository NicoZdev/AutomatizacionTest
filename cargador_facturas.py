import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pdfplumber
import re
import pandas as pd
import os
import threading

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
            'tipo': 'Factura C', 'servicio': 'Sesión de T.O', 'importe': 0.0,
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
        self.root.title("Cargador Contable Pro - v4.0")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        self.root.configure(bg='#f8f9fa')
        
        self.pdfs_seleccionados = []
        self.archivo_excel = ""
        self.extractor = ExtractorFacturas()
        
        # Hacer que la ventana sea responsiva
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        self.crear_interfaz()

    def crear_interfaz(self):
        # Contenedor principal con padding
        self.main = tk.Frame(self.root, bg='#f8f9fa', padx=30, pady=30)
        self.main.grid(sticky="nsew")
        self.main.columnconfigure(0, weight=1) # Centrado de contenido

        # Título
        lbl_titulo = tk.Label(self.main, text="Gestión Automática de Facturas", 
                              font=("Segoe UI", 18, "bold"), bg='#f8f9fa', fg='#2c3e50')
        lbl_titulo.grid(row=0, column=0, pady=(0, 20))

        # --- SECCIÓN EXCEL ---
        frame_excel = tk.LabelFrame(self.main, text=" 1. Destino de datos ", font=("Segoe UI", 10, "bold"), bg='#f8f9fa', padx=15, pady=15)
        frame_excel.grid(row=1, column=0, sticky="ew", pady=10)
        frame_excel.columnconfigure(0, weight=1)

        self.btn_excel = tk.Button(frame_excel, text="Seleccionar Archivo Excel", command=self.seleccionar_excel,
                                   bg='#4a90e2', fg='white', font=("Segoe UI", 10), relief='flat', cursor='hand2', padx=20)
        self.btn_excel.grid(row=0, column=0)
        
        self.lbl_ex = tk.Label(frame_excel, text="Ningún Excel seleccionado", bg='#f8f9fa', fg='#7f8c8d', font=("Segoe UI", 9, "italic"))
        self.lbl_ex.grid(row=1, column=0, pady=(5,0))

        # --- SECCIÓN PDF ---
        frame_pdf = tk.LabelFrame(self.main, text=" 2. Comprobantes PDF ", font=("Segoe UI", 10, "bold"), bg='#f8f9fa', padx=15, pady=15)
        frame_pdf.grid(row=2, column=0, sticky="nsew", pady=10)
        frame_pdf.columnconfigure(0, weight=1)
        frame_pdf.rowconfigure(1, weight=1)

        self.btn_pdf = tk.Button(frame_pdf, text="+ Agregar Facturas", command=self.seleccionar_pdfs,
                                 bg='#27ae60', fg='white', font=("Segoe UI", 10), relief='flat', cursor='hand2', padx=20)
        self.btn_pdf.grid(row=0, column=0, pady=(0,10))
        
        self.listbox = tk.Listbox(frame_pdf, font=("Consolas", 10), relief='flat', borderwidth=1, highlightthickness=1)
        self.listbox.grid(row=1, column=0, sticky="nsew", pady=5)

        # --- PROGRESO Y ACCIÓN ---
        self.progress = ttk.Progressbar(self.main, orient="horizontal", mode="determinate")
        self.progress.grid(row=3, column=0, sticky="ew", pady=10)

        self.btn_run = tk.Button(self.main, text="INICIAR PROCESAMIENTO", command=self.iniciar_hilo, 
                                bg='#e67e22', fg='white', font=("Segoe UI", 12, "bold"), relief='flat', cursor='hand2', pady=12)
        self.btn_run.grid(row=4, column=0, sticky="ew", pady=10)

        # Log de consola
        self.txt_log = scrolledtext.ScrolledText(self.main, height=8, font=("Consolas", 9), bg='#ffffff', borderwidth=1)
        self.txt_log.grid(row=5, column=0, sticky="ew", pady=5)

    def log(self, msj):
        self.txt_log.insert(tk.END, f"> {msj}\n")
        self.txt_log.see(tk.END)

    def seleccionar_excel(self):
        file = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if file:
            self.archivo_excel = file
            self.lbl_ex.config(text=os.path.basename(file), fg='#2c3e50', font=("Segoe UI", 9, "bold"))

    def seleccionar_pdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        for f in files:
            if f not in self.pdfs_seleccionados:
                self.pdfs_seleccionados.append(f)
                self.listbox.insert(tk.END, f" 📄 {os.path.basename(f)}")

    def iniciar_hilo(self):
        if not self.archivo_excel or not self.pdfs_seleccionados:
            messagebox.showwarning("Atención", "Por favor, selecciona el Excel y los PDF primero.")
            return
        # Desactivar botón y arrancar proceso en segundo plano
        self.btn_run.config(state='disabled', text="PROCESANDO...")
        threading.Thread(target=self.cargar_datos, daemon=True).start()

    def cargar_datos(self):
        existentes = set()
        if os.path.exists(self.archivo_excel):
            try:
                df_old = pd.read_excel(self.archivo_excel, header=None)
                existentes = set(df_old.iloc[:, 6].astype(str).str.strip().tolist())
            except: pass

        nuevos_datos = []
        duplicados = 0
        total = len(self.pdfs_seleccionados)
        self.progress["maximum"] = total
        
        for i, p in enumerate(self.pdfs_seleccionados):
            res = self.extractor.extraer_datos(p)
            
            if res['cae'] in existentes:
                self.log(f"Duplicado omitido: {res['comp_nro']}")
                duplicados += 1
            else:
                self.log(f"Procesado con éxito: {res['comp_nro']}")
                nuevos_datos.append([
                    res['fecha'], res['comp_nro'], res['cliente'], res['tipo'],
                    res['servicio'], res['importe'], res['cae'], '', 
                    res['condicion_venta'], res['archivo_pdf'], res['notas']
                ])
            
            # Actualizar barra desde el hilo
            self.progress["value"] = i + 1
            self.root.update_idletasks()

        if nuevos_datos:
            try:
                df_nuevo = pd.DataFrame(nuevos_datos)
                with pd.ExcelWriter(self.archivo_excel, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    sheet_name = writer.book.sheetnames[0]
                    start_row = writer.book[sheet_name].max_row
                    df_nuevo.to_excel(writer, index=False, header=False, startrow=start_row, sheet_name=sheet_name)
                
                self.log("--- CARGA FINALIZADA CON ÉXITO ---")
                messagebox.showinfo("Completado", f"Se cargaron {len(nuevos_datos)} facturas.\nOmitidas por duplicado: {duplicados}")
            except Exception as e:
                self.log(f"ERROR AL GUARDAR: {e}")
                messagebox.showerror("Error", "Cierra el Excel antes de procesar.")
        else:
            messagebox.showinfo("Fin", f"No se encontraron datos nuevos ({duplicados} duplicados).")

        # Resetear interfaz
        self.btn_run.config(state='normal', text="INICIAR PROCESAMIENTO")
        self.progress["value"] = 0
        self.pdfs_seleccionados = []
        self.listbox.delete(0, tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = AplicacionCargador(root)
    root.mainloop()
    