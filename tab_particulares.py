import os
import sys
import json
import sqlite3
import datetime
import tkinter as tk
from tkinter import ttk, messagebox
from fpdf import FPDF
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# ==========================================
# CODIGOS POSTALES PROVINCIA DE BARCELONA
# ==========================================
def _build_bcn_cp_dict():
    cps = {}
    for i in range(8001, 8043):
        cps[f"{i:05d}"] = "Barcelona"
    for i in range(8901, 8910):
        cps[f"{i:05d}"] = "L'Hospitalet de Llobregat"
    for i in range(8911, 8919):
        cps[f"{i:05d}"] = "Badalona"
    for i in range(8921, 8925):
        cps[f"{i:05d}"] = "Santa Coloma de Gramenet"
    for i in range(8201, 8209):
        cps[f"{i:05d}"] = "Sabadell"
    for i in range(8221, 8230):
        cps[f"{i:05d}"] = "Terrassa"
    for i in range(8301, 8305):
        cps[f"{i:05d}"] = "Mataro"
    for i in range(8401, 8404):
        cps[f"{i:05d}"] = "Granollers"
    for i in range(8241, 8244):
        cps[f"{i:05d}"] = "Manresa"
    for i in range(8172, 8175):
        cps[f"{i:05d}"] = "Sant Cugat del Valles"

    otros = {
        "08930": "Sant Adria de Besos", "08940": "Cornella de Llobregat",
        "08950": "Esplugues de Llobregat", "08960": "Sant Just Desvern",
        "08970": "Sant Joan Despi", "08980": "Sant Feliu de Llobregat",
        "08820": "El Prat de Llobregat", "08830": "Sant Boi de Llobregat",
        "08840": "Viladecans", "08850": "Gava", "08860": "Castelldefels",
        "08800": "Vilanova i la Geltru", "08810": "Sant Pere de Ribes",
        "08870": "Sitges", "08880": "Cubelles", "08191": "Rubi",
        "08290": "Cerdanyola del Valles", "08291": "Ripollet",
        "08210": "Barbera del Valles", "08211": "Castellar del Valles",
        "08100": "Mollet del Valles", "08110": "Montcada i Reixac",
        "08120": "La Llagosta", "08130": "Santa Perpetua de Mogoda",
        "08320": "El Masnou", "08330": "Premia de Mar",
        "08340": "Vilassar de Mar", "08350": "Arenys de Mar",
        "08370": "Calella", "08380": "Malgrat de Mar",
        "08390": "Montgat", "08397": "Pineda de Mar",
        "08440": "Cardedeu", "08500": "Vic", "08520": "Les Franqueses del Valles",
        "08600": "Berga", "08620": "Sant Vicenc dels Horts",
        "08630": "Abrera", "08640": "Olesa de Montserrat",
        "08700": "Igualada", "08720": "Vilafranca del Penedes",
        "08740": "Sant Andreu de la Barca", "08750": "Molins de Rei",
        "08758": "Cervello", "08760": "Martorell", "08780": "Palleja"
    }
    cps.update(otros)
    return cps


BCN_CP = _build_bcn_cp_dict()
CIUDADES_BCN = sorted(set(BCN_CP.values()))

MONTH_NAMES = {
    1: "01_Enero", 2: "02_Febrero", 3: "03_Marzo", 4: "04_Abril",
    5: "05_Mayo", 6: "06_Junio", 7: "07_Julio", 8: "08_Agosto",
    9: "09_Septiembre", 10: "10_Octubre", 11: "11_Noviembre", 12: "12_Diciembre"
}

NACIONALIDADES = [
    "Espanola", "Marroqui", "Pakistani", "Colombiana", "Peruana",
    "Hondurena", "Ecuatoriana", "Argentina", "Venezolana", "Rumana",
    "Italiana", "China", "India", "Senegalesa", "Otra"
]

MARCAS = [
    "Bosch", "Balay", "Siemens", "LG", "Samsung", "Beko", "Haier",
    "Whirlpool", "Indesit", "Candy", "Hisense", "Fagor", "Teka",
    "AEG", "Electrolux", "Corbero", "Benavent", "Cecotec", "Otra"
]


def _clean_pdf(text):
    if text is None:
        return ""
    text = str(text).replace("€", "EUR").replace("–", "-").replace("—", "-")
    return text.encode("latin-1", "replace").decode("latin-1")


def _open_file(path):
    if sys.platform == "win32":
        os.startfile(path)
    else:
        messagebox.showinfo("Archivo", path)


def _load_company_info(app_dir):
    cfg_path = os.path.join(app_dir, "config.json")
    default = {
        "company_name": "NOMBRE DE TU EMPRESA S.L.",
        "nif": "B00000000",
        "phone": "+34 600 000 000",
        "stores": ["L'Hospitalet de Llobregat, Barcelona"]
    }
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                default.update(data)
        except Exception:
            pass
    return default


def _generate_contract_pdf(app_dir, contract_num, data):
    company = _load_company_info(app_dir)
    store_addr = company["stores"][0] if company.get("stores") else ""

    try:
        dt = datetime.datetime.strptime(data["fecha"], "%d/%m/%Y").date()
    except ValueError:
        dt = datetime.date.today()

    target_dir = os.path.join(app_dir, "Contratos_Compraventa", str(dt.year), MONTH_NAMES[dt.month])
    os.makedirs(target_dir, exist_ok=True)
    filepath = os.path.join(target_dir, f"{contract_num}.pdf")

    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    # Encabezado minimalista en blanco y negro
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 8, "CONTRATO Y RECIBO DE COMPRAVENTA DE BIEN USADO", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 6, _clean_pdf(f"N. Documento: {contract_num}   |   Fecha: {data['fecha']}"), align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(6)

    # Bloque 1: Vendedor (Particular)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 7, "1. DATOS DEL VENDEDOR (PARTICULAR)", border="B", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)
    pdf.set_font("Helvetica", "", 10)
    vendedor_lines = [
        f"Nombre y Apellidos: {data['vendedor'] or '________________________________________'}",
        f"DNI / NIE / Pasaporte: {data['dni'] or '__________________'}    Nacionalidad: {data['nacionalidad'] or '______________'}",
        f"Direccion: {data['direccion'] or '__________________________________________________'}",
        f"Codigo Postal: {data['cp'] or '________'}    Ciudad: {data['ciudad'] or '______________________'}"
    ]
    for line in vendedor_lines:
        pdf.set_x(10)
        pdf.multi_cell(0, 6, _clean_pdf(line), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)

    # Bloque 2: Comprador (Empresa)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 7, "2. DATOS DEL COMPRADOR (EMPRESA)", border="B", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)
    pdf.set_font("Helvetica", "", 10)
    comprador_lines = [
        f"Empresa: {company['company_name']}    NIF/CIF: {company['nif']}",
        f"Domicilio: {store_addr}",
        f"Telefono: {company.get('phone', '')}"
    ]
    for line in comprador_lines:
        pdf.set_x(10)
        pdf.multi_cell(0, 6, _clean_pdf(line), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)

    # Bloque 3: Objeto y Precio
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 7, "3. ARTICULO OBJETO DE COMPRAVENTA", border="B", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(50, 8, "Marca", border=1)
    pdf.cell(100, 8, "Aparato / Modelo", border=1)
    pdf.cell(40, 8, "Precio Acordado", border=1, align="R", new_x="LMARGIN", new_y="NEXT")

    pdf.set_font("Helvetica", "", 10)
    pdf.cell(50, 8, _clean_pdf(data["marca"] or "-"), border=1)
    pdf.cell(100, 8, _clean_pdf(data["modelo"][:52]), border=1)
    pdf.cell(40, 8, f"{data['precio']:.2f} EUR", border=1, align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(6)

    # Bloque 4: Clausulas
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 7, "4. DECLARACION Y CONDICIONES", border="B", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)
    pdf.set_font("Helvetica", "", 9)
    clausulas = [
        "PRIMERA.- El VENDEDOR declara bajo su responsabilidad ser el legitimo propietario del articulo descrito, que este es de procedencia licita y se encuentra libre de cargas, embargos o reservas de dominio.",
        "SEGUNDA.- El VENDEDOR entrega en este acto el articulo al COMPRADOR, transmitiendo su plena propiedad, y declara haber recibido el importe integro arriba indicado, sirviendo el presente documento como carta de pago y recibo.",
        "TERCERA.- Al tratarse de una compra a un particular para su posterior comercializacion, la operacion queda registrada a efectos del Regimen Especial de Bienes Usados (REBU) y normativa vigente."
    ]
    for cl in clausulas:
        pdf.set_x(10)
        pdf.multi_cell(0, 5, _clean_pdf(cl), new_x="LMARGIN", new_y="NEXT")
        pdf.ln(2)

    # Bloque 5: Firmas
    pdf.ln(14)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(95, 6, "FIRMA DEL VENDEDOR (Particular)", align="C")
    pdf.cell(95, 6, "FIRMA Y SELLO DEL COMPRADOR", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(18)
    pdf.set_font("Helvetica", "", 9)
    pdf.cell(95, 6, "____________________________________", align="C")
    pdf.cell(95, 6, "____________________________________", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(95, 5, _clean_pdf(f"DNI/NIE: {data['dni'] or ''}"), align="C")
    pdf.cell(95, 5, _clean_pdf(company["company_name"]), align="C", new_x="LMARGIN", new_y="NEXT")

    pdf.output(filepath)
    return filepath


def build_particulares_tab(parent, db, app_dir):
    excel_file = os.path.join(app_dir, "compras_particulares.xlsx")

    db.conn.execute("""
        CREATE TABLE IF NOT EXISTS private_purchases (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT,
            vendedor TEXT,
            dni TEXT,
            nacionalidad TEXT,
            direccion TEXT,
            cp TEXT,
            ciudad TEXT,
            marca TEXT,
            modelo TEXT,
            precio REAL,
            pdf_path TEXT
        )
    """)
    # Migracion segura si la tabla ya existia sin pdf_path
    try:
        db.conn.execute("ALTER TABLE private_purchases ADD COLUMN pdf_path TEXT")
    except sqlite3.OperationalError:
        pass
    db.conn.commit()

    # --- Formulario ---
    frm = ttk.LabelFrame(parent, text="Registrar Compra a Particular (Genera Contrato PDF + Excel)", padding=10)
    frm.pack(fill=tk.X, pady=(0, 10))

    v_fecha = tk.StringVar(value=datetime.date.today().strftime("%d/%m/%Y"))
    v_vendedor = tk.StringVar()
    v_dni = tk.StringVar()
    v_nac = tk.StringVar(value="Espanola")
    v_dir = tk.StringVar()
    v_cp = tk.StringVar()
    v_ciudad = tk.StringVar()
    v_marca = tk.StringVar()
    v_modelo = tk.StringVar()
    v_precio = tk.StringVar()

    ttk.Label(frm, text="Fecha:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=3)
    ttk.Entry(frm, textvariable=v_fecha, width=12).grid(row=0, column=1, sticky=tk.W, padx=4)

    ttk.Label(frm, text="Nombre:").grid(row=0, column=2, sticky=tk.W, padx=6)
    ttk.Entry(frm, textvariable=v_vendedor, width=22).grid(row=0, column=3, padx=4)

    ttk.Label(frm, text="DNI / NIE:").grid(row=0, column=4, sticky=tk.W, padx=6)
    ttk.Entry(frm, textvariable=v_dni, width=15).grid(row=0, column=5, padx=4)

    ttk.Label(frm, text="Nacionalidad:").grid(row=0, column=6, sticky=tk.W, padx=6)
    ttk.Combobox(frm, textvariable=v_nac, values=NACIONALIDADES, width=15).grid(row=0, column=7, padx=4)

    ttk.Label(frm, text="Direccion:").grid(row=1, column=0, sticky=tk.W, padx=4, pady=4)
    ttk.Entry(frm, textvariable=v_dir, width=42).grid(row=1, column=1, columnspan=3, sticky=tk.EW, padx=4)

    ttk.Label(frm, text="Codigo Postal:").grid(row=1, column=4, sticky=tk.W, padx=6)
    cb_cp = ttk.Combobox(frm, textvariable=v_cp, values=sorted(BCN_CP.keys()), width=13)
    cb_cp.grid(row=1, column=5, padx=4)

    ttk.Label(frm, text="Ciudad:").grid(row=1, column=6, sticky=tk.W, padx=6)
    cb_ciudad = ttk.Combobox(frm, textvariable=v_ciudad, values=CIUDADES_BCN, width=22)
    cb_ciudad.grid(row=1, column=7, padx=4)

    def on_cp_change(event=None):
        code = v_cp.get().strip()
        if code in BCN_CP:
            v_ciudad.set(BCN_CP[code])

    def on_city_select(event=None):
        city = v_ciudad.get().strip()
        for cp_k, c_val in BCN_CP.items():
            if c_val == city:
                v_cp.set(cp_k)
                break

    cb_cp.bind("<KeyRelease>", on_cp_change)
    cb_cp.bind("<<ComboboxSelected>>", on_cp_change)
    cb_ciudad.bind("<<ComboboxSelected>>", on_city_select)

    ttk.Label(frm, text="Marca:").grid(row=2, column=0, sticky=tk.W, padx=4, pady=4)
    ttk.Combobox(frm, textvariable=v_marca, values=MARCAS, width=18).grid(row=2, column=1, sticky=tk.W, padx=4)

    ttk.Label(frm, text="Aparato / Modelo:").grid(row=2, column=2, sticky=tk.W, padx=6)
    ttk.Entry(frm, textvariable=v_modelo, width=32).grid(row=2, column=3, columnspan=2, sticky=tk.EW, padx=4)

    ttk.Label(frm, text="Precio (EUR):").grid(row=2, column=5, sticky=tk.E, padx=6)
    ttk.Entry(frm, textvariable=v_precio, width=12).grid(row=2, column=6, sticky=tk.W, padx=4)

    # --- Tabla ---
    cols = ("id", "fecha", "vendedor", "dni", "nac", "dir", "cp", "ciudad", "marca", "modelo", "precio")
    tree = ttk.Treeview(parent, columns=cols, show="headings", height=16)
    for c, txt, w in [
        ("id", "ID", 35), ("fecha", "Fecha", 80), ("vendedor", "Vendedor", 120),
        ("dni", "DNI/NIE", 85), ("nac", "Nacionalidad", 90), ("dir", "Direccion", 150),
        ("cp", "C.P.", 55), ("ciudad", "Ciudad", 125), ("marca", "Marca", 85),
        ("modelo", "Modelo / Aparato", 150), ("precio", "Precio", 75)
    ]:
        tree.heading(c, text=txt)
        tree.column(c, width=w)
    tree.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

    def sync_excel():
        wb = Workbook()
        ws = wb.active
        ws.title = "Compras Particulares"
        headers = ["ID", "Fecha", "Vendedor", "DNI/NIE", "Nacionalidad", "Direccion", "CP", "Ciudad", "Marca", "Modelo", "Precio (EUR)", "Contrato PDF"]
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")

        rows = db.conn.execute("SELECT * FROM private_purchases ORDER BY id DESC").fetchall()
        for r in rows:
            ws.append([
                r["id"], r["fecha"], r["vendedor"], r["dni"], r["nacionalidad"],
                r["direccion"], r["cp"], r["ciudad"], r["marca"], r["modelo"], r["precio"],
                r["pdf_path"] or ""
            ])

        for col in ws.columns:
            max_len = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 35)
        try:
            wb.save(excel_file)
        except PermissionError:
            messagebox.showwarning("Excel abierto", "Cierra 'compras_particulares.xlsx' para actualizar el archivo Excel.")

    def refresh_table():
        for r in tree.get_children():
            tree.delete(r)
        for r in db.conn.execute("SELECT * FROM private_purchases ORDER BY id DESC").fetchall():
            tree.insert("", tk.END, values=(
                r["id"], r["fecha"], r["vendedor"], r["dni"], r["nacionalidad"],
                r["direccion"], r["cp"], r["ciudad"], r["marca"], r["modelo"], f"{r['precio']:.2f} EUR"
            ))

    def add_record():
        try:
            precio = float(v_precio.get().strip().replace(",", "."))
        except ValueError:
            messagebox.showerror("Error", "Introduce un precio valido.")
            return
        if not v_modelo.get().strip():
            messagebox.showwarning("Falta modelo", "Indica el aparato o modelo.")
            return

        fecha_str = v_fecha.get().strip()
        try:
            dt = datetime.datetime.strptime(fecha_str, "%d/%m/%Y").date()
        except ValueError:
            dt = datetime.date.today()

        seq = db.get_next_number(dt.year, "COMP")
        contract_num = f"COMP-{dt.year}-{seq:04d}"

        record_data = {
            "fecha": fecha_str,
            "vendedor": v_vendedor.get().strip(),
            "dni": v_dni.get().strip(),
            "nacionalidad": v_nac.get().strip(),
            "direccion": v_dir.get().strip(),
            "cp": v_cp.get().strip(),
            "ciudad": v_ciudad.get().strip(),
            "marca": v_marca.get().strip(),
            "modelo": v_modelo.get().strip(),
            "precio": round(precio, 2)
        }

        pdf_path = _generate_contract_pdf(app_dir, contract_num, record_data)

        db.conn.execute("""
            INSERT INTO private_purchases (fecha, vendedor, dni, nacionalidad, direccion, cp, ciudad, marca, modelo, precio, pdf_path)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            record_data["fecha"], record_data["vendedor"], record_data["dni"],
            record_data["nacionalidad"], record_data["direccion"], record_data["cp"],
            record_data["ciudad"], record_data["marca"], record_data["modelo"],
            record_data["precio"], pdf_path
        ))
        db.conn.commit()
        db.commit_number(dt.year, "COMP", seq)

        for v in (v_vendedor, v_dni, v_dir, v_cp, v_ciudad, v_marca, v_modelo, v_precio):
            v.set("")
        refresh_table()
        sync_excel()
        _open_file(pdf_path)

    def open_selected_pdf():
        sel = tree.selection()
        if not sel:
            return
        rid = tree.item(sel[0])["values"][0]
        row = db.conn.execute("SELECT pdf_path FROM private_purchases WHERE id=?", (rid,)).fetchone()
        if row and row["pdf_path"] and os.path.exists(row["pdf_path"]):
            _open_file(row["pdf_path"])
        else:
            messagebox.showwarning("Sin PDF", "Este registro no tiene un contrato PDF guardado.")

    def delete_record():
        sel = tree.selection()
        if not sel:
            return
        rid = tree.item(sel[0])["values"][0]
        db.conn.execute("DELETE FROM private_purchases WHERE id=?", (rid,))
        db.conn.commit()
        refresh_table()
        sync_excel()

    ttk.Button(frm, text="+ Guardar + Contrato PDF", command=add_record).grid(row=2, column=7, padx=6, ipadx=6)

    bot = ttk.Frame(parent)
    bot.pack(fill=tk.X)
    ttk.Button(bot, text="Eliminar Seleccionado", command=delete_record).pack(side=tk.LEFT, padx=(0, 8))
    ttk.Button(bot, text="Abrir Contrato PDF", command=open_selected_pdf).pack(side=tk.LEFT)
    ttk.Button(bot, text="Abrir Excel (compras_particulares.xlsx)", command=lambda: _open_file(excel_file) if os.path.exists(excel_file) else messagebox.showinfo("Vacio", "Guarda un registro primero.")).pack(side=tk.RIGHT)

    refresh_table()