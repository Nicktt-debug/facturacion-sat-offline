import os
import re
import sys
import csv
import json
import shutil
import sqlite3
import datetime
import urllib.parse
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from fpdf import FPDF
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from tab_particulares import build_particulares_tab
from tab_agenda import build_agenda_tab, HORARIOS, VEHICULOS_DEF, TECNICOS_DEF
from tab_extra import build_extra_tab
# ==========================================
# RUTAS LOCALES (Compatible con .py y .exe)
# ==========================================
if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

DB_FILE = os.path.join(APP_DIR, "gestion.db")
CONFIG_FILE = os.path.join(APP_DIR, "config.json")
LOGO_FILE = os.path.join(APP_DIR, "logo.png")
BACKUP_DIR = os.path.join(APP_DIR, "Backups")

DEFAULT_CONFIG = {
    "company_name": "NOMBRE DE TU EMPRESA S.L.",
    "nif": "B00000000",
    "phone": "+34 600 000 000",
    "email": "info@tuempresa.com",
    "stores": [
        "Tienda 1 - direccion 1 de local",
        "Tienda 2 - direccion 2 de local",
        "Tienda 3 - "
    ],


# cambiar categorias para la facturas segun tu necessidad 
    "categories": [
        "Lavadora", "Nevera", "Congelador", "Secadora",
        "Lavavajillas", "Calentador", "Aire Acondicionado",
        "Horno / Vitroceramica", "Microondas",
        "Mano de Obra / Reparacion", "Instalacion",
        "Desplazamiento / Revision", "Recambio / Pieza",
        "Otros"
    ]
}

TAX_OPTIONS = {
    "Sin IVA": 0.00,
    "21% IVA": 0.21,
    "4% REBU": 0.04,
    "0% Exento": 0.00
}

EXPENSE_CATEGORIES = [
    "Mercaderia / Electrodomesticos",
    "Recambios / Piezas",
    "Luz / Agua / Gas / Internet",
    "Alquiler Local",
    "Nominas / Seguridad Social",
    "Combustible / Furgonetas / Taller",
    "Gestoria / Seguros / Otros"
]

PAYMENT_METHODS = ["Efectivo", "Tarjeta (TPV)", "Bizum", "Transferencia"]
PAYMENT_STATUSES = ["Pagado", "Pago Parcial", "Pendiente"]
DOC_TYPES = ["Factura", "Factura Reparacion", "Presupuesto", "Albaran Entrega", "Recibo Alquiler"]
OUTPUT_MODES = ["Generar + Imprimir A4", "Generar PDF A4", "Solo Ticket (80mm)", "Solo Guardar (Sin PDF)"]

WARRANTY_OPTIONS = {
    "6 meses (Por defecto ventas)": ("Garantia comercial de 6 meses desde la fecha de compra.", 180),
    "3 meses (Reparacion)": ("Garantia de 3 meses sobre la reparacion realizada y piezas sustituidas.", 90),
    "2 meses (Reparacion / Ajuste)": ("Garantia de 2 meses sobre la reparacion o intervencion realizada.", 60),
    "1 ano (Segunda Mano / REBU)": ("Garantia de 1 ano para bienes de segunda mano. Operacion acogida al Regimen Especial de Bienes Usados (REBU).", 365),
    "3 anos (Producto Nuevo)": ("Garantia legal de 3 anos para productos nuevos segun RDL 1/2007.", 1095),
    "Sin garantia / Alquiler": ("", 0)
}

SAT_STATUSES = ["En taller", "Presupuestado", "Reparado", "Entregado"]
SAT_TECHNICIANS = ["Tecnico 1", "Tecnico 2", "Tecnico 3", "Sin asignar"]

MONTH_NAMES = {
    1: "01_Enero", 2: "02_Febrero", 3: "03_Marzo", 4: "04_Abril",
    5: "05_Mayo", 6: "06_Junio", 7: "07_Julio", 8: "08_Agosto",
    9: "09_Septiembre", 10: "10_Octubre", 11: "11_Noviembre", 12: "12_Diciembre"
}


# ==========================================
# UTILIDADES
# ==========================================
def clean_pdf_text(text):
    if text is None:
        return ""
    text = str(text).replace("€", "EUR").replace("–", "-").replace("—", "-")
    return text.encode("latin-1", "replace").decode("latin-1")


def parse_date(date_str):
    try:
        return datetime.datetime.strptime(date_str.strip(), "%d/%m/%Y").date()
    except ValueError:
        return datetime.date.today()


def parse_date_opt(date_str):
    try:
        return datetime.datetime.strptime(date_str.strip(), "%d/%m/%Y").date()
    except ValueError:
        return None


def get_quarter(dt):
    return f"{(dt.month - 1) // 3 + 1}T"


def open_path(path):
    if sys.platform == "win32":
        os.startfile(path)
    else:
        messagebox.showinfo("Ruta", path)


def print_pdf_file(path):
    if sys.platform == "win32":
        try:
            os.startfile(path, "print")
        except Exception:
            os.startfile(path)
    else:
        messagebox.showinfo("Imprimir", f"PDF listo para imprimir:\n{path}")


def open_whatsapp(phone, message):
    digits = re.sub(r"\D", "", phone or "")
    if not digits:
        messagebox.showwarning("Sin telefono", "Este registro no tiene un numero de telefono valido.")
        return
    if len(digits) == 9:
        digits = "34" + digits
    url = f"https://wa.me/{digits}?text={urllib.parse.quote(message)}"
    webbrowser.open(url)


def calc_warranty_status(doc_date_str, warranty_key):
    if not warranty_key:
        return "-"
    if warranty_key not in WARRANTY_OPTIONS:
        return warranty_key[:20]
    _, days = WARRANTY_OPTIONS[warranty_key]
    if days <= 0:
        return "-"
    dt = parse_date(doc_date_str)
    exp = dt + datetime.timedelta(days=days)
    today = datetime.date.today()
    if today <= exp:
        left = (exp - today).days
        return f"EN GARANTIA ({left}d)"
    return f"CADUCADA ({exp.strftime('%d/%m/%Y')})"


# ==========================================
# BASE DE DATOS SQLITE
# ==========================================
class Database:
    def __init__(self, db_path):
        self.db_path = db_path
        self.conn = sqlite3.connect(db_path)
        self.conn.row_factory = sqlite3.Row
        self._init_tables()

    def reconnect(self):
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row
        self._init_tables()

    def _init_tables(self):
        cur = self.conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS counters (
                year INTEGER, prefix TEXT, last_num INTEGER,
                PRIMARY KEY (year, prefix)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS clients (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE, nif TEXT, phone TEXT, address TEXT, zip TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS suppliers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE, nif TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                doc_type TEXT, doc_num TEXT, doc_date TEXT, year INTEGER, quarter TEXT,
                store TEXT, client_name TEXT, client_nif TEXT, client_phone TEXT,
                payment_method TEXT, payment_status TEXT, warranty TEXT, notes TEXT,
                base_amount REAL, tax_amount REAL, discount_amount REAL, total_amount REAL,
                items_json TEXT, pdf_path TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS repairs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ticket_num TEXT, entry_date TEXT, store TEXT,
                client_name TEXT, client_phone TEXT, appliance TEXT, brand_model TEXT,
                issue TEXT, estimated_price REAL, status TEXT, pdf_path TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS purchases (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                inv_date TEXT, year INTEGER, quarter TEXT, supplier TEXT, inv_num TEXT,
                category TEXT, payment_method TEXT, base_amount REAL, tax_rate REAL,
                tax_amount REAL, total_amount REAL, notes TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS agenda (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT, hora TEXT, duracion INTEGER DEFAULT 1,
                tipo TEXT, vehiculo TEXT, tecnico TEXT,
                cliente TEXT, telefono TEXT, direccion TEXT, detalle TEXT, estado TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS private_purchases (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT, vendedor TEXT, dni TEXT, nacionalidad TEXT,
                direccion TEXT, cp TEXT, ciudad TEXT, marca TEXT, modelo TEXT,
                precio REAL, pdf_path TEXT
            )
        """)
        # Migraciones seguras para bases de datos existentes
        migrations = [
            "ALTER TABLE documents ADD COLUMN paid_amount REAL DEFAULT 0",
            "ALTER TABLE documents ADD COLUMN pending_amount REAL DEFAULT 0",
            "ALTER TABLE repairs ADD COLUMN status_date TEXT",
            "ALTER TABLE repairs ADD COLUMN technician TEXT DEFAULT 'Sin asignar'",
            "ALTER TABLE repairs ADD COLUMN client_nif TEXT DEFAULT ''",
            "ALTER TABLE repairs ADD COLUMN client_address TEXT DEFAULT ''",
            "ALTER TABLE purchases ADD COLUMN supplier_nif TEXT DEFAULT ''",
            "ALTER TABLE purchases ADD COLUMN pdf_path TEXT DEFAULT ''"
        ]
        for sql in migrations:
            try:
                cur.execute(sql)
            except sqlite3.OperationalError:
                pass
        self.conn.commit()

    def create_backup(self):
        os.makedirs(BACKUP_DIR, exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = os.path.join(BACKUP_DIR, f"gestion_backup_{stamp}.db")
        self.conn.commit()
        shutil.copy2(self.db_path, dest)
        backups = sorted(
            [os.path.join(BACKUP_DIR, f) for f in os.listdir(BACKUP_DIR) if f.endswith(".db")],
            key=os.path.getmtime
        )
        for old in backups[:-20]:
            try:
                os.remove(old)
            except OSError:
                pass
        return dest

    def get_next_number(self, year, prefix):
        row = self.conn.execute(
            "SELECT last_num FROM counters WHERE year=? AND prefix=?", (year, prefix)
        ).fetchone()
        return (row["last_num"] + 1) if row else 1

    def commit_number(self, year, prefix, num):
        self.conn.execute("""
            INSERT INTO counters (year, prefix, last_num) VALUES (?, ?, ?)
            ON CONFLICT(year, prefix) DO UPDATE SET last_num=excluded.last_num
        """, (year, prefix, num))
        self.conn.commit()

    def save_client(self, name, nif, phone, address, zip_code):
        if not name.strip():
            return
        self.conn.execute("""
            INSERT INTO clients (name, nif, phone, address, zip) VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(name) DO UPDATE SET
                nif=excluded.nif, phone=excluded.phone,
                address=excluded.address, zip=excluded.zip
        """, (name.strip(), nif.strip(), phone.strip(), address.strip(), zip_code.strip()))
        self.conn.commit()

    def get_clients(self):
        return self.conn.execute("SELECT * FROM clients ORDER BY name").fetchall()

    def save_supplier(self, name, nif):
        if not name.strip():
            return
        self.conn.execute("""
            INSERT INTO suppliers (name, nif) VALUES (?, ?)
            ON CONFLICT(name) DO UPDATE SET nif=excluded.nif
        """, (name.strip(), nif.strip()))
        self.conn.commit()

    def get_suppliers(self):
        return self.conn.execute("SELECT * FROM suppliers ORDER BY name").fetchall()


# ==========================================
# APLICACION PRINCIPAL
# ==========================================
class ERPApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestion Integral - Facturas, SAT, Proveedores, Agenda e Impuestos")
        self.root.geometry("1220x840")

        self.db = Database(DB_FILE)
        self.config = self._load_config()
        self.items = []
        self.editing_item_idx = None

        self._build_top_bar()
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.tab_billing = ttk.Frame(self.notebook, padding=10)
        self.tab_sat = ttk.Frame(self.notebook, padding=10)
        self.tab_purchases = ttk.Frame(self.notebook, padding=10)
        self.tab_fiscal = ttk.Frame(self.notebook, padding=10)
        self.tab_history = ttk.Frame(self.notebook, padding=10)
        self.tab_particulares = ttk.Frame(self.notebook, padding=10)
        self.tab_agenda = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(self.tab_billing, text="1. Facturas / Tickets")
        self.notebook.add(self.tab_sat, text="2. SAT / Reparaciones")
        self.notebook.add(self.tab_purchases, text="3. Compras y Gastos")
        self.notebook.add(self.tab_fiscal, text="4. Fiscal, TPV y Caja")
        self.notebook.add(self.tab_history, text="5. Buscador / Abonos")
        self.notebook.add(self.tab_particulares, text="6. Compras Particulares")
        self.notebook.add(self.tab_agenda, text="7. Agenda 11h-21h / Vehiculos")

        self.tab_extra = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_extra, text="8. Stock / Caja Chica / RMA")
        build_extra_tab(self.tab_extra, self.db, APP_DIR, on_sell_stock=self.stock_to_invoice)

        self._build_billing_tab()
        self._build_sat_tab()
        self._build_purchases_tab()
        self._build_fiscal_tab()
        self._build_history_tab()
        build_particulares_tab(self.tab_particulares, self.db, APP_DIR)
        self.agenda_tab = build_agenda_tab(self.tab_agenda, self.db, on_invoice=self.agenda_to_invoice)

        self._refresh_all()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        try:
            self.db.create_backup()
        except Exception:
            pass
        self.db.conn.close()
        self.root.destroy()

    def _refresh_all(self):
        self._refresh_clients_combo()
        self._refresh_suppliers_combo()
        self._refresh_sat_table()
        self._refresh_purchases_table()
        self._refresh_fiscal_summary()
        self._refresh_history_table()
        self._refresh_alerts()

    # ==========================================
    # CONFIGURACION, BACKUP Y ALERTAS
    # ==========================================
    def _load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                if "Alquiler Mensual" in cfg.get("categories", []):
                    cfg["categories"].remove("Alquiler Mensual")
                return cfg
            except Exception:
                pass
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, indent=2, ensure_ascii=False)
        return dict(DEFAULT_CONFIG)

    def _save_config(self):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.config, f, indent=2, ensure_ascii=False)

    def _build_top_bar(self):
        bar = ttk.Frame(self.root, padding=(10, 8))
        bar.pack(fill=tk.X)

        self.lbl_company = ttk.Label(
            bar, text=f"{self.config['company_name']}  |  NIF: {self.config['nif']}",
            font=("Segoe UI", 11, "bold")
        )
        self.lbl_company.pack(side=tk.LEFT)

        self.lbl_alerts = ttk.Label(bar, text="", font=("Segoe UI", 9, "bold"), foreground="#b3261e")
        self.lbl_alerts.pack(side=tk.LEFT, padx=16)

        ttk.Button(bar, text="Configuracion", command=self.open_settings_window).pack(side=tk.RIGHT, padx=3)
        ttk.Button(bar, text="Restaurar Backup", command=self.restore_backup).pack(side=tk.RIGHT, padx=3)
        ttk.Button(bar, text="Copia Seguridad", command=self.manual_backup).pack(side=tk.RIGHT, padx=3)
        ttk.Button(bar, text="Carpeta App", command=lambda: open_path(APP_DIR)).pack(side=tk.RIGHT, padx=3)

    def _refresh_alerts(self):
        pend_row = self.db.conn.execute("""
            SELECT COUNT(*), COALESCE(SUM(CASE WHEN pending_amount > 0 THEN pending_amount ELSE total_amount END), 0)
            FROM documents
            WHERE payment_status IN ('Pendiente', 'Pago Parcial')
              AND doc_type NOT IN ('Presupuesto', 'Albaran Entrega')
        """).fetchone()

        sat_rows = self.db.conn.execute(
            "SELECT entry_date, status_date FROM repairs WHERE status='Reparado'"
        ).fetchall()
        today = datetime.date.today()
        old_repairs = sum(1 for r in sat_rows if (today - parse_date(r["status_date"] or r["entry_date"])).days >= 7)

        parts = []
        if pend_row[0] > 0:
            parts.append(f"Pendientes cobro: {pend_row[0]} ({pend_row[1]:.2f} EUR)")
        if old_repairs > 0:
            parts.append(f"SAT Reparados >7d: {old_repairs}")
        self.lbl_alerts.config(text="  |  ".join(parts) if parts else "Todo al dia")

    def manual_backup(self):
        dest = self.db.create_backup()
        messagebox.showinfo("Copia de Seguridad", f"Copia guardada en:\n{dest}")

    def restore_backup(self):
        path = filedialog.askopenfilename(
            initialdir=BACKUP_DIR if os.path.exists(BACKUP_DIR) else APP_DIR,
            title="Seleccionar copia de seguridad (.db)",
            filetypes=[("Base de datos SQLite", "*.db")]
        )
        if not path:
            return
        if not messagebox.askyesno("Confirmar Restauracion", "Esto reemplazara los datos actuales con la copia seleccionada. Continuar?"):
            return
        try:
            self.db.conn.close()
            shutil.copy2(path, DB_FILE)
            self.db.reconnect()
            self.agenda_tab.reload()
            self._refresh_all()
            messagebox.showinfo("Restaurado", "Todos los datos se han recuperado correctamente desde el backup.")
        except Exception as e:
            self.db.reconnect()
            messagebox.showerror("Error al restaurar", str(e))

    def open_settings_window(self):
        win = tk.Toplevel(self.root)
        win.title("Configuracion de Empresa y Tiendas")
        win.geometry("560x520")
        frm = ttk.Frame(win, padding=15)
        frm.pack(fill=tk.BOTH, expand=True)

        v_name = tk.StringVar(value=self.config["company_name"])
        v_nif = tk.StringVar(value=self.config["nif"])
        v_phone = tk.StringVar(value=self.config["phone"])
        v_email = tk.StringVar(value=self.config["email"])

        for r, (lbl, var) in enumerate([
            ("Nombre Empresa:", v_name), ("NIF / CIF:", v_nif),
            ("Telefono:", v_phone), ("Email:", v_email)
        ]):
            ttk.Label(frm, text=lbl).grid(row=r, column=0, sticky=tk.W, pady=4)
            ttk.Entry(frm, textvariable=var, width=52).grid(row=r, column=1, pady=4, padx=5)

        ttk.Label(frm, text="Tiendas (1 por linea):").grid(row=4, column=0, sticky=tk.NW, pady=6)
        txt_stores = tk.Text(frm, width=40, height=5)
        txt_stores.insert("1.0", "\n".join(self.config["stores"]))
        txt_stores.grid(row=4, column=1, sticky=tk.EW, pady=6, padx=5)

        ttk.Label(frm, text="Categorias (1 por linea):").grid(row=5, column=0, sticky=tk.NW, pady=6)
        txt_cats = tk.Text(frm, width=40, height=8)
        txt_cats.insert("1.0", "\n".join(self.config["categories"]))
        txt_cats.grid(row=5, column=1, sticky=tk.EW, pady=6, padx=5)

        def save_and_close():
            self.config["company_name"] = v_name.get().strip() or DEFAULT_CONFIG["company_name"]
            self.config["nif"] = v_nif.get().strip()
            self.config["phone"] = v_phone.get().strip()
            self.config["email"] = v_email.get().strip()
            stores = [s.strip() for s in txt_stores.get("1.0", tk.END).splitlines() if s.strip()]
            cats = [c.strip() for c in txt_cats.get("1.0", tk.END).splitlines() if c.strip()]
            if stores:
                self.config["stores"] = stores
            if cats:
                self.config["categories"] = cats
            self._save_config()
            self.lbl_company.config(text=f"{self.config['company_name']}  |  NIF: {self.config['nif']}")
            self.store_combo["values"] = self.config["stores"]
            self.sat_store_combo["values"] = self.config["stores"]
            self.cat_combo["values"] = self.config["categories"]
            self.fiscal_store_combo["values"] = ["Todas"] + self.config["stores"]
            self.fiscal_top_store_combo["values"] = ["Todas"] + self.config["stores"]
            messagebox.showinfo("Guardado", "Configuracion actualizada.")
            win.destroy()

        ttk.Button(frm, text="Guardar Cambios", command=save_and_close).grid(row=7, column=0, columnspan=2, pady=10, ipadx=15)

    # ==========================================
    # PESTANA 1: FACTURAS / PRESUPUESTOS / TICKETS
    # ==========================================
    def _get_prefix(self, doc_type):
        return {
            "Factura": "FAC", "Factura Reparacion": "FREP",
            "Presupuesto": "PRE", "Albaran Entrega": "ALB",
            "Recibo Alquiler": "ALQ", "Factura Rectificativa": "RECT"
        }.get(doc_type, "FAC")

    def _build_billing_tab(self):
        top = ttk.LabelFrame(self.tab_billing, text="Datos del Documento y Pago", padding=8)
        top.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(top, text="Tipo:").grid(row=0, column=0, sticky=tk.W, padx=4)
        self.doc_type_var = tk.StringVar(value="Factura")
        dt_cb = ttk.Combobox(top, textvariable=self.doc_type_var, values=DOC_TYPES, state="readonly", width=17)
        dt_cb.grid(row=0, column=1, padx=4)
        dt_cb.bind("<<ComboboxSelected>>", lambda e: self._update_doc_number())

        ttk.Label(top, text="Numero:").grid(row=0, column=2, sticky=tk.W, padx=6)
        self.doc_num_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.doc_num_var, width=16).grid(row=0, column=3, padx=4)

        ttk.Label(top, text="Fecha:").grid(row=0, column=4, sticky=tk.W, padx=6)
        self.doc_date_var = tk.StringVar(value=datetime.date.today().strftime("%d/%m/%Y"))
        ttk.Entry(top, textvariable=self.doc_date_var, width=11).grid(row=0, column=5, padx=4)

        self.btn_convert_inv = ttk.Button(top, text="Convertir a Factura", command=self.convert_quote_to_invoice)
        self.btn_convert_inv.grid(row=0, column=6, padx=6)

        ttk.Button(top, text="Asignar Hora Entrega (Agenda)", command=self.open_assign_delivery_dialog).grid(row=0, column=7, padx=6)

        ttk.Label(top, text="Tienda:").grid(row=1, column=0, sticky=tk.W, padx=4, pady=(6, 0))
        self.store_var = tk.StringVar(value=self.config["stores"][0])
        self.store_combo = ttk.Combobox(top, textvariable=self.store_var, values=self.config["stores"], state="readonly", width=42)
        self.store_combo.grid(row=1, column=1, columnspan=3, sticky=tk.W, padx=4, pady=(6, 0))

        ttk.Label(top, text="Pago:").grid(row=1, column=4, sticky=tk.W, padx=6, pady=(6, 0))
        self.pay_method_var = tk.StringVar(value=PAYMENT_METHODS[0])
        ttk.Combobox(top, textvariable=self.pay_method_var, values=PAYMENT_METHODS, state="readonly", width=13).grid(row=1, column=5, padx=4, pady=(6, 0))

        self.pay_status_var = tk.StringVar(value=PAYMENT_STATUSES[0])
        cb_ps = ttk.Combobox(top, textvariable=self.pay_status_var, values=PAYMENT_STATUSES, state="readonly", width=12)
        cb_ps.grid(row=1, column=6, padx=4, pady=(6, 0))
        cb_ps.bind("<<ComboboxSelected>>", lambda e: self._on_pay_status_change())

        # Pago parcial (cantidad entregada)
        pf = ttk.Frame(top)
        pf.grid(row=1, column=7, sticky=tk.W, padx=4, pady=(6, 0))
        ttk.Label(pf, text="Entregado:").pack(side=tk.LEFT)
        self.paid_amt_var = tk.StringVar(value="0")
        self.ent_paid_amt = ttk.Entry(pf, textvariable=self.paid_amt_var, width=8)
        self.ent_paid_amt.pack(side=tk.LEFT, padx=3)
        self.ent_paid_amt.bind("<KeyRelease>", lambda e: self._refresh_items_table())
        self.lbl_pending_amt = ttk.Label(pf, text="Resta: 0.00", foreground="#b3261e", font=("Segoe UI", 8, "bold"))
        self.lbl_pending_amt.pack(side=tk.LEFT, padx=3)

        # Cliente
        cli = ttk.LabelFrame(self.tab_billing, text="Cliente (Opcional - Selecciona guardado o escribe nuevo)", padding=8)
        cli.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(cli, text="Guardados:").grid(row=0, column=0, sticky=tk.W, padx=4)
        self.client_select_var = tk.StringVar()
        self.client_combo = ttk.Combobox(cli, textvariable=self.client_select_var, width=28)
        self.client_combo.grid(row=0, column=1, padx=4)
        self.client_combo.bind("<<ComboboxSelected>>", self._on_client_selected)

        self.c_name = tk.StringVar()
        self.c_nif = tk.StringVar()
        self.c_phone = tk.StringVar()
        self.c_addr = tk.StringVar()
        self.c_zip = tk.StringVar()

        ttk.Label(cli, text="Nombre:").grid(row=0, column=2, sticky=tk.W, padx=6)
        ttk.Entry(cli, textvariable=self.c_name, width=28).grid(row=0, column=3, padx=4)
        ttk.Label(cli, text="DNI/NIF:").grid(row=0, column=4, sticky=tk.W, padx=6)
        ttk.Entry(cli, textvariable=self.c_nif, width=16).grid(row=0, column=5, padx=4)

        ttk.Label(cli, text="Direccion:").grid(row=1, column=0, sticky=tk.W, padx=4, pady=(4, 0))
        ttk.Entry(cli, textvariable=self.c_addr, width=48).grid(row=1, column=1, columnspan=2, sticky=tk.EW, padx=4, pady=(4, 0))
        ttk.Label(cli, text="C.P.:").grid(row=1, column=3, sticky=tk.E, padx=4, pady=(4, 0))
        ttk.Entry(cli, textvariable=self.c_zip, width=10).grid(row=1, column=4, sticky=tk.W, padx=4, pady=(4, 0))
        ttk.Label(cli, text="Tel:").grid(row=1, column=4, sticky=tk.E, padx=4, pady=(4, 0))
        ttk.Entry(cli, textvariable=self.c_phone, width=16).grid(row=1, column=5, padx=4, pady=(4, 0))

        # Producto
        prod = ttk.LabelFrame(self.tab_billing, text="Agregar / Editar Producto (Doble clic en fila para editar)", padding=8)
        prod.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(prod, text="Categoria:").grid(row=0, column=0, sticky=tk.W, padx=4)
        self.cat_var = tk.StringVar(value=self.config["categories"][0])
        self.cat_combo = ttk.Combobox(prod, textvariable=self.cat_var, values=self.config["categories"], state="readonly", width=22)
        self.cat_combo.grid(row=0, column=1, padx=4)

        ttk.Label(prod, text="Modelo / Detalle:").grid(row=0, column=2, sticky=tk.W, padx=6)
        self.desc_var = tk.StringVar()
        ttk.Entry(prod, textvariable=self.desc_var, width=34).grid(row=0, column=3, padx=4)

        ttk.Label(prod, text="Cant.:").grid(row=0, column=4, sticky=tk.W, padx=6)
        self.qty_var = tk.StringVar(value="1")
        ttk.Entry(prod, textvariable=self.qty_var, width=6).grid(row=0, column=5, padx=4)

        ttk.Label(prod, text="Precio (EUR):").grid(row=1, column=0, sticky=tk.W, padx=4, pady=(5, 0))
        self.price_var = tk.StringVar()
        ttk.Entry(prod, textvariable=self.price_var, width=14).grid(row=1, column=1, sticky=tk.W, padx=4, pady=(5, 0))

        self.price_mode_var = tk.StringVar(value="gross")
        mf = ttk.Frame(prod)
        mf.grid(row=1, column=2, columnspan=2, sticky=tk.W, padx=6, pady=(5, 0))
        ttk.Radiobutton(mf, text="Bruto (IVA inc.)", variable=self.price_mode_var, value="gross").pack(side=tk.LEFT, padx=3)
        ttk.Radiobutton(mf, text="Neto (Sin IVA)", variable=self.price_mode_var, value="net").pack(side=tk.LEFT, padx=3)

        ttk.Label(prod, text="Impuesto:").grid(row=1, column=4, sticky=tk.W, padx=6, pady=(5, 0))
        self.tax_var = tk.StringVar(value="Sin IVA")
        ttk.Combobox(prod, textvariable=self.tax_var, values=list(TAX_OPTIONS.keys()), state="readonly", width=11).grid(row=1, column=5, padx=4, pady=(5, 0))

        self.btn_add_item = ttk.Button(prod, text="+ Anadir", command=self.add_or_update_item)
        self.btn_add_item.grid(row=1, column=6, padx=8, pady=(5, 0))

        # Tabla productos
        tf = ttk.Frame(self.tab_billing)
        tf.pack(fill=tk.BOTH, expand=True, pady=(0, 6))
        cols = ("producto", "cant", "neto", "impuesto", "total")
        self.tree_items = ttk.Treeview(tf, columns=cols, show="headings", height=6)
        for c, txt, w, anc in [
            ("producto", "Producto / Concepto", 420, tk.W),
            ("cant", "Cant.", 60, tk.CENTER),
            ("neto", "Base Imp.", 115, tk.E),
            ("impuesto", "IVA / REBU", 145, tk.E),
            ("total", "Total", 115, tk.E)
        ]:
            self.tree_items.heading(c, text=txt)
            self.tree_items.column(c, width=w, anchor=anc)
        self.tree_items.bind("<Double-1>", self._on_item_double_click)
        sb = ttk.Scrollbar(tf, orient=tk.VERTICAL, command=self.tree_items.yview)
        self.tree_items.configure(yscrollcommand=sb.set)
        self.tree_items.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        # Descuento, Garantia, Notas
        extra = ttk.Frame(self.tab_billing)
        extra.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(extra, text="Descuento:").pack(side=tk.LEFT, padx=(0, 4))
        self.disc_val_var = tk.StringVar(value="0")
        ttk.Entry(extra, textvariable=self.disc_val_var, width=7).pack(side=tk.LEFT)
        self.disc_type_var = tk.StringVar(value="EUR")
        ttk.Combobox(extra, textvariable=self.disc_type_var, values=["EUR", "%"], state="readonly", width=5).pack(side=tk.LEFT, padx=4)
        ttk.Button(extra, text="Aplicar", command=self._refresh_items_table).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(extra, text="Garantia (editable):").pack(side=tk.LEFT, padx=(0, 4))
        self.warranty_var = tk.StringVar(value="6 meses (Por defecto ventas)")
        ttk.Combobox(extra, textvariable=self.warranty_var, values=list(WARRANTY_OPTIONS.keys()), width=28).pack(side=tk.LEFT, padx=4)

        ttk.Label(extra, text="Notas PDF:").pack(side=tk.LEFT, padx=(10, 4))
        self.notes_var = tk.StringVar()
        ttk.Entry(extra, textvariable=self.notes_var, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Totales y controles
        bot = ttk.Frame(self.tab_billing)
        bot.pack(fill=tk.X)
        ttk.Button(bot, text="Eliminar Linea", command=self.remove_selected_item).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(bot, text="Limpiar Formulario", command=self.clear_billing_form).pack(side=tk.LEFT)
        self.lbl_totals = ttk.Label(bot, text="Base: 0.00 EUR | Imp: 0.00 EUR | Dto: 0.00 EUR | TOTAL: 0.00 EUR", font=("Segoe UI", 10, "bold"))
        self.lbl_totals.pack(side=tk.RIGHT, padx=10)

        # Modo de salida (Por defecto: Generar + Imprimir A4)
        out_frm = ttk.Frame(self.tab_billing)
        out_frm.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(out_frm, text="Accion al guardar:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=4)
        self.output_mode_var = tk.StringVar(value=OUTPUT_MODES[0])
        ttk.Combobox(out_frm, textvariable=self.output_mode_var, values=OUTPUT_MODES, state="readonly", width=24).pack(side=tk.LEFT, padx=6)
        ttk.Button(out_frm, text="Ejecutar y Guardar Registro", command=self.generate_document_pdf).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5, padx=(8, 0))

        self._on_pay_status_change()
        self._update_doc_number()

    def _on_pay_status_change(self):
        st = self.pay_status_var.get()
        if st == "Pago Parcial":
            self.ent_paid_amt.state(["!disabled"])
        else:
            self.ent_paid_amt.state(["disabled"])
        self._refresh_items_table()

    def _get_payment_split(self, grand_total):
        st = self.pay_status_var.get()
        if st == "Pagado":
            return round(grand_total, 2), 0.0
        if st == "Pendiente":
            return 0.0, round(grand_total, 2)
        try:
            paid = float(self.paid_amt_var.get().strip().replace(",", "."))
        except ValueError:
            paid = 0.0
        paid = max(0.0, min(paid, grand_total))
        return round(paid, 2), round(grand_total - paid, 2)

    def _update_doc_number(self):
        doc_type = self.doc_type_var.get()
        dt = parse_date(self.doc_date_var.get())
        prefix = self._get_prefix(doc_type)
        seq = self.db.get_next_number(dt.year, prefix)
        self.doc_num_var.set(f"{prefix}-{dt.year}-{seq:04d}")
        if doc_type == "Factura Reparacion":
            self.warranty_var.set("3 meses (Reparacion)")
        elif doc_type == "Recibo Alquiler":
            self.warranty_var.set("Sin garantia / Alquiler")
        elif doc_type == "Factura":
            self.warranty_var.set("6 meses (Por defecto ventas)")

        if doc_type in ("Presupuesto", "Albaran Entrega"):
            self.btn_convert_inv.state(["!disabled"])
        else:
            self.btn_convert_inv.state(["disabled"])

    def convert_quote_to_invoice(self):
        self.doc_type_var.set("Factura")
        self._update_doc_number()
        messagebox.showinfo("Convertido", "Tipo cambiado a Factura manteniendo cliente y productos.")

    def _refresh_clients_combo(self):
        self._clients_cache = {row["name"]: row for row in self.db.get_clients()}
        self.client_combo["values"] = list(self._clients_cache.keys())

    def _on_client_selected(self, event=None):
        row = self._clients_cache.get(self.client_select_var.get())
        if row:
            self.c_name.set(row["name"])
            self.c_nif.set(row["nif"])
            self.c_phone.set(row["phone"])
            self.c_addr.set(row["address"])
            self.c_zip.set(row["zip"])

    def add_or_update_item(self):
        try:
            qty = int(self.qty_var.get().strip())
            unit_price = float(self.price_var.get().strip().replace(",", "."))
            if qty <= 0 or unit_price < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Revisa la cantidad y el precio.")
            return

        cat = self.cat_var.get()
        desc = self.desc_var.get().strip()
        full_name = f"{cat} - {desc}" if desc else cat
        tax_label = self.tax_var.get()
        tax_rate = TAX_OPTIONS[tax_label]
        mode = self.price_mode_var.get()

        line_val = unit_price * qty
        if mode == "net":
            net = line_val
            tax_amt = net * tax_rate
            gross = net + tax_amt
        else:
            gross = line_val
            net = gross / (1 + tax_rate) if tax_rate > 0 else gross
            tax_amt = gross - net

        item = {
            "cat": cat, "desc": desc, "name": full_name, "qty": qty,
            "unit_price": unit_price, "mode": mode, "tax_label": tax_label,
            "net": round(net, 2), "tax_amount": round(tax_amt, 2), "total": round(gross, 2)
        }

        if self.editing_item_idx is not None:
            self.items[self.editing_item_idx] = item
            self.editing_item_idx = None
            self.btn_add_item.config(text="+ Anadir")
        else:
            self.items.append(item)

        self._refresh_items_table()
        self.desc_var.set("")
        self.price_var.set("")
        self.qty_var.set("1")

    def _on_item_double_click(self, event):
        sel = self.tree_items.selection()
        if not sel:
            return
        idx = self.tree_items.index(sel[0])
        it = self.items[idx]
        self.editing_item_idx = idx
        self.cat_var.set(it["cat"])
        self.desc_var.set(it["desc"])
        self.qty_var.set(str(it["qty"]))
        self.price_var.set(str(it["unit_price"]))
        self.price_mode_var.set(it["mode"])
        self.tax_var.set(it["tax_label"])
        self.btn_add_item.config(text="Guardar Cambio")

    def remove_selected_item(self):
        for sel in reversed(self.tree_items.selection()):
            del self.items[self.tree_items.index(sel)]
        self.editing_item_idx = None
        self.btn_add_item.config(text="+ Anadir")
        self._refresh_items_table()

    def clear_billing_form(self):
        self.items.clear()
        self.editing_item_idx = None
        self.btn_add_item.config(text="+ Anadir")
        for v in (self.client_select_var, self.c_name, self.c_nif, self.c_phone, self.c_addr, self.c_zip, self.notes_var):
            v.set("")
        self.disc_val_var.set("0")
        self.paid_amt_var.set("0")
        self.pay_status_var.set("Pagado")
        self._on_pay_status_change()

    def _calc_totals(self):
        sum_net = sum(i["net"] for i in self.items)
        sum_tax = sum(i["tax_amount"] for i in self.items)
        sum_gross = sum(i["total"] for i in self.items)
        try:
            d_val = float(self.disc_val_var.get().strip().replace(",", "."))
        except ValueError:
            d_val = 0.0
        discount = round(sum_gross * (d_val / 100.0), 2) if self.disc_type_var.get() == "%" else round(d_val, 2)
        discount = min(discount, sum_gross)
        ratio = (sum_gross - discount) / sum_gross if sum_gross > 0 else 1.0
        return round(sum_net * ratio, 2), round(sum_tax * ratio, 2), discount, round(sum_gross - discount, 2)

    def _refresh_items_table(self):
        for r in self.tree_items.get_children():
            self.tree_items.delete(r)
        for it in self.items:
            imp_txt = "-" if it["tax_label"] == "Sin IVA" else f"{it['tax_amount']:.2f} ({it['tax_label'].split()[0]})"
            self.tree_items.insert("", tk.END, values=(
                it["name"], it["qty"], f"{it['net']:.2f} EUR", imp_txt, f"{it['total']:.2f} EUR"
            ))
        base, tax, disc, total = self._calc_totals()
        paid, pend = self._get_payment_split(total)
        self.lbl_pending_amt.config(text=f"Resta: {pend:.2f} EUR")
        self.lbl_totals.config(text=f"Base: {base:.2f} | Imp: {tax:.2f} | Dto: {disc:.2f} | TOTAL: {total:.2f} EUR")

    def open_assign_delivery_dialog(self, preset_detail=""):
        win = tk.Toplevel(self.root)
        win.title("Asignar Hora de Entrega (Sincronizado con Agenda)")
        win.geometry("460x320")
        win.transient(self.root)
        frm = ttk.Frame(win, padding=14)
        frm.pack(fill=tk.BOTH, expand=True)

        v_f = tk.StringVar(value=datetime.date.today().strftime("%d/%m/%Y"))
        v_h = tk.StringVar(value=HORARIOS[0])
        v_veh = tk.StringVar(value=VEHICULOS_DEF[0])
        v_tec = tk.StringVar(value=TECNICOS_DEF[0])
        v_dir = tk.StringVar(value=f"{self.c_addr.get().strip()} {self.c_zip.get().strip()}".strip())
        det_auto = preset_detail or ", ".join(f"{it['qty']}x {it['name']}" for it in self.items[:3])
        v_det = tk.StringVar(value=det_auto)

        rows = [
            ("Fecha (DD/MM/AAAA):", ttk.Entry(frm, textvariable=v_f, width=16)),
            ("Hora (11:00-21:00):", ttk.Combobox(frm, textvariable=v_h, values=HORARIOS, state="readonly", width=14)),
            ("Vehiculo:", ttk.Combobox(frm, textvariable=v_veh, values=VEHICULOS_DEF, state="readonly", width=26)),
            ("Tecnico / Repartidor:", ttk.Combobox(frm, textvariable=v_tec, values=TECNICOS_DEF, state="readonly", width=26)),
            ("Direccion:", ttk.Entry(frm, textvariable=v_dir, width=36)),
            ("Detalle / Aparato:", ttk.Entry(frm, textvariable=v_det, width=36)),
        ]
        for i, (lbl, w) in enumerate(rows):
            ttk.Label(frm, text=lbl).grid(row=i, column=0, sticky=tk.W, pady=5)
            w.grid(row=i, column=1, sticky=tk.W, padx=6, pady=5)

        def save_delivery():
            dt = parse_date_opt(v_f.get())
            if dt is None:
                messagebox.showwarning("Fecha invalida", "Usa el formato DD/MM/AAAA.", parent=win)
                return
            data = {
                "fecha": dt.strftime("%Y-%m-%d"),
                "hora": v_h.get(),
                "duracion": 1,
                "tipo": "Entrega Venta",
                "vehiculo": v_veh.get(),
                "tecnico": v_tec.get(),
                "cliente": self.c_name.get().strip() or "Cliente Tienda",
                "telefono": self.c_phone.get().strip(),
                "direccion": v_dir.get().strip(),
                "detalle": v_det.get().strip()
            }
            self.agenda_tab.repo.insert(data)
            self.agenda_tab.reload()
            messagebox.showinfo("Agenda Actualizada", f"Entrega asignada el {v_f.get()} a las {v_h.get()} ({v_veh.get()}).", parent=win)
            win.destroy()

        ttk.Button(frm, text="Guardar en Agenda", command=save_delivery).grid(row=6, column=0, columnspan=2, sticky=tk.EW, pady=(12, 0), ipady=4)

    def _render_ticket_80mm(self, filepath, doc_type, doc_num, doc_date, store, items, grand_total, pay_method, pay_status, paid_amt, pend_amt, warranty_key):
        pdf = FPDF(unit="mm", format=(80, 155 + len(items) * 6))
        pdf.set_margins(4, 4, 4)
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(0, 5, clean_pdf_text(self.config["company_name"]), align="C", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(0, 4, clean_pdf_text(f"NIF: {self.config['nif']} | Tel: {self.config['phone']}"), align="C", new_x="LMARGIN", new_y="NEXT")
        pdf.cell(0, 4, clean_pdf_text(store[:38]), align="C", new_x="LMARGIN", new_y="NEXT")
        pdf.cell(0, 4, "-" * 42, align="C", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(0, 5, clean_pdf_text(f"{doc_type.upper()} {doc_num}  ({doc_date})"), new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 8)
        for it in items:
            pdf.cell(48, 5, clean_pdf_text(f"{it['qty']}x {it['name'][:24]}"))
            pdf.cell(0, 5, f"{it['total']:.2f} EUR", align="R", new_x="LMARGIN", new_y="NEXT")
        pdf.cell(0, 4, "-" * 42, align="C", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 6, f"TOTAL: {grand_total:.2f} EUR ({pay_method})", align="R", new_x="LMARGIN", new_y="NEXT")
        if pay_status == "Pago Parcial":
            pdf.set_font("Helvetica", "B", 8)
            pdf.cell(0, 4, f"Pagado: {paid_amt:.2f} EUR | Pendiente: {pend_amt:.2f} EUR", align="R", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 7)
        pdf.set_x(4)
        pdf.multi_cell(0, 4, clean_pdf_text(f"Garantia: {warranty_key}"), new_x="LMARGIN", new_y="NEXT")
        pdf.output(filepath)

    def _render_document_pdf(self, filepath, doc_type, doc_num, doc_date, quarter, store,
                             c_name, c_nif, c_addr, c_zip, c_phone,
                             pay_method, pay_status, paid_amt, pend_amt, warranty_key, notes,
                             items, base_amt, tax_amt, disc_amt, grand_total):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)

        if os.path.exists(LOGO_FILE):
            try:
                pdf.image(LOGO_FILE, x=10, y=8, w=28)
                pdf.set_x(42)
            except Exception:
                pass

        pdf.set_font("Helvetica", "B", 16)
        pdf.cell(95, 8, clean_pdf_text(doc_type.upper()), border=0)
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 6, clean_pdf_text(f"N.: {doc_num}"), align="R", new_x="LMARGIN", new_y="NEXT")
        pdf.cell(0, 5, clean_pdf_text(f"Fecha: {doc_date} ({quarter})"), align="R", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(8)

        y_start = pdf.get_y()
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(95, 5, clean_pdf_text(self.config["company_name"]), new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 9)
        for line in [f"NIF: {self.config['nif']}", store, f"Tel: {self.config['phone']}", self.config["email"]]:
            if line.strip():
                pdf.cell(95, 5, clean_pdf_text(line), new_x="LMARGIN", new_y="NEXT")
        y_left = pdf.get_y()

        pdf.set_xy(110, y_start)
        c_lines = []
        if c_name: c_lines.append(("B", c_name))
        if c_nif: c_lines.append(("", f"DNI/NIF: {c_nif}"))
        if c_addr: c_lines.append(("", c_addr))
        if c_zip: c_lines.append(("", f"C.P.: {c_zip}"))
        if c_phone: c_lines.append(("", f"Tel: {c_phone}"))

        if c_lines:
            pdf.set_font("Helvetica", "B", 9)
            pdf.cell(90, 5, "CLIENTE / ENTREGA:", new_x="LMARGIN", new_y="NEXT")
            for st, txt in c_lines:
                pdf.set_x(110)
                pdf.set_font("Helvetica", st, 9)
                pdf.cell(90, 5, clean_pdf_text(txt), new_x="LMARGIN", new_y="NEXT")
        pdf.set_y(max(y_left, pdf.get_y()) + 8)

        is_albaran = (doc_type == "Albaran Entrega")
        hide_tax = all(it.get("tax_label") == "Sin IVA" for it in items)

        pdf.set_font("Helvetica", "B", 9)
        pdf.set_fill_color(235, 235, 235)
        if is_albaran:
            pdf.cell(160, 7, "Articulo / Concepto a Entregar", border=1, fill=True)
            pdf.cell(30, 7, "Cant.", border=1, align="C", fill=True, new_x="LMARGIN", new_y="NEXT")
        elif hide_tax:
            pdf.cell(115, 7, "Concepto", border=1, fill=True)
            pdf.cell(15, 7, "Cant.", border=1, align="C", fill=True)
            pdf.cell(30, 7, "Precio Unit.", border=1, align="R", fill=True)
            pdf.cell(30, 7, "Total", border=1, align="R", fill=True, new_x="LMARGIN", new_y="NEXT")
        else:
            pdf.cell(85, 7, "Concepto", border=1, fill=True)
            pdf.cell(15, 7, "Cant.", border=1, align="C", fill=True)
            pdf.cell(30, 7, "Base Imp.", border=1, align="R", fill=True)
            pdf.cell(30, 7, "Impuesto", border=1, align="R", fill=True)
            pdf.cell(30, 7, "Total", border=1, align="R", fill=True, new_x="LMARGIN", new_y="NEXT")

        pdf.set_font("Helvetica", "", 9)
        has_rebu = False
        for it in items:
            if "REBU" in it.get("tax_label", ""):
                has_rebu = True
            if is_albaran:
                pdf.cell(160, 6, clean_pdf_text(it["name"][:80]), border=1)
                pdf.cell(30, 6, str(it["qty"]), border=1, align="C", new_x="LMARGIN", new_y="NEXT")
            elif hide_tax:
                pdf.cell(115, 6, clean_pdf_text(it["name"][:62]), border=1)
                pdf.cell(15, 6, str(it["qty"]), border=1, align="C")
                pdf.cell(30, 6, f"{it.get('unit_price', it['total']):.2f} EUR", border=1, align="R")
                pdf.cell(30, 6, f"{it['total']:.2f} EUR", border=1, align="R", new_x="LMARGIN", new_y="NEXT")
            else:
                pdf.cell(85, 6, clean_pdf_text(it["name"][:46]), border=1)
                pdf.cell(15, 6, str(it["qty"]), border=1, align="C")
                pdf.cell(30, 6, f"{it['net']:.2f} EUR", border=1, align="R")
                pdf.cell(30, 6, f"{it['tax_amount']:.2f} ({it['tax_label'].split()[0]})", border=1, align="R")
                pdf.cell(30, 6, f"{it['total']:.2f} EUR", border=1, align="R", new_x="LMARGIN", new_y="NEXT")

        pdf.ln(6)
        if is_albaran:
            pdf.set_font("Helvetica", "B", 11)
            if pay_status == "Pago Parcial":
                estado_cobro = f"ENTREGADO: {paid_amt:.2f} EUR  |  PENDIENTE EN ENTREGA: {pend_amt:.2f} EUR"
            elif pay_status == "Pendiente":
                estado_cobro = f"PENDIENTE DE COBRO EN ENTREGA: {grand_total:.2f} EUR"
            else:
                estado_cobro = "PAGADO EN TIENDA (NO COBRAR)"
            pdf.cell(0, 8, clean_pdf_text(estado_cobro), border=1, align="C", new_x="LMARGIN", new_y="NEXT")
            pdf.ln(14)
            pdf.set_font("Helvetica", "", 9)
            pdf.cell(95, 6, "Firma Repartidor: __________________________")
            pdf.cell(95, 6, "Firma Cliente (Conforme): __________________________", new_x="LMARGIN", new_y="NEXT")
        else:
            pdf.set_font("Helvetica", "", 9)
            if disc_amt != 0:
                pdf.cell(130, 6, "")
                pdf.cell(30, 6, "Descuento:", align="R")
                pdf.cell(30, 6, f"-{disc_amt:.2f} EUR", align="R", new_x="LMARGIN", new_y="NEXT")
            if not hide_tax:
                pdf.cell(130, 6, "")
                pdf.cell(30, 6, "Base Imponible:", align="R")
                pdf.cell(30, 6, f"{base_amt:.2f} EUR", align="R", new_x="LMARGIN", new_y="NEXT")
                pdf.cell(130, 6, "")
                pdf.cell(30, 6, "Total Impuestos:", align="R")
                pdf.cell(30, 6, f"{tax_amt:.2f} EUR", align="R", new_x="LMARGIN", new_y="NEXT")
            pdf.set_font("Helvetica", "B", 11)
            pdf.cell(130, 7, clean_pdf_text(f"Forma de pago: {pay_method} ({pay_status})"))
            pdf.cell(30, 7, "TOTAL:", border="T", align="R")
            pdf.cell(30, 7, f"{grand_total:.2f} EUR", border="T", align="R", new_x="LMARGIN", new_y="NEXT")
            if pay_status == "Pago Parcial":
                pdf.set_font("Helvetica", "B", 9)
                pdf.cell(130, 6, "")
                pdf.cell(30, 6, "Entregado:", align="R")
                pdf.cell(30, 6, f"{paid_amt:.2f} EUR", align="R", new_x="LMARGIN", new_y="NEXT")
                pdf.cell(130, 6, "")
                pdf.cell(30, 6, "Pendiente:", align="R")
                pdf.cell(30, 6, f"{pend_amt:.2f} EUR", align="R", new_x="LMARGIN", new_y="NEXT")

        pdf.ln(8)
        pdf.set_font("Helvetica", "", 8)
        if notes:
            pdf.set_x(10)
            pdf.multi_cell(0, 5, clean_pdf_text(f"Observaciones: {notes}"), new_x="LMARGIN", new_y="NEXT")
            pdf.ln(2)
        if not is_albaran:
            if warranty_key in WARRANTY_OPTIONS:
                w_text, _ = WARRANTY_OPTIONS[warranty_key]
            else:
                w_text = f"Garantia: {warranty_key}" if warranty_key.strip() else ""
            if w_text:
                pdf.set_x(10)
                pdf.multi_cell(0, 4, clean_pdf_text(w_text), new_x="LMARGIN", new_y="NEXT")
            if has_rebu and "REBU" not in w_text:
                pdf.set_x(10)
                pdf.multi_cell(0, 4, "Operacion acogida al Regimen Especial de Bienes Usados (REBU) - IVA incluido en el margen.", new_x="LMARGIN", new_y="NEXT")
        if doc_type == "Presupuesto":
            pdf.ln(2)
            pdf.set_font("Helvetica", "I", 8)
            pdf.cell(0, 5, "Este presupuesto tiene una validez de 15 dias y no sustituye a una factura.", new_x="LMARGIN", new_y="NEXT")

        pdf.output(filepath)

    def generate_document_pdf(self):
        if not self.items:
            messagebox.showwarning("Vacio", "Anade al menos un producto o concepto.")
            return

        doc_type = self.doc_type_var.get()
        doc_num = self.doc_num_var.get().strip()
        doc_date = self.doc_date_var.get().strip()
        dt = parse_date(doc_date)
        quarter = get_quarter(dt)
        store = self.store_var.get()
        out_mode = self.output_mode_var.get()

        folder_map = {
            "Factura": "Facturas", "Factura Reparacion": "Facturas_Reparacion",
            "Presupuesto": "Presupuestos", "Albaran Entrega": "Albaranes",
            "Recibo Alquiler": "Recibos_Alquiler"
        }
        base_amt, tax_amt, disc_amt, grand_total = self._calc_totals()
        paid_amt, pend_amt = self._get_payment_split(grand_total)
        safe_num = re.sub(r'[\\/*?:"<>|]', "_", doc_num)
        filepath = ""

        if out_mode == "Solo Ticket (80mm)":
            target_dir = os.path.join(APP_DIR, "Tickets", str(dt.year), MONTH_NAMES[dt.month])
            os.makedirs(target_dir, exist_ok=True)
            filepath = os.path.join(target_dir, f"{safe_num}_ticket.pdf")
            self._render_ticket_80mm(
                filepath, doc_type, doc_num, doc_date, store,
                self.items, grand_total, self.pay_method_var.get(),
                self.pay_status_var.get(), paid_amt, pend_amt, self.warranty_var.get()
            )
        elif out_mode in ("Generar PDF A4", "Generar + Imprimir A4"):
            target_dir = os.path.join(APP_DIR, folder_map[doc_type], str(dt.year), MONTH_NAMES[dt.month])
            os.makedirs(target_dir, exist_ok=True)
            filepath = os.path.join(target_dir, f"{safe_num}.pdf")
            self._render_document_pdf(
                filepath, doc_type, doc_num, doc_date, quarter, store,
                self.c_name.get().strip(), self.c_nif.get().strip(),
                self.c_addr.get().strip(), self.c_zip.get().strip(), self.c_phone.get().strip(),
                self.pay_method_var.get(), self.pay_status_var.get(), paid_amt, pend_amt,
                self.warranty_var.get(), self.notes_var.get().strip(),
                self.items, base_amt, tax_amt, disc_amt, grand_total
            )

        self.db.save_client(self.c_name.get(), self.c_nif.get(), self.c_phone.get(), self.c_addr.get(), self.c_zip.get())
        self.db.conn.execute("""
            INSERT INTO documents (
                doc_type, doc_num, doc_date, year, quarter, store,
                client_name, client_nif, client_phone, payment_method, payment_status,
                paid_amount, pending_amount, warranty, notes,
                base_amount, tax_amount, discount_amount, total_amount, items_json, pdf_path
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            doc_type, doc_num, doc_date, dt.year, quarter, store,
            self.c_name.get().strip(), self.c_nif.get().strip(), self.c_phone.get().strip(),
            self.pay_method_var.get(), self.pay_status_var.get(), paid_amt, pend_amt,
            self.warranty_var.get(), self.notes_var.get().strip(),
            base_amt, tax_amt, disc_amt, grand_total, json.dumps(self.items), filepath
        ))
        self.db.conn.commit()

        prefix = self._get_prefix(doc_type)
        cur_seq = self.db.get_next_number(dt.year, prefix)
        self.db.commit_number(dt.year, prefix, cur_seq)

        self._update_doc_number()
        self._refresh_all()

        if out_mode == "Solo Guardar (Sin PDF)":
            messagebox.showinfo("Guardado", f"{doc_type} {doc_num} registrado sin generar PDF.")
        elif out_mode == "Solo Ticket (80mm)":
            if messagebox.askyesno("Ticket Generado", "Deseas IMPRIMIR el ticket ahora?\n(Pulsa 'No' para solo VERLO en pantalla)"):
                print_pdf_file(filepath)
            else:
                open_path(filepath)
        elif out_mode == "Generar + Imprimir A4":
            print_pdf_file(filepath)
        else:
            open_path(filepath)

        if doc_type in ("Factura", "Albaran Entrega"):
            if messagebox.askyesno("Agenda de Entregas", "Deseas asignar hora de entrega en la Agenda para este pedido?"):
                self.open_assign_delivery_dialog(preset_detail=f"{doc_num} - " + ", ".join(it["name"] for it in self.items[:2]))

    def agenda_to_invoice(self, d):
        tipo = d.get("tipo", "")
        reparacion = tipo in ("Averia / SAT", "Reparacion Domicilio")
        self.clear_billing_form()
        self.doc_type_var.set("Factura Reparacion" if reparacion else "Factura")
        self._update_doc_number()

        nombre = (d.get("cliente") or "").strip()
        cli = getattr(self, "_clients_cache", {}).get(nombre)
        self.c_name.set(nombre)
        self.c_phone.set(d.get("telefono") or (cli["phone"] if cli else "") or "")
        self.c_addr.set(d.get("direccion") or (cli["address"] if cli else "") or "")
        if cli:
            self.c_nif.set(cli["nif"] or "")
            self.c_zip.set(cli["zip"] or "")

        self.notes_var.set(f"Ref. Agenda #{d['id']} ({d.get('fecha', '')}) - {tipo}: {d.get('detalle', '')}".strip())
        cat = {"Averia / SAT": "Mano de Obra / Reparacion",
               "Reparacion Domicilio": "Mano de Obra / Reparacion",
               "Instalacion": "Instalacion"}.get(tipo)
        if cat and cat in self.config["categories"]:
            self.cat_var.set(cat)
        self.desc_var.set(d.get("detalle", ""))
        self.notebook.select(self.tab_billing)

    # ==========================================
    # PESTANA 2: SAT / RESGUARDO (Con NIF y Direccion)
    # ==========================================
    def _build_sat_tab(self):
        frm = ttk.LabelFrame(self.tab_sat, text="Nuevo Resguardo de Entrega / Parte de Reparacion (SAT)", padding=10)
        frm.pack(fill=tk.X, pady=(0, 10))

        self.sat_store_var = tk.StringVar(value=self.config["stores"][0])
        self.sat_client_var = tk.StringVar()
        self.sat_nif_var = tk.StringVar()
        self.sat_phone_var = tk.StringVar()
        self.sat_addr_var = tk.StringVar()
        self.sat_app_var = tk.StringVar(value=self.config["categories"][0])
        self.sat_model_var = tk.StringVar()
        self.sat_issue_var = tk.StringVar()
        self.sat_price_var = tk.StringVar(value="0")
        self.sat_tech_var = tk.StringVar(value=SAT_TECHNICIANS[0])

        ttk.Label(frm, text="Tienda:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=3)
        self.sat_store_combo = ttk.Combobox(frm, textvariable=self.sat_store_var, values=self.config["stores"], state="readonly", width=36)
        self.sat_store_combo.grid(row=0, column=1, columnspan=2, sticky=tk.W, padx=4, pady=3)

        ttk.Label(frm, text="Cliente:").grid(row=0, column=3, sticky=tk.W, padx=5)
        ttk.Entry(frm, textvariable=self.sat_client_var, width=20).grid(row=0, column=4, padx=4)
        ttk.Label(frm, text="DNI/NIF:").grid(row=0, column=5, sticky=tk.W, padx=5)
        ttk.Entry(frm, textvariable=self.sat_nif_var, width=14).grid(row=0, column=6, padx=4)

        ttk.Label(frm, text="Direccion:").grid(row=1, column=0, sticky=tk.W, padx=4, pady=3)
        ttk.Entry(frm, textvariable=self.sat_addr_var, width=40).grid(row=1, column=1, columnspan=2, sticky=tk.EW, padx=4)
        ttk.Label(frm, text="Telefono:").grid(row=1, column=3, sticky=tk.W, padx=5)
        ttk.Entry(frm, textvariable=self.sat_phone_var, width=16).grid(row=1, column=4, sticky=tk.W, padx=4)
        ttk.Label(frm, text="Tecnico:").grid(row=1, column=5, sticky=tk.W, padx=5)
        ttk.Combobox(frm, textvariable=self.sat_tech_var, values=SAT_TECHNICIANS, width=14).grid(row=1, column=6, padx=4)

        ttk.Label(frm, text="Aparato:").grid(row=2, column=0, sticky=tk.W, padx=4, pady=3)
        ttk.Combobox(frm, textvariable=self.sat_app_var, values=self.config["categories"], width=18).grid(row=2, column=1, sticky=tk.W, padx=4)
        ttk.Label(frm, text="Marca / Modelo:").grid(row=2, column=2, sticky=tk.W, padx=5)
        ttk.Entry(frm, textvariable=self.sat_model_var, width=20).grid(row=2, column=3, padx=4)
        ttk.Label(frm, text="Est. EUR:").grid(row=2, column=4, sticky=tk.E, padx=5)
        ttk.Entry(frm, textvariable=self.sat_price_var, width=10).grid(row=2, column=5, sticky=tk.W, padx=4)

        ttk.Label(frm, text="Averia / Notas:").grid(row=3, column=0, sticky=tk.W, padx=4, pady=3)
        ttk.Entry(frm, textvariable=self.sat_issue_var, width=75).grid(row=3, column=1, columnspan=4, sticky=tk.EW, padx=4)
        ttk.Button(frm, text="Crear Resguardo + PDF", command=self.create_sat_ticket).grid(row=3, column=5, columnspan=2, padx=6, ipadx=8)

        cols = ("id", "ticket", "fecha", "tecnico", "cliente", "nif", "tel", "aparato", "averia", "estado", "est")
        self.tree_sat = ttk.Treeview(self.tab_sat, columns=cols, show="headings", height=13)
        for c, txt, w in [
            ("id", "ID", 32), ("ticket", "Parte", 105), ("fecha", "Fecha", 76),
            ("tecnico", "Tecnico", 85), ("cliente", "Cliente", 120), ("nif", "NIF", 80),
            ("tel", "Tel", 85), ("aparato", "Aparato / Modelo", 160), ("averia", "Averia", 170),
            ("estado", "Estado", 90), ("est", "Est. EUR", 68)
        ]:
            self.tree_sat.heading(c, text=txt)
            self.tree_sat.column(c, width=w)
        self.tree_sat.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        act = ttk.Frame(self.tab_sat)
        act.pack(fill=tk.X)
        ttk.Label(act, text="Estado:").pack(side=tk.LEFT, padx=3)
        self.sat_new_status = tk.StringVar(value=SAT_STATUSES[2])
        ttk.Combobox(act, textvariable=self.sat_new_status, values=SAT_STATUSES, state="readonly", width=13).pack(side=tk.LEFT, padx=3)
        ttk.Label(act, text="Tecnico:").pack(side=tk.LEFT, padx=3)
        self.sat_new_tech = tk.StringVar(value=SAT_TECHNICIANS[0])
        ttk.Combobox(act, textvariable=self.sat_new_tech, values=SAT_TECHNICIANS, width=13).pack(side=tk.LEFT, padx=3)
        ttk.Button(act, text="Actualizar", command=self.update_sat_status).pack(side=tk.LEFT, padx=5)
        ttk.Button(act, text="Abrir PDF", command=self.open_sat_pdf).pack(side=tk.LEFT, padx=4)
        ttk.Button(act, text="WhatsApp", command=self.whatsapp_sat_client).pack(side=tk.LEFT, padx=4)
        ttk.Button(act, text="Convertir a Factura Reparacion ->", command=self.convert_sat_to_invoice).pack(side=tk.RIGHT, padx=5)

    def create_sat_ticket(self):
        if not self.sat_client_var.get().strip():
            messagebox.showwarning("Falta cliente", "Introduce al menos el nombre del cliente.")
            return
        dt = datetime.date.today()
        seq = self.db.get_next_number(dt.year, "SAT")
        ticket_num = f"SAT-{dt.year}-{seq:04d}"
        date_str = dt.strftime("%d/%m/%Y")
        try:
            est = float(self.sat_price_var.get().replace(",", "."))
        except ValueError:
            est = 0.0

        target_dir = os.path.join(APP_DIR, "Partes_SAT", str(dt.year), MONTH_NAMES[dt.month])
        os.makedirs(target_dir, exist_ok=True)
        filepath = os.path.join(target_dir, f"{ticket_num}.pdf")

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 15)
        pdf.cell(0, 10, clean_pdf_text(f"RESGUARDO DE ENTREGA A REPARAR - {ticket_num}"), new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 6, clean_pdf_text(f"{self.config['company_name']} | {self.sat_store_var.get()} | Tel: {self.config['phone']}"), new_x="LMARGIN", new_y="NEXT")
        pdf.cell(0, 6, clean_pdf_text(f"Fecha entrada: {date_str}   |   Tecnico asignado: {self.sat_tech_var.get()}"), new_x="LMARGIN", new_y="NEXT")
        pdf.ln(6)
        for lbl, val in [
            ("Cliente:", f"{self.sat_client_var.get()}  (DNI/NIF: {self.sat_nif_var.get() or '-'})  Tel: {self.sat_phone_var.get()}"),
            ("Direccion:", self.sat_addr_var.get() or "-"),
            ("Aparato:", f"{self.sat_app_var.get()} - {self.sat_model_var.get()}"),
            ("Averia reportada:", self.sat_issue_var.get()),
            ("Presupuesto orientativo:", f"{est:.2f} EUR")
        ]:
            pdf.set_x(10)
            pdf.set_font("Helvetica", "B", 10)
            pdf.cell(48, 7, lbl)
            pdf.set_font("Helvetica", "", 10)
            pdf.multi_cell(0, 7, clean_pdf_text(val), new_x="LMARGIN", new_y="NEXT")
        pdf.ln(10)
        pdf.set_x(10)
        pdf.set_font("Helvetica", "I", 8)
        pdf.multi_cell(0, 5, "Imprescindible presentar este resguardo para la retirada del aparato. Pasados 90 dias sin retirada, el taller no se hace responsable del almacenamiento.", new_x="LMARGIN", new_y="NEXT")
        pdf.output(filepath)

        self.db.save_client(self.sat_client_var.get(), self.sat_nif_var.get(), self.sat_phone_var.get(), self.sat_addr_var.get(), "")
        self.db.conn.execute("""
            INSERT INTO repairs (
                ticket_num, entry_date, store, client_name, client_nif, client_phone, client_address,
                appliance, brand_model, issue, estimated_price, status, pdf_path, status_date, technician
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (ticket_num, date_str, self.sat_store_var.get(), self.sat_client_var.get().strip(),
              self.sat_nif_var.get().strip(), self.sat_phone_var.get().strip(), self.sat_addr_var.get().strip(),
              self.sat_app_var.get(), self.sat_model_var.get().strip(), self.sat_issue_var.get().strip(),
              est, "En taller", filepath, date_str, self.sat_tech_var.get().strip()))
        self.db.conn.commit()
        self.db.commit_number(dt.year, "SAT", seq)

        for v in (self.sat_client_var, self.sat_nif_var, self.sat_phone_var, self.sat_addr_var, self.sat_model_var, self.sat_issue_var):
            v.set("")
        self.sat_price_var.set("0")
        self._refresh_all()
        open_path(filepath)

    def _refresh_sat_table(self):
        for r in self.tree_sat.get_children():
            self.tree_sat.delete(r)
        for r in self.db.conn.execute("SELECT * FROM repairs ORDER BY id DESC").fetchall():
            tech = r["technician"] if "technician" in r.keys() and r["technician"] else "Sin asignar"
            nif = r["client_nif"] if "client_nif" in r.keys() and r["client_nif"] else ""
            self.tree_sat.insert("", tk.END, values=(
                r["id"], r["ticket_num"], r["entry_date"], tech, r["client_name"], nif, r["client_phone"],
                f"{r['appliance']} {r['brand_model']}", r["issue"], r["status"], f"{r['estimated_price']:.2f}"
            ))

    def update_sat_status(self):
        sel = self.tree_sat.selection()
        if not sel:
            return
        rep_id = self.tree_sat.item(sel[0])["values"][0]
        today_str = datetime.date.today().strftime("%d/%m/%Y")
        self.db.conn.execute(
            "UPDATE repairs SET status=?, status_date=?, technician=? WHERE id=?",
            (self.sat_new_status.get(), today_str, self.sat_new_tech.get().strip(), rep_id)
        )
        self.db.conn.commit()
        self._refresh_sat_table()
        self._refresh_alerts()

    def open_sat_pdf(self):
        sel = self.tree_sat.selection()
        if not sel:
            return
        rep_id = self.tree_sat.item(sel[0])["values"][0]
        row = self.db.conn.execute("SELECT pdf_path FROM repairs WHERE id=?", (rep_id,)).fetchone()
        if row and row["pdf_path"] and os.path.exists(row["pdf_path"]):
            open_path(row["pdf_path"])

    def whatsapp_sat_client(self):
        sel = self.tree_sat.selection()
        if not sel:
            return
        rep_id = self.tree_sat.item(sel[0])["values"][0]
        r = self.db.conn.execute("SELECT * FROM repairs WHERE id=?", (rep_id,)).fetchone()
        msg = (f"Hola {r['client_name']}, le contactamos de {self.config['company_name']} "
               f"sobre su {r['appliance']} {r['brand_model']} (Parte {r['ticket_num']}). "
               f"Estado actual: {r['status']}.")
        open_whatsapp(r["client_phone"], msg)

    def convert_sat_to_invoice(self):
        sel = self.tree_sat.selection()
        if not sel:
            return
        rep_id = self.tree_sat.item(sel[0])["values"][0]
        r = self.db.conn.execute("SELECT * FROM repairs WHERE id=?", (rep_id,)).fetchone()
        self.clear_billing_form()
        self.doc_type_var.set("Factura Reparacion")
        self._update_doc_number()
        self.store_var.set(r["store"])
        self.c_name.set(r["client_name"])
        self.c_nif.set(r["client_nif"] or "")
        self.c_phone.set(r["client_phone"])
        self.c_addr.set(r["client_address"] or "")
        self.warranty_var.set("3 meses (Reparacion)")
        self.notes_var.set(f"Ref. Parte {r['ticket_num']} - {r['appliance']} {r['brand_model']}")
        if r["estimated_price"] > 0:
            self.cat_var.set("Mano de Obra / Reparacion")
            self.desc_var.set(f"{r['appliance']} {r['brand_model']} ({r['issue']})")
            self.price_var.set(str(r["estimated_price"]))
            self.add_or_update_item()
        self.notebook.select(self.tab_billing)

    # ==========================================
    # PESTANA 3: COMPRAS (NIF, Proveedores guardados, Adjuntar PDF)
    # ==========================================
    def _build_purchases_tab(self):
        frm = ttk.LabelFrame(self.tab_purchases, text="Registrar Compra a Proveedor o Gasto Fijo (IVA Soportado)", padding=10)
        frm.pack(fill=tk.X, pady=(0, 10))

        self.p_date = tk.StringVar(value=datetime.date.today().strftime("%d/%m/%Y"))
        self.p_supplier = tk.StringVar()
        self.p_sup_nif = tk.StringVar()
        self.p_num = tk.StringVar()
        self.p_cat = tk.StringVar(value=EXPENSE_CATEGORIES[0])
        self.p_pay = tk.StringVar(value="Transferencia")
        self.p_amount = tk.StringVar()
        self.p_mode = tk.StringVar(value="gross")
        self.p_tax = tk.StringVar(value="21% IVA")
        self.p_notes = tk.StringVar()
        self.p_pdf_path = tk.StringVar()

        ttk.Label(frm, text="Fecha:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=3)
        ttk.Entry(frm, textvariable=self.p_date, width=12).grid(row=0, column=1, sticky=tk.W, padx=4)

        ttk.Label(frm, text="Proveedor:").grid(row=0, column=2, sticky=tk.W, padx=6)
        self.cb_supplier = ttk.Combobox(frm, textvariable=self.p_supplier, width=26)
        self.cb_supplier.grid(row=0, column=3, padx=4)
        self.cb_supplier.bind("<<ComboboxSelected>>", self._on_supplier_selected)

        ttk.Label(frm, text="NIF Prov.:").grid(row=0, column=4, sticky=tk.W, padx=6)
        ttk.Entry(frm, textvariable=self.p_sup_nif, width=14).grid(row=0, column=5, padx=4)

        ttk.Label(frm, text="N. Factura:").grid(row=0, column=6, sticky=tk.W, padx=6)
        ttk.Entry(frm, textvariable=self.p_num, width=16).grid(row=0, column=7, padx=4)

        ttk.Label(frm, text="Tipo Gasto:").grid(row=1, column=0, sticky=tk.W, padx=4, pady=4)
        ttk.Combobox(frm, textvariable=self.p_cat, values=EXPENSE_CATEGORIES, state="readonly", width=28).grid(row=1, column=1, columnspan=2, sticky=tk.W, padx=4)

        ttk.Label(frm, text="Importe:").grid(row=1, column=3, sticky=tk.E, padx=4)
        ttk.Entry(frm, textvariable=self.p_amount, width=11).grid(row=1, column=4, sticky=tk.W, padx=4)
        ttk.Combobox(frm, textvariable=self.p_tax, values=list(TAX_OPTIONS.keys()), state="readonly", width=11).grid(row=1, column=5, sticky=tk.W, padx=4)
        ttk.Combobox(frm, textvariable=self.p_pay, values=PAYMENT_METHODS, state="readonly", width=14).grid(row=1, column=6, columnspan=2, sticky=tk.W, padx=4)

        pmf = ttk.Frame(frm)
        pmf.grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=4, pady=3)
        ttk.Radiobutton(pmf, text="Total Bruto", variable=self.p_mode, value="gross").pack(side=tk.LEFT)
        ttk.Radiobutton(pmf, text="Base Neta", variable=self.p_mode, value="net").pack(side=tk.LEFT, padx=6)

        ttk.Label(frm, text="Notas:").grid(row=2, column=2, sticky=tk.E, padx=4)
        ttk.Entry(frm, textvariable=self.p_notes, width=28).grid(row=2, column=3, columnspan=2, sticky=tk.EW, padx=4)

        ttk.Button(frm, text="Adjuntar PDF...", command=self.attach_purchase_pdf).grid(row=2, column=5, padx=4)
        self.lbl_p_pdf = ttk.Label(frm, text="Sin PDF", width=14)
        self.lbl_p_pdf.grid(row=2, column=6, sticky=tk.W, padx=2)
        ttk.Button(frm, text="+ Guardar Gasto", command=self.add_purchase).grid(row=2, column=7, padx=4, ipadx=6)

        cols = ("id", "fecha", "trim", "tipo", "proveedor", "nif", "num", "base", "iva", "total", "pago", "pdf")
        self.tree_purch = ttk.Treeview(self.tab_purchases, columns=cols, show="headings", height=14)
        for c, txt, w in [
            ("id", "ID", 32), ("fecha", "Fecha", 78), ("trim", "Trim.", 52),
            ("tipo", "Tipo Gasto", 150), ("proveedor", "Proveedor", 150), ("nif", "NIF", 85),
            ("num", "N. Factura", 100), ("base", "Base", 78), ("iva", "IVA", 70),
            ("total", "Total", 80), ("pago", "Pago", 85), ("pdf", "PDF", 45)
        ]:
            self.tree_purch.heading(c, text=txt)
            self.tree_purch.column(c, width=w)
        self.tree_purch.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        bot = ttk.Frame(self.tab_purchases)
        bot.pack(fill=tk.X)
        ttk.Button(bot, text="Eliminar Seleccionado", command=self.delete_purchase).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(bot, text="Abrir PDF Adjunto", command=self.open_purchase_pdf).pack(side=tk.LEFT)

    def _refresh_suppliers_combo(self):
        self._suppliers_cache = {r["name"]: r["nif"] for r in self.db.get_suppliers()}
        self.cb_supplier["values"] = list(self._suppliers_cache.keys())

    def _on_supplier_selected(self, event=None):
        nif = self._suppliers_cache.get(self.p_supplier.get(), "")
        if nif:
            self.p_sup_nif.set(nif)

    def attach_purchase_pdf(self):
        path = filedialog.askopenfilename(title="Seleccionar Factura PDF del Proveedor", filetypes=[("PDF", "*.pdf")])
        if path:
            self.p_pdf_path.set(path)
            self.lbl_p_pdf.config(text=os.path.basename(path)[:14])

    def add_purchase(self):
        if not self.p_supplier.get().strip():
            messagebox.showwarning("Falta proveedor", "Escribe el nombre del proveedor o concepto.")
            return
        try:
            amt = float(self.p_amount.get().strip().replace(",", "."))
        except ValueError:
            messagebox.showerror("Error", "Importe no valido.")
            return

        dt = parse_date(self.p_date.get())
        rate = TAX_OPTIONS[self.p_tax.get()]
        if self.p_mode.get() == "net":
            base = round(amt, 2)
            tax = round(base * rate, 2)
            total = round(base + tax, 2)
        else:
            total = round(amt, 2)
            base = round(total / (1 + rate), 2) if rate > 0 else total
            tax = round(total - base, 2)

        self.db.save_supplier(self.p_supplier.get(), self.p_sup_nif.get())
        self.db.conn.execute("""
            INSERT INTO purchases (
                inv_date, year, quarter, supplier, supplier_nif, inv_num, category,
                payment_method, base_amount, tax_rate, tax_amount, total_amount, notes, pdf_path
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (dt.strftime("%d/%m/%Y"), dt.year, get_quarter(dt), self.p_supplier.get().strip(),
              self.p_sup_nif.get().strip(), self.p_num.get().strip(), self.p_cat.get(),
              self.p_pay.get(), base, rate, tax, total, self.p_notes.get().strip(), self.p_pdf_path.get().strip()))
        self.db.conn.commit()
        for v in (self.p_supplier, self.p_sup_nif, self.p_num, self.p_amount, self.p_notes, self.p_pdf_path):
            v.set("")
        self.lbl_p_pdf.config(text="Sin PDF")
        self._refresh_all()

    def _refresh_purchases_table(self):
        for r in self.tree_purch.get_children():
            self.tree_purch.delete(r)
        for r in self.db.conn.execute("SELECT * FROM purchases ORDER BY id DESC").fetchall():
            nif = r["supplier_nif"] if "supplier_nif" in r.keys() and r["supplier_nif"] else ""
            has_pdf = "SI" if ("pdf_path" in r.keys() and r["pdf_path"]) else "-"
            self.tree_purch.insert("", tk.END, values=(
                r["id"], r["inv_date"], f"{r['year']}-{r['quarter']}", r["category"],
                r["supplier"], nif, r["inv_num"], f"{r['base_amount']:.2f}", f"{r['tax_amount']:.2f}",
                f"{r['total_amount']:.2f}", r["payment_method"], has_pdf
            ))

    def open_purchase_pdf(self):
        sel = self.tree_purch.selection()
        if not sel:
            return
        pid = self.tree_purch.item(sel[0])["values"][0]
        row = self.db.conn.execute("SELECT pdf_path FROM purchases WHERE id=?", (pid,)).fetchone()
        if row and row["pdf_path"] and os.path.exists(row["pdf_path"]):
            open_path(row["pdf_path"])
        else:
            messagebox.showwarning("Sin PDF", "Esta compra no tiene PDF adjunto o el archivo se ha movido.")

    def delete_purchase(self):
        sel = self.tree_purch.selection()
        if not sel:
            return
        pid = self.tree_purch.item(sel[0])["values"][0]
        self.db.conn.execute("DELETE FROM purchases WHERE id=?", (pid,))
        self.db.conn.commit()
        self._refresh_all()

    # ==========================================
    # PESTANA 4: FISCAL, FILTROS, EXCEL/PDF, TPV Y CAJA
    # ==========================================
    def _build_fiscal_tab(self):
        filt = ttk.Frame(self.tab_fiscal)
        filt.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(filt, text="Ano:").pack(side=tk.LEFT, padx=3)
        self.fiscal_year_var = tk.StringVar(value=str(datetime.date.today().year))
        ttk.Entry(filt, textvariable=self.fiscal_year_var, width=6).pack(side=tk.LEFT, padx=3)

        ttk.Label(filt, text="Tienda:").pack(side=tk.LEFT, padx=4)
        self.fiscal_top_store_var = tk.StringVar(value="Todas")
        self.fiscal_top_store_combo = ttk.Combobox(filt, textvariable=self.fiscal_top_store_var, values=["Todas"] + self.config["stores"], state="readonly", width=26)
        self.fiscal_top_store_combo.pack(side=tk.LEFT, padx=3)

        ttk.Button(filt, text="Recalcular", command=self._refresh_fiscal_summary).pack(side=tk.LEFT, padx=5)

        ttk.Button(filt, text="Exportar PDF", command=self.export_fiscal_pdf).pack(side=tk.RIGHT, padx=3)
        ttk.Button(filt, text="Exportar Excel (.xlsx)", command=self.export_fiscal_excel).pack(side=tk.RIGHT, padx=3)
        self.gest_q_var = tk.StringVar(value="Todos")
        ttk.Combobox(filt, textvariable=self.gest_q_var, values=["Todos", "1T", "2T", "3T", "4T"], state="readonly", width=6).pack(side=tk.RIGHT, padx=3)
        ttk.Label(filt, text="Trim.:").pack(side=tk.RIGHT, padx=2)

        cols = ("periodo", "v21", "v_rebu", "v_iva", "c_base", "c_iva", "iva_303", "beneficio")
        self.tree_fiscal = ttk.Treeview(self.tab_fiscal, columns=cols, show="headings", height=6)
        for c, txt, w in [
            ("periodo", "Periodo", 90), ("v21", "Ventas 21%/SinIVA", 125), ("v_rebu", "Ventas REBU", 115),
            ("v_iva", "IVA Reperc.", 110), ("c_base", "Gastos Base", 115), ("c_iva", "IVA Soport.", 110),
            ("iva_303", "IVA Mod. 303", 115), ("beneficio", "Beneficio Neto", 120)
        ]:
            self.tree_fiscal.heading(c, text=txt)
            self.tree_fiscal.column(c, width=w, anchor=tk.E if c != "periodo" else tk.CENTER)
        self.tree_fiscal.pack(fill=tk.X, pady=(0, 10))

        lower = ttk.Frame(self.tab_fiscal)
        lower.pack(fill=tk.BOTH, expand=True)

        caja = ttk.LabelFrame(lower, text="Cierre de Caja Diario", padding=8)
        caja.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))
        cf = ttk.Frame(caja)
        cf.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(cf, text="Fecha:").pack(side=tk.LEFT, padx=3)
        self.caja_date_var = tk.StringVar(value=datetime.date.today().strftime("%d/%m/%Y"))
        ttk.Entry(cf, textvariable=self.caja_date_var, width=11).pack(side=tk.LEFT, padx=3)
        self.fiscal_store_var = tk.StringVar(value="Todas")
        self.fiscal_store_combo = ttk.Combobox(cf, textvariable=self.fiscal_store_var, values=["Todas"] + self.config["stores"], state="readonly", width=26)
        self.fiscal_store_combo.pack(side=tk.LEFT, padx=4)
        ttk.Button(cf, text="Ver", command=self._refresh_caja).pack(side=tk.LEFT, padx=4)
        self.txt_caja = tk.Text(caja, height=11, font=("Consolas", 9), state="disabled")
        self.txt_caja.pack(fill=tk.BOTH, expand=True)

        tpv = ttk.LabelFrame(lower, text="Cuadre TPV (Tarjeta) por Tienda", padding=8)
        tpv.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(6, 0))
        tf = ttk.Frame(tpv)
        tf.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(tf, text="Mes (MM/AAAA o vacio=Ano):").pack(side=tk.LEFT, padx=3)
        self.tpv_month_var = tk.StringVar(value=datetime.date.today().strftime("%m/%Y"))
        ttk.Entry(tf, textvariable=self.tpv_month_var, width=10).pack(side=tk.LEFT, padx=4)
        ttk.Button(tf, text="Cuadrar TPV", command=self._refresh_tpv).pack(side=tk.LEFT, padx=4)
        self.txt_tpv = tk.Text(tpv, height=11, font=("Consolas", 9), state="disabled")
        self.txt_tpv.pack(fill=tk.BOTH, expand=True)

    def _refresh_fiscal_summary(self):
        for r in self.tree_fiscal.get_children():
            self.tree_fiscal.delete(r)
        try:
            yr = int(self.fiscal_year_var.get())
        except ValueError:
            yr = datetime.date.today().year
        st_flt = self.fiscal_top_store_var.get()

        cur = self.db.conn.cursor()
        totals = [0.0] * 7
        self._fiscal_rows_cache = []
        for q in ("1T", "2T", "3T", "4T"):
            sql_d = "SELECT base_amount, tax_amount, total_amount, items_json FROM documents WHERE year=? AND quarter=? AND doc_type NOT IN ('Presupuesto', 'Albaran Entrega')"
            p_d = [yr, q]
            if st_flt != "Todas":
                sql_d += " AND store=?"
                p_d.append(st_flt)
            docs = cur.execute(sql_d, p_d).fetchall()

            v21, v_rebu, v_base_tot, v_iva_tot = 0.0, 0.0, 0.0, 0.0
            for d in docs:
                v_base_tot += d["base_amount"]
                v_iva_tot += d["tax_amount"]
                try:
                    items = json.loads(d["items_json"] or "[]")
                    if any("REBU" in it.get("tax_label", "") for it in items):
                        v_rebu += d["total_amount"]
                    else:
                        v21 += d["total_amount"]
                except Exception:
                    v21 += d["total_amount"]

            c = cur.execute("SELECT COALESCE(SUM(base_amount),0), COALESCE(SUM(tax_amount),0) FROM purchases WHERE year=? AND quarter=?", (yr, q)).fetchone()
            iva_303 = v_iva_tot - c[1]
            benef = v_base_tot - c[0]
            vals = [v21, v_rebu, v_iva_tot, c[0], c[1], iva_303, benef]
            for i in range(7):
                totals[i] += vals[i]
            row_t = (f"{yr} - {q}", *[f"{x:.2f} EUR" for x in vals])
            self._fiscal_rows_cache.append((f"{yr} - {q}", *vals))
            self.tree_fiscal.insert("", tk.END, values=row_t)

        self._fiscal_rows_cache.append((f"TOTAL {yr}", *totals))
        self.tree_fiscal.insert("", tk.END, values=(f"TOTAL {yr}", *[f"{x:.2f} EUR" for x in totals]))
        self._refresh_caja()
        self._refresh_tpv()

    def _refresh_caja(self):
        d_str = self.caja_date_var.get().strip()
        st = self.fiscal_store_var.get()
        query = """
            SELECT payment_method, COUNT(*),
                   COALESCE(SUM(CASE WHEN payment_status='Pagado' THEN total_amount WHEN payment_status='Pago Parcial' THEN paid_amount ELSE 0 END), 0),
                   COALESCE(SUM(CASE WHEN payment_status='Pendiente' THEN total_amount WHEN payment_status='Pago Parcial' THEN pending_amount ELSE 0 END), 0)
            FROM documents
            WHERE doc_date=? AND doc_type NOT IN ('Presupuesto', 'Albaran Entrega')
        """
        params = [d_str]
        if st != "Todas":
            query += " AND store=?"
            params.append(st)
        query += " GROUP BY payment_method"

        rows = self.db.conn.execute(query, params).fetchall()
        lines = [f"CAJA {d_str} | {st[:30]}", "-" * 50]
        paid_tot, pend_tot = 0.0, 0.0
        for pm, cnt, p_amt, pend_amt in rows:
            lines.append(f"{pm:<14} {cnt:>2}x | Cob: {p_amt:>8.2f} | Pend: {pend_amt:>7.2f}")
            paid_tot += p_amt
            pend_tot += pend_amt
        lines.append("-" * 50)
        lines.append(f"TOTAL COBRADO:   {paid_tot:>10.2f} EUR")
        lines.append(f"TOTAL PENDIENTE: {pend_tot:>10.2f} EUR")
        self.txt_caja.config(state="normal")
        self.txt_caja.delete("1.0", tk.END)
        self.txt_caja.insert("1.0", "\n".join(lines))
        self.txt_caja.config(state="disabled")

    def _refresh_tpv(self):
        m_str = self.tpv_month_var.get().strip()
        yr = self.fiscal_year_var.get().strip()
        if m_str:
            rows = self.db.conn.execute("""
                SELECT store, COUNT(*), COALESCE(SUM(total_amount), 0)
                FROM documents
                WHERE payment_method='Tarjeta (TPV)' AND doc_type NOT IN ('Presupuesto', 'Albaran Entrega') AND doc_date LIKE ?
                GROUP BY store
            """, (f"%/{m_str}",)).fetchall()
            header = f"TPV TARJETA - Mes: {m_str}"
        else:
            rows = self.db.conn.execute("""
                SELECT store, COUNT(*), COALESCE(SUM(total_amount), 0)
                FROM documents
                WHERE payment_method='Tarjeta (TPV)' AND doc_type NOT IN ('Presupuesto', 'Albaran Entrega') AND year=?
                GROUP BY store
            """, (yr,)).fetchall()
            header = f"TPV TARJETA - Ano: {yr}"

        lines = [header, "-" * 48]
        tot = 0.0
        for st, cnt, amt in rows:
            lines.append(f"{st[:26]:<26} {cnt:>3}x {amt:>10.2f} EUR")
            tot += amt
        lines.append("-" * 48)
        lines.append(f"TOTAL TPV: {tot:>10.2f} EUR")
        self.txt_tpv.config(state="normal")
        self.txt_tpv.delete("1.0", tk.END)
        self.txt_tpv.insert("1.0", "\n".join(lines))
        self.txt_tpv.config(state="disabled")

    def export_fiscal_excel(self):
        yr = self.fiscal_year_var.get().strip()
        q = self.gest_q_var.get()
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"Fiscal_{yr}_{q}.xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        wb = Workbook()
        # Hoja 1: Resumen
        ws = wb.active
        ws.title = "Resumen Fiscal"
        ws.append(["Periodo", "Ventas 21%/SinIVA", "Ventas REBU", "IVA Repercutido", "Gastos Base", "IVA Soportado", "IVA Mod. 303", "Beneficio Neto"])
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")
        for r in self._fiscal_rows_cache:
            ws.append(list(r))

        q_flt = "" if q == "Todos" else " AND quarter=?"
        params = [yr] if q == "Todos" else [yr, q]

        # Hoja 2: Emitidas
        ws2 = wb.create_sheet("Emitidas")
        ws2.append(["Tipo", "Numero", "Fecha", "Trimestre", "Tienda", "Cliente", "NIF", "Pago", "Estado", "Pagado", "Pendiente", "Base", "IVA", "Total"])
        for r in self.db.conn.execute(f"SELECT * FROM documents WHERE year=? AND doc_type NOT IN ('Presupuesto', 'Albaran Entrega'){q_flt} ORDER BY id", params).fetchall():
            ws2.append([r["doc_type"], r["doc_num"], r["doc_date"], r["quarter"], r["store"], r["client_name"], r["client_nif"],
                        r["payment_method"], r["payment_status"], r["paid_amount"], r["pending_amount"], r["base_amount"], r["tax_amount"], r["total_amount"]])

        # Hoja 3: Recibidas
        ws3 = wb.create_sheet("Compras y Gastos")
        ws3.append(["Fecha", "Trimestre", "Tipo_Gasto", "Proveedor", "NIF", "N_Factura", "Pago", "Base", "IVA", "Total", "Notas"])
        for r in self.db.conn.execute(f"SELECT * FROM purchases WHERE year=?{q_flt} ORDER BY id", params).fetchall():
            ws3.append([r["inv_date"], r["quarter"], r["category"], r["supplier"], r["supplier_nif"], r["inv_num"],
                        r["payment_method"], r["base_amount"], r["tax_amount"], r["total_amount"], r["notes"]])

        wb.save(path)
        messagebox.showinfo("Exportado", f"Excel fiscal guardado en:\n{path}")
        open_path(path)

    def export_fiscal_pdf(self):
        yr = self.fiscal_year_var.get().strip()
        path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=f"Resumen_Fiscal_{yr}.pdf", filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        pdf = FPDF(orientation="L")
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 14)
        pdf.cell(0, 8, clean_pdf_text(f"RESUMEN FISCAL {yr} - {self.config['company_name']}"), new_x="LMARGIN", new_y="NEXT")
        pdf.ln(4)
        headers = ["Periodo", "Ventas 21%", "Ventas REBU", "IVA Reperc.", "Gastos Base", "IVA Soport.", "IVA 303", "Beneficio"]
        widths = [32, 34, 34, 34, 34, 34, 34, 36]
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_fill_color(235, 235, 235)
        for h, w in zip(headers, widths):
            pdf.cell(w, 8, h, border=1, align="C", fill=True)
        pdf.ln(8)
        pdf.set_font("Helvetica", "", 9)
        for row in self._fiscal_rows_cache:
            pdf.cell(widths[0], 7, str(row[0]), border=1, align="C")
            for val, w in zip(row[1:], widths[1:]):
                pdf.cell(w, 7, f"{val:.2f} EUR", border=1, align="R")
            pdf.ln(7)
        pdf.output(path)
        open_path(path)

    # ==========================================
    # PESTANA 5: BUSCADOR AVANZADO (Ventas / Compras / SAT + Fecha + Precio)
    # ==========================================
    def _build_history_tab(self):
        top = ttk.LabelFrame(self.tab_history, text="Filtros de Busqueda", padding=8)
        top.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(top, text="Seccion:").grid(row=0, column=0, padx=3)
        self.hist_sec_var = tk.StringVar(value="Ventas / Facturas")
        cb_sec = ttk.Combobox(top, textvariable=self.hist_sec_var, values=["Ventas / Facturas", "Compras / Gastos", "Reparaciones SAT"], state="readonly", width=18)
        cb_sec.grid(row=0, column=1, padx=4)
        cb_sec.bind("<<ComboboxSelected>>", lambda e: self._refresh_history_table())

        ttk.Label(top, text="Texto:").grid(row=0, column=2, padx=4)
        self.search_var = tk.StringVar()
        ent = ttk.Entry(top, textvariable=self.search_var, width=22)
        ent.grid(row=0, column=3, padx=4)
        ent.bind("<KeyRelease>", lambda e: self._refresh_history_table())

        ttk.Label(top, text="Desde (DD/MM/AAAA):").grid(row=0, column=4, padx=4)
        self.hist_from_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.hist_from_var, width=11).grid(row=0, column=5, padx=3)

        ttk.Label(top, text="Hasta:").grid(row=0, column=6, padx=4)
        self.hist_to_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.hist_to_var, width=11).grid(row=0, column=7, padx=3)

        ttk.Label(top, text="Min EUR:").grid(row=0, column=8, padx=4)
        self.hist_min_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.hist_min_var, width=7).grid(row=0, column=9, padx=3)

        ttk.Label(top, text="Max EUR:").grid(row=0, column=10, padx=4)
        self.hist_max_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.hist_max_var, width=7).grid(row=0, column=11, padx=3)

        ttk.Button(top, text="Filtrar", command=self._refresh_history_table).grid(row=0, column=12, padx=6)

        cols = ("id", "tipo", "num", "fecha", "cliente", "tel", "pago", "estado", "garantia", "total")
        self.tree_hist = ttk.Treeview(self.tab_history, columns=cols, show="headings", height=19)
        for c, txt, w in [
            ("id", "ID", 35), ("tipo", "Tipo / Cat.", 120), ("num", "Numero", 115),
            ("fecha", "Fecha", 80), ("cliente", "Cliente / Prov.", 155), ("tel", "Tel / NIF", 90),
            ("pago", "Pago / Tec.", 95), ("estado", "Estado", 85), ("garantia", "Garantia / Detalle", 160), ("total", "Total", 90)
        ]:
            self.tree_hist.heading(c, text=txt)
            self.tree_hist.column(c, width=w)
        self.tree_hist.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        bot = ttk.Frame(self.tab_history)
        bot.pack(fill=tk.X)
        ttk.Button(bot, text="Abrir PDF", command=self.open_history_pdf).pack(side=tk.LEFT, padx=4)
        ttk.Button(bot, text="WhatsApp", command=self.whatsapp_history_client).pack(side=tk.LEFT, padx=4)
        ttk.Button(bot, text="Marcar PAGADO", command=self.mark_doc_paid).pack(side=tk.LEFT, padx=4)
        ttk.Button(bot, text="Anadir Pago Parcial", command=self.add_partial_payment).pack(side=tk.LEFT, padx=4)
        ttk.Button(bot, text="Abrir SAT (Garantia)", command=self.invoice_to_sat).pack(side=tk.LEFT, padx=4)
        ttk.Button(bot, text="Emitir Abono (Rectificativa)", command=self.issue_credit_note).pack(side=tk.RIGHT, padx=4)

    def _refresh_history_table(self):
        for r in self.tree_hist.get_children():
            self.tree_hist.delete(r)

        sec = self.hist_sec_var.get()
        q_txt = self.search_var.get().strip().lower()
        d_from = parse_date_opt(self.hist_from_var.get())
        d_to = parse_date_opt(self.hist_to_var.get())
        try:
            p_min = float(self.hist_min_var.get().replace(",", ".")) if self.hist_min_var.get().strip() else None
        except ValueError:
            p_min = None
        try:
            p_max = float(self.hist_max_var.get().replace(",", ".")) if self.hist_max_var.get().strip() else None
        except ValueError:
            p_max = None

        def match_filters(fecha_str, price, *texts):
            if q_txt and not any(q_txt in (str(t or "").lower()) for t in texts):
                return False
            dt = parse_date_opt(fecha_str)
            if d_from and dt and dt < d_from:
                return False
            if d_to and dt and dt > d_to:
                return False
            if p_min is not None and price < p_min:
                return False
            if p_max is not None and price > p_max:
                return False
            return True

        if sec == "Ventas / Facturas":
            for r in self.db.conn.execute("SELECT * FROM documents ORDER BY id DESC LIMIT 500").fetchall():
                if not match_filters(r["doc_date"], r["total_amount"], r["doc_num"], r["client_name"], r["client_phone"], r["doc_type"]):
                    continue
                g_status = calc_warranty_status(r["doc_date"], r["warranty"])
                st_txt = f"Parcial ({r['pending_amount']:.0f}p)" if r["payment_status"] == "Pago Parcial" else r["payment_status"]
                self.tree_hist.insert("", tk.END, values=(
                    r["id"], r["doc_type"], r["doc_num"], r["doc_date"], r["client_name"],
                    r["client_phone"], r["payment_method"], st_txt, g_status, f"{r['total_amount']:.2f} EUR"
                ))
        elif sec == "Compras / Gastos":
            for r in self.db.conn.execute("SELECT * FROM purchases ORDER BY id DESC LIMIT 500").fetchall():
                if not match_filters(r["inv_date"], r["total_amount"], r["supplier"], r["inv_num"], r["category"], r["supplier_nif"]):
                    continue
                self.tree_hist.insert("", tk.END, values=(
                    r["id"], r["category"][:18], r["inv_num"], r["inv_date"], r["supplier"],
                    r["supplier_nif"] or "-", r["payment_method"], "Registrada", r["notes"][:24], f"{r['total_amount']:.2f} EUR"
                ))
        else:
            for r in self.db.conn.execute("SELECT * FROM repairs ORDER BY id DESC LIMIT 500").fetchall():
                if not match_filters(r["entry_date"], r["estimated_price"], r["ticket_num"], r["client_name"], r["client_phone"], r["appliance"], r["brand_model"]):
                    continue
                self.tree_hist.insert("", tk.END, values=(
                    r["id"], "SAT", r["ticket_num"], r["entry_date"], r["client_name"],
                    r["client_phone"], r["technician"] or "-", r["status"], f"{r['appliance']} {r['brand_model']}"[:24], f"{r['estimated_price']:.2f} EUR"
                ))

    def open_history_pdf(self):
        sel = self.tree_hist.selection()
        if not sel:
            return
        rid = self.tree_hist.item(sel[0])["values"][0]
        sec = self.hist_sec_var.get()
        tbl = "documents" if sec == "Ventas / Facturas" else ("purchases" if sec == "Compras / Gastos" else "repairs")
        row = self.db.conn.execute(f"SELECT pdf_path FROM {tbl} WHERE id=?", (rid,)).fetchone()
        if row and row["pdf_path"] and os.path.exists(row["pdf_path"]):
            open_path(row["pdf_path"])
        else:
            messagebox.showwarning("Sin PDF", "Este registro no tiene un PDF asociado.")

    def whatsapp_history_client(self):
        sel = self.tree_hist.selection()
        if not sel or self.hist_sec_var.get() == "Compras / Gastos":
            return
        rid = self.tree_hist.item(sel[0])["values"][0]
        if self.hist_sec_var.get() == "Ventas / Facturas":
            r = self.db.conn.execute("SELECT * FROM documents WHERE id=?", (rid,)).fetchone()
            msg = f"Hola {r['client_name']}, le escribimos de {self.config['company_name']} sobre su {r['doc_type']} {r['doc_num']}."
            open_whatsapp(r["client_phone"], msg)
        else:
            r = self.db.conn.execute("SELECT * FROM repairs WHERE id=?", (rid,)).fetchone()
            msg = f"Hola {r['client_name']}, sobre su {r['appliance']} (Parte {r['ticket_num']}), estado: {r['status']}."
            open_whatsapp(r["client_phone"], msg)

    def mark_doc_paid(self):
        sel = self.tree_hist.selection()
        if not sel or self.hist_sec_var.get() != "Ventas / Facturas":
            return
        doc_id = self.tree_hist.item(sel[0])["values"][0]
        self.db.conn.execute("UPDATE documents SET payment_status='Pagado', paid_amount=total_amount, pending_amount=0 WHERE id=?", (doc_id,))
        self.db.conn.commit()
        self._refresh_all()

    def issue_credit_note(self):
        sel = self.tree_hist.selection()
        if not sel or self.hist_sec_var.get() != "Ventas / Facturas":
            return
        doc_id = self.tree_hist.item(sel[0])["values"][0]
        orig = self.db.conn.execute("SELECT * FROM documents WHERE id=?", (doc_id,)).fetchone()
        if orig["doc_type"] in ("Presupuesto", "Albaran Entrega", "Factura Rectificativa"):
            messagebox.showwarning("No aplicable", "Solo se pueden emitir abonos sobre facturas o recibos.")
            return
        if not messagebox.askyesno("Confirmar Abono", f"Emitir Factura Rectificativa (en negativo) para {orig['doc_num']}?"):
            return

        dt = datetime.date.today()
        quarter = get_quarter(dt)
        seq = self.db.get_next_number(dt.year, "RECT")
        rect_num = f"RECT-{dt.year}-{seq:04d}"
        date_str = dt.strftime("%d/%m/%Y")

        orig_items = json.loads(orig["items_json"] or "[]")
        neg_items = []
        for it in orig_items:
            ni = dict(it)
            ni["net"] = -abs(ni["net"])
            ni["tax_amount"] = -abs(ni["tax_amount"])
            ni["total"] = -abs(ni["total"])
            neg_items.append(ni)

        base_amt = -abs(orig["base_amount"])
        tax_amt = -abs(orig["tax_amount"])
        disc_amt = -abs(orig["discount_amount"])
        grand_total = -abs(orig["total_amount"])
        notes = f"Abono / Rectificacion de {orig['doc_num']} ({orig['doc_date']})."

        target_dir = os.path.join(APP_DIR, "Facturas_Rectificativas", str(dt.year), MONTH_NAMES[dt.month])
        os.makedirs(target_dir, exist_ok=True)
        filepath = os.path.join(target_dir, f"{rect_num}.pdf")

        self._render_document_pdf(
            filepath, "Factura Rectificativa", rect_num, date_str, quarter, orig["store"],
            orig["client_name"], orig["client_nif"], "", "", orig["client_phone"],
            orig["payment_method"], "Pagado", grand_total, 0.0, "Sin garantia / Alquiler", notes,
            neg_items, base_amt, tax_amt, disc_amt, grand_total
        )

        self.db.conn.execute("""
            INSERT INTO documents (
                doc_type, doc_num, doc_date, year, quarter, store,
                client_name, client_nif, client_phone, payment_method, payment_status,
                paid_amount, pending_amount, warranty, notes,
                base_amount, tax_amount, discount_amount, total_amount, items_json, pdf_path
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            "Factura Rectificativa", rect_num, date_str, dt.year, quarter, orig["store"],
            orig["client_name"], orig["client_nif"], orig["client_phone"],
            orig["payment_method"], "Pagado", grand_total, 0.0, "Sin garantia / Alquiler", notes,
            base_amt, tax_amt, disc_amt, grand_total, json.dumps(neg_items), filepath
        ))
        self.db.conn.commit()
        self.db.commit_number(dt.year, "RECT", seq)
        self._refresh_all()
        messagebox.showinfo("Abono Emitido", f"Factura Rectificativa guardada en:\n{filepath}")
        open_path(filepath)

    def stock_to_invoice(self, s):
        self.clear_billing_form()
        self.doc_type_var.set("Factura")
        self._update_doc_number()
        if s.get("store") in self.config["stores"]:
            self.store_var.set(s["store"])
        if s.get("category") in self.config["categories"]:
            self.cat_var.set(s["category"])
        self.desc_var.set(f"{s.get('brand', '')} {s.get('model', '')} (S/N: {s.get('serial_num') or '-'})".strip())
        self.price_var.set(str(s.get("sale_price", 0)))
        self.add_or_update_item()
        self.notebook.select(self.tab_billing)

    def add_partial_payment(self):
        sel = self.tree_hist.selection()
        if not sel or self.hist_sec_var.get() != "Ventas / Facturas":
            return
        doc_id = self.tree_hist.item(sel[0])["values"][0]
        r = self.db.conn.execute("SELECT * FROM documents WHERE id=?", (doc_id,)).fetchone()
        if r["payment_status"] == "Pagado":
            messagebox.showinfo("Pagado", "Este documento ya esta pagado por completo.")
            return
        pend_actual = r["pending_amount"] if r["pending_amount"] > 0 else r["total_amount"]
        win = tk.Toplevel(self.root)
        win.title("Anadir Pago Parcial")
        frm = ttk.Frame(win, padding=14)
        frm.pack()
        ttk.Label(frm, text=f"{r['doc_num']} - Pendiente actual: {pend_actual:.2f} EUR").pack(pady=(0, 8))
        v_abono = tk.StringVar(value=f"{pend_actual:.2f}")
        ttk.Entry(frm, textvariable=v_abono, width=15).pack(pady=4)

        def guardar():
            try:
                abono = float(v_abono.get().replace(",", "."))
                if abono <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Error", "Importe no valido.", parent=win)
                return
            nuevo_pagado = round((r["paid_amount"] or 0) + abono, 2)
            nuevo_pend = max(0.0, round(r["total_amount"] - nuevo_pagado, 2))
            nuevo_est = "Pagado" if nuevo_pend <= 0.01 else "Pago Parcial"
            self.db.conn.execute(
                "UPDATE documents SET paid_amount=?, pending_amount=?, payment_status=? WHERE id=?",
                (min(nuevo_pagado, r["total_amount"]), nuevo_pend, nuevo_est, doc_id)
            )
            self.db.conn.commit()
            win.destroy()
            self._refresh_all()

        ttk.Button(frm, text="Registrar Pago", command=guardar).pack(pady=(8, 0), fill=tk.X)

    def invoice_to_sat(self):
        sel = self.tree_hist.selection()
        if not sel or self.hist_sec_var.get() != "Ventas / Facturas":
            return
        doc_id = self.tree_hist.item(sel[0])["values"][0]
        r = self.db.conn.execute("SELECT * FROM documents WHERE id=?", (doc_id,)).fetchone()
        self.sat_store_var.set(r["store"] if r["store"] in self.config["stores"] else self.config["stores"][0])
        self.sat_client_var.set(r["client_name"] or "")
        self.sat_nif_var.set(r["client_nif"] or "")
        self.sat_phone_var.set(r["client_phone"] or "")
        try:
            items = json.loads(r["items_json"] or "[]")
            if items:
                self.sat_model_var.set(items[0].get("name", ""))
        except Exception:
            pass
        self.sat_issue_var.set(f"Incidencia Garantia - Ref. {r['doc_num']} ({r['doc_date']})")
        self.notebook.select(self.tab_sat)


if __name__ == "__main__":
    root = tk.Tk()
    app = ERPApp(root)
    root.mainloop()