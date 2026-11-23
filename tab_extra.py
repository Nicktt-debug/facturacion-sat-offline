import os
import sys
import json
import sqlite3
import datetime
import tkinter as tk
from tkinter import ttk, messagebox
from fpdf import FPDF


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


def _load_company(app_dir):
    cfg_path = os.path.join(app_dir, "config.json")
    default = {"company_name": "MI EMPRESA", "stores": ["Tienda 1"], "categories": ["Lavadora", "Nevera", "Otros"]}
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                default.update(json.load(f))
        except Exception:
            pass
    return default


def build_extra_tab(parent, db, app_dir, on_sell_stock=None):
    cfg = _load_company(app_dir)
    stores = cfg.get("stores", ["Tienda 1"])
    categories = cfg.get("categories", ["Lavadora", "Nevera", "Otros"])

    def ensure_tables():
        c = db.conn
        c.execute("""
            CREATE TABLE IF NOT EXISTS stock (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                store TEXT, category TEXT, brand TEXT, model TEXT,
                serial_num TEXT, cost_price REAL, sale_price REAL,
                warranty TEXT, status TEXT DEFAULT 'Disponible'
            )
        """)
        c.execute("""
            CREATE TABLE IF NOT EXISTS petty_cash (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT, store TEXT, tipo TEXT, concepto TEXT, importe REAL
            )
        """)
        c.execute("""
            CREATE TABLE IF NOT EXISTS rma (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT, supplier TEXT, rma_num TEXT, item TEXT,
                amount REAL, status TEXT DEFAULT 'Pendiente', notes TEXT
            )
        """)
        c.commit()

    ensure_tables()

    sub_nb = ttk.Notebook(parent)
    sub_nb.pack(fill=tk.BOTH, expand=True)

    tab_stock = ttk.Frame(sub_nb, padding=8)
    tab_cash = ttk.Frame(sub_nb, padding=8)
    tab_rma = ttk.Frame(sub_nb, padding=8)

    sub_nb.add(tab_stock, text="Stock y Etiquetas")
    sub_nb.add(tab_cash, text="Caja Chica (Efectivo)")
    sub_nb.add(tab_rma, text="RMA / Devoluciones Proveedor")

    # ==========================================
    # SUB-PESTANA 1: STOCK Y ETIQUETAS DE PRECIO
    # ==========================================
    sf = ttk.LabelFrame(tab_stock, text="Agregar Electrodomestico al Stock", padding=8)
    sf.pack(fill=tk.X, pady=(0, 8))

    s_store = tk.StringVar(value=stores[0])
    s_cat = tk.StringVar(value=categories[0])
    s_brand = tk.StringVar()
    s_model = tk.StringVar()
    s_serial = tk.StringVar()
    s_cost = tk.StringVar(value="0")
    s_sale = tk.StringVar(value="0")
    s_warr = tk.StringVar(value="6 meses")

    ttk.Label(sf, text="Tienda:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=3)
    ttk.Combobox(sf, textvariable=s_store, values=stores, state="readonly", width=34).grid(row=0, column=1, columnspan=2, sticky=tk.W, padx=4)
    ttk.Label(sf, text="Categoria:").grid(row=0, column=3, sticky=tk.W, padx=6)
    ttk.Combobox(sf, textvariable=s_cat, values=categories, width=18).grid(row=0, column=4, padx=4)
    ttk.Label(sf, text="Marca:").grid(row=0, column=5, sticky=tk.W, padx=6)
    ttk.Entry(sf, textvariable=s_brand, width=16).grid(row=0, column=6, padx=4)

    ttk.Label(sf, text="Modelo:").grid(row=1, column=0, sticky=tk.W, padx=4, pady=3)
    ttk.Entry(sf, textvariable=s_model, width=22).grid(row=1, column=1, padx=4)
    ttk.Label(sf, text="N. Serie:").grid(row=1, column=2, sticky=tk.E, padx=4)
    ttk.Entry(sf, textvariable=s_serial, width=18).grid(row=1, column=3, padx=4)
    ttk.Label(sf, text="Coste EUR:").grid(row=1, column=4, sticky=tk.E, padx=4)
    ttk.Entry(sf, textvariable=s_cost, width=10).grid(row=1, column=5, sticky=tk.W, padx=4)
    ttk.Label(sf, text="Venta EUR:").grid(row=1, column=5, sticky=tk.E, padx=4)
    ttk.Entry(sf, textvariable=s_sale, width=10).grid(row=1, column=6, sticky=tk.W, padx=4)

    ttk.Label(sf, text="Garantia:").grid(row=2, column=0, sticky=tk.W, padx=4, pady=3)
    ttk.Combobox(sf, textvariable=s_warr, values=["6 meses", "1 ano (REBU)", "3 anos (Nuevo)", "3 meses"], width=18).grid(row=2, column=1, sticky=tk.W, padx=4)

    cols_s = ("id", "tienda", "cat", "marca", "modelo", "serie", "coste", "venta", "garantia", "estado")
    tree_s = ttk.Treeview(tab_stock, columns=cols_s, show="headings", height=13)
    for c, txt, w in [
        ("id", "ID", 35), ("tienda", "Tienda", 160), ("cat", "Categoria", 100),
        ("marca", "Marca", 85), ("modelo", "Modelo", 130), ("serie", "N. Serie", 110),
        ("coste", "Coste", 70), ("venta", "Precio Venta", 85), ("garantia", "Garantia", 90), ("estado", "Estado", 80)
    ]:
        tree_s.heading(c, text=txt)
        tree_s.column(c, width=w)
    tree_s.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

    def refresh_stock():
        ensure_tables()
        for r in tree_s.get_children():
            tree_s.delete(r)
        for r in db.conn.execute("SELECT * FROM stock ORDER BY id DESC").fetchall():
            tree_s.insert("", tk.END, values=(
                r["id"], r["store"][:24], r["category"], r["brand"], r["model"],
                r["serial_num"], f"{r['cost_price']:.2f}", f"{r['sale_price']:.2f} EUR",
                r["warranty"], r["status"]
            ))

    def add_stock():
        try:
            cost = float(s_cost.get().replace(",", "."))
            sale = float(s_sale.get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Error", "Revisa el precio de coste y venta.")
            return
        db.conn.execute("""
            INSERT INTO stock (store, category, brand, model, serial_num, cost_price, sale_price, warranty, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'Disponible')
        """, (s_store.get(), s_cat.get(), s_brand.get().strip(), s_model.get().strip(),
              s_serial.get().strip(), cost, sale, s_warr.get().strip()))
        db.conn.commit()
        for v in (s_brand, s_model, s_serial):
            v.set("")
        s_cost.set("0")
        s_sale.set("0")
        refresh_stock()

    def print_label(fmt="A4"):
        sel = tree_s.selection()
        if not sel:
            messagebox.showinfo("Selecciona", "Selecciona un aparato de la tabla.")
            return
        sid = tree_s.item(sel[0])["values"][0]
        r = db.conn.execute("SELECT * FROM stock WHERE id=?", (sid,)).fetchone()
        etq_dir = os.path.join(app_dir, "Etiquetas")
        os.makedirs(etq_dir, exist_ok=True)
        path = os.path.join(etq_dir, f"Etiqueta_ID{sid}_{fmt}.pdf")

        if fmt == "80mm":
            pdf = FPDF(unit="mm", format=(80, 90))
            pdf.set_margins(4, 4, 4)
            pdf.add_page()
            pdf.set_font("Helvetica", "B", 10)
            pdf.cell(0, 5, _clean_pdf(cfg["company_name"]), align="C", new_x="LMARGIN", new_y="NEXT")
            pdf.cell(0, 4, "-" * 38, align="C", new_x="LMARGIN", new_y="NEXT")
            pdf.set_font("Helvetica", "B", 11)
            pdf.multi_cell(0, 6, _clean_pdf(f"{r['category']} {r['brand']} {r['model']}"), align="C", new_x="LMARGIN", new_y="NEXT")
            pdf.set_font("Helvetica", "", 8)
            pdf.cell(0, 5, _clean_pdf(f"Ref: #{r['id']}  |  S/N: {r['serial_num'] or '-'}"), align="C", new_x="LMARGIN", new_y="NEXT")
            pdf.ln(3)
            pdf.set_font("Helvetica", "B", 22)
            pdf.cell(0, 12, f"{r['sale_price']:.2f} EUR", border=1, align="C", new_x="LMARGIN", new_y="NEXT")
            pdf.ln(3)
            pdf.set_font("Helvetica", "B", 9)
            pdf.cell(0, 5, _clean_pdf(f"Garantia: {r['warranty']}"), align="C", new_x="LMARGIN", new_y="NEXT")
        else:
            pdf = FPDF(orientation="L", format="A5")
            pdf.add_page()
            pdf.set_font("Helvetica", "B", 16)
            pdf.cell(0, 10, _clean_pdf(cfg["company_name"]), align="C", border="B", new_x="LMARGIN", new_y="NEXT")
            pdf.ln(6)
            pdf.set_font("Helvetica", "B", 20)
            pdf.multi_cell(0, 10, _clean_pdf(f"{r['category'].upper()} - {r['brand']} {r['model']}"), align="C", new_x="LMARGIN", new_y="NEXT")
            pdf.ln(4)
            pdf.set_font("Helvetica", "", 12)
            pdf.cell(0, 8, _clean_pdf(f"Codigo Interno: #{r['id']}    |    N. Serie: {r['serial_num'] or '-'}"), align="C", new_x="LMARGIN", new_y="NEXT")
            pdf.ln(8)
            pdf.set_font("Helvetica", "B", 42)
            pdf.cell(0, 24, f"{r['sale_price']:.2f} EUR", border=1, align="C", new_x="LMARGIN", new_y="NEXT")
            pdf.ln(8)
            pdf.set_font("Helvetica", "B", 14)
            pdf.cell(0, 8, _clean_pdf(f"GARANTIA INCLUIDA: {r['warranty']}"), align="C", new_x="LMARGIN", new_y="NEXT")

        pdf.output(path)
        _open_file(path)

    def sell_selected():
        sel = tree_s.selection()
        if not sel:
            return
        sid = tree_s.item(sel[0])["values"][0]
        r = db.conn.execute("SELECT * FROM stock WHERE id=?", (sid,)).fetchone()
        db.conn.execute("UPDATE stock SET status='Vendido' WHERE id=?", (sid,))
        db.conn.commit()
        refresh_stock()
        if on_sell_stock:
            on_sell_stock(dict(r))

    def del_stock():
        sel = tree_s.selection()
        if not sel:
            return
        sid = tree_s.item(sel[0])["values"][0]
        db.conn.execute("DELETE FROM stock WHERE id=?", (sid,))
        db.conn.commit()
        refresh_stock()

    ttk.Button(sf, text="+ Guardar en Stock", command=add_stock).grid(row=2, column=5, columnspan=2, sticky=tk.EW, padx=4)

    sb_s = ttk.Frame(tab_stock)
    sb_s.pack(fill=tk.X)
    ttk.Button(sb_s, text="Etiqueta A5 (Escaparate)", command=lambda: print_label("A4")).pack(side=tk.LEFT, padx=4)
    ttk.Button(sb_s, text="Etiqueta 80mm (Ticket)", command=lambda: print_label("80mm")).pack(side=tk.LEFT, padx=4)
    ttk.Button(sb_s, text="Eliminar", command=del_stock).pack(side=tk.LEFT, padx=4)
    ttk.Button(sb_s, text="Vender y Pasar a Factura ->", command=sell_selected).pack(side=tk.RIGHT, padx=4)

    # ==========================================
    # SUB-PESTANA 2: CAJA CHICA (EFECTIVO)
    # ==========================================
    cf = ttk.LabelFrame(tab_cash, text="Movimientos de Efectivo (Gasolina, Recambios, Retirada Banco...)", padding=8)
    cf.pack(fill=tk.X, pady=(0, 8))

    c_fecha = tk.StringVar(value=datetime.date.today().strftime("%d/%m/%Y"))
    c_store = tk.StringVar(value=stores[0])
    c_tipo = tk.StringVar(value="Salida")
    c_conc = tk.StringVar()
    c_imp = tk.StringVar()

    ttk.Label(cf, text="Fecha:").grid(row=0, column=0, padx=4, pady=3)
    ttk.Entry(cf, textvariable=c_fecha, width=12).grid(row=0, column=1, padx=4)
    ttk.Label(cf, text="Tienda:").grid(row=0, column=2, padx=4)
    ttk.Combobox(cf, textvariable=c_store, values=stores, state="readonly", width=32).grid(row=0, column=3, padx=4)
    ttk.Label(cf, text="Tipo:").grid(row=0, column=4, padx=4)
    ttk.Combobox(cf, textvariable=c_tipo, values=["Salida", "Entrada"], state="readonly", width=10).grid(row=0, column=5, padx=4)

    ttk.Label(cf, text="Concepto:").grid(row=1, column=0, padx=4, pady=4)
    ttk.Entry(cf, textvariable=c_conc, width=48).grid(row=1, column=1, columnspan=3, sticky=tk.EW, padx=4)
    ttk.Label(cf, text="Importe EUR:").grid(row=1, column=4, padx=4)
    ttk.Entry(cf, textvariable=c_imp, width=10).grid(row=1, column=5, padx=4)

    cols_c = ("id", "fecha", "tienda", "tipo", "concepto", "importe")
    tree_c = ttk.Treeview(tab_cash, columns=cols_c, show="headings", height=13)
    for c, txt, w in [
        ("id", "ID", 35), ("fecha", "Fecha", 85), ("tienda", "Tienda", 190),
        ("tipo", "Tipo", 80), ("concepto", "Concepto", 360), ("importe", "Importe", 100)
    ]:
        tree_c.heading(c, text=txt)
        tree_c.column(c, width=w)
    tree_c.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

    lbl_cash_tot = ttk.Label(tab_cash, text="", font=("Segoe UI", 10, "bold"))
    lbl_cash_tot.pack(side=tk.RIGHT, padx=8)

    def refresh_cash():
        ensure_tables()
        for r in tree_c.get_children():
            tree_c.delete(r)
        ent_tot, sal_tot = 0.0, 0.0
        for r in db.conn.execute("SELECT * FROM petty_cash ORDER BY id DESC LIMIT 200").fetchall():
            if r["tipo"] == "Entrada":
                ent_tot += r["importe"]
            else:
                sal_tot += r["importe"]
            tree_c.insert("", tk.END, values=(
                r["id"], r["fecha"], r["store"][:28], r["tipo"], r["concepto"], f"{r['importe']:.2f} EUR"
            ))
        lbl_cash_tot.config(text=f"Entradas: +{ent_tot:.2f} EUR  |  Salidas: -{sal_tot:.2f} EUR  |  Neto: {ent_tot - sal_tot:.2f} EUR")

    def add_cash():
        try:
            imp = float(c_imp.get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Error", "Importe invalido.")
            return
        if not c_conc.get().strip():
            return
        db.conn.execute("INSERT INTO petty_cash (fecha, store, tipo, concepto, importe) VALUES (?, ?, ?, ?, ?)",
                        (c_fecha.get().strip(), c_store.get(), c_tipo.get(), c_conc.get().strip(), round(imp, 2)))
        db.conn.commit()
        c_conc.set("")
        c_imp.set("")
        refresh_cash()

    def del_cash():
        sel = tree_c.selection()
        if not sel:
            return
        cid = tree_c.item(sel[0])["values"][0]
        db.conn.execute("DELETE FROM petty_cash WHERE id=?", (cid,))
        db.conn.commit()
        refresh_cash()

    ttk.Button(cf, text="+ Registrar Movimiento", command=add_cash).grid(row=1, column=6, padx=6)
    ttk.Button(tab_cash, text="Eliminar Seleccionado", command=del_cash).pack(side=tk.LEFT, padx=4)

    # ==========================================
    # SUB-PESTANA 3: RMA / DEVOLUCIONES PROVEEDOR
    # ==========================================
    rf = ttk.LabelFrame(tab_rma, text="Seguimiento de Devoluciones / RMA a Proveedores", padding=8)
    rf.pack(fill=tk.X, pady=(0, 8))

    r_fecha = tk.StringVar(value=datetime.date.today().strftime("%d/%m/%Y"))
    r_sup = tk.StringVar()
    r_num = tk.StringVar()
    r_item = tk.StringVar()
    r_amt = tk.StringVar(value="0")
    r_notes = tk.StringVar()

    ttk.Label(rf, text="Fecha:").grid(row=0, column=0, padx=4, pady=3)
    ttk.Entry(rf, textvariable=r_fecha, width=12).grid(row=0, column=1, padx=4)
    ttk.Label(rf, text="Proveedor:").grid(row=0, column=2, padx=4)
    ttk.Entry(rf, textvariable=r_sup, width=22).grid(row=0, column=3, padx=4)
    ttk.Label(rf, text="N. RMA:").grid(row=0, column=4, padx=4)
    ttk.Entry(rf, textvariable=r_num, width=15).grid(row=0, column=5, padx=4)

    ttk.Label(rf, text="Aparato/Pieza:").grid(row=1, column=0, padx=4, pady=4)
    ttk.Entry(rf, textvariable=r_item, width=28).grid(row=1, column=1, columnspan=2, sticky=tk.EW, padx=4)
    ttk.Label(rf, text="Importe EUR:").grid(row=1, column=3, sticky=tk.E, padx=4)
    ttk.Entry(rf, textvariable=r_amt, width=10).grid(row=1, column=4, sticky=tk.W, padx=4)
    ttk.Entry(rf, textvariable=r_notes, width=20).grid(row=1, column=5, padx=4)

    cols_r = ("id", "fecha", "prov", "rma", "item", "importe", "estado", "notas")
    tree_r = ttk.Treeview(tab_rma, columns=cols_r, show="headings", height=13)
    for c, txt, w in [
        ("id", "ID", 35), ("fecha", "Fecha", 80), ("prov", "Proveedor", 150),
        ("rma", "N. RMA", 100), ("item", "Articulo / Pieza", 200),
        ("importe", "Importe", 85), ("estado", "Estado", 95), ("notas", "Notas", 170)
    ]:
        tree_r.heading(c, text=txt)
        tree_r.column(c, width=w)
    tree_r.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

    lbl_rma_pend = ttk.Label(tab_rma, text="", font=("Segoe UI", 10, "bold"), foreground="#b3261e")
    lbl_rma_pend.pack(side=tk.RIGHT, padx=8)

    def refresh_rma():
        ensure_tables()
        for r in tree_r.get_children():
            tree_r.delete(r)
        pend = 0.0
        for r in db.conn.execute("SELECT * FROM rma ORDER BY id DESC").fetchall():
            if r["status"] in ("Pendiente", "Enviado"):
                pend += r["amount"]
            tree_r.insert("", tk.END, values=(
                r["id"], r["fecha"], r["supplier"], r["rma_num"], r["item"],
                f"{r['amount']:.2f} EUR", r["status"], r["notes"]
            ))
        lbl_rma_pend.config(text=f"Total RMA pendiente de abono: {pend:.2f} EUR")

    def add_rma():
        try:
            amt = float(r_amt.get().replace(",", "."))
        except ValueError:
            amt = 0.0
        if not r_sup.get().strip():
            return
        db.conn.execute("""
            INSERT INTO rma (fecha, supplier, rma_num, item, amount, status, notes)
            VALUES (?, ?, ?, ?, ?, 'Pendiente', ?)
        """, (r_fecha.get().strip(), r_sup.get().strip(), r_num.get().strip(), r_item.get().strip(), round(amt, 2), r_notes.get().strip()))
        db.conn.commit()
        for v in (r_sup, r_num, r_item, r_notes):
            v.set("")
        r_amt.set("0")
        refresh_rma()

    r_new_st = tk.StringVar(value="Abonado")

    def upd_rma():
        sel = tree_r.selection()
        if not sel:
            return
        rid = tree_r.item(sel[0])["values"][0]
        db.conn.execute("UPDATE rma SET status=? WHERE id=?", (r_new_st.get(), rid))
        db.conn.commit()
        refresh_rma()

    def del_rma():
        sel = tree_r.selection()
        if not sel:
            return
        rid = tree_r.item(sel[0])["values"][0]
        db.conn.execute("DELETE FROM rma WHERE id=?", (rid,))
        db.conn.commit()
        refresh_rma()

    ttk.Button(rf, text="+ Guardar RMA", command=add_rma).grid(row=1, column=6, padx=6)
    rb = ttk.Frame(tab_rma)
    rb.pack(side=tk.LEFT)
    ttk.Combobox(rb, textvariable=r_new_st, values=["Pendiente", "Enviado", "Abonado", "Rechazado"], state="readonly", width=12).pack(side=tk.LEFT, padx=4)
    ttk.Button(rb, text="Cambiar Estado", command=upd_rma).pack(side=tk.LEFT, padx=4)
    ttk.Button(rb, text="Eliminar", command=del_rma).pack(side=tk.LEFT, padx=4)

    refresh_stock()
    refresh_cash()
    refresh_rma()