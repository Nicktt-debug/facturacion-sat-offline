"""
Pestana 7 - Agenda de servicios (entregas, averias, instalaciones...).

Cambios principales respecto a la version anterior:
  * Fechas guardadas como AAAA-MM-DD (ordenan bien); se muestran DD/MM/AAAA.
    Las filas antiguas se convierten solas la primera vez.
  * Validacion de fecha / hora / telefono y aviso de doble reserva.
  * Duracion por servicio (ocupa varios huecos).
  * Editar, reprogramar, estados con colores, busqueda y filtros.
  * Autocompletar cliente desde la tabla 'clients' de facturacion.
  * Acciones rapidas: copiar telefono, Google Maps, WhatsApp, crear factura.
  * Vista semanal y hoja de ruta en PDF.
  * Listas de vehiculos / tecnicos / tipos configurables en config.json:
        "agenda_vehiculos": [...], "agenda_tecnicos": [...], "agenda_tipos": [...]

Se usa igual que antes:  build_agenda_tab(parent, db)
Opcional:                build_agenda_tab(parent, db, on_invoice=callback)
"""
import os
import re
import sys
import json
import sqlite3
import datetime
import webbrowser
import urllib.parse
import tkinter as tk
from tkinter import ttk, messagebox

# ==========================================
# CONSTANTES
# ==========================================
HORA_INICIO = 11
HORA_FIN = 22  # el ultimo hueco empieza a las 21:00 y termina a las 22:00
HORARIOS = [f"{h:02d}:00" for h in range(HORA_INICIO, HORA_FIN)]

SIN_VEHICULO = "Sin vehiculo"
SIN_TECNICO = "Sin asignar"
VEHICULOS_DEF = ["Furgoneta 1", "Furgoneta 2", "Vehiculo 3 / Particular", SIN_VEHICULO]
TECNICOS_DEF = ["Tecnico 1", "Tecnico 2", "Repartidor 1", "Repartidor 2", SIN_TECNICO]
TIPOS_SERVICIO = ["Entrega Venta", "Averia / SAT", "Reparacion Domicilio",
                  "Instalacion", "Recogida", "Otro (No registrado)"]

ESTADOS = ["Pendiente", "En camino", "Completado", "Cancelado", "Reprogramado"]
ESTADO_TAG = {
    "Pendiente": "pend", "En camino": "camino", "Completado": "hecho",
    "Cancelado": "cancel", "Reprogramado": "repro",
}
DURACIONES = ["1", "2", "3", "4"]
DIAS = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado", "Domingo"]


# ==========================================
# UTILIDADES
# ==========================================
def _s(x):
    return "" if x is None else str(x)


def parse_fecha(texto):
    """'DD/MM/AAAA' -> date, o None si no es valida."""
    try:
        return datetime.datetime.strptime(_s(texto).strip(), "%d/%m/%Y").date()
    except ValueError:
        return None


def to_iso(d):
    return d.strftime("%Y-%m-%d")


def iso_to_display(iso):
    try:
        return datetime.datetime.strptime(_s(iso), "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return _s(iso)


def hora_a_int(hora):
    try:
        return int(_s(hora)[:2])
    except ValueError:
        return HORA_INICIO


def expandir_horas(row):
    """Devuelve [(hora_int, es_inicio)] de los huecos que ocupa un servicio."""
    ini = hora_a_int(row["hora"])
    dur = max(1, int(row["duracion"] or 1))
    return [(h, h == ini) for h in range(ini, min(ini + dur, HORA_FIN))]


def _clean_pdf(text):
    text = _s(text).replace("\u20ac", "EUR").replace("\u2013", "-").replace("\u2014", "-")
    return text.encode("latin-1", "replace").decode("latin-1")


def _abrir_archivo(path):
    if sys.platform == "win32":
        os.startfile(path)
    else:
        messagebox.showinfo("Archivo", path)


def _telefono_wa(telefono):
    digits = re.sub(r"\D", "", _s(telefono))
    if len(digits) == 9:
        digits = "34" + digits
    return digits


def _load_lists(app_dir):
    """Lee las listas de la agenda desde config.json (si existen)."""
    cfg = {}
    path = os.path.join(app_dir, "config.json")
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except (OSError, ValueError):
            cfg = {}

    def lista(key, default, obligatorio):
        val = cfg.get(key)
        if isinstance(val, list):
            val = [str(x).strip() for x in val if str(x).strip()]
        else:
            val = []
        if not val:
            val = list(default)
        if obligatorio and obligatorio not in val:
            val.append(obligatorio)
        return val

    return {
        "vehiculos": lista("agenda_vehiculos", VEHICULOS_DEF, SIN_VEHICULO),
        "tecnicos": lista("agenda_tecnicos", TECNICOS_DEF, SIN_TECNICO),
        "tipos": lista("agenda_tipos", TIPOS_SERVICIO, None),
        "empresa": _s(cfg.get("company_name")),
    }


# ==========================================
# ACCESO A DATOS (la interfaz nunca escribe SQL)
# ==========================================
class AgendaRepository:
    FIELDS = ("fecha", "hora", "duracion", "tipo", "vehiculo", "tecnico",
              "cliente", "telefono", "direccion", "detalle")

    def __init__(self, db):
        self.db = db

    @property
    def conn(self):
        # Siempre se pide a db: tras "Restaurar Backup" la conexion cambia.
        return self.db.conn

    def ensure_schema(self):
        c = self.conn
        c.row_factory = sqlite3.Row
        c.execute("""
            CREATE TABLE IF NOT EXISTS agenda (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT, hora TEXT, tipo TEXT, vehiculo TEXT, tecnico TEXT,
                cliente TEXT, telefono TEXT, direccion TEXT, detalle TEXT, estado TEXT
            )
        """)
        cols = {row[1] for row in c.execute("PRAGMA table_info(agenda)").fetchall()}
        if "duracion" not in cols:
            c.execute("ALTER TABLE agenda ADD COLUMN duracion INTEGER DEFAULT 1")

        # Migracion: fechas DD/MM/AAAA -> AAAA-MM-DD (idempotente)
        pendientes = c.execute(
            "SELECT id, fecha FROM agenda WHERE fecha IS NOT NULL AND fecha NOT LIKE '____-__-__'"
        ).fetchall()
        for rid, fecha in pendientes:
            d = parse_fecha(fecha)
            if d:
                c.execute("UPDATE agenda SET fecha=? WHERE id=?", (to_iso(d), rid))

        c.execute("UPDATE agenda SET duracion=1 WHERE duracion IS NULL OR duracion<1")
        c.execute("UPDATE agenda SET estado='Pendiente' WHERE estado IS NULL OR estado=''")
        c.execute("CREATE INDEX IF NOT EXISTS idx_agenda_fecha ON agenda(fecha, hora)")
        c.commit()

    def _write(self, sql, params=()):
        try:
            self.conn.execute(sql, params)
            self.conn.commit()
        except sqlite3.Error:
            self.conn.rollback()
            raise

    # --- lectura ---
    def get(self, rid):
        return self.conn.execute("SELECT * FROM agenda WHERE id=?", (rid,)).fetchone()

    def day(self, fecha_iso):
        return self.conn.execute(
            "SELECT * FROM agenda WHERE fecha=? ORDER BY hora, id", (fecha_iso,)
        ).fetchall()

    def search(self, fecha_iso=None, text="", estado="Todos"):
        sql = "SELECT * FROM agenda WHERE 1=1"
        params = []
        if fecha_iso:
            sql += " AND fecha=?"
            params.append(fecha_iso)
        if estado and estado != "Todos":
            sql += " AND estado=?"
            params.append(estado)
        if text:
            like = f"%{text}%"
            sql += (" AND (cliente LIKE ? OR telefono LIKE ? OR tecnico LIKE ?"
                    " OR vehiculo LIKE ? OR direccion LIKE ? OR detalle LIKE ?)")
            params += [like] * 6
        if fecha_iso:
            sql += " ORDER BY hora, id"
        else:
            sql += " ORDER BY fecha DESC, hora, id LIMIT 500"
        return self.conn.execute(sql, params).fetchall()

    def range(self, desde_iso, hasta_iso):
        return self.conn.execute(
            "SELECT * FROM agenda WHERE fecha BETWEEN ? AND ? "
            "AND COALESCE(estado,'')!='Cancelado' ORDER BY fecha, hora, id",
            (desde_iso, hasta_iso)
        ).fetchall()

    def conflicts(self, fecha_iso, hora, dur, vehiculo, tecnico, exclude_id=None):
        """Servicios que se solapan y comparten vehiculo o tecnico."""
        sql = "SELECT * FROM agenda WHERE fecha=? AND COALESCE(estado,'')!='Cancelado'"
        params = [fecha_iso]
        if exclude_id is not None:
            sql += " AND id!=?"
            params.append(exclude_id)
        ini = hora_a_int(hora)
        fin = ini + dur
        out = []
        for r in self.conn.execute(sql, params).fetchall():
            r_ini = hora_a_int(r["hora"])
            r_fin = r_ini + int(r["duracion"] or 1)
            if r_ini < fin and ini < r_fin:
                mismo_v = vehiculo != SIN_VEHICULO and r["vehiculo"] == vehiculo
                mismo_t = tecnico != SIN_TECNICO and r["tecnico"] == tecnico
                if mismo_v or mismo_t:
                    out.append((r, mismo_v, mismo_t))
        return out

    def clients(self):
        try:
            return self.conn.execute(
                "SELECT name, phone, address FROM clients ORDER BY name"
            ).fetchall()
        except sqlite3.Error:
            return []

    # --- escritura ---
    def insert(self, d):
        self._write(
            "INSERT INTO agenda (fecha, hora, duracion, tipo, vehiculo, tecnico, cliente,"
            " telefono, direccion, detalle, estado) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            tuple(d[k] for k in self.FIELDS) + ("Pendiente",)
        )

    def update(self, rid, d):
        self._write(
            "UPDATE agenda SET fecha=?, hora=?, duracion=?, tipo=?, vehiculo=?, tecnico=?,"
            " cliente=?, telefono=?, direccion=?, detalle=? WHERE id=?",
            tuple(d[k] for k in self.FIELDS) + (rid,)
        )

    def set_estado(self, rid, estado):
        self._write("UPDATE agenda SET estado=? WHERE id=?", (estado, rid))

    def move(self, rid, fecha_iso, hora):
        self._write("UPDATE agenda SET fecha=?, hora=?, estado='Reprogramado' WHERE id=?",
                    (fecha_iso, hora, rid))

    def delete(self, rid):
        self._write("DELETE FROM agenda WHERE id=?", (rid,))


# ==========================================
# INTERFAZ
# ==========================================
class AgendaTab:
    def __init__(self, parent, db, on_invoice=None):
        self.parent = parent
        self.db = db
        self.on_invoice = on_invoice
        self.repo = AgendaRepository(db)
        self.repo.ensure_schema()

        db_path = getattr(db, "db_path", None)
        self.app_dir = os.path.dirname(os.path.abspath(db_path)) if db_path else os.getcwd()
        listas = _load_lists(self.app_dir)
        self.vehiculos = listas["vehiculos"]
        self.tecnicos = listas["tecnicos"]
        self.tipos = listas["tipos"]
        self.empresa = listas["empresa"]

        self.client_map = {}
        self.edit_id = None
        self.edit_fecha_iso = None

        self.v_fecha = tk.StringVar(value=datetime.date.today().strftime("%d/%m/%Y"))
        self.v_hora = tk.StringVar(value=HORARIOS[0])
        self.v_dur = tk.StringVar(value="1")
        self.v_tipo = tk.StringVar(value=self.tipos[0])
        self.v_veh = tk.StringVar(value=self.vehiculos[0])
        self.v_tec = tk.StringVar(value=self.tecnicos[0])
        self.v_cli = tk.StringVar()
        self.v_tel = tk.StringVar()
        self.v_dir = tk.StringVar()
        self.v_det = tk.StringVar()
        self.v_buscar = tk.StringVar()
        self.v_est_filtro = tk.StringVar(value="Todos")
        self.v_todas = tk.BooleanVar(value=False)
        self.v_estado_nuevo = tk.StringVar(value=ESTADOS[1])

        self._build_ui()
        self._load_clients()
        self.v_fecha.trace_add("write", self._on_fecha_change)
        self.v_buscar.trace_add("write", lambda *_: self.refresh_views())
        self._bind_shortcuts()
        self._on_fecha_change()

    # ---------- API publica ----------
    def reload(self):
        """Llamar tras restaurar un backup u otro cambio externo de la BD."""
        try:
            self.repo.ensure_schema()
        except sqlite3.Error as e:
            messagebox.showerror("Agenda", f"No se pudo preparar la agenda:\n{e}")
            return
        self._load_clients()
        self.refresh_views()

    # ---------- construccion ----------
    def _build_ui(self):
        p = self.parent
        ttk.Style(p).configure("Bad.TEntry", foreground="#b3261e")

        # --- Barra superior (fecha) ---
        top = ttk.Frame(p)
        top.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(top, text="Fecha Agenda (DD/MM/AAAA):", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=4)
        self.ent_fecha = ttk.Entry(top, textvariable=self.v_fecha, width=13)
        self.ent_fecha.pack(side=tk.LEFT, padx=4)
        self.lbl_dia = ttk.Label(top, width=18, font=("Segoe UI", 9, "bold"))
        self.lbl_dia.pack(side=tk.LEFT, padx=2)

        ttk.Button(top, text="< Dia Anterior", command=lambda: self.cambiar_dia(-1)).pack(side=tk.LEFT, padx=3)
        ttk.Button(top, text="Hoy", command=self.ir_a_hoy).pack(side=tk.LEFT, padx=3)
        ttk.Button(top, text="Dia Siguiente >", command=lambda: self.cambiar_dia(1)).pack(side=tk.LEFT, padx=3)
        ttk.Button(top, text="Actualizar", command=self.refresh_views).pack(side=tk.LEFT, padx=8)
        ttk.Button(top, text="Hoja de Ruta (PDF)", command=self.export_route_sheet).pack(side=tk.RIGHT, padx=3)
        ttk.Button(top, text="Vista Semanal", command=self.open_week_view).pack(side=tk.RIGHT, padx=3)

        # --- Cuerpo ---
        main_pane = ttk.Frame(p)
        main_pane.pack(fill=tk.BOTH, expand=True)

        left_frm = ttk.LabelFrame(main_pane, text="Huecos 11:00 - 21:00 (Clic para elegir hora)", padding=6)
        left_frm.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 8))

        self.tree_slots = ttk.Treeview(left_frm, columns=("hora", "estado", "resumen"), show="headings", height=14)
        self.tree_slots.heading("hora", text="Hora")
        self.tree_slots.heading("estado", text="Estado")
        self.tree_slots.heading("resumen", text="Servicios / Vehiculos")
        self.tree_slots.column("hora", width=55, anchor=tk.CENTER)
        self.tree_slots.column("estado", width=85, anchor=tk.CENTER)
        self.tree_slots.column("resumen", width=190)
        self.tree_slots.tag_configure("libre", background="#e8f5e9")
        self.tree_slots.tag_configure("ocupado", background="#ffebee")
        self.tree_slots.pack(fill=tk.BOTH, expand=True)
        self.tree_slots.bind("<<TreeviewSelect>>", self._on_slot_click)
        ttk.Label(left_frm, text="* = continua del hueco anterior (servicio de varias horas)",
                  font=("Segoe UI", 8, "italic")).pack(anchor=tk.W, pady=(4, 0))

        right_frm = ttk.Frame(main_pane)
        right_frm.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # --- Formulario ---
        self.form = ttk.LabelFrame(right_frm, text="Asignar Entrega / Averia / Servicio (Registrado o Manual)", padding=8)
        self.form.pack(fill=tk.X, pady=(0, 8))
        f = self.form

        ttk.Label(f, text="Hora:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=3)
        ttk.Combobox(f, textvariable=self.v_hora, values=HORARIOS, state="readonly", width=8).grid(row=0, column=1, sticky=tk.W, padx=4)
        ttk.Label(f, text="Duracion (h):").grid(row=0, column=2, sticky=tk.W, padx=6)
        ttk.Combobox(f, textvariable=self.v_dur, values=DURACIONES, state="readonly", width=6).grid(row=0, column=3, sticky=tk.W, padx=4)
        ttk.Label(f, text="Tipo:").grid(row=0, column=4, sticky=tk.W, padx=6)
        ttk.Combobox(f, textvariable=self.v_tipo, values=self.tipos, state="readonly", width=20).grid(row=0, column=5, sticky=tk.W, padx=4)

        ttk.Label(f, text="Vehiculo:").grid(row=1, column=0, sticky=tk.W, padx=4, pady=3)
        ttk.Combobox(f, textvariable=self.v_veh, values=self.vehiculos, state="readonly", width=20).grid(row=1, column=1, columnspan=2, sticky=tk.W, padx=4)
        ttk.Label(f, text="Tecnico/Repart.:").grid(row=1, column=3, sticky=tk.W, padx=6)
        ttk.Combobox(f, textvariable=self.v_tec, values=self.tecnicos, state="readonly", width=16).grid(row=1, column=4, sticky=tk.W, padx=4)
        ttk.Label(f, text="Cliente:").grid(row=1, column=5, sticky=tk.W, padx=6)
        self.cb_cli = ttk.Combobox(f, textvariable=self.v_cli, width=24)
        self.cb_cli.grid(row=1, column=6, sticky=tk.W, padx=4)
        self.cb_cli.bind("<<ComboboxSelected>>", self._on_client_selected)
        self.cb_cli.bind("<KeyRelease>", self._filter_clients)
        self.cb_cli.bind("<FocusIn>", lambda e: self._load_clients())
        self.cb_cli.bind("<FocusOut>", self._on_client_focus_out)

        ttk.Label(f, text="Telefono:").grid(row=2, column=0, sticky=tk.W, padx=4, pady=3)
        ttk.Entry(f, textvariable=self.v_tel, width=14).grid(row=2, column=1, columnspan=2, sticky=tk.W, padx=4)
        ttk.Label(f, text="Direccion:").grid(row=2, column=3, sticky=tk.W, padx=6)
        ttk.Entry(f, textvariable=self.v_dir, width=38).grid(row=2, column=4, columnspan=2, sticky=tk.EW, padx=4)
        ttk.Label(f, text="Aparato/Nota:").grid(row=2, column=6, sticky=tk.W, padx=6)
        ttk.Entry(f, textvariable=self.v_det, width=24).grid(row=2, column=7, sticky=tk.W, padx=4)

        btns = ttk.Frame(f)
        btns.grid(row=3, column=0, columnspan=8, sticky=tk.EW, pady=(8, 0))
        btns.columnconfigure(0, weight=1)
        self.btn_save = ttk.Button(btns, text="+ Guardar en Agenda", command=self.save)
        self.btn_save.grid(row=0, column=0, sticky=tk.EW)
        self.btn_cancel_edit = ttk.Button(btns, text="Cancelar edicion", command=self._reset_form)
        self.btn_cancel_edit.grid(row=0, column=1, padx=(6, 0))
        self.btn_cancel_edit.state(["disabled"])

        for w in f.winfo_children():
            if isinstance(w, (ttk.Entry, ttk.Combobox)):
                w.bind("<Return>", lambda e: self.save())

        # --- Lista del dia ---
        list_frm = ttk.LabelFrame(right_frm, text="Servicios Programados para la Fecha", padding=6)
        list_frm.pack(fill=tk.BOTH, expand=True)

        flt = ttk.Frame(list_frm)
        flt.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(flt, text="Buscar:").pack(side=tk.LEFT, padx=(0, 4))
        ttk.Entry(flt, textvariable=self.v_buscar, width=24).pack(side=tk.LEFT)
        ttk.Label(flt, text="Estado:").pack(side=tk.LEFT, padx=(10, 4))
        cb_est = ttk.Combobox(flt, textvariable=self.v_est_filtro, values=["Todos"] + ESTADOS, state="readonly", width=13)
        cb_est.pack(side=tk.LEFT)
        cb_est.bind("<<ComboboxSelected>>", lambda e: self.refresh_views())
        ttk.Checkbutton(flt, text="Buscar en todas las fechas", variable=self.v_todas,
                        command=self.refresh_views).pack(side=tk.LEFT, padx=12)

        self.cols = ("id", "fecha", "hora", "dur", "tipo", "vehiculo", "tecnico", "cliente", "dir", "detalle", "estado")
        tf = ttk.Frame(list_frm)
        tf.pack(fill=tk.BOTH, expand=True, pady=(0, 6))
        tf.rowconfigure(0, weight=1)
        tf.columnconfigure(0, weight=1)
        self.tree_day = ttk.Treeview(tf, columns=self.cols, show="headings", height=11)
        for c, txt, w in [
            ("id", "ID", 32), ("fecha", "Fecha", 78), ("hora", "Hora", 48), ("dur", "Dur.", 34),
            ("tipo", "Tipo", 100), ("vehiculo", "Vehiculo", 90), ("tecnico", "Tecnico", 85),
            ("cliente", "Cliente", 140), ("dir", "Direccion", 125), ("detalle", "Detalle", 110),
            ("estado", "Estado", 85),
        ]:
            self.tree_day.heading(c, text=txt)
            self.tree_day.column(c, width=w)
        vs = ttk.Scrollbar(tf, orient=tk.VERTICAL, command=self.tree_day.yview)
        hs = ttk.Scrollbar(tf, orient=tk.HORIZONTAL, command=self.tree_day.xview)
        self.tree_day.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree_day.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")

        self.tree_day.tag_configure("pend", background="#ffffff")
        self.tree_day.tag_configure("camino", background="#fff8e1")
        self.tree_day.tag_configure("hecho", background="#e8f5e9", foreground="#2e7d32")
        self.tree_day.tag_configure("cancel", background="#f2f2f2", foreground="#888888")
        self.tree_day.tag_configure("repro", background="#e3f2fd")
        self.tree_day.tag_configure("atrasado", background="#ffebee", foreground="#b3261e")
        self.tree_day.bind("<Double-1>", self.start_edit)
        self.tree_day.bind("<Return>", self.start_edit)
        self.tree_day.bind("<Delete>", self.delete_selected)

        bar1 = ttk.Frame(list_frm)
        bar1.pack(fill=tk.X, pady=(0, 4))
        ttk.Button(bar1, text="Editar", command=self.start_edit).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(bar1, text="Marcar Completado", command=lambda: self.set_estado("Completado")).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(bar1, text="Estado:").pack(side=tk.LEFT, padx=(6, 3))
        ttk.Combobox(bar1, textvariable=self.v_estado_nuevo, values=ESTADOS, state="readonly", width=13).pack(side=tk.LEFT)
        ttk.Button(bar1, text="Aplicar", command=lambda: self.set_estado(self.v_estado_nuevo.get())).pack(side=tk.LEFT, padx=(4, 10))
        ttk.Button(bar1, text="Reprogramar...", command=self.reschedule).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(bar1, text="Eliminar", command=self.delete_selected).pack(side=tk.LEFT)

        bar2 = ttk.Frame(list_frm)
        bar2.pack(fill=tk.X)
        ttk.Button(bar2, text="Copiar Telefono", command=self.copy_phone).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(bar2, text="Abrir en Maps", command=self.open_maps).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(bar2, text="WhatsApp Recordatorio", command=self.whatsapp_reminder).pack(side=tk.LEFT, padx=(0, 6))
        self.btn_invoice = ttk.Button(bar2, text="Crear Factura ->", command=self.create_invoice)
        self.btn_invoice.pack(side=tk.RIGHT)
        if self.on_invoice is None:
            self.btn_invoice.state(["disabled"])

    def _bind_shortcuts(self):
        top = self.parent.winfo_toplevel()
        top.bind("<Alt-Left>", lambda e: self._shortcut(lambda: self.cambiar_dia(-1)), add="+")
        top.bind("<Alt-Right>", lambda e: self._shortcut(lambda: self.cambiar_dia(1)), add="+")
        top.bind("<F5>", lambda e: self._shortcut(self.refresh_views), add="+")

    def _shortcut(self, fn):
        # Solo actua si esta pestana es la visible
        if self.parent.winfo_ismapped():
            fn()

    # ---------- fecha ----------
    def _on_fecha_change(self, *_):
        d = parse_fecha(self.v_fecha.get())
        if d is None:
            self.ent_fecha.configure(style="Bad.TEntry")
            self.lbl_dia.config(text="Fecha no valida", foreground="#b3261e")
            return
        self.ent_fecha.configure(style="TEntry")
        txt = DIAS[d.weekday()] + ("  (HOY)" if d == datetime.date.today() else "")
        self.lbl_dia.config(text=txt, foreground="#1b5e20" if d == datetime.date.today() else "")
        self.refresh_views()

    def cambiar_dia(self, delta):
        d = parse_fecha(self.v_fecha.get()) or datetime.date.today()
        self.v_fecha.set((d + datetime.timedelta(days=delta)).strftime("%d/%m/%Y"))

    def ir_a_hoy(self):
        self.v_fecha.set(datetime.date.today().strftime("%d/%m/%Y"))

    # ---------- refresco ----------
    def refresh_views(self):
        d = parse_fecha(self.v_fecha.get())
        todas = self.v_todas.get()
        if d is None and not todas:
            return
        sel = self.tree_day.selection()
        try:
            if d is not None:
                self._refresh_slots(d)
            self._refresh_list(d, todas)
        except sqlite3.Error as e:
            messagebox.showerror("Error de base de datos", str(e))
            return
        keep = [i for i in sel if self.tree_day.exists(i)]
        if keep:
            self.tree_day.selection_set(keep)

    def _refresh_slots(self, d):
        rows = [r for r in self.repo.day(to_iso(d)) if _s(r["estado"]) != "Cancelado"]
        ocup = {h: [] for h in range(HORA_INICIO, HORA_FIN)}
        for r in rows:
            for h, inicio in expandir_horas(r):
                ocup[h].append((r, inicio))

        for item in self.tree_slots.get_children():
            self.tree_slots.delete(item)
        for h in range(HORA_INICIO, HORA_FIN):
            hora = f"{h:02d}:00"
            items = ocup[h]
            if items:
                resumen = ", ".join(
                    f"{_s(r['vehiculo'])[:10]}/{_s(r['tecnico'])[:8]}" + ("" if ini else "*")
                    for r, ini in items
                )
                self.tree_slots.insert("", tk.END, values=(hora, f"Ocupado ({len(items)})", resumen), tags=("ocupado",))
            else:
                self.tree_slots.insert("", tk.END, values=(hora, "Libre", "-"), tags=("libre",))

    def _refresh_list(self, d, todas):
        rows = self.repo.search(None if todas else to_iso(d),
                                self.v_buscar.get().strip(), self.v_est_filtro.get())
        hoy = to_iso(datetime.date.today())
        self.tree_day["displaycolumns"] = self.cols if todas else [c for c in self.cols if c != "fecha"]

        for item in self.tree_day.get_children():
            self.tree_day.delete(item)
        for r in rows:
            estado = _s(r["estado"]) or "Pendiente"
            if _s(r["fecha"]) < hoy and estado in ("Pendiente", "En camino"):
                tag = "atrasado"
            else:
                tag = ESTADO_TAG.get(estado, "pend")
            cli = _s(r["cliente"])
            if r["telefono"]:
                cli = f"{cli} ({r['telefono']})"
            self.tree_day.insert("", tk.END, iid=str(r["id"]), tags=(tag,), values=(
                r["id"], iso_to_display(r["fecha"]), r["hora"], f"{int(r['duracion'] or 1)}h",
                r["tipo"], r["vehiculo"], r["tecnico"], cli, r["direccion"], r["detalle"], estado
            ))

    def _on_slot_click(self, event=None):
        sel = self.tree_slots.selection()
        if sel:
            self.v_hora.set(self.tree_slots.item(sel[0])["values"][0])

    # ---------- clientes (autocompletar) ----------
    def _load_clients(self):
        try:
            rows = self.repo.clients()
        except sqlite3.Error:
            rows = []
        self.client_map = {_s(r["name"]): r for r in rows if _s(r["name"]).strip()}
        self.cb_cli["values"] = list(self.client_map.keys())[:200]

    def _filter_clients(self, event):
        if event.keysym in ("Up", "Down", "Left", "Right", "Return", "Tab", "Escape",
                            "Shift_L", "Shift_R", "Control_L", "Control_R"):
            return
        txt = self.v_cli.get().strip().lower()
        names = [n for n in self.client_map if txt in n.lower()] if txt else list(self.client_map)
        self.cb_cli["values"] = names[:50]

    def _on_client_selected(self, event=None):
        c = self.client_map.get(self.v_cli.get())
        if c:
            if c["phone"]:
                self.v_tel.set(c["phone"])
            if c["address"]:
                self.v_dir.set(c["address"])

    def _on_client_focus_out(self, event=None):
        c = self.client_map.get(self.v_cli.get().strip())
        if c and not self.v_tel.get().strip() and not self.v_dir.get().strip():
            self._on_client_selected()

    # ---------- formulario ----------
    def _read_form(self):
        """Valida y devuelve un dict listo para guardar. Lanza ValueError con el motivo."""
        if self.edit_id is not None:
            fecha_iso = self.edit_fecha_iso
        else:
            d = parse_fecha(self.v_fecha.get())
            if d is None:
                raise ValueError("La fecha no es valida. Usa el formato DD/MM/AAAA.")
            fecha_iso = to_iso(d)

        cli = self.v_cli.get().strip()
        det = self.v_det.get().strip()
        if not cli and not det:
            raise ValueError("Escribe al menos el cliente o el detalle del servicio.")

        hora = self.v_hora.get().strip()
        if hora not in HORARIOS:
            raise ValueError("Elige una hora de la lista (11:00 a 21:00).")
        try:
            dur = max(1, int(self.v_dur.get()))
        except ValueError:
            dur = 1
        if hora_a_int(hora) + dur > HORA_FIN:
            raise ValueError(f"Un servicio de {dur}h a las {hora} acabaria despues de las "
                             f"{HORA_FIN}:00. Reduce la duracion o elige una hora anterior.")

        tel = self.v_tel.get().strip()
        if tel and len(re.sub(r"\D", "", tel)) < 9:
            raise ValueError("El telefono parece incompleto (minimo 9 digitos).")

        return {
            "fecha": fecha_iso, "hora": hora, "duracion": dur,
            "tipo": self.v_tipo.get().strip(), "vehiculo": self.v_veh.get().strip(),
            "tecnico": self.v_tec.get().strip(), "cliente": cli, "telefono": tel,
            "direccion": self.v_dir.get().strip(), "detalle": det,
        }

    def _confirm_conflicts(self, fecha_iso, hora, dur, vehiculo, tecnico, exclude_id=None, parent=None):
        try:
            conf = self.repo.conflicts(fecha_iso, hora, dur, vehiculo, tecnico, exclude_id)
        except sqlite3.Error as e:
            messagebox.showerror("Error de base de datos", str(e), parent=parent or self.parent)
            return False
        if not conf:
            return True
        lines = []
        for r, mismo_v, mismo_t in conf[:5]:
            quien = " y ".join(x for x, ok in ((r["vehiculo"], mismo_v), (r["tecnico"], mismo_t)) if ok)
            lines.append(f"- {r['hora']} ({int(r['duracion'] or 1)}h): {quien} ya esta con "
                         f"{_s(r['cliente']) or _s(r['detalle']) or 'otro servicio'}")
        return messagebox.askyesno(
            "Posible doble reserva",
            "Hay solapamientos ese dia:\n\n" + "\n".join(lines) + "\n\nGuardar de todas formas?",
            icon="warning", parent=parent or self.parent
        )

    def save(self):
        try:
            data = self._read_form()
        except ValueError as e:
            messagebox.showwarning("Revisa los datos", str(e))
            return
        if not self._confirm_conflicts(data["fecha"], data["hora"], data["duracion"],
                                       data["vehiculo"], data["tecnico"], self.edit_id):
            return
        try:
            if self.edit_id is not None:
                self.repo.update(self.edit_id, data)
            else:
                self.repo.insert(data)
        except sqlite3.Error as e:
            messagebox.showerror("Error de base de datos", f"No se pudo guardar:\n{e}")
            return
        self._reset_form()
        self.refresh_views()

    def _reset_form(self):
        self.edit_id = None
        self.edit_fecha_iso = None
        for v in (self.v_cli, self.v_tel, self.v_dir, self.v_det):
            v.set("")
        self.v_dur.set("1")
        self.form.config(text="Asignar Entrega / Averia / Servicio (Registrado o Manual)")
        self.btn_save.config(text="+ Guardar en Agenda")
        self.btn_cancel_edit.state(["disabled"])

    # ---------- seleccion ----------
    def _selected_row(self, aviso=True):
        sel = self.tree_day.selection()
        row = None
        if sel:
            try:
                row = self.repo.get(int(sel[0]))
            except sqlite3.Error as e:
                messagebox.showerror("Error de base de datos", str(e))
                return None
        if row is None and aviso:
            messagebox.showinfo("Selecciona un servicio", "Selecciona primero un servicio de la lista.")
        return row

    # ---------- editar / estado / borrar ----------
    def start_edit(self, event=None):
        r = self._selected_row(aviso=event is None)
        if not r:
            return
        self.edit_id = r["id"]
        self.edit_fecha_iso = r["fecha"]
        self.v_hora.set(r["hora"] if r["hora"] in HORARIOS else HORARIOS[0])
        self.v_dur.set(str(int(r["duracion"] or 1)))
        self.v_tipo.set(_s(r["tipo"]))
        self.v_veh.set(_s(r["vehiculo"]))
        self.v_tec.set(_s(r["tecnico"]))
        self.v_cli.set(_s(r["cliente"]))
        self.v_tel.set(_s(r["telefono"]))
        self.v_dir.set(_s(r["direccion"]))
        self.v_det.set(_s(r["detalle"]))
        self.form.config(text=f"Editando servicio #{r['id']} del {iso_to_display(r['fecha'])}")
        self.btn_save.config(text="Guardar cambios")
        self.btn_cancel_edit.state(["!disabled"])

    def set_estado(self, estado):
        r = self._selected_row()
        if not r:
            return
        try:
            self.repo.set_estado(r["id"], estado)
        except sqlite3.Error as e:
            messagebox.showerror("Error de base de datos", str(e))
            return
        self.refresh_views()

    def delete_selected(self, event=None):
        r = self._selected_row(aviso=event is None)
        if not r:
            return
        if not messagebox.askyesno(
            "Eliminar servicio",
            f"Eliminar definitivamente este servicio?\n\n{r['hora']} - {r['tipo']}\n"
            f"{_s(r['cliente']) or _s(r['detalle'])}\n\n"
            "Si solo quieres anularlo, ponle el estado 'Cancelado'.",
            icon="warning"
        ):
            return
        try:
            self.repo.delete(r["id"])
        except sqlite3.Error as e:
            messagebox.showerror("Error de base de datos", str(e))
            return
        if self.edit_id == r["id"]:
            self._reset_form()
        self.refresh_views()

    def reschedule(self):
        r = self._selected_row()
        if not r:
            return
        win = tk.Toplevel(self.parent)
        win.title("Reprogramar servicio")
        win.transient(self.parent.winfo_toplevel())
        win.resizable(False, False)
        frm = ttk.Frame(win, padding=14)
        frm.pack()

        v_f = tk.StringVar(value=iso_to_display(r["fecha"]))
        v_h = tk.StringVar(value=r["hora"] if r["hora"] in HORARIOS else HORARIOS[0])
        ttk.Label(frm, text=_s(r["cliente"]) or _s(r["detalle"]), font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, columnspan=2, pady=(0, 8))
        ttk.Label(frm, text="Nueva fecha:").grid(row=1, column=0, sticky=tk.W, pady=3)
        ttk.Entry(frm, textvariable=v_f, width=13).grid(row=1, column=1, padx=6)
        ttk.Label(frm, text="Nueva hora:").grid(row=2, column=0, sticky=tk.W, pady=3)
        ttk.Combobox(frm, textvariable=v_h, values=HORARIOS, state="readonly", width=10).grid(row=2, column=1, padx=6)

        def ok():
            nd = parse_fecha(v_f.get())
            if nd is None:
                messagebox.showwarning("Fecha no valida", "Usa el formato DD/MM/AAAA.", parent=win)
                return
            dur = int(r["duracion"] or 1)
            if hora_a_int(v_h.get()) + dur > HORA_FIN:
                messagebox.showwarning("Hora no valida",
                                       f"Con {dur}h de duracion acabaria despues de las {HORA_FIN}:00.", parent=win)
                return
            if not self._confirm_conflicts(to_iso(nd), v_h.get(), dur, r["vehiculo"], r["tecnico"], r["id"], parent=win):
                return
            try:
                self.repo.move(r["id"], to_iso(nd), v_h.get())
            except sqlite3.Error as e:
                messagebox.showerror("Error de base de datos", str(e), parent=win)
                return
            win.destroy()
            self.refresh_views()

        ttk.Button(frm, text="Mover", command=ok).grid(row=3, column=0, columnspan=2, sticky=tk.EW, pady=(10, 0))
        win.wait_visibility()
        win.grab_set()

    # ---------- acciones rapidas ----------
    def copy_phone(self):
        r = self._selected_row()
        if not r:
            return
        if not _s(r["telefono"]).strip():
            messagebox.showinfo("Sin telefono", "Este servicio no tiene telefono.")
            return
        self.parent.clipboard_clear()
        self.parent.clipboard_append(_s(r["telefono"]).strip())

    def open_maps(self):
        r = self._selected_row()
        if not r:
            return
        if not _s(r["direccion"]).strip():
            messagebox.showinfo("Sin direccion", "Este servicio no tiene direccion.")
            return
        q = urllib.parse.quote(_s(r["direccion"]).strip())
        webbrowser.open(f"https://www.google.com/maps/search/?api=1&query={q}")

    def whatsapp_reminder(self):
        r = self._selected_row()
        if not r:
            return
        digits = _telefono_wa(r["telefono"])
        if not digits:
            messagebox.showwarning("Sin telefono", "Este servicio no tiene un telefono valido.")
            return
        firma = f" Un saludo, {self.empresa}." if self.empresa else ""
        msg = (f"Hola {_s(r['cliente'])}, le recordamos que el {iso_to_display(r['fecha'])} a las "
               f"{r['hora']} tenemos programado: {r['tipo']}"
               f"{' (' + _s(r['detalle']) + ')' if r['detalle'] else ''}."
               f"{' Direccion: ' + _s(r['direccion']) + '.' if r['direccion'] else ''}{firma}")
        webbrowser.open(f"https://wa.me/{digits}?text={urllib.parse.quote(msg)}")

    def create_invoice(self):
        if self.on_invoice is None:
            return
        r = self._selected_row()
        if not r:
            return
        if _s(r["estado"]) == "Cancelado":
            messagebox.showinfo("Servicio cancelado", "Este servicio esta cancelado.")
            return
        self.on_invoice({
            "id": r["id"], "fecha": iso_to_display(r["fecha"]), "tipo": _s(r["tipo"]),
            "cliente": _s(r["cliente"]), "telefono": _s(r["telefono"]),
            "direccion": _s(r["direccion"]), "detalle": _s(r["detalle"]),
        })

    # ---------- vista semanal ----------
    def open_week_view(self):
        base = parse_fecha(self.v_fecha.get()) or datetime.date.today()
        state = {"lunes": base - datetime.timedelta(days=base.weekday())}

        win = tk.Toplevel(self.parent)
        win.title("Vista semanal")
        win.geometry("1120x470")
        win.transient(self.parent.winfo_toplevel())

        top = ttk.Frame(win, padding=8)
        top.pack(fill=tk.X)
        lbl = ttk.Label(top, font=("Segoe UI", 10, "bold"))
        v_filtro = tk.StringVar(value="Todos")
        opciones = ["Todos"] + [v for v in self.vehiculos if v != SIN_VEHICULO] + \
                   [t for t in self.tecnicos if t != SIN_TECNICO]

        cols = ("hora",) + tuple(f"d{i}" for i in range(7))
        tree = ttk.Treeview(win, columns=cols, show="headings", height=11)
        tree.heading("hora", text="Hora")
        tree.column("hora", width=55, anchor=tk.CENTER)
        for i in range(7):
            tree.column(f"d{i}", width=145)
        tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))
        ttk.Label(win, text="Doble clic en un dia para abrirlo en la agenda.   * = continua del hueco anterior",
                  font=("Segoe UI", 8, "italic")).pack(anchor=tk.W, padx=10, pady=(0, 8))

        def dias():
            return [state["lunes"] + datetime.timedelta(days=i) for i in range(7)]

        def texto_celda(items):
            partes = []
            for r, ini in items[:2]:
                nombre = (_s(r["cliente"]) or _s(r["detalle"]) or "?")[:16]
                partes.append(nombre + ("" if ini else "*"))
            if len(items) > 2:
                partes.append(f"+{len(items) - 2}")
            return " | ".join(partes)

        def dibujar(*_):
            ds = dias()
            lbl.config(text=f"Semana del {ds[0]:%d/%m/%Y} al {ds[6]:%d/%m/%Y}")
            try:
                rows = self.repo.range(to_iso(ds[0]), to_iso(ds[6]))
            except sqlite3.Error as e:
                messagebox.showerror("Error de base de datos", str(e), parent=win)
                return
            flt = v_filtro.get()
            grid = {}
            for r in rows:
                if flt != "Todos" and flt not in (r["vehiculo"], r["tecnico"]):
                    continue
                for h, ini in expandir_horas(r):
                    grid.setdefault((r["fecha"], h), []).append((r, ini))
            for i, d in enumerate(ds):
                marca = " *HOY*" if d == datetime.date.today() else ""
                tree.heading(f"d{i}", text=f"{DIAS[i][:3]} {d:%d/%m}{marca}")
            for item in tree.get_children():
                tree.delete(item)
            for h in range(HORA_INICIO, HORA_FIN):
                celdas = [texto_celda(grid.get((to_iso(d), h), [])) for d in ds]
                tree.insert("", tk.END, values=(f"{h:02d}:00", *celdas))

        def mover(delta):
            state["lunes"] += datetime.timedelta(days=7 * delta)
            dibujar()

        ttk.Button(top, text="< Semana anterior", command=lambda: mover(-1)).pack(side=tk.LEFT, padx=3)
        lbl.pack(side=tk.LEFT, padx=10)
        ttk.Button(top, text="Semana siguiente >", command=lambda: mover(1)).pack(side=tk.LEFT, padx=3)
        ttk.Label(top, text="Ver:").pack(side=tk.RIGHT, padx=(10, 3))
        cb = ttk.Combobox(top, textvariable=v_filtro, values=opciones, state="readonly", width=24)
        cb.pack(side=tk.RIGHT)
        cb.bind("<<ComboboxSelected>>", dibujar)
        ttk.Label(top, text="Filtrar vehiculo / tecnico").pack(side=tk.RIGHT, padx=4)

        def on_dbl(event):
            col = tree.identify_column(event.x)
            try:
                idx = int(col.lstrip("#")) - 2
            except ValueError:
                return
            if 0 <= idx < 7:
                self.v_fecha.set(dias()[idx].strftime("%d/%m/%Y"))
                win.destroy()

        tree.bind("<Double-1>", on_dbl)
        dibujar()

    # ---------- hoja de ruta ----------
    def export_route_sheet(self):
        d = parse_fecha(self.v_fecha.get())
        if d is None:
            messagebox.showwarning("Fecha no valida", "Corrige la fecha (DD/MM/AAAA) antes de exportar.")
            return
        try:
            rows = [r for r in self.repo.day(to_iso(d)) if _s(r["estado"]) != "Cancelado"]
        except sqlite3.Error as e:
            messagebox.showerror("Error de base de datos", str(e))
            return
        if not rows:
            messagebox.showinfo("Sin servicios", "No hay servicios para esta fecha.")
            return
        try:
            from fpdf import FPDF
        except ImportError:
            messagebox.showerror("Falta fpdf", "Instala la libreria con:  pip install fpdf2")
            return

        grupos = {}
        for r in rows:
            grupos.setdefault(_s(r["vehiculo"]) or SIN_VEHICULO, []).append(r)

        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font("Helvetica", "B", 14)
        pdf.cell(0, 8, _clean_pdf(f"HOJA DE RUTA - {DIAS[d.weekday()]} {d:%d/%m/%Y}"), new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 9)
        pdf.cell(0, 5, _clean_pdf(f"{self.empresa}   |   Servicios: {len(rows)}"), new_x="LMARGIN", new_y="NEXT")
        pdf.ln(3)

        for veh, items in grupos.items():
            pdf.set_font("Helvetica", "B", 11)
            pdf.set_fill_color(230, 230, 230)
            pdf.cell(0, 7, _clean_pdf(f"{veh}  -  {len(items)} servicio(s)"), fill=True, new_x="LMARGIN", new_y="NEXT")
            pdf.ln(1)
            for r in items:
                ini = hora_a_int(r["hora"])
                fin = ini + int(r["duracion"] or 1)
                pdf.set_font("Helvetica", "B", 10)
                pdf.set_x(10)
                pdf.multi_cell(0, 6, _clean_pdf(f"{ini:02d}:00-{fin:02d}:00   {_s(r['tipo'])}   [{_s(r['estado'])}]"),
                               new_x="LMARGIN", new_y="NEXT")
                pdf.set_font("Helvetica", "", 10)
                lineas = []
                if r["cliente"] or r["telefono"]:
                    lineas.append(f"Cliente: {_s(r['cliente'])}   Tel: {_s(r['telefono'])}")
                if r["direccion"]:
                    lineas.append(f"Direccion: {_s(r['direccion'])}")
                if r["detalle"]:
                    lineas.append(f"Detalle: {_s(r['detalle'])}")
                if r["tecnico"] and r["tecnico"] != SIN_TECNICO:
                    lineas.append(f"Tecnico / Repartidor: {_s(r['tecnico'])}")
                for ln in lineas:
                    pdf.set_x(14)
                    pdf.multi_cell(0, 5, _clean_pdf(ln), new_x="LMARGIN", new_y="NEXT")
                pdf.ln(3)
            pdf.ln(3)

        try:
            carpeta = os.path.join(self.app_dir, "Hojas_Ruta", str(d.year))
            os.makedirs(carpeta, exist_ok=True)
            path = os.path.join(carpeta, f"Ruta_{to_iso(d)}.pdf")
            pdf.output(path)
        except OSError as e:
            messagebox.showerror("No se pudo guardar", f"{e}\n\n(Si el PDF esta abierto, cierralo e intenta de nuevo.)")
            return
        _abrir_archivo(path)


def build_agenda_tab(parent, db, on_invoice=None):
    """Punto de entrada compatible con la version anterior. Devuelve el objeto AgendaTab."""
    return AgendaTab(parent, db, on_invoice=on_invoice)
