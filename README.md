# 🧾 Mini-ERP y TPV Offline Gratis en Python | Facturación, SAT, Stock y Agenda

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![SQLite](https://img.shields.io/badge/Base%20de%20Datos-SQLite%20100%25%20Offline-003B57?logo=sqlite&logoColor=white)
![OS](https://img.shields.io/badge/Plataforma-Windows%20%7C%20Linux-28a745)
![License](https://img.shields.io/badge/Licencia-MIT%20(Gratis)-orange)

**Software de gestión comercial (Mini-ERP / TPV) 100% offline, gratuito y sin cuotas mensuales**, desarrollado en Python (`Tkinter` + `SQLite` + `fpdf2` + `openpyxl`).

Pensado para **autónomos, tiendas de electrodomésticos (nuevos y segunda mano / REBU), servicios técnicos (SAT), talleres de reparación y pequeños comercios** que buscan facturar, controlar su caja, gestionar entregas y llevar su contabilidad en local sin depender de la nube.

---

## ✨ Funcionalidades Principales (8 Módulos)

1. **Facturas, Tickets y Albaranes**
   - Emite **Facturas A4**, **Tickets térmicos (80mm)**, **Facturas de Reparación**, **Presupuestos**, **Albaranes de Entrega** y **Recibos**.
   - Cálculo automático en modo **Bruto** o **Neto** con soporte para **21% IVA**, **4% REBU**, **0% Exento** o **Sin IVA** (sin desglose fiscal en el PDF).
   - **Pagos parciales** (calcula automáticamente lo entregado y lo pendiente), descuentos (`€` o `%`) y garantías editables (6, 3, 2 meses, 1 o 3 años).
   - Opción de **Solo Guardar (sin PDF)**, **Generar PDF**, **Generar + Imprimir** o **Ver/Imprimir Ticket**.
   - Sincronización directa con la **Agenda de Entregas** y conversión de presupuestos a factura en 1 clic.

2. **SAT / Partes de Reparación**
   - Registro de entrada de aparatos con datos del cliente (Nombre, DNI/NIF, dirección, teléfono), avería y **técnico asignado**.
   - Genera **Resguardo de Entrega en PDF** listo para imprimir.
   - Control de estados (*En taller, Presupuestado, Reparado, Entregado*), aviso directo por **WhatsApp** y conversión a **Factura de Reparación** en 1 clic.

3. **Compras a Proveedores y Gastos Fijos**
   - Registro de facturas recibidas con cálculo de IVA soportado y clasificación por tipo de gasto (*Mercadería, Recambios, Suministros, Alquiler, Nóminas, Furgonetas...*).
   - Autoguardado de proveedores con su NIF y opción de **adjuntar el PDF original** de la factura local.

4. **Resumen Fiscal (Modelo 303), TPV y Cierre de Caja**
   - Cálculo automático por **trimestres (1T–4T) y año**: Ventas 21%, Ventas REBU, IVA Repercutido, IVA Soportado, **estimación Modelo 303** y **Beneficio Neto**.
   - **Cierre de Caja diario** por tienda y método de pago (*Efectivo, Tarjeta, Bizum, Transferencia*).
   - **Cuadre de TPV (datáfonos)** mensual o anual por tienda.
   - Exportación directa a **Excel (`.xlsx`)** y **PDF** para entregar a la gestoría.

5. **Buscador Avanzado, Garantías y Abonos**
   - Búsqueda unificada en **Ventas, Compras y SAT** con filtros por texto, rango de fechas y precio mínimo/máximo.
   - **Verificador automático de garantía** (indica días restantes o si ha caducado).
   - Emisión de **Facturas Rectificativas (Abonos)** en negativo, registro de **pagos parciales adicionales** y apertura de partes SAT desde una factura vendida.

6. **Compras a Particulares (Libro Registro / REBU)**
   - Registro de compras a particulares con autocompletado de códigos postales y ciudades.
   - Genera automáticamente el **Contrato / Recibo de Compraventa en PDF** (en blanco y negro para ahorrar tinta) y actualiza un registro en **Excel (`compras_particulares.xlsx`)**.

7. **Agenda de Entregas, Vehículos y Técnicos (11:00 – 21:00)**
   - Cuadrante horario con código de colores (*verde libre / rojo ocupado*, permitiendo siempre añadir más de un servicio en la misma hora).
   - Asignación de vehículos, técnicos/repartidores y duración por servicio, con aviso de solapamientos.
   - **Vista semanal**, recordatorios por **WhatsApp**, apertura de dirección en **Google Maps**, paso a factura y **Hoja de Ruta en PDF**.

8. **Stock, Etiquetas, Caja Chica y RMA**
   - Inventario por tienda con número de serie y generación de **Etiquetas de Precio** en formato **A5 (escaparate)** y **80mm (ticket)**.
   - **Caja Chica** para entradas y salidas de efectivo en tienda.
   - Seguimiento de **RMA / Devoluciones a proveedores** pendientes de abono.

---

## 💻 Requisitos Previos (Desde 0)

1. **Python 3.10 o superior** (recomendado Python 3.13):
   - Descárgalo desde [python.org/downloads](https://www.python.org/downloads/).
   - **IMPORTANTE en Windows:** Marca la casilla **"Add python.exe to PATH"** antes de pulsar *Install Now*.
2. **Visual Studio Code (VS Code):**
   - Descárgalo desde [code.visualstudio.com](https://code.visualstudio.com/).
   - Al abrir VS Code, ve a **Extensiones** (`Ctrl + Shift + X`), busca **Python** (de Microsoft) e instálala.

---

## 🚀 Instalación y Ejecución Paso a Paso en VS Code

### Paso 1: Descargar el código
- **Opción A (ZIP - Sin Git):** Pulsa el botón verde **`<> Code`** → **`Download ZIP`** y descomprime la carpeta en tu PC (ej. `C:\Gestion_Tienda`).
- **Opción B (Con Git):**
  ```bash
  git clone [https://github.com/TU_USUARIO/TU_REPOSITORIO.git](https://github.com/TU_USUARIO/TU_REPOSITORIO.git)
  cd TU_REPOSITORIO
  ```

### Paso 2: Abrir el proyecto en VS Code
1. Abre **Visual Studio Code**.
2. Ve a **Archivo → Abrir carpeta...** (`File → Open Folder...`) y selecciona la carpeta donde descomprimiste el proyecto.

### Paso 3: Seleccionar el intérprete de Python
1. Pulsa `Ctrl + Shift + P` en VS Code.
2. Escribe `Python: Select Interpreter` y pulsa **Enter**.
3. Elige tu versión de Python instalada (por ejemplo, `Python 3.13`).

### Paso 4: Abrir el terminal e instalar dependencias (`pip`)
1. En VS Code, ve a **Terminal → Nuevo terminal** (o pulsa `Ctrl + Ñ`).
2. Ejecuta estos dos comandos para actualizar `pip` e instalar `fpdf2` y `openpyxl`:
   ```powershell
   python -m pip install --upgrade pip
   python -m pip install -r requirements.txt
   ```

### Paso 5: Probar y ejecutar la aplicación
En el mismo terminal de VS Code, ejecuta:
```powershell
python fac.py
```
Si todo está correcto, se abrirá la ventana principal del programa.

---

## ⚙️ Cómo Personalizarlo para tu Negocio y Ciudad

### A) Personalización rápida sin tocar código (Desde la App)
1. Abre la aplicación y pulsa el botón **Configuración** (esquina superior derecha).
2. Cambia el **Nombre de tu empresa**, **NIF/CIF**, **Teléfono** y **Email**.
3. En **Tiendas (1 por línea)**, escribe las direcciones reales de tus locales.
4. En **Categorías**, **Vehículos** y **Técnicos**, pon los de tu negocio.
5. Pulsa **Guardar Cambios**. Todo se guardará en tu archivo local `config.json`.
6. **Logo en facturas:** Guarda una imagen llamada `logo.png` en la misma carpeta que `fac.py` y aparecerá automáticamente en tus PDFs.

### B) Personalización avanzada en el código (Desde VS Code)
- **Locales, categorías, impuestos y garantías por defecto (`fac.py`):**
  - Edita `DEFAULT_CONFIG` para cambiar los valores iniciales de empresa, tiendas y productos.
  - Edita `TAX_OPTIONS` si necesitas otros tipos de IVA.
  - Edita `WARRANTY_OPTIONS` para cambiar los textos o días de garantía por defecto.
- **Códigos Postales de otra ciudad o provincia (`tab_particulares.py`):**
  - Por defecto incluye los códigos postales de la provincia de Barcelona.
  - Abre `tab_particulares.py`, busca `_build_bcn_cp_dict()` y sustituye los códigos postales y ciudades por los de tu provincia (ej. `"28001": "Madrid"`, `"46001": "Valencia"`).
  - En el mismo archivo puedes editar las listas `NACIONALIDADES` y `MARCAS`.
- **Horario de la Agenda (`tab_agenda.py`):**
  - Por defecto cubre de **11:00 a 21:00**. Cambia `HORA_INICIO` y `HORA_FIN` al principio de `tab_agenda.py` si necesitas otro horario.

---

## 📦 Compilar a Ejecutable `.exe` desde el Terminal de VS Code

Cuando hayas probado tus cambios con `python fac.py`, puedes crear un único archivo `.exe` para abrirlo con doble clic en cualquier Windows (sin necesitar Python instalado).

### 1. Instalar PyInstaller
```powershell
python -m pip install pyinstaller
```

### 2. Comprobar que no hay errores antes de compilar
```powershell
python -m py_compile fac.py tab_agenda.py tab_particulares.py tab_extra.py
```
*(Si no aparece ningún texto, no hay errores de sintaxis).*

### 3. Compilar el `.exe`
```powershell
python -m PyInstaller --onefile --windowed --name "Gestion_Tienda" fac.py
```
- `--onefile`: Empaqueta todo en un solo archivo `.exe`.
- `--windowed`: Oculta la consola negra al abrir la aplicación.
- `--name "Gestion_Tienda"`: Nombre final del programa.
- *(Opcional)* Para añadir tu icono, guarda `icono.ico` en la carpeta y añade `--icon="icono.ico"` al comando.

### 4. Usar tu `.exe`
1. Al terminar (30–60 segundos), entra en la carpeta **`dist/`** y busca **`Gestion_Tienda.exe`**.
2. Copia ese `.exe` a la carpeta definitiva donde quieras trabajar (ej. `C:\Mi_Negocio\`). Al abrirlo, creará automáticamente `gestion.db`, `config.json` y las carpetas de PDFs a su lado.
3. Ya puedes borrar las carpetas temporales `build/`, `dist/` y el archivo `.spec`.

> **Nota sobre Windows Defender:** Al ser un `.exe` compilado por ti sin firma digital de pago, Windows puede mostrar un aviso azul (*"Windows protegió su PC"*). Pulsa **Más información → Ejecutar de todas formas**.

---

## 🛠️ Solución de Errores Comunes

- **`ModuleNotFoundError: No module named 'fpdf'` (o `'openpyxl'`):**
  Tienes varias versiones de Python instaladas. Fuerza la misma versión para instalar y ejecutar:
  ```powershell
  py -3.13 -m pip install -r requirements.txt
  py -3.13 fac.py
  ```
- **El `.exe` se cierra solo al abrirlo:**
  Compílalo una vez quitando `--windowed` y ejecútalo desde el terminal para ver el error exacto.
- **Error `PermissionError` al guardar Excel:**
  Cierra el archivo `.xlsx` en Excel antes de guardar un nuevo registro o exportar.

---

## 📁 Estructura del Proyecto y Privacidad

```text
├── fac.py                 # Aplicación principal (Facturación, SAT, Compras, Fiscal, Buscador)
├── tab_agenda.py          # Agenda 11h-21h, Vista Semanal y Hojas de Ruta en PDF
├── tab_particulares.py    # Compras a Particulares (Códigos Postales, Contratos PDF + Excel)
├── tab_extra.py           # Stock, Etiquetas A5/80mm, Caja Chica y RMA
├── requirements.txt       # Dependencias (fpdf2, openpyxl)
├── .gitignore             # Protege gestion.db, config.json y PDFs de clientes
└── README.md              # Documentación
```

- **100% Privado y Offline:** Ningún dato sale de tu ordenador; todo se guarda en `gestion.db`.
- **Backups automáticos:** Al cerrar el programa se guarda una copia en `Backups/` (mantiene las últimas 20) y puedes restaurarla en 1 clic desde **Restaurar Backup**.

---

## 📄 Licencia
Distribuido bajo licencia **MIT**. Libre para usar, modificar y adaptar a cualquier negocio.
