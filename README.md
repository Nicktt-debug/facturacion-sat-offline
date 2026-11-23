# 🧾 Mini-ERP & TPV Offline en Python (Facturación, SAT, Stock y Agenda)

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)
![Database](https://img.shields.io/badge/Base%20de%20Datos-SQLite%20Offline-lightgrey?logo=sqlite)
![Platform](https://img.shields.io/badge/Plataforma-Windows%20%7C%20Linux-green)
![License](https://img.shields.io/badge/Licencia-MIT-orange)

Aplicación de escritorio **100% offline y sin cuotas mensuales**, desarrollada en Python (`Tkinter` + `SQLite` + `fpdf2` + `openpyxl`). Diseñada originalmente para tiendas de electrodomésticos nuevos/segunda mano y talleres de reparación (SAT), pero **totalmente adaptable a cualquier comercio o autónomo** desde su menú de configuración.

---

## ✨ Funcionalidades Principales (8 Módulos)

1. **Facturas, Tickets y Albaranes**
   - Emite **Facturas A4**, **Tickets térmicos (80mm)**, **Facturas de Reparación**, **Presupuestos**, **Albaranes de Entrega** y **Recibos**.
   - Cálculo automático en modo **Bruto** o **Neto** con soporte para **21% IVA**, **4% REBU**, **0% Exento** o **Sin IVA** (sin desglose fiscal en el PDF).
   - **Pagos parciales** (calcula automáticamente lo entregado y lo pendiente), descuentos (`€` o `%`) y garantías editables (6, 3, 2 meses, 1 o 3 años).
   - Opción de **Solo Guardar (sin PDF)** para ahorrar espacio, **Generar PDF**, **Generar + Imprimir** o **Ver/Imprimir Ticket**.
   - Sincronización directa con la **Agenda de Entregas** y conversión de presupuestos a factura en 1 clic.

2. **SAT / Partes de Reparación**
   - Registro de entrada de aparatos con datos del cliente (Nombre, DNI/NIF, dirección, teléfono), avería y **técnico asignado**.
   - Genera **Resguardo de Entrega en PDF** listo para imprimir.
   - Control de estados (*En taller, Presupuestado, Reparado, Entregado*), aviso directo por **WhatsApp** y conversión a **Factura de Reparación** en 1 clic.

3. **Compras a Proveedores y Gastos Fijos**
   - Registro de facturas recibidas con cálculo de IVA soportado y clasificación por tipo de gasto (*Mercadería, Recambios, Suministros, Alquiler, Nóminas, Furgonetas...*).
   - Autoguardado de proveedores con su NIF y opción de **adjuntar el PDF original** de la factura local.

4. **Resumen Fiscal, TPV y Cierre de Caja**
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
   - Genera automáticamente el **Contrato / Recibo de Compraventa en PDF** (diseño en blanco y negro para ahorrar tinta) y actualiza un registro en **Excel (`compras_particulares.xlsx`)**.

7. **Agenda de Entregas, Vehículos y Técnicos (11:00 – 21:00)**
   - Cuadrante horario con código de colores (*verde libre / rojo ocupado*, permitiendo siempre añadir más de un servicio en la misma hora).
   - Asignación de vehículos, técnicos/repartidores y duración por servicio, con aviso de solapamientos.
   - **Vista semanal**, recordatorios por **WhatsApp**, apertura de dirección en **Google Maps**, paso a factura y **Hoja de Ruta en PDF**.

8. **Stock, Etiquetas, Caja Chica y RMA**
   - Inventario por tienda con número de serie y generación de **Etiquetas de Precio** en formato **A5 (escaparate)** y **80mm (ticket)**.
   - **Caja Chica** para entradas y salidas de efectivo en tienda.
   - Seguimiento de **RMA / Devoluciones a proveedores** pendientes de abono.

---

## 🔒 Privacidad y Copias de Seguridad
- **100% Local:** Todos los datos se guardan en tu propio ordenador (`gestion.db`), sin nube ni servidores externos.
- **Backups automáticos:** Al cerrar la aplicación se crea una copia de seguridad en la carpeta `Backups/` (conservando las últimas 20).
- **Restauración en 1 clic:** Incluye botón **Restaurar Backup** en la barra superior.

---

## 🚀 Instalación y Uso

### Requisitos
- Python **3.10 o superior** (recomendado Python 3.13).

### Pasos
1. Clona o descarga este repositorio:
   ```bash
   git clone [https://github.com/TU_USUARIO/TU_REPOSITORIO.git](https://github.com/TU_USUARIO/TU_REPOSITORIO.git)
   cd TU_REPOSITORIO