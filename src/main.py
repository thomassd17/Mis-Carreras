"""
Registro de Carreras de Transporte
----------------------------------
App en un solo archivo hecha con Flet + SQLite3 (sin servidor externo).

Instalación:  pip install flet
Ejecución:    python main.py          (o:  flet run main.py)
"""

import io
import math
import os
import sqlite3
import zipfile
from datetime import datetime
from xml.sax.saxutils import escape

import flet as ft

# En el celular (APK) la carpeta del código es de solo lectura; Flet indica en
# FLET_APP_STORAGE_DATA una carpeta donde sí se puede escribir y que persiste.
# En el PC se usa la carpeta del archivo.
_CARPETA_DATOS = os.getenv("FLET_APP_STORAGE_DATA") or os.path.dirname(os.path.abspath(__file__))
os.makedirs(_CARPETA_DATOS, exist_ok=True)
DB_PATH = os.path.join(_CARPETA_DATOS, "transporte.db")

MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# En Flet recientes `ElevatedButton` fue reemplazado por `Button`.
BotonGuardar = getattr(ft, "ElevatedButton", None) or ft.Button


# ---------------------------------------------------------------------------
# 1. BASE DE DATOS
# ---------------------------------------------------------------------------
def conectar() -> sqlite3.Connection:
    """Abre una conexión a la base de datos SQLite."""
    return sqlite3.connect(DB_PATH)


def iniciar_bd() -> None:
    """Crea las tablas si no existen y actualiza bases de datos antiguas."""
    with conectar() as con:
        # Carreras realizadas (historial).
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS carreras (
                id      INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre  TEXT NOT NULL,
                precio  REAL NOT NULL,
                fecha   TEXT,
                detalle TEXT DEFAULT ''
            )
            """
        )
        # Si la base ya existía sin la columna `detalle`, se la agrega.
        columnas = [fila[1] for fila in con.execute("PRAGMA table_info(carreras)")]
        if "detalle" not in columnas:
            con.execute("ALTER TABLE carreras ADD COLUMN detalle TEXT DEFAULT ''")

        # Lista de carreras frecuentes con su precio (para elegir sin reescribir).
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS destinos (
                id     INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT NOT NULL,
                precio REAL NOT NULL
            )
            """
        )


# --- Carreras (historial) ---------------------------------------------------
def insertar_carrera(nombre: str, precio: float, detalle: str = "") -> None:
    """Guarda una carrera con la fecha y hora actuales."""
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with conectar() as con:
        con.execute(
            "INSERT INTO carreras (nombre, precio, fecha, detalle) VALUES (?, ?, ?, ?)",
            (nombre, precio, fecha, detalle),
        )


def eliminar_carrera(id_carrera: int) -> None:
    """Elimina una carrera por su id."""
    with conectar() as con:
        con.execute("DELETE FROM carreras WHERE id = ?", (id_carrera,))


def eliminar_todas() -> None:
    """Elimina TODAS las carreras guardadas (no toca la lista de frecuentes)."""
    with conectar() as con:
        con.execute("DELETE FROM carreras")


def obtener_carreras() -> list:
    """Devuelve todas las carreras, de la más reciente a la más antigua."""
    with conectar() as con:
        return con.execute(
            "SELECT id, nombre, precio, fecha, detalle FROM carreras ORDER BY id DESC"
        ).fetchall()


def contar_carreras() -> int:
    """Devuelve cuántas carreras hay guardadas."""
    with conectar() as con:
        return con.execute("SELECT COUNT(*) FROM carreras").fetchone()[0]


def obtener_total() -> float:
    """Devuelve la suma de todos los precios (0 si no hay registros)."""
    with conectar() as con:
        resultado = con.execute("SELECT SUM(precio) FROM carreras").fetchone()[0]
    return resultado or 0.0


# --- Lista de carreras frecuentes -------------------------------------------
def guardar_destino(nombre: str, precio: float) -> bool:
    """
    Agrega un destino a la lista de frecuentes.
    Si ya existe uno con el mismo nombre, solo actualiza su precio.
    Devuelve True si fue nuevo y False si se actualizó.
    """
    with conectar() as con:
        fila = con.execute(
            "SELECT id FROM destinos WHERE nombre = ? COLLATE NOCASE", (nombre,)
        ).fetchone()
        if fila:
            con.execute("UPDATE destinos SET precio = ? WHERE id = ?", (precio, fila[0]))
            return False
        con.execute("INSERT INTO destinos (nombre, precio) VALUES (?, ?)", (nombre, precio))
        return True


def eliminar_destino(id_destino: int) -> None:
    """Quita un destino de la lista de frecuentes."""
    with conectar() as con:
        con.execute("DELETE FROM destinos WHERE id = ?", (id_destino,))


def obtener_destinos(filtro: str = "") -> list:
    """Devuelve los destinos ordenados por nombre, filtrados por el texto buscado."""
    with conectar() as con:
        filas = con.execute(
            "SELECT id, nombre, precio FROM destinos ORDER BY nombre COLLATE NOCASE"
        ).fetchall()
    filtro = filtro.strip().casefold()
    if not filtro:
        return filas
    return [f for f in filas if filtro in f[1].casefold()]


# ---------------------------------------------------------------------------
# 2. UTILIDADES
# ---------------------------------------------------------------------------
def leer_precio(texto: str):
    """Convierte el texto a un precio válido (> 0). Devuelve None si no es válido."""
    try:
        precio = float((texto or "").strip().replace(",", "."))
    except ValueError:
        return None
    if math.isnan(precio) or math.isinf(precio) or precio <= 0:
        return None
    return precio


# ---------------------------------------------------------------------------
# 3. EXPORTAR A EXCEL (.xlsx) - sin librerías extra, solo zipfile
# ---------------------------------------------------------------------------
def _xml_seguro(texto) -> str:
    """Quita caracteres que XML no permite y escapa los especiales (& < >)."""
    limpio = "".join(
        c for c in str(texto or "")
        if c in "\t\n\r"
        or 0x20 <= ord(c) <= 0xD7FF
        or 0xE000 <= ord(c) <= 0xFFFD
        or ord(c) >= 0x10000
    )
    return escape(limpio)


def _celda_texto(ref: str, texto, estilo: int = 0) -> str:
    return (f'<c r="{ref}" s="{estilo}" t="inlineStr"><is>'
            f'<t xml:space="preserve">{_xml_seguro(texto)}</t></is></c>')


def _celda_numero(ref: str, valor, estilo: int = 0) -> str:
    return f'<c r="{ref}" s="{estilo}"><v>{float(valor)!r}</v></c>'


def _serial_excel(fecha_texto):
    """Convierte 'YYYY-MM-DD HH:MM:SS' al número de fecha que usa Excel."""
    try:
        fecha = datetime.strptime(fecha_texto, "%Y-%m-%d %H:%M:%S")
    except (ValueError, TypeError):
        return None
    delta = fecha - datetime(1899, 12, 30)
    return delta.days + delta.seconds / 86400


_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_CABECERA_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

_ESTILOS_XML = _CABECERA_XML + f"""<styleSheet xmlns="{_NS}">
<numFmts count="2">
<numFmt numFmtId="164" formatCode="#,##0.00"/>
<numFmt numFmtId="165" formatCode="yyyy\\-mm\\-dd\\ hh:mm:ss"/>
</numFmts>
<fonts count="3">
<font><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><name val="Calibri"/></font>
</fonts>
<fills count="3">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1565C0"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="6">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="164" fontId="2" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""

# Índices de estilo (posición dentro de <cellXfs> de arriba).
_E_CABECERA, _E_PRECIO, _E_FECHA, _E_NEGRITA, _E_TOTAL = 1, 2, 3, 4, 5


def generar_xlsx():
    """
    Genera un archivo Excel (.xlsx) con todas las carreras, de la más antigua a la
    más reciente, y una fila de TOTAL al final. Devuelve los bytes del archivo,
    o None si no hay carreras para exportar.
    """
    carreras = list(reversed(obtener_carreras()))  # orden cronológico
    if not carreras:
        return None

    filas = []
    # Fila 1: encabezados.
    titulos = ["N°", "Fecha", "Nombre / Destino", "Detalle", "Precio ($)"]
    celdas = "".join(_celda_texto(f"{col}1", t, _E_CABECERA)
                     for col, t in zip("ABCDE", titulos))
    filas.append(f'<row r="1">{celdas}</row>')

    # Filas de datos.
    total = 0.0
    for numero, (_id, nombre, precio, fecha, detalle) in enumerate(carreras, start=1):
        r = numero + 1
        serial = _serial_excel(fecha)
        celda_fecha = (_celda_numero(f"B{r}", serial, _E_FECHA) if serial is not None
                       else _celda_texto(f"B{r}", fecha))
        celdas = (
            _celda_numero(f"A{r}", numero)
            + celda_fecha
            + _celda_texto(f"C{r}", nombre)
            + _celda_texto(f"D{r}", detalle)
            + _celda_numero(f"E{r}", precio, _E_PRECIO)
        )
        filas.append(f'<row r="{r}">{celdas}</row>')
        total += precio

    # Fila de TOTAL (deja una fila vacía de separación).
    ultima = len(carreras) + 1
    r_total = ultima + 2
    filas.append(
        f'<row r="{r_total}">'
        + _celda_texto(f"D{r_total}", "TOTAL", _E_NEGRITA)
        + f'<c r="E{r_total}" s="{_E_TOTAL}"><f>SUM(E2:E{ultima})</f><v>{round(total, 2)!r}</v></c>'
        + "</row>"
    )

    hoja = _CABECERA_XML + f"""<worksheet xmlns="{_NS}">
<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
<cols>
<col min="1" max="1" width="6" customWidth="1"/>
<col min="2" max="2" width="20" customWidth="1"/>
<col min="3" max="3" width="32" customWidth="1"/>
<col min="4" max="4" width="36" customWidth="1"/>
<col min="5" max="5" width="14" customWidth="1"/>
</cols>
<sheetData>{"".join(filas)}</sheetData>
<autoFilter ref="A1:E{ultima}"/>
</worksheet>"""

    libro = _CABECERA_XML + (
        f'<workbook xmlns="{_NS}" xmlns:r="{_NS_REL}"><sheets>'
        f'<sheet name="Carreras" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )
    rels_libro = _CABECERA_XML + (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{_NS_REL}/worksheet" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId2" Type="{_NS_REL}/styles" Target="styles.xml"/>'
        "</Relationships>"
    )
    rels_raiz = _CABECERA_XML + (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{_NS_REL}/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    tipos = _CABECERA_XML + (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        "</Types>"
    )

    salida = io.BytesIO()
    with zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", tipos)
        z.writestr("_rels/.rels", rels_raiz)
        z.writestr("xl/workbook.xml", libro)
        z.writestr("xl/_rels/workbook.xml.rels", rels_libro)
        z.writestr("xl/styles.xml", _ESTILOS_XML)
        z.writestr("xl/worksheets/sheet1.xml", hoja)
    return salida.getvalue()


def nombre_archivo_excel() -> str:
    """Nombre sugerido para el archivo, con la fecha y hora actuales."""
    return f"carreras_{datetime.now():%Y-%m-%d_%H%M}.xlsx"


# ---------------------------------------------------------------------------
# 4. INTERFAZ
# ---------------------------------------------------------------------------
def main(page: ft.Page) -> None:
    page.title = "Mis Carreras"
    page.padding = 16
    page.scroll = None  # El scroll lo maneja el ListView.

    iniciar_bd()

    # --- Controles del formulario principal ---------------------------------
    txt_total = ft.Text("Total: $ 0.00", size=30, weight=ft.FontWeight.BOLD,
                        color=ft.Colors.WHITE)

    tarjeta_total = ft.Container(
        content=ft.Column(
            [
                ft.Text("TOTAL ACUMULADO", size=12, color="#B3FFFFFF"),
                txt_total,
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=4,
        ),
        bgcolor=ft.Colors.BLUE_700,
        border_radius=16,
        padding=20,
        alignment=ft.Alignment(0, 0),
        width=float("inf"),
    )

    campo_nombre = ft.TextField(label="Nombre / Destino", expand=True)
    campo_precio = ft.TextField(
        label="Precio ($)",
        keyboard_type=ft.KeyboardType.NUMBER,
        expand=True,
    )
    campo_detalle = ft.TextField(label="Detalle (opcional)")

    lista = ft.ListView(expand=True, spacing=6)

    # --- Servicios (guardar y compartir archivos) ---------------------------
    def registrar_servicio(servicio) -> None:
        """Agrega un servicio de Flet (FilePicker, Share) a la página."""
        if hasattr(page, "services"):
            page.services.append(servicio)
        else:  # Flet anterior
            page.overlay.append(servicio)

    selector_archivos = ft.FilePicker()
    registrar_servicio(selector_archivos)

    compartidor = ft.Share() if hasattr(ft, "Share") else None
    if compartidor is not None:
        registrar_servicio(compartidor)

    # --- Utilidades de pantalla ---------------------------------------------
    def mostrar_mensaje(texto: str, error: bool = False) -> None:
        """Muestra un SnackBar (compatible con distintas versiones de Flet)."""
        snack = ft.SnackBar(
            ft.Text(texto),
            bgcolor=ft.Colors.RED_700 if error else ft.Colors.GREEN_700,
        )
        if hasattr(page, "show_dialog"):  # Flet >= 0.80
            page.show_dialog(snack)
        else:  # Flet anterior
            page.open(snack)

    def abrir_dialogo(dialogo) -> None:
        if hasattr(page, "show_dialog"):
            page.show_dialog(dialogo)
        else:
            page.open(dialogo)

    def cerrar_dialogo(dialogo) -> None:
        if hasattr(page, "pop_dialog"):
            page.pop_dialog()
        else:
            page.close(dialogo)

    def confirmar(titulo: str, mensaje: str, texto_si: str, accion) -> None:
        """
        Muestra una ventana de confirmación con Cancelar / <texto_si>.
        Solo si se presiona <texto_si> se ejecuta `accion()`.
        """
        def al_confirmar(e) -> None:
            cerrar_dialogo(dialogo)
            accion()

        dialogo = ft.AlertDialog(
            modal=True,
            title=ft.Text(titulo),
            content=ft.Text(mensaje),
            actions=[
                ft.TextButton("Cancelar", on_click=lambda e: cerrar_dialogo(dialogo)),
                ft.TextButton(
                    texto_si,
                    on_click=al_confirmar,
                    style=ft.ButtonStyle(color=ft.Colors.RED_700),
                ),
            ],
        )
        abrir_dialogo(dialogo)

    def refrescar() -> None:
        """Recarga el total y la lista de carreras desde la base de datos."""
        txt_total.value = f"Total: $ {obtener_total():,.2f}"

        lista.controls.clear()
        carreras = obtener_carreras()
        if not carreras:
            lista.controls.append(
                ft.Text("Aún no hay carreras registradas.",
                        italic=True, color=ft.Colors.GREY)
            )
        for id_carrera, nombre, precio, fecha, detalle in carreras:
            lista.controls.append(crear_item(id_carrera, nombre, precio, fecha, detalle))
        page.update()

    def crear_item(id_carrera: int, nombre: str, precio: float,
                   fecha: str, detalle: str) -> ft.Control:
        """Construye una fila de la lista para una carrera."""
        detalle = (detalle or "").strip()
        texto = f"$ {precio:,.2f}  •  {fecha}"
        if detalle:
            texto += f"\n{detalle}"
        return ft.Card(
            content=ft.ListTile(
                leading=ft.Icon(ft.Icons.LOCAL_TAXI, color=ft.Colors.BLUE_700),
                title=ft.Text(nombre, weight=ft.FontWeight.W_600),
                subtitle=ft.Text(texto, max_lines=2),
                is_three_line=bool(detalle),
                trailing=ft.IconButton(
                    icon=ft.Icons.DELETE_OUTLINE,
                    icon_color=ft.Colors.RED_600,
                    tooltip="Eliminar",
                    on_click=lambda e, i=id_carrera, n=nombre, p=precio: al_eliminar(i, n, p),
                ),
            )
        )

    # --- Eventos del formulario principal -----------------------------------
    def al_guardar(e) -> None:
        nombre = (campo_nombre.value or "").strip()
        detalle = (campo_detalle.value or "").strip()

        if not nombre or not (campo_precio.value or "").strip():
            mostrar_mensaje("Completa el nombre y el precio.", error=True)
            return

        precio = leer_precio(campo_precio.value)
        if precio is None:
            mostrar_mensaje("El precio debe ser un número válido mayor que 0.", error=True)
            return

        insertar_carrera(nombre, precio, detalle)
        campo_nombre.value = ""
        campo_precio.value = ""
        campo_detalle.value = ""
        refrescar()
        mostrar_mensaje("Carrera guardada.")

    # --- Borrar una carrera (pide confirmación) -----------------------------
    def al_eliminar(id_carrera: int, nombre: str, precio: float) -> None:
        def borrar() -> None:
            eliminar_carrera(id_carrera)
            refrescar()
            mostrar_mensaje("Carrera eliminada.")

        confirmar(
            "¿Eliminar esta carrera?",
            f"{nombre}\n$ {precio:,.2f}\n\nEsta acción no se puede deshacer.",
            "Eliminar",
            borrar,
        )

    # --- Borrar todas las carreras (pide confirmación) ----------------------
    def al_borrar_todo(e) -> None:
        cantidad = contar_carreras()
        if cantidad == 0:
            mostrar_mensaje("No hay carreras para borrar.", error=True)
            return

        def borrar_todo() -> None:
            eliminar_todas()
            refrescar()
            mostrar_mensaje("Se borraron todas las carreras.")

        confirmar(
            "¿Borrar TODAS las carreras?",
            f"Se eliminarán las {cantidad} carreras guardadas y el total volverá a $ 0.00."
            "\n\nEsta acción no se puede deshacer.",
            "Sí, borrar todo",
            borrar_todo,
        )

    # --- Exportar a Excel: descargar y compartir ----------------------------
    async def al_descargar(e) -> None:
        datos = generar_xlsx()
        if datos is None:
            mostrar_mensaje("No hay carreras para exportar.", error=True)
            return
        try:
            ruta = await selector_archivos.save_file(
                dialog_title="Guardar tabla de carreras",
                file_name=nombre_archivo_excel(),
                file_type=ft.FilePickerFileType.CUSTOM,
                allowed_extensions=["xlsx"],
                src_bytes=datos,
            )
        except Exception as error:  # noqa: BLE001 - se informa al usuario
            mostrar_mensaje(f"No se pudo guardar el archivo: {error}", error=True)
            return
        if page.web:
            # En el navegador la descarga arranca sola y no devuelve una ruta.
            mostrar_mensaje("Descarga iniciada.")
        elif ruta:
            mostrar_mensaje("Archivo Excel guardado.")
        else:
            mostrar_mensaje("Guardado cancelado.", error=True)

    async def al_compartir(e) -> None:
        datos = generar_xlsx()
        if datos is None:
            mostrar_mensaje("No hay carreras para compartir.", error=True)
            return
        archivo = ft.ShareFile.from_bytes(
            datos, mime_type=MIME_XLSX, name=nombre_archivo_excel()
        )
        try:
            await compartidor.share_files(
                [archivo], title="Mis carreras", subject="Mis carreras",
                text="Tabla de carreras",
            )
        except Exception as error:  # noqa: BLE001 - se informa al usuario
            mostrar_mensaje(f"No se pudo compartir: {error}", error=True)

    btn_descargar = BotonGuardar(
        "Descargar Excel", icon=ft.Icons.DOWNLOAD, on_click=al_descargar,
    )
    btn_compartir = BotonGuardar(
        "Compartir", icon=ft.Icons.SHARE, on_click=al_compartir,
        visible=compartidor is not None,
    )
    btn_borrar_todo = BotonGuardar(
        "Borrar todo",
        icon=ft.Icons.DELETE_FOREVER,
        bgcolor=ft.Colors.RED_700,
        color=ft.Colors.WHITE,
        on_click=al_borrar_todo,
    )

    # --- Lista de carreras frecuentes (ventana con buscador) ----------------
    campo_buscar = ft.TextField(
        label="Buscar en mi lista",
        on_change=lambda e: refrescar_catalogo(),
    )
    lista_catalogo = ft.ListView(expand=True, spacing=2)
    txt_aviso = ft.Text("", size=12)
    nuevo_nombre = ft.TextField(label="Nombre / Destino", expand=True)
    nuevo_precio = ft.TextField(
        label="Precio ($)",
        keyboard_type=ft.KeyboardType.NUMBER,
        expand=True,
    )

    def refrescar_catalogo() -> None:
        """Muestra los destinos guardados que coinciden con lo escrito en el buscador."""
        lista_catalogo.controls.clear()
        destinos = obtener_destinos(campo_buscar.value or "")
        if not destinos:
            lista_catalogo.controls.append(
                ft.Text("No hay resultados.", italic=True, color=ft.Colors.GREY)
            )
        for id_destino, nombre, precio in destinos:
            lista_catalogo.controls.append(
                ft.ListTile(
                    title=ft.Text(nombre, weight=ft.FontWeight.W_600),
                    subtitle=ft.Text(f"$ {precio:,.2f}"),
                    on_click=lambda e, n=nombre, p=precio: elegir_destino(n, p),
                    trailing=ft.IconButton(
                        icon=ft.Icons.DELETE_OUTLINE,
                        icon_color=ft.Colors.RED_600,
                        tooltip="Quitar de la lista",
                        on_click=lambda e, i=id_destino: quitar_destino(i),
                    ),
                )
            )
        page.update()

    def elegir_destino(nombre: str, precio: float) -> None:
        """Pasa el destino elegido al formulario principal."""
        campo_nombre.value = nombre
        campo_precio.value = f"{precio:.2f}"
        cerrar_dialogo(dlg_catalogo)
        page.update()

    def quitar_destino(id_destino: int) -> None:
        eliminar_destino(id_destino)
        refrescar_catalogo()

    def al_agregar_destino(e) -> None:
        nombre = (nuevo_nombre.value or "").strip()
        precio = leer_precio(nuevo_precio.value)

        if not nombre or precio is None:
            txt_aviso.value = "Escribe un nombre y un precio válido mayor que 0."
            txt_aviso.color = ft.Colors.RED_700
            page.update()
            return

        es_nuevo = guardar_destino(nombre, precio)
        nuevo_nombre.value = ""
        nuevo_precio.value = ""
        txt_aviso.value = "Agregado a la lista." if es_nuevo else "Ese nombre ya existía: precio actualizado."
        txt_aviso.color = ft.Colors.GREEN_700
        refrescar_catalogo()

    dlg_catalogo = ft.AlertDialog(
        modal=True,
        title=ft.Text("Mi lista de carreras"),
        content=ft.Column(
            [
                campo_buscar,
                ft.Container(content=lista_catalogo, height=200),
                ft.Divider(),
                ft.Text("Agregar a la lista", weight=ft.FontWeight.BOLD),
                ft.Row([nuevo_nombre, nuevo_precio]),
                BotonGuardar("Agregar a la lista", icon=ft.Icons.ADD, on_click=al_agregar_destino),
                txt_aviso,
            ],
            tight=True,
            width=320,
        ),
        actions=[ft.TextButton("Cerrar", on_click=lambda e: cerrar_dialogo(dlg_catalogo))],
    )

    def al_abrir_catalogo(e) -> None:
        campo_buscar.value = ""
        txt_aviso.value = ""
        refrescar_catalogo()
        abrir_dialogo(dlg_catalogo)

    # --- Vista principal ----------------------------------------------------
    page.add(
        tarjeta_total,
        BotonGuardar(
            "Elegir de mi lista",
            icon=ft.Icons.LIST_ALT,
            bgcolor=ft.Colors.BLUE_GREY_700,
            color=ft.Colors.WHITE,
            on_click=al_abrir_catalogo,
            width=float("inf"),
            height=44,
        ),
        ft.Row([campo_nombre, campo_precio]),
        campo_detalle,
        BotonGuardar(
            "Guardar Carrera",
            icon=ft.Icons.SAVE,
            on_click=al_guardar,
            width=float("inf"),
            height=48,
        ),
        ft.Divider(),
        ft.Row(
            [btn_descargar, btn_compartir, btn_borrar_todo],
            wrap=True,
            spacing=8,
            run_spacing=8,
        ),
        lista,
    )

    refrescar()


if __name__ == "__main__":
    # `ft.run` en Flet recientes; `ft.app` en versiones anteriores.
    if hasattr(ft, "run"):
        ft.run(main)
    else:
        ft.app(target=main)

