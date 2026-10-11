"""
Registro de Carreras de Transporte
----------------------------------
App en un solo archivo hecha con Flet + SQLite3 (sin servidor externo).

Instalación:  pip install flet
Ejecución:    python main.py          (o:  flet run main.py)
"""

import math
import os
import sqlite3
from datetime import datetime

import flet as ft

# En el celular (APK) la carpeta del código es de solo lectura; Flet indica en
# FLET_APP_STORAGE_DATA una carpeta donde sí se puede escribir y que persiste.
# En el PC se usa la carpeta del archivo.
_CARPETA_DATOS = os.getenv("FLET_APP_STORAGE_DATA") or os.path.dirname(os.path.abspath(__file__))
os.makedirs(_CARPETA_DATOS, exist_ok=True)
DB_PATH = os.path.join(_CARPETA_DATOS, "transporte.db")


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
# 3. INTERFAZ
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
                    on_click=lambda e, i=id_carrera: al_eliminar(i),
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

    def al_eliminar(id_carrera: int) -> None:
        eliminar_carrera(id_carrera)
        refrescar()
        mostrar_mensaje("Carrera eliminada.")

    # --- Borrar todas las carreras (requiere marcar la casilla) -------------
    def al_cambiar_confirmacion(e) -> None:
        """Habilita el botón de borrado total solo si la casilla está marcada."""
        btn_borrar_todo.disabled = not chk_confirmar.value
        page.update()

    def al_borrar_todo(e) -> None:
        if not chk_confirmar.value:
            mostrar_mensaje("Marca la casilla para confirmar.", error=True)
            return
        eliminar_todas()
        # Se desmarca la casilla para que la próxima vez haya que confirmar de nuevo.
        chk_confirmar.value = False
        btn_borrar_todo.disabled = True
        refrescar()
        mostrar_mensaje("Se borraron todas las carreras.")

    chk_confirmar = ft.Checkbox(
        label="Confirmo que quiero borrar todo",
        value=False,
        on_change=al_cambiar_confirmacion,
    )
    btn_borrar_todo = BotonGuardar(
        "Borrar todo",
        icon=ft.Icons.DELETE_FOREVER,
        bgcolor=ft.Colors.RED_700,
        color=ft.Colors.WHITE,
        disabled=True,
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
            [chk_confirmar, btn_borrar_todo],
            wrap=True,
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
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
