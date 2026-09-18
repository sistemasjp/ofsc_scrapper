#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ofsc_scraper.py
================
Scraper automatizado para la Consola de despacho de Oracle Field Service
(claro-ec.fs.ocs.oraclecloud.com) que recolecta, para una fecha dada y una
lista de cuadrillas, los materiales/equipos consumidos por actividad
"finalizada", incluyendo el CODIGO SAP (cuando OFSC lo expone), cantidad y
número de serie.

CÓMO FUNCIONA (en corto)
------------------------
En vez de leer la tabla visual (HTML) de "Lista de inventarios" -que tarda en
renderizar y puede cambiar de layout-, este script intercepta las respuestas
JSON que la propia aplicación ya pide a su backend interno:

  1) ?m=Grid&a=get&itype=manage&output=ajax
     -> Se dispara al seleccionar una cuadrilla en el árbol izquierdo.
     Devuelve TODAS las actividades del día para esa cuadrilla
     (activitiesRows), con estado (astatus), tipo (aworktype), orden
     (appt_number), cliente (cname), zona/ciudad (aworkzone), etc. Con esto
     sabemos qué actividades abrir SIN necesidad de abrir cada una solo
     para ver su estado.

  2) ?m=sync&a=write
     -> Canal genérico de sincronización offline-first de la app (no un
     endpoint dedicado por actividad). Su respuesta trae delta.Inventory:
     cada material/equipo "en foco" en ese momento, con sus propiedades
     internas y el campo inv_aid indicando a qué actividad pertenece. La
     propiedad "186" (cuando existe) es el CODIGO SAP del material. Como el
     inventario de una actividad puede llegar en una respuesta anterior a
     cuando se abre esa actividad puntual, el script acumula TODAS las
     respuestas de este canal durante toda la sesión (ver clase
     InventoryCache) y consulta ese acumulado por inv_aid, en vez de
     esperar una respuesta nueva cada vez.

Para EQUIPOS (ONT, Mesh, STB, decos) OFSC casi nunca expone el código SAP en
esa respuesta -sí expone modelo, marca y serie-, y esto se considera normal:
el número de SERIE ya identifica la unidad de forma suficiente, así que no
se advierte ni se busca en ninguna tabla externa cuando falta. Para
MATERIALES sí se exige el código SAP (rara vez tienen serial propio).

IMPORTANTE - CREDENCIALES
--------------------------
El script en sí NUNCA escribe tu usuario/clave a ningún archivo. Por defecto
(--auto-login NO activado) abre el navegador, te deja iniciar sesión
MANUALMENTE (incluye SSO/MFA si tu empresa lo usa) y continúa solo cuando tú
presionas ENTER en la terminal. Es el modo recomendado.

Si tu login es un formulario simple usuario+clave sin SSO, puedes activar
--auto-login: en ese caso el script pide el usuario/clave por consola con
`getpass` (no quedan visibles ni se guardan) o los toma de las variables de
entorno OFSC_USERNAME / OFSC_PASSWORD. La forma más cómoda de dejarlas
listas (para uso desatendido, cron, u otra persona del equipo) es copiar
`.env.example` a `.env` junto a este script y completarlas ahí: se cargan
solas al arrancar (python-dotenv). Ese `.env` lo creas y controlas tú —
nunca se sube a control de versiones (está en .gitignore) y nadie más que
quien lo escribió debería tener acceso a él. Antes de usar --auto-login
DEBES verificar/ajustar los selectores LOGIN_USER_SELECTORS /
LOGIN_PASS_SELECTORS / LOGIN_SUBMIT_SELECTORS más abajo, inspeccionando el
formulario real con F12 -> clic derecho sobre el campo -> "Inspeccionar".

INSTALACIÓN
-----------
    pip install -r requirements.txt
    playwright install chromium

USO
---
    python ofsc_scraper.py --date 2026-09-08 --output consumo_2026-09-08.xlsx
    python ofsc_scraper.py --date 2026-09-08 --headless          # sin ventana visible (requiere --auto-login)
    python ofsc_scraper.py --date 2026-09-08 --auto-login        # login automático con .env / env vars / getpass

    ./reporte.sh 2026-09-08   # atajo: --auto-login --headless + nombre de salida automático

El script escribe el Excel INCREMENTALMENTE (guarda tras cada cuadrilla), así
que si se corta a la mitad, no pierdes lo ya recolectado: vuelve a correr con
el mismo --output y usa --resume para saltar cuadrillas ya completas.
"""

import argparse
import asyncio
import getpass
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

from playwright.async_api import async_playwright, TimeoutError as PWTimeout

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Falta openpyxl. Corre: pip install -r requirements.txt")
    sys.exit(1)


SCRIPT_DIR = Path(__file__).resolve().parent

# Carga OFSC_USERNAME / OFSC_PASSWORD desde un archivo .env junto al script,
# si existe (ver .env.example). No sobreescribe variables ya exportadas en el
# entorno. Si python-dotenv no está instalado, simplemente no hace nada: el
# flujo de getpass/variables de entorno exportadas manualmente sigue
# funcionando igual.
try:
    from dotenv import load_dotenv
    load_dotenv(SCRIPT_DIR / ".env")
except ImportError:
    pass

# ---------------------------------------------------------------------------
# Selectores de LOGIN - AJUSTAR según tu página real de inicio de sesión.
# Se prueban en orden hasta que uno matchee (primer campo visible).
# Inspecciona el formulario con F12 y reemplaza estos valores.
# ---------------------------------------------------------------------------
LOGIN_USER_SELECTORS = [
    '#username',
    'input[name="username"]',
    '#IDToken1',
    'input[name="user"]',
    'input[type="text"]',
    'input[type="email"]',
]
LOGIN_PASS_SELECTORS = [
    '#password',
    'input[name="password"]',
    '#IDToken2',
    'input[type="password"]',
]
LOGIN_SUBMIT_SELECTORS = [
    '#sign-in',
    'button:has-text("Iniciar")',
    'button[type="submit"]',
    'input[type="submit"]',
    '#loginButton',
    'button:has-text("Login")',
    'button:has-text("Sign in")',
]

# Tipos de actividad administrativos que jamás consumen material.
DEFAULT_SKIP_TYPES = {"Almuerzo", "Bodega Inicio Día", "Bodega Fin Día", "BLOQUEO DE FRANJA"}


def load_config(path: Path) -> dict:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# LOGIN
# ---------------------------------------------------------------------------
async def do_login(page, base_url: str, auto_login: bool):
    await page.goto(base_url, wait_until="domcontentloaded")
    # El formulario de login es una app Oracle JET que se renderiza vía JS
    # después de domcontentloaded; sin esta espera, la búsqueda de campos de
    # abajo puede ejecutarse antes de que existan en el DOM.
    try:
        await page.wait_for_load_state("networkidle", timeout=15000)
    except PWTimeout:
        pass

    if not auto_login:
        print("\n" + "=" * 70)
        print(" Inicia sesión MANUALMENTE en la ventana del navegador que se abrió.")
        print(" Cuando ya veas la 'Consola de despacho', vuelve aquí y presiona ENTER.")
        print("=" * 70)
        input(">>> Presiona ENTER para continuar... ")
        return

    username = os.environ.get("OFSC_USERNAME") or input("Usuario OFSC: ").strip()
    password = os.environ.get("OFSC_PASSWORD") or getpass.getpass("Clave OFSC: ")

    await _submit_login_form(page, username, password)
    await page.wait_for_load_state("networkidle", timeout=30000)

    # OFSC puede rechazar el login con un aviso de "se ha superado el número
    # máximo de sesiones" en vez de dejarnos entrar: las corridas headless de
    # este script nunca hacen logout explícito (solo cierran el navegador),
    # así que van dejando sesiones huérfanas del lado del servidor hasta topar
    # el límite de la cuenta. Cuando aparece, hay que tildar la casilla para
    # terminar la sesión más antigua y reenviar el formulario.
    aviso_limite = page.get_by_text("número máximo de sesiones", exact=False).first
    try:
        await aviso_limite.wait_for(state="visible", timeout=5000)
        limite_alcanzado = True
    except Exception:
        limite_alcanzado = False

    if limite_alcanzado:
        print("Aviso de OFSC: se alcanzó el máximo de sesiones concurrentes de esta cuenta.")
        print("Terminando la sesión más antigua y reintentando el login...")
        checkbox = page.get_by_text("Suprimir la sesión y conexión de usuario más antiguas", exact=False).first
        if await checkbox.count():
            await checkbox.click()
        await _submit_login_form(page, username, password)
        await page.wait_for_load_state("networkidle", timeout=30000)


async def _submit_login_form(page, username: str, password: str):
    user_field = await _first_visible(page, LOGIN_USER_SELECTORS)
    pass_field = await _first_visible(page, LOGIN_PASS_SELECTORS)
    if not user_field or not pass_field:
        raise RuntimeError(
            "No se encontró el formulario de login con los selectores configurados. "
            "Ajusta LOGIN_USER_SELECTORS / LOGIN_PASS_SELECTORS en el script, o corre "
            "sin --auto-login para iniciar sesión manualmente."
        )
    await user_field.fill(username)
    await pass_field.fill(password)

    submit = await _first_visible(page, LOGIN_SUBMIT_SELECTORS)
    if submit:
        await submit.click()
    else:
        await pass_field.press("Enter")


async def _first_visible(page, selectors, timeout_ms=4000):
    """Prueba cada selector en orden, esperando activamente (no solo mirando
    el estado instantáneo) a que se vuelva visible, ya que en apps Oracle JET
    el campo puede tardar en renderizarse tras la navegación."""
    for sel in selectors:
        loc = page.locator(sel).first
        try:
            await loc.wait_for(state="visible", timeout=timeout_ms)
            return loc
        except Exception:
            continue
    return None


# ---------------------------------------------------------------------------
# NAVEGACIÓN
# ---------------------------------------------------------------------------
async def wait_no_overlay(page, timeout=15000):
    """Espera a que el overlay de carga global de la app (#plugin-overlay-window,
    z-index 9999) deje de bloquear la pantalla. Aparece brevemente durante
    transiciones de página; si el siguiente clic cae mientras sigue visible,
    Playwright lo reporta como '<div id="plugin-overlay-window"> intercepts
    pointer events' y reintenta hasta agotar su timeout."""
    try:
        await page.locator("#plugin-overlay-window").wait_for(state="hidden", timeout=timeout)
    except Exception:
        pass


async def go_to_console(page, max_steps=4):
    """Vuelve a la vista raíz 'Consola de despacho' desde cualquier sub-página,
    navegando un paso a la vez por el breadcrumb LOCAL de cada pantalla (el
    botón '‹ Consola de despacho' / '‹ Detalles de actividad' que aparece
    junto al título) en vez de saltar directo con el link de navegación
    global del menú lateral.

    Por qué: ese link global (a.global-navigation-item--activities) funciona
    bien desde 'Detalles de actividad', pero desde una subpágina más
    profunda como 'Lista de inventarios' dispara una RECARGA COMPLETA de la
    SPA en vez de una transición fluida — costosa (puede tardar bastante en
    sesiones sin caché tibio) y que resetea el árbol de cuadrillas y la
    vista de lista, dejando fallando TODAS las cuadrillas siguientes. El
    breadcrumb local, en cambio, navega un nivel a la vez sin recargar."""
    nav = page.locator('a.global-navigation-item--activities').first
    for _ in range(max_steps):
        if await nav.count():
            try:
                classes = await nav.get_attribute("class") or ""
                if "current" in classes:
                    break  # ya estamos en la Consola de despacho
            except Exception:
                pass
        breadcrumb = page.locator(
            'button:has-text("Consola de despacho"), button:has-text("Detalles de actividad")'
        ).first
        if await breadcrumb.count() == 0:
            # Sin breadcrumb local visible: único recurso, el link global
            # (puede ser lento si dispara una recarga completa).
            if await nav.count():
                try:
                    await nav.click(timeout=5000)
                    await page.wait_for_load_state("networkidle", timeout=30000)
                except Exception:
                    pass
            break
        try:
            await breadcrumb.click(timeout=5000)
            await page.wait_for_timeout(400)
        except Exception:
            break
    await wait_no_overlay(page)


async def ensure_list_view(page):
    """La Consola de despacho puede arrancar en 'Vista de tiempo' (Gantt) en
    sesiones sin preferencias guardadas (una sesión nueva de Playwright no
    hereda la preferencia de vista de tu navegador normal). El resto del
    script asume la 'Vista de lista' (columnas de texto, celdas
    div.oj-datagrid-cell con el ID de actividad, etc.), así que la forzamos
    explícitamente antes de tocar el árbol de cuadrillas."""
    btn = page.locator('button[aria-label="Vista de lista"]').first
    try:
        # count()==0 aquí sería un chequeo instantáneo: en frío/headless el botón
        # puede tardar en existir aunque la red ya esté quieta (Knockout todavía
        # pintando), así que esperamos activamente en vez de rendirnos de una.
        await btn.wait_for(state="visible", timeout=20000)
    except Exception:
        return  # esta UI no tiene selector de vista (versión distinta); no hay nada que forzar
    try:
        classes = await btn.get_attribute("class") or ""
        if "radio-selected" in classes:
            return  # ya está en vista de lista
        await btn.click(timeout=5000)
        await page.wait_for_timeout(500)
    except Exception:
        pass


async def set_date(page, target_date: str):
    """target_date: 'YYYY-MM-DD'. Usa las flechas Anterior/Siguiente del selector
    de fecha (div.toolbar-datepicker-wrapper) hasta llegar a la fecha, comparando
    contra el aria-label/title del botón de fecha (ej. 'Viernes 11 Septiembre 2026')."""
    target = datetime.strptime(target_date, "%Y-%m-%d").date()
    meses = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto",
             "septiembre","octubre","noviembre","diciembre"]

    wrapper = page.locator("div.toolbar-datepicker-wrapper")
    date_btn = wrapper.locator("button.toolbar-date-picker-button").first
    prev_btn = wrapper.locator('button[aria-label="Anterior"]').first
    next_btn = wrapper.locator('button[aria-label="Siguiente"]').first
    # En frío/headless el toolbar completo puede tardar bastante más de 15s en
    # pintarse (mismo patrón visto en el resto del script: la red se calla
    # antes de que Knockout termine de renderizar).
    await date_btn.wait_for(state="visible", timeout=30000)

    for _ in range(60):  # tope de seguridad: no más de ~2 meses de diferencia
        # El contenedor puede volverse "visible" antes de que Knockout rellene
        # el aria-label/title con la fecha real, así que reintentamos con
        # polling en vez de fallar en el primer intento vacío.
        header_text = await _wait_header_date_text(date_btn, timeout_ms=15000)
        cur = _parse_header_date(header_text, meses)
        if cur is None:
            raise RuntimeError(
                f"No se pudo leer la fecha actual del encabezado (texto: '{header_text}'). "
                "Revisa el selector 'div.toolbar-datepicker-wrapper button.toolbar-date-picker-button' "
                "con F12 si la UI de OFSC cambió."
            )
        if cur == target:
            return
        btn = next_btn if cur < target else prev_btn
        await btn.click()
        await page.wait_for_timeout(700)
    raise RuntimeError(f"No se pudo llegar a la fecha {target_date} tras varios intentos.")


async def _get_header_date_text(date_btn):
    # El aria-label/title del botón trae la fecha completa, ej: "Martes 08 Septiembre 2026".
    try:
        return (await date_btn.get_attribute("aria-label")
                or await date_btn.get_attribute("title")
                or (await date_btn.inner_text()).strip())
    except Exception:
        return ""


async def _wait_header_date_text(date_btn, timeout_ms=15000, poll_ms=300):
    """Como _get_header_date_text, pero reintenta hasta timeout_ms si el botón
    todavía no tiene texto (visible != poblado en apps Knockout/Oracle JET)."""
    elapsed = 0
    text = ""
    while elapsed <= timeout_ms:
        text = await _get_header_date_text(date_btn)
        if text:
            return text
        await date_btn.page.wait_for_timeout(poll_ms)
        elapsed += poll_ms
    return text


def _parse_header_date(text, meses):
    if not text:
        return None
    m = re.search(r"(\d{1,2})\s+([A-Za-zñÑ]+)\s+(\d{4})", text)
    if not m:
        return None
    day, mon_name, year = m.groups()
    mon_name_low = mon_name.lower()
    if mon_name_low not in meses:
        return None
    month = meses.index(mon_name_low) + 1
    try:
        return datetime(int(year), month, int(day)).date()
    except ValueError:
        return None


async def expand_crew_tree(page, max_clicks=200):
    """Expande todos los grupos colapsados del árbol de cuadrillas (izquierda).

    En una sesión nueva (sin historial previo de clics en ese navegador), el
    árbol arranca con sus grupos colapsados (botón button.edt-open con clase
    'ptplus'); los nodos de cuadrilla dentro de un grupo colapsado a veces ni
    siquiera existen en el DOM todavía, por lo que select_crew() los reporta
    como "no encontrados". Expandimos todo una vez, al inicio, antes de
    recorrer las cuadrillas."""
    for _ in range(max_clicks):
        collapsed = page.locator("button.edt-open.ptplus").first
        if await collapsed.count() == 0:
            break
        try:
            await collapsed.click(timeout=3000)
        except Exception:
            break
        await page.wait_for_timeout(200)

    # Cada expansión dispara una carga de datos del grupo (m=Provider&a=opentree)
    # que en una sesión sin caché puede tardar varios segundos; si no dejamos que
    # esas peticiones terminen, compiten con la siguiente selección de cuadrilla
    # y su respuesta m=Grid&a=get puede tardar 15-20s o más en llegar.
    try:
        await page.wait_for_load_state("networkidle", timeout=30000)
    except PWTimeout:
        pass


async def select_crew(page, crew_code: str):
    """Hace clic en el nodo de la cuadrilla en el árbol izquierdo y espera la
    respuesta m=Grid&a=get con la lista de actividades del día."""
    await wait_no_overlay(page)
    # :visible filtra los nodos ocultos/duplicados que deja el árbol virtualizado
    # (ramas colapsadas o ya renderizadas pero fuera de vista).
    loc = page.locator(f'button.edt-label:visible:has-text("{crew_code}")').first
    if await loc.count() == 0:
        # El árbol puede volver a colapsarse al navegar a "Detalles de
        # actividad"/"Inventario" y regresar (no es un estado que sobreviva
        # una sola expansión al inicio de la corrida), dejando la cuadrilla
        # invisible de nuevo. Reexpandimos bajo demanda antes de rendirnos.
        await expand_crew_tree(page)
        loc = page.locator(f'button.edt-label:visible:has-text("{crew_code}")').first
    if await loc.count() == 0:
        loc = page.get_by_text(crew_code, exact=True).first
    if await loc.count() == 0:
        return None

    # En sesiones sin caché tibio, esta respuesta puede tardar bastante más de
    # lo que uno esperaría de un clic simple (medido hasta ~20s en frío).
    async with page.expect_response(lambda r: "m=Grid" in r.url and "a=get" in r.url, timeout=45000) as resp_info:
        await loc.click()
    resp = await resp_info.value
    try:
        data = await resp.json()
    except Exception:
        return None
    return data


async def return_to_crew_list(page, crew_code: str):
    """Vuelve a la Consola de despacho y reselecciona la cuadrilla tras
    procesar una actividad. A diferencia de la selección inicial de
    cuadrilla (protegida en el loop principal), este paso se repite muchas
    veces por corrida (una vez por actividad), así que absorbe cualquier
    excepción en vez de dejarla escapar: un timeout aquí no debe tumbar todo
    el script, solo esta cuadrilla. Devuelve True si quedamos listos para la
    siguiente actividad, False si hay que abandonar esta cuadrilla."""
    try:
        await go_to_console(page)
        await select_crew(page, crew_code)
        return True
    except Exception as e:
        print(f"  ! No se pudo volver a la lista de {crew_code}: {e}")
        return False


def _parse_city_from_zone(aworkzone: str) -> str:
    """aworkzone viene como 'TECNOLOGIA.REGION.PROVINCIA/CIUDAD/PARROQUIA'
    (ej. ' GPON.R2.GUAYAS/GUAYAQUIL/XIMENA'); la ciudad es el segundo
    segmento separado por '/'. Coincide con el campo 'Ciudad' que se ve en
    el detalle de la actividad."""
    if not aworkzone:
        return ""
    parts = aworkzone.strip().split("/")
    return parts[1].strip() if len(parts) >= 2 else ""


def extract_activities_rows(grid_json: dict):
    """Devuelve lista de dicts normalizados desde activitiesRows."""
    rows = grid_json.get("activitiesRows") or []
    out = []
    for r in rows:
        out.append({
            "aid": str(r.get("aid") or r.get("key") or ""),
            "tipo": r.get("aworktype") or "",
            "orden": r.get("appt_number") or "",
            "cliente": (r.get("cname") or "").strip(),
            "estado": r.get("astatus") or "",
            "ciudad": _parse_city_from_zone(r.get("aworkzone") or ""),
        })
    return out


def _is_sync_write(resp):
    return "m=sync" in resp.url and "a=write" in resp.url


def _parse_inventory_item(inv_id, it):
    invtype = it.get("invtype")
    sap_raw = it.get("186")
    return {
        "tipo": "Equipos" if invtype == 1 else "Materiales",
        "descripcion": it.get("200") or "",
        "modelo": it.get("187") or "",
        "serie": it.get("invsn") or "",
        "id_inventario": str(it.get("invid") or inv_id),
        "codigo_sap": str(int(sap_raw)) if sap_raw not in (None, "") else None,
        "cantidad": it.get("quantity"),
    }


class InventoryCache:
    """Acumula, durante TODA la sesión, los ítems de inventario vistos en
    cualquier respuesta ?m=sync&a=write (delta.Inventory), indexados por
    inv_aid.

    Por qué existe: el detalle de una actividad puede no disparar ninguna
    respuesta de sync nueva si sus datos ya llegaron antes (por ejemplo, en
    la sincronización grande que dispara select_crew() al elegir la
    cuadrilla, que puede traer de antemano el inventario de varias
    actividades del día). Si solo escucháramos respuestas "durante" cada
    open_activity_and_get_inventory(), esas actividades reportarían 0 ítems
    aunque sí tengan consumo. Por eso el listener se registra UNA sola vez
    para toda la página, y cada actividad simplemente consulta este caché
    acumulado en el momento en que la necesita."""

    def __init__(self):
        self.by_aid = {}  # aid_str -> {id_inventario: item}
        self._pending = []  # Response objects aún no procesados

    def listener(self, resp):
        if _is_sync_write(resp):
            self._pending.append(resp)

    async def _drain(self):
        pending, self._pending = self._pending, []
        for resp in pending:
            try:
                data = await resp.json()
            except Exception:
                continue
            inv = (data.get("delta") or {}).get("Inventory") or {}
            for inv_id, it in inv.items():
                if it.get("invpool") != "install":
                    continue  # solo lo INSTALADO (no "Recurso"/"Cliente")
                aid_str = str(it.get("inv_aid"))
                item = _parse_inventory_item(inv_id, it)
                self.by_aid.setdefault(aid_str, {})[item["id_inventario"]] = item

    async def items_for(self, aid):
        await self._drain()
        return list(self.by_aid.get(str(aid), {}).values())


async def open_activity_and_get_inventory(page, aid: str, inventory_cache: InventoryCache):
    """Doble clic en la fila de la actividad (por su ID) y devuelve su
    inventario instalado, consultando el InventoryCache global (ver su
    docstring: los datos pueden haber llegado antes, no necesariamente
    durante esta apertura puntual)."""
    cell = page.locator(f'div.oj-datagrid-cell:has-text("{aid}")').first
    if await cell.count() == 0:
        cell = page.get_by_text(aid, exact=True).first
    if await cell.count() == 0:
        raise RuntimeError(f"No se encontró la fila de la actividad {aid} en la grilla.")

    await wait_no_overlay(page)
    await cell.dblclick()
    try:
        await page.wait_for_load_state("networkidle", timeout=20000)
    except PWTimeout:
        pass
    # Margen extra: una respuesta que llegó justo al filo del "idle" puede
    # seguir procesándose (parseo del body) un instante más.
    await page.wait_for_timeout(500)

    items = await inventory_cache.items_for(aid)

    # Si aún no hay datos cacheados para esta actividad, forzamos la carga
    # abriendo explícitamente la pestaña "Inventario". El botón puede no
    # existir todavía en el DOM justo tras el "networkidle" (la plantilla de
    # Knockout puede terminar de pintarse después de que la red se calla),
    # así que esperamos activamente a que aparezca en vez de mirar el estado
    # instantáneo con count().
    if not items:
        inv_btn = page.locator('button[aria-label="Inventario"]').first
        try:
            await inv_btn.wait_for(state="visible", timeout=10000)
            await wait_no_overlay(page)
            await inv_btn.click(timeout=10000)
            await page.wait_for_load_state("networkidle", timeout=20000)
            await page.wait_for_timeout(500)
            items = await inventory_cache.items_for(aid)
        except PWTimeout:
            pass

    return items


# ---------------------------------------------------------------------------
# EXCEL
# ---------------------------------------------------------------------------
DETAIL_COLS = ["CUADRILLA", "FECHA", "ID_ACTIVIDAD", "ORDEN_TRABAJO", "CLIENTE", "CIUDAD",
               "TIPO_ACTIVIDAD", "ESTADO_ACTIVIDAD", "TIPO_INVENTARIO", "DESCRIPCION",
               "MODELO", "SERIE", "ID_INVENTARIO", "CODIGO_SAP", "CANTIDAD"]
SUMMARY_COLS = ["CUADRILLA", "FECHA", "ID_ACTIVIDAD", "ORDEN_TRABAJO", "CLIENTE", "CIUDAD",
                "TIPO_ACTIVIDAD", "ESTADO_ACTIVIDAD", "N_ITEMS_CONSUMIDOS"]


def _style_header(ws, cols):
    header_fill = PatternFill("solid", fgColor="C00000")
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.append(cols)
    for c in range(1, len(cols) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border


def init_workbook(path: Path):
    if path.exists():
        wb = openpyxl.load_workbook(path)
        return wb
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Consumo Detallado"
    _style_header(ws, DETAIL_COLS)
    ws.freeze_panes = "A2"
    ws2 = wb.create_sheet("Resumen Actividades")
    _style_header(ws2, SUMMARY_COLS)
    ws2.freeze_panes = "A2"
    return wb


def already_done_crews(wb) -> set:
    ws2 = wb["Resumen Actividades"]
    return {row[0].value for row in ws2.iter_rows(min_row=2) if row[0].value}


def _sap_cell_value(it):
    """Decide qué mostrar en CODIGO_SAP y si la fila debe advertirse (amarillo).

    Regla de negocio: para Equipos, el número de serie ya identifica la
    unidad de forma suficiente — OFSC casi nunca expone su código SAP en la
    respuesta que interceptamos, y eso es normal, no un error a verificar.
    Solo se advierte un Equipo si falta el serial (con o sin SAP). Para
    Materiales sí se exige el código SAP como antes (rara vez tienen serial
    propio en el que apoyarse)."""
    sap = it["codigo_sap"]
    if it["tipo"] == "Equipos":
        if sap:
            return sap, False
        if it["serie"]:
            return "", False
        return "SIN SERIAL NI CODIGO SAP - VERIFICAR", True
    if sap:
        return sap, False
    return "SIN CODIGO SAP - VERIFICAR", True


def append_activity(wb, crew, fecha, act, items):
    ws = wb["Consumo Detallado"]
    ws2 = wb["Resumen Actividades"]
    warn_fill = PatternFill("solid", fgColor="FFF2CC")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for it in items:
        sap, warn = _sap_cell_value(it)
        row_vals = [crew, fecha, act["aid"], act["orden"], act["cliente"], act["ciudad"], act["tipo"],
                    act["estado"], it["tipo"], it["descripcion"], it["modelo"], it["serie"],
                    it["id_inventario"], sap, it["cantidad"]]
        ws.append(row_vals)
        row_idx = ws.max_row
        for c in range(1, len(DETAIL_COLS) + 1):
            cell = ws.cell(row=row_idx, column=c)
            cell.font = Font(name="Arial", size=10)
            cell.border = border
        if warn:
            for c in range(1, len(DETAIL_COLS) + 1):
                ws.cell(row=row_idx, column=c).fill = warn_fill

    ws2.append([crew, fecha, act["aid"], act["orden"], act["cliente"], act["ciudad"], act["tipo"],
                act["estado"], len(items)])
    row_idx = ws2.max_row
    for c in range(1, len(SUMMARY_COLS) + 1):
        ws2.cell(row=row_idx, column=c).font = Font(name="Arial", size=10)
        ws2.cell(row=row_idx, column=c).border = border


def autosize_and_save(wb, path: Path):
    for ws, widths in [(wb["Consumo Detallado"], [16, 12, 13, 16, 28, 16, 20, 14, 14, 38, 16, 20, 13, 34, 10]),
                        (wb["Resumen Actividades"], [16, 12, 13, 16, 28, 16, 20, 14, 18])]:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    wb.save(path)


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
async def run(args):
    cfg = load_config(SCRIPT_DIR / "config.json")
    skip_types = set(cfg.get("skip_activity_types", DEFAULT_SKIP_TYPES))
    consuming_states = set(cfg.get("estados_con_posible_consumo", ["finalizada"]))
    output_path = SCRIPT_DIR / (args.output or cfg.get("output_file", "consumo_materiales.xlsx"))

    wb = init_workbook(output_path)
    done_crews = already_done_crews(wb) if args.resume else set()
    if done_crews:
        print(f"Reanudando: ya hay datos para {len(done_crews)} cuadrilla(s), se omiten.")

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=args.headless,
            slow_mo=cfg.get("slow_mo_ms", 0),
            args=["--start-maximized"] if not args.headless else None,
        )
        context = await browser.new_context(
            no_viewport=not args.headless,
            viewport=None if not args.headless else {"width": 1600, "height": 900},
        )
        page = await context.new_page()
        page.set_default_timeout(cfg.get("nav_timeout_ms", 30000))

        # Registrado una sola vez para toda la sesión: ver InventoryCache.
        inventory_cache = InventoryCache()
        page.on("response", inventory_cache.listener)

        await do_login(page, cfg["base_url"], args.auto_login)

        # Preparar la Consola (vista de lista + fecha) puede fallar por lentitud
        # o inestabilidad puntual del servidor de OFSC (no depende de qué fecha
        # se pida: esto corre ANTES de navegar a ninguna fecha en particular,
        # la consola todavía muestra "hoy"). Se reintenta una vez recargando
        # antes de darse por vencido, en vez de solo subir timeouts a ciegas.
        ultimo_error = None
        for intento in range(2):
            try:
                await go_to_console(page)
                await page.wait_for_timeout(1000)
                try:
                    await page.wait_for_load_state("networkidle", timeout=20000)
                except PWTimeout:
                    pass
                await ensure_list_view(page)
                await set_date(page, args.date)
                ultimo_error = None
                break
            except Exception as e:
                ultimo_error = e
                if intento == 0:
                    print(f"  ! Falló preparar la consola para la fecha {args.date} (intento 1/2): {e}")
                    print("    Recargando la página y reintentando...")
                    try:
                        await page.reload(wait_until="domcontentloaded")
                        await page.wait_for_timeout(2000)
                    except Exception:
                        pass

        if ultimo_error is not None:
            print(f"  ! Error preparando la consola para la fecha {args.date} tras 2 intentos: {ultimo_error}")
            try:
                print(f"    URL en el momento del error: {page.url}")
                print(f"    Título de la página: {await page.title()}")
            except Exception:
                pass
            if cfg.get("screenshot_on_error"):
                try:
                    shot = SCRIPT_DIR / f"error_inicio_{args.date}.png"
                    await page.screenshot(path=str(shot))
                    print(f"    Captura guardada en {shot}")
                except Exception:
                    pass
            raise ultimo_error
        await expand_crew_tree(page)

        for crew in cfg["crews"]:
            if crew in done_crews:
                print(f"[SKIP] {crew} ya procesada (--resume).")
                continue

            print(f"\n=== Cuadrilla {crew} ===")
            # Por si la cuadrilla/actividad anterior nos dejó fuera de la
            # Consola (p.ej. tras un error de la iteración previa) — no-op
            # si ya estamos ahí.
            await go_to_console(page)
            try:
                grid_json = await select_crew(page, crew)
            except Exception as e:
                print(f"  ! No se pudo abrir la cuadrilla {crew}: {e}")
                continue
            if grid_json is None:
                print(f"  ! No se encontró la cuadrilla '{crew}' en el árbol (revisa el nombre en config.json).")
                continue

            activities = extract_activities_rows(grid_json)
            print(f"  {len(activities)} actividad(es) el {args.date}.")

            for act in activities:
                if act["tipo"] in skip_types:
                    print(f"  - {act['aid']} ({act['tipo']}): administrativa, se omite.")
                    continue
                if act["estado"] not in consuming_states:
                    print(f"  - {act['aid']} ({act['tipo']}, {act['estado']}): estado sin consumo esperado, se omite del Excel.")
                    continue

                try:
                    items = await open_activity_and_get_inventory(page, act["aid"], inventory_cache)
                except Exception as e:
                    print(f"  ! Error abriendo actividad {act['aid']}: {e}")
                    if cfg.get("screenshot_on_error"):
                        shot = SCRIPT_DIR / f"error_{crew}_{act['aid']}.png"
                        await page.screenshot(path=str(shot))
                        print(f"    (captura guardada en {shot})")
                    if not await return_to_crew_list(page, crew):
                        break  # no pudimos recuperar el estado: pasamos a la siguiente cuadrilla
                    continue

                print(f"  - {act['aid']} ({act['tipo']}, {act['estado']}): {len(items)} item(s).")
                append_activity(wb, crew, args.date, act, items)

                # volver a la lista de la cuadrilla para la siguiente actividad
                if not await return_to_crew_list(page, crew):
                    break  # no pudimos recuperar el estado: pasamos a la siguiente cuadrilla

            autosize_and_save(wb, output_path)
            print(f"  Guardado parcial en {output_path}")

        await browser.close()

    autosize_and_save(wb, output_path)
    print(f"\nListo. Archivo final: {output_path}")


def parse_args():
    ap = argparse.ArgumentParser(description="Scraper de consumo de materiales OFSC (Claro Ecuador).")
    ap.add_argument("--date", required=True, help="Fecha a extraer, formato YYYY-MM-DD (ej. 2026-09-08).")
    ap.add_argument("--output", default=None, help="Nombre del archivo Excel de salida.")
    ap.add_argument("--headless", action="store_true", help="Correr sin ventana visible (requiere --auto-login).")
    ap.add_argument("--auto-login", action="store_true", help="Login automático (usuario/clave simples, sin SSO).")
    ap.add_argument("--resume", action="store_true", help="Si el Excel de salida ya existe, omite cuadrillas ya completas.")
    return ap.parse_args()


if __name__ == "__main__":
    args = parse_args()
    if args.headless and not args.auto_login:
        print("--headless requiere --auto-login (no hay forma de iniciar sesión manualmente sin ventana).")
        sys.exit(1)
    asyncio.run(run(args))
