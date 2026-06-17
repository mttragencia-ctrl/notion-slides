"""Lectura de una base de datos de Notion y mapeo a items para el deck.

Expone:
    fetch_items(token, database_id) -> (deck_title, items)

`items` es la lista de dicts que consume slides.build_deck.
El mapeo es por NOMBRE de propiedad y tolerante a faltantes, para funcionar con
las distintas bases mensuales que usen el mismo template.
"""

from datetime import date

from notion_client import Client

# Iconos según el formato detectado.
ICON_REEL    = "\U0001f4fd"   # 📽
ICON_CARRUSEL = "\U0001f5bc"  # 🖼
ICON_HISTORIA = "⏰"       # ⏰
ICON_DEFAULT = "\U0001f4c4"   # 📄

# Nombres de propiedades esperados (con alternativas tolerantes).
PROP_FORMATO   = ["Formato"]
PROP_PILAR     = ["Pilar de contenido", "Pilar"]
PROP_TIPO      = ["Tipo de historia", "Tipo"]
PROP_OBJETIVO  = ["Objetivo"]
PROP_ESTADO    = ["Estado"]
PROP_FECHA     = ["Fecha"]


# ─── Helpers de extracción de propiedades ─────────────────────────────────────
def _find_prop(props, names):
    for n in names:
        if n in props:
            return props[n]
    return None


def _get_title(props):
    for p in props.values():
        if p.get("type") == "title":
            parts = p.get("title", [])
            return "".join(t.get("plain_text", "") for t in parts).strip()
    return ""


def _get_date(props):
    p = _find_prop(props, PROP_FECHA)
    if not p:
        for cand in props.values():
            if cand.get("type") == "date":
                p = cand
                break
    if p and p.get("type") == "date" and p.get("date"):
        return p["date"].get("start")
    return None


def _get_multi(props, names):
    p = _find_prop(props, names)
    if p and p.get("type") == "multi_select":
        return [o.get("name", "") for o in p.get("multi_select", [])]
    if p and p.get("type") == "select" and p.get("select"):
        return [p["select"].get("name", "")]
    return []


def _get_status(props, names):
    p = _find_prop(props, names)
    if p and p.get("type") == "status" and p.get("status"):
        return p["status"].get("name", "")
    if p and p.get("type") == "select" and p.get("select"):
        return p["select"].get("name", "")
    return ""


# ─── Derivaciones de presentación ────────────────────────────────────────────
def _derive_red(formatos):
    """Deriva la red social a partir del Formato (misma lógica del armado manual)."""
    f = " ".join(formatos).lower()
    if "historia" in f:
        return "Instagram Stories"
    if "carrusel" in f or "posteo" in f:
        return "Instagram Feed"
    if "reel" in f:
        return "Instagram"
    return "Instagram"


def _derive_icon(formatos):
    f = " ".join(formatos).lower()
    if "reel" in f:
        return ICON_REEL
    if "carrusel" in f or "posteo" in f:
        return ICON_CARRUSEL
    if "historia" in f:
        return ICON_HISTORIA
    return ICON_DEFAULT


_MESES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre",
    12: "Diciembre",
}


def _format_date(iso):
    """'2026-07-13' -> '13 Jul'."""
    if not iso:
        return ""
    try:
        d = date.fromisoformat(iso[:10])
        return f"{d.day:02d} {_MESES[d.month][:3]}"
    except (ValueError, KeyError):
        return iso[:10]


# ─── Lectura del contenido de la página (guión / copy) ────────────────────────
def _rich_text_to_str(rich):
    return "".join(t.get("plain_text", "") for t in rich).strip()


def _block_text(block):
    """Devuelve (tipo_logico, texto) de un bloque, o (None, '')."""
    btype = block.get("type")
    data = block.get(btype, {})
    rich = data.get("rich_text", [])
    text = _rich_text_to_str(rich)
    return btype, text


def _read_page_content(client, page_id):
    """Lee los bloques de la página y separa en (guion[list], copy[str|None]).

    Agrupa por encabezados de nivel 1 ('# Guión...', '# Copy'). Todo lo que cae
    bajo una sección cuyo título contiene 'copy' va al copy; el resto al guión.
    """
    guion = []
    copy_parts = []
    section = "guion"  # sección activa por defecto

    cursor = None
    while True:
        resp = client.blocks.children.list(block_id=page_id, start_cursor=cursor)
        for block in resp.get("results", []):
            btype, text = _block_text(block)
            if not text:
                continue
            low = text.lower()

            # Encabezado de nivel 1 => cambia de sección
            if btype == "heading_1":
                section = "copy" if "copy" in low else "guion"
                continue
            # Encabezados internos tipo "Guión del carrusel" / "Inspo": omitir
            if btype == "heading_2" and ("gui" in low or "inspo" in low):
                continue

            if section == "copy":
                copy_parts.append(text)
            else:
                # Saltar marcadores de encabezado de sección sueltos
                if btype in ("heading_2",) and ("gui" in low):
                    continue
                guion.append(text)

        if not resp.get("has_more"):
            break
        cursor = resp.get("next_cursor")

    copy = " ".join(copy_parts).strip() or None
    return guion, copy


# ─── Función principal ────────────────────────────────────────────────────────
def fetch_items(token, database_id):
    """Lee la base y devuelve (deck_title, items ordenados por fecha)."""
    client = Client(auth=token)

    # Título de la base (para portada y nombre de archivo).
    deck_title = "Planificacion"
    try:
        db = client.databases.retrieve(database_id=database_id)
        title_parts = db.get("title", [])
        raw_title = "".join(t.get("plain_text", "") for t in title_parts).strip()
        if raw_title:
            deck_title = raw_title
    except Exception:
        pass

    # Query paginado de todas las filas.
    rows = []
    cursor = None
    while True:
        resp = client.databases.query(database_id=database_id, start_cursor=cursor)
        rows.extend(resp.get("results", []))
        if not resp.get("has_more"):
            break
        cursor = resp.get("next_cursor")

    items = []
    for page in rows:
        props = page.get("properties", {})
        name = _get_title(props)
        iso_date = _get_date(props)
        formatos = _get_multi(props, PROP_FORMATO)
        pilares = _get_multi(props, PROP_PILAR)
        tipos = _get_multi(props, PROP_TIPO)
        objetivos = _get_multi(props, PROP_OBJETIVO)
        estado = _get_status(props, PROP_ESTADO)

        guion, copy = _read_page_content(client, page["id"])
        is_draft = (not guion) and (not copy)

        items.append({
            "_iso": iso_date or "9999-12-31",
            "date": _format_date(iso_date),
            "name": name or "(sin nombre)",
            "icon": _derive_icon(formatos),
            "formato": " / ".join(formatos) if formatos else None,
            "red": _derive_red(formatos) if formatos else None,
            "pilar": " / ".join(pilares) if pilares else None,
            "tipo": " / ".join(tipos) if tipos else None,
            "objetivo": " / ".join(objetivos) if objetivos else None,
            "estado": estado or ("Borrador" if is_draft else "—"),
            "borrador": is_draft,
            "guion": guion or None,
            "copy": copy,
        })

    items.sort(key=lambda it: it["_iso"])
    return deck_title, items
