"""Extraer el database_id (o page_id) de una URL de Notion.

Notion usa IDs de 32 caracteres hexadecimales. En una URL pueden aparecer:
  - https://www.notion.so/workspace/Nombre-De-La-Base-36d5784b9abc8093bb94fd7a36c3bde7
  - https://app.notion.com/p/36d5784b9abc8093bb94fd7a36c3bde7?v=2d75...&source=copy_link
  - https://notion.so/36d5784b-9abc-8093-bb94-fd7a36c3bde7

El parámetro ?v=... es el ID de una VISTA, no de la base: hay que ignorarlo y
quedarse con el ID que está en el path.
"""

import re

_HEX32 = re.compile(r"[0-9a-fA-F]{32}")


def _strip_query(url: str) -> str:
    """Devuelve solo la parte del path, sin query string (?v=...&source=...)."""
    return url.split("?", 1)[0]


def extract_id(url_or_id: str) -> str:
    """Devuelve un ID de Notion con guiones (formato UUID).

    Acepta tanto una URL completa como un ID pegado directamente.
    Lanza ValueError si no encuentra un ID válido.
    """
    if not url_or_id or not url_or_id.strip():
        raise ValueError("No ingresaste ninguna URL de Notion.")

    text = url_or_id.strip()

    # Tomar solo el path para evitar capturar el ID de la vista (?v=...).
    path = _strip_query(text)

    # Buscar el ÚLTIMO id de 32 hex en el path: en URLs tipo /Nombre-<id>
    # el id real es el último segmento.
    matches = _HEX32.findall(path)
    if not matches:
        # Tal vez el usuario pegó un ID ya con guiones; quitar guiones y reintentar.
        compact = path.replace("-", "")
        matches = _HEX32.findall(compact)
    if not matches:
        raise ValueError(
            "No se encontró un ID de Notion en el link. "
            "Asegurate de copiar el link de la base de datos completo."
        )

    raw = matches[-1].lower()
    # Formatear como UUID 8-4-4-4-12
    return f"{raw[0:8]}-{raw[8:12]}-{raw[12:16]}-{raw[16:20]}-{raw[20:32]}"


if __name__ == "__main__":
    # Pruebas rápidas
    casos = [
        "https://app.notion.com/p/36d5784b9abc8093bb94fd7a36c3bde7?v=2d75784b9abc8299ab6f08c14683e949&source=copy_link",
        "https://www.notion.so/agencia/Planificacion-Julio-36d5784b9abc8093bb94fd7a36c3bde7",
        "36d5784b-9abc-8093-bb94-fd7a36c3bde7",
        "36d5784b9abc8093bb94fd7a36c3bde7",
    ]
    for c in casos:
        print(f"{c[:50]:50s} -> {extract_id(c)}")
