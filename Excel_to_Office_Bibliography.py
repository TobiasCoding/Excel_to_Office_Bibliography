#!/usr/bin/env python3
# excel2word_bib.py
# Convierte un Excel a Sources.xml para Word (OOXML Bibliography).
# Requisitos: openpyxl (pip install openpyxl)
# Uso: python excel2word_bib.py input.xlsx output.xml
# Notas:
# - Lee la hoja activa del Excel.
# - Columnas esperadas (camelCase): ver listado en el encabezado del archivo.
# - Para autores persona, usar "Last, First Middle" separados por ";"
#   Ej: "Dwork, Cynthia; McSherry, Frank; Nissim, Kobbi; Smith, Adam"
# - No realiza sanitización avanzada ni valores por defecto.

import sys
from typing import Dict, Any, List, Optional
from openpyxl import load_workbook
import xml.etree.ElementTree as ET

# Namespace OOXML de bibliografía (Word)
NS = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
ET.register_namespace("b", NS)           # prefijo b:
ET.register_namespace("", NS)            # default namespace igual al de bibliografía

# Columnas esperadas en camelCase
EXPECTED_COLUMNS = [
    "tag", "sourceType", "authorType", "authors", "title",
    "journalName", "bookTitle", "internetSiteTitle",
    "year", "month", "day",
    "volume", "issue", "pages", "edition",
    "publisher", "city", "url"
]

REQUIRED_COLUMNS = ["tag", "sourceType", "authorType", "title", "year"]

def get(cell_val: Any) -> Optional[str]:
    """Convierte a str y recorta espacios; None si vacío."""
    if cell_val is None:
        return None
    s = str(cell_val).strip()
    return s if s else None

def add_text(parent: ET.Element, tag: str, text: Optional[str]) -> Optional[ET.Element]:
    """Crea un hijo <b:tag> con texto si text no es None."""
    if text is None:
        return None
    el = ET.SubElement(parent, f"{{{NS}}}{tag}")
    el.text = text
    return el

def parse_person_authors(authors_str: str) -> List[Dict[str, str]]:
    """
    Parsea "Last, First Middle; Last2, First2" -> lista de dicts con 'Last' y 'First'.
    - 'First' contendrá 'First Middle' completo (Word acepta <First> con nombres compuestos).
    """
    persons = []
    for part in [p.strip() for p in authors_str.split(";") if p.strip()]:
        if "," in part:
            last, first = part.split(",", 1)
            persons.append({"Last": last.strip(), "First": first.strip()})
        else:
            # Si no hay coma, lo tratamos como 'Last' solamente.
            persons.append({"Last": part.strip(), "First": ""})
    return persons

def build_author_block(parent: ET.Element, row: Dict[str, Optional[str]]) -> None:
    """
    Construye:
      <b:Author>
        <b:Author>
          <b:NameList>.../o <b:Corporate>...</b:Corporate>
        </b:Author>
      </b:Author>
    según authorType = person | corporate
    """
    outer_author = ET.SubElement(parent, f"{{{NS}}}Author")
    inner_author = ET.SubElement(outer_author, f"{{{NS}}}Author")

    a_type = (row.get("authorType") or "").lower()
    authors = row.get("authors")

    if a_type == "person" and authors:
        name_list = ET.SubElement(inner_author, f"{{{NS}}}NameList")
        for person in parse_person_authors(authors):
            p = ET.SubElement(name_list, f"{{{NS}}}Person")
            add_text(p, "Last", person.get("Last"))
            # Word requiere <First>; <Middle> es opcional. Pondremos todo en First.
            add_text(p, "First", person.get("First"))
    elif a_type == "corporate" and authors:
        add_text(inner_author, "Corporate", authors)
    else:
        # Si no hay autores o tipo inválido, generamos un autor corporativo genérico
        add_text(inner_author, "Corporate", "Unknown")

def build_source(parent: ET.Element, row: Dict[str, Optional[str]]) -> None:
    """Crea un bloque <b:Source> con los campos disponibles."""
    src = ET.SubElement(parent, f"{{{NS}}}Source")
    add_text(src, "Tag", row.get("tag"))
    add_text(src, "SourceType", row.get("sourceType"))

    # Autor (estructura que Word interpreta correctamente)
    build_author_block(src, row)

    # Títulos y datos bibliográficos comunes
    add_text(src, "Title", row.get("title"))
    add_text(src, "JournalName", row.get("journalName"))
    add_text(src, "BookTitle", row.get("bookTitle"))
    add_text(src, "InternetSiteTitle", row.get("internetSiteTitle"))

    add_text(src, "Year", row.get("year"))
    add_text(src, "Month", row.get("month"))
    add_text(src, "Day", row.get("day"))

    add_text(src, "Volume", row.get("volume"))
    add_text(src, "Issue", row.get("issue"))
    add_text(src, "Pages", row.get("pages"))
    add_text(src, "Edition", row.get("edition"))

    add_text(src, "Publisher", row.get("publisher"))
    add_text(src, "City", row.get("city"))
    add_text(src, "URL", row.get("url"))

def read_excel_rows(xlsx_path: str) -> List[Dict[str, Optional[str]]]:
    """Lee la hoja activa y devuelve filas como dict con claves de EXPECTED_COLUMNS."""
    wb = load_workbook(filename=xlsx_path, data_only=True)
    ws = wb.active

    # Mapeo encabezados -> índice de columna
    header_row = None
    for r in ws.iter_rows(min_row=1, max_row=1, values_only=True):
        header_row = [get(h) for h in r]
        break
    if not header_row:
        raise ValueError("No se encontró fila de encabezados en el Excel (fila 1).")

    col_index: Dict[str, int] = {}
    for idx, name in enumerate(header_row):
        if name in EXPECTED_COLUMNS:
            col_index[name] = idx

    # Verificación de obligatorias
    missing = [c for c in REQUIRED_COLUMNS if c not in col_index]
    if missing:
        raise ValueError(f"Faltan encabezados obligatorios en el Excel: {', '.join(missing)}")

    rows: List[Dict[str, Optional[str]]] = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        # Saltear filas completamente vacías
        if all(v is None or str(v).strip() == "" for v in r):
            continue
        row: Dict[str, Optional[str]] = {}
        for col in EXPECTED_COLUMNS:
            if col in col_index and col_index[col] < len(r):
                row[col] = get(r[col_index[col]])
            else:
                row[col] = None
        rows.append(row)
    return rows

def build_sources_xml(rows: List[Dict[str, Optional[str]]]) -> ET.ElementTree:
    """Arma el árbol XML <b:Sources> con todos los <b:Source>."""
    root = ET.Element(f"{{{NS}}}Sources", attrib={"SelectedStyle": ""})
    for row in rows:
        # Validación mínima por fila
        for req in REQUIRED_COLUMNS:
            if not row.get(req):
                # Si falta obligatorio, salteamos esta fila
                continue
        build_source(root, row)
    return ET.ElementTree(root)

def main() -> None:
    if len(sys.argv) != 3:
        print("Uso: python excel2word_bib.py input.xlsx output.xml")
        sys.exit(1)

    xlsx_path = sys.argv[1]
    xml_path = sys.argv[2]

    rows = read_excel_rows(xlsx_path)
    tree = build_sources_xml(rows)

    # Guardado con declaración XML y UTF-8
    tree.write(xml_path, encoding="utf-8", xml_declaration=True)
    print(f"OK: escrito '{xml_path}' con {len(rows)} fila(s) de entrada.")

if __name__ == "__main__":
    main()
