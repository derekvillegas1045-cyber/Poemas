#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera poemas.xlsx — la hoja de cálculo que sirve de fuente de verdad
para la colección de poemas.

Uso:
    python tools/generate_poemas.py
    python tools/generate_poemas.py --out ruta/a/poemas.xlsx

Solo usa la librería estándar de Python (zipfile + xml), no necesita
instalar nada.

Estructura del xlsx generado:
    Hoja "Poemas":  id, title, dedicatory, author, theme, badge, columns
    Hoja "Versos":  poem_id, line_no, text

Edita el .xlsx en Excel, Google Sheets o LibreOffice y vuelve a abrir
la página; los cambios aparecen solos.
"""

import argparse
import re
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

# ----------------------------------------------------------------------------
# Colección de poemas (fuente inicial; edítala aquí o directamente en el xlsx)
# ----------------------------------------------------------------------------
POEMS = [
    {
        "title": "Tu Existencia",
        "dedicatory": "Mar Mar Mar, siempre eres tu Mar.",
        "author": "Anónimo",
        "theme": "rosa",
        "badge": "",
        "columns": "no",
        "lines": [
            "Tu existencia es un regalo",
            "Tus palabras una oda a la belleza",
            "Quiero abrazarte hasta que muera",
            "Y eliminar todo lo que es malo",
            "",
            "Estaré para ti",
            "quizá no con las manos cocidas",
            "Pero sí con el corazón fusionado",
            "Llegamos a ser uno mismo",
            "siempre y todos los días",
            "",
            "Vivimos en la incertidumbre",
            "del ¿Cuánto me quiere?",
            "pero contigo no existe",
            "porque me lo demuestras diariamente.",
            "",
            "Permíteme seguir estando",
            "Permíteme seguir hablando",
            "pero también seguir escuchando.",
        ],
    },
    {
        "title": "Regalo No. 23",
        "dedicatory": "¿Dónde se supone que estás?",
        "author": "Anónimo",
        "theme": "rosa",
        "badge": "23",
        "columns": "sí",
        "lines": [
            "Siento que no estoy para ti, que mereces más",
            "Creo que tu amor va muy lejos, aquí y allá",
            "Tus abrazos cálidos y que dan mucha paz",
            "Reconfortantes como el rojo de tu cabellera",
            "Quemándome como un millón de soles",
            "Tratando de mantenerse un poco estables",
            "Colapsando para dar paso a una vida nueva",
            "Una aventura sin nubes grises que lluevan",
            "Porque no importa cuantas lagrimas caigan",
            "Son gotas mojadas que a ti nunca te apagan",
            "Tremenda dualidad que tienes en tu ser",
            "Llevando el Mar en tu nombre primero",
            "Pero en el corazón un fuego eterno",
            "Que luchan por un espacio y permanecer",
            "Quedándose como yo en tus días",
            "Incrustado en tus buenas memorias",
            "En tus cumpleaños completos siempre",
            "Y en los medios sin falta estar",
            "Siendo un solo ser buscando ser libre",
            "Y espero que eso no llegue a faltar",
            "Te debo mi vida, al ser mi gran ancla",
            "Te ofrezco mis palabras, mi pobre alma",
            "Proporcionas algo, como tranquilidad",
            "Porque tu siempre me dices la verdad",
            "Sin importar el cómo pueda yo actuar",
            "Cuando genuinamente volveré a estar",
            "No pienso que me veas ausentar",
            "Pero si llegara yo no más nunca cantar",
            "Que no me puedas otra vez escuchar",
            "Sepas que mi corazón en pausa teatral",
            "Seguirá recordando tu suave y fino tocar",
            "Esperando volver a sentir el vibrar",
            "De tus dedos pasando sin algo afinar",
            "Extrañando el sonido de tu risa",
            "Pero deseando una reunión a prisa",
            "Aun sin desear una reunión fatídica",
            "Algo que se sienta como crítica",
            "Si me voy no espero que me acompañes",
            "Porque tu vives bien, aunque quizá llores",
            "No quiero ser el ancla de tus olas",
            "Más bien las nubes que acompañan",
            "Las sales disueltas las cuales te bañan",
            "No retenerte, para que nunca estés sola",
            "Fumar tus ideas con mis labios partidos",
            "Disfrutar la pureza de tus palabras",
            "Deleitando la inmundicia de mis oídos",
            "Entrando las ondas en las grietas",
            "Propagando la belleza, dejando marca",
            "Como una mordida roja, que sangra",
            "Permanentemente tuya",
            "Blanca como muñequita de porcelana.",
        ],
    },
]

VALID_THEMES = ["rosa", "medianoche", "atardecer", "esmeralda", "sepia"]
VALID_BOOL = ["sí", "si", "no", "yes", "true", "false", "1", "0", ""]


def slugify(text):
    slug = re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")
    return slug or "poema"


def cell(value):
    """Convierte a texto plano para la celda."""
    return "" if value is None else str(value)


def write_xlsx(path, poems):
    """Escribe un .xlsx mínimo y válido usando solo la librería estándar."""
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ET.register_namespace("", ns)

    def col_letter(idx):
        # 1 -> A, 2 -> B, ...
        letters = ""
        while idx:
            idx, rem = divmod(idx - 1, 26)
            letters = chr(65 + rem) + letters
        return letters

    def sheet_xml(rows):
        root = ET.Element(f"{{{ns}}}worksheet")
        sheet_data = ET.SubElement(root, f"{{{ns}}}sheetData")
        for r, row in enumerate(rows, start=1):
            row_el = ET.SubElement(sheet_data, f"{{{ns}}}row", {"r": str(r)})
            for c, value in enumerate(row, start=1):
                cell_el = ET.SubElement(
                    row_el,
                    f"{{{ns}}}c",
                    {"r": f"{col_letter(c)}{r}", "t": "inlineStr"},
                )
                is_el = ET.SubElement(cell_el, f"{{{ns}}}is")
                t_el = ET.SubElement(is_el, f"{{{ns}}}t")
                t_el.text = cell(value)
        return ET.tostring(root, encoding="UTF-8", xml_declaration=True)

    poem_rows = [["id", "title", "dedicatory", "author", "theme", "badge", "columns"]]
    verse_rows = [["poem_id", "line_no", "text"]]
    for poem in poems:
        pid = slugify(poem["title"])
        poem_rows.append([
            pid,
            poem["title"],
            poem["dedicatory"],
            poem["author"],
            poem["theme"],
            poem["badge"],
            poem["columns"],
        ])
        for n, line in enumerate(poem["lines"], start=1):
            verse_rows.append([pid, str(n), line])

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '</Types>'
    ).encode("utf-8")

    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    ).encode("utf-8")

    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets>'
        '<sheet name="Poemas" sheetId="1" r:id="rId1"/>'
        '<sheet name="Versos" sheetId="2" r:id="rId2"/>'
        '</sheets>'
        '</workbook>'
    ).encode("utf-8")

    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
        '</Relationships>'
    ).encode("utf-8")

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = tmp.name
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
                z.writestr("[Content_Types].xml", content_types)
                z.writestr("_rels/.rels", root_rels)
                z.writestr("xl/workbook.xml", workbook)
                z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
                z.writestr("xl/worksheets/sheet1.xml", sheet_xml(poem_rows))
                z.writestr("xl/worksheets/sheet2.xml", sheet_xml(verse_rows))
        Path(tmp_path).replace(path)
    finally:
        if tmp_path and Path(tmp_path).exists():
            Path(tmp_path).unlink()


def validate(poems):
    """Avisos de consistencia antes de escribir el archivo."""
    seen = set()
    for poem in poems:
        if poem["theme"] not in VALID_THEMES:
            raise SystemExit(f"Tema inválido en «{poem['title']}»: {poem['theme']} (usa: {', '.join(VALID_THEMES)})")
        if poem["columns"] not in VALID_BOOL:
            raise SystemExit(f"Valor de columns inválido en «{poem['title']}»: {poem['columns']} (usa sí/no)")
        pid = slugify(poem["title"])
        if pid in seen:
            raise SystemExit(f"Título duplicado (mismo id): {pid}")
        seen.add(pid)
        if not [line for line in poem["lines"] if line.strip()]:
            raise SystemExit(f"El poema «{poem['title']}» no tiene versos")


def main():
    parser = argparse.ArgumentParser(description="Genera poemas.xlsx")
    parser.add_argument("--out", default="poemas.xlsx", help="Ruta de salida (por defecto: poemas.xlsx)")
    args = parser.parse_args()

    validate(POEMS)
    out = Path(args.out)
    backup = out.with_suffix(out.suffix + ".bak")
    if out.exists():
        shutil.copy2(out, backup)
        print(f"Copia de seguridad: {backup.name}")

    write_xlsx(out, POEMS)
    total_lines = sum(len(p["lines"]) for p in POEMS)
    print(f"OK -> {out} ({len(POEMS)} poemas, {total_lines} versos)")


if __name__ == "__main__":
    main()
