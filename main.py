import json
import re
import pandas as pd
from groq import Groq
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

client = Groq(api_key="<REDACTED>")

instrucciones_excel = """
Respond ONLY with valid JSON.

Structure:
{
  "datos": [[...], [...]],
  "columnas": ["col1", "col2"],
  "estilo": {
    "font_size": <number, default 11>,
    "header_color": <hex sin #, default "2F4F7F">,
    "font_color_header": <hex sin #, default "FFFFFF">,
    "row_alt_color": <hex sin #, default "F2F2F2">
  }
}

Rules:
- No explanations, no markdown, no extra text
- columnas y datos deben coincidir en cantidad
- Las descripciones deben ser detalladas, específicas y distintas entre sí
- Los datos deben ser realistas y coherentes con el contexto
- Nunca repitas el mismo valor en una columna de descripción
- El estilo debe ser coherente con el tema del Excel pedido
"""

instrucciones_word = """
Respond ONLY with valid JSON.

Structure:
{
  "titulo": "Título del documento",
  "terminos": [
    {
      "nombre": "Nombre del concepto",
      "definicion": "Definición detallada del concepto.",
      "palabras_clave": ["palabra1", "palabra2"]
    }
  ]
}

Rules:
- No explanations, no markdown, no extra text
- Las definiciones deben ser claras, completas y académicas
- palabras_clave son los términos más importantes dentro de la definición (2-4 por concepto)
- Nunca repitas definiciones
"""


def iu_basica():
    seleccion = input("Introduzca 1 para generar un Excel y 2 para un Word: ")
    if seleccion == "1":
        prompt = input("Describa el Excel: ")
    else:
        prompt = input("Describa el Word: ")
    return prompt, seleccion


def generacion_json(prompt, instrucciones):
    completion = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[
            {"role": "system", "content": instrucciones},
            {"role": "user", "content": prompt}
        ],
        temperature=0
    )
    return completion.choices[0].message.content


def parsear_json(texto):
    try:
        return json.loads(texto)
    except json.JSONDecodeError:
        pass
    match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", texto, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(1))
        except json.JSONDecodeError:
            pass
    match = re.search(r"\{.*\}", texto, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(0))
        except json.JSONDecodeError:
            pass
    return None


def formatear_excel(path, estilo):
    wb = load_workbook(path)
    ws = wb.active

    header_fill = PatternFill("solid", fgColor=estilo.get("header_color", "2F4F7F"))
    header_font = Font(bold=True, color=estilo.get("font_color_header", "FFFFFF"), size=estilo.get("font_size", 11))

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    fill_alt = PatternFill("solid", fgColor=estilo.get("row_alt_color", "F2F2F2"))
    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        for cell in row:
            if i % 2 == 0:
                cell.fill = fill_alt
            cell.font = Font(size=estilo.get("font_size", 11))
            cell.alignment = Alignment(horizontal="left", wrap_text=True)

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border

    wb.save(path)


def generar_word(data, output_path):
    titulo = data.get("titulo", "Documento")
    terminos = data.get("terminos", [])

    doc = Document()

    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3)
        section.right_margin = Cm(3)

    titulo_par = doc.add_paragraph()
    titulo_par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    titulo_par.paragraph_format.space_after = Pt(24)
    run_titulo = titulo_par.add_run(titulo)
    run_titulo.bold = True
    run_titulo.font.size = Pt(24)
    run_titulo.font.name = "Arial"

    for t in terminos:
        nombre = t.get("nombre", "")
        definicion = t.get("definicion", "")
        palabras_clave = [pk.lower() for pk in t.get("palabras_clave", [])]

        par = doc.add_paragraph()
        par.paragraph_format.space_before = Pt(10)
        par.paragraph_format.space_after = Pt(10)

        run_nombre = par.add_run(nombre + ": ")
        run_nombre.bold = True
        run_nombre.font.size = Pt(12)
        run_nombre.font.name = "Arial"

        restante = definicion
        while restante:
            primer_idx = len(restante)
            primer_palabra = None

            for pk in palabras_clave:
                idx = restante.lower().find(pk)
                if idx != -1 and idx < primer_idx:
                    primer_idx = idx
                    primer_palabra = pk

            if primer_palabra is None:
                run = par.add_run(restante)
                run.font.size = Pt(12)
                run.font.name = "Arial"
                break
            else:
                if primer_idx > 0:
                    run = par.add_run(restante[:primer_idx])
                    run.font.size = Pt(12)
                    run.font.name = "Arial"

                run_sub = par.add_run(restante[primer_idx:primer_idx + len(primer_palabra)])
                run_sub.underline = True
                run_sub.font.size = Pt(12)
                run_sub.font.name = "Arial"

                restante = restante[primer_idx + len(primer_palabra):]

    doc.save(output_path)
    print("Word creado")


# MAIN
prompt, seleccion = iu_basica()

if seleccion == "1":
    contenido = generacion_json(prompt, instrucciones_excel)
else:
    contenido = generacion_json(prompt, instrucciones_word)

print("RESPUESTA CRUDA:")
print(contenido)
print("=" * 50)

data = parsear_json(contenido)
if data is None:
    print("No se pudo extraer JSON válido")
    exit()

if seleccion == "1":
    try:
        path = "archivo.xlsx"
        df = pd.DataFrame(data["datos"], columns=data["columnas"])
        df.to_excel(path, index=False)
        estilo = data.get("estilo", {})
        formatear_excel(path, estilo)
        print("Excel creado y formateado")
    except Exception as e:
        print("Error:", e)
else:
    try:
        generar_word(data, "archivo.docx")
    except Exception as e:
        print("Error:", e)

input("Presioná Enter para cerrar...")