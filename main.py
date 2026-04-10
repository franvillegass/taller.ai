import json
import re
import subprocess
import tempfile
import pandas as pd
from groq import Groq
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os

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


JS_WORD = r"""
const { Document, Packer, Paragraph, TextRun, AlignmentType, UnderlineType } = require('docx');
const fs = require('fs');

const input = JSON.parse(process.argv[2]);
const outputPath = process.argv[3];
const titulo = input.titulo;
const terminos = input.terminos;

const children = [];

children.push(
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 600 },
    children: [
      new TextRun({
        text: titulo,
        bold: true,
        size: 48,
        font: "Arial"
      })
    ]
  })
);

for (const t of terminos) {
  const palabrasClave = t.palabras_clave || [];
  const definicion = t.definicion;
  const runs = [];

  runs.push(new TextRun({
    text: t.nombre + ": ",
    bold: true,
    size: 24,
    font: "Arial"
  }));

  let restante = definicion;
  while (restante.length > 0) {
    let primerIdx = restante.length;
    let primerPalabra = null;

    for (const pk of palabrasClave) {
      const idx = restante.toLowerCase().indexOf(pk.toLowerCase());
      if (idx !== -1 && idx < primerIdx) {
        primerIdx = idx;
        primerPalabra = pk;
      }
    }

    if (primerPalabra === null) {
      runs.push(new TextRun({ text: restante, size: 24, font: "Arial" }));
      break;
    } else {
      if (primerIdx > 0) {
        runs.push(new TextRun({ text: restante.slice(0, primerIdx), size: 24, font: "Arial" }));
      }
      runs.push(new TextRun({
        text: restante.slice(primerIdx, primerIdx + primerPalabra.length),
        size: 24,
        font: "Arial",
        underline: { type: UnderlineType.SINGLE }
      }));
      restante = restante.slice(primerIdx + primerPalabra.length);
    }
  }

  children.push(
    new Paragraph({
      spacing: { before: 300, after: 300 },
      children: runs
    })
  );
}

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outputPath, buffer);
  console.log("Word creado");
});
"""


def generar_word(data, output_path):
    titulo = data.get("titulo", "Documento")
    terminos = data.get("terminos", [])
    payload = json.dumps({"titulo": titulo, "terminos": terminos}, ensure_ascii=False)

    # Guardar el JS en la carpeta del proyecto, no en Temp
    js_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_word_gen.js")
    with open(js_path, "w", encoding="utf-8") as f:
        f.write(JS_WORD)

    result = subprocess.run(
        ["node", js_path, payload, output_path],
        capture_output=True, text=True,
        cwd=os.path.dirname(os.path.abspath(__file__))  # ejecutar desde la carpeta del proyecto
    )

    if result.returncode != 0:
        print("Error al generar Word:", result.stderr)
    else:
        print(result.stdout.strip())

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