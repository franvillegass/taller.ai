from groq import Groq

client = Groq(api_key="dont see brouououou")

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
- Only include "grafico" field if the user explicitly asks for a chart
- "grafico" types: "bar", "line", "pie"
- "columna_x" and "columna_y" must match exact column names in "columnas"
- If a calculated column is needed, use EXACTLY these names: "Total", "Subtotal", "Promedio", "Cantidad"
- "Total" = price × quantity
- "Subtotal" = same as Total before taxes
- "Promedio" = average of a numeric column
- You MUST use the real data provided under "Real data found" as the primary source. Never invent information that contradicts it.
- don't make up or say anything that isn't proven
- No explanations, no markdown, no extra text
- columnas y datos must match in quantity
- Descriptions must be detailed, specific and distinct from each other
- Data must be realistic and coherent with the context
- Never repeat the same value in a description column
- Style must be coherent with the requested Excel theme
- All text content (column names, data) must be written in Spanish
- NEVER fill price or monetary columns with values, always set them to 0. This is mandatory.
"""

instrucciones_word = """
Respond ONLY with valid JSON.

Structure:
{
  "titulo": "Document title in Spanish",
  "terminos": [
    {
      "nombre": "Concept name in Spanish",
      "definicion": "Detailed definition in Spanish.",
      "palabras_clave": ["word1", "word2"]
    }
  ]
}

Rules:
- You MUST use the real data provided under "Real data found" as the primary source. Never invent information that contradicts it.
- No explanations, no markdown, no extra text
- Definitions must be clear, complete and academic, written in Spanish
- palabras_clave are the most important terms within the definition (2-4 per concept)
- Never repeat definitions
- don't make up or say anything that isn't proven
"""