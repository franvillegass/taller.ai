from groq import Groq

client = Groq(api_key="dont sebrou")

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
- CRITICAL: Each row must represent ONE product. Never split products across columns by brand. All products go in the same columns regardless of brand.

Example of ideal output:
User asks for: "table of yerba mate products with name, description and origin"

{
  "datos": [
    ["Taragüi", "Yerba mate líder del mercado argentino, producida por Establecimiento Las Marías en Corrientes desde 1924. Sabor intenso y equilibrado, con más de 60.000 toneladas producidas por año.", "Corrientes, Argentina"],
    ["Rosamonte", "Yerba tradicional fundada en 1936 en Apóstoles, Misiones. Reconocida por su sabor robusto y su proceso de curado prolongado con palo de rosa, que le da un aroma distintivo.", "Misiones, Argentina"],
    ["CBSé", "Empresa familiar fundada en 1978 en San Francisco, Córdoba. Primera yerba mate compuesta de Argentina, pionera en mezclas con hierbas serranas como menta, poleo y peperina.", "Córdoba, Argentina"]
  ],
  "columnas": ["Nombre", "Descripción", "Origen"],
  "estilo": {
    "font_size": 11,
    "header_color": "2E7D32",
    "font_color_header": "FFFFFF",
    "row_alt_color": "E8F5E9"
  }
}

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

Example of ideal output:
User asks for: "definitions of ente and sociedad simple"

{
  "titulo": "Conceptos Fundamentales de Organizaciones",
  "terminos": [
    {
      "nombre": "Ente",
      "definicion": "Un ente es toda entidad, física o jurídica, con capacidad para adquirir derechos y contraer obligaciones dentro del ordenamiento jurídico. En el ámbito económico, un ente es cualquier organización con existencia propia, independiente de las personas que la integran, capaz de realizar actos con efectos legales y patrimoniales.",
      "palabras_clave": ["entidad", "derechos", "obligaciones", "ordenamiento jurídico"]
    },
    {
      "nombre": "Sociedad Simple",
      "definicion": "La sociedad simple es la forma asociativa más básica del derecho privado argentino, caracterizada por la ausencia de personalidad jurídica propia y la responsabilidad ilimitada y solidaria de sus socios frente a terceros. Se constituye cuando dos o más personas acuerdan ejercer en común una actividad económica sin adoptar otro tipo societario.",
      "palabras_clave": ["responsabilidad ilimitada", "personalidad jurídica", "socios", "actividad económica"]
    }
  ]
}

"""