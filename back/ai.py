from back.config import client
import json
import re
import requests
from ddgs import DDGS

def mejorar_prompt(prompt_usuario, tipo):
    tipo_str = "Excel con datos tabulares" if tipo == "1" else "Word con definiciones"

    completion = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[
            {"role": "system", "content": f"""
You are an assistant that improves prompts for generating {tipo_str}.
Take the user input and rewrite it clearly, specifically and in detail.
Add missing context, specify number of columns/terms if not mentioned.
Always request detailed descriptions regardless of what the user says.
The final output (Excel/Word content) must be in Spanish.
Return ONLY the improved prompt, no explanations or comments.
"""},
            {"role": "user", "content": prompt_usuario}
        ],
        temperature=0.3
    )
    return completion.choices[0].message.content


def generacion_json(prompt, instrucciones):
    completion = client.chat.completions.create(
        model="openai/gpt-oss-120b", # otra opcion valida es llama-3.3-70b-versatile otro es openai/gpt-oss-120b otra es groq/compound 
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

def buscar_datos_web(prompt_original):
    completion = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[
            {"role": "system", "content": "Extract ONLY the product or brand names from the text. Return them separated by commas, nothing else. Example output: taragui, rosamonte, cbse"},
            {"role": "user", "content": prompt_original}
        ],
        temperature=0,
        max_tokens=30
    )
    keywords = completion.choices[0].message.content.strip()
    print("KEYWORDS:", keywords)

    resultados = []
    for kw in [k.strip() for k in keywords.split(",")]:
        results = DDGS().text(f"{kw} yerba mate Argentina", max_results=2)
        for r in results:
            resultados.append(r["body"])

    return "\n".join(resultados[:6]) if resultados else "No se encontraron datos."