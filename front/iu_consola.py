from back.word import generar_word
from back.excel import formatear_excel
from back.ai import parsear_json
import os
import pandas as pd

def iu_basica():
    seleccion = input("Introduzca 1 para generar un Excel y 2 para un Word: ")
    if seleccion == "1":
        prompt = input("Describa el Excel: ")
    else:
        prompt = input("Describa el Word: ")
    return prompt, seleccion

def post_generacion(contenido, seleccion):
    data = parsear_json(contenido)
    if data is None:
        print("No se pudo extraer JSON válido")
        input("Presioná Enter para cerrar...")
        exit()

    if seleccion == "1":
        try:
            nombre = input("¿Qué nombre le ponés al archivo Excel? (sin extensión): ")
            os.makedirs("excels", exist_ok=True)
            path = f"excels/{nombre}.xlsx"
            df = pd.DataFrame(data["datos"], columns=data["columnas"])
            df.to_excel(path, index=False)
            estilo = data.get("estilo", {})
            formatear_excel(path, estilo)
            print(f"Excel creado: {path}")
            print("NOTA, MUY IMPORTANTE POR FAVOR LEEEEEEEEEEEEEEEEEEEEEEEEEEEEEER: holaaa soy fran porfa porfa revisa lo que haya generado  porque se puede equivocar, el modulo con el modelo de ia que busca la info esta medio gaga aveces, gracias x usar:D ")
        except Exception as e:
            print("Error:", e)
    else:
        try:
            nombre = input("¿Qué nombre le ponés al archivo Word? (sin extensión): ")
            os.makedirs("words", exist_ok=True)
            path = f"words/{nombre}.docx"
            generar_word(data, path)
        except Exception as e:
            print("Error:", e)