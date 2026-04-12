from back.ai import mejorar_prompt, generacion_json
from front.iu_consola import iu_basica, post_generacion
from back.config import instrucciones_excel, instrucciones_word
from back.data import buscar_datos_web

# MAIN
prompt, seleccion = iu_basica()

prompt_mejorado = mejorar_prompt(prompt, seleccion)
print("PROMPT MEJORADO:", prompt_mejorado)
print("=" * 50)

print("Buscando datos en la web...")
datos_web = buscar_datos_web(prompt)
print("DATOS WEB:", datos_web)
print("=" * 50)

if seleccion == "1":
    contenido = generacion_json(f"{prompt_mejorado}\n\nReal data found:\n{datos_web}", instrucciones_excel)
else:
    contenido = generacion_json(f"{prompt_mejorado}\n\nReal data found:\n{datos_web}", instrucciones_word)

print("RESPUESTA CRUDA:")
print(contenido)
print("=" * 50)
post_generacion(seleccion, contenido)
