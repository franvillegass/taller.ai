import customtkinter as ctk
import threading
import json
import os
from datetime import datetime

# Importar backend
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Colores y tema
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

COLORS = {
    "bg": "#0f0f0f",
    "surface": "#1a1a1a",
    "surface2": "#242424",
    "accent": "#00ff88",
    "accent_dim": "#00cc6a",
    "text": "#ffffff",
    "text_dim": "#888888",
    "ai_bubble": "#1e2d1e",
    "user_bubble": "#1a1a2e",
    "border": "#2a2a2a",
    "error": "#ff4444",
}

BIBLIOTECA_PATH = "biblioteca.json"

def cargar_biblioteca():
    if os.path.exists(BIBLIOTECA_PATH):
        with open(BIBLIOTECA_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def guardar_biblioteca(entries):
    with open(BIBLIOTECA_PATH, "w", encoding="utf-8") as f:
        json.dump(entries, f, ensure_ascii=False, indent=2)

def agregar_a_biblioteca(nombre, tipo, path):
    entries = cargar_biblioteca()
    entries.insert(0, {
        "nombre": nombre,
        "tipo": tipo,
        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "path": path
    })
    guardar_biblioteca(entries)


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("taller.ai")
        self.geometry("800x600")
        self.minsize(700, 500)
        self.configure(fg_color=COLORS["bg"])
        self.current_frame = None
        self.mostrar_menu()

    def cambiar_frame(self, frame_class, **kwargs):
        if self.current_frame:
            self.current_frame.destroy()
        self.current_frame = frame_class(self, **kwargs)
        self.current_frame.pack(fill="both", expand=True)

    def mostrar_menu(self):
        self.cambiar_frame(MenuFrame)

    def mostrar_chat(self, tipo):
        self.cambiar_frame(ChatFrame, tipo=tipo)

    def mostrar_biblioteca(self):
        self.cambiar_frame(BibliotecaFrame)


class MenuFrame(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color=COLORS["bg"])

        # Título
        ctk.CTkLabel(
            self,
            text="taller.ai",
            font=ctk.CTkFont(family="Courier New", size=48, weight="bold"),
            text_color=COLORS["accent"]
        ).pack(pady=(80, 8))

        ctk.CTkLabel(
            self,
            text="generador de documentos con IA",
            font=ctk.CTkFont(size=14),
            text_color=COLORS["text_dim"]
        ).pack(pady=(0, 60))

        # Botones
        btn_cfg = dict(
            width=260, height=52,
            font=ctk.CTkFont(size=15, weight="bold"),
            corner_radius=8,
        )

        ctk.CTkButton(
            self,
            text="📊  Generar Excel",
            fg_color=COLORS["surface2"],
            hover_color=COLORS["accent"],
            text_color=COLORS["text"],
            command=lambda: master.mostrar_chat("excel"),
            **btn_cfg
        ).pack(pady=8)

        ctk.CTkButton(
            self,
            text="📄  Generar Word",
            fg_color=COLORS["surface2"],
            hover_color=COLORS["accent"],
            text_color=COLORS["text"],
            command=lambda: master.mostrar_chat("word"),
            **btn_cfg
        ).pack(pady=8)

        ctk.CTkButton(
            self,
            text="📁  Biblioteca",
            fg_color=COLORS["surface"],
            hover_color=COLORS["surface2"],
            text_color=COLORS["text_dim"],
            command=master.mostrar_biblioteca,
            **btn_cfg
        ).pack(pady=8)


class ChatFrame(ctk.CTkFrame):
    def __init__(self, master, tipo):
        super().__init__(master, fg_color=COLORS["bg"])
        self.master = master
        self.tipo = tipo  # "excel" o "word"
        self.seleccion = "1" if tipo == "excel" else "2"
        self.estado = "esperando_prompt"  # estados: esperando_prompt, esperando_nombre, editando
        self.ultimo_json = None
        self.ultimo_contenido = None

        self._build_ui()
        self.after(300, self._saludo_inicial)

    def _build_ui(self):
        # Header
        header = ctk.CTkFrame(self, fg_color=COLORS["surface"], height=52, corner_radius=0)
        header.pack(fill="x")
        header.pack_propagate(False)

        ctk.CTkButton(
            header, text="←", width=40, height=36,
            fg_color="transparent", hover_color=COLORS["surface2"],
            text_color=COLORS["text_dim"],
            command=self.master.mostrar_menu,
            font=ctk.CTkFont(size=18)
        ).pack(side="left", padx=8, pady=8)

        tipo_str = "Excel" if self.tipo == "excel" else "Word"
        ctk.CTkLabel(
            header,
            text=f"Generar {tipo_str}",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color=COLORS["text"]
        ).pack(side="left", padx=4)

        # Área de chat
        self.chat_scroll = ctk.CTkScrollableFrame(
            self, fg_color=COLORS["bg"], corner_radius=0
        )
        self.chat_scroll.pack(fill="both", expand=True, padx=0, pady=0)

        # Input
        input_bar = ctk.CTkFrame(self, fg_color=COLORS["surface"], height=64, corner_radius=0)
        input_bar.pack(fill="x")
        input_bar.pack_propagate(False)

        self.input_box = ctk.CTkEntry(
            input_bar,
            placeholder_text="Escribí acá...",
            fg_color=COLORS["surface2"],
            border_color=COLORS["border"],
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=13),
            height=40,
            corner_radius=8
        )
        self.input_box.pack(side="left", fill="x", expand=True, padx=(12, 8), pady=12)
        self.input_box.bind("<Return>", lambda e: self._enviar())

        self.btn_enviar = ctk.CTkButton(
            input_bar,
            text="Enviar",
            width=90, height=40,
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_dim"],
            text_color="#000000",
            font=ctk.CTkFont(size=13, weight="bold"),
            corner_radius=8,
            command=self._enviar
        )
        self.btn_enviar.pack(side="right", padx=(0, 12), pady=12)

    def _agregar_burbuja(self, texto, es_ia=True):
        color = COLORS["ai_bubble"] if es_ia else COLORS["user_bubble"]
        anchor = "w" if es_ia else "e"
        prefix = "🤖  " if es_ia else ""

        wrapper = ctk.CTkFrame(self.chat_scroll, fg_color="transparent")
        wrapper.pack(fill="x", pady=4, padx=12)

        burbuja = ctk.CTkLabel(
            wrapper,
            text=prefix + texto,
            fg_color=color,
            text_color=COLORS["text"],
            font=ctk.CTkFont(size=13),
            corner_radius=10,
            wraplength=520,
            justify="left",
            anchor="w",
            padx=14, pady=10
        )
        burbuja.pack(anchor=anchor)
        self.after(50, lambda: self.chat_scroll._parent_canvas.yview_moveto(1.0))

    def _set_input(self, habilitado):
        state = "normal" if habilitado else "disabled"
        self.input_box.configure(state=state)
        self.btn_enviar.configure(state=state)

    def _saludo_inicial(self):
        tipo_str = "Excel" if self.tipo == "excel" else "Word"
        self._agregar_burbuja(f"Hola! Describí el {tipo_str} que querés generar.")

    def _enviar(self):
        texto = self.input_box.get().strip()
        if not texto:
            return

        self.input_box.delete(0, "end")
        self._agregar_burbuja(texto, es_ia=False)
        self._set_input(False)

        if self.estado == "esperando_prompt":
            self._agregar_burbuja("Procesando tu pedido...")
            threading.Thread(target=self._generar, args=(texto,), daemon=True).start()

        elif self.estado == "esperando_nombre":
            self._guardar_archivo(texto)

        elif self.estado == "editando":
            self._agregar_burbuja("Aplicando cambios...")
            threading.Thread(target=self._editar, args=(texto,), daemon=True).start()

    def _generar(self, prompt):
        try:
            from back.ai import mejorar_prompt, buscar_datos_web, generacion_json, parsear_json
            from back.config import instrucciones_excel, instrucciones_word

            self._agregar_burbuja("Mejorando tu prompt...")
            prompt_mejorado = mejorar_prompt(prompt, self.seleccion)

            self._agregar_burbuja("Buscando datos en la web...")
            datos_web = buscar_datos_web(prompt)

            instrucciones = instrucciones_excel if self.seleccion == "1" else instrucciones_word
            contenido = generacion_json(
                f"{prompt_mejorado}\n\nReal data found:\n{datos_web}",
                instrucciones
            )

            data = parsear_json(contenido)
            if data is None:
                self._agregar_burbuja("No pude generar un resultado válido. Intentá de nuevo.")
                self._set_input(True)
                return

            self.ultimo_json = data
            self.ultimo_contenido = contenido
            self.estado = "esperando_nombre"
            self._agregar_burbuja(f"¡Listo! ¿Qué nombre le ponés al archivo? (sin extensión)")
            self._set_input(True)

        except Exception as e:
            self._agregar_burbuja(f"Error: {str(e)}")
            self._set_input(True)

    def _guardar_archivo(self, nombre):
        try:
            import pandas as pd
            from back.excel import formatear_excel, aplicar_formulas
            from back.word import generar_word
            from openpyxl import load_workbook

            if self.seleccion == "1":
                os.makedirs("excels", exist_ok=True)
                path = f"excels/{nombre}.xlsx"
                df = pd.DataFrame(self.ultimo_json["datos"], columns=self.ultimo_json["columnas"])
                df.to_excel(path, index=False)
                estilo = self.ultimo_json.get("estilo", {})
                formatear_excel(path, estilo)
                wb = load_workbook(path)
                aplicar_formulas(wb.active)
                wb.save(path)
            else:
                os.makedirs("words", exist_ok=True)
                path = f"words/{nombre}.docx"
                generar_word(self.ultimo_json, path)

            agregar_a_biblioteca(nombre, self.tipo, path)
            self._agregar_burbuja(f"Archivo guardado en {path} ✓")
            self._agregar_burbuja("¿Querés modificar algo? Describí los cambios o escribí 'no' para terminar.")
            self.estado = "editando"
            self._set_input(True)

        except Exception as e:
            self._agregar_burbuja(f"Error al guardar: {str(e)}")
            self._set_input(True)

    def _editar(self, pedido):
        if pedido.lower() in ["no", "no gracias", "listo", "nada"]:
            self._agregar_burbuja("Perfecto. Podés volver al menú cuando quieras.")
            self._set_input(True)
            return

        try:
            from back.ai import editar_json, parsear_json
            from back.config import instrucciones_excel, instrucciones_word
            from back.excel import formatear_excel, aplicar_formulas
            from back.word import generar_word
            from openpyxl import load_workbook
            import pandas as pd

            instrucciones = instrucciones_excel if self.seleccion == "1" else instrucciones_word
            contenido = editar_json(
                json.dumps(self.ultimo_json, ensure_ascii=False),
                pedido,
                instrucciones
            )

            data = parsear_json(contenido)
            if data is None:
                self._agregar_burbuja("No pude aplicar los cambios. Intentá describir la modificación de otra forma.")
                self._set_input(True)
                return

            self.ultimo_json = data

            # Reescribir el archivo existente
            entries = cargar_biblioteca()
            path = entries[0]["path"] if entries else None

            if path:
                if self.seleccion == "1":
                    df = pd.DataFrame(data["datos"], columns=data["columnas"])
                    df.to_excel(path, index=False)
                    estilo = data.get("estilo", {})
                    formatear_excel(path, estilo)
                    wb = load_workbook(path)
                    aplicar_formulas(wb.active)
                    wb.save(path)
                else:
                    generar_word(data, path)

            self._agregar_burbuja("Cambios aplicados ✓ ¿Algo más?")
            self._set_input(True)

        except Exception as e:
            self._agregar_burbuja(f"Error: {str(e)}")
            self._set_input(True)


class BibliotecaFrame(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color=COLORS["bg"])

        # Header
        header = ctk.CTkFrame(self, fg_color=COLORS["surface"], height=52, corner_radius=0)
        header.pack(fill="x")
        header.pack_propagate(False)

        ctk.CTkButton(
            header, text="←", width=40, height=36,
            fg_color="transparent", hover_color=COLORS["surface2"],
            text_color=COLORS["text_dim"],
            command=master.mostrar_menu,
            font=ctk.CTkFont(size=18)
        ).pack(side="left", padx=8, pady=8)

        ctk.CTkLabel(
            header,
            text="Biblioteca",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color=COLORS["text"]
        ).pack(side="left", padx=4)

        # Lista
        scroll = ctk.CTkScrollableFrame(self, fg_color=COLORS["bg"])
        scroll.pack(fill="both", expand=True, padx=16, pady=16)

        entries = cargar_biblioteca()

        if not entries:
            ctk.CTkLabel(
                scroll,
                text="No hay archivos generados todavía.",
                text_color=COLORS["text_dim"],
                font=ctk.CTkFont(size=14)
            ).pack(pady=40)
            return

        for entry in entries:
            self._agregar_entrada(scroll, entry)

    def _agregar_entrada(self, parent, entry):
        row = ctk.CTkFrame(parent, fg_color=COLORS["surface"], corner_radius=8)
        row.pack(fill="x", pady=5)

        icono = "📊" if entry["tipo"] == "excel" else "📄"
        ctk.CTkLabel(
            row,
            text=f"{icono}  {entry['nombre']}",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COLORS["text"],
            anchor="w"
        ).pack(side="left", padx=14, pady=12)

        ctk.CTkLabel(
            row,
            text=entry["fecha"],
            font=ctk.CTkFont(size=12),
            text_color=COLORS["text_dim"]
        ).pack(side="left", padx=8)

        ctk.CTkButton(
            row,
            text="Abrir",
            width=70, height=30,
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_dim"],
            text_color="#000000",
            font=ctk.CTkFont(size=12, weight="bold"),
            corner_radius=6,
            command=lambda p=entry["path"]: os.startfile(p) if os.path.exists(p) else None
        ).pack(side="right", padx=14, pady=12)


if __name__ == "__main__":
    app = App()
    app.mainloop()