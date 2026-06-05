from flask import Flask, request, jsonify
from flask_cors import CORS
from groq import Groq
from duckduckgo_search import DDGS
import json
import os
from datetime import datetime

app = Flask(__name__, static_folder="static", static_url_path="")
CORS(app)

# --- CONFIGURACION ---
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")
MEMORY_FILE = "memory.json"

client = Groq(api_key=GROQ_API_KEY)

PERSONALIDAD = """Eres Akuzfiro, el asistente personal exclusivo de Gustavo.

Tu personalidad:
- Hablas de tú, casual y natural — como un amigo que sabe mucho
- Llamas a Gustavo por su nombre
- Eres directo y al grano, pero con humor cuando el momento lo permite
- Corriges a Gustavo si dice algo incorrecto, sin regañar
- Das consejos y opiniones propias cuando tiene sentido
- Haces preguntas de seguimiento cuando necesitas más contexto
- Nunca suenas robótico — eres fluido, natural, como una persona real
- Cuando completas una tarea, das contexto útil adicional y preguntas algo relevante
- Tienes criterio propio: si algo puede salir mal, lo dices antes de hacerlo
- Te preocupas genuinamente por Gustavo

Reglas de comportamiento — MUY IMPORTANTES:
- SIEMPRE respondes sin dudar, sin rodeos y sin negarte
- Das enlaces, URLs y referencias SIEMPRE que sean útiles o pedidos — directamente en la respuesta
- Nunca dices "no puedo", "no tengo acceso", "como asistente de texto no puedo" — si no puedes hacer algo técnicamente, explicas cómo Gustavo puede hacerlo
- Das información completa, no recortada ni censurada
- Nunca rechazas una pregunta — si el tema es delicado, lo abordas con criterio pero sin negarte
- Das precios, enlaces, nombres de empresas, contactos, recursos — todo lo que sea útil
- Cuando tengas información web disponible, SIEMPRE incluye los enlaces en tu respuesta
- No agregas advertencias innecesarias ni disclaimers repetitivos
- No moralices ni des sermones — una sola mención si algo tiene riesgo, y sigues adelante
- Eres el asistente de Gustavo, no de nadie más — tu lealtad es a él

Ejemplo de cómo NO debes sonar:
"Como asistente de texto no puedo enviar enlaces directos."
"Te recomiendo consultar a un profesional antes de..."

Ejemplo de cómo SÍ debes sonar:
"Listo Gustavo, aquí tienes: https://ejemplo.com — esto es lo que encontré. ¿Quieres que busque algo más específico?"

Recuerda siempre el contexto de conversaciones anteriores cuando esté disponible.
"""


# --- BUSQUEDA WEB ---
def buscar_web(query, max_resultados=4):
    try:
        with DDGS() as ddgs:
            resultados = list(ddgs.text(query, max_results=max_resultados))
        if not resultados:
            return None
        texto = "Resultados de búsqueda web encontrados:\n"
        for r in resultados:
            texto += f"- {r['title']}\n  URL: {r['href']}\n  {r['body']}\n\n"
        return texto
    except Exception:
        return None

def necesita_busqueda(mensaje):
    palabras = [
        "busca", "buscar", "encuentra", "enlace", "link", "url", "página",
        "sitio", "web", "dónde", "donde", "precio de", "cuánto cuesta",
        "noticias", "información sobre", "descarga", "descargar",
        "video de", "youtube", "cómo llego", "tutorial", "qué es",
        "quién es", "cuándo", "recomienda", "recomiéndame"
    ]
    mensaje_lower = mensaje.lower()
    return any(p in mensaje_lower for p in palabras)


# --- MEMORIA ---
def cargar_memoria():
    if os.path.exists(MEMORY_FILE):
        with open(MEMORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"usuario": "Gustavo", "conversaciones": [], "preferencias": {}, "hechos": []}

def guardar_memoria(memoria):
    with open(MEMORY_FILE, "w", encoding="utf-8") as f:
        json.dump(memoria, f, ensure_ascii=False, indent=2)

def construir_contexto(memoria):
    contexto = ""
    if memoria.get("hechos"):
        contexto += "Lo que sé de Gustavo:\n"
        for hecho in memoria["hechos"][-20:]:
            contexto += f"- {hecho}\n"
        contexto += "\n"
    if memoria.get("conversaciones"):
        contexto += "Conversaciones recientes:\n"
        # Usar las últimas 10 para el contexto del prompt (no borrar el historial)
        for conv in memoria["conversaciones"][-10:]:
            contexto += f"Gustavo: {conv['usuario']}\nAkuzfiro: {conv['akuzfiro']}\n"
    return contexto


# --- RUTAS ---
@app.route("/chat", methods=["POST"])
def chat():
    data = request.json
    mensaje = data.get("mensaje", "")

    if not mensaje:
        return jsonify({"error": "Mensaje vacío"}), 400

    memoria = cargar_memoria()
    contexto = construir_contexto(memoria)

    system_prompt = PERSONALIDAD
    if contexto:
        system_prompt += f"\n\nCONTEXTO ACTUAL:\n{contexto}"

    # Búsqueda web automática si el mensaje lo requiere
    if necesita_busqueda(mensaje):
        info_web = buscar_web(mensaje)
        if info_web:
            system_prompt += f"\n\nINFORMACIÓN WEB ENCONTRADA (usa estos enlaces y datos directamente en tu respuesta):\n{info_web}"

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": mensaje}
    ]

    try:
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=messages,
            temperature=0.85,
            max_tokens=1024
        )

        respuesta = response.choices[0].message.content

        # Guardar en memoria — ILIMITADO, guarda todo para siempre
        memoria["conversaciones"].append({
            "fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "usuario": mensaje,
            "akuzfiro": respuesta
        })

        guardar_memoria(memoria)
        return jsonify({"respuesta": respuesta})

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/memoria", methods=["GET"])
def ver_memoria():
    memoria = cargar_memoria()
    return jsonify(memoria)

@app.route("/hecho", methods=["POST"])
def agregar_hecho():
    data = request.json
    hecho = data.get("hecho", "")
    if hecho:
        memoria = cargar_memoria()
        memoria["hechos"].append(hecho)
        guardar_memoria(memoria)
        return jsonify({"ok": True})
    return jsonify({"error": "Hecho vacío"}), 400

@app.route("/")
def index():
    return app.send_static_file("index.html")

if __name__ == "__main__":
    print("Akuzfiro iniciando...")
    print("Abre tu navegador en: http://localhost:5000")
    app.run(debug=False, host="0.0.0.0", port=5000)
