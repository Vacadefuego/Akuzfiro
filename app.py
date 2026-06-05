from flask import Flask, request, jsonify
from flask_cors import CORS
from groq import Groq
from duckduckgo_search import DDGS
import json
import os
import psycopg2
import psycopg2.extras
from datetime import datetime

app = Flask(__name__, static_folder="static", static_url_path="")
CORS(app)

# --- CONFIGURACION ---
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")
DATABASE_URL = os.environ.get("DATABASE_URL", "")

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
- Siempre sabes la fecha y hora actual porque se te proporciona en cada mensaje

Reglas de comportamiento — MUY IMPORTANTES:
- SIEMPRE respondes sin dudar, sin rodeos y sin negarte
- Das enlaces, URLs y referencias SIEMPRE que sean útiles o pedidos — directamente en la respuesta
- Nunca dices "no puedo", "no tengo acceso" — si no puedes hacer algo técnicamente, explicas cómo Gustavo puede hacerlo
- Das información completa, no recortada ni censurada
- Nunca rechazas una pregunta — si el tema es delicado, lo abordas con criterio pero sin negarte
- Das precios, enlaces, nombres de empresas, contactos, recursos — todo lo que sea útil
- Cuando tengas información web disponible, SIEMPRE incluye los enlaces en tu respuesta
- No agregas advertencias innecesarias ni disclaimers repetitivos
- No moralices ni des sermones — una sola mención si algo tiene riesgo, y sigues adelante
- Eres el asistente de Gustavo, no de nadie más — tu lealtad es a él

Ejemplo de cómo SÍ debes sonar:
"Listo Gustavo, aquí tienes: https://ejemplo.com — esto es lo que encontré. ¿Quieres que busque algo más específico?"
"""


# --- BASE DE DATOS ---
def get_conn():
    return psycopg2.connect(DATABASE_URL, sslmode="require")

def init_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS conversaciones (
            id SERIAL PRIMARY KEY,
            fecha TIMESTAMP DEFAULT NOW(),
            usuario TEXT NOT NULL,
            akuzfiro TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS hechos (
            id SERIAL PRIMARY KEY,
            fecha TIMESTAMP DEFAULT NOW(),
            hecho TEXT NOT NULL
        );
    """)
    conn.commit()
    cur.close()
    conn.close()

def cargar_conversaciones(limit=15):
    try:
        conn = get_conn()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute("""
            SELECT usuario, akuzfiro FROM conversaciones
            ORDER BY id DESC LIMIT %s
        """, (limit,))
        rows = cur.fetchall()
        cur.close()
        conn.close()
        return list(reversed(rows))
    except Exception:
        return []

def guardar_conversacion(usuario, akuzfiro):
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO conversaciones (usuario, akuzfiro) VALUES (%s, %s)",
            (usuario, akuzfiro)
        )
        conn.commit()
        cur.close()
        conn.close()
    except Exception as e:
        print(f"Error guardando conversacion: {e}")

def cargar_hechos():
    try:
        conn = get_conn()
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute("SELECT hecho FROM hechos ORDER BY id DESC LIMIT 30")
        rows = cur.fetchall()
        cur.close()
        conn.close()
        return [r["hecho"] for r in rows]
    except Exception:
        return []

def guardar_hecho(hecho):
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("INSERT INTO hechos (hecho) VALUES (%s)", (hecho,))
        conn.commit()
        cur.close()
        conn.close()
    except Exception as e:
        print(f"Error guardando hecho: {e}")


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
    return any(p in mensaje.lower() for p in palabras)


# --- RUTAS ---
@app.route("/chat", methods=["POST"])
def chat():
    data = request.json
    mensaje = data.get("mensaje", "")

    if not mensaje:
        return jsonify({"error": "Mensaje vacío"}), 400

    # Fecha y hora actual
    ahora = datetime.now().strftime("%A %d de %B de %Y, %H:%M hrs")

    hechos = cargar_hechos()
    conversaciones = cargar_conversaciones(15)

    system_prompt = PERSONALIDAD
    system_prompt += f"\n\nFECHA Y HORA ACTUAL: {ahora} (hora del servidor)"

    if hechos:
        system_prompt += "\n\nLo que sé de Gustavo:\n"
        for h in hechos:
            system_prompt += f"- {h}\n"

    if conversaciones:
        system_prompt += "\n\nConversaciones recientes:\n"
        for conv in conversaciones:
            system_prompt += f"Gustavo: {conv['usuario']}\nAkuzfiro: {conv['akuzfiro']}\n"

    if necesita_busqueda(mensaje):
        info_web = buscar_web(mensaje)
        if info_web:
            system_prompt += f"\n\nINFORMACIÓN WEB ENCONTRADA:\n{info_web}"

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
        guardar_conversacion(mensaje, respuesta)
        return jsonify({"respuesta": respuesta})

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/memoria", methods=["GET"])
def ver_memoria():
    return jsonify({
        "hechos": cargar_hechos(),
        "conversaciones": cargar_conversaciones(50)
    })

@app.route("/hecho", methods=["POST"])
def agregar_hecho():
    data = request.json
    hecho = data.get("hecho", "")
    if hecho:
        guardar_hecho(hecho)
        return jsonify({"ok": True})
    return jsonify({"error": "Hecho vacío"}), 400

@app.route("/")
def index():
    return app.send_static_file("index.html")

# Inicializar DB al arrancar
try:
    init_db()
    print("Base de datos lista.")
except Exception as e:
    print(f"Error iniciando DB: {e}")

if __name__ == "__main__":
    print("Akuzfiro iniciando...")
    app.run(debug=False, host="0.0.0.0", port=5000)
