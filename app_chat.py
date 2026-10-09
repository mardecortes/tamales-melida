import requests
from datetime import datetime
import io
import time
import pandas as pd
import streamlit as st

# 1. Configuración de la ventana web
st.set_page_config(
    page_title="Tamales Melida", page_icon="🫔", layout="wide", initial_sidebar_state="auto" 
)

# 2. Estilos visuales personalizados (CSS)
st.markdown(
    """
    <style>
    .stApp {
        background-color: #FFFDF7;
        color: #3E2723;
    }
    /* Mostrar el botón del sidebar y destacarlo */
    [data-testid="stSidebarCollapseButton"], 
    [data-testid="collapsedControl"] {
        display: block !important;
        visibility: visible !important;
        position: fixed !important;
        top: 12px !important;
        left: 12px !important;
        z-index: 99999 !important;
        background-color: #D32F2F !important;
        color: white !important;
        border-radius: 8px !important;
        padding: 4px !important;
    }
    .header-banner {
        background-color: #D32F2F;
        padding: 20px;
        border-radius: 15px;
        text-align: center;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .header-banner h1 {
        color: #FFFFFF !important;
        margin: 0;
        font-size: 2.2rem;
    }
    .header-banner p {
        color: #FFECB3 !important;
        margin-top: 5px;
        margin-bottom: 5px;
        font-size: 1.1rem;
    }
    .header-banner .horario-badge {
        display: inline-block;
        background-color: #B71C1C;
        color: #FFFFFF;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.9rem;
        margin-top: 8px;
        font-weight: bold;
    }
    .stChatMessage {
        background-color: #FFF8E7;
        border: 1px solid #FFE082;
        border-radius: 12px;
        padding: 12px;
        margin-bottom: 10px;
    }
    .info-box {
        background-color: #FFF3E0;
        border-left: 5px solid #FF9800;
        padding: 12px 15px;
        border-radius: 8px;
        margin-bottom: 20px;
        font-size: 0.95rem;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# 3. Base de datos del negocio
INFO_NEGOCIO = {
    "precio": "Todos nuestros tamales tienen un costo de **$20 c/u**.",
    "horario": (
        "Nuestro horario de atención es de **10:00 am a 8:00 pm** todos los"
        " días."
    ),
    "pagos": (
        "Aceptamos 4 métodos de pago:\n• 💵 **Efectivo**\n• 💳 **Tarjeta de"
        " débito/crédito**\n• 📱 **Transferencia bancaria**\n• 🤝 **Pago contra"
        " entrega**"
    ),
    "pedidos_especiales": (
        "Los pedidos especiales se deben realizar con mínimo **2 días de"
        " anticipación**."
    ),
    "domicilio": (
        "El servicio a domicilio aplica únicamente en **pedidos mayores a 8"
        " tamales** y tiene un costo extra según la zona."
    ),
    "contacto": (
        "Puedes hacer tus pedidos o cotizaciones al WhatsApp **+52 644 123"
        " 4567**."
    ),
}

MENU_TAMALES = {
    "salados": ["Carne", "Acelga", "Hoja de plátano", "Elote salado"],
    "dulces": ["Elote dulce", "Fresa", "Piña"],
}


# --- FUNCIÓN PARA GENERAR EL ARCHIVO EXCEL EN MEMORIA ---
def generar_excel_bytes(
    nombre_cliente, detalle_sabores, total_tamales, total_precio, direccion=""
):
  resumen_sabores = ", ".join(
      [f"{cant}x {sabor}" for sabor, cant in detalle_sabores.items() if cant > 0]
  )

  nuevo_registro = {
      "Fecha y Hora": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
      "Cliente": [nombre_cliente if nombre_cliente else "Cliente Anónimo"],
      "Detalle del Pedido": [resumen_sabores],
      "Total Tamales": [total_tamales],
      "Monto Total ($ MXN)": [total_precio],
      "Dirección de Envío": [
          direccion if direccion else "Recoge en sucursal / No aplica"
      ],
  }

  for sabor, cant in detalle_sabores.items():
    nuevo_registro[f"Cant. {sabor}"] = [cant]

  df = pd.DataFrame(nuevo_registro)

  buffer = io.BytesIO()
  with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Pedido")
  buffer.seek(0)
  return buffer


# Inicializar estados en sesión
if "mostrar_confirmacion" not in st.session_state:
  st.session_state["mostrar_confirmacion"] = False

# 4. Barra lateral con Cotizador
with st.sidebar:
  st.header("🧮 Cotizador y Registro")
  st.write("Selecciona la cantidad por cada sabor:")

  nombre_cliente = st.text_input("Nombre del Cliente")

  cantidades_sabores = {}

  st.markdown("### 🥑 Tamales Salados")
  for sabor in MENU_TAMALES["salados"]:
    cantidades_sabores[sabor] = st.number_input(
        f"{sabor} ($20 c/u)", min_value=0, max_value=100, value=0, step=1
    )

  st.markdown("### 🍓 Tamales Dulces")
  for sabor in MENU_TAMALES["dulces"]:
    cantidades_sabores[sabor] = st.number_input(
        f"{sabor} ($20 c/u)", min_value=0, max_value=100, value=0, step=1
    )

  total_tamales = sum(cantidades_sabores.values())
  total_precio = total_tamales * 20

  st.markdown("---")
  st.subheader(f"Total: **${total_precio} MXN**")
  st.caption(f"Cantidad total: {total_tamales} tamales")

  # CASILLA DE DIRECCIÓN A PARTIR DE 8 TAMALES
  direccion_envio = ""
  if total_tamales > 0 and total_tamales < 8:
    st.warning("⚠️ El envío a domicilio requiere un mínimo de 8 tamales.")
  elif total_tamales >= 8:
    st.success("✅ ¡Califica para envío a domicilio!")
    direccion_envio = st.text_input(
        "📍 Dirección de envío (Obligatoria para domicilio)",
        placeholder="Ej. Calle Hidalgo #123, Col. Centro",
    )

  st.markdown("---")

  # BOTÓN "GUARDAR"
  if st.button("💾 Guardar", use_container_width=True):
    if total_tamales == 0:
      st.error("⚠️ Debes seleccionar al menos 1 tamal.")
    elif total_tamales >= 8 and not direccion_envio.strip():
      st.error("⚠️ Por favor ingresa la dirección de envío para tu pedido.")
    else:
      st.success("✅ Pedido registrado exitosamente.")

  # BOTÓN "ENVIAR"
  if st.button("📤 Enviar", use_container_width=True):
    if total_tamales == 0:
      st.error("⚠️ Debes seleccionar al menos 1 tamal antes de enviar.")
    elif total_tamales >= 8 and not direccion_envio.strip():
      st.error("⚠️ Por favor ingresa la dirección de envío para tu pedido.")
    else:
      st.session_state["mostrar_confirmacion"] = True

  # CUADRO DE CONFIRMACIÓN
  if st.session_state["mostrar_confirmacion"]:
      # -------------------------------------------------------------
    # Envío del pedido a Google Sheets
    # -------------------------------------------------------------
    WEBHOOK_URL = (
        "https://script.google.com/macros/s/AKfycbysSHnDFF__Iv8IX4bXzLKY4TDZA_SEjpFELfP6Mc5ATd_CW2Z7MfTElH3tOChbsKyT/exec"
    )

    # Convertimos el diccionario/lista de sabores a texto para el resumen
    resumen_sabores = str(cantidades_sabores)

    payload = {
        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "cliente": nombre_cliente if nombre_cliente else "Cliente Anónimo",
        "pedido": resumen_sabores,
        "total_tamales": total_tamales,
        "total_precio": total_precio,
        "direccion": (
            direccion_envio if direccion_envio else "Recoge en sucursal"
        ),
    }

    try:
        response = requests.post(WEBHOOK_URL, json=payload)
        if response.status_code == 200:
            st.success("🎉 ¡Tu pedido ha sido guardado en Google Sheets!")
        else:
            st.error("Hubo un problema al registrar el pedido en la hoja.")
    except Exception as e:
        st.error(f"Error de conexión con Google Sheets: {e}")
    st.markdown("---")
    st.info("📋 **Confirmar Pedido**")
    st.caption(
        f"Cliente: {nombre_cliente if nombre_cliente else 'Cliente Anónimo'}"
    )
    st.caption(f"Total: {total_tamales} tamales (${total_precio} MXN)")

    # Generar el archivo Excel
    excel_data = generar_excel_bytes(
        nombre_cliente,
        cantidades_sabores,
        total_tamales,
        total_precio,
        direccion_envio,
    )

    # Botón de Confirmar que descarga el archivo
    st.download_button(
        label="✅ Confirmar",
        data=excel_data,
        file_name="pedido_tamales.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    if st.button("❌ Cancelar", use_container_width=True):
      st.session_state["mostrar_confirmacion"] = False
      st.rerun(
           
    )
    # Pon tu URL de Apps Script aquí

    
  st.markdown("---")
  st.markdown("💳 **Métodos de pago aceptados:**")
  st.markdown(
      "• Efectivo\n• Transferencia\n• Tarjetas de crédito/débito\n• Pago"
      " contra entrega"
  )

# 5. Encabezado decorado
st.markdown(
    """
    <div class="header-banner">
        <h1>🫔 Tamales Melida</h1>
        <p>¡Los mejores tamales de la región! Consulta menú, precios y haz tu pedido.</p>
        <div class="horario-badge">⏰ Horario: 10:00 am a 8:00 pm</div>
    </div>
""",
    unsafe_allow_html=True,
)

# Tarjeta informativa
st.markdown(
    """
    <div class="info-box">
        💡 <b>Tip rápido:</b> Pregúntame sobre el <b>menú</b>, <b>métodos de pago</b>, <b>horario</b>, servicio a <b>domicilio</b> o usa el cotizador a la izquierda.
    </div>
""",
    unsafe_allow_html=True,
)
st.warning(
    "↖️ **¿Quieres hacer un pedido?** Toca el botón **`>`** en la **esquina superior izquierda** para desplegar el cotizador.",
    icon="🫔"
)

# 6. Chatbot
def responder(mensaje):
  texto = mensaje.lower().strip()

  if any(
      p in texto
      for p in [
          "pago",
          "pagar",
          "tarjeta",
          "efectivo",
          "transferencia",
          "cobro",
          "aceptan",
      ]
  ):
    return INFO_NEGOCIO["pagos"]

  elif any(
      p in texto
      for p in [
          "horario",
          "hora",
          "abren",
          "cierran",
          "abierto",
          "cerrado",
          "atencion",
      ]
  ):
    return INFO_NEGOCIO["horario"]

  elif any(
      p in texto for p in ["tienen de", "tienen ", "buscar ", "hay de", "hay "]
  ):
    termino = (
        texto.replace("tienen de", "")
        .replace("tienen", "")
        .replace("buscar", "")
        .replace("hay de", "")
        .replace("hay", "")
        .replace("tamal de", "")
        .strip()
    )
    coincidencias = []
    if termino:
      for categoria, sabores in MENU_TAMALES.items():
        for sabor in sabores:
          if termino in sabor.lower():
            coincidencias.append(
                f"Tamal de {sabor} ({categoria.capitalize()})"
            )
      if coincidencias:
        return "¡Sí tenemos! Disponible en el menú:\n\n" + "\n".join(
            [f"• {c}" for c in coincidencias]
        )
      return (
          f"Lo siento, no preparamos tamales de '{termino}'. Te invitamos a"
          " revisar nuestro menú disponible."
      )

  elif any(
      p in texto
      for p in ["precio", "costo", "cuanto cuesta", "cuanto valen", "valen"]
  ):
    return INFO_NEGOCIO["precio"]

  elif any(p in texto for p in ["domicilio", "envio", "entrega", "llevar"]):
    return INFO_NEGOCIO["domicilio"]
  elif any(
      p in texto for p in ["especial", "anticipacion", "encargo", "evento"]
  ):
    return INFO_NEGOCIO["pedidos_especiales"]

  elif any(p in texto for p in ["menu", "tamales", "sabores", "dulce", "salado"]):
    for categoria, sabores in MENU_TAMALES.items():
      if categoria in texto:
        return (
            f"Sabores **{categoria.capitalize()}** ($20 c/u):\n\n"
            + "\n".join([f"• Tamal de {s}" for s in sabores])
        )

    menu_completo = ""
    for cat, sabores in MENU_TAMALES.items():
      menu_completo += (
          f"\n**{cat.capitalize()}:**\n"
          + "\n".join([f"• Tamal de {s}" for s in sabores])
          + "\n"
      )
    return (
        f"Todos nuestros tamales cuestan **$20 c/u**. Este es nuestro"
        f" menú:\n{menu_completo}"
    )

  elif any(
      p in texto for p in ["contacto", "telefono", "whatsapp", "pedir", "pedido"]
  ):
    return (
        f"📱 {INFO_NEGOCIO['contacto']}\n\n💳 **Formas de pago:**"
        f" {INFO_NEGOCIO['pagos']}\n\n⏰ Horario: {INFO_NEGOCIO['horario']}"
    )

  else:
    return "¡Hola! Puedo darte detalles sobre nuestro 'menú', 'formas de pago', 'horario', 'precio', servicio a 'domicilio' o cotizar tu pedido."


def generar_efecto_escritura(texto):
  for palabra in texto.split(" "):
    yield palabra + " "
    time.sleep(0.03)


# 7. Interfaz del chat
if "mensajes" not in st.session_state:
  st.session_state.mensajes = [
      {
          "rol": "assistant",
          "contenido": (
              "¡Hola! 🫔 Bienvenido a **Tamales Melida**. ¿En qué te puedo"
              " ayudar hoy? Puedes preguntarme por el menú, precios o métodos"
              " de pago."
          ),
      }
  ]

for msg in st.session_state.mensajes:
  avatar = "👤" if msg["rol"] == "user" else "🫔"
  with st.chat_message(msg["rol"], avatar=avatar):
    st.write(msg["contenido"])

if entrada := st.chat_input("Escribe tu duda sobre los tamales aquí..."):
  st.session_state.mensajes.append({"rol": "user", "contenido": entrada})
  with st.chat_message("user", avatar="👤"):
    st.write(entrada)

  respuesta = responder(entrada)

  with st.chat_message("assistant", avatar="🫔"):
    respuesta_completa = st.write_stream(generar_efecto_escritura(respuesta))

  st.session_state.mensajes.append(
      {"rol": "assistant", "contenido": respuesta_completa}
  )
