#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Interfaz web (Streamlit) para el Limpiador de correos - Fase 1.

Sube un archivo (CSV o Excel), corre el mismo pipeline que la CLI
(estandarización automática de separador/encoding + clasificación de
correos) y ofrece 3 botones de descarga: buenos.xlsx / revisar.xlsx /
eliminar.xlsx. No duplica ninguna lógica: todo se reutiliza directamente
de limpiador_correos_fase1.py.

Uso:
    streamlit run streamlit_app.py

AVISO: si esta app se despliega en un hosting compartido/gratuito que
bloquea el puerto 25 saliente (común en muchos servicios cloud), la
verificación SMTP no va a poder conectarse y todo quedará en REVISAR.
Corré esta interfaz en la misma red/máquina donde ya sabés que la CLI
puede conectarse por SMTP, o usá el checkbox "Desactivar verificación
SMTP" de acá abajo si necesitás igual filtrar por sintaxis/dominio/desechables.
"""

import smtplib
import threading
import time
import urllib.request

import streamlit as st
from streamlit.runtime.scriptrunner import add_script_run_ctx, get_script_run_ctx

from limpiador_correos_fase1 import (
    CONCURRENCIA_DEFAULT,
    CONCURRENCIA_POR_DOMINIO_DEFAULT,
    DNS_TIMEOUT_DEFAULT,
    RUTA_LISTA_NEGRA_LOCAL,
    SMTP_HELO_DOMINIO,
    SMTP_MAIL_FROM,
    SMTP_TIMEOUT_DEFAULT,
    SMTP_TIMEOUT_PROVEEDOR_MASIVO_DEFAULT,
    cargar_o_crear_lista_negra_local,
    clasificar_dataframe,
    estandarizar_entrada,
    generar_excel_en_memoria,
    leer_lista_contactos,
    particionar_por_accion,
    verificar_mx,
)

st.set_page_config(page_title="Limpiador de correos - Fase 1", page_icon="📧")

st.title("📧 Limpiador de correos - Fase 1")
st.caption(
    "Subí tu lista de contactos, verificamos cada correo (sintaxis, dominio, "
    "desechables y SMTP) y descargá los resultados separados en 3 archivos."
)

# ==========================================================================
# BLOQUE TEMPORAL DE DIAGNOSTICO -- SACAR DESPUES DE USAR
# ==========================================================================
# Objetivo: comparar el comportamiento de red de ESTE entorno (Streamlit
# Cloud) contra el de una red local, para explicar por que un mismo archivo
# da MANTENER muy distinto en cada uno. No toca clasificar_email() ni
# ninguna logica de clasificacion real -- hace su propia conversacion SMTP
# cruda (HELO/MAIL FROM/RCPT TO) por fuera del pipeline, solo para mostrar
# el codigo y mensaje TAL CUAL los devuelve el servidor, tanto para una
# direccion real como para una inventada en el mismo dominio.
#
# Interpretacion:
#   - Real=250, Inventada=algo distinto de 250 -> el servidor SI esta
#     verificando el buzon de verdad. El MANTENER de este entorno es confiable
#     para ese dominio.
#   - Real=250 e Inventada=250 -> o el dominio es catch-all (acepta
#     cualquier direccion), o algo en la red de este entorno esta aceptando
#     el RCPT TO sin llegar a verificar contra el buzon real.
def _diagnostico_rcpt_crudo(email: str, mx_host: str, timeout: float):
    smtp = None
    try:
        smtp = smtplib.SMTP(timeout=timeout)
        smtp.connect(mx_host, 25)
        smtp.helo(SMTP_HELO_DOMINIO)
        smtp.mail(SMTP_MAIL_FROM)
        codigo, mensaje = smtp.rcpt(email)
        mensaje_texto = mensaje.decode("utf-8", errors="ignore") if isinstance(mensaje, bytes) else str(mensaje)
        return codigo, mensaje_texto, None
    except Exception as e:
        return None, None, repr(e)
    finally:
        if smtp is not None:
            try:
                smtp.quit()
            except Exception:
                pass


def _obtener_ip_publica_saliente():
    try:
        with urllib.request.urlopen("https://api.ipify.org", timeout=5) as resp:
            return resp.read().decode().strip()
    except Exception:
        return None


with st.expander("🔧 Diagnóstico temporal de red SMTP (sacar este bloque después de usar)"):
    st.caption(
        "Prueba, dominio por dominio, una dirección real (opcional) y una "
        "dirección INVENTADA en el mismo dominio, mostrando el código y "
        "mensaje SMTP crudos que devuelve el servidor -- sin pasar por "
        "ninguna lógica de clasificación. Sirve para ver si este entorno "
        "está verificando de verdad contra el buzón o si algo en la red "
        "está aceptando todo sin comprobar nada."
    )
    entradas_diagnostico = st.text_area(
        "Un caso por línea: 'email_real@dominio.com' (se prueba esa + una inventada en el mismo dominio)",
        value="director@araiindia.com\ncontact@cafindia.org\ndirector@caritasindia.org",
        height=100,
    )

    if st.button("Correr diagnóstico SMTP crudo"):
        with st.spinner("Consultando la IP pública de salida de este entorno..."):
            ip_publica = _obtener_ip_publica_saliente()
        if ip_publica:
            st.write(
                f"**IP pública de salida de este entorno:** `{ip_publica}` — "
                f"[chequear reputación en MXToolbox](https://mxtoolbox.com/SuperTool.aspx?action=blacklist%3a{ip_publica})"
            )
        else:
            st.write("**IP pública de salida de este entorno:** no se pudo determinar.")

        for linea in entradas_diagnostico.splitlines():
            linea = linea.strip()
            if not linea:
                continue
            email_real = linea if "@" in linea else None
            dominio = linea.split("@")[1] if email_real else linea

            st.markdown(f"---\n**Dominio:** `{dominio}`")

            resultado_mx, mx_host, reason_mx = verificar_mx(dominio, DNS_TIMEOUT_DEFAULT)
            if resultado_mx != "ok":
                st.error(f"No se pudo resolver MX ({resultado_mx} / {reason_mx})")
                continue
            st.write(f"MX resuelto: `{mx_host}`")

            email_inventado = f"esteusuarionoexiste-diagnostico-{int(time.time())}@{dominio}"

            if email_real:
                codigo_real, mensaje_real, excepcion_real = _diagnostico_rcpt_crudo(
                    email_real, mx_host, SMTP_TIMEOUT_DEFAULT
                )
                st.code(
                    f"REAL      {email_real}\n"
                    f"          codigo={codigo_real!r}  mensaje={mensaje_real!r}  excepcion={excepcion_real!r}"
                )

            codigo_falso, mensaje_falso, excepcion_falso = _diagnostico_rcpt_crudo(
                email_inventado, mx_host, SMTP_TIMEOUT_DEFAULT
            )
            st.code(
                f"INVENTADA {email_inventado}\n"
                f"          codigo={codigo_falso!r}  mensaje={mensaje_falso!r}  excepcion={excepcion_falso!r}"
            )

            if email_real:
                if codigo_real == 250 and codigo_falso == 250:
                    st.warning(
                        "Las DOS direcciones fueron aceptadas con 250 -> este dominio es "
                        "catch-all, o la red de este entorno está aceptando el RCPT TO sin "
                        "verificar de verdad contra el buzón real."
                    )
                elif codigo_real == 250 and codigo_falso != 250:
                    st.success(
                        "La real fue aceptada (250) y la inventada NO -> el servidor está "
                        "verificando de verdad. El MANTENER de este entorno es confiable "
                        "para este dominio."
                    )
                elif codigo_real != 250:
                    st.info(
                        f"La dirección real NO dio 250 acá (dio {codigo_real!r}) -- "
                        "no coincide con lo esperado para este caso, revisar a mano."
                    )
# ==========================================================================
# FIN DEL BLOQUE TEMPORAL DE DIAGNOSTICO
# ==========================================================================


archivo_subido = st.file_uploader(
    "Archivo de contactos (CSV o Excel)", type=["csv", "xlsx", "xls"]
)

with st.expander("Opciones avanzadas"):
    sin_smtp = st.checkbox(
        "Desactivar verificación SMTP",
        value=False,
        help="Útil si esta app corre en una red/host que bloquea el puerto 25. "
             "Los correos con sintaxis y dominio válidos quedarán en REVISAR "
             "en vez de MANTENER/ELIMINAR.",
    )
    columna_seleccionada = None
    if archivo_subido is not None:
        try:
            archivo_subido.seek(0)
            df_vista_previa = estandarizar_entrada(archivo_subido, archivo_subido.name)
            archivo_subido.seek(0)
            columnas_disponibles = ["Detectar automáticamente"] + list(df_vista_previa.columns)
            columna_elegida = st.selectbox("Columna de email", columnas_disponibles)
            if columna_elegida != "Detectar automáticamente":
                columna_seleccionada = columna_elegida
        except Exception as e:
            st.warning(f"No se pudo leer el archivo todavía para elegir columna: {e}")

# --------------------------------------------------------------------------
# Estado persistente entre reruns (necesario porque la verificación corre en
# un hilo de fondo y el progreso se muestra vía un fragment que se refresca
# solo, de forma independiente al resto del script -- ver más abajo).
# --------------------------------------------------------------------------
if "hilo_clasificacion" not in st.session_state:
    st.session_state.hilo_clasificacion = None
    st.session_state.progreso_completados = 0
    st.session_state.progreso_total = 0
    st.session_state.resultado_final = None
    st.session_state.error_final = None

hay_verificacion_en_curso = st.session_state.hilo_clasificacion is not None

correr = st.button(
    "Verificar correos",
    type="primary",
    disabled=archivo_subido is None or hay_verificacion_en_curso,
)

if correr and archivo_subido is not None:
    try:
        with st.spinner("Leyendo y estandarizando el archivo..."):
            archivo_subido.seek(0)
            df_entrada, columna_email = leer_lista_contactos(
                archivo_subido, columna_seleccionada, archivo_subido.name
            )

        st.info(f"Columna de email usada: **{columna_email}** — {len(df_entrada)} correo(s) a verificar.")

        lista_negra_local = cargar_o_crear_lista_negra_local(RUTA_LISTA_NEGRA_LOCAL)

        st.session_state.progreso_completados = 0
        st.session_state.progreso_total = len(df_entrada)
        st.session_state.resultado_final = None
        st.session_state.error_final = None

        # Streamlit solo resuelve st.session_state al "contexto real" de esta
        # sesion (ScriptRunContext) en el hilo que Streamlit maneja para esta
        # corrida. Un threading.Thread nuevo NO hereda ese contexto solo: sin
        # add_script_run_ctx, cualquier escritura a st.session_state hecha
        # desde ese hilo (o desde los hilos del ThreadPoolExecutor dentro de
        # clasificar_dataframe) queda aislada en un estado "mock" interno de
        # Streamlit y nunca llega a esta sesion real -- por eso el contador
        # de progreso quedaba pegado en 0. Capturamos el contexto actual aca
        # (hilo principal, con contexto real) para propagarlo explicitamente.
        ctx_streamlit = get_script_run_ctx()

        def _actualizar_progreso(completados, total):
            add_script_run_ctx(ctx=ctx_streamlit)
            st.session_state.progreso_completados = completados
            st.session_state.progreso_total = total

        def _correr_clasificacion():
            try:
                st.session_state.resultado_final = clasificar_dataframe(
                    df_entrada, columna_email, lista_negra_local,
                    dns_timeout=DNS_TIMEOUT_DEFAULT,
                    smtp_timeout=SMTP_TIMEOUT_DEFAULT,
                    smtp_timeout_proveedor_masivo=SMTP_TIMEOUT_PROVEEDOR_MASIVO_DEFAULT,
                    verificar_smtp_activo=not sin_smtp,
                    verificacion_paciente=False,
                    concurrencia=CONCURRENCIA_DEFAULT,
                    concurrencia_por_dominio=CONCURRENCIA_POR_DOMINIO_DEFAULT,
                    mostrar_barra_progreso=False,
                    callback_progreso=_actualizar_progreso,
                )
            except Exception as e:
                st.session_state.error_final = e

        hilo = threading.Thread(target=_correr_clasificacion, daemon=True)
        add_script_run_ctx(hilo, ctx=ctx_streamlit)
        st.session_state.hilo_clasificacion = hilo
        hilo.start()
        # Fuerza un rerun inmediato para entrar en el bloque de progreso de
        # abajo ya en este mismo instante (en vez de esperar a la próxima
        # interacción del usuario).
        st.rerun()

    except ValueError as e:
        st.error(str(e))
    except Exception as e:
        st.error(f"Ocurrió un error inesperado: {e}")


# --------------------------------------------------------------------------
# Panel de progreso: usa st.fragment(run_every=...), el mecanismo nativo de
# Streamlit para refrescar SOLO este bloque cada ~2s sin volver a correr todo
# el script y sin que los hilos de verificación toquen la UI directamente
# (los hilos solo escriben en st.session_state vía callback_progreso; quien
# lee ese estado y actualiza la pantalla es siempre este fragment, corriendo
# en el hilo principal de Streamlit). Esto reemplaza el sondeo anterior con
# un "while + time.sleep()" bloqueante, que no refrescaba la pantalla en vivo.
# --------------------------------------------------------------------------
if st.session_state.hilo_clasificacion is not None:

    @st.fragment(run_every="2s")
    def _panel_progreso():
        hilo = st.session_state.hilo_clasificacion
        if hilo is None:
            return

        completados = st.session_state.progreso_completados
        total = st.session_state.progreso_total

        if hilo.is_alive():
            st.text(f"Verificando... {completados} de {total} correos verificados (aproximado).")
            if total:
                st.progress(min(completados / total, 1.0))
        else:
            # El hilo de fondo ya terminó: liberamos el estado de "en curso"
            # y forzamos un rerun completo para mostrar el resumen y las
            # descargas más abajo.
            st.session_state.hilo_clasificacion = None
            st.rerun()

    _panel_progreso()


# --------------------------------------------------------------------------
# Resultado final: se muestra apenas hay un resultado (o error) guardado.
# --------------------------------------------------------------------------
if st.session_state.error_final is not None:
    st.error(f"Ocurrió un error inesperado durante la verificación: {st.session_state.error_final}")

elif st.session_state.resultado_final is not None:
    df_resultado, tiempo_total_segundos = st.session_state.resultado_final
    minutos, segundos = divmod(tiempo_total_segundos, 60)
    st.success(f"Listo en {int(minutos)} min {segundos:.0f} s.")

    particiones = particionar_por_accion(df_resultado)
    total = len(df_resultado)

    def _pct(cantidad):
        return f"{cantidad / total * 100:.1f}%" if total else "0%"

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total", total)
    col2.metric("MANTENER", len(particiones["buenos"]), _pct(len(particiones["buenos"])))
    col3.metric("REVISAR", len(particiones["revisar"]), _pct(len(particiones["revisar"])))
    col4.metric("ELIMINAR", len(particiones["eliminar"]), _pct(len(particiones["eliminar"])))

    st.subheader("Descargar resultados")
    nombres_archivo = {"buenos": "buenos.xlsx", "revisar": "revisar.xlsx", "eliminar": "eliminar.xlsx"}
    etiquetas = {
        "buenos": "⬇️ Descargar buenos.xlsx (MANTENER)",
        "revisar": "⬇️ Descargar revisar.xlsx (REVISAR)",
        "eliminar": "⬇️ Descargar eliminar.xlsx (ELIMINAR)",
    }
    col_a, col_b, col_c = st.columns(3)
    for columna_ui, clave in zip((col_a, col_b, col_c), ("buenos", "revisar", "eliminar")):
        columna_ui.download_button(
            etiquetas[clave],
            data=generar_excel_en_memoria(particiones[clave]),
            file_name=nombres_archivo[clave],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
