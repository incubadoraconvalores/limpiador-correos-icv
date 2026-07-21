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

import threading

import streamlit as st
from streamlit.runtime.scriptrunner import add_script_run_ctx, get_script_run_ctx

from limpiador_correos_fase1 import (
    CONCURRENCIA_DEFAULT,
    CONCURRENCIA_POR_DOMINIO_DEFAULT,
    DNS_TIMEOUT_DEFAULT,
    RUTA_LISTA_NEGRA_LOCAL,
    SMTP_TIMEOUT_DEFAULT,
    SMTP_TIMEOUT_PROVEEDOR_MASIVO_DEFAULT,
    cargar_o_crear_lista_negra_local,
    clasificar_dataframe,
    estandarizar_entrada,
    generar_excel_en_memoria,
    leer_lista_contactos,
    particionar_por_accion,
)

st.set_page_config(page_title="Limpiador de correos - Fase 1", page_icon="📧")

st.title("📧 Limpiador de correos - Fase 1")
st.caption(
    "Subí tu lista de contactos, verificamos cada correo (sintaxis, dominio, "
    "desechables y SMTP) y descargá los resultados separados en 3 archivos."
)

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
