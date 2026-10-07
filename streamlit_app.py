#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Interfaz web (Streamlit) para el Limpiador de correos - Fase 1.

Sube un archivo (CSV o Excel), corre el mismo pipeline que la CLI
(estandarización automática de separador/encoding + clasificación de
correos) y ofrece 2 botones de descarga: buenos.xlsx / eliminar.xlsx.
No duplica ninguna lógica: todo se reutiliza directamente de
limpiador_correos_fase1.py.

Uso:
    streamlit run streamlit_app.py

AVISO: si esta app se despliega en un hosting compartido/gratuito que
bloquea el puerto 25 saliente (común en muchos servicios cloud), la
verificación SMTP no va a poder conectarse y todo quedará en MANTENER
sin confirmación real (ACTUALIZACIÓN 6 de limpiador_correos_fase1.py: ya
no existe REVISAR, y sin señal SMTP no hay evidencia para mandar nada a
ELIMINAR). Corré esta interfaz en la misma red/máquina donde ya sabés que
la CLI puede conectarse por SMTP, o usá el checkbox "Desactivar verificación
SMTP" de acá abajo si necesitás igual filtrar por sintaxis/dominio/desechables.
"""

import threading
import time
import urllib.request
import uuid

import dns.resolver
import streamlit as st

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
    nombre_archivo_salida,
    particionar_por_accion,
)

st.set_page_config(page_title="Limpiador de correos - Fase 1", page_icon="📧")

st.title("📧 Limpiador de correos - Fase 1")
st.caption(
    "Subí tu lista de contactos, verificamos cada correo (sintaxis, dominio, "
    "desechables y SMTP) y descargá los resultados separados en 2 archivos."
)

# ==========================================================================
# BLOQUE TEMPORAL DE DIAGNOSTICO -- SACAR DESPUES DE USAR
# ==========================================================================
# Objetivo: revisar si la IP de salida ACTUAL de este entorno (que puede
# cambiar entre reboots de Streamlit Cloud) esta listada en Spamhaus Zen
# (SBL/XBL/PBL) en este momento. Muchas verificaciones seguidas en poco
# tiempo pueden hacer que la IP quede listada "en caliente" durante el dia,
# lo que explicaria resultados que empeoran corrida tras corrida sin que
# haya cambiado nada en el codigo de clasificacion.
def _obtener_ip_publica_saliente():
    try:
        with urllib.request.urlopen("https://api.ipify.org", timeout=5) as resp:
            return resp.read().decode().strip()
    except Exception:
        return None


with st.expander("🔧 Diagnóstico temporal: IP de salida y Spamhaus (sacar después de usar)", expanded=True):
    if st.button("Chequear IP de salida y Spamhaus"):
        with st.spinner("Detectando IP pública de salida de este entorno..."):
            ip_publica = _obtener_ip_publica_saliente()

        if not ip_publica:
            st.error("No se pudo determinar la IP pública de salida.")
        else:
            st.write(f"**IP pública de salida de este entorno:** `{ip_publica}`")
            ip_invertida = ".".join(reversed(ip_publica.split(".")))
            consulta = f"{ip_invertida}.zen.spamhaus.org"
            try:
                respuestas = dns.resolver.resolve(consulta, "A", lifetime=5)
                codigos = [r.to_text() for r in respuestas]
                st.error(
                    f"⚠️ LISTADA en Spamhaus Zen ahora mismo: {', '.join(codigos)} "
                    f"(consulta `{consulta}`). Esto explicaría que las verificaciones SMTP "
                    "den resultados ambiguos/negativos en casi todos los dominios sin "
                    "importar el proveedor."
                )
            except dns.resolver.NXDOMAIN:
                st.success(f"✅ NO está listada en Spamhaus Zen (SBL/XBL/PBL) ahora mismo. Consulta: `{consulta}`")
            except Exception as e:
                st.warning(f"No se pudo consultar Spamhaus: {e}")
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
             "Los correos con sintaxis y dominio válidos quedarán en MANTENER "
             "sin confirmación real (reason 'unsupported').",
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
# Verificaciones guardadas en el SERVIDOR, no en la sesión del navegador.
#
# Antes todo el estado (hilo, progreso, resultado) vivía en st.session_state,
# que es por sesión de navegador: si la pestaña perdía la conexión durante
# una corrida larga (portátil suspendido, cambio de red, pestaña mucho rato
# en segundo plano), Streamlit abría una sesión NUEVA y la app "volvía a
# inicio" sin ningún error, mientras la verificación seguía corriendo en el
# servidor sin que nadie pudiera ver el resultado (caso real: BBDD
# Emprendedores Español, ~7.000 correos, octubre 2026).
#
# Ahora cada verificación es un TrabajoVerificacion guardado en un registro
# del proceso (st.cache_resource, compartido entre sesiones), identificado
# por un id que también va en la URL (?trabajo=...). Al reconectar o
# recargar la página, la sesión nueva lee ese id de la URL y retoma el mismo
# trabajo: progreso en vivo o resultados listos para descargar.
# Límite conocido: un reinicio del servidor (ej. cada redeploy tras un push)
# sí borra el registro y corta las verificaciones en curso.
# --------------------------------------------------------------------------
HORAS_RETENCION_TRABAJOS = 24


class TrabajoVerificacion:
    def __init__(self, nombre_archivo_original: str, columna_email: str, total: int):
        self.nombre_archivo_original = nombre_archivo_original
        self.columna_email = columna_email
        self.creado_en = time.time()
        self.completados = 0
        self.total = total
        self.resultado = None  # (df_resultado, tiempo_total_segundos)
        self.error = None
        self.hilo = None
        self._excels = {}
        self._lock_excels = threading.Lock()

    def en_curso(self) -> bool:
        return self.hilo is not None and self.hilo.is_alive()

    def excel(self, particiones: dict, clave: str) -> bytes:
        # Se genera una sola vez por archivo: los reruns de Streamlit (y las
        # reconexiones) no vuelven a armar el .xlsx de miles de filas.
        with self._lock_excels:
            if clave not in self._excels:
                self._excels[clave] = generar_excel_en_memoria(particiones[clave])
            return self._excels[clave]


@st.cache_resource
def _registro_trabajos() -> dict:
    return {}


def _limpiar_trabajos_viejos(registro: dict):
    limite = time.time() - HORAS_RETENCION_TRABAJOS * 3600
    for id_trabajo, trabajo in list(registro.items()):
        if trabajo.creado_en < limite and not trabajo.en_curso():
            registro.pop(id_trabajo, None)


registro_trabajos = _registro_trabajos()
id_trabajo_actual = st.query_params.get("trabajo")
trabajo_actual = registro_trabajos.get(id_trabajo_actual) if id_trabajo_actual else None

if id_trabajo_actual and trabajo_actual is None:
    st.warning(
        "No se encontró la verificación de este enlace. Puede que el servidor se "
        "haya reiniciado (por ejemplo, tras una actualización de la app) o que "
        f"tenga más de {HORAS_RETENCION_TRABAJOS} h. Vuelve a subir el archivo."
    )

hay_verificacion_en_curso = trabajo_actual is not None and trabajo_actual.en_curso()

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

        lista_negra_local = cargar_o_crear_lista_negra_local(RUTA_LISTA_NEGRA_LOCAL)

        _limpiar_trabajos_viejos(registro_trabajos)
        trabajo = TrabajoVerificacion(archivo_subido.name, columna_email, len(df_entrada))

        # Los hilos de verificación solo escriben en el objeto 'trabajo' (no
        # en st.session_state), así que no necesitan el contexto de Streamlit
        # de ninguna sesión: siguen funcionando aunque el navegador se
        # desconecte.
        def _actualizar_progreso(completados, total):
            trabajo.completados = completados
            trabajo.total = total

        def _correr_clasificacion():
            try:
                trabajo.resultado = clasificar_dataframe(
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
                trabajo.error = e

        trabajo.hilo = threading.Thread(target=_correr_clasificacion, daemon=True)
        id_nuevo = uuid.uuid4().hex
        registro_trabajos[id_nuevo] = trabajo
        trabajo.hilo.start()

        st.query_params["trabajo"] = id_nuevo
        # Rerun inmediato para entrar ya en el panel de progreso de abajo.
        st.rerun()

    except ValueError as e:
        st.error(str(e))
    except Exception as e:
        st.error(f"Ocurrió un error inesperado: {e}")


# --------------------------------------------------------------------------
# Panel de progreso: st.fragment(run_every=...) refresca SOLO este bloque
# cada ~2s leyendo el objeto del trabajo; cuando el hilo termina, fuerza un
# rerun completo para mostrar el resumen y las descargas.
# --------------------------------------------------------------------------
if trabajo_actual is not None and trabajo_actual.en_curso():
    st.info(
        f"Verificando **{trabajo_actual.nombre_archivo_original}** (columna de email: "
        f"**{trabajo_actual.columna_email}**). Puedes recargar "
        "la página o volver más tarde con este mismo enlace (la dirección de esta "
        "pestaña): la verificación sigue en el servidor aunque se corte la conexión."
    )

    @st.fragment(run_every="2s")
    def _panel_progreso():
        completados = trabajo_actual.completados
        total = trabajo_actual.total
        if trabajo_actual.en_curso():
            st.text(f"Verificando... {completados} de {total} correos verificados (aproximado).")
            if total:
                st.progress(min(completados / total, 1.0))
        else:
            st.rerun()

    _panel_progreso()


# --------------------------------------------------------------------------
# Resultado final: se muestra apenas el trabajo tiene un resultado (o error).
# --------------------------------------------------------------------------
elif trabajo_actual is not None and trabajo_actual.error is not None:
    st.error(f"Ocurrió un error inesperado durante la verificación: {trabajo_actual.error}")

elif trabajo_actual is not None and trabajo_actual.resultado is not None:
    df_resultado, tiempo_total_segundos = trabajo_actual.resultado
    minutos, segundos = divmod(tiempo_total_segundos, 60)
    st.success(
        f"**{trabajo_actual.nombre_archivo_original}**: listo en "
        f"{int(minutos)} min {segundos:.0f} s."
    )

    particiones = particionar_por_accion(df_resultado)
    total = len(df_resultado)

    def _pct(cantidad):
        return f"{cantidad / total * 100:.1f}%" if total else "0%"

    col1, col2, col3 = st.columns(3)
    col1.metric("Total", total)
    col2.metric("MANTENER", len(particiones["buenos"]), _pct(len(particiones["buenos"])))
    col3.metric("ELIMINAR", len(particiones["eliminar"]), _pct(len(particiones["eliminar"])))

    st.subheader("Descargar resultados")
    nombres_archivo = {
        clave: nombre_archivo_salida(trabajo_actual.nombre_archivo_original, clave)
        for clave in ("buenos", "eliminar")
    }
    etiquetas_accion = {"buenos": "MANTENER", "eliminar": "ELIMINAR"}
    col_a, col_b = st.columns(2)
    for columna_ui, clave in zip((col_a, col_b), ("buenos", "eliminar")):
        columna_ui.download_button(
            f"⬇️ Descargar {nombres_archivo[clave]} ({etiquetas_accion[clave]})",
            data=trabajo_actual.excel(particiones, clave),
            file_name=nombres_archivo[clave],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
