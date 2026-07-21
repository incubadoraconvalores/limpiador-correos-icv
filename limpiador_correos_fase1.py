#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Limpiador de correos - Fase 1 (local, sin coste)
CON VALORES - Departamento de Innovación

Lee una lista de contactos (CSV o Excel), verifica cada correo con los
mismos criterios que usaba Bouncer (userbouncer.com) hasta donde es posible
replicar de forma gratuita y local, y genera TRES archivos de salida:

    - buenos.xlsx   -> filas con accion = "MANTENER" (correos confirmados o
                       mantenidos por confianza en proveedores masivos/
                       reputación de IP). Lista para enviar campañas.
    - revisar.xlsx  -> filas con accion = "REVISAR", para revisión manual.
    - eliminar.xlsx -> filas con accion = "ELIMINAR" (ya validado con 100%
                       de precisión contra Bouncer, no necesita revisión).

    Los 3 con todas las columnas originales + status/reason/accion, con
    AutoFiltro activado en la fila de encabezados y esa fila congelada
    (freeze panes) para filtrar sin pasos manuales adicionales.

Novedades de esta versión:
    - Estandarización automática de la entrada: si un .csv usa ";" en vez de
      "," como separador de columnas, o tiene un encoding que rompe tildes
      (ej. "PAÍS" mal codificado), se detecta y corrige automáticamente
      antes de procesar el archivo (ver estandarizar_entrada()). Ya no hace
      falta corregir archivos a mano antes de correr el script.
    - Si una celda de email contiene varias direcciones separadas por ";" o ",",
      se separan automáticamente en filas independientes, cada una evaluada
      por su cuenta (sintaxis, dominio, SMTP, desechables). Esto es distinto
      del separador de COLUMNAS del csv: uno se resuelve en estandarizar_entrada(),
      el otro en expandir_emails_multiples(), sobre el valor ya extraído de la celda.
    - Se conservan TODAS las columnas originales del archivo de entrada
      (organización, ubicación, teléfono, fuente, etc.) en el archivo de
      salida; solo se añaden al final las columnas status/reason/accion.
      Si una fila original tenía 2 correos, el resto de sus datos se repite
      en las 2 filas resultantes.
    - Reintentos automáticos en la verificación SMTP cuando el resultado no
      es concluyente (timeout, 4xx, conexión rechazada), y timeout SMTP más
      largo para proveedores de correo masivos (hotmail, outlook, live, aol,
      yahoo y sus variantes regionales), que suelen tardar más en responder
      por sus sistemas anti-abuso. Ver PROVEEDORES_MASIVOS más abajo.
    - Verificación SMTP en paralelo (hilos, no asyncio): varios correos se
      verifican en simultáneo (--concurrencia, por defecto
      CONCURRENCIA_DEFAULT), con un límite aparte de conexiones simultáneas
      contra el MISMO dominio (--concurrencia-por-dominio) para no abrir
      varias conexiones a la vez al mismo proveedor masivo. Al final se
      informa el tiempo total de verificación.
    - Interfaz web opcional (streamlit_app.py): subir un archivo, correr todo
      el pipeline y descargar los 3 documentos sin usar la terminal.
      Ver "streamlit run streamlit_app.py".

LÓGICA DE CLASIFICACIÓN (conservadora, sin heurísticas aproximadas):

    Sintaxis inválida               -> Undeliverable / invalid_email          -> ELIMINAR
    Correo desechable/temporal      -> Risky / low_quality                    -> ELIMINAR
    Dominio sin registros MX        -> Undeliverable / invalid_domain         -> ELIMINAR
    Error de DNS / timeout DNS      -> Unknown / dns_error                    -> REVISAR
    SMTP confirma el buzón (250), dominio NO catch-all -> Deliverable / accepted_email -> MANTENER
    SMTP confirma el buzón (250), pero el dominio
      es catch-all (acepta CUALQUIER dirección)   -> probable_deliverable / dominio_catch_all -> REVISAR
                                                      (el 250 no distingue nada, no confirma ESE buzón)
    No se pudo determinar si el dominio es
      catch-all (prueba con dirección inventada
      no concluyente tras reintentos)             -> Unknown / catchall_no_verificable -> REVISAR
    SMTP rechaza (5xx) el buzón, mensaje genuino  -> Undeliverable / rejected_email -> ELIMINAR
    SMTP rechaza (5xx) pero el mensaje indica
      bloqueo por reputación de NUESTRA IP
      (Spamhaus/PBL/RBL/blacklist, no el buzón)   -> probable_deliverable / rechazo_por_reputacion_ip_propia -> REVISAR
                                                      (no es rechazo genuino, pero tampoco confirmación real)
    SMTP no concluyente (dominio propio)          -> Unknown / unavailable_smtp|timeout -> REVISAR
    SMTP no concluyente (proveedor masivo,
      tras reintentos)                            -> probable_deliverable / smtp_bloqueado_proveedor_masivo -> REVISAR
                                                      (no hay confirmación SMTP real)
    (buzón lleno no se detecta en Fase 1: ver Fase 2)

NOTA IMPORTANTE (detectado en julio 2026, comparando contra Bouncer):
    Muchos servidores corporativos (Microsoft 365 / mail.protection.outlook.com,
    Mimecast, etc.) devuelven un 550 INMEDIATO a cualquier dirección -incluso
    inventadas- cuando la IP que ejecuta este script está listada en Spamhaus
    PBL/RBL. Ese 550 no dice nada sobre si el buzón existe: es un rechazo por
    reputación de la IP emisora, no del destinatario. Se detecta por el TEXTO
    del mensaje SMTP (ver PALABRAS_CLAVE_RECHAZO_POR_REPUTACION).

    ACTUALIZACIÓN (julio 2026): tanto este caso como el de proveedores
    masivos sin confirmar SE MANDAN A REVISAR, no a MANTENER. Al comparar
    contra Bouncer, solo ~41% de los casos con reason "rechazo_por_reputacion_ip_propia"
    o "smtp_bloqueado_proveedor_masivo" resultaron ser realmente entregables
    -- mandarlos directo a campaña arriesgaría rebotes en la mayoría de los
    casos. buenos.xlsx queda reservado para reason == "accepted_email"
    (confirmación SMTP 100% real).

    ACTUALIZACIÓN 2 (julio 2026): un 250 a RCPT TO tampoco es, por sí solo,
    prueba de que la dirección exista: si el dominio es "catch-all" (acepta
    con 250 CUALQUIER dirección, real o inventada -- muy común en tenants
    de Microsoft 365 mal configurados), el 250 no distingue nada. Se agregó
    DetectorCatchAll: antes de confiar en un accepted_email, se prueba UNA
    VEZ POR DOMINIO (no por email) una dirección claramente inventada. Si
    también da 250, el dominio es catch-all y el correo real va a REVISAR
    (reason "dominio_catch_all"), no a MANTENER. Detectado comparando contra
    Bouncer: en un lote de 63 accepted_email de un dominio catch-all, Bouncer
    solo confirmó 3 como realmente entregables (acceptAll=true en el resto).

Uso:
    python limpiador_correos_fase1.py entrada.xlsx
    python limpiador_correos_fase1.py entrada.csv --columna correo_electronico
    python limpiador_correos_fase1.py entrada.xlsx --salida ./resultados --smtp-timeout 8
    python limpiador_correos_fase1.py entrada.xlsx --concurrencia 15 --concurrencia-por-dominio 2

Requisitos (instalar una sola vez):
    pip install pandas openpyxl dnspython tqdm disposable-email-domains email-validator
"""

import argparse
import concurrent.futures
import csv
import io
import os
import re
import smtplib
import socket
import sys
import threading
import time
import uuid
from pathlib import Path

import dns.resolver
import dns.exception
import pandas as pd
from email_validator import validate_email, EmailNotValidError
from openpyxl.utils import get_column_letter
from tqdm import tqdm

try:
    from disposable_email_domains import blocklist as DISPOSABLE_BLOCKLIST
except ImportError:
    # Nombre alternativo usado en algunas versiones de la librería
    from disposable_email_domains import disposable_domains as DISPOSABLE_BLOCKLIST


# --------------------------------------------------------------------------
# CONFIGURACIÓN GENERAL
# --------------------------------------------------------------------------

DNS_TIMEOUT_DEFAULT = 5        # segundos para resolver registros MX
SMTP_TIMEOUT_DEFAULT = 8       # segundos para la conexión SMTP
SMTP_HELO_DOMINIO = "convalores.org"     # dominio que se anuncia en el HELO (ajustar)
SMTP_MAIL_FROM = "verificacion@convalores.org"  # remitente usado en la prueba (ajustar)

RUTA_LISTA_NEGRA_LOCAL = "lista_negra_local.txt"

# Códigos SMTP que consideramos rechazo permanente (buzón no existe)
CODIGOS_SMTP_RECHAZO_PERMANENTE = {550, 551, 553, 554}
# Código SMTP que consideramos aceptación explícita del buzón
CODIGO_SMTP_ACEPTADO = 250

# Palabras clave que, dentro del texto de un rechazo 5xx, indican que el
# servidor está bloqueando NUESTRA IP por reputación (Spamhaus, PBL, otras
# RBL/blacklists) y no está diciendo nada sobre si el buzón destino existe.
# Detectado comparando contra Bouncer en julio 2026: servidores de Microsoft
# 365 y Mimecast devuelven 550 a CUALQUIER dirección (incluso inventadas)
# cuando la IP emisora está listada, lo que generaba falsos ELIMINAR masivos
# en dominios corporativos legítimos (ONGs, bancos, etc.).
PALABRAS_CLAVE_RECHAZO_POR_REPUTACION = (
    "spamhaus", "pbl", "rbl", "blocklist", "block list",
    "blacklist", "black list", "denylist", "deny list",
    "listed by", "reputation", "reputación", "backscatter",
    "policy block list",
)

# Reintentos automáticos cuando la verificación SMTP no da una respuesta
# concluyente (timeout, 4xx, conexión rechazada), antes de darla por perdida.
SMTP_REINTENTOS = 3
SMTP_PAUSA_ENTRE_REINTENTOS = 2.0  # segundos (dominios NO masivos, sin cambios)

# Proveedores de correo masivos: aplican sistemas anti-abuso agresivos que
# hacen que la verificación SMTP tarde más o no responda de forma concluyente
# aunque el buzón exista. Se identifican por la primera etiqueta del dominio,
# lo que cubre automáticamente variantes regionales (hotmail.es, hotmail.co.uk,
# yahoo.com.mx, yahoo.co.jp, outlook.de, etc.) sin tener que listarlas todas.
PROVEEDORES_MASIVOS = {"hotmail", "outlook", "live", "msn", "aol", "yahoo"}
SMTP_TIMEOUT_PROVEEDOR_MASIVO_DEFAULT = 20  # segundos

# Para proveedores masivos, en vez de una pausa fija entre reintentos se usa
# una pausa progresiva más larga (espaciar más las conexiones reduce la
# probabilidad de que el proveedor las detecte como abuso). Un valor por
# posición: pausas_progresivas[0] = espera antes del 2do intento,
# pausas_progresivas[1] = espera antes del 3er intento, etc.
SMTP_PAUSAS_PROGRESIVAS_PROVEEDOR_MASIVO = [5.0, 15.0]  # segundos

# Factor multiplicador aplicado a las pausas de proveedores masivos cuando
# se activa --verificacion-paciente (prioriza tasa de acierto sobre velocidad).
FACTOR_VERIFICACION_PACIENTE = 2.0

# Verificación SMTP en paralelo (varios correos a la vez, usando hilos: cada
# verificación es I/O de red bloqueante, así que el GIL se libera durante la
# espera y varios hilos avanzan en simultáneo sin necesidad de asyncio).
CONCURRENCIA_DEFAULT = 12          # verificaciones SMTP simultáneas, en total
CONCURRENCIA_POR_DOMINIO_DEFAULT = 2  # simultáneas contra el MISMO dominio


def es_proveedor_masivo(dominio: str) -> bool:
    etiqueta = dominio.lower().split(".")[0]
    return etiqueta in PROVEEDORES_MASIVOS


class LimitadorConcurrenciaPorDominio:
    """
    Limita cuántas verificaciones SMTP corren en simultáneo contra el MISMO
    dominio, sin importar cuántos hilos haya libres en el pool global
    (--concurrencia). Sin este límite, varios hilos podrían abrir conexiones
    simultáneas al mismo proveedor masivo (ej. 5 de 12 hilos contra
    hotmail.com a la vez), lo que agravaría el riesgo de bloqueo por abuso
    que las pausas progresivas ya intentan evitar. Con este límite, la
    concurrencia se reparte entre dominios DISTINTOS, que es de donde sale
    la mayor parte de la ganancia de velocidad en listas con muchos dominios.
    """

    def __init__(self, limite_por_dominio: int):
        self._limite = limite_por_dominio
        self._lock_diccionario = threading.Lock()
        self._semaforos_por_dominio = {}

    def _semaforo_de(self, dominio: str) -> threading.Semaphore:
        with self._lock_diccionario:
            semaforo = self._semaforos_por_dominio.get(dominio)
            if semaforo is None:
                semaforo = threading.Semaphore(self._limite)
                self._semaforos_por_dominio[dominio] = semaforo
            return semaforo

    def adquirir(self, dominio: str):
        self._semaforo_de(dominio).acquire()

    def liberar(self, dominio: str):
        self._semaforo_de(dominio).release()


class DetectorCatchAll:
    """
    Para cada dominio, verifica UNA SOLA VEZ (memoizado) si el servidor
    acepta con 250 una dirección claramente inventada. Si la acepta, el
    dominio es "catch-all": acepta CUALQUIER dirección sin verificar el
    buzón real, por lo que un 250 a una dirección real de ese dominio NO es
    una confirmación genuina de que ESA dirección específica exista.

    Detectado comparando contra Bouncer (julio 2026): de 63 correos con
    accepted_email en dominios M365 configurados como catch-all, Bouncer
    solo confirmó 3 como realmente entregables (el resto, acceptAll=true) --
    sin esta detección, los 63 hubiesen ido a MANTENER como falsos positivos.

    Se prueba una sola vez POR DOMINIO, no una vez por email, para no
    duplicar el tráfico SMTP: si 50 contactos comparten un dominio, se hace
    1 verificación de catch-all + 50 verificaciones reales, no 100.
    Thread-safe: si varios hilos llegan al mismo dominio antes de que la
    primera prueba termine, esperan su resultado en vez de probar cada uno
    por su cuenta.
    """

    def __init__(self):
        self._lock = threading.Lock()
        self._resultados = {}  # dominio -> True / False / None (inconcluyente)
        self._eventos = {}     # dominio -> threading.Event, mientras se prueba

    def es_catch_all(self, dominio: str, mx_host: str, timeout: float, pausas,
                      limitador_dominio: "LimitadorConcurrenciaPorDominio" = None):
        """
        Devuelve True (catch-all confirmado: el 250 real NO es de fiar),
        False (no catch-all: el 250 real SÍ confirma el buzón), o None
        (no se pudo determinar tras los reintentos -- inconcluyente, el
        llamador debe tratarlo como no confirmado).
        """
        with self._lock:
            if dominio in self._resultados:
                return self._resultados[dominio]
            evento = self._eventos.get(dominio)
            me_toca_probar = evento is None
            if me_toca_probar:
                evento = threading.Event()
                self._eventos[dominio] = evento

        if not me_toca_probar:
            evento.wait()
            with self._lock:
                return self._resultados.get(dominio)

        valor = None
        try:
            email_inventado = f"noexiste-verificacion-catchall-{uuid.uuid4().hex}@{dominio}"
            if limitador_dominio is not None:
                limitador_dominio.adquirir(dominio)
            try:
                resultado_smtp, _ = verificar_smtp_con_reintentos(
                    email_inventado, mx_host, timeout,
                    intentos=SMTP_REINTENTOS, pausas=pausas,
                )
            finally:
                if limitador_dominio is not None:
                    limitador_dominio.liberar(dominio)

            if resultado_smtp == "aceptado":
                valor = True
            elif resultado_smtp in ("rechazado", "rechazado_reputacion"):
                valor = False
            # cualquier otro caso (no_concluyente) deja valor = None
        finally:
            with self._lock:
                self._resultados[dominio] = valor
            evento.set()

        return valor


def es_rechazo_por_reputacion_propia(mensaje_smtp) -> bool:
    """
    Analiza el texto de un rechazo 5xx para distinguir un bloqueo por
    reputación de NUESTRA IP (Spamhaus/PBL/RBL/blacklist) de un rechazo
    genuino de buzón inexistente. Ver PALABRAS_CLAVE_RECHAZO_POR_REPUTACION.
    """
    if mensaje_smtp is None:
        return False
    texto = mensaje_smtp.decode("utf-8", errors="ignore") if isinstance(mensaje_smtp, bytes) else str(mensaje_smtp)
    texto = texto.lower()
    return any(palabra in texto for palabra in PALABRAS_CLAVE_RECHAZO_POR_REPUTACION)


# --------------------------------------------------------------------------
# LISTA NEGRA LOCAL DE DESECHABLES (editable a mano por Innovación)
# --------------------------------------------------------------------------

def cargar_o_crear_lista_negra_local(ruta: str) -> set:
    """
    Carga los dominios desechables definidos a mano por el equipo.
    Si el archivo no existe, lo crea con un ejemplo y comentarios de ayuda,
    para que cualquier persona no técnica pueda añadir dominios sin tocar
    el código.
    """
    path = Path(ruta)

    if not path.exists():
        contenido_inicial = (
            "# Lista negra local de dominios desechables/temporales\n"
            "# ------------------------------------------------------\n"
            "# Añade aquí (uno por línea) dominios que detectéis como\n"
            "# desechables y que la librería 'disposable-email-domains'\n"
            "# todavía no incluya. No hace falta tocar el código.\n"
            "#\n"
            "# Reglas:\n"
            "#   - Una línea = un dominio (sin @, solo la parte de after '@')\n"
            "#   - Las líneas que empiezan por '#' se ignoran (comentarios)\n"
            "#   - No hace falta orden alfabético ni mayúsculas/minúsculas\n"
            "#\n"
            "# Ejemplo:\n"
            "# mailinator-clone.com\n"
            "# tempcorreo24.net\n"
        )
        path.write_text(contenido_inicial, encoding="utf-8")
        print(f"[INFO] No existía '{ruta}'. Se ha creado con contenido de ejemplo.")

    dominios = set()
    with path.open(encoding="utf-8") as f:
        for linea in f:
            linea = linea.strip().lower()
            if not linea or linea.startswith("#"):
                continue
            dominios.add(linea)

    return dominios


def es_desechable(dominio: str, lista_negra_local: set) -> bool:
    dominio = dominio.lower()
    if dominio in lista_negra_local:
        return True
    if dominio in DISPOSABLE_BLOCKLIST:
        return True
    return False


# --------------------------------------------------------------------------
# VALIDACIÓN DE SINTAXIS
# --------------------------------------------------------------------------

def validar_sintaxis(email: str):
    """
    Devuelve (True, email_normalizado, dominio) si la sintaxis es válida,
    o (False, None, None) si no lo es.
    No se comprueba entregabilidad aquí (eso lo hacen DNS/SMTP más adelante).
    """
    try:
        resultado = validate_email(email, check_deliverability=False)
        email_normalizado = resultado.normalized
        dominio = email_normalizado.split("@")[1]
        return True, email_normalizado, dominio
    except EmailNotValidError:
        return False, None, None


# --------------------------------------------------------------------------
# VERIFICACIÓN DNS (REGISTROS MX)
# --------------------------------------------------------------------------

def verificar_mx(dominio: str, timeout: float):
    """
    Devuelve una tupla (resultado, mx_host, reason) donde resultado es uno de:
        "ok"         -> hay registros MX, mx_host es el de mayor prioridad
        "sin_mx"     -> el dominio no tiene registros MX (Undeliverable)
        "dns_error"  -> no se pudo resolver por timeout/error de red (Unknown)
    """
    try:
        resolver = dns.resolver.Resolver()
        resolver.timeout = timeout
        resolver.lifetime = timeout
        respuestas = resolver.resolve(dominio, "MX")
        registros = sorted(respuestas, key=lambda r: r.preference)
        mx_host = str(registros[0].exchange).rstrip(".")
        return "ok", mx_host, None

    except (dns.resolver.NXDOMAIN, dns.resolver.NoAnswer):
        return "sin_mx", None, "invalid_domain"

    except (dns.exception.Timeout, dns.resolver.LifetimeTimeout):
        return "dns_error", None, "dns_error"

    except Exception:
        # Cualquier otro fallo de resolución se trata como error transitorio,
        # no como "el dominio no existe" (para evitar falsos positivos).
        return "dns_error", None, "dns_error"


# --------------------------------------------------------------------------
# VERIFICACIÓN SMTP
# --------------------------------------------------------------------------

def verificar_smtp(email: str, mx_host: str, timeout: float):
    """
    Intenta una verificación SMTP básica (HELO / MAIL FROM / RCPT TO).

    Devuelve (resultado, reason) donde resultado es uno de:
        "aceptado"            -> el servidor confirmó el buzón (250)
        "rechazado"           -> el servidor rechazó el buzón de forma permanente
                                  (5xx) con un mensaje genuino de "no existe"
        "rechazado_reputacion" -> 5xx pero el mensaje indica que el bloqueo es
                                  por reputación de NUESTRA IP (Spamhaus/PBL/RBL),
                                  no una confirmación real de que el buzón no exista
        "no_concluyente"      -> timeout, conexión bloqueada, código 4xx, etc.

    IMPORTANTE (alcance de Fase 1):
        No se intenta distinguir catch-all ni buzón lleno; cualquier
        respuesta que no sea un 250 claro o un 5xx claro de "no existe"
        se trata como "no_concluyente" -> Unknown, siguiendo el criterio
        conservador acordado (sin heurísticas aproximadas).
    """
    smtp = None
    try:
        smtp = smtplib.SMTP(timeout=timeout)
        smtp.connect(mx_host, 25)
        smtp.helo(SMTP_HELO_DOMINIO)
        smtp.mail(SMTP_MAIL_FROM)
        codigo, mensaje = smtp.rcpt(email)

        if codigo == CODIGO_SMTP_ACEPTADO:
            return "aceptado", "accepted_email"
        elif codigo in CODIGOS_SMTP_RECHAZO_PERMANENTE:
            if es_rechazo_por_reputacion_propia(mensaje):
                return "rechazado_reputacion", "rechazo_por_reputacion_ip_propia"
            return "rechazado", "rejected_email"
        else:
            # Códigos 4xx (temporales) u otros no contemplados: no concluyente
            return "no_concluyente", "unavailable_smtp"

    except (socket.timeout, TimeoutError):
        return "no_concluyente", "timeout"

    except (smtplib.SMTPServerDisconnected, smtplib.SMTPConnectError,
            ConnectionRefusedError, OSError):
        # Incluye el caso típico de proveedores domésticos que bloquean
        # el puerto 25 de salida (ver limitación en el plan técnico).
        return "no_concluyente", "unavailable_smtp"

    except Exception:
        return "no_concluyente", "unknown"

    finally:
        if smtp is not None:
            try:
                smtp.quit()
            except Exception:
                pass


def verificar_smtp_con_reintentos(email: str, mx_host: str, timeout: float,
                                   intentos: int = SMTP_REINTENTOS,
                                   pausas=SMTP_PAUSA_ENTRE_REINTENTOS):
    """
    Repite verificar_smtp() hasta 'intentos' veces mientras el resultado sea
    "no_concluyente" (timeout, 4xx, conexión rechazada). Se detiene de
    inmediato en cuanto obtiene una respuesta definitiva ("aceptado",
    "rechazado" o "rechazado_reputacion"), sin gastar reintentos ni esperar
    pausas de más (reintentar un bloqueo de reputación de nuestra IP no
    cambiaría el resultado dentro de la misma ejecución).

    'pausas' admite:
        - un único número: esa misma pausa fija antes de cada reintento
          (comportamiento para dominios no masivos).
        - una lista de 'intentos - 1' números: pausa progresiva distinta
          antes de cada reintento (usado para proveedores masivos, donde
          conviene espaciar cada vez más las conexiones).
    """
    if isinstance(pausas, (int, float)):
        pausas = [pausas] * (intentos - 1)

    resultado, reason = "no_concluyente", "unavailable_smtp"
    for intento in range(1, intentos + 1):
        resultado, reason = verificar_smtp(email, mx_host, timeout)
        if resultado != "no_concluyente":
            return resultado, reason
        if intento < intentos:
            time.sleep(pausas[intento - 1])
    return resultado, reason


# --------------------------------------------------------------------------
# CLASIFICACIÓN COMPLETA DE UN CORREO
# --------------------------------------------------------------------------

def clasificar_email(email_original: str, lista_negra_local: set,
                      dns_timeout: float, smtp_timeout: float,
                      smtp_timeout_proveedor_masivo: float,
                      verificar_smtp_activo: bool,
                      verificacion_paciente: bool = False,
                      limitador_dominio: "LimitadorConcurrenciaPorDominio" = None,
                      detector_catch_all: "DetectorCatchAll" = None):
    """
    Aplica el pipeline completo y devuelve un dict con:
        email, status, reason, accion
    """
    email_original = str(email_original).strip()

    # 1) Sintaxis
    valido, email_normalizado, dominio = validar_sintaxis(email_original)
    if not valido:
        return {
            "email": email_original,
            "status": "undeliverable",
            "reason": "invalid_email",
            "accion": "ELIMINAR",
        }

    # 2) Desechable / temporal (chequeo barato, antes de tocar red)
    if es_desechable(dominio, lista_negra_local):
        return {
            "email": email_normalizado,
            "status": "risky",
            "reason": "low_quality",
            "accion": "ELIMINAR",
        }

    # 3) DNS / registros MX
    resultado_mx, mx_host, reason_mx = verificar_mx(dominio, dns_timeout)

    if resultado_mx == "sin_mx":
        return {
            "email": email_normalizado,
            "status": "undeliverable",
            "reason": "invalid_domain",
            "accion": "ELIMINAR",
        }

    if resultado_mx == "dns_error":
        return {
            "email": email_normalizado,
            "status": "unknown",
            "reason": "dns_error",
            "accion": "REVISAR",
        }

    # 4) SMTP (opcional, puede desactivarse si la red bloquea el puerto 25)
    if not verificar_smtp_activo:
        return {
            "email": email_normalizado,
            "status": "unknown",
            "reason": "unsupported",
            "accion": "REVISAR",
        }

    proveedor_masivo = es_proveedor_masivo(dominio)
    timeout_smtp_real = smtp_timeout_proveedor_masivo if proveedor_masivo else smtp_timeout
    factor_paciencia = FACTOR_VERIFICACION_PACIENTE if verificacion_paciente else 1.0

    if proveedor_masivo:
        # Pausa progresiva (más larga en cada reintento) para espaciar las
        # conexiones y reducir la probabilidad de bloqueo por abuso.
        pausas_smtp = [p * factor_paciencia for p in SMTP_PAUSAS_PROGRESIVAS_PROVEEDOR_MASIVO]
    else:
        # Dominios no masivos: pausa fija, sin cambios respecto a antes.
        pausas_smtp = SMTP_PAUSA_ENTRE_REINTENTOS

    # El límite por dominio se adquiere recién acá (no antes de resolver MX),
    # para no bloquear otros hilos mientras se hace DNS de un dominio distinto.
    if limitador_dominio is not None:
        limitador_dominio.adquirir(dominio)
    try:
        resultado_smtp, reason_smtp = verificar_smtp_con_reintentos(
            email_normalizado, mx_host, timeout_smtp_real,
            intentos=SMTP_REINTENTOS, pausas=pausas_smtp,
        )
    finally:
        if limitador_dominio is not None:
            limitador_dominio.liberar(dominio)

    if resultado_smtp == "aceptado":
        # Un 250 a la dirección real no es, por sí solo, prueba de que ESA
        # dirección específica exista: si el dominio es "catch-all" (acepta
        # con 250 CUALQUIER dirección, real o inventada), el 250 no confirma
        # nada. Se prueba una vez por dominio (memoizado en detector_catch_all,
        # no se repite por cada email) antes de confiar en el accepted_email.
        if detector_catch_all is not None:
            es_catch_all = detector_catch_all.es_catch_all(
                dominio, mx_host, timeout_smtp_real, pausas_smtp, limitador_dominio,
            )
            if es_catch_all is None:
                return {
                    "email": email_normalizado,
                    "status": "unknown",
                    "reason": "catchall_no_verificable",
                    "accion": "REVISAR",
                }
            if es_catch_all:
                # Detectado comparando contra Bouncer (julio 2026): en
                # dominios catch-all, solo ~5% de los accepted_email
                # resultaron ser realmente entregables (3 de 63). El 250 de
                # la dirección real no distingue nada de un 250 a una
                # dirección inventada, así que no se puede confiar en él.
                return {
                    "email": email_normalizado,
                    "status": "probable_deliverable",
                    "reason": "dominio_catch_all",
                    "accion": "REVISAR",
                }
        return {
            "email": email_normalizado,
            "status": "deliverable",
            "reason": reason_smtp,
            "accion": "MANTENER",
        }

    if resultado_smtp == "rechazado":
        return {
            "email": email_normalizado,
            "status": "undeliverable",
            "reason": reason_smtp,
            "accion": "ELIMINAR",
        }

    if resultado_smtp == "rechazado_reputacion":
        # El 5xx no es una confirmación real de "buzón inexistente": el
        # mensaje indica que el rechazo es por reputación de NUESTRA IP
        # (Spamhaus/PBL/RBL). No es tampoco una confirmación de que el buzón
        # SÍ exista: comparado contra Bouncer, solo ~41% de estos casos
        # resultan ser realmente entregables. Por eso va a REVISAR (no
        # MANTENER): mandarlo directo a campaña arriesgaría rebotes en la
        # mayoría de los casos.
        return {
            "email": email_normalizado,
            "status": "probable_deliverable",
            "reason": reason_smtp,  # "rechazo_por_reputacion_ip_propia"
            "accion": "REVISAR",
        }

    # no_concluyente -> Unknown (incluye catch-all y buzón lleno no detectables
    # en esta fase, además de timeouts y bloqueos de puerto 25).
    #
    # Excepción: si es un proveedor masivo conocido y ya se agotaron los
    # reintentos, se etiqueta con la razón específica "smtp_bloqueado_proveedor_masivo"
    # para dejar visible que no hubo una confirmación SMTP 100% real, pero
    # igual va a REVISAR (no MANTENER): comparado contra Bouncer, solo ~41%
    # de estos casos resultan ser realmente entregables.
    if proveedor_masivo:
        return {
            "email": email_normalizado,
            "status": "probable_deliverable",
            "reason": "smtp_bloqueado_proveedor_masivo",
            "accion": "REVISAR",
        }

    return {
        "email": email_normalizado,
        "status": "unknown",
        "reason": reason_smtp,
        "accion": "REVISAR",
    }


# --------------------------------------------------------------------------
# LECTURA / ESCRITURA DE ARCHIVOS
# --------------------------------------------------------------------------

SEPARADOR_MULTIPLES_EMAILS = re.compile(r"[;,]")

# Palabras que, si aparecen en el nombre de una columna candidata a email,
# suelen indicar que en realidad es otra cosa (nombre de empresa, estado,
# origen del dato, un ID, etc.) y no la columna de correos en sí. Ej: en un
# export de Apollo, "Company Name for Emails" contiene "email" pero es el
# nombre de la empresa, no una dirección de correo.
PALABRAS_EXCLUSION_COLUMNA_EMAIL = ("name", "nombre", "company", "empresa", "status", "source", "id")


def _normalizar_nombre_columna(nombre) -> str:
    return re.sub(r"[\s_-]+", "", str(nombre).strip().lower())


def detectar_columna_email(columnas):
    """
    Detecta automáticamente la columna de email entre 'columnas', con esta
    prioridad:
        1) Coincidencia EXACTA con "email" o "correo" (prioridad absoluta).
        2) El nombre EMPIEZA con "email" o "correo" (ej. "Email Address").
        3) El nombre CONTIENE "email" o "correo" en cualquier parte
           (comportamiento anterior, usado solo si 1 y 2 no dieron nada).
    En los niveles 2 y 3 se descartan columnas cuyo nombre también contenga
    palabras como "name"/"company"/"status"/"id" (ver
    PALABRAS_EXCLUSION_COLUMNA_EMAIL), salvo que sea la única opción
    disponible en todo el archivo (en ese caso, se usa igual como último
    recurso en vez de fallar).
    Devuelve el nombre de columna elegido, o None si ninguna coincide.
    """
    normalizados = [(col, _normalizar_nombre_columna(col)) for col in columnas]

    # Nivel 1: coincidencia exacta, prioridad absoluta.
    exactas = [col for col, norm in normalizados if norm in ("email", "correo")]
    if exactas:
        return exactas[0]

    def contiene_palabra_exclusion(norm: str) -> bool:
        return any(palabra in norm for palabra in PALABRAS_EXCLUSION_COLUMNA_EMAIL)

    # Nivel 2: empieza con "email"/"correo". Nivel 3: lo contiene en cualquier parte.
    empiezan = [(col, norm) for col, norm in normalizados if norm.startswith("email") or norm.startswith("correo")]
    contienen = [(col, norm) for col, norm in normalizados if "email" in norm or "correo" in norm]

    # Primero se intenta sin columnas "sospechosas" (name/company/status/...),
    # respetando la prioridad "empieza" > "contiene".
    for nivel in (empiezan, contienen):
        sin_exclusion = [col for col, norm in nivel if not contiene_palabra_exclusion(norm)]
        if sin_exclusion:
            return sin_exclusion[0]

    # Si no hubo ninguna columna "limpia", se acepta como último recurso la
    # primera candidata aunque contenga una palabra de exclusión.
    for nivel in (empiezan, contienen):
        if nivel:
            return nivel[0][0]

    return None


def expandir_emails_multiples(df: pd.DataFrame, columna_email: str) -> pd.DataFrame:
    """
    Divide las celdas de la columna de email que contengan varias direcciones
    separadas por ';' o ',' en filas independientes, duplicando el resto de
    columnas originales (organización, ubicación, teléfono, etc.) en cada una.

    Ejemplo:
        "juan@x.org; ana@x.org"  -> dos filas, una con "juan@x.org" y otra
                                     con "ana@x.org", ambas con el resto de
                                     datos de la fila original repetidos.
    """
    df = df.copy()

    # Partir cada celda en una lista de correos (quitando espacios sobrantes
    # y descartando trozos vacíos que puedan quedar por separadores repetidos
    # o al final de la celda, ej: "juan@x.org;" o "juan@x.org;;ana@x.org").
    def partir_celda(valor):
        if pd.isna(valor):
            return [valor]
        trozos = SEPARADOR_MULTIPLES_EMAILS.split(str(valor))
        trozos = [t.strip() for t in trozos if t.strip() != ""]
        return trozos if trozos else [valor]

    df[columna_email] = df[columna_email].apply(partir_celda)
    df = df.explode(columna_email, ignore_index=True)

    return df


# --------------------------------------------------------------------------
# ESTANDARIZADOR DE ENTRADA (separador de columnas + encoding)
# --------------------------------------------------------------------------
# Antes esto se hacía a mano: si un .csv venía con ";" en vez de "," como
# separador de columnas, o con tildes rotas (encoding incorrecto, ej. "PAÍS"
# convertido en mojibake), alguien generaba una copia "(corregido).csv" con
# un script aparte antes de correr el limpiador. Ahora se detecta y corrige
# automáticamente acá, en memoria, tanto para la CLI como para Streamlit.
#
# OJO: esto NO tiene nada que ver con expandir_emails_multiples() (más abajo),
# que separa varios correos dentro de UNA CELDA de email (ej. "a@x.com;b@x.com").
# Son dos separadores ";"/"," distintos en dos lugares distintos: uno es el
# separador de COLUMNAS del archivo completo, el otro es el separador DENTRO
# de una celda ya extraída. No se pisan entre sí.

ENCODINGS_A_PROBAR = ("utf-8-sig", "cp1252", "latin-1")  # latin-1 nunca falla


def _detectar_extension(fuente, nombre_archivo: str = None) -> str:
    """
    Devuelve la extensión (en minúsculas, con punto, ej. ".csv") de 'fuente'.
    'fuente' puede ser una ruta (str/Path) o un objeto tipo archivo (bytes,
    BytesIO, o el UploadedFile de st.file_uploader, que ya expone '.name').
    """
    if nombre_archivo:
        return Path(nombre_archivo).suffix.lower()
    if isinstance(fuente, (str, Path)):
        return Path(fuente).suffix.lower()
    nombre = getattr(fuente, "name", None)
    if nombre:
        return Path(nombre).suffix.lower()
    raise ValueError(
        "No se pudo determinar la extensión del archivo. "
        "Pasá 'nombre_archivo' si 'fuente' no es una ruta ni tiene atributo '.name'."
    )


def _leer_bytes_de_fuente(fuente) -> bytes:
    """Devuelve los bytes crudos de 'fuente', sea ruta, bytes o file-like."""
    if isinstance(fuente, bytes):
        return fuente
    if isinstance(fuente, (str, Path)):
        return Path(fuente).read_bytes()
    # Objeto tipo archivo (BytesIO, UploadedFile de Streamlit, etc.)
    if hasattr(fuente, "seek"):
        fuente.seek(0)
    datos = fuente.read()
    if hasattr(fuente, "seek"):
        fuente.seek(0)  # lo deja reutilizable para quien llame después
    return datos


def _detectar_encoding_y_texto(datos: bytes):
    """
    Prueba encodings en cascada y devuelve (encoding_usado, texto_decodificado).
    utf-8-sig decodifica bien tanto UTF-8 con BOM como sin BOM. cp1252 cubre
    el caso típico de tildes rotas en exports de Windows/Excel en español.
    latin-1 es el último recurso: nunca lanza UnicodeDecodeError.
    """
    for encoding in ENCODINGS_A_PROBAR:
        try:
            return encoding, datos.decode(encoding)
        except UnicodeDecodeError:
            continue
    # No debería llegar acá porque latin-1 no falla, pero por si acaso:
    return "latin-1", datos.decode("latin-1", errors="replace")


def _detectar_separador_csv(texto: str) -> str:
    """
    Detecta el separador de COLUMNAS de un csv (no el de multiples emails
    dentro de una celda) usando csv.Sniffer sobre una muestra de las
    primeras líneas, con fallback a contar ';' vs ',' en la primera línea
    si el Sniffer no puede decidir.
    """
    lineas = texto.splitlines()
    muestra = "\n".join(lineas[:20]) if lineas else texto
    try:
        return csv.Sniffer().sniff(muestra, delimiters=";,").delimiter
    except csv.Error:
        primera_linea = lineas[0] if lineas else ""
        return ";" if primera_linea.count(";") > primera_linea.count(",") else ","


def estandarizar_entrada(fuente, nombre_archivo: str = None) -> pd.DataFrame:
    """
    Lee 'fuente' (ruta en disco, o bytes/objeto tipo archivo como el
    UploadedFile de Streamlit) y devuelve un DataFrame ya estandarizado:
    para .csv, detecta y corrige separador de columnas y encoding
    automáticamente; para .xlsx/.xls, delega directo en pd.read_excel
    (formato binario, sin separador ni encoding que resolver).
    """
    extension = _detectar_extension(fuente, nombre_archivo)

    if extension in (".xlsx", ".xls"):
        return pd.read_excel(fuente)

    if extension != ".csv":
        raise ValueError(f"Formato no soportado: {extension}. Usa .csv, .xlsx o .xls")

    datos = _leer_bytes_de_fuente(fuente)
    encoding_usado, texto = _detectar_encoding_y_texto(datos)
    separador_usado = _detectar_separador_csv(texto)

    return pd.read_csv(io.StringIO(texto), sep=separador_usado)


def leer_lista_contactos(ruta_entrada, columna_email: str = None, nombre_archivo: str = None):
    df = estandarizar_entrada(ruta_entrada, nombre_archivo)

    if columna_email is None:
        columna_email = detectar_columna_email(df.columns)
        if columna_email is None:
            raise ValueError(
                "No se pudo detectar automáticamente la columna de email. "
                "Especifícala con --columna NOMBRE_COLUMNA. "
                f"Columnas disponibles: {list(df.columns)}"
            )
        print(f"[INFO] Columna de email detectada automáticamente: '{columna_email}'")

    if columna_email not in df.columns:
        raise ValueError(
            f"La columna '{columna_email}' no existe en el archivo. "
            f"Columnas disponibles: {list(df.columns)}"
        )

    filas_antes = len(df)

    # Quitar filas sin ningún valor en la columna de email
    df = df.dropna(subset=[columna_email]).copy()
    df = df[df[columna_email].astype(str).str.strip() != ""]
    df = df.reset_index(drop=True)

    # ID único por contacto ORIGINAL (antes de separar celdas con varios
    # correos), para poder reagrupar más adelante todos los correos que
    # vinieron de una misma fila/contacto de origen.
    df.insert(0, "id_contacto_original", range(1, len(df) + 1))

    # Separar celdas con múltiples correos en filas independientes,
    # conservando el resto de columnas originales (incluido el ID) tal cual
    df = expandir_emails_multiples(df, columna_email)
    df = df.dropna(subset=[columna_email]).copy()
    df = df[df[columna_email].astype(str).str.strip() != ""]
    df = df.reset_index(drop=True)

    filas_despues = len(df)
    if filas_despues != filas_antes:
        print(f"[INFO] {filas_antes} fila(s) originales -> {filas_despues} correo(s) individuales "
              f"tras separar celdas con múltiples direcciones.")

    return df, columna_email


def _guardar_hoja_excel(df: pd.DataFrame, destino, sheet_name: str = "Resultado"):
    """
    Guarda 'df' en 'destino' con AutoFiltro activado en la fila de encabezados
    (para poder filtrar por "accion", "reason", etc. sin pasos manuales) y
    esa fila congelada (freeze panes) para que se mantenga visible al hacer
    scroll. 'destino' puede ser una ruta en disco (Path/str) o un buffer en
    memoria (io.BytesIO) -- pd.ExcelWriter acepta ambos por igual.
    """
    with pd.ExcelWriter(destino, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]

        n_filas = len(df) + 1  # +1 por la fila de encabezados
        n_columnas = len(df.columns)
        ultima_columna = get_column_letter(n_columnas)

        worksheet.auto_filter.ref = f"A1:{ultima_columna}{n_filas}"
        worksheet.freeze_panes = "A2"


def generar_excel_en_memoria(df: pd.DataFrame, sheet_name: str = "Resultado") -> bytes:
    """
    Igual que _guardar_hoja_excel, pero devuelve los bytes del .xlsx en vez
    de escribir a disco (para los botones de descarga de Streamlit).
    """
    buffer = io.BytesIO()
    _guardar_hoja_excel(df, buffer, sheet_name)
    buffer.seek(0)
    return buffer.getvalue()


def particionar_por_accion(df_resultado: pd.DataFrame) -> dict:
    """
    Devuelve {'buenos': df, 'revisar': df, 'eliminar': df} filtrando por
    accion == "MANTENER"/"REVISAR"/"ELIMINAR" respectivamente. No cambia
    ningún criterio de clasificación, solo particiona el resultado ya
    calculado por clasificar_email().
    """
    return {
        "buenos": df_resultado[df_resultado["accion"] == "MANTENER"].reset_index(drop=True),
        "revisar": df_resultado[df_resultado["accion"] == "REVISAR"].reset_index(drop=True),
        "eliminar": df_resultado[df_resultado["accion"] == "ELIMINAR"].reset_index(drop=True),
    }


def guardar_resultados(df_resultado: pd.DataFrame, carpeta_salida: str):
    """
    Guarda tres archivos:
        - buenos.xlsx   -> accion == "MANTENER" (listo para campañas)
        - revisar.xlsx  -> accion == "REVISAR" (revisión manual)
        - eliminar.xlsx -> accion == "ELIMINAR" (descarte, ya validado)
    Los 3 con todas las columnas originales + status/reason/accion, AutoFiltro
    en los encabezados y esa fila congelada (ver _guardar_hoja_excel).
    """
    carpeta = Path(carpeta_salida)
    carpeta.mkdir(parents=True, exist_ok=True)

    particiones = particionar_por_accion(df_resultado)
    rutas = {nombre: carpeta / f"{nombre}.xlsx" for nombre in particiones}

    for nombre, df_particion in particiones.items():
        _guardar_hoja_excel(df_particion, rutas[nombre])

    return rutas


def clasificar_dataframe(df_entrada: pd.DataFrame, columna_email: str, lista_negra_local: set,
                          dns_timeout: float, smtp_timeout: float, smtp_timeout_proveedor_masivo: float,
                          verificar_smtp_activo: bool, verificacion_paciente: bool,
                          concurrencia: int, concurrencia_por_dominio: int,
                          mostrar_barra_progreso: bool = True, callback_progreso=None):
    """
    Clasifica todos los emails de 'df_entrada' en paralelo (hilos, con
    LimitadorConcurrenciaPorDominio para no saturar un mismo dominio) y
    devuelve (df_resultado, tiempo_total_segundos). df_resultado conserva
    todas las columnas originales + email normalizado + status/reason/accion.
    Reutilizada tanto por main() (CLI) como por streamlit_app.py.

    'callback_progreso', si se pasa, se llama como callback_progreso(completados,
    total) cada vez que un future termina -- es un hook de observación pasivo
    (no cambia el ThreadPoolExecutor, el limitador por dominio, ni ningún
    criterio de clasificación), pensado para que un frontend externo (ej.
    streamlit_app.py) pueda mostrar un contador de avance sin que los hilos
    de verificación toquen directamente ninguna UI.
    """
    limitador_dominio = LimitadorConcurrenciaPorDominio(concurrencia_por_dominio)
    detector_catch_all = DetectorCatchAll()
    emails = list(df_entrada[columna_email])
    filas_clasificacion = [None] * len(emails)

    def _clasificar_indexado(indice, email):
        fila = clasificar_email(
            email,
            lista_negra_local=lista_negra_local,
            dns_timeout=dns_timeout,
            smtp_timeout=smtp_timeout,
            smtp_timeout_proveedor_masivo=smtp_timeout_proveedor_masivo,
            verificar_smtp_activo=verificar_smtp_activo,
            verificacion_paciente=verificacion_paciente,
            limitador_dominio=limitador_dominio,
            detector_catch_all=detector_catch_all,
        )
        return indice, fila

    tiempo_inicio = time.perf_counter()

    total_emails = len(emails)
    completados = 0

    with concurrent.futures.ThreadPoolExecutor(max_workers=concurrencia) as executor:
        futuros = [executor.submit(_clasificar_indexado, indice, email)
                   for indice, email in enumerate(emails)]

        for futuro in tqdm(concurrent.futures.as_completed(futuros), total=len(futuros),
                            desc="Verificando correos", unit="email", disable=not mostrar_barra_progreso):
            indice, fila = futuro.result()
            filas_clasificacion[indice] = fila
            completados += 1
            if callback_progreso is not None:
                callback_progreso(completados, total_emails)

    tiempo_total_segundos = time.perf_counter() - tiempo_inicio

    df_clasificacion = pd.DataFrame(filas_clasificacion, columns=["email", "status", "reason", "accion"])

    # Se sustituye la columna de email original por la versión normalizada
    # (minúsculas, sin espacios) y se añaden status/reason/accion al final,
    # conservando el resto de columnas originales tal cual venían.
    df_resultado = df_entrada.drop(columns=[columna_email]).reset_index(drop=True)
    df_resultado.insert(1, "email", df_clasificacion["email"])
    df_resultado["status"] = df_clasificacion["status"]
    df_resultado["reason"] = df_clasificacion["reason"]
    df_resultado["accion"] = df_clasificacion["accion"]

    return df_resultado, tiempo_total_segundos


# --------------------------------------------------------------------------
# PROGRAMA PRINCIPAL
# --------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Limpiador de correos Fase 1 - CON VALORES (sin dependencias externas de pago)"
    )
    parser.add_argument("entrada", help="Archivo CSV o Excel con la lista de contactos")
    parser.add_argument("--columna", default=None,
                         help="Nombre de la columna con los emails (si no se indica, se detecta automáticamente)")
    parser.add_argument("--salida", default="./resultados_limpieza",
                         help="Carpeta donde guardar 'buenos.xlsx', 'revisar.xlsx' y 'eliminar.xlsx' "
                              "(por defecto ./resultados_limpieza)")
    parser.add_argument("--guardar-estandarizado", default=None,
                         help="Si se pasa una ruta, guarda ahí una copia del archivo de entrada ya "
                              "estandarizado (separador de columnas normalizado a ',' y encoding "
                              "corregido a utf-8-sig), para inspeccionar cómo quedó. Opcional, no "
                              "afecta el resultado de la limpieza.")
    parser.add_argument("--lista-negra-local", default=RUTA_LISTA_NEGRA_LOCAL,
                         help=f"Ruta al archivo editable de dominios desechables (por defecto {RUTA_LISTA_NEGRA_LOCAL})")
    parser.add_argument("--dns-timeout", type=float, default=DNS_TIMEOUT_DEFAULT,
                         help=f"Timeout en segundos para consultas DNS (por defecto {DNS_TIMEOUT_DEFAULT})")
    parser.add_argument("--smtp-timeout", type=float, default=SMTP_TIMEOUT_DEFAULT,
                         help=f"Timeout en segundos para la conexión SMTP (por defecto {SMTP_TIMEOUT_DEFAULT})")
    parser.add_argument("--smtp-timeout-proveedor-masivo", type=float,
                         default=SMTP_TIMEOUT_PROVEEDOR_MASIVO_DEFAULT,
                         help="Timeout en segundos para la conexión SMTP con proveedores masivos "
                              f"(hotmail, outlook, live, aol, yahoo...) (por defecto {SMTP_TIMEOUT_PROVEEDOR_MASIVO_DEFAULT})")
    parser.add_argument("--sin-smtp", action="store_true",
                         help="Desactiva la verificación SMTP (útil si tu red bloquea el puerto 25). "
                              "Todo lo que pase el filtro de sintaxis/desechables/DNS quedará en Unknown/REVISAR.")
    parser.add_argument("--verificacion-paciente", action="store_true",
                         help="Duplica las pausas progresivas usadas entre reintentos con proveedores "
                              "masivos, para maximizar la tasa de verificación real a costa de que "
                              "el script tarde más en total.")
    parser.add_argument("--concurrencia", type=int, default=CONCURRENCIA_DEFAULT,
                         help="Cantidad de verificaciones SMTP simultáneas en total "
                              f"(por defecto {CONCURRENCIA_DEFAULT}). Subirlo acelera el procesamiento "
                              "pero aumenta el riesgo de bloqueos de reputación de IP.")
    parser.add_argument("--concurrencia-por-dominio", type=int, default=CONCURRENCIA_POR_DOMINIO_DEFAULT,
                         help="Máximo de verificaciones SMTP simultáneas contra el MISMO dominio "
                              f"(por defecto {CONCURRENCIA_POR_DOMINIO_DEFAULT}), para no abrir varias "
                              "conexiones a la vez al mismo proveedor sin importar --concurrencia.")

    args = parser.parse_args()

    print("=" * 70)
    print("LIMPIADOR DE CORREOS - FASE 1 (CON VALORES)")
    print("=" * 70)

    lista_negra_local = cargar_o_crear_lista_negra_local(args.lista_negra_local)
    print(f"[INFO] {len(lista_negra_local)} dominio(s) en la lista negra local.")

    print(f"[INFO] Leyendo contactos desde: {args.entrada}")
    df_entrada, columna_email = leer_lista_contactos(args.entrada, args.columna)
    print(f"[INFO] {len(df_entrada)} correos encontrados (tras separar celdas con múltiples direcciones).")
    print(f"[INFO] Columnas originales que se conservarán: {list(df_entrada.columns)}")

    if args.sin_smtp:
        print("[AVISO] Verificación SMTP DESACTIVADA. Los correos con sintaxis y "
              "dominio válidos quedarán como Unknown/REVISAR.")

    if args.verificacion_paciente:
        print("[INFO] Modo --verificacion-paciente activado: pausas dobles con proveedores "
              "masivos para priorizar la tasa de verificación sobre la velocidad.")

    print(f"[INFO] Concurrencia: hasta {args.concurrencia} verificaciones SMTP simultáneas "
          f"(máx. {args.concurrencia_por_dominio} contra el mismo dominio a la vez).")

    if args.guardar_estandarizado:
        df_entrada.to_csv(args.guardar_estandarizado, index=False, encoding="utf-8-sig")
        print(f"[INFO] Entrada ya estandarizada (separador/encoding corregidos) guardada en: "
              f"{args.guardar_estandarizado}")

    df_resultado, tiempo_total_segundos = clasificar_dataframe(
        df_entrada, columna_email, lista_negra_local,
        dns_timeout=args.dns_timeout,
        smtp_timeout=args.smtp_timeout,
        smtp_timeout_proveedor_masivo=args.smtp_timeout_proveedor_masivo,
        verificar_smtp_activo=not args.sin_smtp,
        verificacion_paciente=args.verificacion_paciente,
        concurrencia=args.concurrencia,
        concurrencia_por_dominio=args.concurrencia_por_dominio,
    )

    rutas = guardar_resultados(df_resultado, args.salida)

    print("\n" + "=" * 70)
    print("RESUMEN")
    print("=" * 70)
    resumen = df_resultado["accion"].value_counts()
    for accion in ["MANTENER", "ELIMINAR", "REVISAR"]:
        print(f"  {accion:10s}: {resumen.get(accion, 0)}")
    print(f"  TOTAL     : {len(df_resultado)}")

    minutos, segundos = divmod(tiempo_total_segundos, 60)
    print(f"\nTiempo total de verificación: {int(minutos)} min {segundos:.1f} s "
          f"(concurrencia={args.concurrencia}, por dominio={args.concurrencia_por_dominio})")

    # Desglose de proveedores masivos: cuántos se verificaron por SMTP con
    # certeza (aceptado/rechazado real) vs. cuántos quedaron sin confirmar
    # por bloqueo del proveedor, para poder medir si el ajuste de pausas
    # mejora la tasa de verificación real.
    es_dominio_masivo = df_resultado["email"].astype(str).str.rsplit("@", n=1).str[-1].apply(
        lambda d: es_proveedor_masivo(d) if isinstance(d, str) else False
    )
    confirmados_masivos = int((es_dominio_masivo & df_resultado["reason"].isin(["accepted_email", "rejected_email"])).sum())
    bloqueados_masivos = int((es_dominio_masivo & (df_resultado["reason"] == "smtp_bloqueado_proveedor_masivo")).sum())
    total_masivos_evaluados = confirmados_masivos + bloqueados_masivos

    if total_masivos_evaluados > 0:
        tasa_confirmacion = confirmados_masivos / total_masivos_evaluados * 100
        print("\nProveedores masivos (hotmail/outlook/live/aol/yahoo y variantes):")
        print(f"  Verificados con certeza (deliverable/undeliverable): {confirmados_masivos}")
        print(f"  Bloqueados por el proveedor (probable_deliverable) : {bloqueados_masivos}")
        print(f"  Tasa de verificación real                         : {tasa_confirmacion:.1f}%")

    # Rechazos 5xx que en realidad son bloqueo de reputación de NUESTRA IP
    # (Spamhaus/PBL/RBL), no una confirmación real de "buzón inexistente".
    # Aplica a cualquier dominio, no solo a proveedores masivos.
    bloqueados_reputacion = int((df_resultado["reason"] == "rechazo_por_reputacion_ip_propia").sum())
    if bloqueados_reputacion > 0:
        print(f"\nRechazos 5xx por reputación de nuestra IP (no del buzón), enviados "
              f"a REVISAR (probable_deliverable): {bloqueados_reputacion}")

    # Dominios catch-all: el 250 a la dirección real no confirmó nada porque
    # el servidor acepta cualquier dirección (ver DetectorCatchAll).
    catch_all_detectados = int((df_resultado["reason"] == "dominio_catch_all").sum())
    if catch_all_detectados > 0:
        print(f"\nDominios catch-all detectados (accepted_email no confiable, "
              f"enviados a REVISAR): {catch_all_detectados}")

    catch_all_no_verificable = int((df_resultado["reason"] == "catchall_no_verificable").sum())
    if catch_all_no_verificable > 0:
        print(f"No se pudo determinar si el dominio es catch-all (enviados a REVISAR): "
              f"{catch_all_no_verificable}")

    print("\nArchivos generados:")
    print(f"  - {rutas['buenos']}    (accion = MANTENER, listo para campañas)")
    print(f"  - {rutas['revisar']}   (accion = REVISAR, revisión manual)")
    print(f"  - {rutas['eliminar']}  (accion = ELIMINAR, descarte)")

    print("\nListo. Los 3 archivos tienen AutoFiltro y encabezado congelado. "
          "'buenos.xlsx' contiene solo reason == 'accepted_email' de dominios confirmados "
          "NO catch-all (confirmación SMTP real de ESA dirección específica). Dentro de "
          "'revisar.xlsx', filtrá la columna 'reason' para ver 'smtp_bloqueado_proveedor_masivo' "
          "(hotmail/outlook/live/aol/yahoo sin confirmar), 'rechazo_por_reputacion_ip_propia' "
          "(5xx bloqueado por reputación de nuestra IP, no del buzón), 'dominio_catch_all' "
          "(el dominio acepta cualquier dirección, el 250 no confirma nada) o "
          "'catchall_no_verificable' (no se pudo determinar) — ninguno de estos es una "
          "confirmación real, por eso van a revisión manual en vez de a la campaña directa.")


if __name__ == "__main__":
    main()
