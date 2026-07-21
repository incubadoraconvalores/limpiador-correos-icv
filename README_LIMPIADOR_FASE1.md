# Limpiador de correos - Fase 1 (CON VALORES)

## Instalación (una sola vez)

1. Instalar Python 3.10+ (https://www.python.org)
2. Abrir una terminal en esta carpeta y ejecutar:

   pip install -r requirements.txt

## Uso básico (línea de comandos)

    python limpiador_correos_fase1.py mi_lista.xlsx

- Si el archivo es un `.csv` con separador `;` en vez de `,`, o con tildes
  mal codificadas (encoding roto), se detecta y corrige automáticamente
  antes de procesar — no hace falta corregirlo a mano.
- Detecta automáticamente la columna de email (busca "email" o "correo" en
  el nombre de columna, con prioridad a coincidencias exactas).
- Si no la detecta, indícala manualmente:

    python limpiador_correos_fase1.py mi_lista.csv --columna correo_electronico

## Interfaz web (Streamlit)

Para usar el limpiador sin la terminal (subir un archivo y descargar los
resultados desde el navegador):

    streamlit run streamlit_app.py

Se abre una página local con: subir archivo → botón "Verificar correos" →
resumen visual (total y cantidad/% por MANTENER/REVISAR/ELIMINAR) → 3
botones de descarga.

**Aviso:** si esta interfaz se despliega en un hosting compartido/gratuito
que bloquea el puerto 25 saliente (común en varios servicios cloud), la
verificación SMTP no podrá conectarse y todo quedará en REVISAR. Corré la
interfaz en la misma red/máquina donde ya sabés que la CLI puede conectarse
por SMTP, o usá el checkbox "Desactivar verificación SMTP" de la interfaz.

## Archivos que genera (carpeta `resultados_limpieza/` por defecto)

- `buenos.xlsx`   -> accion = MANTENER (correos confirmados, o mantenidos por
                     confianza en proveedores masivos/reputación de IP) —
                     esta es la lista lista para la campaña (ej. Acumbamail)
- `revisar.xlsx`  -> accion = REVISAR — para revisión manual
- `eliminar.xlsx` -> accion = ELIMINAR — descarte (sintaxis inválida,
                     desechables, dominio inexistente, rechazo SMTP genuino);
                     ya validado con 100% de precisión contra Bouncer

Los 3 archivos incluyen todas las columnas originales del archivo de entrada
más `status`/`reason`/`accion`, con AutoFiltro activado en los encabezados y
esa fila congelada.

## Lista negra local de desechables

El archivo `lista_negra_local.txt` se puede editar a mano (sin tocar el código)
para añadir dominios desechables que la librería `disposable-email-domains`
todavía no incluya. Una línea = un dominio, sin @. Las líneas con # son comentarios.

## Limitación importante: puerto SMTP (25)

Algunos proveedores de internet domésticos (y muchos hostings compartidos/
gratuitos) bloquean el puerto 25 de salida, necesario para la verificación
SMTP. Si tu conexión lo bloquea, todos los correos con sintaxis y dominio
válidos caerán en REVISAR (unavailable_smtp) en lugar de MANTENER/ELIMINAR.

Si sabes que tu red bloquea el puerto 25 y quieres saltarte directamente
la verificación SMTP (para ir más rápido), usa:

    python limpiador_correos_fase1.py mi_lista.xlsx --sin-smtp

## Otras opciones

    --salida CARPETA                     Carpeta de salida (por defecto ./resultados_limpieza)
    --guardar-estandarizado RUTA         Guarda una copia del archivo de entrada ya
                                          estandarizado (separador ',' + encoding utf-8-sig),
                                          por si querés inspeccionarlo. Opcional.
    --lista-negra-local RUTA             Ruta al archivo editable de desechables
    --dns-timeout SEGUNDOS                Timeout para consultas DNS (por defecto 5)
    --smtp-timeout SEGUNDOS                Timeout para conexión SMTP (por defecto 8)
    --smtp-timeout-proveedor-masivo SEG    Timeout SMTP para hotmail/outlook/live/aol/yahoo (por defecto 20)
    --verificacion-paciente               Duplica las pausas de reintento con proveedores masivos
                                          (más lento, mayor tasa de verificación real)
    --concurrencia N                      Verificaciones SMTP simultáneas en total (por defecto 12)
    --concurrencia-por-dominio N          Máximo simultáneo contra el MISMO dominio (por defecto 2)
