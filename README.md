# Prospección diaria en Google Maps (nube)

Todos los días a las **6:07 a. m. (hora Colombia)**, GitHub busca en Google Maps odontólogos, clínicas, médicos estéticos, cirujanos plásticos y otros consultorios privados en Manizales, Pereira, Medellín y Bogotá. Después limpia la lista y agrega **solo los negocios nuevos** a `Prospectos.xlsx`, ordenados por prioridad.

No necesitas tener el Mac encendido ni instalar nada.

## Uso diario

1. Entra al repositorio y haz clic en **`Prospectos.xlsx`** → botón **Download** (o "Download raw file").
2. Empieza por la hoja **Nuevos hoy**.
3. Si actualizas **Estado** y **Notas**, súbelo de vuelta **el mismo día**: en la página principal del repositorio, entra a **Add file → Upload files**, arrastra el Excel (con el mismo nombre) y haz clic en **Commit changes**.

> ⚠️ Sube siempre la versión que descargaste **ese mismo día**. Si subes una copia vieja, se pierden los prospectos que entraron después. En todo caso, nada se pierde del todo: en el historial del archivo (botón **History**) están todas las versiones anteriores.

## Correrlo a mano

Pestaña **Actions → Prospección diaria → Run workflow**. Tarda entre 30 y 60 minutos. En cada corrida verás un resumen con cuántos prospectos nuevos entraron.

## Cómo leer el Excel

| Hoja | Para qué sirve |
|---|---|
| **Nuevos hoy** | Lo que entró en la última corrida |
| **Prospectos** | Toda tu base acumulada, con **Estado** (lista desplegable) y **Notas** |
| **Resumen** | Totales por ciudad, nicho y etapa del embudo |

**Prioridad**, calculada con puntos:
- Sin página web: +3
- Solo tiene Instagram, Facebook o Doctoralia: +2
- Calificación menor a 4.0: +2
- Menos de 15 reseñas: +1
- Tiene celular (WhatsApp): +1
- Sin teléfono: −2

Con 4 puntos o más es **Alta**, con 2 o 3 es **Media** y con menos es **Baja**. La columna **WhatsApp** abre el chat con un clic. Solo aparece cuando el número es un celular colombiano.

## Personalizar (`config.json`)

Haz clic en el archivo → ícono del lápiz ✏️ → edita → **Commit changes**.
- `nichos`: agrega o quita tipos de negocio.
- `ciudades`: los barrios o zonas de cada ciudad. Cada día se busca en una zona distinta para no repetir negocios. Agrega más zonas para que tarde más en repetirse.
- `zonas_por_dia`: súbelo a 2 si quieres más volumen. Ojo con los minutos gratis (ver abajo).
- `calificacion_baja` y `pocas_resenas`: los umbrales para marcar oportunidades.

Para cambiar la hora, edita la línea `cron` en `.github/workflows/prospeccion.yml`. La hora está en UTC: Colombia es UTC−5.

## Límites del plan gratis

- Un repositorio privado tiene **2.000 minutos al mes** de GitHub Actions. Una corrida diaria de 40 a 60 minutos gasta entre 1.200 y 1.800 minutos. Puedes revisar tu consumo en GitHub → Settings → Billing.
- Si te acercas al límite, baja `profundidad_scroll` a 2 o quita algunos nichos.

## Si un día no llega nada

- Revisa la corrida en **Actions**. Si falló el paso "Scrapear Google Maps", lo más común es que Google haya bloqueado esa IP ese día. Al día siguiente sale de otra IP y suele funcionar.
- Si falla varios días seguidos, puede que Google haya cambiado su página. Avísame y lo ajustamos.

## Uso responsable

- Hacer scraping de Google Maps va contra sus términos de servicio. Es una práctica común, pero el riesgo es tuyo.
- Ley 1581 de 2012: escribe de forma **personalizada y uno a uno**, nunca con envíos masivos. Ofrece la opción de no volver a escribirle. A quien te lo pida, márcalo como **"No contactar"**.
- Envía como máximo entre 20 y 30 mensajes de WhatsApp al día desde tu número, para no perderlo por reportes.

Las plantillas de mensajes están en `mensajes_whatsapp.md`.
