# TAXITRONIC TX12460

App de control de caja diaria para taxistas, con estética de taxímetro
Taxitrónico. Registra Efectivo, Tarjeta, Uber, Freenow, Bolt, Emisora,
Combustible y Otros gastos por día, guarda un histórico y calcula la
facturación entre dos fechas.

Todo el código es un único archivo `index.html` (HTML + CSS + JS, sin
dependencias de build). Los datos se guardan en el propio navegador
(`localStorage`-style vía la API de almacenamiento del entorno de Claude).
Opcionalmente puede sincronizar cada día con una Google Sheet mediante el
script de `apps-script/`.

## 1. Subir a GitHub y publicarlo (GitHub Pages)

1. Crea un repositorio nuevo en GitHub (por ejemplo `taxitronic`).
2. Sube estos archivos manteniendo la estructura:
   ```
   index.html
   apps-script/Code.gs
   apps-script/appsscript.json
   README.md
   ```
3. En el repo: **Settings → Pages → Source: Deploy from a branch**,
   elige la rama `main` y la carpeta `/ (root)`. Guarda.
4. En un par de minutos tu app estará en
   `https://<tu-usuario>.github.io/<nombre-repo>/`.
5. Desde el móvil, abre esa URL en Chrome/Safari y usa
   "Añadir a pantalla de inicio" para que quede como una app instalada.

No hace falta ningún paso de compilación: es HTML puro, funciona tal cual
lo subas.

## 2. Conectar con Google Sheets (opcional, para tener respaldo/backup)

Esto hace que, además de guardarse en el propio navegador, cada día que
edites en la app se escriba también como fila en tu Google Sheet
"taxitronic" (carpeta de Drive) — así tienes los datos accesibles desde
cualquier sitio y puedes trabajar con ellos en Excel/Sheets.

1. Abre tu Google Sheet **taxitronic** en Drive → pestaña **Datos diarios**.
2. Menú **Extensiones → Apps Script**.
3. Borra el contenido de `Code.gs` que aparece por defecto y pega el
   contenido de `apps-script/Code.gs` de este repo.
4. Con el icono de engranaje (⚙) a la izquierda, marca
   **"Mostrar archivo de manifiesto appsscript.json"**. Ábrelo y sustituye
   su contenido por el de `apps-script/appsscript.json` de este repo.
5. Arriba a la derecha: **Implementar → Nueva implementación**.
   - Tipo: **Aplicación web**.
   - Ejecutar como: **Yo (tu cuenta)**.
   - Quién tiene acceso: **Cualquier usuario**.
6. Pulsa **Implementar**, autoriza los permisos que pida Google, y copia
   la **URL de la aplicación web** que te entrega (termina en `/exec`).
7. Abre `index.html` (en tu repo o en GitHub), busca la línea:
   ```js
   const SYNC_URL = '';
   ```
   y pega tu URL entre las comillas:
   ```js
   const SYNC_URL = 'https://script.google.com/macros/s/XXXXXXXX/exec';
   ```
8. Guarda y vuelve a subir el archivo a GitHub (o edítalo directamente
   en GitHub con el lápiz de edición y haz commit).

A partir de ahí, cada vez que añadas una carrera, un gasto, deshagas,
corrijas un día o envíes la hoja, la app intentará mandar ese día
actualizado a la fila correspondiente de "Datos diarios" (sin bloquear
la app si no hay conexión: los datos siempre quedan a salvo en local
primero).

### Notas sobre el script

- Busca si ya existe una fila con esa fecha **y esa licencia** en las
  columnas A y O; si la encuentra, la actualiza. Si no, usa la primera
  fila libre del rango ya preparado (hasta la fila 400).
- No toca las columnas **Facturado**, **Gastos** ni **Neto**: esas ya
  llevan fórmulas en la plantilla y se recalculan solas.
- Si algún día ves el error "No quedan filas libres" en los registros
  del script (Ejecuciones, en el editor de Apps Script), simplemente
  amplía el rango de fórmulas en el Excel más allá de la fila 400 y
  sube `LAST_PREPARED_ROW` en `Code.gs` al mismo número.

## 3. Usuarios y contraseñas

La app ahora pide iniciar sesión. Hay 5 licencias más un usuario admin:

| Usuario | Contraseña | Fijo diario |
|---|---|---|
| 12460 | 0 | 120 € |
| 12453 | 11 | 130 € |
| 3493 | 222 | 130 € |
| 12614 | 3333 | 120 € |
| 12553 | 4444 | 130 € |
| admin | admin00 | (no tiene) |

- Cada licencia solo ve y edita **sus propios días** (los datos se guardan
  por usuario en el propio dispositivo). El usuario **admin** no tiene
  consola para meter carreras — solo entra en un panel donde puede
  **cambiar la contraseña y el fijo diario** de cada licencia. Esos
  cambios se guardan en el mismo dispositivo donde el admin inicie
  sesión y edite — si cada conductor usa su propio móvil en vez de un
  único dispositivo compartido en el coche, el admin tendría que repetir
  los cambios en cada uno (no hay backend central que los reparta solo).
- Puedes cambiar las contraseñas/fijos por defecto directamente en el
  código (`DEFAULT_USERS` en `index.html`) si quieres que arranquen ya
  con otros valores de fábrica.

## 4. Cambios que tienes que hacer tú a mano en el Google Sheet

1. En la pestaña **"Datos diarios"**, añade dos cabeceras nuevas:
   - Celda **O1**: escribe `Licencia`.
   - Celda **P1**: escribe `Refuerzo`.
   No hace falta ponerles fórmula ni formato especial, solo el texto.
2. Tras pegar el `Code.gs` nuevo, tienes que **crear una nueva versión de
   la implementación** (Implementar → Gestionar implementaciones → lápiz
   de editar → Versión "Nueva versión" → Implementar), porque el script
   pide permisos de Drive y Documentos (para poder generar el PDF) —
   Google te pedirá autorizar de nuevo, acéptalo.
3. Los PDF se guardan solos, organizados **por licencia**, dentro de tu
   carpeta "taxitronic" de Drive, así:
   ```
   taxitronic/PDFs/12460/recaudacion_12460_2026-09-15.pdf
   taxitronic/PDFs/3493/recaudacion_3493_2026-09-15.pdf
   ...
   ```
   El script crea solo la carpeta "PDFs" y, dentro, una subcarpeta por
   cada licencia la primera vez que hace falta — no tienes que crear nada
   tú a mano.
4. Si el PDF sigue sin crearse, la app ahora te muestra el motivo exacto
   en la línea de estado tras pulsar "Enviar hoja diaria" (por ejemplo,
   un error de permisos de Drive o de Documentos). Revisa ese mensaje —
   casi siempre significa que falta el paso 2 (nueva versión + reautorizar).

## 5. "No graba nada hasta enviar hoja diaria"

Mientras vas metiendo carreras y gastos del día, todo se guarda al
momento **en el propio dispositivo**, pero ese día no aparece todavía en
la pantalla de **Histórico** ni se manda al Google Sheet. Ambas cosas
pasan a la vez, solo al pulsar **"ENVIAR HOJA DIARIA"**: ahí se guarda la
fila en el Excel, se genera el PDF, y el día pasa a verse en el
Histórico. Una vez enviado, ese día queda bloqueado para seguir
añadiendo directamente — solo se puede tocar desde el botón **Corregir**.

## 6. Actualizar la app más adelante

Como es un único `index.html`, para publicar cambios solo tienes que
editar el archivo y hacer commit — GitHub Pages lo republica solo en
1-2 minutos.
