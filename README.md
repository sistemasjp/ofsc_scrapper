# Scraper de consumo de materiales - OFSC Claro Ecuador

Automatiza lo que hicimos manualmente: recorre la Consola de despacho,
identifica actividades "finalizada" por cuadrilla/fecha, y extrae los
materiales/equipos instalados (con código SAP, cantidad y serie) directamente
desde las respuestas internas que la página ya consulta — es más rápido y más
confiable que leer la tabla visual.

## 1. Instalación (una sola vez, en cada computadora)

Copia esta carpeta completa **sin la subcarpeta `venv/`** (no es portable
entre computadoras/sistemas operativos) y sin ningún `.xlsx` o `.env` que
traiga tu copia — genera todo eso de nuevo ahí:

```bash
python3 -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
playwright install chromium
```

## 2. Configurar `config.json`

- `crews`: lista de códigos de cuadrilla exactamente como aparecen en el
  árbol izquierdo de la consola.
- `skip_activity_types`: tipos administrativos que nunca tienen consumo
  (Almuerzo, Bodega, etc.) — no se abren ni se registran en el Excel (ni en
  "Consumo Detallado" ni en "Resumen Actividades"), solo se imprimen en la
  consola. Lo mismo aplica a cualquier actividad cuyo estado no esté en
  `estados_con_posible_consumo`.

## 3. Primer uso (recomendado: login manual)

```bash
python ofsc_scraper.py --date 2026-09-08 --output consumo_2026-09-08.xlsx
```

Se abre una ventana de Chrome. Inicia sesión tú mismo (usuario/clave, y
cualquier SSO/MFA que tenga tu empresa) y cuando veas la Consola de
despacho, vuelve a la terminal y presiona ENTER. A partir de ahí todo es
automático: cambia la fecha, recorre cada cuadrilla, abre solo las
actividades finalizadas, y va guardando el Excel tras cada cuadrilla (si se
corta a la mitad, no pierdes lo ya hecho).

## 4. Uso desatendido (sin ventana, sin intervención humana)

Solo tiene sentido si tu login **no** usa SSO/MFA (usuario+clave simple).

Guarda tus credenciales una sola vez en un archivo `.env` (nunca se sube a
control de versiones, ya está en `.gitignore`):

```bash
cp .env.example .env
nano .env   # completa OFSC_USERNAME y OFSC_PASSWORD
```

El script lee automáticamente ese `.env` (sin necesidad de `export` manual)
gracias a `python-dotenv`. Si prefieres no usar `.env`, sigue funcionando
exportar las variables tú mismo, o dejarlas vacías para que te las pida por
consola (`getpass`, oculto).

Antes de ir a headless, corre una vez con `--auto-login` y ventana visible
para verificar que detecta bien el formulario:

```bash
python ofsc_scraper.py --date 2026-09-08 --auto-login
```

Si el script no encuentra los campos de usuario/clave, abre el formulario de
login con F12 -> clic derecho sobre el campo -> "Inspeccionar", y ajusta las
listas `LOGIN_USER_SELECTORS` / `LOGIN_PASS_SELECTORS` /
`LOGIN_SUBMIT_SELECTORS` al inicio de `ofsc_scraper.py`.

Una vez validado, ya puedes correr desatendido:

```bash
python ofsc_scraper.py --date 2026-09-08 --auto-login --headless
```

### Comando corto (`reporte.sh` / `reporte.bat`)

Para que cualquiera del equipo lo corra sin recordar flags, pásale solo la
fecha. Hay dos versiones — usa la que corresponda a tu sistema operativo,
**no son intercambiables**:

- **Linux / macOS** (terminal):
  ```bash
  ./reporte.sh 2026-09-08
  ```
- **Windows** (símbolo del sistema / `cmd`, o doble clic y luego escribir la
  fecha cuando la pida):
  ```bat
  reporte.bat 2026-09-08
  ```

Ambos hacen lo mismo: `--auto-login --headless`, y guardan el resultado como
`consumo_2026-09-08.xlsx` en esta misma carpeta. Ambos fallan con un mensaje
claro si no existe `.env` todavía (o, en Windows, si no existe `venv\`).

> `reporte.sh` **no funciona en Windows** con doble clic ni desde `cmd`/PowerShell
> (es un script de shell) — ahí necesitan `reporte.bat`. Al revés tampoco:
> `reporte.bat` no sirve en Linux/macOS.

Y se puede programar para que corra solo cada noche con la fecha del día:

```bash
# Linux/macOS: crontab -e, ejemplo a las 23:30 todos los días
30 23 * * * /ruta/a/ofsc_scraper/reporte.sh $(date +\%Y-\%m-\%d)
```

En Windows, con el "Programador de tareas" (Task Scheduler): crear una
tarea que ejecute `reporte.bat` con el argumento `%date:~-4,4%-%date:~-10,2%-%date:~-7,2%`
como fecha (el formato exacto de `%date%` depende de la configuración
regional de Windows — más simple y confiable es crear un `.bat` de una línea
con la fecha fija y regenerarlo/editarlo cada día, o pedirle a la persona de
sistemas que lo arme con PowerShell usando `Get-Date -Format "yyyy-MM-dd"`).

## 5. Reanudar tras un corte

```bash
python ofsc_scraper.py --date 2026-09-08 --output consumo_2026-09-08.xlsx --resume
```

Omite las cuadrillas que ya tienen filas en la hoja "Resumen Actividades" del
Excel indicado, y sigue con las que faltan.

## Salida

Un `.xlsx` con dos hojas:

- **Consumo Detallado**: una fila por material/equipo consumido, con
  cuadrilla, fecha, actividad, orden, cliente, **ciudad**, tipo y estado de
  la actividad, tipo de inventario, descripción, modelo, serie, ID de
  inventario, código SAP y cantidad. Regla de identificación: para
  **Equipos** basta con el número de serie (OFSC casi nunca expone su
  código SAP y eso es normal, no se advierte); para **Materiales** sí se
  exige el código SAP. Solo se resalta en amarillo una fila cuando falta
  el dato realmente requerido para ese tipo (Equipos sin serial, o
  Materiales sin código SAP).
- **Resumen Actividades**: una fila por cada actividad "finalizada" que el
  script realmente abrió (con su estado, ciudad, y cuántos ítems tuvo —
  puede ser 0 si de verdad no consumió nada). Las actividades que nunca se
  abren (administrativas, o en un estado sin consumo esperado) no aparecen
  en ninguna hoja — ver `skip_activity_types` / `estados_con_posible_consumo`
  arriba.

## Notas / riesgos a tener en cuenta

- Este script usa endpoints **internos, no documentados** de la app (no es
  la API REST pública de OFSC). Si Oracle actualiza la versión de la
  consola, la estructura puede cambiar y el script dejar de funcionar. Si
  eso pasa, lo normal es que un `expect_response(...)` truene por timeout;
  revisa con F12 -> Network qué URL/formato cambió.
- Si en el futuro planean mandar este consumo automáticamente a su sistema
  principal (SAP u otro), vale la pena evaluar si su cuenta OFSC tiene
  acceso a la **API REST oficial de Oracle Field Service** (Core/Metadata,
  con OAuth). Es la vía soportada y estable para integraciones — se los
  comento en el mensaje del chat, no es parte de este script.
- La selección de fecha (`set_date`) navega con las flechas
  anterior/siguiente del encabezado; si tu instancia tiene un selector de
  calendario distinto, puede requerir ajuste (está aislado en su propia
  función para facilitarlo).
