# Auditorías 5S — Frontend (GitHub Pages)

App estática (HTML/CSS/JS, sin build) para cargar y consultar Auditorías 5S
de Escorial (Planta / Pañol / Oficina). Se conecta a la API en
`auditorias-5s-api` (Cloudflare Worker + Neon + R2).

## 1. Configurar la URL de la API

Editá `js/config.js` y reemplazá el placeholder por la URL real del Worker
(la que te dio `wrangler deploy` en el repo de la API):

```js
export const API_BASE = window.__API_BASE__ || 'https://auditorias-5s-api.TU-SUBDOMINIO.workers.dev';
```

## 2. Publicar en GitHub Pages

```bash
git init
git add .
git commit -m "Auditorías 5S — v1"
git remote add origin https://github.com/diegortizdao-collab/Auditorias-5S.git
git push -u origin main
```

En GitHub: **Settings → Pages → Deploy from a branch → main / (root)**.
La app queda en `https://diegortizdao-collab.github.io/Auditorias-5S/`.

Después volvé al repo de la API y poné esa URL en `ALLOWED_ORIGIN`
(`wrangler.toml`), y volvé a desplegar el Worker (`npx wrangler deploy`) para
que el CORS deje pasar los pedidos del frontend ya publicado.

## Estructura

- `index.html` — shell de la app (topbar, navegación, `<main>`).
- `css/style.css` — mismo sistema de diseño ya validado en el preview
  (paleta celeste Escorial, Barlow Condensed + IBM Plex Sans/Mono).
- `js/config.js` — URL de la API.
- `js/api.js` — wrapper de `fetch` para cada endpoint.
- `js/app.js` — router (hash) + las pantallas (Selección, Formulario,
  Evaluación, Dashboard, Informe, Plan de Acción).
- `js/charts.js` — genera los SVG de radar / evolución / comparativo con
  datos reales (misma geometría que el preview de diseño, ahora calculada
  dinámicamente).
- `js/xlsx-export.js` — genera el Informe en `.xlsx`: Hoja 1 (réplica exacta
  de la solapa "8. Auditoría 5S y KPI") + hoja "Fotos" (cada foto cargada,
  incrustada de verdad, no como link — se descarga de R2 y se normaliza a
  PNG con un `<canvas>` antes de insertarla, así no importa si la cámara la
  guardó en jpg/png/webp) + hoja "Plan de Accion" con las acciones de esa
  auditoría, usando ExcelJS (cargado por CDN).
- `assets/` — logos de Escorial y SPE.

## Cómo se usa

1. **Selección** — elegís tipo (Pañol / Planta / Oficina) y dónde es la
   auditoría, y creás el registro. Para tipo Planta la ubicación es
   Planta → UET → Sector, en cascada real: el combo de Sector solo se
   habilita después de elegir la UET, y muestra únicamente los sectores de
   esa UET (agrupados en subgrupos como "Soldadura" o "Pintura" cuando la
   UET los tiene). Si ese sector ya tiene acciones correctivas pendientes de
   una auditoría anterior, se pasa directo a **Revisión** en vez de al
   Formulario.
2. **Revisión de Plan de Acción** *(solo aparece si hay algo pendiente de la
   vez pasada en ese mismo sector)* — se listan las acciones abiertas /
   en proceso de auditorías anteriores, agrupadas por ítem. Por cada una se
   puede marcar **"✓ Marcar cumplida"** (la cierra, con fecha de cierre de
   hoy), y para el ítem correspondiente aparece la escala 0/1/3/5 para
   puntuarlo ahí mismo, en base a lo que se ve al revisar si la acción
   realmente se cumplió. Ese puntaje queda cargado en la auditoría de hoy —
   así es como el cierre de una acción se termina reflejando en el puntaje
   de la auditoría siguiente, no como algo separado. Al continuar, ese ítem
   ya aparece puntuado (y editable) al llegar a Evaluación.
3. **Formulario** — completás el encabezado (evaluador, turno, etc.). Tiene
   también un acceso directo a "Revisar plan de acción anterior" por si se
   quiere volver a mirar.
4. **Evaluación** — puntuás cada ítem (0/1/3/5), con comentario y foto
   opcionales. El puntaje se guarda al toque, no hace falta "guardar todo"
   al final. Al lado del nombre de cada ítem hay un ícono "ℹ" que despliega
   los 4 criterios (qué distingue a un 0 de un 1, un 3 o un 5) — colapsado
   por defecto para no alargar la pantalla, pero pensado para que alguien
   que recién arranca con la app pueda consultarlo antes de elegir un
   puntaje en vez de puntuar a ciegas; una vez puntuado, el criterio elegido
   queda resaltado ahí mismo. Si un ítem queda en **0 o 1**, la IA
   (Cloudflare Workers AI) sugiere automáticamente una acción correctiva
   para ese ítem — se muestra ahí mismo, debajo del ítem, con un botón "↻"
   para pedirle otra sugerencia si no convence. También se puede agregar a
   mano cualquier cantidad de acciones manuales sobre un ítem ("+ Agregar
   acción manual"), con responsable y fecha de vencimiento opcionales.
5. **Informe** — al finalizar, se genera automáticamente la Hoja 1 (idéntica
   al Excel que exige la norma) más una **Hoja 3 · Plan de Acción** con
   todas las acciones que dejó esa auditoría (editable ahí mismo, con el
   mismo selector de estado), y se puede descargar todo en `.xlsx`
   (3 hojas: "Hoja 1", "Fotos" —con las fotos realmente incrustadas, no un
   link— y "Plan de Accion"). Al lado del botón de descarga hay un botón
   **"✉ Enviar por correo"**: descarga el mismo `.xlsx` y, en el mismo click,
   abre el cliente de correo predeterminado del usuario (Outlook u otro) con
   un borrador prellenado (sector, fecha y puntaje) y un recordatorio de
   adjuntar el archivo recién descargado. No hay envío automático ni
   servidor de correo de por medio — por eso no requiere dominio propio ni
   configuración adicional; el usuario decide si finalmente lo manda, y
   puede editar el borrador o cancelarlo antes de enviar.
6. **Dashboard** — filtrable por tipo / planta / UET / sector / año / mes
   ("Todos" en año o mes promedia todas las fechas). El radar y el puntaje
   total muestran el **promedio** de todas las auditorías que matchean el
   segmentador elegido (no solo la última — con un dato de contexto aparte,
   "Última auditoría", para saber quién la hizo y cuándo), la evolución
   mensual sigue siendo Objetivo vs. Auditado para un año puntual, y hay tres
   comparativos (todos como promedio, no como última auditoría): por
   sector, por UET y por planta.
7. **Plan de Acción** — pantalla independiente de la Evaluación (accesible
   desde Selección, Dashboard y desde cada Evaluación) con todas las
   acciones —sugeridas por IA y manuales— de todas las auditorías, marcadas
   con su origen (🤖 IA / ✋ Manual). Se filtra por estado (abierta / en
   proceso / cerrada), tipo, planta y sector, y el estado de cada acción se
   puede cambiar ahí mismo, sin entrar a la auditoría.

Sin login (v1): cualquiera con el link puede cargar y ver auditorías.

Pie de página con crédito de autoría ("Desarrollado por Mgter. Diego Ortiz —
Líder de Mejora Continua | Consultor ISO") — estándar a incluir en toda app
nueva que se genere para Escorial / Silver / freelance (título genérico
multiempresa, igual al de su LinkedIn, en vez del cargo puntual en cada
empresa).

Todo el flujo (crear → revisión de acciones pendientes → puntuar → foto →
cerrar → informe con Hoja 3 → dashboard → plan de acción) fue probado de
punta a punta con datos reales antes de la entrega, incluyendo el caso sin
conexión a Workers AI (el puntaje se guarda igual, la sugerencia
simplemente no aparece esa vez) y el ciclo completo de abrir una auditoría
nueva sobre un sector con una acción pendiente, cerrarla en Revisión,
puntuar el ítem ahí mismo, y verlo ya cargado al llegar a Evaluación.
