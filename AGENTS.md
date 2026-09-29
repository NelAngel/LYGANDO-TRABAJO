# MEMORIA COMPLETA DEL PROYECTO — LEE SIEMPRE (se carga automáticamente al iniciar sesión)

## Qué es esto
Página web pública de ofertas de empleo del Perú: `index.html` (HTML puro, CSS integrado),
leída en el navegador con **Live Server de Visual Studio Code**. NADA de PHP, MySQL ni phpMyAdmin.
Base de datos = DOS+Tres Excels leídos con SheetJS (`xlsx.full.min.js` local):

- **`ofertas.xlsx`**                     -> anuncios de FOTOS (manual, con OCR).
- **`convocatorias.xlsx`**               -> convocatorias ONPE por SCRAPING, generado por `scraper_onpe.py`.
- **`convocatoriasdetrabajo.xlsx`**      -> convocatorias del Estado en general (CAS/728/276/
  prácticas de cualquier institución) por SCRAPING de `www.convocatoriasdetrabajo.com`,
  generado por `scraper_convocatoriasdetrabajo.py`.

`script.js` carga todos (`FUENTES = ['ofertas.xlsx','convocatorias.xlsx',
'convocatoriasdetrabajo.xlsx', ...]`), fusiona, ordena por `fecha` desc y marca cada
tarjeta con su etiqueta de origen. Si falta un Excel, se salta sin romper la página.

Características de la página: banner "+N ofertas en M regiones del Perú", estadísticas
(total / con enlace / con whatsapp / regiones), oferta destacada (`destacado=si`, o la más
reciente), chips de región top 12, buscador + filtros (TODAS / CON ENLACE / CON WHATSAPP),
sección "Cómo aplicar" y footer con descarga de la plantilla Excel.

## REGLA DE ORO (nunca olvidar / instrucción directa del usuario)
Cada oferta DEBE tener SIEMPERE estos 4 campos llenos:

1. **TRABAJO**     -> columna `titulo`
2. **LUGAR**       -> columna `ubicacion`
3. **DESCRIPCIÓN** -> columna `descripcion`
4. **CONTACTO**    -> columna `whatsapp` (o `enlace` si la empresa da un link)

Si en un anuncio falta alguno: **NO INVENTAR NADA** -> dejo el campo vacío, aviso al
usuario y espero SU decisión. Sin campo completado, esa oferta no se publica
(se marca `visible=no` y queda en el Excel).

## REGLA DE VIGENCIA — decidida por el usuario 2026-09-15 (criterio ONPE añadido 2026-09-25)
- **FOTOS (`ofertas.xlsx`): 10 días desde la subida**. `script.js` NO los muestra
  (filtro `esReciente()`, `DIAS_MAX = 10`). `purgar.py` los BORRA del Excel:
  `python3 purgar.py` (usa columna `fecha`).
- **La fecha de subida manda (decidido por el usuario 2026-09-17)**: los 10 días se cuentan
  DESDE la fecha en que se sube la imagen (columna `fecha` = día de subida, no el día del
  cartel). Si hoy se suben 20 fotos, todas llevan `fecha` = fecha de HOY. Al registrar cada
  foto nueva en `ofertas.xlsx` SIEMPRE completar `fecha` con el día de subida (hoy).
- **CONVOCATORIAS ONPE (`convocatorias.xlsx`): se mantienen hasta su fecha LÍMITE**
  (decisión del usuario 2026-09-25). La regla de 10 días NO aplica a las convocatorias:
  - El scraper extrae la columna **`vigencia`** (ISO) parseando "Vigente: Hasta el
    DD/MM/YYYY" del contenido del feed (`extraer_vigencia()`). Vacía si la convocatoria
    dice "Hasta completar vacantes, según ODPE" (sin fecha concreta).
  - `script.js` (`esReciente()`): convocatoria con `vigencia` >= hoy sigue VISIBLE; con
    `vigencia` YA PASADA se oculta; SIN `vigencia` se queda mientras el scraping siga
    publicándola (re-correr `scraper_onpe.py` cada 2-3 días: lo que el feed deje de
    traer desaparece solo). NO pasan por `DIAS_MAX`.
  - **Chip de la tarjeta ONPE** (2026-09-25): en vez de la fecha de subida muestra
    "**Hasta el d mmm**" (fecha límite; tooltip "Vence el ...").
  - **Chips DE LAS CONVOCATORIAS: las DOS fechas (2026-09-28, pedido del usuario)**:
    `fechaChip()` en script.js ya NO se queda solo con `vigencia`. Toda convocatoria
    scrapeada (ONPE `Portaltrabajo` y Estado `Convocatorias`) muestra
    **"Publicado el d mmm · Hasta el d mmm"** = fecha de publicación (`fecha`, la que
    trae el scraping) + fecha límite (`vigencia`). Antes solo salía la de fin, y las
    41 de ONPE "hasta completar vacantes" (sin `vigencia`) NO mostraban ninguna fecha.
    - Sin `vigencia` -> "**Publicado el d mmm · Sin fecha límite**" (tooltip "... ·
      Sin fecha límite: hasta completar vacantes"). Con `vigencia` pero sin `fecha`
      (los 114 PUESTOS por carrera `Universitarios`/`Técnicos`, cuya ficha no trae
      fecha) -> solo "Hasta el d mmm".
    - Si algún día existiera la columna `inicio` (cronograma), tiene prioridad y sale
      el rango "Del d mmm al d mmm" (rama dormida: hoy ningún Excel tiene `inicio`).
    - En las **destacadas** (compacto) va "pub. 17 set. · hasta 27 set." / "pub. 8 jul. ·
      sin límite". El tooltip siempre lleva las fechas EXACTAS con año.
    - Las **FOTOS** no cambian: "Subido hoy / ayer / hace N días / Subido el d mmm".
    Helpers nuevos: `fechaCorta(iso)` ("15 sep", con año si no es el actual) y
    `fechaLarga(iso)` ("15 de setiembre de 2026").
  - **PURGA AUTOMÁTICA**: `scraper_onpe.py` llama a `purgar_por_vigencia()` SOLO al
    terminar de guardar `convocatorias.xlsx` (borra filas con vigencia ya pasada; no toca
    `ofertas.xlsx` ni las sin fecha). El usuario YA NO tiene que correr `purgar.py`
    manualmente después de cada scraping: ocurre solo.
- **CONVOCATORIAS DEL ESTADO (`convocatoriasdetrabajo.xlsx`): se tratan IGUAL que las
  ONPE** (decisión del usuario 2026-09-25): se mantienen hasta su fecha LÍMITE (columna
  `vigencia`), nunca pasan por `DIAS_MAX`, y `script.js` (ramas de `esReciente()` y
  `fechaChip()` para `_origen === 'Convocatorias'`) las oculta y muestra igual: chip
  "**Hasta el d mmm**" y botón **VER CONVOCATORIA** (enlace a la convocatoria original).
- **`purgar.py` hace cada regla en su archivo**: `purgar()` (10 días) para `ofertas.xlsx`
  y `purgar_por_vigencia()` para `convocatorias.xlsx` y `convocatoriasdetrabajo.xlsx`.
  `python3 purgar.py` limpia TODOS con su regla correspondiente.

## CURSOS GRATUITOS Capacita-T (MTPE) — 3ª fuente (decidida por el usuario 2026-09-15)
- **Origen**: `scraper_capacita.py` (scraping a `capacitacionlaboral.trabajo.gob.pe/cursos/`;
  el listado solo muestra ~12, el catálogo completo de **214 cursos** sale del sitemap
  `/wp-sitemap-posts-cursos-1.xml` con `--todo`). Genera `cursos.xlsx` (18 columnas: las 13
  de control + `duracion`, `certifica`, `segmento`, `empezar_curso`, `ver_ruta`) y `cursos.json`.
- **SEGMENTO = entidad que EMITE el curso** (decisión del usuario 2026-09-17): cada curso se
  agrupa según la institución que lo emite, deducida del LMS de `empezar_curso`
  (`DOMINIO_SEGMENTO` en `scraper_capacita.py`): `capacitate.trabajo.gob.pe`->MTPE,
  `lms.becasgruporomero.pe`->Fundación Romero, `www.netacad.com`->Cisco,
  `e.huawei.com`->Huawei, `www.viabcp.com`->ABC del BCP. En la página la sección `#cursos`
  muestra cada segmento con su CABECERA de color (verde MTPE, azul Cisco, rojo Huawei,
  naranja Fundación Romero, teal ABC del BCP) y su contador, con las tarjetas debajo.
- **FILTRO SUPERIOR por institución** (decisión del usuario 2026-09-17): los botones de
  color por segmento con contador ("Todos", MTPE, Cisco, Huawei, Fundación Romero, ABC del
  BCP) van en una barra ARRIBA de la lista de cursos (no en talón lateral); al pulsar uno
  la lista muestra SOLO ese segmento (combina con `buscadorCursos`). En móvil (≤900px) la
  barra se corre en horizontal.
- **FILTRO SUPERIOR por institución en TRABAJOS** (2026-09-25): `#filtroOrigen` ya NO es
  por origen (Fotos/Portaltrabajo/Convocatorias, quitado). Es un **select desplegable**
  (`filtroSelInstitucion`, clase `.select-institucion`) con el nombre de las instituciones
  (columna `empresa`, cada una con su contador). **Desde 2026-09-28 es OPTGROUP POR SEDE**
  (dado por el usuario: "en caso de las instituciones los juntaste, reparte por sedes pero
  en el mismo lugar osea que sea desplegable"): cada institución es un grupo
  (`<optgroup label="PODER JUDICIAL (6)">`) con la opción "Todas las sedes (N)" + UNA
  opción por SEDE/ubicación con su contador ("Lima (4)", "Arequipa (2)"); las de una sola
  sede van como opción directa. Las tarjetas sin institución (fotos) entran en el grupo
  "Sin institución" (marcador interno `INST_SIN`). El valor del `<option>` es
  `"institución"` (todas las sedes) o `"institución|sede"` (una sede); `filtrar()` combina
  institución + sede + TODAS/CON ENLACE/CON WHATSAPP/CON LLAMADA + chips de región +
  buscador. Cada tarjeta lleva `data-institucion` y `data-sede` (ubicación normalizada).
- **TARJETAS DESTACADAS con >2 botones** (2026-09-28, decisión del usuario): la sección
  "★ DESTACADA" ya NO son las 3 más recientes, sino las 3 más recientes ENTRE las ofertas
  con MÁS DE 2 botones accionables (`botonesDe(o)` en script.js: `funciones`/`postular`
  suman 1, `ver_detalles` y `guia_registro` suman 1 c/u, el respaldo `enlace` suma 1, y
  `whatsapp` suma 2 por LLAMAR+WHATSAPP). Razón del usuario: "en algunos casos no podemos
  postular, dice que podemos ver la convocatoria pero no nos sirve de nada si después no
  podemos postular" → una tarjeta que solo manda a "VER CONVOCATORIA" no se destaca.
- **Columnas mostradas**: col `duracion` (ej. "8 horas"), `certifica` (si/no) y DOBLE botón:
  EMPEZAR CURSO (`empezar_curso` = plataforma LMS del curso) + VER RUTA (`ver_ruta` =
  ruta formativa). La tarjeta NO lleva LLAMAR/WHATSAPP (los cursos no tienen contacto).
- **NO CADUCAN**: los cursos NO pasan por la REGLA DE VIGENCIA (10 días) ni por `purgar.py`
  — los anuncios de cursos se quedan siempre. En `script.js` el filtro es
  `esCurso(o) || esReciente(o)` (un curso NUNCA se oculta por antigüedad).
- **En la página** se muestran en su PROPIA sección (`#cursos` debajo de las ofertas/
  convocatorias), NO mezclados con los anuncios.
- El sello/logotipo del curso manda igual que en el SELLO DE EMPRESA: ver AGENTS.md → SELLOS.

## Cómo leer los anuncios (metodología del usuario — aplicar SIEMPRE)
Por lo general, en las imágenes de ofertas el orden es fijo:

1. **PRIMER TÍTULO = `titulo`**: el texto grande/inicial del cartel es el puesto (ej.
   "COCINERO (BÁSICO)", "SE NECESITA MESERO"). No copiar frases como "ANUNCIA CON NOSOTROS".
2. **Texto siguiente = tipo de trabajo** (jornada, modalidad, horario, sueldo, referencia).
3. **Después van los requisitos** (experiencia, documentos, disponibilidad).
4. **Si hay números, mirar el TEXTO QUE VA ANTES DEL NÚMERO**:
   - "CEL:" (o "CEL 9xx...") antes del número -> **trato por llamada** -> `contacto=llamada`.
   - "WhatsApp", "SOLO WHATSAPP", "Consultas WhatsApp" o icono verde de WhatsApp ->
     **WhatsApp** -> `contacto=whatsapp`.
   - "20 x 10", "14x7" -> jornada, NO es teléfono.
5. **Personas/fotos decorativas del cartel -> IGNORARLAS**: trabajar SOLO con los textos.
   No inventar datos a partir de la imagen.

Esto vale para las FOTOS de la Agencia de Publicidad (Espinar) y los anuncios sueltos.
Aprender a leer el icono/estilo del número en la foto:
- Si el número va como **"CEL:"** (o "CEL 9xx...") **sin icono de WhatsApp** -> el trato es
  **POR LLAMADA** -> columna `contacto = llamada` -> la página muestra botón **LLAMAR**
  (`tel:+51<numero>`).
- Si el anuncio dice **"SOLO WHATSAPP"**, "WhatsApp", "Consultas WhatsApp" o el número
  **lleva el icono verde de WhatsApp** -> columna `contacto = whatsapp` (o vacío) ->
  botón **WHATSAPP** (`wa.me/51<numero>?text=...`).
- **IMPORTANTE PARA EL OCR:** si el número está **resaltado en amarillo** en la foto, ese
  resaltado es la pista de que es el número de contacto; el OCR a veces lo pierde -> hay
  que detectar la banda amarilla (píxeles r>150, g>150, b<130) y re-OCR esa zona con umbral.

### SELLOS DE EMPRESA / CARTELES CON LOGOTIPO — EXCEPCIÓN (cuidado extra)
Algunos carteles de **empresa/agencia** llevan un **SELLO o logotipo** con el nombre real
(leído mal o perdido por el OCR), distinto del texto del cuerpo. Caso real registrado:
imagen `WhatsApp Image 2026-09-15 at 18.46.16.jpeg` (operadores de cisterna/retroexcavadora
de la mina TINTO Ccapmarca): el psm 6/3 daba "TINTO", pero el **sello** del cartel decía
**"EMPRESA APU LLALLAWA"**. Se corrigió empresa=TINTO -> **Apu Llallawa** (fila 14).

**Regla para esta clase de imágenes (SIEMPRE):**
1. Correr OCR en **VARIOS modos**: `python3 ocr_multi.py "images/<archivo>"`.
   (Ejecuta tesseract con PSM 3, 6, 11 y 4 y los compara.)
2. Si la imagen tiene logotipo/sello, no confiar en el primer resultado: buscar en TODOS
   los modos la palabra "EMPRESA"/"S.A.C"/"CONTRATISTA"/"CONSTRUCTORA" (los sellos suelen
   decir "EMPRESA XXXX" o "<NOMBRE> S.A.C").
3. El nombre del SELLO manda sobre el texto del cuerpo cuando es claro. Si el OCR sigue
   ambiguo, preguntar al usuario.
4. Esto aplica también al trasladarlo al Excel: usar el nombre real que diga el sello.

## Decisiones ya tomadas por el usuario (NO volver a preguntar)
- Anuncios de la **Agencia de Publicidad** (formato "ANUNCIA CON NOSOTROS:
  988850241 - 942668911") sin ciudad -> **Espinar - Cusco**.
- "Repartidor de gas en moto lineal" -> **Espinar - Cusco**.
- "**20 x 10**" = jornada (20 días trabajando, 10 de descanso). NO es un teléfono.
- Ubicación unificada "Espinar - Cusco" (agencia y Antapaccay).
- **REGIONES REALES (solo departamentos)**: las "regiones" de la página (banner,
  estadística, chips y footer) son SIEMPRE departamentos reales del Perú (Cusco,
  Arequipa, Lima, La Libertad, Lambayeque, Ica...). NUNCA una ciudad suelta ni una
  calle. CUIDADO: hay avenidas llamadas "Av. Lima" que NO son la región Lima (pueden
  estar dentro de Espinar). Al llenar `ubicacion` escribir
  "Distrito/Provincia - DEPARTAMENTO" (ej. "Espinar - Cusco") o directamente el
  departamento; si la foto solo da una dirección de calle, deducir la región real por
  contexto y no anotar el nombre de la avenida. En `script.js`, `regionDe()` normaliza
  la ubicación a departamento (ignora calles, "Trujillo"->La Libertad, "Chiclayo"->
  Lambayeque). Opcionalmente la hoja puede llevar columna `region` (departamento) que
  manda sobre la deducción automática.
- El texto "ANUNCIA CON NOSOTROS" NO va dentro de las ofertas.
- Minera (TINTO/operadores) Ccapmarca -> **Ccapmarca - Cusco** (no Espinar).
- Minera P/AKIM que pide Antapaccay aunque el pie diga "Cerro Colorado, Arequipa" ->
  **Espinar - Cusco** (decisión del usuario 2026-09-15).
- Convocatorias de campaña del cartel "Choque de Negocios / UNIDOS JUNTOS" (MiniMarket,
  casa, ayudantes de cocina) sin ciudad -> **Espinar - Cusco** (agencia).

## Scraping (convocatorias.xlsx)
- `scraper_onpe.py` (Python: requests + openpyxl). Uso:
  - `python3 scraper_onpe.py` -> últimas 150 recientes.
  - `--max N` -> últimas N.  `--todo` -> todas (767+). `--verificar` -> chequea los enlaces.
- Fuente: feed JSON de Blogger `portaltrabajos.pe/feeds/posts/default/-/ONPE`
  (el feed trae título, fecha, enlace y contenido con Institución/Vacantes/Ubicación/
  Salario/Vigencia/PUESTO estructurados). No baja las 767 páginas una a una.
- Los enlaces son los `rel="alternate"` del feed (OFICIALES, no inventados).
  `portaltrabajo.org` quedó como índice sin listados -> no usar.
- **IMPORTANTE (PDFs de los anuncios):** cada entrada del feed trae en su HTML los links
  que en la web se ven como `[ VER MÁS DETALLES ]`, `[ FUNCIONES ]`, `[ POSTULAR ]` y
  `[ VER GUÍA DE REGISTRO ]`. El scraper los extrae con `extraer_enlaces()` LEYENDO EL HTML
  CRUDO (no debe pasarse por `a_plaintext` antes, o se pierden los href) y los guarda en
  columnas propias: `ver_detalles` (PDF Google Drive de la base), `funciones` (PDF),
  `guia_registro` (PDF), `postular` (sistema reclutamiento.onpe.gob.pe).
- Esos PDF van directo en los botones FUNCIONES / POSTULAR de la página (los detalles
  ya se ven en la tarjeta, por eso no hay VER OFERTA ni VER DETALLES).
- Regla de oro aplicada: convocatorias SIN `ubicacion` -> `visible=no`.
- `--verificar` -> oculta enlaces caídos (status!=200, dominio raro, página <5KB).
  Ojo: este template de Blogger sirve `<data:...>` sin renderizar incluso en posts vivos;
  NO usar eso como señal de página muerta.
- Ubicaciones "Nivel Nacional - (Según ODPE Disponible)" -> "Nivel Nacional".
- Genera además `convocatorias_onpe.json` (respaldo).

## Scraping del Estado (convocatoriasdetrabajo.xlsx) — 2026-09-25
- **Origen**: `scraper_convocatoriasdetrabajo.py` (requests + bs4 + openpyxl; sin pandas,
  porque no está instalado). Scrapea `www.convocatoriasdetrabajo.com` (empleos del Estado:
  CAS, 728, 276, prácticas, etc. de cualquier institución). Uso:
  - `python3 scraper_convocatoriasdetrabajo.py`        -> TODAS las vigentes de la portada.
  - `--max N` -> limita los detalles a N (solo pruebas).
- **Detecta SOLO las vigentes de la PORTADA** (2026-09-28, el sitio cambió el marcador):
  cada `article.convocatoria` trae un pill **"Finaliza en N días" / "Finaliza hoy"**
  (`dias_de_pill()`); de ahí se calcula `vigencia` = hoy + N días. Se guarda además el
  patrón viejo "Vigente hasta el DD/MM/AAAA" como fallback. Se deduplican y se salta el
  detalle de las que ya pasaron en el propio listado. La portada publica ~270 vigentes
  (una tarea periódica de 2-3 días la mantiene sola: lo que deje de estar en la portada
  desaparece). Por defecto NO hay tope
  (decisión del usuario 2026-09-25: quería ver TODAS las vigentes, no solo 150).
- **TRABAJOS POR CARRERA** (2026-09-28, pedido del usuario "aparece trabajos por carrera"):
  además de la portada, scrapea los listados `carreras-profesionales-universitarias.php`
  (~4 páginas) y `carreras-profesionales-tecnicas.php` (~2 páginas), cada tarjeta es un
  PUESTO por código CAS/item (slug `oportunidad-laboral-...`). `extraer_activas()` recibe
  el `slug` esperado en cada página; se marcan con la columna **`categoria`**
  (`Universitarios`/`Técnicos`; portada = vacía). Las páginas de puesto que ya 404
  (posición cubierta entre listado y detalle) se descartan solas.
- OJO: el sitio muestra un total de "3000+" empleos porque cuenta CADA PUESTO por separado
  (incluye muchísimas finalizadas); solo se capturan los VIGENTES (con pill).
- **Una tarjeta por convocatoria/puesto**: filas con las 18 columnas (`id, titulo, empresa,
  ubicacion, sueldo, descripcion, enlace, whatsapp, fecha, fuente, destacado, visible,
  ver_detalles, funciones, guia_registro, postular, vigencia, categoria`).
- **El detalle tiene DOS plantillas** (el sitio rediseñó la ficha 2026-09-28):
  - CONVOCATORIA (`oferta-de-empleo-...`): panel "Datos clave" con `div.dato` /
    `.dato__label` / `.dato__value` (campos SIN dos puntos) -> `extraer_datos_clave()`.
    `empresa`=Institución, `sueldo`=Remuneración, `fecha`=`fecha_corta_a_iso()` ("25 sept
    2026"), `vigencia`="Finaliza: dd de MMMM del AAAA" (ISO).
  - PUESTO (`oportunidad-laboral-...`): NO tiene ese panel; usa el bloque "Claves del
    puesto" con el VALOR ENCIMA de la etiqueta -> `extraer_claves_puesto()` ("Ucayali\n
    Lugar de labores", "S/ 240 / Remuneración", "4 oct 2026 / Plazo para postular").
    La institución sale del BREADCRUMB ("INSTITUCIÓN > Convocatoria ...") ->
    `extraer_breadcrumb()`; el `titulo` limpia el prefijo "Convocatoria {INST} ".
  - `ubicacion`="Lugar de labores" (departamentos reales), `descripcion` = resumen
    (contrato, "Hay plazas para", "Cantidad de plazas", remuneración, puestos destacados
    como "CAS N° 015 - ITEM N° 01: MADRE SUSTITUTA").
- **ENLACES DIRECTOS DE LA FICHA** (2026-09-28, pedido del usuario: NO llevar a
  convocatoriasdetrabajo.com sino directo al trámite): `extraer_enlaces()` clasifica los
  links que la propia convocatoria publica y los guarda en las columnas de enlace:
  `postular` = botón "POSTULA AQUÍ"/"INSCRÍBETE"/"APLICA" (portales de la institución,
  ej. `aplicativo.pj.gob.pe`, `reclutamiento.onpe.gob.pe`), `ver_detalles` = "Ver aquí
  Bases (convocatoria completa y cronograma)"/anexos (PDF en Google Drive o repositorio
  institucional), `guia_registro` = "(VIDEO) Cómo postular" (tutorial YouTube). Se ignoran
  los links internos del sitio y texto genérico del menú. Muchas convocatorias NO traen
  ningún enlace directo (solo texto) -> esas columnas quedan vacías.
- Si `titulo` o `ubicacion` vienen vacíos (páginas compilación sin datos), la fila sale
  `visible=no`. Las que dicen "CONVOCATORIA FINALIZADA" no se publican.
- **PURGA AUTOMÁTICA**: al terminar llama a `purgar_por_vigencia()` sobre su propio xlsx
  (borra las con vigencia ya pasada). `purgar.py` también lo incluye en `purgar_todo()`.
- Genera respaldo `convocatoriasdetrabajo.json`.
- En la página (2026-09-28): botones DIRECTOS por tarjeta de "Convocatorias":
  **POSTULAR** (`botonPostular()` -> `postular`), **VER BASES** (`botonBases()` ->
  `ver_detalles`, clase `.boton-bases` #4f46e5) y **CÓMO POSTULAR** (`botonVideo()` ->
  `guia_registro` video, clase `.boton-video` #dc2626). **VER CONVOCATORIA**
  (`botonConvocatoria()`) SOLO queda de respaldo cuando la convocatoria no trae ningún
  enlace directo (lleva a la ficha original). Etiqueta **"Convocatorias"** (color
  #0284c7), chip "Hasta el d mmm" y CON ENLACE cuenta con `esEnlazada()` (enlace de esta
  fuente ya cuenta como botón web).
- **SOLO SE PUBLICAN LAS QUE SE PUEDEN POSTULAR** (decisión del usuario 2026-09-28:
  "de qué me sirve que los vean si no pueden postular"): `puedePostular()` en script.js
  oculta de la página las convocatorias del Estado que NO traen ningún enlace directo
  (solo el respaldo VER CONVOCATORIA, columna `postular`/`ver_detalles`/`guia_registro`
  vacías). Era ~196 de 372 → quedan ~176 visibles. Fotos (se contacta por
  llamada/whatsapp) y ONPE (su `postular` siempre lleno) no se tocan. El Excel se
  conserva íntegro y se regenera igual; solo es un filtro de visibilidad de la página.
- En `fotoRubro()` las Convocatorias NO pasan por la rama `onpe`: caen en el mapeo por
  palabras clave de `FOTOS_RUBROS` (foto del rubro del puesto, ej. "médicos"->salud,
  "conductores"->chofer, "prácticas de oficina"->oficina).

## Conceptos importantes (lecciones aprendidas — aplicar SIEMPRE)
- **NUNCA comparar horas sueltas con fechas "solo-día"**: `new Date("2026-09-17")` se
  parsea como MEDIANOCHE UTC (por ser formato ISO de fecha sin hora), mientras que
  `new Date()` trae la hora local del visitante. En Perú (UTC-5), a partir del mediodía
  local la diferencia con la medianoche UTC ya "cruza" al día siguiente y el chip salía
  como "Subido ayer" para ofertas del MISMO día.
- **Solución (chip de fecha)**: comparar DÍAS CALENDARIO, no horas. En `script.js`, el
  helper `diasDesdeHoy(iso)` convierte la fecha a medianoche LOCAL
  (`new Date(iso + 'T00:00:00')`) y normaliza "hoy" a medianoche local también
  (`hoy.setHours(0,0,0,0)`); con eso, `Math.round((hoy - f)/86400000)` = 0 para lo
  subido hoy a cualquier hora, 1="ayer", 2..6="hace N días", >=7 fecha larga.
  `fechaSubida()`, `fechaSubidaCorta()` y la rama de FOTOS de `esReciente()` (regla de 10
  días) usan `diasDesdeHoy()`. OJO: `Math.round` sobre `Date.now()-fecha` SIN normalizar "hoy" a
  medianoche local vuelve a fallar en las tardes/noches.
- **Anti-duplicados**: ver regla exacta en ANTI-DUPLICADOS. La página solo oculta cuando
  TODO se repite (`claveOferta` = titulo + número + ubicacion + empresa); si cambia algo,
  se muestra y el aviso lo da `verificar_duplicados.py --ver "PUESTO" "LUGAR" "NUMERO"`.
- **SELLO de empresa manda**: ver SELLOS. En dudas de nombre real (sello/logotipo),
  correr `ocr_multi.py` y confirmar con el usuario.

## DISEÑO / GIT (2026-09-18)
- **GIT conectado**: repo GitHub `NelAngel/LYGANDO-TRABAJO` (remote `origin`),
  rama `main`. Se subió con `.gitignore` que ignora `__pycache__/`, `*.pyc` y
  `.~lock.*#` (bloqueos de LibreOffice); esos archivos ya se quitaron del tracking
  (`git rm --cached`). Al cambiar algo importante, recordar: `git add -A && git commit`
  y `git push` (SOLO si el usuario lo pide).
- **La página entra SIEMPRE por arriba** (decisión del usuario 2026-09-18): al cargar,
  `script.js` fuerza scroll al inicio (`scrollRestoration='manual'` + `scrollTo(0,0)` al
  inicio y al final, y `history.replaceState` limpia el ancla). Sirve para el caso "lo
  comparto por git y a la otra persona le aparecía 'Cómo usar la página' en vez del
  inicio" (el navegador restauraba el scroll de la visita anterior).
- **Fondo de la página BLANCO** (`--fondo: #faf8f6`). Se probó naranja Ubuntu de fondo
  (2026-09-18) y al usuario NO le gustó: se revirtió. El color va en las TARJETAS, no en
  el fondo.
- **TARJETAS naranja Ubuntu con estilo** (2026-09-18): `.tarjeta` (trabajos y cursos,
  los cursos usan `class="tarjeta tarjeta-curso"`) tienen fondo en GRADIENTE naranja
  Ubuntu (`linear-gradient(165deg,#f5713a,#e95420,#d4400e)`), borde `#c03a0a`, sombra
  naranja, línea superior brillante (`::before`) y hover que las eleva. `destacado-tarjeta`
  igual + brillo radial superior. Texto de la tarjeta en BLANCO (`titulo` con sombra,
  `meta`/`descripcion` blanco translúcido); el pin "DESTACADO" (`dest-pin`) y el número
  (`desta-rango`) pasaron a píldora BLANCA con texto naranja para resaltar.
- **Badge "Certificado" de los cursos** (2026-09-18): antes salía verde apagado; ahora
  `.tipo-certificado` es una píldora DORADA con brillo (gradiente `#ffd25e→#d18a0c`),
  estrellita `✦` antes del texto y un DESTELLO que recorre el botón cada ~2.8s
  (`@keyframes centellaCertificado`). El `.categoria` verde se quita cuando hay
  certificado (`.cur .categoria:has(.tipo-certificado) { background: transparent }`).
  OJO: el CSS viejo `.cur .badge-certificado` se ELIMINÓ (nunca se usaba; el JS emite
  `tipo-certificado`), no reintroducirlo.
- **TIPOGRAFÍA Ubuntu** (2026-09-18): la página usa la fuente **Ubuntu** de Google Fonts
  (`<link>` en el `<head>` de `index.html`, weights 300-700, fallback Segoe UI/Arial).
  Títulos con `font-weight:700`, `letter-spacing` y algunos en mayúsculas (h1 del hero,
  títulos de tarjeta, cabeceras de segmento) para un look más ordenado; cuerpo con
  `line-height` ~1.5-1.6. Todos los `font-weight:800` se cambiaron a `700` (Ubuntu no
  tiene el 800). Si el visitante está sin internet, cae a la fuente del sistema.
- **Título dinámico por pestaña** (2026-09-18): al pulsar CURSOS, `activarPestana()` en
  `script.js` cambia `document.title` a "LYGANDO CURSOS - Cursos gratuitos con certificado
  del Perú" y el h1 del hero a "LYGANDO CURSOS"; al volver a TRABAJOS se restaura.
- **Texto justificado** (2026-09-18): los párrafos (`.hero p`, `.tarjeta .descripcion`,
  `.destacado-tarjeta .entradilla`, `.como-aplicar p`, `.paso p`, `.piefooter p`,
  `.vacio/.cargando`) van con `text-align: justify` + `hyphens: auto` para que no queden
  desalineados.

## FOTOS ILUSTRATIVAS EN LAS TARJETAS (2026-09-24, pedido del usuario)
- Cada tarjeta de TRABAJO muestra ARRIBA una foto genérica del RUBRO del puesto
  (si pide cocinero -> cocina; minero -> mina; venta de tienda -> tienda...). No son las
  fotos de `images/` (esa carpeta sigue siendo SOLO fuente OCR): son 60 imágenes Unsplash
  libres de regalías descargadas en la carpeta local **`fotos/`** para que funcionen sin
  internet (como la fuente Ubuntu con fallback): **4 fotos por rubro** (archivo base
  `cocina.jpg` + `cocina-2/-3/-4.jpg`, etc.) para que las tarjetas del mismo rubro no
  repitan siempre la misma imagen.
- `FOTOS_ARCHIVOS` mapea cada rubro a su lista de 4 archivos; `fotoAzar(rubro)` elige uno
  AL AZAR evitando repetir la foto inmediatamente anterior del mismo rubro
  (registro por rubro en `FOTO_ULTIMA`), así dos tarjetas seguidas de cocina muestran
  fotos distintas. `fotoRubro(o)` devuelve `'fotos/' + fotoAzar(rubro)`.
- `fotoRubro(o)` en `script.js` elige la foto buscando claves en `titulo + descripcion`
  DESDE EL INICIO de una palabra (regex `\b`+clave SIN límite final): así "inventario" NO
  activa la clave "venta" (no está al inicio de palabra), pero "telecomunicaciones"
  activa "telecomunica" y "oficiales" activa "oficial". Manda el primer rubro que coincide.
- Mapeo `FOTOS_RUBROS` (orden importa): cocina/cocinero/pasteleria/vajillero->cocina;
  mozo/mesero/restaurante/chifa/discobar/tragos->restaurante; hotel/habitaciones->hotel
  (OJO: NO usar clave "hospedaje": un anuncio de mina que dice "cubrimos hospedaje"
  daría hotel); telecomunica/celular/laptop/lan center->telecom; dental/odontolog/
  clinica/enfermera->salud; tienda/venta/minimarket/bazar/libreria/perfumeria/optica->
  tienda; limpieza/lavanderia/ninera/casa particular->limpieza; mina/minero/perforista/
  antapaccay/tintaya/ccapmarca/retroexcavadora/cisterna/lampero->mina; construccion/
  oficial/encofrador/fierrero/obrero/vidrieria->construccion; moto/delivery->moto;
  chofer/conductor/camioneta/semitrailer->chofer; mecanico/automotriz/taller->mecanico;
  almacen/picking/inventario/despacho->almacen; practicante/oficina/logistica/gerente->
  oficina; seguridad/hse/ssoma/vigilante->construccion (supervisores de seguridad van
  con casco); sin coincidencia -> `trabajo.jpg` (genérica, NO tiene -2/-3/-4).
- **FOTOS PROPIAS PARA LAS CONVOCATORIAS ONPE** (2026-09-25): `fotoRubro()` tiene una
  rama que antepone el rubro `onpe` cuando `_origen === 'Portaltrabajo'`, así TODAS las
  convocatorias de portaltrabajos.pe muestran fotos de elecciones/votación (`fotos/onpe.jpg`
  + `onpe-2/-3/-4.jpg`, elegidas al azar con `fotoAzar('onpe')` sin repetir la anterior)
  y NUNCA la genérica `trabajo.jpg`. Las 4 vienen de Unsplash (voto en urna, votante en
  mesa, letrero "polling station", voto por correo).
- Inyección: `<img class="tarjeta-foto" src="${fotoRubro(o)}" alt="..." loading="lazy">`
  como PRIMER hijo del `<article>` en `crearTarjeta()` y `crearTarjetaDestacada()`.
  CSS: alto 150px (165px en la destacada), `object-fit:cover`, `border-radius:17px
  17px 0 0`, margen negativo para tocar los bordes (`-22px -22px 8px`; la destacada
  `-22px -20px` porque lleva padding 22/20) y `mask-image` que desvanece la base sobre
  el naranja de la tarjeta.
- **FOTOS EN LOS CURSOS por INSTITUCIÓN** (2026-09-24, pedido del usuario): como los
  trabajos, CADA tarjeta de curso muestra arriba una foto del SEGMENTO/institución que
  emite el curso (columna `segmento`): MTPE (capacitación/aula), Cisco (redes/servidores),
  Huawei (IA/circuitos), Fundación Romero (oficina/PC), ABC del BCP (ahorro/dinero).
  Son 20 imágenes más en `fotos/` (4 por institución: `mtpe.jpg` + `mtpe-2/-3/-4.jpg`,
  `cisco.*`, `huawei.*`, `romero.*`, `bcp.*`). `FOTOS_ARCHIVOS_SEGMENTOS` mapea segmento ->
  sus 4 archivos; `fotoSegmento(seg)` elige al azar sin repetir la inmediatamente anterior
  (registro por segmento en `FOTO_ULTIMA_SEG`), igual que `fotoAzar`. Segmento desconocido
  ("Otro") cae en la genérica `trabajo.jpg`. Se inyecta como PRIMER hijo del `<article>`
  en `crearTarjetaCurso()` (misma `.tarjeta-foto`, mismo CSS, sin cambios de style).

## Estado actual (snapshot 2026-09-18 — 17 imágenes nuevas subidas)

**17 imágenes nuevas del 18/09 en `images/`, de las cuales 2 eran DUPLICADOS SHA-1 de
fotos ya procesadas** (se IGNORARON, no se registraron):
- `08.57.59` == `08.55.44` (ya registrada) -> IGNORAR
- `08.58.13` == `4aa46031` (ya registrada) -> IGNORAR

**14 ofertas nuevas registradas en `ofertas.xlsx` (ids 60-73, todas `fecha`=2026-09-18,
visible=si, Espinar - Cusco salvo El Descanso):**
60. Personal para Unidad Minera Antapaccay — "Sé parte de grandes retos" — WA 994433777
61. Vendedora y apoyo en tienda de cerámicas (Av. San Martín 507) — llamada 996006035
62. Maestro perforista (cobre) y 2 peones — llamada 932875727 (decisión usuario: vacante
    nueva aparte; id 59 Chilloroya sigue pendiente visible=no)
63. Personal de limpieza para hotel — llamada 959745906
64. Niñera para cuidar a dos niñas (S/1,700) — llamada 912882120
65. Jaladores para Espinar y Juliaca (S/1,000 quincenal / 2,000 mensual) — llamada 912882120
66. Atención al cliente tienda de telecomunicaciones (S/1,500 + incentivos) — llamada 912882120
67. Cocinero(a) para "El Descanso" a 40 min de Espinar (S/2,500 a más) — llamada 915229427
68. Personal de limpieza hotel medio tiempo + ayudante cocina — llamada 925610536
69. Atención librería - internet medio tiempo — llamada 956604147
70. Atención al cliente Agencia Shalom Encomiendas — WA 931759057
71. Conductores A-IIC y A-IIB transporte de personal — llamada 932384102
72. Mozos(as) y ayudante de cocina (restaurante, Plaza de Armas con Av. Arequipa; llamadas
    984329666 / WhatsApp 957249676) — WA 957249676
73. Mozas de línea y vajillero — concesionaria de alimentos (régimen 14x7, Espinar) — WA 959700225
**OMITIDA (no es trabajo):** `15.43.24` = "SE VENDE LOTE 250m² Barrio Sol de Yauri".

`ofertas.xlsx` queda con 73 filas (ids 1-73); **página con 75 tarjetas visibles** (71 fotos
+ 4 ONPE), sin duplicados ocultos (verificado con Node/SheetJS). Aviso anti-duplicados
nuevo: ids 64/65/66 comparten el n.º 912882120 (patrón agencia, como KANGYX 19/20/21) y
ids 7/60 son la misma convocatoria Antapaccay con distinto número (ambas se publican).
**NUEVAS FOTOS del 17/09 (ids 36-54, visibles), todas Espinar - Cusco salvo Chilloroya:**
36. Chofer de semitrailer lic. A-IIC — llamada 974203922
37. Oficiales y ayudantes de construcción civil — llamada 954051254
38. Conductor lic. A-IB y ayudantes — llamada 908649914
39. Atención en Lan Center (turno tarde) — WHATSAPP 997376957
40. Vajillero y ayudante de cocina (concesionario corredor minero, 20x10) — llamada 994964334
41. Asistenta dental (15:00-20:00) — llamada 973272462
42. Cocinero, ayudante de cocina y moza (concesionaria alimentos) — WHATSAPP 931823945
43. 2 ayudantes de pastelería — llamada 913951928
44. Lamperos para mina en Uchucarco — llamada 906691332
45. Venta de ropa — llamada 929041208
46. Cocinero(a) y moza para restaurante — llamada 928571840
47. Técnico ayudante mecánico automotriz — llamada 931212106
48. Atención en perfumería (medio tiempo) — llamada 984130344
49. Conductor camioneta/operario minero/supervisor telecom (INTELSI S.A.C., Antapaccay) — llamada 958763782
50. Moza chifa 7 Esquinas (14:00-22:00, S/900) — llamada 968189349
51. Conductores A-IIA/A-IIC transporte de personal — llamada 914124410
52. Retroexcavadora/HSE/ing. civil/operarios/encofrador/fierrero/conductor múltiple — WHATSAPP 973207341
53. Atención al público (tragos y piteados) — llamada 972495327
54. Venta de celulares y laptops — CANCELADA (visible=no): mismo n.º 961796226 que id 57, usuario eligió solo la doméstica
55. Mecánicos y técnicos parada setiembre (VICCES S.A.C., Antapaccay) — llamada 914623339
56. Oficiales y ayudantes civiles (n.º 976397152 = mismo que id 29; decisión usuario: publicar aparte) — llamada
57. Trabajo doméstico en casa particular (familia pequeña) — WHATSAPP 961796226
**PENDIENTES (sin número de contacto, en Excel con `visible=no`, REGLA DE ORO):**
58. Encargado de movimiento de tierras — SEGEL UMASI (Antapaccay) — **whatsapp 994636754 (ACTIVADO 17/09)**
59. Maestro perforista, 2 ayudantes y cocinera — mina artesanal Chilloroya - Cusco — falta n.º
**NO registradas de hoy:** 17.17.00 = "SE VENDE LOTE 270m²" (no es trabajo).
FOTOS anteriores y básicas/ejemplo: ids 1-35 (ver snapshot 16/09 abajo).

**Básicas/ejemplo (ids 1-6):** Vendedor tienda (Lima, enlace), Chofer de reparto (Arequipa, wa),
Cocinero/a (Cusco, wa+enlace), Personal de limpieza (Trujillo, wa), Ayudante de construcción
(Chiclayo, enlace), Operador de almacén (Ica, wa).

**Espinar - Cusco y minas (ids 7-18):**
7. Personal para Unidad Minera Antapaccay — wa 935913774 — destacado
8. Repartidor de gas en moto lineal — wa 961575864
9. Vendedor en tienda de cerámica — wa 982345114
10. Cajera y moza para discobar — wa 991268161
11. Operadores y conductores - Zona Sur — Arequipa - Condesuyos — wa 957085375 ("20 x 10")
12. Cocinero (básico) turno mañana — SOLO WHATSAPP 910 006 315 (Hotel Santa Isabel)
13. Operadores cisterna/retroexcavadora/conductores — Ccapmarca - Cusco — CEL 932 558 281 (Apu Llallawa)
14. Atención en minimarket (medio tiempo) — llamada 910966662
15. Ayudantes de cocina (2) — CEL 978 664 185 (llamada)
16. Personal de mantenimiento para minera Antapaccay — WhatsApp 973584106 (P/AKIM)
17. Ayudante en casa particular — llamada 984848183
18. Mesero o mesera para restaurante — CEL 910 388 449 (Qori Restaurant, llamada)

**NUEVAS 16-set (ids 19-35):**
19. Practicante automotriz — KANGYX — wa 958223633
20. Operario automotriz — KANGYX — wa 958223633
21. Asistente de tienda de repuestos — KANGYX — wa 958223633
22. Ayudante de cocina/moza/vajillero (Tintaya-Marquiri) — llamada 959203948
23. Moza o mozo — Choque de Negocios, llamada 921619323
24. Ayudante de cocina (Espinar, 8:45-15:00) — llamada 939669142
25. Lavandería medio tiempo — llamada 910779442
26. Limpieza de hotel — llamada 913045411
27. 2 ayudantes de cocina + 1 moza — llamada 925930823 (también 980546861)
28. Supervisor/conductor/operador volquete — Consorcio Byas — Chilloroya - Cusco — wa 920766777 (icono verde)
29. Supervisores seguridad Ing. HSE/SSOMA (Antapaccay) — wa 976397152
30. Atención en óptica (turno tarde, S/750) — llamada 966179195
31. 2 ayudantes cocina + 1 moza restaurante — llamada 961339157
32. Ayudantes para vidriería — llamada 984331868 (también 931747697, banda amarilla)
33. Obreros/perforistas/carreros — Corp. Minera Lunar de Oro E.I.R.L., Colquemarca - Cusco — wa 988561671 ("20 x 10")
34. Asistente dental (centro odontológico) — llamada 942340234
35. Ventas en Espinar — Distribuidora Molitalia — llamada 984615636

Omitida: `17.45.40` es "CASA EN VENTA" (no es trabajo). Duplicado: `08.55.18` = `14.54.45` (ya publicada).

SCRAPING en `convocatorias.xlsx` (2026-09-25): 46 convocatorias ONPE (41 "hasta completar
vacantes" sin fecha límite + 5 con vigencia futura), todas visibles. La purga por vigencia
borró 104 de la corrida anterior. El resto del feed queda en `convocatorias_onpe.json`
(respaldo) y vuelve al Excel solo al re-correr `scraper_onpe.py` (allí se vuelve a purgar
con la regla de vigencia).

SCRAPING en `convocatoriasdetrabajo.xlsx` (2026-09-28): **372 filas** (258 convocatorias de
la portada + 111 puestos Universitarios + 3 Técnicos), todas visibles y con `vigencia` >= hoy.
La portada publica ~270 vigentes con pill "Finaliza en N días"; al re-correr el scraper cada
2-3 días las que vencen desaparecen solas, y los puestos por carrera ya cubiertos (página
404 en el detalle) se descartan solos. Cobertura de enlaces directos: 133 con POSTULAR,
167 con VER BASES, 1 con CÓMO POSTULAR (video); el resto queda con el respaldo VER
CONVOCATORIA. El total "3000+" del sitio cuenta por PUESTO sin filtrar (incluye
finalizadas) y no se usa. LA PÁGINA muestra ~176 de esas 372 (solo las postulables,
ver `puedePostular()`); el Excel conserva las 372.

## Snapshot CURSOS 2026-09-17 (cambios de este día)
`scraper_capacita.py` ahora incluye la columna **`segmento`** = entidad que EMITE cada
curso, deducida del LMS de `empezar_curso` (`DOMINIO_SEGMENTO`). `cursos.xlsx`/`cursos.json`
regenerados con `--todo` (214 cursos). Reparto: Fundación Romero 163, MTPE 18, Cisco 18,
Huawei 14, ABC del BCP 1 (todos visibles, ninguno sin segmento). En la página:
- La sección `#cursos` agrupa las tarjetas por SEGMENTO con cabecera/título de color y
  contador (clases `seg-mtpe`, `seg-cisco`, `seg-huawei`, `seg-romero`, `seg-bcp`).
- **Filtro SUPERIOR** por institución en `#segmentosFiltros` (botones `segfiltro` de color
  con contador: Todos/MTPE/Cisco/Huawei/Fundación Romero/ABC del BCP). El usuario pidió que
  NO sea lateral/bajo sino ARRIBA (2026-09-17). En móvil ≤900px la barra corre en horizontal.
- **Filtro de instituciones en TRABAJOS**: `#filtroOrigen` es un **select desplegable**
  (`renderFiltroOrigen`) con "Todas las instituciones" + cada institución (columna
  `empresa`) con su contador, arriba de la lista de ofertas (2026-09-25, pedido del usuario:
  ya NO son botones de origen). Las tarjetas llevan `data-institucion={empresa normalizada}`
  y `filtrar()` combina institución + TODAS/CON ENLACE/CON WHATSAPP/CON LLAMADA + chips
  de región + buscador.
- Cada filtro consulta su propia barra (`#segmentosFiltros ... .segfiltro.activa` en cursos,
  `#filtroOrigen #filtroSelInstitucion` en trabajos) para no pisarse entre sí.
- Búsqueda de cursos ahora incluye `o.segmento` en el texto indexado (buscar "huawei" o
  "romero" encuentra esos cursos).

## Sistema / Detalles técnicos
- WhatsApp/LLAMADA: en la hoja va SOLO el número local 9 dígitos; `script.js`, cuando hay
  número, muestra SIEMPRE los DOS botones: **LLAMAR** (`tel:+51<numero>`) y **WHATSAPP**
  (`https://wa.me/51<numero>?text=Hola, me interesa el puesto de <titulo>`). La columna
  `contacto` (llamada/whatsapp/vacío) marca el canal principal para las etiquetas y filtros.
- `visible=no` oculta sin borrar; `destacado=si` marca la tarjeta principal.
- Búsqueda (`script.js`): insensible a tildes (normaliza NFD), TODAS las palabras deben
  coincidir y admite prefijos ("reparti" -> "repartidor").
- Columnas base de los Excels (13): id, titulo, empresa, ubicacion, sueldo, descripcion,
  enlace, whatsapp, fecha, fuente, destacado, visible, contacto. Hoja llamada "Ofertas".
  `contacto` = `llamada` (botón LLAMAR) o `whatsapp`/vacío (botón WHATSAPP).
- `convocatorias.xlsx` tiene 5 columnas extra (solo scraping): `ver_detalles`, `funciones`,
  `guia_registro`, `postular` (links a PDF de Google Drive y al sistema de ONPE) y
  `vigencia` (fecha límite ISO, vacía si es "hasta completar vacantes").
  `script.js` lee por NOMBRE de columna, así que `ofertas.xlsx` (sin extra) sigue funcionando.
- Botones por tarjeta: FUNCIONES (PDF), POSTULAR (ONPE), WHATSAPP o LLAMAR. La tarjeta ya
  muestra los detalles (título, descripción, sueldo), por eso NO hay botones VER OFERTA ni
  VER DETALLES. El filtro "CON ENLACE" cuenta cualquier botón web: `esEnlazada()`.
- Filtros de la página: TODAS / CON ENLACE / CON WHATSAPP / CON LLAMADA (`esLlamada()`).
- `fecha` numérica de Excel -> se convierte a ISO con `serieFecha()`.

## Flujo de trabajo
1. FOTOS: el usuario deja imágenes en la carpeta **`images/`** (así se mantiene el orden).
   OCR con `tesseract -l spa` leerá SIEMPRE las imágenes desde `images/`, aplicar REGLA
   DE ORO (preguntar si falta algo), agregar filas a `ofertas.xlsx` con openpyxl.
   (La página no muestra las fotos: las imágenes solo son fuente para el OCR.)
   **ANTES de registrar una imagen, verificar que NO esté ya publicada** (ver
   ANTI-DUPLICADOS): si la misma foto (aunque cambie el nombre) ya se procesó, no
   registrarla de nuevo.
2. SCRAPING: `python3 scraper_onpe.py` (regenera `convocatorias.xlsx` + JSON) y
   `python3 scraper_convocatoriasdetrabajo.py` (regenera `convocatoriasdetrabajo.xlsx` +
   JSON). Ambos purgan automáticamente al final las convocatorias cuya vigencia ya pasó
   (ver REGLA DE VIGENCIA). Correr cada 2-3 días: las "hasta completar vacantes" y las
   vigentes se quedan solas conforme caducan. No hace falta correr purga manual.
3. PURGAR fotos: `python3 purgar.py` (borra de `ofertas.xlsx` los anuncios con >10 días y
   de los dos Excels de convocatorias las de vigencia pasada — cada archivo con su regla).
4. Verificar con Node (mismo flujo SheetJS) y el usuario refresca Live Server.

## ANTI-DUPLICADOS (imágenes y datos repetidos) — decidido por el usuario 2026-09-17
- Al extraer la información de las fotos puede pasarnos que haya **imágenes repetidas**
  (mismo contenido con distinto nombre, p.ej. `WhatsApp Image 2026-09-14 at 14.54.45` ==
  `WhatsApp Image 2026-09-16 at 08.55.18`) y se registre la oferta dos veces.
- **REGLA**: si TODOS los datos se repiten (PUESTO + LUGAR + NÚMERO) -> **IGNORAR**, es
  la misma oferta: no se registra de nuevo. Si algo cambió (LUGAR, PUESTO o NÚMERO) ->
  **AVISAR** al usuario para que decida si es la misma foto corregida o una vacante nueva.
- **`python3 verificar_duplicados.py`** detecta y clasifica por esa regla:
  - imágenes repetidas en `images/` (mismo contenido por SHA-1, aunque cambie el nombre)
    -> IGNORAR (registrar una sola vez);
  - **MISMA FOTO re-enviada/re-comprimida**: comparación de PÍXELES sobre miniatura 64x64
    en grises (no pHash: las plantillas de la agencia comparten layout y pHash confunde
    todo). Calibrado con los carteles reales: el duplicado conocido `08.55.18`==`14.54.45`
    da 0.0% y los carteles DISTINTOS de la misma agencia parten de >=10.4% de diferencia.
    Umbrales: `<=6%` -> MISMA FOTO (IGNORAR, aunque el archivo cambió de SHA-1 por
    re-compresión), `6-10%` -> PARECIDA (REVISAR), `>=10%` -> distinta. Requiere Pillow
    (`pip install pillow`); sin Pillow solo funciona el SHA-1.
  - filas duplicadas EXACTAS en `ofertas.xlsx` (mismo puesto+lugar+número) -> IGNORAR;
  - ofertas PARECIDAS con datos cambiados (mismo número pero otro puesto/lugar, o mismo
    puesto+lugar pero otro número) -> AVISO.
  **Escanea AMBAS fuentes** (`ofertas.xlsx` + `convocatorias.xlsx`): una foto que repita
  una convocatoria ONPE (mismo puesto+lugar+número) también sale como IGNORAR, y los avisos
  indican la fuente (`id 1 (PortalTrabajo)`). Con `--filtrar` solo marca en `ofertas.xlsx`
  (nunca toca `convocatorias.xlsx`, lo regenera el scraper).
  **Números normalizados en la comparación** (en `verificar_duplicados.py` y en
  `claveOferta()`/`normalizarNumero()` de `script.js`): se quita el código de país `51`
  y un `0` inicial -> `997...`, `+51 997...` y `0997...` se tratan como el MISMO número
  y la misma oferta no se publica dos veces en la página.
  Con `--filtrar` marca `visible=no` los duplicados exactos (conserva el reciente).
  Con `--ver "PUESTO" "LUGAR" "NUMERO"` compara una oferta recién leída con OCR contra lo
  ya publicado (responde IGNORAR / AVISO con el id encontrado / OK).
- **La página no repite**: `script.js` filtra con `dedupeOfertas()` / `claveOferta()`
  (puesto + número + lugar + empresa). Solo oculta la tarjeta cuando TODO se repite; si
  cambia el lugar, el puesto o el número, la tarjeta se muestra (esa separación de
  "parecidas" la decide el usuario con el aviso del script).
- Regla práctica: al leer una foto nueva, si coincide con una ya registrada -> no
  registrar; si coincide en parte pero cambió algo -> avisar al usuario con `--ver`.

## Archivos de la carpeta
`index.html`, `script.js`, `ofertas.xlsx`, `convocatorias.xlsx`, `convocatorias_onpe.json`,
`convocatoriasdetrabajo.xlsx`, `convocatoriasdetrabajo.json`, `cursos.xlsx`, `cursos.json`,
`xlsx.full.min.js`, `scraper_onpe.py`, `scraper_convocatoriasdetrabajo.py`,
`scraper_capacita.py`, `purgar.py`, `verificar_duplicados.py`, la carpeta **`images/`**
con los JPEG de los anuncios (ya procesadas por OCR) y la carpeta **`fotos/`** con las 84
imágenes ilustrativas (60 de rubro de trabajos + 20 de institución de cursos + 4 de votación
para las convocatorias ONPE; 4 por rubro/segmento, ver FOTOS ILUSTRATIVAS).

## CURSOS GRATUITOS CAPACITA-T (MTPE) — 3ª FUENTE (decidida por el usuario 2026-09-16)
- **Scraper**: `scraper_capacita.py`. Origen `capacitacionlaboral.trabajo.gob.pe/cursos/`.
  **El listado `/cursos/` solo muestra ~12 cursos** (y `?pg=N` repite la misma página; los
  `/cursos/page/N/` del HTML dan 404). El **catálogo COMPLETO son 214 cursos** y sale del
  **sitemap de WordPress**: `/wp-sitemap-posts-cursos-1.xml`.
  `python3 scraper_capacita.py --todo` lee ese sitemap y scrapea las 214 fichas (una a una,
  con `time.sleep(0.15)` de cortesía). Sin `--todo` solo baja el listado de portada (~12).
- **Excel generado**: `cursos.xlsx`, hoja "Cursos", **18 columnas** = las 13 de control
  (`id, titulo, empresa, ubicacion, sueldo, descripcion, enlace, whatsapp, fecha, fuente,
  destacado, visible, contacto`) + 5 de curso: `duracion`, `certifica` (si/no),
  `segmento` (entidad emisora: MTPE/Cisco/Huawei/Fundación Romero/ABC del BCP),
  `empezar_curso` (URL plataforma LMS), `ver_ruta` (URL ruta formativa). Respaldos:
  `cursos.json` + `cursos.xlsx`.
- **Se muestran APARTE en la página**: sección `#cursos` con tarjeta verde (badge
  "CON CERTIFICADO", botones **EMPEZAR CURSO** y **VER RUTA**), **agrupados por
  SEGMENTO** (cabecera de color + contador por entidad emisora; ver AGENTS.md → SEGMENTO).
  Separada de los trabajos por una **FRANJA GIGANTE** naranja→verde ("TRABAJOS ·vs·
  CURSOS con certificado") — decisión del usuario 2026-09-16 para que se note al instante.
- **NO CADUCAN (se quedan SIEMPRE)**: los cursos NO pasan por `esReciente()` (regla de 10
  días) ni por `purgar.py`/purga automática. En `script.js` el filtro es
  `esCurso(o) || esReciente(o)`. Si `titulo` está vacío en el listado, `visible=no`.
- Cursos con `certifica=si` muestran el badge verde; `duracion` (ej. "8 horas") se lee de
  su columna. Los botones whatsapp/llamada NO aplican a cursos (no tienen contacto).