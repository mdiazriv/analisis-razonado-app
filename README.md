# Generador de Análisis Razonado — Hortifrut

App local (no sube nada a internet) que toma el Excel trimestral de datos financieros y el
Word del Análisis Razonado, y devuelve un Word actualizado: con las tablas (imágenes) y gran
parte del texto narrativo regenerados automáticamente a partir del Excel.

**Novedades de esta versión:** las tablas de Indicadores de Actividad y de Rentabilidad ya se
actualizan solas (antes quedaban con los valores del trimestre anterior porque la hoja de la que
se leían en el Excel cambió de nombre/estructura y el renderizador no se había actualizado para
seguirla); además, un barrido final corrige fechas/etiquetas de período sueltas en todo el
documento (ver "Qué se automatiza" abajo) y varias frases de causa-efecto se reconocen mejor en
los informes reales de septiembre y diciembre.

**Últimos ajustes (reportados por Melissa sobre un informe real generado con la app):**
- **Descarga en inglés:** al hacer clic en "Descargar Excel en inglés" (o en "Descargar Word en
  inglés"), el otro botón de descarga ya no desaparece. Antes ambos vivían dentro del bloque que
  solo se vuelve a dibujar en el mismo clic que generó los archivos; como Streamlit vuelve a
  correr toda la app en cada clic (incluido el de un botón de descarga), al hacer clic en uno el
  otro quedaba fuera de ese bloque y no se volvía a mostrar.
- **Exposición Cambiaria (letra muy chica):** esta tabla es más ancha que el ancho útil de la
  página (para que no se corten las etiquetas largas de cada moneda), así que Word la achicaba
  completa —letra incluida— al insertarla. Ahora la app calcula el ancho útil de tu página y
  agranda la letra de antemano en esa misma proporción, para que, ya achicada por Word, quede del
  tamaño pensado en vez de diminuta.
- **Tabla de Riesgos:** el texto de "Descripción del Riesgo" y "Controles Implementados" ahora se
  justifica (ambos márgenes parejos). Además, cada fila mide solo lo que su texto necesita (antes
  tenía una altura fija generosa pensada para el peor caso, así que con descripciones cortas
  quedaba con mucho espacio en blanco y la tabla se veía enorme) y, si algún trimestre la hoja
  'Riesgos' trae una fila reservada sin contenido, no se dibuja.
- **Flujo de Efectivo:** ahora SÍ se actualiza sola (hoja `FdeC` del Excel). A diferencia de las
  demás tablas (que son imágenes), esta es una tabla real de Word, así que se actualiza celda por
  celda: cada fila del Word se empareja con su fila en el Excel por el texto de la etiqueta (no por
  posición), porque la plantilla suele mostrar solo un subconjunto de las líneas del Excel (se
  omiten ítems poco significativos, pero los totales igual cuadran con el Excel completo); una fila
  del Word cuya etiqueta ya no calce con ninguna del Excel se deja tal cual, con un aviso, en vez de
  adivinar. La columna de **Comentarios nunca se toca** — queda igual que como la dejaste, para que
  la completes tú misma cada trimestre. Esto aplica también a la **versión en inglés** (tabla "Main
  sources and uses of funds...", columna "Comments" sin tocar): cada etiqueta del Excel se traduce
  primero (ver `FDEC_LABELS_EN` en `ar_scripts/assemble_en.py`) y luego se busca en el Word en
  inglés de la misma forma.

**Últimos ajustes (ronda 2, reportados por Melissa):**
- **Alineación de texto:** a pedido explícito, TODO el texto de las tablas (etiquetas, encabezados
  y números) ahora va alineado a la izquierda en vez de centrado — se notaba sobre todo en la
  tabla de Deuda Financiera Neta, pero el cambio aplica a todas las tablas, en español e inglés.
  Los títulos de sección que ocupan varias columnas (ej. "AÑO CALENDARIO (6 meses)") se mantienen
  centrados porque son títulos, no datos.
- **Exposición en Monedas:** la columna con el nombre de la cuenta se ensanchó (a costa de las 8
  columnas de moneda, que siguen entrando cómodas) para que las etiquetas largas quepan en menos
  líneas, y cada fila ahora mide solo lo que su texto necesita (antes todas las filas tenían el
  mismo alto fijo, pensado para la etiqueta más larga, así que la tabla se veía mucho más extensa
  hacia abajo de lo necesario).
- **Riesgos críticos duplicados (versión en inglés):** en algunas plantillas Word en inglés (se
  vio en la de marzo 2026, no en la de junio) queda una tabla nativa vieja, con contenido genérico
  de un borrador anterior, justo debajo del texto "Main critical risks faced as of ... are the
  following:" — esta tabla no existe en la versión en español, donde esa misma sección solo trae
  la imagen con el detalle real de riesgos (ya traducida). La app ahora detecta y elimina esa
  tabla vieja automáticamente, para que la sección quede igual que en español. Si tu plantilla en
  inglés de un trimestre antiguo tiene una estructura de imágenes muy distinta a la de junio 2026
  (por ejemplo, si a esa plantilla le falta alguna imagen de riesgos), otras tablas de esa misma
  plantilla podrían no quedar en el lugar correcto — avísame si ves algo así y lo revisamos para
  ese archivo puntual.
- **Traducción en vivo de la tabla de Riesgos (DeepL):** a pedido explícito de Melissa, aceptando
  el trade-off de que ese texto puntual sale hacia un servicio externo. Queda **desactivada por
  defecto** (si no configuras la API key, la app sigue funcionando igual que antes, 100% local) —
  ver "Traducción en vivo (DeepL)" más abajo para cómo activarla.

**Últimos ajustes (ronda 3, reportados por Melissa):** la traducción no se quedó solo en la tabla
de Riesgos — ahora, al generar la versión en inglés, la app compara **todo el documento, párrafo
por párrafo**, contra el Word en español **definitivo** de este trimestre (subirlo ahora es
**obligatorio**, ya no opcional) y detecta cualquier texto libre (título, subtítulos de sección, y
sobre todo el bloque de factores de riesgo / seguros / valor justo de la fruta) que no salga de las
plantillas ni del Excel. Ese texto se traduce vía DeepL y se inserta en el Word en inglés,
resaltado en amarillo, para que lo verifiques antes de publicar — en vez de quedar con la redacción
del trimestre anterior sin que nadie lo notara, que era justo el problema original. Si no
configuras DeepL, ese mismo texto queda en español, resaltado, pendiente de que lo traduzcas a
mano — ver "Traducción en vivo (DeepL)" y "Texto libre de todo el documento" más abajo.

## Instalación (una sola vez)

1. Instala Python 3.10 o superior si no lo tienes.
2. Abre una terminal en esta carpeta (`ar_app/`) y ejecuta:

   ```
   pip install -r requirements.txt
   ```

## Uso cada trimestre

1. En la terminal, en esta carpeta, ejecuta:

   ```
   streamlit run app.py
   ```

   Se abrirá una pestaña en tu navegador (típicamente en `http://localhost:8501`). Todo corre
   en tu computador; los archivos no se envían a ningún servidor externo.

2. Sube el Excel del trimestre (con la misma estructura de hojas de siempre) y el Word del
   Análisis Razonado **del trimestre inmediatamente anterior** — se usa como plantilla; su
   estructura de títulos y secciones debe mantenerse igual. Es decir: para generar junio subes
   el Word de marzo, para septiembre el de junio, para diciembre el de septiembre, y para marzo
   del año siguiente el de diciembre. La app funciona con los 4 cierres (marzo, junio,
   septiembre, diciembre): el período del informe (fechas, "Jun26"/"Dic25"/"T25/26", cantidad de
   meses, etc.) se deriva solo del Excel que subas, sin importar de qué cierre venga el Word
   base.

3. Haz clic en **"Generar Análisis Razonado actualizado"** y descarga el Word resultante.

4. **Antes de enviar el informe**, revisa:
   - Todo el texto **resaltado en amarillo**: son causas de negocio específicas (ej. "recambio
     varietal en China y Perú", "provisiones no recurrentes") que no están en el Excel y que
     tú debes completar o confirmar cada trimestre.
   - El panel "Ver detalle del proceso" dentro de la app: lista qué párrafos se actualizaron
     bien (`[OK]`), cuáles no se encontraron (`[AVISO]`) y errores (`[ERROR]`).
   - La columna de Comentarios del Flujo de Efectivo y lo que esta versión todavía no cubre (ver
     "Limitaciones" abajo).

## Qué se automatiza

- **El título, el párrafo preámbulo y los encabezados de sección** ("Al [fecha]", "Análisis
  EBITDA acumulado a...", "Análisis Resultado temporada...", etc.) se regeneran con la fecha y
  el período correctos de este cierre — así el informe nunca queda con la fecha del trimestre
  anterior (la del Word que subiste) pegada en el título o en los encabezados.
- **Todas las cifras y porcentajes** de los párrafos cubiertos (ver lista en
  `ar_scripts/assemble.py`, variable `SINGLE_TEMPLATES`): EBITDA, Resultado (bridge de 6 meses
  y de 12 meses/temporada), Ingresos (con driver volumen/precio), Costos y Gastos, Indicadores
  de Actividad, Indicadores Financieros y de Rentabilidad, Deuda Financiera Neta, y el bridge
  del Estado de Situación Financiera (Activos, Pasivos, Patrimonio).
- Las **tablas en forma de imagen** del Word (EBITDA, Ingresos, Costos, Balance, DFN,
  Indicadores de Actividad, Indicadores de Rentabilidad, Exposición Cambiaria, Riesgos, Seguros,
  Vencimientos, Fair Value de fruta, etc.) se vuelven a generar desde el Excel y se insertan
  reemplazando las imágenes anteriores.
- El **Flujo de Efectivo** (tabla nativa de Word, hoja `FdeC` del Excel): se actualizan los montos
  del período actual y del período anterior de cada fila, emparejando por el texto de la etiqueta.
  La columna de Comentarios nunca se toca — es la única parte de esta tabla que sigue siendo
  manual, a propósito.
- El **punteo de factores** que explican la variación de la Ganancia Atribuible (6M y 12M):
  se seleccionan automáticamente los de mayor magnitud (umbral US$3 millones), ordenados de
  mayor a menor impacto, separando los que ayudaron del resultado de los que lo compensaron
  (igual que en el documento original). El número de factores mostrados puede variar de
  trimestre a trimestre según qué tan grandes sean las variaciones ese período.
- Las frases de "causa y efecto" comparativas (ej. "explicado por el crecimiento en los
  ingresos de X%, compensado por Y") se regeneran solas comparando las variaciones del Excel —
  **no** quedan marcadas para revisión.
- **Barrido final de fechas/etiquetas sueltas:** al terminar, la app recorre todo el documento
  una vez más y corrige cualquier mención de la fecha, el "MesAA" o la temporada del trimestre
  BASE que haya quedado sin actualizar — útil sobre todo en secciones que todavía no tienen su
  propio párrafo cubierto (Riesgos, Seguros, Exposición Cambiaria, Vencimientos, Fair Value de
  fruta) y en frases de contexto libre que acompañan a los bridges de resultado, cuya redacción
  varía demasiado de trimestre a trimestre como para anclarlas una por una.

## Qué queda resaltado en amarillo (revisión manual)

Las causas de negocio específicas que no se pueden derivar del Excel: por ejemplo el motivo
exacto de un deterioro de activos (qué país, qué campo), el motivo de mayores gastos de
administración (qué provisión), o frases de contexto puntuales. La cifra que acompaña a esa
frase (ej. "US$5,79 millones") sí se actualiza sola; solo la explicación cualitativa queda
resaltada.

La app **mantiene la frase de causa tal como venía escrita en el Word que subiste** (no inventa
ni reemplaza el motivo) y le agrega, también resaltado, el aviso en mayúscula "CAUSA NO DERIVABLE
DEL EXCEL: actualizar si corresponde." — así solo tienes que leerla y confirmar que sigue
aplicando ese trimestre, o corregirla si cambió el motivo. Si en algún trimestre futuro cambia
tanto la redacción del Word que la app ya no logra identificar dónde estaba la frase de causa
anterior, el párrafo queda solo con el aviso en amarillo (sin frase previa) para que la completes
tú misma desde cero — revisa el detalle del proceso (`[AVISO]`) si esto ocurre.

## Limitaciones de esta primera versión (V1)

- **Flujo de Efectivo:** ya se actualiza sola (ver "Últimos ajustes" más arriba). Solo la columna
  de Comentarios sigue siendo manual — a propósito, para que la completes tú misma cada trimestre.
- **Texto libre de Riesgos, Seguros, Vencimientos y Fair Value de fruta:** las tablas en sí
  (montos, porcentajes) de estas secciones **sí se regeneran** desde el Excel igual que las demás
  tablas en forma de imagen (ver "Qué se automatiza" arriba); lo que no se regenera es el texto
  libre que las acompaña en el cuerpo del informe (si lo hay), y por eso el barrido final (ver más
  abajo) corrige ahí cualquier fecha suelta del trimestre anterior que haya quedado.
- **"Valor libro de la acción":** este indicador puntual no está cubierto todavía (ni la cifra ni
  la fecha se actualizan solas); revísalo a mano cada trimestre.
- Pequeñas diferencias de redacción pueden aparecer respecto al estilo exacto del informe
  original en algunos párrafos del bridge de Balance (orden de los ítems, conectores como
  "y"/"," ). El contenido numérico siempre es correcto; son matices de estilo que puedes
  ajustar con un vistazo rápido si quieres el calce perfecto.
- **Si subes como Word base un informe real de septiembre** (en vez de uno generado por esta
  misma app, que siempre queda con la redacción "estilo junio"): el informe original de
  septiembre junta en un solo párrafo breve, en el resumen del período, contenido que en
  marzo/junio/diciembre aparece más adelante como intro + 2 viñetas separadas. Esos 3 párrafos
  puntuales pueden no encontrarse (apareciendo como `[AVISO]`) en ese caso específico. Esto no
  aplica si el Word base ya fue generado por esta app en un trimestre anterior, ni a los Words
  originales de marzo, junio o diciembre.
- **El "barrido final de fechas/etiquetas sueltas"** (ver el detalle del proceso dentro de la
  app) solo puede reconstruir el período del trimestre base si su **título original no fue
  modificado a mano** ("Al [fecha]", en la primera página). Si subes como Word base un informe que
  ya pasó por esta app y luego editaste manualmente, o uno cuyo título ya no corresponde a su
  propio trimestre, este barrido no tendrá de dónde partir y lo avisará en el detalle del proceso
  (`[AVISO] No pude determinar el período del documento base...`). Lo normal — partir siempre del
  Word tal como salió del trimestre anterior — no tiene este problema.

## Versión en inglés (Word + Excel)

Más abajo en la misma página de la app hay una sección aparte **"🇬🇧 Versión en inglés"**:

1. Reutiliza automáticamente el Excel que ya subiste arriba (sección en español) — no hay que
   subirlo de nuevo. Tienes que subir el Word del Análisis Razonado **en inglés del trimestre
   anterior**: es la misma lógica de plantilla que la versión en español (le da el
   formato/estructura), **no** es el Word en español que acabas de generar este trimestre.
2. **Opcional, pero recomendado:** sube también el Word en **español** que la propia app generó
   este trimestre, ya con tus causas de negocio revisadas/completadas (resaltadas en amarillo).
   Con esto, la app toma exactamente ese texto resaltado — la causa tal como la escribiste, más
   cualquier otro cambio que le hayas hecho — y lo deja resaltado en el Word en inglés, en
   español, con el aviso "SPANISH — TRANSLATE INTO ENGLISH:" en vez del aviso genérico o de la
   causa (ya desactualizada) del Word en inglés del trimestre pasado. Solo te queda traducir esa
   frase corta al inglés; el contenido es exactamente el mismo en los dos idiomas porque parte de
   lo mismo que escribiste. Si no subes este Word, la sección funciona igual que siempre (ver más
   abajo).
3. Haz clic en **"Generar versión en inglés (Word + Excel)"**. La app entrega dos descargas:
   - El **Word en inglés** actualizado (mismas frases/tablas que la versión en español, ya
     traducidas y verificadas contra los informes en inglés de referencia).
   - Una copia del **Excel traducida al inglés**, acotada a las 18 tablas que aparecen en el Word
     (EBITDA, Ingresos, Costos, Balance, Exposición Cambiaria, Riesgos, Seguros, etc.): se quitan
     las pestañas que no se usan en el informe (notas internas, hojas de chequeo, Flujo de
     Efectivo, Covenants, etc.), y dentro de cada hoja que sí se conserva también se quitan las
     filas y columnas que no pertenecen a esa tabla (filas de trabajo por debajo, columnas
     auxiliares que no aparecen en la imagen) — queda igual de acotado a lo que se ve en el Word.
     En lo que se conserva, las etiquetas de fila/columna de cada tabla, los rangos de fecha
     ("Ene26 - Jun26" → "Jan26 - Jun26") y unidades comunes ("MUS$"→"ThUS$", "veces"→"times",
     "Días"→"Days", "variación"→"Variation") quedan traducidos. **No** se traducen los nombres de
     las pestañas (hojas), y las fórmulas del Excel quedan congeladas en su valor ya calculado
     (varias fórmulas del libro apuntan a otros archivos externos; recalcularlas con esta app
     arriesgaría romper esos vínculos, y como este Excel es para entregar —no para seguir
     editando en vivo— los valores estáticos son seguros). Los banners que en el Excel original
     combinan varias columnas (ej. "AÑO CALENDARIO (6 meses)" en EBITDA) quedan como texto normal
     en una sola celda en vez de centrados sobre varias columnas, para evitar que al borrar
     columnas el rango combinado quede desalineado del contenido.
4. Revisa igual que la versión en español: texto resaltado en amarillo y el detalle del proceso.

**Cómo funciona el traspaso de causas de negocio (sin traducción automática):** el Word en inglés
se regenera de forma independiente a partir del Excel, igual que el de español — no es una
traducción del Word en español, palabra por palabra. Las cifras, porcentajes y estructura del
texto se generan solos desde el Excel en ambos idiomas por separado (y ya están verificados contra
informes reales en inglés), pero las causas de negocio puntuales que no se pueden derivar del
Excel (el motivo exacto de un deterioro, una provisión no recurrente, etc.) sí necesitan ser
exactamente las mismas en los dos idiomas, porque las escribe una persona, no el Excel. Por eso,
subir el Word en español **definitivo** de este trimestre (paso 2 arriba) es **obligatorio** para
generar la versión en inglés:

- La app lee directamente el texto resaltado en amarillo de ese documento — tal como quedó, lo
  hayas dejado igual, reescrito, o le hayas agregado algo más — y lo inserta en el lugar
  correspondiente del Word en inglés. Sin DeepL configurado, queda todavía en español, resaltado,
  con el aviso "SPANISH — TRANSLATE INTO ENGLISH:", para que tú traduzcas esa frase corta (el resto
  del párrafo — cifras, porcentajes — ya viene generado correctamente en inglés). Con DeepL
  configurado, la tabla de Riesgos sí se traduce en vivo (ver "Traducción en vivo (DeepL)" más
  abajo); las causas de negocio puntuales, por ahora, siguen quedando para que las traduzcas tú,
  porque ahí un término técnico mal traducido es más riesgoso.
- Ese mismo Word definitivo también se usa para la traducción general de texto libre (ver "Texto
  libre de todo el documento" más abajo).

**Verificación extra — por si editaste algo más allá de la causa resaltada:** cuando subes el Word
en español ya editado (paso 2), la app además compara, párrafo por párrafo, la parte que **no**
está resaltada contra lo que el motor generaría de nuevo a partir de este mismo Excel. Si alguna
vez cambias a mano algo fuera del resaltado amarillo —una cifra, una palabra, lo que sea—, la app
no lo traspasa sola al inglés (insertar un fragmento recortado automáticamente en un documento
financiero formal es arriesgado), pero sí te deja un aviso explícito en el detalle del proceso,
con el nombre del párrafo, para que revises a mano que el Word en inglés diga lo mismo ahí. En el
uso normal (sin tocar nada fuera de lo resaltado en amarillo) no debería aparecer ningún aviso de
este tipo — si aparece alguno sin que hayas tocado nada, puede deberse a que Word reformateó algo
al guardar (comillas, espacios); igual revisa el párrafo señalado para estar segura.

Si la sección no aparece en tu instalación, falta copiar `ar_scripts/assemble_en.py`,
`ar_scripts/periods_en.py`, `ar_scripts/ar_engine_en.py` y `ar_scripts/excel_translator_en.py`
(ya incluidos en este zip) a tu carpeta de la app.

### Texto libre de todo el documento (comparación con el Word en español definitivo)

Además de las tablas y las causas de negocio, el Word trae texto que no sale ni del Excel ni de
ninguna plantilla de este motor: el título, algunos subtítulos de sección y, sobre todo, el bloque
largo de factores de riesgo / seguros / valor justo de la fruta (unos 70 párrafos). Antes, ese
texto quedaba tal cual venía en el Word en inglés **base** que subes — es decir, con la redacción
del **trimestre anterior** — sin ningún aviso de que podía estar desactualizado. Esto es justo lo
que reportaste: "la traducción no es solo de riesgos ... la app debe comparar la versión en
español definitiva con la versión en inglés, identificar las variaciones y proponer traducciones
(dejar destacado), para que el usuario verifique".

Ahora, al generar la versión en inglés, la app recorre **todo el documento, párrafo por párrafo**,
y clasifica cada uno como "cubierto" (viene de una plantilla, de una causa de negocio, o es una
viñeta de bridge) o "libre" (todo lo demás). Hace lo mismo con el Word en español **definitivo**
que subiste (obligatorio, ver más arriba) y empareja los tramos de texto libre de uno y otro
documento por el orden en que aparecen — se verificó contra los Word reales de junio 2026 que
español e inglés traen la misma cantidad de tramos libres, en el mismo orden.

**Importante: esto no traduce línea por línea sin más.** Cada párrafo libre en español se compara,
uno por uno, contra ese mismo párrafo en el Word en español del **trimestre anterior** (el "Word
base a actualizar" que ya subes en la sección en español, arriba — no hace falta subirlo de nuevo
acá). Si el texto es el mismo (ignorando espacios/comillas tipográficas sueltas), el párrafo en
inglés se deja **tal cual está** — su traducción ya se usó y aprobó en un informe anterior, así que
no hay razón para reemplazarla por una traducción nueva de DeepL que podría redactarse distinto
aunque diga lo mismo. Solo cuando el texto en español **sí cambió** respecto al trimestre pasado, el
párrafo se traduce (vía DeepL, si está configurado) y reemplaza al correspondiente en inglés,
**resaltado en amarillo** con la etiqueta "DeepL translation — please verify", para que lo revises
antes de publicar — nunca se publica una traducción automática sin que la persona la confirme. Si
DeepL no está configurado, o falla para algún párrafo puntual, ese párrafo (el que cambió) queda en
**español**, resaltado, con la etiqueta "SPANISH — DeepL not configured, translate manually".

Si no se puede comparar contra el trimestre anterior (no subiste ese Word, o la cantidad de
párrafos de ese tramo no calza entre los dos trimestres), la app trata ese párrafo como "cambiado" y
lo traduce de todos modos — mejor eso que dejar pasar un cambio real sin avisar. Y si algún
trimestre la cantidad de tramos libres no calza entre español e inglés (por ejemplo, porque se
agregó o se borró un párrafo de redacción libre en alguno de los dos), la app alinea los primeros
tramos en común y deja un `[AVISO]` en el detalle del proceso pidiendo revisar a mano el resto —
nunca inventa una alineación.

### Traducción en vivo (DeepL) — opcional, desactivada por defecto

Las traducciones de las tablas (EBITDA, Balance, Exposición Cambiaria, etc.) son diccionarios fijos
verificados a mano una vez contra el Word en inglés de referencia, y no cambian de un trimestre a
otro porque las etiquetas de esas tablas tampoco cambian. La tabla de **Riesgos críticos** y el
**texto libre de todo el documento** (ver sección anterior) son distintos: su contenido es texto
libre que sí puede cambiar cada trimestre, así que un diccionario fijo (o dejar la redacción vieja)
se desactualiza apenas se edita la redacción en español o aparece algo nuevo — esto fue justo lo
que reportaste como "el texto en inglés no está quedando exactamente igual al español".

Para resolverlo, la app puede traducir ese texto **en vivo** con la API de DeepL. Esto es opcional
y queda **desactivado por defecto** — importante que lo sepas porque implica un trade-off real: ese
texto (que todavía no es público ese trimestre) sale de este computador hacia los servidores de
DeepL para traducirse. Si no activas esto, la app sigue funcionando exactamente igual que antes:
100% local, con el diccionario fijo como respaldo para Riesgos, y el texto libre del resto del
documento queda en español, resaltado, pendiente de que lo traduzcas tú misma.

Para activarlo:
1. Crea una cuenta gratuita en [deepl.com/pro-api](https://www.deepl.com/pro-api) (el plan DeepL
   API Free no tiene costo, con un límite mensual de caracteres traducidos — de sobra para esto).
   Copia tu API key desde el panel de tu cuenta.
2. En Streamlit Community Cloud, entra a tu app → **Settings → Secrets**, y agrega (con comillas,
   o Streamlit lo rechaza por no ser TOML válido):
   ```
   DEEPL_API_KEY = "tu-api-key-aquí"
   ```
   (Nunca pongas la API key directamente en el código ni la subas a GitHub — el repositorio es
   público; "Secrets" de Streamlit Cloud es justamente el lugar seguro para esto, invisible para
   cualquiera que solo vea el código.)
3. Guarda — Streamlit Cloud vuelve a desplegar la app sola (toma alrededor de un minuto). Cuando
   generes la versión en inglés, vas a ver un aviso verde "🌐 Traducción en vivo vía DeepL
   activada" arriba del botón.

Si la API key no está configurada, o DeepL no responde (sin internet, cuota agotada, servicio
caído), la app no se cae: usa el diccionario fijo como respaldo para Riesgos, deja el resto del
texto libre en español resaltado, y sigue generando el resto del informe con normalidad.

## Cómo funciona (para referencia futura)

- El Word se procesa con `python-docx`: cada párrafo cubierto se **regenera completo** a partir
  de los datos del Excel (no es un find/replace de texto viejo), así que no depende de que el
  texto anterior sea exactamente el de un trimestre específico.
- Los párrafos se ubican por **patrones de texto** (ancla), no por posición fija — por eso el
  Word que subas debe mantener el mismo orden y los mismos títulos de sección de siempre. Si
  cambias la redacción de alguna frase "ancla" en el Word base, la app puede dejar de encontrar
  ese párrafo (aparecerá como `[AVISO]` en el detalle).
- Las fechas/periodos (Jun26, Dic25, T25/26, etc.) se leen del Excel (hoja `DF Neta`, fechas de
  cierre), así que la app debería seguir funcionando en trimestres futuros sin tocar el código,
  siempre que la estructura de hojas y columnas del Excel se mantenga igual.
- El código fuente de la lógica de generación está en `ar_scripts/`:
  - `ar_engine.py`: genera cada frase/párrafo a partir de celdas del Excel.
  - `periods.py`: deriva las etiquetas de período desde las fechas del Excel.
  - `assemble.py`: ubica los párrafos en el Word, los reemplaza, maneja el punteo de factores,
    y vuelve a renderizar las tablas-imagen.
  - `render_tables_v5.py`: el renderizador de tablas a imagen (compartido por español e inglés;
    incluye `TABLES`/`TABLES_EN` con las traducciones verificadas de cada tabla).
  - `ar_engine_en.py` / `periods_en.py` / `assemble_en.py`: equivalentes en inglés de
    `ar_engine.py` / `periods.py` / `assemble.py`.
  - `excel_translator_en.py`: genera la copia del Excel traducida al inglés (ver sección
    "Versión en inglés" arriba).

Si el Excel cambia de estructura (hojas, columnas) en algún trimestre, o si quieres sumar lo que
todavía queda pendiente (Valor libro de la acción), avísame y actualizo
`ar_engine.py`/`assemble.py`/`render_tables_v5.py` en consecuencia.
