# Generador de Análisis Razonado — Hortifrut

App local (no sube nada a internet) que toma el Excel trimestral de datos financieros y el
Word del Análisis Razonado, y devuelve un Word actualizado: con las tablas (imágenes) y gran
parte del texto narrativo regenerados automáticamente a partir del Excel.

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
   Análisis Razonado (puede ser el del trimestre anterior — se usa como plantilla; su
   estructura de títulos y secciones debe mantenerse igual).

3. Haz clic en **"Generar Análisis Razonado actualizado"** y descarga el Word resultante.

4. **Antes de enviar el informe**, revisa:
   - Todo el texto **resaltado en amarillo**: son causas de negocio específicas (ej. "recambio
     varietal en China y Perú", "provisiones no recurrentes") que no están en el Excel y que
     tú debes completar o confirmar cada trimestre.
   - El panel "Ver detalle del proceso" dentro de la app: lista qué párrafos se actualizaron
     bien (`[OK]`), cuáles no se encontraron (`[AVISO]`) y errores (`[ERROR]`).
   - Las dos secciones no cubiertas todavía (ver "Limitaciones" abajo).

## Qué se automatiza

- **Todas las cifras y porcentajes** de los párrafos cubiertos (ver lista en
  `ar_scripts/assemble.py`, variable `SINGLE_TEMPLATES`): EBITDA, Resultado (bridge de 6 meses
  y de 12 meses/temporada), Ingresos (con driver volumen/precio), Costos y Gastos, Indicadores
  de Actividad, Indicadores Financieros y de Rentabilidad, Deuda Financiera Neta, y el bridge
  del Estado de Situación Financiera (Activos, Pasivos, Patrimonio).
- Las **tablas en forma de imagen** del Word (EBITDA, Ingresos, Costos, Balance, DFN, etc.) se
  vuelven a generar desde el Excel y se insertan reemplazando las imágenes anteriores.
- El **punteo de factores** que explican la variación de la Ganancia Atribuible (6M y 12M):
  se seleccionan automáticamente los de mayor magnitud (umbral US$3 millones), ordenados de
  mayor a menor impacto, separando los que ayudaron del resultado de los que lo compensaron
  (igual que en el documento original). El número de factores mostrados puede variar de
  trimestre a trimestre según qué tan grandes sean las variaciones ese período.
- Las frases de "causa y efecto" comparativas (ej. "explicado por el crecimiento en los
  ingresos de X%, compensado por Y") se regeneran solas comparando las variaciones del Excel —
  **no** quedan marcadas para revisión.

## Qué queda resaltado en amarillo (revisión manual)

Las causas de negocio específicas que no se pueden derivar del Excel: por ejemplo el motivo
exacto de un deterioro de activos (qué país, qué campo), el motivo de mayores gastos de
administración (qué provisión), o frases de contexto puntuales. La cifra que acompaña a esa
frase (ej. "US$5,79 millones") sí se actualiza sola; solo la explicación cualitativa queda
resaltada.

## Limitaciones de esta primera versión (V1)

- **Indicadores de Actividad y Rentabilidad (tablas de imagen):** estas dos hojas cambiaron de
  formato de columnas en el Excel más reciente respecto al archivo original con el que se
  construyó el renderizador de tablas, y no se volvieron a especificar por precaución (para no
  arriesgar un renderizado incorrecto). Estas dos imágenes **no se actualizan** todavía — hay
  que reemplazarlas a mano, o pedir que se agregue el soporte una vez confirmado el layout
  definitivo de esas hojas.
- **Flujo de Efectivo:** es una tabla nativa de Word (no una imagen), no está cubierta en esta
  versión.
- **Sección de Riesgos** (exposición cambiaria, sensibilidad de tasa, margen de valor
  razonable de fruta): los números sueltos de esta sección no están cubiertos todavía.
- Pequeñas diferencias de redacción pueden aparecer respecto al estilo exacto del informe
  original en algunos párrafos del bridge de Balance (orden de los ítems, conectores como
  "y"/"," ). El contenido numérico siempre es correcto; son matices de estilo que puedes
  ajustar con un vistazo rápido si quieres el calce perfecto.

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
  - `render_tables_v5.py`: el renderizador de tablas a imagen (reutilizado del pipeline
    anterior).

Si el Excel cambia de estructura (hojas, columnas) en algún trimestre, o si quieres sumar las
secciones pendientes (Flujo de Efectivo, Riesgos, Indicadores de Actividad/Rentabilidad), avísame
y actualizo `ar_engine.py`/`assemble.py` en consecuencia.
