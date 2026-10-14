# -*- coding: utf-8 -*-
"""
App local para automatizar el Análisis Razonado de Hortifrut.

Cómo correrla (una sola vez, en la carpeta de la app):
    pip install -r requirements.txt
    streamlit run app.py

Cada trimestre: subir el Excel con los datos financieros del trimestre y el Word del
Análisis Razonado del trimestre ANTERIOR (que sirve de plantilla de formato), y descargar
el Word actualizado. Todo el procesamiento ocurre en este computador; ningún archivo se
envía a servidores externos — con UNA excepción, opcional y desactivada por defecto: si se
configura una API key de DeepL (ver "Secrets" de Streamlit Cloud / variable DEEPL_API_KEY), la
versión en inglés traduce en vivo, mandándolo a los servidores de DeepL: el texto de la tabla de
Riesgos críticos, y (al generar la versión en inglés) todo el texto libre del documento que no
salga de las plantillas ni del Excel — comparando párrafo por párrafo el Word en español
DEFINITIVO de este trimestre contra el Word en inglés recién generado (a pedido explícito de
Melissa, para que la traducción no quede pegada a la de un trimestre anterior). Sin esa API key
configurada, nada cambia salvo que ese texto libre queda en español, resaltado, pendiente de que
la persona lo traduzca a mano: la app sigue funcionando 100% local, igual que antes.
"""
import os
import sys
import tempfile
import traceback
from datetime import datetime

import streamlit as st

# Funciona tanto si ar_engine.py/assemble.py/etc. están en una subcarpeta "ar_scripts" (como en el
# zip original) como si quedaron en la misma carpeta que este archivo (ej. al subirlos sueltos a un
# repositorio de GitHub) — usa la que exista.
_here = os.path.dirname(os.path.abspath(__file__))
_scripts_dir = os.path.join(_here, "ar_scripts")
sys.path.insert(0, _scripts_dir if os.path.isdir(_scripts_dir) else _here)
import assemble  # noqa: E402

# assemble_en.py / excel_translator_en.py son opcionales: si todavía no se subieron a la carpeta
# ar_scripts, la app sigue funcionando en español y simplemente no muestra la sección en inglés.
try:
    import assemble_en  # noqa: E402
    import excel_translator_en  # noqa: E402
    import deepl_translate  # noqa: E402
    _EN_AVAILABLE = True
except ImportError:
    _EN_AVAILABLE = False

# Traducción dinámica (DeepL) — tabla de Riesgos críticos Y, al generar la versión en inglés, todo
# el texto libre del documento no cubierto por las plantillas/Excel (ver assemble_en.
# translate_free_text_from_definitive_es) — opcional y desactivada por defecto. Se activa SOLO si
# configuraste la API key en "Secrets" de Streamlit Cloud (Settings → Secrets de tu app) así:
#     DEEPL_API_KEY = "tu-api-key-aquí"
# Sin esa key, _DEEPL_ON queda en False y la app sigue funcionando exactamente igual que antes
# (con el diccionario fijo de traducciones como respaldo) — ver deepl_translate.py.
_DEEPL_ON = False
if _EN_AVAILABLE:
    try:
        _deepl_key = st.secrets.get("DEEPL_API_KEY")
    except Exception:
        _deepl_key = None
    deepl_translate.configure_deepl(_deepl_key)
    _DEEPL_ON = deepl_translate.is_configured()

st.set_page_config(page_title="Análisis Razonado — Hortifrut", page_icon="📊", layout="centered")

st.title("📊 Generador de Análisis Razonado")
st.caption(
    "Sube el Excel del trimestre y el Word del Análisis Razonado del trimestre anterior. "
    "La app actualiza las tablas y gran parte del texto narrativo automáticamente, y deja "
    "resaltado en amarillo lo que requiere revisión manual (causas de negocio que no están "
    "en el Excel)."
)

st.divider()

col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("1. Excel del trimestre (.xlsx)", type=["xlsx"])
with col2:
    word_file = st.file_uploader("2. Word base a actualizar (.docx)", type=["docx"])

st.caption(
    "El Word base debe tener la misma estructura de siempre (mismos títulos y orden de "
    "secciones); la app busca los párrafos por su texto, no por posición fija."
)

generate_clicked = st.button("Generar Análisis Razonado actualizado", type="primary", disabled=not (excel_file and word_file))

if generate_clicked:
    with tempfile.TemporaryDirectory() as tmp:
        excel_path = os.path.join(tmp, "input.xlsx")
        word_path = os.path.join(tmp, "input.docx")
        out_path = os.path.join(tmp, "Analisis_Razonado_actualizado.docx")

        with open(excel_path, "wb") as f:
            f.write(excel_file.getbuffer())
        with open(word_path, "wb") as f:
            f.write(word_file.getbuffer())

        with st.spinner("Generando documento..."):
            try:
                log = assemble.generate(excel_path, word_path, out_path)
                ok = True
            except Exception:
                log = [traceback.format_exc()]
                ok = False

        data = None
        if ok and os.path.exists(out_path):
            with open(out_path, "rb") as f:
                data = f.read()

        # Se guarda en session_state (en vez de usarse solo dentro de este "if") para que el botón
        # de descarga y el detalle del proceso sigan apareciendo aunque Streamlit vuelva a correr el
        # script por otra interacción (ej. hacer clic en el propio botón de descarga) — si no, al
        # hacer clic en "Descargar" este bloque completo desaparecía en el rerun siguiente.
        st.session_state["es_result"] = {
            "ok": ok,
            "data": data,
            "fname": f"Analisis_Razonado_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
            "log": log,
        }

result = st.session_state.get("es_result")
if result:
    log = result["log"]
    alertas = [l for l in log if l.startswith("[ALERTA]")]

    if result["ok"] and result["data"]:
        if alertas:
            st.warning(
                "⚠️ El documento se generó, pero probablemente a partir del Word equivocado: "
                "casi todo el texto narrativo quedó SIN actualizar (revisa el detalle abajo antes "
                "de usar este archivo)."
            )
        else:
            st.success("Documento generado correctamente.")
        st.download_button(
            "⬇️ Descargar Word actualizado",
            data=result["data"],
            file_name=result["fname"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="dl_word_es",
        )
    else:
        st.error("Hubo un error generando el documento. Revisa el detalle abajo.")

    with st.expander("Ver detalle del proceso (qué se actualizó, avisos y errores)", expanded=not result["ok"] or bool(alertas)):
        avisos = [l for l in log if l.startswith("[AVISO]")]
        errores = [l for l in log if l.startswith("[ERROR]") or "Traceback" in l]
        oks = [l for l in log if l.startswith("[OK]")]
        infos = [l for l in log if l.startswith("[INFO]")]

        if alertas:
            st.markdown("**⚠️ Posible Word equivocado:**")
            st.code("\n".join(alertas))
        if infos:
            st.write("\n".join(infos))
        if errores:
            st.markdown("**Errores:**")
            st.code("\n".join(errores))
        if avisos:
            st.markdown("**Avisos (revisar):**")
            st.code("\n".join(avisos))
        st.markdown(f"**{len(oks)} elementos actualizados correctamente.**")
        st.code("\n".join(oks) if oks else "(sin detalle)")

st.divider()
st.caption(
    "Recuerda revisar antes de enviar: todo el texto resaltado en amarillo en el Word "
    "resultante, la columna de Comentarios del Flujo de Efectivo (se deja para que la completes "
    "tú misma) y las secciones que esta versión todavía no cubre (Valor libro de la acción) — ver "
    "README."
)

# ---------------------------------------------------------------------------
# Versión en inglés (Word + Excel traducido)
# ---------------------------------------------------------------------------
st.divider()
st.header("🇬🇧 Versión en inglés")

if not _EN_AVAILABLE:
    st.info(
        "La versión en inglés todavía no está instalada en esta carpeta: faltan "
        "`assemble_en.py` y/o `excel_translator_en.py` dentro de `ar_scripts/`."
    )
else:
    # El Excel del trimestre hace falta sí o sí para esta sección (de ahí salen las 18 tablas que
    # se renderizan de nuevo, en inglés) — pero no todo el mundo necesita generar primero el Word en
    # ESPAÑOL con la sección de arriba: si ya tienes el Word en español definitivo de este trimestre
    # (revisado y editado a mano, fuera de la app, o de una sesión anterior donde Streamlit ya
    # "olvidó" lo que habías subido arriba — ver nota de la sesión más abajo), lo único que falta es
    # el Excel y los Word de esta sección. Antes, si no habías subido el Excel arriba, esta sección
    # entera quedaba bloqueada con un mensaje que igual te mandaba a la sección en español aunque no
    # la necesitaras para nada — esto es lo que Melissa pidió poder saltarse ("ir directo al paso
    # 2"). Ahora, si el Excel ya está subido arriba se reutiliza igual que antes; si no, esta
    # sección trae su propio uploader para no depender de la sección en español en absoluto.
    excel_file_en_own = None
    if not excel_file:
        st.info(
            "¿Ya tienes el Word en español **definitivo** de este trimestre listo (no necesitas "
            "generar el Word en español con la app)? Puedes saltarte la sección de arriba: sube "
            "aquí el Excel del trimestre y sigue directo con la versión en inglés."
        )
        excel_file_en_own = st.file_uploader("Excel del trimestre (.xlsx) — para la versión en inglés", type=["xlsx"], key="excel_en_own")
    _excel_candidate = excel_file or excel_file_en_own

    if not _excel_candidate:
        pass  # el aviso de arriba ya deja claro qué falta; nada más que mostrar todavía
    else:
        st.caption(
            f"Usando el Excel {'ya subido más arriba' if excel_file else 'subido acá'} "
            f"({_excel_candidate.name}). Sube además el Word del Análisis Razonado en INGLÉS "
            "**del trimestre anterior** (sirve solo de plantilla de formato — títulos, orden de "
            "secciones, estilos — igual que el Word base de la sección en español; no es el Word "
            "en español que acabas de generar). La app genera el Word en inglés actualizado y, "
            "además, una copia del Excel con las etiquetas de las tablas y las fechas ya "
            "traducidas al inglés (mismos valores y fórmulas congeladas en sus valores "
            "calculados; no se traduce el texto libre fuera de las tablas ni los nombres de "
            "pestaña)."
        )
        word_file_en = st.file_uploader("Word base en inglés del trimestre anterior (.docx)", type=["docx"], key="word_en")

        word_file_es_edited = st.file_uploader(
            "Obligatorio — Word en ESPAÑOL **definitivo** de este trimestre, ya revisado y editado (.docx)",
            type=["docx"],
            key="word_es_edited",
        )
        st.caption(
            "Súbelo después de dejar listo el Word en español definitivo de este trimestre (causas de "
            "negocio completadas, y cualquier otro ajuste de redacción ya hecho). La app toma de acá "
            "las causas de negocio resaltadas en amarillo tal como quedaron escritas."
        )

        # Word en español del TRIMESTRE ANTERIOR: punto de comparación para saber qué texto libre
        # cambió de verdad este trimestre (ver assemble_en.translate_free_text_from_definitive_es).
        # Antes esto se tomaba en silencio del "2. Word base a actualizar" de la sección en español de
        # arriba (mismo archivo, reutilizado) — pero si esta sección se usa sola, en una sesión donde
        # esa otra sección no tiene nada subido (o Melissa solo quiere regenerar el inglés sin volver a
        # tocar la sección en español), ese campo queda vacío sin que se note, y la comparación se
        # pierde: TODO el texto libre se trata como "cambiado" y queda marcado para traducir, incluso
        # texto que en realidad no cambió (esto fue justamente lo que reportó Melissa: "toda la sección
        # de riesgo quedó como pendiente de traducción... esto no debería haber cambiado"). Por eso
        # ahora hay un uploader propio, independiente del de la sección en español, con un estado
        # explícito de si hay o no punto de comparación ANTES de generar.
        word_file_es_base_en = st.file_uploader(
            "Word en ESPAÑOL del TRIMESTRE ANTERIOR (para comparar qué cambió; recomendado) (.docx)",
            type=["docx"],
            key="word_es_base_en",
        )
        _es_base_candidate = word_file_es_base_en or word_file
        if word_file_es_base_en:
            st.caption(
                "✅ Tengo el Word en español del trimestre anterior (subido acá mismo): se compara "
                "párrafo por párrafo y solo se marca para traducir el texto libre que de verdad cambió."
            )
        elif word_file:
            st.caption(
                "✅ Tengo el Word en español del trimestre anterior — reutilizando el mismo que subiste "
                "arriba en \"2. Word base a actualizar\" (sección en español). Si es un archivo distinto, "
                "súbelo en el campo de arriba."
            )
        else:
            st.caption(
                "⚠️ No tengo el Word en español del trimestre anterior (ni aquí ni en \"2. Word base a "
                "actualizar\" arriba): no voy a poder distinguir qué texto libre cambió de verdad — "
                "**todo** el texto libre (título, subtítulos, factores de riesgo, seguros, valor justo de "
                "la fruta, etc.) se marcará como pendiente de traducir, aunque no haya cambiado. Sube el "
                "Word en español del trimestre anterior arriba para evitar esto."
            )

        if _DEEPL_ON:
            st.caption(
                "🌐 Traducción en vivo vía DeepL activada: el texto libre que cambió este trimestre (y el "
                "de la tabla de Riesgos críticos) se traduce mandándolo a los servidores de DeepL, en vez "
                "de dejarlo en español pendiente de traducir a mano."
            )
        else:
            st.caption(
                "Traducción en vivo vía DeepL: desactivada (no hay API key configurada en \"Secrets\" "
                "de Streamlit Cloud). El texto libre que cambió este trimestre queda en español, "
                "resaltado, pendiente de que tú lo traduzcas a mano; la tabla de Riesgos usa su "
                "traducción fija guardada."
            )

        generate_en_clicked = st.button(
            "Generar versión en inglés (Word + Excel)",
            type="primary",
            disabled=not (word_file_en and word_file_es_edited),
        )
        if word_file_en and not word_file_es_edited:
            st.caption("⚠️ Falta subir el Word en español definitivo de este trimestre (obligatorio) para poder generar el inglés.")

        if generate_en_clicked:
            with tempfile.TemporaryDirectory() as tmp_en:
                excel_path_en = os.path.join(tmp_en, "input_en.xlsx")
                word_path_en = os.path.join(tmp_en, "input_en.docx")
                out_word_path_en = os.path.join(tmp_en, "Analisis_Razonado_EN.docx")
                out_excel_path_en = os.path.join(tmp_en, "Datos_EN.xlsx")

                with open(excel_path_en, "wb") as f:
                    f.write(_excel_candidate.getbuffer())
                with open(word_path_en, "wb") as f:
                    f.write(word_file_en.getbuffer())

                word_es_edited_path = os.path.join(tmp_en, "input_es_edited.docx")
                with open(word_es_edited_path, "wb") as f:
                    f.write(word_file_es_edited.getbuffer())

                # Punto de comparación para saber qué texto libre cambió este trimestre (ver
                # assemble_en.translate_free_text_from_definitive_es): el uploader propio de esta
                # sección si se usó, si no el de la sección en español (ver _es_base_candidate y los
                # avisos de arriba) — si tampoco hay eso, la app igual funciona, pero traduce todo el
                # texto libre en vez de solo lo que cambió (se avisa en el detalle del proceso).
                word_es_base_path = None
                if _es_base_candidate:
                    word_es_base_path = os.path.join(tmp_en, "input_es_base.docx")
                    with open(word_es_base_path, "wb") as f:
                        f.write(_es_base_candidate.getbuffer())

                with st.spinner("Generando Word en inglés..."):
                    try:
                        log_en = assemble_en.generate_en(
                            excel_path_en, word_path_en, out_word_path_en,
                            word_es_edited_path, word_es_base_path,
                        )
                        ok_word_en = True
                    except Exception:
                        log_en = [traceback.format_exc()]
                        ok_word_en = False

                with st.spinner("Generando Excel traducido al inglés..."):
                    try:
                        log_xl_en = excel_translator_en.translate_workbook_to_en(excel_path_en, out_excel_path_en)
                        ok_excel_en = os.path.exists(out_excel_path_en) and not any(
                            l.startswith("[ERROR]") for l in log_xl_en
                        )
                    except Exception:
                        log_xl_en = [traceback.format_exc()]
                        ok_excel_en = False

                data_word_en = None
                if ok_word_en and os.path.exists(out_word_path_en):
                    with open(out_word_path_en, "rb") as f:
                        data_word_en = f.read()

                data_excel_en = None
                if ok_excel_en:
                    with open(out_excel_path_en, "rb") as f:
                        data_excel_en = f.read()

                # Igual que en la sección en español: se guarda todo en session_state para que los dos
                # botones de descarga (Word + Excel) y el detalle del proceso sigan visibles aunque se
                # haga clic en uno de ellos — antes, al estar solo dentro de este "if generate_en_clicked",
                # un clic en "Descargar Excel en inglés" hacía desaparecer el botón de "Descargar Word en
                # inglés" (y viceversa) en el rerun siguiente, porque Streamlit vuelve a correr el script
                # completo en cada clic y generate_en_clicked vuelve a ser False.
                st.session_state["en_results"] = {
                    "ok_word_en": ok_word_en,
                    "data_word_en": data_word_en,
                    "fname_word_en": f"Analisis_Razonado_EN_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                    "ok_excel_en": ok_excel_en,
                    "data_excel_en": data_excel_en,
                    "fname_excel_en": f"Datos_AR_EN_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    "log_en": log_en,
                    "log_xl_en": log_xl_en,
                }

        en_results = st.session_state.get("en_results")
        if en_results:
            log_en = en_results["log_en"]
            alertas_en = [l for l in log_en if l.startswith("[ALERTA]")]
            alertas_deepl_en = [l for l in log_en if l.startswith("[ALERTA-DEEPL]")]

            if en_results["ok_word_en"] and en_results["data_word_en"]:
                if alertas_en:
                    st.warning(
                        "⚠️ El Word en inglés se generó, pero probablemente a partir del Word base "
                        "equivocado: casi todo el texto narrativo quedó SIN actualizar (revisa el "
                        "detalle abajo antes de usar este archivo)."
                    )
                if alertas_deepl_en:
                    st.warning(
                        "⚠️ DeepL está configurado, pero una o más traducciones fallaron igual: ese "
                        "texto quedó en español, resaltado, pendiente de traducir a mano (revisa el "
                        "motivo exacto en el detalle abajo)."
                    )
                if not alertas_en and not alertas_deepl_en:
                    st.success("Word en inglés generado correctamente.")
                st.download_button(
                    "⬇️ Descargar Word en inglés",
                    data=en_results["data_word_en"],
                    file_name=en_results["fname_word_en"],
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl_word_en",
                )
            else:
                st.error("Hubo un error generando el Word en inglés. Revisa el detalle abajo.")

            if en_results["ok_excel_en"] and en_results["data_excel_en"]:
                st.success("Excel traducido al inglés generado correctamente.")
                st.download_button(
                    "⬇️ Descargar Excel en inglés",
                    data=en_results["data_excel_en"],
                    file_name=en_results["fname_excel_en"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_excel_en",
                )
            else:
                st.error("Hubo un error generando el Excel en inglés. Revisa el detalle abajo.")

            with st.expander(
                "Ver detalle del proceso (Word + Excel en inglés)",
                expanded=not (en_results["ok_word_en"] and en_results["ok_excel_en"]) or bool(alertas_en) or bool(alertas_deepl_en),
            ):
                st.markdown("**Word en inglés:**")
                avisos_en = [l for l in log_en if l.startswith("[AVISO]")]
                errores_en = [l for l in log_en if l.startswith("[ERROR]") or "Traceback" in l]
                oks_en = [l for l in log_en if l.startswith("[OK]")]
                if alertas_en:
                    st.markdown("_⚠️ Posible Word base equivocado:_")
                    st.code("\n".join(alertas_en))
                if alertas_deepl_en:
                    st.markdown("_⚠️ DeepL configurado pero con fallas:_")
                    st.code("\n".join(alertas_deepl_en))
                if errores_en:
                    st.markdown("_Errores:_")
                    st.code("\n".join(errores_en))
                if avisos_en:
                    st.markdown("_Avisos (revisar):_")
                    st.code("\n".join(avisos_en))
                st.markdown(f"_{len(oks_en)} elementos actualizados correctamente._")
                st.code("\n".join(oks_en) if oks_en else "(sin detalle)")

                st.markdown("**Excel en inglés:**")
                st.code("\n".join(en_results["log_xl_en"]) if en_results["log_xl_en"] else "(sin detalle)")

        st.caption(
            "El Excel en inglés queda acotado a las 18 tablas que aparecen en el Word: se quitan las "
            "demás pestañas (notas internas, Flujo de Efectivo, etc.) y, dentro de cada hoja que se "
            "conserva, también las filas/columnas que no son parte de esa tabla. En lo que se "
            "conserva se traducen las etiquetas de fila/columna, los rangos de fecha, "
            "'MUS$'→'ThUS$', 'veces'→'times', 'Días'→'Days' y 'variación'→'Variation'. Los nombres "
            "de las pestañas quedan tal como están en el Excel original."
        )
