# -*- coding: utf-8 -*-
"""
App local para automatizar el Análisis Razonado de Hortifrut.

Cómo correrla (una sola vez, en la carpeta de la app):
    pip install -r requirements.txt
    streamlit run app.py

Cada trimestre: subir el Excel con los datos financieros del trimestre y el Word del
Análisis Razonado del trimestre ANTERIOR (que sirve de plantilla de formato), y descargar
el Word actualizado. Todo el procesamiento ocurre en este computador; ningún archivo se
envía a servidores externos.
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
    _EN_AVAILABLE = True
except ImportError:
    _EN_AVAILABLE = False

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

        if ok and os.path.exists(out_path):
            st.success("Documento generado correctamente.")
            with open(out_path, "rb") as f:
                data = f.read()
            fname = f"Analisis_Razonado_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
            st.download_button(
                "⬇️ Descargar Word actualizado",
                data=data,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        else:
            st.error("Hubo un error generando el documento. Revisa el detalle abajo.")

        with st.expander("Ver detalle del proceso (qué se actualizó, avisos y errores)", expanded=not ok):
            avisos = [l for l in log if l.startswith("[AVISO]")]
            errores = [l for l in log if l.startswith("[ERROR]") or "Traceback" in l]
            oks = [l for l in log if l.startswith("[OK]")]
            infos = [l for l in log if l.startswith("[INFO]")]

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
    "resultante, y las secciones que esta versión aún no regenera (Flujo de Efectivo, "
    "Riesgos/Seguros/Exposición Cambiaria, Valor libro de la acción) — ver README."
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
    st.caption(
        "Sube el mismo Excel del trimestre y el Word del Análisis Razonado en INGLÉS del "
        "trimestre anterior (sirve de plantilla de formato). La app genera el Word en inglés "
        "actualizado y, además, una copia del Excel con las etiquetas de las tablas y las "
        "fechas ya traducidas al inglés (mismos valores y fórmulas congeladas en sus valores "
        "calculados; no se traduce el texto libre fuera de las tablas ni los nombres de "
        "pestaña)."
    )

    col1_en, col2_en = st.columns(2)
    with col1_en:
        excel_file_en = st.file_uploader("1. Excel del trimestre (.xlsx)", type=["xlsx"], key="excel_en")
    with col2_en:
        word_file_en = st.file_uploader("2. Word base en inglés a actualizar (.docx)", type=["docx"], key="word_en")

    generate_en_clicked = st.button(
        "Generar versión en inglés (Word + Excel)",
        type="primary",
        disabled=not (excel_file_en and word_file_en),
    )

    if generate_en_clicked:
        with tempfile.TemporaryDirectory() as tmp_en:
            excel_path_en = os.path.join(tmp_en, "input_en.xlsx")
            word_path_en = os.path.join(tmp_en, "input_en.docx")
            out_word_path_en = os.path.join(tmp_en, "Analisis_Razonado_EN.docx")
            out_excel_path_en = os.path.join(tmp_en, "Datos_EN.xlsx")

            with open(excel_path_en, "wb") as f:
                f.write(excel_file_en.getbuffer())
            with open(word_path_en, "wb") as f:
                f.write(word_file_en.getbuffer())

            with st.spinner("Generando Word en inglés..."):
                try:
                    log_en = assemble_en.generate_en(excel_path_en, word_path_en, out_word_path_en)
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

            if ok_word_en and os.path.exists(out_word_path_en):
                st.success("Word en inglés generado correctamente.")
                with open(out_word_path_en, "rb") as f:
                    data_word_en = f.read()
                fname_word_en = f"Analisis_Razonado_EN_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
                st.download_button(
                    "⬇️ Descargar Word en inglés",
                    data=data_word_en,
                    file_name=fname_word_en,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl_word_en",
                )
            else:
                st.error("Hubo un error generando el Word en inglés. Revisa el detalle abajo.")

            if ok_excel_en:
                st.success("Excel traducido al inglés generado correctamente.")
                with open(out_excel_path_en, "rb") as f:
                    data_excel_en = f.read()
                fname_excel_en = f"Datos_AR_EN_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                st.download_button(
                    "⬇️ Descargar Excel en inglés",
                    data=data_excel_en,
                    file_name=fname_excel_en,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_excel_en",
                )
            else:
                st.error("Hubo un error generando el Excel en inglés. Revisa el detalle abajo.")

            with st.expander("Ver detalle del proceso (Word + Excel en inglés)", expanded=not (ok_word_en and ok_excel_en)):
                st.markdown("**Word en inglés:**")
                avisos_en = [l for l in log_en if l.startswith("[AVISO]")]
                errores_en = [l for l in log_en if l.startswith("[ERROR]") or "Traceback" in l]
                oks_en = [l for l in log_en if l.startswith("[OK]")]
                if errores_en:
                    st.markdown("_Errores:_")
                    st.code("\n".join(errores_en))
                if avisos_en:
                    st.markdown("_Avisos (revisar):_")
                    st.code("\n".join(avisos_en))
                st.markdown(f"_{len(oks_en)} elementos actualizados correctamente._")
                st.code("\n".join(oks_en) if oks_en else "(sin detalle)")

                st.markdown("**Excel en inglés:**")
                st.code("\n".join(log_xl_en) if log_xl_en else "(sin detalle)")

    st.caption(
        "El Excel en inglés queda acotado a las 18 tablas que aparecen en el Word: se quitan las "
        "demás pestañas (notas internas, Flujo de Efectivo, etc.) y, dentro de cada hoja que se "
        "conserva, también las filas/columnas que no son parte de esa tabla. En lo que se "
        "conserva se traducen las etiquetas de fila/columna, los rangos de fecha, "
        "'MUS$'→'ThUS$', 'veces'→'times', 'Días'→'Days' y 'variación'→'Variation'. Los nombres "
        "de las pestañas quedan tal como están en el Excel original."
    )
