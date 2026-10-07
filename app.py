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

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "ar_scripts"))
import assemble  # noqa: E402

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
    "resultante, y las dos tablas de imagen que esta versión no regenera todavía "
    "(Indicadores de Actividad y Rentabilidad) — ver README."
)
