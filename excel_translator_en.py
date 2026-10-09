# -*- coding: utf-8 -*-
"""
Genera una copia en INGLÉS del Excel de datos del trimestre (ar_v2.xlsx y sucesores), para que
Melissa pueda descargarlo junto con el Word en inglés — sin tener que traducir el Excel a mano
cada trimestre.

Dos pasadas de traducción, ambas reutilizando exactamente la lógica ya verificada que usa
render_tables_v5.py para las tablas-imagen del Word en inglés (misma fuente de verdad, cero
traducciones nuevas inventadas aquí):

  1) Pasada GENÉRICA sobre TODO el libro (todas las hojas, todas las celdas de texto): traduce
     rangos de fecha ("Ene26 - Jun26" → "Jan26 - Jun26"), unidades ("MUS$" → "ThUS$", "veces" →
     "times", "Días" → "Days", "variación" → "Variation") y fechas largas dinámicas
     ("Exposición Neta al 30 junio 2026" → "Net Exposure as of June 30, 2026"), vía
     RT._translate_common_strings_en(). Esto cubre texto suelto fuera de las tablas mapeadas en
     TABLES (ej. la hoja 'Tabla inicial') sin tener que listarlo celda por celda.
  2) Pasada DIRIGIDA, tabla por tabla (RT.TABLES_EN): escribe las etiquetas de fila/columna ya
     traducidas y verificadas contra el Excel/Word en inglés de referencia que envió Melissa
     (trimestre Jun26), usando la fila real de cada tabla resuelta con RT.resolve_row_nums() —
     la MISMA resolución (incluido el ajuste por 'anchor') que usa render_table() para que la
     celda coincida aunque la plantilla de un trimestre futuro haya corrido una fila (ver
     dic25/sep25/mar26). Esta resolución se calcula UNA SOLA VEZ, antes de traducir nada, y se
     reutiliza después para recortar filas/columnas (ver _trim_unused_rows_cols): el anchor busca
     texto literal en español, así que recalcularlo sobre la hoja ya traducida lo dejaría "ciego"
     justo donde más importa.

Decisiones de diseño (confirmadas contra el Excel en inglés de referencia de Melissa):
  - Se carga el libro con data_only=True: las fórmulas quedan congeladas en su valor ya
    calculado. Varias fórmulas del libro apuntan a OTROS workbooks externos (ej.
    '=+[1]EBITDA!C4'), y reescribir/guardar esos vínculos con openpyxl arriesga corromperlos;
    como este Excel es para ENTREGAR (no para seguir recalculando en vivo), los valores
    estáticos son seguros y muestran exactamente lo mismo que ve el lector.
  - Los NOMBRES DE PESTAÑA (hojas) no se traducen — el Excel en inglés de referencia de Melissa
    mantiene los mismos nombres de pestaña en español/mixto que el Excel fuente.
  - Los number_format de las celdas no se tocan — son idénticos entre el Excel en español y el
    de referencia en inglés (el separador de miles/decimales depende de la configuración
    regional de Excel al abrirlo, no del string de formato).
  - Las celdas de un rango combinado (merge) que no son la celda superior-izquierda son
    MergedCell de solo lectura en openpyxl: se saltan en la pasada genérica, y para los banners
    de grupo (group_headers, ej. EBITDA/VctosPtmos) se escribe solo en la celda superior-
    izquierda del rango combinado, igual que usa render_table() para dibujar el banner.
  - Solo se conservan las hojas que alimentan alguna tabla del Word (las mismas 18 de
    render_tables_v5.TABLES/IMAGE_TO_TABLE) — a pedido explícito de Melissa, para que el Excel
    en inglés quede más limpio. El resto de las hojas del libro fuente (ej. 'Tabla inicial',
    'FdeC'/Flujo de Efectivo, u otras no usadas por ninguna tabla del informe) se eliminan del
    archivo de salida.
  - Dentro de cada hoja que SÍ se conserva, también se eliminan las filas y columnas que no
    pertenecen a la tabla que se muestra en el Word (ej. filas de trabajo/notas por debajo de la
    tabla, o columnas auxiliares que no aparecen en la imagen) — a pedido explícito de Melissa.
    Se calculan con la misma fuente de verdad que usa el renderizador (resolve_row_nums() +
    spec['cols']), y cuando una hoja alimenta más de una tabla (ej. 'ExposicTC' → ExposicTC_tabla
    + EfectoTC) se conserva la UNIÓN de filas/columnas de todas. Los rangos combinados (merge) de
    la hoja se desarman antes de borrar (ver _trim_unused_rows_cols): openpyxl no reubica sus
    coordenadas al borrar filas/columnas, así que dejarlos combinados arriesga perder texto en
    silencio (se detectó con el banner de EBITDA). Los banners de grupo quedan como texto normal
    en una celda en vez de centrados sobre varias columnas; el contenido nunca se pierde.
"""
import os
import sys

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter, column_index_from_string

_here = os.path.dirname(os.path.abspath(__file__))
if _here not in sys.path:
    sys.path.insert(0, _here)
import render_tables_v5 as RT  # noqa: E402


def _trim_unused_rows_cols(wb, row_nums_by_tkey, log):
    """Dentro de cada hoja conservada, borra las filas/columnas que no pertenecen a ninguna tabla
    mapeada a esa hoja (ver nota de diseño en el docstring del módulo). Debe llamarse DESPUÉS de
    escribir las traducciones, usando `row_nums_by_tkey` — la resolución de filas calculada ANTES
    de traducir nada (ver nota sobre anchors más abajo), no una recalculada sobre la hoja ya
    traducida.

    IMPORTANTE sobre 'anchor': NO se puede volver a llamar RT.resolve_row_nums() sobre la hoja
    después de traducir, porque el anchor busca el texto literal en ESPAÑOL (ej. 'Ingresos',
    'Otros Ingresos (egresos)') para detectar si la plantilla corrió una fila — y la pasada de
    traducción dirigida ya sobrescribió esa misma celda con su texto en inglés. Volver a resolver
    ahí encontraría el anchor "perdido" y devolvería las filas SIN desplazar, desalineando todo el
    recorte en las plantillas donde el anchor sí corrigió un corrimiento (se detectó así: en
    dic25.xlsx las etiquetas de 'Ingresos'/'Otros ingresos(gastos)' quedaban corridas una fila).
    Por eso `row_nums_by_tkey` se calcula una sola vez, antes de tocar ninguna celda, y se reutiliza
    aquí tal cual.

    IMPORTANTE sobre merges: openpyxl NO reubica las coordenadas de un rango combinado (merge)
    cuando se borran filas/columnas — el valor de la celda sí se desplaza, pero el merge se queda
    apuntando a las coordenadas ORIGINALES, quedando desalineado del contenido real (se verificó
    con el banner de grupo de EBITDA: tras borrar columnas, el merge seguía diciendo "G1:I1"
    mientras el valor real ya había quedado en otra columna, y el texto del banner desaparecía al
    releer el archivo). Para evitar esa corrupción silenciosa, se desarman TODOS los merges de la
    hoja antes de borrar — los banners de grupo (EBITDA/VctosPtmos) quedan como texto normal en
    una sola celda en vez de centrados sobre varias columnas, pero el contenido nunca se pierde.

    IMPORTANTE sobre ancho de columna / alto de fila: por el mismo motivo, openpyxl tampoco
    reubica ws.column_dimensions/row_dimensions al borrar — el ancho de la columna A (por ejemplo,
    un margen angosto en el Excel original) se queda asignado a la letra "A" aunque ahora A
    contenga la columna de etiquetas (que necesita ser ancha), y de ahí venía el problema que
    reportó Melissa ("no se ve bien"). Se capturan los anchos/altos ORIGINALES por número de
    columna/fila antes de borrar, y se reasignan después a la posición compactada que le
    corresponda a cada columna/fila conservada (las que no tenían un ancho/alto propio quedan con
    el default de la hoja, que es lo normal)."""
    sheets_specs = {}
    for tkey, spec in RT.TABLES.items():
        sheets_specs.setdefault(spec['sheet'], []).append((tkey, spec))

    for sheet_name, tkey_specs in sheets_specs.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]

        keep_rows = set()
        keep_cols = set()
        for tkey, spec in tkey_specs:
            if tkey not in row_nums_by_tkey:
                continue
            keep_rows.update(row_nums_by_tkey[tkey])
            keep_cols.update(c for c, _ in spec['cols'])

        if not keep_rows or not keep_cols:
            continue

        # Captura los anchos/altos ORIGINALES (por número de columna/fila) antes de borrar nada —
        # ver nota de diseño arriba sobre por qué hay que remapearlos a mano después.
        orig_col_widths = {}
        for letter, dim in list(ws.column_dimensions.items()):
            if dim.width is not None:
                try:
                    orig_col_widths[column_index_from_string(letter)] = dim.width
                except ValueError:
                    continue
        orig_row_heights = {idx: dim.height for idx, dim in list(ws.row_dimensions.items())
                            if dim.height is not None}

        for rng in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(rng))

        rows_to_delete = sorted((r for r in range(1, ws.max_row + 1) if r not in keep_rows), reverse=True)
        cols_to_delete = sorted((c for c in range(1, ws.max_column + 1) if c not in keep_cols), reverse=True)

        for r in rows_to_delete:
            ws.delete_rows(r, 1)
        for c in cols_to_delete:
            ws.delete_cols(c, 1)

        # Reasigna los anchos/altos capturados a la posición NUEVA (compactada) de cada columna/
        # fila conservada, y limpia cualquier entrada vieja que haya quedado apuntando a una
        # posición que ya no corresponde a esa columna/fila.
        for letter in list(ws.column_dimensions.keys()):
            del ws.column_dimensions[letter]
        for idx in list(ws.row_dimensions.keys()):
            del ws.row_dimensions[idx]
        for new_idx, orig_idx in enumerate(sorted(keep_cols), start=1):
            if orig_idx in orig_col_widths:
                ws.column_dimensions[get_column_letter(new_idx)].width = orig_col_widths[orig_idx]
        for new_idx, orig_idx in enumerate(sorted(keep_rows), start=1):
            if orig_idx in orig_row_heights:
                ws.row_dimensions[new_idx].height = orig_row_heights[orig_idx]

        if rows_to_delete or cols_to_delete:
            log.append(f"[OK] Hoja '{sheet_name}': se quitaron {len(rows_to_delete)} filas y "
                       f"{len(cols_to_delete)} columnas fuera de la tabla, y se reajustó el ancho "
                       f"de columna / alto de fila a la posición compactada")


def translate_workbook_to_en(src_path, out_path, log=None):
    """Lee el Excel en español `src_path` y guarda una copia traducida al inglés en `out_path`.
    Devuelve la lista de mensajes de log (misma convención [OK]/[AVISO]/[ERROR] que assemble.py)."""
    if log is None:
        log = []

    try:
        wb = load_workbook(src_path, data_only=True)
    except Exception as ex:
        log.append(f"[ERROR] No se pudo abrir el Excel para traducir: {ex}")
        return log

    # Mismos parches de nombre de hoja que usa assemble.collect_and_render_tables(), para que
    # TABLES_EN encuentre la hoja correcta en plantillas donde el nombre de pestaña varía.
    RT.TABLES['IngresosxSegmento']['sheet'] = 'IngresosxSegmento'
    if 'Indicadores1' not in wb.sheetnames and 'Ind. Financieros' in wb.sheetnames:
        RT.TABLES['Indicadores1']['sheet'] = 'Ind. Financieros'

    # 0) Deja solo las hojas que alimentan alguna de las 18 tablas del Word (ver nota de diseño
    #    arriba) — se hace ANTES de traducir para no gastar tiempo traduciendo hojas que de
    #    todas formas se van a eliminar.
    used_sheets = {spec['sheet'] for spec in RT.TABLES.values()}
    removed_sheets = [name for name in wb.sheetnames if name not in used_sheets]
    for name in removed_sheets:
        wb.remove(wb[name])
    if removed_sheets:
        log.append(f"[OK] Se quitaron {len(removed_sheets)} hojas que no se usan en el Word: "
                   + ', '.join(removed_sheets))

    # Resuelve y CONGELA la fila real de cada tabla (incluido el ajuste por 'anchor', que busca
    # texto literal en español) ANTES de traducir ninguna celda — ver nota en
    # _trim_unused_rows_cols sobre por qué no se puede recalcular esto después.
    row_nums_by_tkey = {}
    for tkey, spec in RT.TABLES.items():
        sheet_name = spec['sheet']
        if sheet_name in wb.sheetnames:
            row_nums_by_tkey[tkey] = RT.resolve_row_nums(wb[sheet_name], spec)

    # 1) Pasada genérica sobre todas las hojas/celdas de texto (ya solo quedan las usadas).
    n_generic = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                v = cell.value
                if isinstance(v, str) and v.strip():
                    new_v = RT._translate_common_strings_en(v)
                    if new_v != v:
                        cell.value = new_v
                        n_generic += 1
    log.append(f"[OK] Traducción genérica (fechas, MUS$, veces, Días, variación...): {n_generic} celdas")

    # 2) Pasada dirigida: etiquetas de cada tabla (TABLES_EN), ya verificadas contra el Excel/Word
    #    en inglés de referencia.
    n_tables = 0
    n_tables_skipped = 0
    for tkey, spec in RT.TABLES.items():
        over = RT.TABLES_EN.get(tkey)
        if not over:
            continue
        sheet_name = spec['sheet']
        if sheet_name not in wb.sheetnames or tkey not in row_nums_by_tkey:
            n_tables_skipped += 1
            continue
        ws = wb[sheet_name]
        row_nums = row_nums_by_tkey[tkey]
        col_defs = spec['cols']

        for (ri, ci), text in over.get('label_overrides', {}).items():
            if ri >= len(row_nums) or ci >= len(col_defs):
                continue
            r = row_nums[ri]
            c = col_defs[ci][0]
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell):
                continue
            cell.value = text
            n_tables += 1

        for (gh_ri, gh_c1, gh_c2, gh_txt) in over.get('group_headers', []):
            if gh_ri >= len(row_nums) or gh_c1 >= len(col_defs):
                continue
            r = row_nums[gh_ri]
            c = col_defs[gh_c1][0]
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell):
                continue
            # En el Excel los banners de grupo usan espacios, no saltos de línea (ver celda
            # original "AÑO CALENDARIO             (6 meses)" en EBITDA) — se reproduce igual.
            cell.value = ' '.join(gh_txt.replace('\n', ' ').split())
            n_tables += 1

    log.append(f"[OK] Traducción de etiquetas de tabla: {n_tables} celdas en {len(RT.TABLES_EN)} tablas"
               + (f" ({n_tables_skipped} tablas omitidas: hoja no presente en este Excel)" if n_tables_skipped else ""))

    # 3) Recorta cada hoja a las filas/columnas que realmente pertenecen a su tabla (se hace
    #    AL FINAL, reutilizando row_nums_by_tkey calculado ANTES de traducir — ver nota arriba).
    _trim_unused_rows_cols(wb, row_nums_by_tkey, log)

    try:
        wb.save(out_path)
        log.append(f"[OK] Excel en inglés guardado: {os.path.basename(out_path)}")
    except Exception as ex:
        log.append(f"[ERROR] No se pudo guardar el Excel traducido: {ex}")

    return log


if __name__ == '__main__':
    src, out = sys.argv[1], sys.argv[2]
    log = translate_workbook_to_en(src, out)
    for line in log:
        print(line)
