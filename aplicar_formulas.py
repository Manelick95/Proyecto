import xlrd
import xlwt
from xlutils.copy import copy


def aplicar_formulas(input_path, output_path):
    try:
        # 🔹 Abrir el archivo
        rb = xlrd.open_workbook(input_path)
        wb = copy(rb)
        sheet = wb.get_sheet(0)
        num_rows = rb.sheet_by_index(0).nrows
        num_cols = rb.sheet_by_index(0).ncols

        # 🔹 Obtener encabezados
        headers = [rb.sheet_by_index(0).cell_value(0, col) for col in range(num_cols)]
        try:
            utilidad_index = headers.index("UTILIDAD")
            porcentaje_index = headers.index("PORCENTAJE")
            comision_index = headers.index("COMISION")
            total_column_index = headers.index("TOTAL.")
            financiam_index = headers.index("C. FINANCIAM")
            cxa_bbva_index = headers.index("CXA  (BBVA)")
            com_g_ext_index = headers.index("COM. G. EXT.")
            com_accs_index = headers.index("COM. ACCS")
            com_seguros_index = headers.index("VF3")
            com_satfind_index = headers.index("COM. SATFIND")
            bastidor_index = headers.index("BASTIDOR . VO RECOG")
            com_por_tom_index = headers.index("COM. POR TOM")
            bono_index = headers.index("BONO")
            nivel_index = headers.index("NIVEL")
        except ValueError as e:
            print(f"❌ Error: No se encontraron las columnas necesarias. {e}")
            return None

        # 🔹 Aplicar fórmula a columna COMISION (UTILIDAD * PORCENTAJE)
        for row_idx in range(1, num_rows):
            util_cell = xlwt.Utils.rowcol_to_cell(row_idx, utilidad_index)
            porc_cell = xlwt.Utils.rowcol_to_cell(row_idx, porcentaje_index)
            sheet.write(row_idx, comision_index, f"={util_cell}*{porc_cell}")

        # 🔹 Aplicar lógica a C. FINANCIAM según el NIVEL
        for row_idx in range(1, num_rows):
            nivel = rb.sheet_by_index(0).cell_value(row_idx, nivel_index).strip().upper()
            valor_actual = rb.sheet_by_index(0).cell_value(row_idx, financiam_index)

            if valor_actual not in ["", None]:
                try:
                    valor_numerico = float(valor_actual)
                    if nivel == "BASICO":
                        sheet.write(row_idx, financiam_index, "NA")
                    elif nivel == "CONFORT":
                        sheet.write(row_idx, financiam_index, valor_numerico / 2)
                    else:
                        sheet.write(row_idx, financiam_index, valor_numerico)
                except ValueError:
                    sheet.write(row_idx, financiam_index, valor_actual)

        # 🔹 Aplicar fórmula TOTAL = SUM(...)
        for row_idx in range(1, num_rows):
            total_formula = f"=SUM({xlwt.Utils.rowcol_to_cell(row_idx, comision_index)}, {xlwt.Utils.rowcol_to_cell(row_idx, financiam_index)}, {xlwt.Utils.rowcol_to_cell(row_idx, cxa_bbva_index)}, {xlwt.Utils.rowcol_to_cell(row_idx, com_g_ext_index)}, {xlwt.Utils.rowcol_to_cell(row_idx, com_accs_index)}, {xlwt.Utils.rowcol_to_cell(row_idx, com_seguros_index)}, {xlwt.Utils.rowcol_to_cell(row_idx, com_satfind_index)}, {xlwt.Utils.rowcol_to_cell(row_idx, bastidor_index)}, {xlwt.Utils.rowcol_to_cell(row_idx, com_por_tom_index)}, {xlwt.Utils.rowcol_to_cell(row_idx, bono_index)})"
            sheet.write(row_idx, total_column_index, total_formula)

        wb.save(output_path)
        print(f"✅ Fórmulas aplicadas correctamente en {output_path}")
        return output_path

    except Exception as e:
        print(f"❌ Error al aplicar fórmulas: {e}")
        return None
