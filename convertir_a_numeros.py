import pandas as pd
import xlsxwriter

def convertir_a_numeros(output_path):
    try:
        print("🔹 Forzando formato de número en columnas clave...")

        # Lista base (sin 'COM. SEGUROS' que puede haber sido renombrada)
        columnas_numericas = [
            "COMISION", "C. FINANCIAM", "CXA  (BBVA)", "COM. G. EXT.",
            "COM. ACCS", "COM. SATFIND", "BASTIDOR . VO RECOG",
            "COM. POR TOM", "BONO", "TOTAL.", "VF3"  # ← se usa "VF3" directamente
        ]

        df_final = pd.read_excel(output_path, sheet_name="Informe")

        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            df_final.to_excel(writer, sheet_name="Informe", index=False)
            workbook = writer.book
            worksheet = writer.sheets["Informe"]
            formato_numero = workbook.add_format({'num_format': '#,##0.00'})

            for col_idx, col_name in enumerate(df_final.columns):
                if col_name in columnas_numericas:
                    worksheet.set_column(col_idx, col_idx, 12, formato_numero)
                else:
                    max_len = max(df_final[col_name].astype(str).map(len).max(), len(col_name)) + 2
                    worksheet.set_column(col_idx, col_idx, max_len)

        print(f"✅ Columnas convertidas a número correctamente en {output_path}")
        return output_path

    except Exception as e:
        print(f"❌ Error al convertir celdas: {e}")
        return None
