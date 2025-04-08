import pandas as pd

def ajustar_estetica(input_path, output_path):
    try:
        xls = pd.ExcelFile(input_path)
        df = pd.read_excel(xls, sheet_name="Informe", dtype=str)

        if df is None:
            print("❌ Error: No se encontró la hoja 'Informe'.")
            return None

        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="Informe", index=False)
            workbook = writer.book
            worksheet = writer.sheets["Informe"]

            # 🔹 **Autoajustar el ancho de las columnas**
            for idx, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(idx, idx, max_len)

        print(f"✅ Estética ajustada en {output_path}")
        return output_path

    except Exception as e:
        print(f"❌ Error al ajustar estética: {e}")
        return None
