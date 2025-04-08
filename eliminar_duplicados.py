import pandas as pd

def eliminar_duplicados(input_path, output_path):
    try:
        xls = pd.ExcelFile(input_path)
        df = pd.read_excel(xls, sheet_name="Informe", dtype=str)

        if df is None:
            print("❌ Error: No se encontró la hoja 'Informe'.")
            return None

        # 🔹 Reemplazar NaN explícitamente por cadenas vacías
        df.fillna("", inplace=True)

        # 🔹 **Eliminar columnas innecesarias**
        columnas_a_eliminar = ["C#NO REPUVE.1", "C#CONSTANCIA.1", "COMISION 6%", "COMISION 8%", "COMISION 10%", "COMISION 12%"]
        df.drop(columns=[col for col in columnas_a_eliminar if col in df.columns], inplace=True, errors='ignore')

        # 🔹 **Columnas donde NO se deben eliminar valores duplicados**
        columnas_excluidas = ["C. FINANCIAM", "COM. G. EXT.", "COM. ACCS", "VF3", "BONO"]

        # 🔹 **Hacer una copia de las columnas antes de eliminar duplicados**
        df_original = df.copy()

        # 🔹 **Eliminar duplicados en VIN y REFERENCIA**
        if {"VIN", "REFERENCIA"}.issubset(df.columns):
            df = df.drop_duplicates(subset=["VIN", "REFERENCIA"], keep='first')

        # 🔹 **Restaurar los datos en las columnas excluidas SOLO donde estén vacías**
        for col in columnas_excluidas:
            if col in df.columns:
                df[col] = df[col].where(df[col] != "", df_original[col])

        # 🔹 **Reemplazar valores erróneos**
        df.replace("2.69653970229347E+308", "", inplace=True)

        # 🔹 **Asegurar que los valores copiados no se eliminen**
        df = df.applymap(lambda x: "" if (pd.isna(x) or str(x).strip() == "") else x)

        # 🔹 **Guardar el archivo ajustando la estética**
        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="Informe", index=False)
            workbook = writer.book
            worksheet = writer.sheets["Informe"]

            for idx, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(idx, idx, max_len)

        print(f"✅ Duplicados eliminados sin afectar datos en {output_path}")
        return output_path

    except Exception as e:
        print(f"❌ Error al eliminar duplicados: {e}")
        return None
