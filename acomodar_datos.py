import pandas as pd

def acomodar_datos(input_path, output_path):
    try:
        xls = pd.ExcelFile(input_path)
        df = pd.read_excel(xls, sheet_name="Informe", dtype=str)

        if df is None:
            print("❌ Error: No se encontró la hoja 'Informe'.")
            return None

        # 🔹 Limpiar nombres de columnas
        df.columns = df.columns.str.strip()

        # 🔹 Validar columnas esenciales (sin 'COM. SEGUROS')
        required_columns = [
            "DESC.PRODUCTO", "PRECIO", "C. FINANCIAM", "COM. G. EXT.",
            "COM. ACCS", "BONO", "CXA  (BBVA)", "COM. POR TOM"
        ]
        for col in required_columns:
            if col not in df.columns:
                print(f"⚠️ Advertencia: La columna obligatoria '{col}' no existe.")
                return None

        # 🔍 Detectar columna que se parezca a "COM. SEGUROS"
        columna_com_seguros = next((c for c in df.columns if "COM. SEGUROS" in c.upper().replace(" ", " ").strip()), None)
        existe_com_seguros = columna_com_seguros is not None
        existe_vf3 = "VF3" in df.columns

        for row_idx in range(len(df)):
            desc_values = df.at[row_idx, "DESC.PRODUCTO"].split(" | ") if isinstance(df.at[row_idx, "DESC.PRODUCTO"], str) else []
            precio_values = df.at[row_idx, "PRECIO"].split(" | ") if isinstance(df.at[row_idx, "PRECIO"], str) else []
            precio_values = [p for p in precio_values if p.strip()]

            if "INCENTIVO DEALER" in desc_values:
                index = desc_values.index("INCENTIVO DEALER")
                if index < len(precio_values):
                    df.at[row_idx, "C. FINANCIAM"] = precio_values[index]

            for etiqueta, columna in [
                ("COMISION GARANTIA EXTENDIDA", "COM. G. EXT."),
                ("COMISIÓN ACCESORIOS REFACCIONE", "COM. ACCS"),
                ("AUTO ADQUIRIDO POR SELECTIVITY", "BONO")
            ]:
                for i, desc in enumerate(desc_values):
                    if etiqueta in desc and i < len(precio_values):
                        df.at[row_idx, columna] = precio_values[i]

            for i, desc in enumerate(desc_values):
                if "COMISION VF3" in desc and i < len(precio_values):
                    destino = "VF3" if existe_vf3 else columna_com_seguros if existe_com_seguros else None
                    if destino:
                        df.at[row_idx, destino] = precio_values[i]

        # 🔄 Renombrar COM. SEGUROS → VF3 si se usó
        if not existe_vf3 and existe_com_seguros and columna_com_seguros in df.columns:
            df.rename(columns={columna_com_seguros: "VF3"}, inplace=True)

        columnas_numericas = [
            "C. FINANCIAM", "COM. G. EXT.", "COM. ACCS", "BONO",
            "CXA  (BBVA)", "COM. POR TOM", "COMISION", "TOTAL.", "VF3"
        ]
        for col in columnas_numericas:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')

        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="Informe", index=False)
            worksheet = writer.sheets["Informe"]
            for idx, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(idx, idx, max_len)

        print(f"✅ Datos acomodados correctamente en {output_path}")
        return output_path

    except Exception as e:
        print(f"❌ Error al acomodar datos: {e}")
        return None
