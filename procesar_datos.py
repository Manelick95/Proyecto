import pandas as pd
import xlrd
import xlwt
from xlutils.copy import copy
from eliminar_duplicados import eliminar_duplicados
from aplicar_formulas import aplicar_formulas
from acomodar_datos import acomodar_datos
from convertir_a_numeros import convertir_a_numeros
from datetime import datetime
from dateutil.relativedelta import relativedelta
import re

def procesar_archivo(file1_path, file2_path, output_path):
    try:
        rb = xlrd.open_workbook(file1_path, formatting_info=True)
        wb = copy(rb)
        sheets = {sheet: pd.read_excel(file1_path, sheet_name=sheet, dtype=str) for sheet in rb.sheet_names()}
        df_informe = sheets.get("Informe")

        if df_informe is None:
            print("❌ Error: No se encontró la hoja 'Informe'.")
            return None

        df_informe.columns = df_informe.columns.str.strip()
        df_informe["FEC.FACT"] = pd.to_datetime(df_informe["FEC.FACT"], errors='coerce', dayfirst=True)
        df_informe["VENDEDOR"] = df_informe["VENDEDOR"].astype(str).str.strip().str.upper()

        porcentaje_map = {
            "BASICO": "6%",
            "CONFORT": "10%",
            "EXTRAMILLA": "12%",
            "HUERPEL": "15%"
        }

        # 🟩 Mapeo CONCESI → hoja de archivo de niveles
        mapeo_concesis = {
            "1": "PACHUCA",
            "19": "TIZAYUCA",
            "2": "TULANCINGO",
            "3": "HUAUCHINANGO"
        }

        xls_niveles = pd.ExcelFile(file2_path)

        for concesi_codigo, nombre_clave in mapeo_concesis.items():
            hoja_encontrada = next((s for s in xls_niveles.sheet_names if nombre_clave in s.upper()), None)
            if hoja_encontrada is None:
                print(f"❌ No se encontró la hoja correspondiente a CONCESI {concesi_codigo} con nombre '{nombre_clave}'")
                continue

            df_niveles = pd.read_excel(xls_niveles, sheet_name=hoja_encontrada, header=None, dtype=str)

            # Detectar secciones "KPI’S MES ..."
            secciones = []
            for i, row in df_niveles.iterrows():
                for cell in row:
                    if isinstance(cell, str) and "KPI" in cell.upper():
                        match = re.search(r"(ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SEPTIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE)\s+(\d{4})", cell.upper())
                        if match:
                            mes_str, anio_str = match.groups()
                            meses_es = {
                                "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
                                "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12
                            }
                            mes = meses_es.get(mes_str)
                            anio = int(anio_str)
                            fecha_objetivo = datetime(anio, mes, 1) - relativedelta(months=1)
                            secciones.append({
                                "fila_inicio": i + 1,
                                "mes": fecha_objetivo.month,
                                "anio": fecha_objetivo.year
                            })
                        break

            if not secciones:
                print(f"⚠️ No se encontraron secciones en hoja {hoja_encontrada}")
                continue

            for idx in range(len(secciones)):
                if idx < len(secciones) - 1:
                    secciones[idx]["fila_fin"] = secciones[idx + 1]["fila_inicio"] - 1
                else:
                    secciones[idx]["fila_fin"] = len(df_niveles)

            for seccion in secciones:
                inicio = seccion["fila_inicio"]
                fin = seccion["fila_fin"]
                mes_obj = seccion["mes"]
                anio_obj = seccion["anio"]

                df_tabla_raw = df_niveles.iloc[inicio:fin].dropna(how='all', axis=1)
                if df_tabla_raw.empty:
                    continue

                df_tabla_raw.columns = df_tabla_raw.iloc[0]
                df_tabla = df_tabla_raw[1:].copy()
                df_tabla.columns = [str(col).strip() for col in df_tabla.columns]

                col_vendedor = next((c for c in df_tabla.columns if "VENDEDOR" in c.upper()), None)
                col_nivel = next((c for c in df_tabla.columns if "NIVEL" in c.upper()), None)

                if not col_vendedor or not col_nivel:
                    print(f"⚠️ Columnas 'VENDEDOR' o 'NIVEL' ausentes en filas {inicio}-{fin}")
                    continue

                try:
                    df_tabla = df_tabla[[col_vendedor, col_nivel]]
                    df_tabla.columns = ["VENDEDOR", "NIVEL"]
                    df_tabla["VENDEDOR"] = df_tabla["VENDEDOR"].astype(str).str.strip().str.upper()
                    df_tabla["NIVEL"] = df_tabla["NIVEL"].astype(str).str.strip().str.upper()
                    niveles_dict = df_tabla.set_index("VENDEDOR")["NIVEL"].to_dict()

                    registros = df_informe[
                        (df_informe["FEC.FACT"].dt.month == mes_obj) &
                        (df_informe["FEC.FACT"].dt.year == anio_obj) &
                        (df_informe["CONCESI"] == concesi_codigo)
                    ]

                    df_informe.loc[registros.index, "NIVEL"] = registros["VENDEDOR"].map(niveles_dict).fillna("")

                except Exception as e:
                    print(f"⚠️ Error procesando tabla en filas {inicio}-{fin}: {e}")
                    continue

        df_informe["NIVEL"] = df_informe["NIVEL"].fillna("").str.upper()
        df_informe["PORCENTAJE"] = df_informe["NIVEL"].map(porcentaje_map).fillna("")
        df_informe = df_informe.loc[:, ~df_informe.columns.duplicated()]

        for col in df_informe.columns:
            if "fec" in col.lower() or "fecha" in col.lower():
                df_informe[col] = pd.to_datetime(df_informe[col], errors='coerce', dayfirst=True).dt.strftime('%d/%m/%Y')

        df_informe.replace("2.69653970229347E+308", "", inplace=True)

        if "PRECIO" in df_informe.columns:
            df_informe["PRECIO"] = df_informe["PRECIO"].astype(str).replace(["nan", "NaN", "None"], "")
            df_informe["PRECIO"] = df_informe["PRECIO"].str.replace(r'\|\s*\|', '|', regex=True)
            df_informe["PRECIO"] = df_informe["PRECIO"].str.strip('|')

        columnas_sin_duplicados = ["BASTIDOR . VO RECOG", "IMP PENDIE", "Cta.personal", "FECHACER", "C#NO REPUVE"]
        for col in columnas_sin_duplicados:
            if col in df_informe.columns:
                df_informe[col] = df_informe[col].mask(df_informe[col].duplicated(), "")

        if "VENDEDOR" in df_informe.columns and "Curp" in df_informe.columns:
            vendedores_curp = df_informe[df_informe["Curp"].notna()].set_index("VENDEDOR")["Curp"].to_dict()
            df_informe["Curp"] = df_informe["VENDEDOR"].map(vendedores_curp)
            df_informe.loc[df_informe["VENDEDOR"] == "INTERCAMBIOS TULANCINGO", "Curp"] = ""

        columnas_vacias_originales = df_informe.columns[df_informe.isna().all()].tolist()
        columnas_a_rellenar = [col for col in df_informe.columns if col not in columnas_vacias_originales]
        df_informe[columnas_a_rellenar] = df_informe[columnas_a_rellenar].ffill().bfill()
        df_informe[columnas_vacias_originales] = ""

        if {"PRODUCTO", "DESC.PRODUCTO", "PRECIO", "VIN", "REFERENCIA"}.issubset(df_informe.columns):
            df_informe["PRODUCTO"] = df_informe.groupby(["VIN", "REFERENCIA"])["PRODUCTO"].transform(lambda x: " | ".join(x.dropna().unique()))
            df_informe["DESC.PRODUCTO"] = df_informe.groupby(["VIN", "REFERENCIA"])["DESC.PRODUCTO"].transform(lambda x: " | ".join(x.dropna().unique()))
            df_informe["PRECIO"] = df_informe.groupby(["VIN", "REFERENCIA"])["PRECIO"].transform(lambda x: " | ".join(x.dropna().unique()))

        for idx, sheet_name in enumerate(rb.sheet_names()):
            sheet = wb.get_sheet(idx)
            if sheet_name == "Informe":
                for col_idx, col_name in enumerate(df_informe.columns):
                    sheet.write(0, col_idx, col_name)
                for row_idx, row in enumerate(df_informe.itertuples(index=False), start=1):
                    for col_idx, value in enumerate(row):
                        if isinstance(value, str) and value.replace('.', '', 1).isdigit():
                            sheet.write(row_idx, col_idx, round(float(value), 2))
                        else:
                            sheet.write(row_idx, col_idx, value)

        wb.save(output_path)
        print(f"✅ Archivo procesado correctamente en {output_path}")
        eliminar_duplicados(output_path, output_path)
        acomodar_datos(output_path, output_path)
        aplicar_formulas(output_path, output_path)
        convertir_a_numeros(output_path)

        return output_path

    except Exception as e:
        print(f"❌ Error al procesar archivo: {e}")
        return None
