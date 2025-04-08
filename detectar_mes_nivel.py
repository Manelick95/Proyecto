import pandas as pd
import calendar

def obtener_mes_anterior_desde_niveles(file_path):
    try:
        # Leer el archivo sin encabezado, porque buscamos en cualquier parte del archivo
        df = pd.read_excel(file_path, header=None, dtype=str)

        # Buscar en cada celda la palabra "KPI" y un mes en español
        for row in df[0]:
            if isinstance(row, str) and "KPI" in row.upper():
                # Lista de meses en español
                meses_es = {
                    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4,
                    "MAYO": 5, "JUNIO": 6, "JULIO": 7, "AGOSTO": 8,
                    "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12
                }

                partes = row.upper().split()
                mes_actual = None
                anio = None

                for parte in partes:
                    if parte in meses_es:
                        mes_actual = meses_es[parte]
                    elif parte.isdigit() and len(parte) == 4:
                        anio = int(parte)

                if mes_actual and anio:
                    if mes_actual == 1:
                        return (12, anio - 1)  # Diciembre del año anterior
                    else:
                        return (mes_actual - 1, anio)  # Mes anterior del mismo año

        print("❌ Error: No se pudo detectar el mes y año desde el archivo de niveles.")
        return None

    except Exception as e:
        print(f"❌ Error detectando mes/año: {e}")
        return None
