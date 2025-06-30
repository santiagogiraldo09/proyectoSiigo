import pandas as pd
import requests
from datetime import datetime
import time # Para añadir un pequeño retraso entre peticiones a la API

# --- Tu función para obtener la TRM ---
def get_trm_from_datos_abiertos(date_str):
    """
    Consulta la TRM desde la API de Datos Abiertos Colombia (Socrata).

    Args:
        date_str (str): La fecha en formato 'YYYY-MM-DD'.

    Returns:
        float or None: El valor de la TRM, o None si no se encuentra o hay un error.
    """
    BASE_URL = "https://www.datos.gov.co/resource/mcec-87by.json"
    params = {
        "vigenciadesde": f"{date_str}T00:00:00.000"
    }

    try:
        response = requests.get(BASE_URL, params=params, timeout=10) # Añadir un timeout
        response.raise_for_status()

        data = response.json()

        if data:
            return float(data[0].get('valor'))
        else:
            return None
    except requests.exceptions.RequestException as e:
        print(f"Error de conexión o HTTP al consultar Datos Abiertos para {date_str}: {e}")
        return None
    except (ValueError, IndexError, TypeError) as e:
        print(f"Error al parsear o acceder a los datos de la TRM para {date_str}: {e}")
        return None
    except Exception as e:
        print(f"Error inesperado al consultar Datos Abiertos para {date_str}: {e}")
        return None


def procesar_y_guardar_excel_completo(ruta_archivo_entrada, nombres_columnas_a_eliminar, ruta_archivo_salida):
    """
    Lee un archivo de Excel, elimina filas con 'Tipo clasificación' vacío,
    elimina múltiples columnas, actualiza la columna 'Total',
    y rellena 'Tasa de cambio' con TRM de API si es necesario,
    luego guarda el resultado.

    Args:
        ruta_archivo_entrada (str): La ruta completa al archivo de Excel original.
        nombres_columnas_a_eliminar (list): Una lista con los nombres de las columnas que se desean eliminar.
        ruta_archivo_salida (str): La ruta completa donde se guardará el nuevo archivo de Excel.
    """
    try:
        # Leer el archivo de Excel
        df = pd.read_excel(ruta_archivo_entrada)
        print(f"Archivo '{ruta_archivo_entrada}' leído. Filas iniciales: {len(df)}")

        # 1. Eliminar filas donde "Tipo clasificación" esté vacío/NaN
        if "Tipo clasificación" in df.columns:
            filas_antes_eliminacion = len(df)
            df_procesado = df.dropna(subset=["Tipo clasificación"])
            filas_despues_eliminacion = len(df_procesado)
            print(f"Filas con 'Tipo clasificación' vacío eliminadas: {filas_antes_eliminacion - filas_despues_eliminacion}. Filas restantes: {filas_despues_eliminacion}")
        else:
            df_procesado = df.copy()
            print("La columna 'Tipo clasificación' no se encontró. No se eliminaron filas vacías.")

        # 2. Eliminar columnas especificadas
        columnas_existentes_para_eliminar = [col for col in nombres_columnas_a_eliminar if col in df_procesado.columns]
        columnas_no_existentes_para_eliminar = [col for col in nombres_columnas_a_eliminar if col not in df_procesado.columns]

        if columnas_existentes_para_eliminar:
            df_procesado = df_procesado.drop(columns=columnas_existentes_para_eliminar)
            print(f"Las columnas {columnas_existentes_para_eliminar} han sido eliminadas.")
        else:
            print("Ninguna de las columnas especificadas para eliminar se encontró. No se eliminaron columnas.")

        if columnas_no_existentes_para_eliminar:
            print(f"Advertencia: Las siguientes columnas especificadas para eliminación no se encontraron: {columnas_no_existentes_para_eliminar}")

        # 3. Actualizar la columna "Total" existente
        if "Cantidad" in df_procesado.columns and "Valor unitario" in df_procesado.columns and "Total" in df_procesado.columns:
            df_procesado["Cantidad"] = pd.to_numeric(df_procesado["Cantidad"], errors='coerce')
            df_procesado["Valor unitario"] = pd.to_numeric(df_procesado["Valor unitario"], errors='coerce')
            df_procesado["Total"] = df_procesado["Cantidad"] * df_procesado["Valor unitario"]
            df_procesado["Total"] = df_procesado["Total"].fillna(0)
            print("La columna 'Total' ha sido actualizada con el cálculo 'Cantidad * Valor unitario'.")
        else:
            print("Advertencia: No se pudieron encontrar las columnas 'Cantidad', 'Valor unitario' y/o 'Total'. La columna 'Total' no se actualizó.")

        # 4. Rellenar celdas vacías o con 0 en "Tasa de cambio" con la TRM
        if "Tasa de cambio" in df_procesado.columns and "Fecha elaboración" in df_procesado.columns:
            print("Iniciando el proceso de rellenado de 'Tasa de cambio'...")
            
            # Convertir "Fecha elaboración" a formato de fecha de Pandas y luego a datetime para manipularla
            df_procesado['Fecha elaboración_dt'] = pd.to_datetime(df_procesado['Fecha elaboración'], format='%d/%m/%Y', errors='coerce')

            # Iterar sobre las filas del DataFrame para buscar y aplicar la TRM
            for index, row in df_procesado.iterrows():
                tasa_actual = row["Tasa de cambio"]
                fecha_elaboracion_dt = row["Fecha elaboración_dt"]

                if (pd.isna(tasa_actual) or tasa_actual == 0) and pd.notna(fecha_elaboracion_dt):
                    fecha_str_api = fecha_elaboracion_dt.strftime('%Y-%m-%d')
                    print(f"Buscando TRM para la fecha: {fecha_str_api} (Fila {index})...")
                    trm_valor = get_trm_from_datos_abiertos(fecha_str_api)

                    if trm_valor is not None:
                        df_procesado.at[index, "Tasa de cambio"] = trm_valor
                        print(f"  > TRM encontrada: {trm_valor}")
                    else:
                        print(f"  > No se pudo obtener TRM para {fecha_str_api}. La celda permanecerá sin cambios.")
                    
                    time.sleep(0.1) # Pequeña pausa para no sobrecargar la API

            df_procesado = df_procesado.drop(columns=['Fecha elaboración_dt'])
            print("Proceso de rellenado de 'Tasa de cambio' completado.")
        else:
            print("Advertencia: No se encontraron las columnas 'Tasa de cambio' y/o 'Fecha elaboración'. No se buscó la TRM.")

        # 5. Guardar el DataFrame modificado en un nuevo archivo Excel
        df_procesado.to_excel(ruta_archivo_salida, index=False)
        print(f"El archivo final ha sido guardado en: {ruta_archivo_salida}")

    except FileNotFoundError:
        print(f"Error: El archivo de entrada no se encontró en la ruta: {ruta_archivo_entrada}")
    except Exception as e:
        print(f"Se produjo un error durante el procesamiento: {e}")


if __name__ == "__main__":
    # Define la ruta de tu archivo de entrada
    # Reemplaza 'tu_archivo.xlsx' con el nombre real y la ruta completa.
    archivo_entrada = "C:/Users/Lenovo Thinkpad E14/Downloads/ARCHIVO SIIGO.xlsx"

    # Define la lista de nombres de las columnas a eliminar
    columnas_a_eliminar = [
        "Nombre tercero",
        "Código",
        "Consecutivo",
        "Tipo transacción"
    ]

    # Define la ruta y el nombre para el nuevo archivo de salida
    archivo_salida = "C:/Users/Lenovo Thinkpad E14/Downloads/archivo_siigo_modificado.xlsx"

    # Llama a la función principal para procesar y guardar el Excel
    procesar_y_guardar_excel_completo(archivo_entrada, columnas_a_eliminar, archivo_salida)
