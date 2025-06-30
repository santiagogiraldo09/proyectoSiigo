import streamlit as st
import pandas as pd
import requests
from datetime import datetime
import time
import io

# --- Tu función para obtener la TRM (sin cambios mayores) ---
@st.cache_data(ttl=3600) # Almacena en caché los resultados de la TRM por 1 hora para evitar peticiones repetidas
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
        response = requests.get(BASE_URL, params=params, timeout=10)
        response.raise_for_status()

        data = response.json()

        if data:
            return float(data[0].get('valor'))
        else:
            return None
    except requests.exceptions.RequestException as e:
        # st.warning(f"Error de conexión o HTTP al consultar Datos Abiertos para {date_str}: {e}")
        return None
    except (ValueError, IndexError, TypeError) as e:
        # st.warning(f"Error al parsear o acceder a los datos de la TRM para {date_str}: {e}")
        return None
    except Exception as e:
        # st.warning(f"Error inesperado al consultar Datos Abiertos para {date_str}: {e}")
        return None

# --- Función Principal de Procesamiento (adaptada para Streamlit) ---
def procesar_excel_para_streamlit(uploaded_file):
    """
    Procesa el archivo de Excel subido:
    - Elimina filas con 'Tipo clasificación' vacío.
    - Elimina columnas no deseadas.
    - Actualiza la columna 'Total'.
    - Rellena 'Tasa de cambio' con TRM de API.

    Args:
        uploaded_file (streamlit.UploadedFile): El archivo Excel subido por el usuario.

    Returns:
        pandas.DataFrame or None: El DataFrame procesado o None si hay un error.
    """
    try:
        # Leer el archivo de Excel desde el objeto de archivo subido
        # Pandas puede leer directamente desde el objeto BytesIO de Streamlit
        df = pd.read_excel(uploaded_file)
        st.info(f"Archivo cargado exitosamente. Filas iniciales: **{len(df)}**.")

        # Columnas a eliminar predefinidas (puedes hacerlas configurables en Streamlit si lo deseas)
        nombres_columnas_a_eliminar = [
            "Nombre tercero",
            "Código",
            "Consecutivo",
            "Tipo transacción"
        ]

        df_procesado = df.copy() # Usamos una copia para no modificar el DataFrame original si hay errores

        # 1. Eliminar filas donde "Tipo clasificación" esté vacío/NaN
        if "Tipo clasificación" in df_procesado.columns:
            filas_antes_eliminacion = len(df_procesado)
            df_procesado.dropna(subset=["Tipo clasificación"], inplace=True) # inplace=True modifica el DataFrame directamente
            filas_despues_eliminacion = len(df_procesado)
            st.success(f"Filas con 'Tipo clasificación' vacío eliminadas: **{filas_antes_eliminacion - filas_despues_eliminacion}**. Filas restantes: **{filas_despues_eliminacion}**.")
        else:
            st.warning("La columna **'Tipo clasificación'** no se encontró. No se eliminaron filas vacías.")

        # 2. Eliminar columnas especificadas
        columnas_existentes_para_eliminar = [col for col in nombres_columnas_a_eliminar if col in df_procesado.columns]
        columnas_no_existentes_para_eliminar = [col for col in nombres_columnas_a_eliminar if col not in df_procesado.columns]

        if columnas_existentes_para_eliminar:
            df_procesado.drop(columns=columnas_existentes_para_eliminar, inplace=True)
            st.success(f"Columnas eliminadas: **{', '.join(columnas_existentes_para_eliminar)}**.")
        else:
            st.info("Ninguna de las columnas especificadas para eliminar se encontró. No se eliminaron columnas.")

        if columnas_no_existentes_para_eliminar:
            st.warning(f"Advertencia: Las siguientes columnas especificadas para eliminación no se encontraron: **{', '.join(columnas_no_existentes_para_eliminar)}**.")

        # 3. Actualizar la columna "Total" existente
        if "Cantidad" in df_procesado.columns and "Valor unitario" in df_procesado.columns and "Total" in df_procesado.columns:
            df_procesado["Cantidad"] = pd.to_numeric(df_procesado["Cantidad"], errors='coerce')
            df_procesado["Valor unitario"] = pd.to_numeric(df_procesado["Valor unitario"], errors='coerce')
            df_procesado["Total"] = df_procesado["Cantidad"] * df_procesado["Valor unitario"]
            df_procesado["Total"] = df_procesado["Total"].fillna(0)
            st.success("La columna **'Total'** ha sido actualizada con el cálculo **'Cantidad * Valor unitario'**.")
        else:
            st.warning("Advertencia: No se pudieron encontrar las columnas **'Cantidad'**, **'Valor unitario'** y/o **'Total'**. La columna **'Total'** no se actualizó.")

        # 4. Rellenar celdas vacías o con 0 en "Tasa de cambio" con la TRM
        if "Tasa de cambio" in df_procesado.columns and "Fecha elaboración" in df_procesado.columns:
            st.info("Iniciando el proceso de rellenado de **'Tasa de cambio'** con TRM desde Datos Abiertos...")
            
            # Convertir "Fecha elaboración" a formato de fecha de Pandas y luego a datetime
            df_procesado['Fecha elaboración_dt'] = pd.to_datetime(df_procesado['Fecha elaboración'], format='%d/%m/%Y', errors='coerce')

            trm_placeholder = st.empty() # Placeholder para mostrar el progreso de la TRM
            total_trm_consultas = 0
            
            # Iterar sobre las filas del DataFrame para buscar y aplicar la TRM
            for index, row in df_procesado.iterrows():
                tasa_actual = row["Tasa de cambio"]
                fecha_elaboracion_dt = row["Fecha elaboración_dt"]

                if (pd.isna(tasa_actual) or tasa_actual == 0) and pd.notna(fecha_elaboracion_dt):
                    fecha_str_api = fecha_elaboracion_dt.strftime('%Y-%m-%d')
                    trm_placeholder.text(f"Buscando TRM para la fecha: {fecha_str_api} (Fila {index})...")
                    trm_valor = get_trm_from_datos_abiertos(fecha_str_api)
                    total_trm_consultas +=1 # Conteo de las consultas realizadas

                    if trm_valor is not None:
                        df_procesado.at[index, "Tasa de cambio"] = trm_valor
                        trm_placeholder.text(f"TRM encontrada: {trm_valor} para {fecha_str_api}. (Fila {index})")
                    else:
                        trm_placeholder.warning(f"No se pudo obtener TRM para {fecha_str_api}. La celda permanecerá sin cambios.")
                    
                    time.sleep(0.05) # Pequeña pausa para no sobrecargar la API

            # Limpiar el placeholder de TRM
            trm_placeholder.empty()
            df_procesado.drop(columns=['Fecha elaboración_dt'], inplace=True)
            st.success(f"Proceso de rellenado de **'Tasa de cambio'** completado. Total de consultas TRM: **{total_trm_consultas}**.")
        else:
            st.warning("Advertencia: No se encontraron las columnas **'Tasa de cambio'** y/o **'Fecha elaboración'**. No se buscó la TRM.")

        st.success("¡Procesamiento completado con éxito!")
        return df_procesado

    except Exception as e:
        st.error(f"Se produjo un error durante el procesamiento: {e}")
        return None

st.set_page_config(page_title="Procesador de Excel Automático", layout="centered")

st.title("📊 Procesador de Archivos Excel")

st.markdown("---")

uploaded_file = st.file_uploader(
    "Sube tu archivo Excel (.xlsx)",
    type=["xlsx"],
    help="Arrastra y suelta tu archivo Excel aquí o haz clic para buscar."
)

df_result = None # Inicializamos df_result fuera del bloque if para que sea accesible para la descarga

if uploaded_file is not None:
    st.success(f"Archivo **'{uploaded_file.name}'** cargado correctamente.")
    
    # Botón para iniciar el procesamiento
    if st.button("Iniciar Procesamiento"):
        with st.spinner("Procesando tu archivo... Esto puede tardar unos minutos, especialmente al consultar la TRM..."):
            df_result = procesar_excel_para_streamlit(uploaded_file)
        
        if df_result is not None:
            st.subheader("Vista previa del archivo procesado:")
            st.dataframe(df_result.head()) # Muestra las primeras filas del DataFrame procesado

            # Convertir DataFrame a bytes para la descarga
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_result.to_excel(writer, index=False, sheet_name='Procesado')
            processed_data = output.getvalue()

            # Botón de descarga
            st.download_button(
                label="Descargar Archivo Procesado",
                data=processed_data,
                file_name=f"procesado_{uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.info("Tu archivo ha sido procesado y está listo para descargar.")
else:
    st.info("Por favor, sube el archivo Excel para comenzar.")

