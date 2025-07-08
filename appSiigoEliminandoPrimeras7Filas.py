import streamlit as st
import pandas as pd
import requests
from datetime import datetime
import time
import io

# --- FUNCIÓN MEJORADA PARA OBTENER LA TRM ---
@st.cache_data(ttl=3600) # Almacena en caché los resultados de la TRM por 1 hora
def get_trm_from_datos_abiertos(target_date_str):
    """
    Consulta la TRM desde la API de Datos Abiertos Colombia (Socrata),
    manejando fines de semana y festivos, buscando la TRM vigente en o antes de la fecha.

    Args:
        target_date_str (str): La fecha para la cual se desea consultar la TRM, en formato 'YYYY-MM-DD'.

    Returns:
        float or None: El valor de la TRM, o None si no se encuentra o hay un error.
    """
    BASE_URL = "https://www.datos.gov.co/resource/mcec-87by.json"

    try:
        # Construir la consulta Socrata Query Language (SoQL)
        # $where: vigenciadesde <= 'fecha_solicitada' (busca TRM vigente en o antes de la fecha)
        # $order: vigenciadesde DESC (para obtener la más reciente primero)
        # $limit: 1 (para obtener solo un resultado)
        soql_query = f"$where=vigenciadesde <= '{target_date_str}T23:59:59.000'&$order=vigenciadesde DESC&$limit=1"

        response = requests.get(f"{BASE_URL}?{soql_query}", timeout=10) # Añadir timeout
        response.raise_for_status() # Lanza una excepción para códigos de estado HTTP 4xx/5xx

        data = response.json()

        if data:
            # Si hay datos, el primer elemento es la TRM más reciente y vigente para esa fecha o anterior
            trm_data = data[0]
            return float(trm_data.get('valor')) # Devuelve solo el valor float
        else:
            # st.warning(f"No se encontró TRM vigente para la fecha {target_date_str} o anterior en Datos Abiertos.")
            return None # Si no hay datos, devuelve None

    except requests.exceptions.RequestException as e:
        # st.error(f"Error de conexión o HTTP al consultar Datos Abiertos para {target_date_str}: {e}")
        return None
    except (ValueError, IndexError, TypeError) as e:
        # st.error(f"Error al parsear o acceder a los datos de la TRM para {target_date_str}: {e}")
        return None
    except Exception as e:
        # st.error(f"Error inesperado al consultar Datos Abiertos para {target_date_str}: {e}")
        return None

# --- Función Principal de Procesamiento ---
def procesar_excel_para_streamlit(uploaded_file):
    """
    Procesa el archivo de Excel subido:
    - Ignora las primeras 7 filas al cargar el archivo (asumiendo que los encabezados están en la fila 8).
    - Elimina filas con 'Tipo clasificación' vacío.
    - Elimina columnas no deseadas.
    - Actualiza la columna 'Total'.
    - Rellena 'Tasa de cambio' con TRM de API bajo condiciones específicas.

    Args:
        uploaded_file (streamlit.UploadedFile): El archivo Excel subido por el usuario.

    Returns:
        pandas.DataFrame or None: El DataFrame procesado o None si hay un error.
    """
    try:
        # Usar skiprows para que Pandas lea el encabezado correcto
        df = pd.read_excel(uploaded_file, skiprows=7) # La fila 8 (índice 7) se toma como encabezado

        # Verifica si el DataFrame tiene columnas después de skiprows.
        if df.empty or df.columns.empty:
            st.error("Parece que el archivo no tiene datos o encabezados después de saltar las primeras 7 filas. Por favor, verifica el formato del archivo.")
            return None

        st.info(f"Archivo cargado exitosamente. Se saltaron las primeras 7 filas. Filas iniciales (después de saltar): **{len(df)}**.")

        # Columnas a eliminar predefinidas
        nombres_columnas_a_eliminar = [
            "Sucursal",
            "Centro costo",
            "Fecha creación",
            "Fecha modificación",
            "Correo electrónico",
            "Tipo de registro",
            "Referencia fábrica",
            "Bodega",
            "Identificación Vendedor",
            "Nombre vendedor",
            "Valor desc.",
            "Base AIU",
            "Impuesto cargo",
            "Valor Impuesto Cargo",
            "Impuesto Cargo 2",
            "Valor Impuesto Cargo 2",
            "Impuesto retención",
            "Valor Impuesto Retención",
            "Base retención (ICA/IVA)",
            "Cargo en totales",
            "Descuento en totales",
            "Moneda",
            "Forma pago",
            "Fecha vencimiento"
        ]

        df_procesado = df.copy()

        # 1. Eliminar filas donde "Tipo clasificación" esté vacío/NaN
        if "Tipo clasificación" in df_procesado.columns:
            filas_antes_eliminacion = len(df_procesado)
            df_procesado.dropna(subset=["Tipo clasificación"], inplace=True)
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
        if "Tasa de cambio" in df_procesado.columns and "Fecha elaboración" in df_procesado.columns and "Número comprobante" in df_procesado.columns:
            st.info("Iniciando el proceso de rellenado de **'Tasa de cambio'** con TRM desde Datos Abiertos...")
            
            df_procesado['Fecha elaboración_dt'] = pd.to_datetime(df_procesado['Fecha elaboración'], format='%d/%m/%Y', errors='coerce')

            trm_progress_bar = st.progress(0)
            trm_placeholder = st.empty()
            
            # Identificar las filas que necesitan consulta de TRM
            filas_a_consultar = df_procesado[
                (pd.isna(df_procesado["Tasa de cambio"]) | (df_procesado["Tasa de cambio"] == 0)) &
                pd.notna(df_procesado["Fecha elaboración_dt"]) &
                (df_procesado["Número comprobante"].astype(str).str.startswith("FV", na=False))
            ]
            total_trm_consultas_necesarias = len(filas_a_consultar)
            consultas_realizadas = 0

            for index, row in df_procesado.iterrows():
                tasa_actual = row["Tasa de cambio"]
                fecha_elaboracion_dt = row["Fecha elaboración_dt"]
                numero_comprobante = row["Número comprobante"]

                # --- INICIO DE LA MODIFICACIÓN ---
                # Se añade la tercera condición: str(numero_comprobante).startswith("FV")
                # Esto convierte el valor a texto de forma segura y comprueba si empieza con "FV"
                condicion_tasa = pd.isna(tasa_actual) or tasa_actual == 0
                condicion_fecha = pd.notna(fecha_elaboracion_dt)
                condicion_comprobante = str(numero_comprobante).startswith("FV")

                if condicion_tasa and condicion_fecha and condicion_comprobante:
                # --- FIN DE LA MODIFICACIÓN ---
                    fecha_str_api = fecha_elaboracion_dt.strftime('%Y-%m-%d')
                    trm_placeholder.text(f"Buscando TRM para la fecha: {fecha_str_api} (Fila {index+2} de Excel original)...")
                    trm_valor = get_trm_from_datos_abiertos(fecha_str_api)
                    consultas_realizadas += 1
                    
                    if total_trm_consultas_necesarias > 0:
                        progress_percentage = int((consultas_realizadas / total_trm_consultas_necesarias) * 100)
                        trm_progress_bar.progress(progress_percentage)

                    if trm_valor is not None:
                        df_procesado.at[index, "Tasa de cambio"] = trm_valor
                        trm_placeholder.text(f"TRM encontrada: {trm_valor} para {fecha_str_api}. (Fila {index+2} de Excel original)")
                    else:
                        trm_placeholder.warning(f"No se pudo obtener TRM para {fecha_str_api}. La celda permanecerá sin cambios. (Fila {index+2} de Excel original)")
                    
                    time.sleep(0.05)

            trm_placeholder.empty()
            trm_progress_bar.empty()
            df_procesado.drop(columns=['Fecha elaboración_dt'], inplace=True)
            st.success(f"Proceso de rellenado de **'Tasa de cambio'** completado. Total de consultas TRM realizadas: **{consultas_realizadas}**.")
        else:
            st.warning("Advertencia: No se encontraron las columnas **'Tasa de cambio'**, **'Fecha elaboración'** y/o **'Número comprobante'**. No se buscó la TRM.")

        st.success("¡Procesamiento completado con éxito!")
        return df_procesado

    except Exception as e:
        st.error(f"Se produjo un error durante el procesamiento: {e}")
        return None

# --- Interfaz de Usuario de Streamlit ---
st.set_page_config(page_title="Procesador de Excel Automático", layout="centered")

st.title("📊 Procesador de Archivos Excel")
st.markdown("---")

uploaded_file = st.file_uploader(
    "Sube tu archivo Excel (.xlsx)",
    type=["xlsx"],
    help="Arrastra y suelta tu archivo Excel aquí o haz clic para buscar."
)

df_result = None

if uploaded_file is not None:
    st.success(f"Archivo **'{uploaded_file.name}'** cargado correctamente.")
    
    if st.button("Iniciar Procesamiento"):
        with st.spinner("Procesando tu archivo... Esto puede tardar unos minutos, especialmente al consultar la TRM..."):
            df_result = procesar_excel_para_streamlit(uploaded_file)
        
        if df_result is not None:
            st.subheader("Vista previa del archivo procesado:")
            st.dataframe(df_result.head())

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_result.to_excel(writer, index=False, sheet_name='Procesado')
            processed_data = output.getvalue()

            st.download_button(
                label="Descargar Archivo Procesado",
                data=processed_data,
                file_name=f"procesado_{uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.info("Tu archivo ha sido procesado y está listo para descargar.")
else:
    st.info("Por favor, sube un archivo Excel para comenzar.")


