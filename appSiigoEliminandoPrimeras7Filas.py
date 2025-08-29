import streamlit as st
import pandas as pd
import time
import io
import numpy as np

# --- INICIO DE LA HERRAMIENTA DE DIAGNÓSTICO AVANZADA ---
def diagnosticar_problemas_de_conversion(uploaded_file):
    """
    Lee un archivo de Excel y muestra los valores exactos que fallan al
    intentar convertirlos a números en las columnas clave.
    """
    st.header("🕵️‍♂️ Herramienta de Diagnóstico de Tipos de Datos")
    st.info("Esta herramienta te ayudará a encontrar los valores exactos que están causando problemas de conversión en tus columnas numéricas.")

    try:
        df = pd.read_excel(uploaded_file, skiprows=7)
        st.success("Archivo leído correctamente. Analizando columnas...")

        columnas_a_revisar = ['Cantidad', 'Valor unitario', 'Tasa de cambio']
        hay_problemas = False

        for columna in columnas_a_revisar:
            st.subheader(f"Análisis de la columna: `{columna}`")

            if columna not in df.columns:
                st.warning(f"La columna '{columna}' no fue encontrada en el archivo.")
                continue

            # Forzar la columna a string para un análisis consistente
            # y eliminar filas vacías que no aportan información
            col_texto = df[columna].dropna().astype(str)

            # Intentar la conversión numérica directa
            col_numerica = pd.to_numeric(col_texto, errors='coerce')

            # Encontrar los valores que fallaron la conversión (se volvieron NaT/NaN)
            fallos_mask = col_numerica.isna()
            valores_problematicos = col_texto[fallos_mask].unique()

            if len(valores_problematicos) > 0:
                hay_problemas = True
                st.error(f"Se encontraron {len(valores_problematicos)} valores únicos en `{columna}` que NO se pueden convertir a número:")
                
                df_diag = pd.DataFrame({
                    'Valor Original Problemático': valores_problematicos
                })
                
                # Aplicar la lógica de limpieza propuesta para ver qué hace
                texto_limpio = pd.Series(valores_problematicos).astype(str).str.strip()
                texto_limpio = texto_limpio.str.replace(',', '.', regex=False)
                texto_limpio = texto_limpio.str.replace(r'\.(?=[^.]*\.)', '', regex=True)
                
                df_diag['Resultado Tras Limpieza'] = texto_limpio
                df_diag['¿Se Convierte a Número?'] = pd.to_numeric(texto_limpio, errors='coerce').notna()

                st.dataframe(df_diag)
                st.warning(f"Observa la tabla de `{columna}`. Si la columna '¿Se Convierte a Número?' muestra 'False', entonces los valores originales tienen un formato que la limpieza actual no resuelve.")

            else:
                st.success(f"¡Buenas noticias! Todos los valores en la columna `{columna}` se convierten a número correctamente.")
        
        if not hay_problemas:
            st.balloons()
            st.success("¡Diagnóstico completado! Parece que todas las columnas clave se pueden convertir a números sin problemas.")


    except Exception as e:
        st.error(f"Ocurrió un error durante el diagnóstico: {e}")
# --- FIN DE LA HERRAMIENTA DE DIAGNÓSTICO ---


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

        df_procesado = df.copy()

        def limpiar_y_convertir_a_numero(columna):
            """
            Toma una columna de pandas, la limpia de formatos mixtos (comas/puntos)
            y la convierte a un tipo de dato numérico.
            """
            # Solo procesa si la columna contiene texto
            if pd.api.types.is_string_dtype(columna) or columna.dtype == 'object':
                columna_texto = columna.astype(str).str.strip()
                
                # Reemplaza la coma decimal por un punto.
                # "4.500,25" -> "4.500.25"
                # "4,042.50" -> "4.042.50" (no cambia esta)
                columna_texto = columna_texto.str.replace(',', '.', regex=False)
                
                # Ahora que solo hay puntos, eliminamos todos los que actúan como
                # separadores de miles (es decir, todos menos el último).
                # Usamos una expresión regular para esto.
                # "4.500.25" -> "4500.25"
                # "4.042.50" -> "4042.50"
                columna_texto = columna_texto.str.replace(r'\.(?=[^.]*\.)', '', regex=True)

                return pd.to_numeric(columna_texto, errors='coerce')
            
            # Si ya es numérica, solo la devuelve
            return pd.to_numeric(columna, errors='coerce')


        # --- APLICAR LIMPIEZA ANTES DE CUALQUIER CÁLCULO ---
        st.info("Estandarizando formatos numéricos...")
        columnas_a_limpiar = ['Cantidad', 'Valor unitario', 'Tasa de cambio']
        for col_nombre in columnas_a_limpiar:
            if col_nombre in df_procesado.columns:
                df_procesado[col_nombre] = limpiar_y_convertir_a_numero(df_procesado[col_nombre])

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
            "Fecha vencimiento",
            "Nombre contacto"
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

        # 4. Crear y posicionar la nueva columna "Numero comprobante"
        columnas_necesarias = ['Número comprobante', 'Consecutivo', 'Factura proveedor']
        if all(col in df_procesado.columns for col in columnas_necesarias):
            # Definir las condiciones
            conditions = [
                df_procesado['Número comprobante'] == 'FV-1',
                df_procesado['Número comprobante'] == 'FV-2'
            ]
            
            # Definir los valores a asignar para cada condición
            choices = [
                'FLE-' + df_procesado['Consecutivo'].astype('Int64').astype(str),
                'FSE-' + df_procesado['Consecutivo'].astype('Int64').astype(str)
            ]
            
            # Usar np.select para crear los valores de la nueva columna
            # El valor por defecto será un texto vacío ''
            valores_nueva_columna = np.select(conditions, choices, default='')
            
            # Encontrar la posición de la columna "Factura proveedor" para insertar antes
            posicion_insercion = df_procesado.columns.get_loc('Factura proveedor')
            
            # Insertar la nueva columna en la posición encontrada
            df_procesado.insert(posicion_insercion, 'Numero comprobante', valores_nueva_columna)
            
            st.success("Se ha creado y llenado la nueva columna **'Numero comprobante'**.")
            
        else:
            st.warning("Advertencia: No se encontraron las columnas necesarias ('Número comprobante', 'Consecutivo', 'Factura proveedor') para crear la nueva columna.")
        
        # 5. Extraer TRM de 'Observaciones' y sobrescribir 'Tasa de cambio'
        if "Tasa de cambio" in df_procesado.columns and "Observaciones" in df_procesado.columns:
            st.info("Actualizando 'Tasa de cambio' con los valores encontrados en 'Observaciones'...")

            df_procesado['Observaciones'] = df_procesado['Observaciones'].astype(str)
            # Extrae el contenido de las llaves '{}'. El resultado será el texto o NaN si no hay llaves.
            trm_extraida = df_procesado['Observaciones'].str.extract(r'\{(.*?)\}')[0]
            # Elimina las filas donde no se encontró nada (NaN), para quedarnos solo con los valores a actualizar.
            trm_extraida.dropna(inplace=True)
            # Aseguramos que la columna 'Tasa de cambio' pueda recibir texto sin problemas.
            df_procesado['Tasa de cambio'] = df_procesado['Tasa de cambio'].astype(object)
            # Actualiza la columna 'Tasa de cambio' SÓLO con los valores encontrados.
            # El método .update() alinea por índice y solo modifica donde hay coincidencia.
            df_procesado['Tasa de cambio'].update(trm_extraida)
            
            filas_actualizadas = len(trm_extraida)
            st.success(f"Se actualizaron **{filas_actualizadas}** filas en 'Tasa de cambio'. Los valores existentes se respetaron donde no se encontró un valor entre {{}}.")
        else:
            st.warning("Advertencia: No se encontraron las columnas **'Tasa de cambio'** y/o **'Observaciones'**.")

        # 5.1. Calcular la nueva columna 'Valor Total ME'
        st.info("Calculando la nueva columna 'Valor Total ME'...")
        if 'Total' in df_procesado.columns and 'Tasa de cambio' in df_procesado.columns:
            # Para evitar errores, convertimos 'Tasa de cambio' a número. 
            # Los valores no numéricos se volverán NaN (Not a Number).
            tasa_numerica = pd.to_numeric(df_procesado['Tasa de cambio'], errors='coerce')
            
            # Reemplazamos 0 con NaN para evitar errores de división por cero.
            tasa_numerica.replace(0, np.nan, inplace=True)

            # Realizamos la división. Si se divide por NaN, el resultado será NaN.
            df_procesado['Valor Total ME'] = df_procesado['Total'] / tasa_numerica
            
            # Rellenamos cualquier resultado inválido (NaN) con 0 para mantener la consistencia.
            df_procesado['Valor Total ME'].fillna(0, inplace=True)
            
            st.success("Se ha creado y calculado la columna **'Valor Total ME'**.")
        else:
            st.warning("No se pudieron encontrar las columnas 'Total' y/o 'Tasa de cambio'. No se pudo calcular 'Valor Total ME'.")

        # 6. Relacionar documentos FV-1 con DS-1 y FC-1
        st.info("Iniciando el proceso de relacionamiento de documentos...")
        
        # Separar el DataFrame en los dos grupos principales
        df_destino = df_procesado[df_procesado['Número comprobante'].isin(['FV-1', 'FV-2'])].copy()
        df_fuente = df_procesado[df_procesado['Número comprobante'].isin(['DS-1', 'FC-1'])].copy()

        if not df_fuente.empty:
            # Preparar el DataFrame fuente (DS-1, FC-1)
            df_fuente['NIT_relacion'] = df_fuente['Observaciones'].str.extract(r'\((.*?)\)')[0]
            
            df_destino['Identificación'] = df_destino['Identificación'].astype('Int64').astype(str)
            df_destino['Código'] = df_destino['Código'].astype(str)
            
            df_fuente['NIT_relacion'] = df_fuente['NIT_relacion'].astype(str)
            df_fuente['Código'] = df_fuente['Código'].astype(str)
            
            # Añadir prefijo a las columnas para evitar colisiones y dar claridad
            df_fuente = df_fuente.add_prefix('REL_')
            
            # Realizar la unión externa (outer join)
            df_final = pd.merge(
                df_destino,
                df_fuente,
                how='outer',
                left_on=['Identificación', 'Código'],
                right_on=['REL_NIT_relacion', 'REL_Código']
            )
            
            st.success("Relacionamiento completado. Los documentos sin pareja se han conservado.")
            df_procesado = df_final
        else:
            st.warning("No se encontraron documentos DS-1 o FC-1 para relacionar. El archivo final no tendrá columnas de relación.")
        
        # 7. Organizar y Limpiar Columnas Finales
        st.info("Organizando el formato final del archivo...")
        
        # A. Renombrar la columna "Tipo clasificación" a "Tipo Bien"
        # Verificamos si la columna existe antes de intentar renombrarla
        if "Tipo clasificación" in df_procesado.columns:
            df_procesado.rename(columns={"Tipo clasificación": "Tipo Bien"}, inplace=True)
            st.info("La columna **'Tipo clasificación'** ha sido renombrada a **'Tipo Bien'**.")
        
        if 'Tipo Bien' in df_procesado.columns:
            # Creamos un diccionario con los valores a reemplazar
            mapeo_valores = {
                'Servicio': 'S',
                'Producto': 'P'
            }
            df_procesado['Tipo Bien'].replace(mapeo_valores, inplace=True)
            st.info("Valores en 'Tipo Bien' actualizados: 'Servicio' a 'S' y 'Producto' a 'P'.")
        
        #Creación de la nueva columna "Vendedor"
        if 'Vendedor' not in df_procesado.columns:
            df_procesado['Vendedor'] = ''
            
        #Creación de la nueva columna "Clasificación Producto"
        if 'Clasificación Producto' not in df_procesado.columns:
            df_procesado['Clasificación Producto'] = ''
            
        #Creación de la nueva columna "Línea"
        if 'Línea' not in df_procesado.columns:
            df_procesado['Línea'] = ''
            
        #Creación de la nueva columna "Descripción Línea"
        if 'Descripción Línea' not in df_procesado.columns:
            df_procesado['Descripción Línea'] = ''
            
        #Creación de la nueva columna "Sublínea"
        if 'Sublínea' not in df_procesado.columns:
            df_procesado['Sublínea'] = ''
            
        #Creación de la nueva columna "Descripción Sublínea"
        if 'Descripción Sublínea' not in df_procesado.columns:
            df_procesado['Descripción Sublínea'] = ''
            
        
        #Se define el orden y la selección final de las columnas
        columnas_finales = [
            # Columnas del lado izquierdo (FV)
            'Tipo Bien', 'Clasificación Producto', 'Línea', 'Descripción Línea', 'Sublínea', 'Descripción Sublínea', 'Código', 'Nombre', 'Número comprobante', 'Numero comprobante',
            'Fecha elaboración', 'Identificación', 'Nombre tercero', 'Vendedor', 'Cantidad',
            'Valor unitario', 'Total', 'Tasa de cambio', 'Valor Total ME', 'Observaciones',
            
            # Columnas del lado derecho (REL_)
            'REL_Número comprobante', 'REL_Consecutivo',
            'REL_Factura proveedor', 'REL_Identificación', 'REL_Nombre tercero', 'REL_Cantidad',
            'REL_Valor unitario',  'REL_Tasa de cambio', 'REL_Total', 'REL_Valor Total ME'
        ]
        
        # Filtrar la lista para incluir solo las columnas que realmente existen en el DataFrame
        # Esto hace el código más robusto si alguna columna faltara
        columnas_existentes_ordenadas = [col for col in columnas_finales if col in df_procesado.columns]

        # Reordenar y eliminar las columnas no deseadas de una sola vez
        df_procesado = df_procesado[columnas_existentes_ordenadas]

        st.success("Columnas reorganizadas y limpiadas con éxito.")

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
    
    # --- CÓDIGO MODIFICADO PARA DIAGNÓSTICO ---
    # Llama a la herramienta de diagnóstico directamente al subir el archivo
    # No necesitas presionar un botón.
    diagnosticar_problemas_de_conversion(uploaded_file)
    
    #if st.button("Iniciar Procesamiento"):
        #with st.spinner("Procesando tu archivo... Esto puede tardar unos minutos, especialmente al consultar la TRM..."):
            #df_result = procesar_excel_para_streamlit(uploaded_file)
        
        #if df_result is not None:
            #st.subheader("Vista previa del archivo procesado:")
            #st.dataframe(df_result.head())

            #output = io.BytesIO()
            #with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                #df_result.to_excel(writer, index=False, sheet_name='Procesado')
            #processed_data = output.getvalue()

            #st.download_button(
                #label="Descargar Archivo Procesado",
                #data=processed_data,
                #file_name=f"procesado_{uploaded_file.name}",
                #mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            #)
            #st.info("Tu archivo ha sido procesado y está listo para descargar.")
else:
    st.info("Por favor, sube un archivo Excel para comenzar.")


