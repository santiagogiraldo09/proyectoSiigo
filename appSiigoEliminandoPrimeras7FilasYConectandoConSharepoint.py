import streamlit as st
import pandas as pd
import io
import numpy as np
import locale
from datetime import datetime
import requests
from msal import ConfidentialClientApplication
import urllib.parse


# ==============================================================================
# SECCIÓN 1: CONFIGURACIÓN DE SHAREPOINT Y AZURE
# ==============================================================================
CLIENT_ID = "b469ba00-b7b6-434c-91bf-d3481c171da5"
CLIENT_SECRET = "8nS8Q~tAYqkeISRUQyOBBAsLn6b_Z8LdNQR23dnn"
TENANT_ID = "f20cbde7-1c45-44a0-89c5-63a25c557ef8"
SHAREPOINT_HOSTNAME = "iacsas.sharepoint.com"
SITE_NAME = "PruebasProyectosSantiago"
# ==============================================================================
# SECCIÓN 2: FUNCIONES DE AUTENTICACIÓN Y CONEXIÓN
# ==============================================================================
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

def get_access_token():
    """Se autentica para obtener un token de acceso."""
    app = ConfidentialClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" in result:
        # MENSAJE DE ÉXITO AÑADIDO
        st.success("✅ Token de acceso obtenido con éxito.")
        return result['access_token']
    else:
        st.error(f"Error al obtener token: {result.get('error_description')}")
        return None

def get_sharepoint_site_id(access_token):
    """Obtiene el ID del sitio de SharePoint y confirma el éxito."""
    if not access_token:
        return None
        
    headers = {'Authorization': f'Bearer {access_token}'}
    site_url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_HOSTNAME}:/sites/{SITE_NAME}"
    try:
        response = requests.get(site_url, headers=headers)
        response.raise_for_status()
        site_data = response.json()
        site_id = site_data.get('id')
        if site_id:
            # MENSAJE DE ÉXITO AÑADIDO
            st.success(f"✅ Conexión exitosa con el sitio SharePoint: '{SITE_NAME}'")
            return site_id
        else:
            # ERROR MÁS CLARO
            st.error("Respuesta de la API exitosa, pero no se encontró un 'id' para el sitio. Verifica que el 'SITE_NAME' sea correcto.")
            return None
    except requests.exceptions.RequestException as e:
        st.error(f"Error al obtener site_id. Verifica que 'SHAREPOINT_HOSTNAME' y 'SITE_NAME' son correctos.")
        # Muestra el error devuelto por el servidor de Microsoft para dar más pistas
        st.json(e.response.json())
        return None

import urllib.parse
import pandas as pd

def explorar_raiz_sharepoint(site_id, headers):
    """
    Explora la raíz del drive de SharePoint para ver qué carpetas existen realmente
    """
    st.info("🗂️ Explorando la raíz del sitio SharePoint...")
    
    endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
    
    try:
        response = requests.get(endpoint, headers=headers)
        if response.status_code == 200:
            items = response.json().get('value', [])
            st.success(f"✅ Encontrados {len(items)} elementos en la raíz:")
            
            # Crear tabla para mejor visualización
            data = []
            for item in items:
                tipo = "📁 Carpeta" if item.get('folder') else "📄 Archivo"
                nombre = item.get('name', 'Sin nombre')
                tamano = item.get('size', 0)
                fecha = item.get('lastModifiedDateTime', 'N/A')[:10] if item.get('lastModifiedDateTime') else 'N/A'
                
                data.append({
                    'Tipo': tipo,
                    'Nombre': nombre,
                    'Tamaño (bytes)': tamano,
                    'Última modificación': fecha
                })
            
            df = pd.DataFrame(data)
            st.dataframe(df, use_container_width=True)
            return items
        else:
            st.error(f"❌ No se pudo explorar la raíz. HTTP {response.status_code}")
            st.json(response.json())
            return []
    except Exception as e:
        st.error(f"Error al explorar raíz: {e}")
        return []

def explorar_carpeta_especifica(site_id, headers, ruta_carpeta):
    """
    Explora una carpeta específica y muestra su contenido
    """
    st.info(f"📂 Explorando carpeta: '{ruta_carpeta}'")
    ruta_encoded = urllib.parse.quote(ruta_carpeta, safe='/')
    endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{ruta_encoded}:/children"
    
    try:
        response = requests.get(endpoint, headers=headers)
        if response.status_code == 200:
            items = response.json().get('value', [])
            st.success(f"✅ Encontrados {len(items)} elementos en '{ruta_carpeta}':")
            
            # Crear tabla
            data = []
            for item in items:
                tipo = "📁 Carpeta" if item.get('folder') else "📄 Archivo"
                nombre = item.get('name', 'Sin nombre')
                tamano = item.get('size', 0)
                fecha = item.get('lastModifiedDateTime', 'N/A')[:10] if item.get('lastModifiedDateTime') else 'N/A'
                
                data.append({
                    'Tipo': tipo,
                    'Nombre': nombre,
                    'Tamaño (bytes)': tamano,
                    'Última modificación': fecha
                })
            
            df = pd.DataFrame(data)
            st.dataframe(df, use_container_width=True)
            return items
        else:
            st.error(f"❌ No se pudo explorar '{ruta_carpeta}'. HTTP {response.status_code}")
            error_response = response.json()
            st.json(error_response)
            return []
    except Exception as e:
        st.error(f"Error al explorar carpeta '{ruta_carpeta}': {e}")
        return []

def buscar_archivo_globalmente(site_id, headers, nombre_archivo):
    """
    Busca un archivo específico en todo el sitio SharePoint
    """
    st.info(f"🔍 Buscando '{nombre_archivo}' en todo el sitio...")
    
    # URL encode del término de búsqueda
    query_encoded = urllib.parse.quote(nombre_archivo)
    endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/search(q='{query_encoded}')"
    
    try:
        response = requests.get(endpoint, headers=headers)
        if response.status_code == 200:
            items = response.json().get('value', [])
            
            if items:
                st.success(f"✅ Se encontraron {len(items)} resultados para '{nombre_archivo}':")
                
                data = []
                for item in items:
                    tipo = "📁 Carpeta" if item.get('folder') else "📄 Archivo"
                    nombre = item.get('name', 'Sin nombre')
                    
                    # Construir la ruta completa
                    parent_path = item.get('parentReference', {}).get('path', '')
                    if parent_path:
                        # Limpiar la ruta (quitar /drive/root: del inicio)
                        ruta_limpia = parent_path.replace('/drive/root:', '').replace('/drive/root', '')
                        ruta_completa = f"{ruta_limpia}/{nombre}" if ruta_limpia else nombre
                    else:
                        ruta_completa = nombre
                    
                    web_url = item.get('webUrl', 'N/A')
                    
                    data.append({
                        'Tipo': tipo,
                        'Nombre': nombre,
                        'Ruta Completa': ruta_completa,
                        'URL Web': web_url
                    })
                
                df = pd.DataFrame(data)
                st.dataframe(df, use_container_width=True)
                return items
            else:
                st.warning(f"⚠️ No se encontraron resultados para '{nombre_archivo}'")
                return []
        else:
            st.error(f"❌ Error en búsqueda global. HTTP {response.status_code}")
            st.json(response.json())
            return []
    except Exception as e:
        st.error(f"Error en búsqueda global: {e}")
        return []

def probar_nombres_comunes_documentos(site_id, headers):
    """
    Prueba nombres comunes para la carpeta de documentos en SharePoint
    """
    st.info("🧪 Probando nombres comunes para la carpeta de documentos...")
    
    nombres_comunes = [
        "Shared Documents",
        "Documents", 
        "Documentos Compartidos",
        "Documentos compartidos",
        "documentos compartidos",
        "General",
        "Sitio"
    ]
    
    carpetas_encontradas = []
    
    for nombre in nombres_comunes:
        st.write(f"Probando: '{nombre}'...")
        ruta_encoded = urllib.parse.quote(nombre, safe='/')
        endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{ruta_encoded}"
        
        try:
            response = requests.get(endpoint, headers=headers)
            if response.status_code == 200:
                st.success(f"✅ ¡Encontrada!: '{nombre}'")
                carpetas_encontradas.append(nombre)
            else:
                st.write(f"❌ No existe: '{nombre}'")
        except Exception as e:
            st.write(f"❌ Error: '{nombre}' - {e}")
    
    return carpetas_encontradas

def generar_rutas_sugeridas(carpetas_documentos, archivo_objetivo="TRM.xlsx"):
    """
    Genera rutas sugeridas basadas en las carpetas encontradas
    """
    st.info("💡 Rutas sugeridas basadas en carpetas encontradas:")
    
    rutas_sugeridas = []
    
    for carpeta_base in carpetas_documentos:
        # Rutas posibles con diferentes variaciones
        rutas_posibles = [
            f"{carpeta_base}/01 Archivos Area Administrativa/{archivo_objetivo}",
            f"{carpeta_base}/Archivos Area Administrativa/{archivo_objetivo}",
            f"{carpeta_base}/Area Administrativa/{archivo_objetivo}",
            f"{carpeta_base}/Administrativa/{archivo_objetivo}",
            f"{carpeta_base}/{archivo_objetivo}"
        ]
        
        rutas_sugeridas.extend(rutas_posibles)
    
    for i, ruta in enumerate(rutas_sugeridas, 1):
        st.code(f"Opción {i}: {ruta}")
    
    return rutas_sugeridas

# Función principal de exploración
def explorador_completo_sharepoint(site_id, headers):
    """
    Herramienta completa de exploración de SharePoint
    """
    st.header("🔍 Explorador Completo de SharePoint")
    
    tab1, tab2, tab3, tab4 = st.tabs(["📁 Raíz", "🔍 Búsqueda", "📂 Carpetas", "💡 Sugerencias"])
    
    with tab1:
        st.subheader("Contenido de la Raíz")
        if st.button("Explorar Raíz del Sitio"):
            items_raiz = explorar_raiz_sharepoint(site_id, headers)
    
    with tab2:
        st.subheader("Búsqueda Global")
        archivo_buscar = st.text_input("Nombre del archivo a buscar:", "TRM.xlsx")
        if st.button("Buscar Archivo"):
            if archivo_buscar:
                resultados = buscar_archivo_globalmente(site_id, headers, archivo_buscar)
    
    with tab3:
        st.subheader("Explorar Carpeta Específica")
        ruta_manual = st.text_input("Ruta de carpeta a explorar:", "")
        if st.button("Explorar Carpeta") and ruta_manual:
            explorar_carpeta_especifica(site_id, headers, ruta_manual)
        
        st.markdown("---")
        st.subheader("Probar Nombres Comunes")
        if st.button("Buscar Carpetas de Documentos Comunes"):
            carpetas_encontradas = probar_nombres_comunes_documentos(site_id, headers)
            if carpetas_encontradas:
                st.success(f"Carpetas encontradas: {carpetas_encontradas}")
    
    with tab4:
        st.subheader("Generador de Rutas")
        if st.button("Generar Rutas Sugeridas"):
            # Primero buscar carpetas comunes
            carpetas_encontradas = probar_nombres_comunes_documentos(site_id, headers)
            if carpetas_encontradas:
                generar_rutas_sugeridas(carpetas_encontradas)
            else:
                st.warning("No se encontraron carpetas de documentos comunes. Explora la raíz primero.")
    
def verificar_archivo_por_ruta(site_id, headers, ruta_archivo):
    """
    Verifica si un archivo o carpeta existe en una ruta específica.
    """
    st.info(f"Verificando ruta fija: '{ruta_archivo}'...")
    
    # MEJORA 1: URL encode de la ruta para manejar espacios y caracteres especiales
    ruta_encoded = urllib.parse.quote(ruta_archivo, safe='/')
    endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{ruta_encoded}"
    
    st.info(f"🔍 URL construida: {endpoint}")
    
    try:
        response = requests.get(endpoint, headers=headers)
        
        # MEJORA 2: Mejor manejo de errores con más información de debug
        st.info(f"📊 Código de respuesta HTTP: {response.status_code}")
        
        if response.status_code == 200:
            st.success(f"✅ Ruta encontrada: '{ruta_archivo}'")
            return True
        elif response.status_code == 404:
            st.warning(f"⚠️ Archivo no encontrado: '{ruta_archivo}'")
            # Mostrar la respuesta de error para más contexto
            try:
                error_details = response.json()
                st.error(f"Detalles del error 404: {error_details}")
            except:
                st.error("No se pudo parsear la respuesta de error")
            return False
        else:
            st.error(f"❌ Error HTTP {response.status_code}: {response.text}")
            return False
            
    except requests.exceptions.RequestException as e:
        st.error(f"Error de conexión al verificar la ruta fija: {e}")
        return False

def verificar_archivo_alternativo(site_id, headers, ruta_archivo):
    """
    Función alternativa que también verifica diferentes formatos de ruta
    """
    rutas_a_probar = [
        ruta_archivo,  # Ruta original
        ruta_archivo.replace(" ", "%20"),  # Con espacios URL encoded manualmente
        ruta_archivo.replace("/", "\\"),   # Con backslashes (formato Windows)
    ]
    
    for i, ruta in enumerate(rutas_a_probar):
        st.info(f"🔄 Probando formato {i+1}: '{ruta}'")
        ruta_encoded = urllib.parse.quote(ruta, safe='/')
        endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{ruta_encoded}"
        
        try:
            response = requests.get(endpoint, headers=headers)
            st.info(f"📊 Respuesta para formato {i+1}: HTTP {response.status_code}")
            
            if response.status_code == 200:
                st.success(f"✅ ¡Archivo encontrado con formato {i+1}!: '{ruta}'")
                return True, ruta
        except Exception as e:
            st.warning(f"Error con formato {i+1}: {e}")
            continue
    
    st.error("❌ No se pudo encontrar el archivo con ninguno de los formatos probados")
    return False, None

def listar_contenido_carpeta(site_id, headers, ruta_carpeta="Documentos compartidos"):
    """
    Función auxiliar para listar el contenido de una carpeta y ayudar en debug
    """
    st.info(f"📂 Listando contenido de: '{ruta_carpeta}'")
    ruta_encoded = urllib.parse.quote(ruta_carpeta, safe='/')
    endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{ruta_encoded}:/children"
    
    try:
        response = requests.get(endpoint, headers=headers)
        if response.status_code == 200:
            items = response.json().get('value', [])
            st.success(f"✅ Encontrados {len(items)} elementos en '{ruta_carpeta}':")
            
            for item in items[:10]:  # Mostrar solo los primeros 10 items
                tipo = "📁" if item.get('folder') else "📄"
                nombre = item.get('name', 'Sin nombre')
                st.write(f"{tipo} {nombre}")
                
            if len(items) > 10:
                st.info(f"... y {len(items) - 10} elementos más.")
            return True
        else:
            st.error(f"❌ No se pudo listar la carpeta. HTTP {response.status_code}")
            return False
    except Exception as e:
        st.error(f"Error al listar carpeta: {e}")
        return False

def encontrar_archivo_del_mes_en_carpeta(site_id, headers, ruta_carpeta):
    """
    Busca dentro de una CARPETA específica un archivo del mes actual,
    sin depender del locale del servidor.
    """
    try:
        # --- SOLUCIÓN: Usar una lista propia para los meses en español ---
        meses_es = [
            "enero", "febrero", "marzo", "abril", "mayo", "junio", 
            "julio", "agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
        ]
        
        fecha_actual = datetime.now()
        mes_numero = fecha_actual.month
        # Obtenemos el nombre del mes de nuestra lista (índice es mes - 1)
        mes_nombre = meses_es[mes_numero - 1]
        
        st.info(f"Buscando archivo de '{mes_nombre.capitalize()}' en la carpeta: '{ruta_carpeta}'...")
        
        search_endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{ruta_carpeta}:/search(q='{mes_nombre}')"
        
        response = requests.get(search_endpoint, headers=headers)
        response.raise_for_status()
        search_results = response.json()
        
        for item in search_results.get('value', []):
            nombre_archivo = item.get('name', '')
            if mes_nombre.lower() in nombre_archivo.lower() and str(mes_numero) in nombre_archivo:
                st.success(f"✅ Archivo del mes encontrado: {nombre_archivo}")
                ruta_relativa = item.get('parentReference', {}).get('path', '').split('root:')[-1]
                ruta_completa = f"{ruta_relativa}/{nombre_archivo}"
                return nombre_archivo, ruta_completa
        
        st.warning(f"⚠️ No se encontró archivo para '{mes_nombre.capitalize()}' en la carpeta especificada.")
        return None, None

    except requests.exceptions.RequestException as e:
        st.error(f"Error de conexión al buscar el archivo del mes: {e.response.text}")
        return None, None
    except Exception as e:
        st.error(f"Error inesperado durante la búsqueda del mes: {e}")
        return None, None


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
        
        # --- FUNCIÓN DE LIMPIEZA SIMPLE ---
        def convertir_a_numero_limpiando_comas(columna):
            if not pd.api.types.is_string_dtype(columna):
                columna = columna.astype(str)
            columna_limpia = columna.str.replace(',', '', regex=False)
            return pd.to_numeric(columna_limpia, errors='coerce')

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
        #if "Tasa de cambio" in df_procesado.columns and "Observaciones" in df_procesado.columns:
            #st.info("Actualizando 'Tasa de cambio' con los valores encontrados en 'Observaciones'...")

            #df_procesado['Observaciones'] = df_procesado['Observaciones'].astype(str)
            # Extrae el contenido de las llaves '{}'. El resultado será el texto o NaN si no hay llaves.
            #trm_extraida = df_procesado['Observaciones'].str.extract(r'\{(.*?)\}')[0]
            # Elimina las filas donde no se encontró nada (NaN), para quedarnos solo con los valores a actualizar.
            #trm_extraida.dropna(inplace=True)
            # Aseguramos que la columna 'Tasa de cambio' pueda recibir texto sin problemas.
            #df_procesado['Tasa de cambio'] = df_procesado['Tasa de cambio'].astype(object)
            # Actualiza la columna 'Tasa de cambio' SÓLO con los valores encontrados.
            # El método .update() alinea por índice y solo modifica donde hay coincidencia.
            #df_procesado['Tasa de cambio'].update(trm_extraida)
            
            #filas_actualizadas = len(trm_extraida)
            #st.success(f"Se actualizaron **{filas_actualizadas}** filas en 'Tasa de cambio'. Los valores existentes se respetaron donde no se encontró un valor entre {{}}.")
        #else:
            #st.warning("Advertencia: No se encontraron las columnas **'Tasa de cambio'** y/o **'Observaciones'**.")
        # 5. Extraer, LIMPIAR y sobrescribir 'Tasa de cambio' desde 'Observaciones' (LÓGICA CORREGIDA Y ENFOCADA)
        if "Tasa de cambio" in df_procesado.columns and "Observaciones" in df_procesado.columns:
            
            # Para evitar problemas, nos aseguramos de que la columna 'Tasa de cambio' sea numérica desde el principio.
            # Usamos la limpieza simple de comas que ya definimos.
            df_procesado['Tasa de cambio'] = convertir_a_numero_limpiando_comas(df_procesado['Tasa de cambio']).fillna(0)

            # 1. EXTRAER el valor de las observaciones como texto.
            trm_extraida = df_procesado['Observaciones'].astype(str).str.extract(r'\{(.*?)\}')[0]
            
            # Quitamos las filas donde no se encontró nada.
            trm_extraida.dropna(inplace=True)

            if not trm_extraida.empty:
                st.info("Valores de TRM encontrados en 'Observaciones'. Limpiando y actualizando...")

                # 2. LIMPIAR el texto extraído (quitamos comas de miles).
                # Ejemplo: "4,061.36" se convierte en "4061.36"
                trm_limpia = trm_extraida.str.replace(',', '', regex=False)

                # 3. CONVERTIR el texto limpio a un formato numérico.
                trm_numerica = pd.to_numeric(trm_limpia, errors='coerce')
                
                # Quitamos las filas donde la conversión a número pudo haber fallado.
                trm_numerica.dropna(inplace=True)

                # 4. ACTUALIZAR la columna 'Tasa de cambio' con los valores ya numéricos y limpios.
                # El método .update() alinea por índice y solo modifica donde encuentra correspondencia.
                df_procesado['Tasa de cambio'].update(trm_numerica)
                st.success(f"Se actualizaron **{len(trm_numerica)}** filas en 'Tasa de cambio' con valores numéricos limpios desde 'Observaciones'.")


        # 5.1. Calcular la nueva columna 'Valor Total ME' (VERSIÓN CORREGIDA FINAL)
        st.info("Calculando 'Valor Total ME'...")
        if 'Total' in df_procesado.columns and 'Tasa de cambio' in df_procesado.columns:
            
            # PASO CLAVE: Nos aseguramos de que 'Tasa de cambio' sea numérica OTRA VEZ,
            # justo antes de la división, para revertir el cambio a 'object' del paso anterior.
            tasa_numerica = pd.to_numeric(df_procesado['Tasa de cambio'], errors='coerce')
            
            # Reemplazamos 0 con NaN para evitar errores de división por cero.
            tasa_numerica.replace(0, np.nan, inplace=True)

            # Realizamos la división.
            df_procesado['Valor Total ME'] = df_procesado['Total'] / tasa_numerica
            
            # Rellenamos cualquier resultado inválido (NaN) con 0.
            df_procesado['Valor Total ME'].fillna(0, inplace=True)
            
            st.success("Se ha creado y calculado la columna **'Valor Total ME'**.")
        else:
            st.warning("No se pudo calcular 'Valor Total ME'.")

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

# --- PASO 1: ENTRADA DE DATOS Y VERIFICACIÓN ---
st.header("1. Verificación de Archivos en SharePoint")

# Usamos st.session_state para guardar el estado de la conexión
if 'conectado' not in st.session_state:
    st.session_state.conectado = False
    st.session_state.headers = None
    st.session_state.site_id = None
    st.session_state.verificacion_exitosa = False

# Inputs para las rutas de los archivos
ruta_fija = st.text_input(
    "Ruta completa del archivo FIJO en SharePoint",
    "Documentos compartidos/01 Archivos Area Administrativa/TRM.xlsx"
)
ruta_carpeta_mensual = st.text_input(
    "Ruta de la CARPETA que contiene los archivos mensuales",
    "Documentos compartidos/Ventas con ciudad 2025"
)

if st.button("Conectar y Verificar Archivos"):
    with st.spinner("Autenticando y buscando archivos..."):
        # Limpiamos el estado anterior para una nueva verificación
        st.session_state.conectado = False
        st.session_state.verificacion_exitosa = False

        token = get_access_token()
        
        # VERIFICACIÓN PASO A PASO
        if token:
            st.session_state.headers = {'Authorization': f'Bearer {token}'}
            site_id = get_sharepoint_site_id(token)
            
            if site_id:
                st.session_state.site_id = site_id
                st.session_state.conectado = True

# Esta parte se ejecuta FUERA del spinner para que los mensajes finales sean visibles
if st.session_state.conectado:
    st.markdown("---")
    st.success("🎉 ¡Conexión con SharePoint exitosa!")
    
    # HERRAMIENTAS DE EXPLORACIÓN
    explorador_completo_sharepoint(st.session_state.site_id, st.session_state.headers)
    
    st.markdown("---")
    st.header("📋 Verificación Manual de Rutas")
    st.info("Una vez que hayas encontrado las rutas correctas arriba, ingrésalas aquí:")
    
    # Permitir al usuario ingresar las rutas correctas encontradas
    col1, col2 = st.columns(2)
    
    with col1:
        ruta_fija_corregida = st.text_input(
            "Ruta CORRECTA del archivo TRM.xlsx:",
            value=ruta_fija,  # Valor por defecto
            help="Usa las herramientas de exploración arriba para encontrar la ruta exacta"
        )
    
    with col2:
        ruta_carpeta_corregida = st.text_input(
            "Ruta CORRECTA de la carpeta mensual:",
            value=ruta_carpeta_mensual,  # Valor por defecto
            help="Carpeta que contiene los archivos mensuales"
        )
    
    # Botón para verificar con las rutas corregidas
    if st.button("🚀 Verificar con Rutas Correctas"):
        if ruta_fija_corregida and ruta_carpeta_corregida:
            with st.spinner("Verificando con las rutas corregidas..."):
                
                st.subheader("🔍 Verificando archivo fijo")
                check1 = verificar_archivo_alternativo(
                    st.session_state.site_id, 
                    st.session_state.headers, 
                    ruta_fija_corregida
                )[0]
                
                st.subheader("📅 Verificando archivo mensual") 
                nombre_mes, ruta_mes = encontrar_archivo_del_mes_en_carpeta(
                    st.session_state.site_id, 
                    st.session_state.headers, 
                    ruta_carpeta_corregida
                )
                
                if check1 and nombre_mes:
                    st.session_state.verificacion_exitosa = True
                    st.success("🎉 ¡Todas las verificaciones fueron exitosas!")
                    st.balloons()
                    
                    # Guardar las rutas correctas en session state
                    st.session_state.ruta_fija_final = ruta_fija_corregida
                    st.session_state.ruta_carpeta_final = ruta_carpeta_corregida
                    st.session_state.nombre_archivo_mes = nombre_mes
                    st.session_state.ruta_archivo_mes = ruta_mes
                    
                    # Mostrar resumen final
                    st.markdown("---")
                    st.subheader("📋 Resumen de Rutas Verificadas")
                    st.success(f"✅ Archivo fijo encontrado: `{ruta_fija_corregida}`")
                    st.success(f"✅ Archivo mensual encontrado: `{nombre_mes}` en `{ruta_carpeta_corregida}`")
                    
                else:
                    st.session_state.verificacion_exitosa = False
                    st.error("❌ Una o ambas verificaciones fallaron con las rutas corregidas.")
                    st.info("💡 Usa las herramientas de exploración arriba para encontrar las rutas exactas.")
        else:
            st.warning("⚠️ Por favor, ingresa ambas rutas antes de verificar.")

else:
    st.error("El proceso se detuvo porque la conexión con SharePoint falló. Revisa las credenciales y nombres del sitio.")

# También agrega esta sección de ayuda al final:
st.markdown("---")
st.subheader("❓ Cómo usar las herramientas de exploración")

with st.expander("📖 Guía de uso", expanded=False):
    st.markdown("""
    ### Pasos recomendados:
    
    1. **🔍 Explorar Raíz**: 
       - Ve a la pestaña "📁 Raíz" y haz clic en "Explorar Raíz del Sitio"
       - Esto te mostrará todas las carpetas principales
    
    2. **🧪 Probar Nombres Comunes**:
       - Ve a la pestaña "📂 Carpetas" 
       - Haz clic en "Buscar Carpetas de Documentos Comunes"
       - Esto probará nombres típicos como "Shared Documents"
    
    3. **🔍 Búsqueda Global**:
       - Ve a la pestaña "🔍 Búsqueda"
       - Busca "TRM.xlsx" para encontrar el archivo exacto y su ubicación
    
    4. **💡 Generar Sugerencias**:
       - Ve a la pestaña "💡 Sugerencias"  
       - Esto generará rutas probables basadas en las carpetas encontradas
    
    5. **✅ Verificar**:
       - Usa las rutas encontradas en la sección "📋 Verificación Manual"
    
    ### Errores comunes:
    - **"Documentos compartidos"** a menudo es **"Shared Documents"** en inglés
    - Las rutas son **case-sensitive** (importan mayúsculas/minúsculas)
    - Los espacios deben coincidir exactamente
    """)

# --- PASO 2: PROCESAMIENTO DEL ARCHIVO LOCAL (Solo si la verificación fue exitosa) ---
if st.session_state.get('verificacion_exitosa'):
    st.markdown("---")
    st.header("2. Procesamiento del Archivo Local")
    st.info("Las verificaciones en SharePoint fueron exitosas. Ahora puedes subir y procesar tu archivo.")

    uploaded_file = st.file_uploader(
        "Sube tu archivo Excel (.xlsx) para procesar",
        type=["xlsx"]
    )

    if uploaded_file is not None:
        if st.button("Iniciar Procesamiento"):
            df_result = procesar_excel_para_streamlit(uploaded_file)
            if df_result is not None:
                st.dataframe(df_result.head())
                st.success("Tu archivo ha sido procesado y está listo para los siguientes pasos (ej. ser combinado y subido a SharePoint).")
                # Aquí iría la lógica para combinar df_result con los datos de SharePoint y subirlos.
else:
    st.info("Por favor, introduce las rutas de SharePoint y haz clic en 'Conectar y Verificar' para comenzar.")
    
    



