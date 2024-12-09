import pandas as pd
from docx import Document
import io  # Para manejar archivos en memoria
import pyperclip
import numpy as np

def procesar_factura(file):
    """
    Procesa un archivo Word cargado como un objeto en memoria (file) 
    y devuelve un DataFrame limpio.
    """
    # Leer el archivo Word desde el objeto en memoria
    document = Document(io.BytesIO(file.read()))

    # Identificar el texto relevante y excluir patrones innecesarios
    start_text = "********  INVOICE CUM PACKING LIST ANNEXURE ********"
    marker_text = "===================="
    found_start = False
    found_marker = False
    extracted_content = []

    # Definir los patrones a excluir
    exclude_patterns = [
        "PAGE NO :", "|REGISTERED OFFICE:", "BOX ITEM TOTAL", "|MARUTI", "********","|Plot No.","|Vasant Kunj","|Pan","|DECLARATION"
    ]

    # Iterar sobre los párrafos en el documento
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        
        # Verificar si el texto del párrafo coincide con el inicio de la sección deseada
        if start_text in text:
            found_start = True
        
        # Si se encuentra el punto de inicio, verificar si aparece el marcador
        if found_start and marker_text in text:
            found_marker = True
            continue
        
        # Si se encuentra el marcador, comenzar a recopilar contenido
        if found_marker:
            # Omitir si el texto coincide con algún patrón de exclusión
            if any(text.startswith(pattern) for pattern in exclude_patterns):
                continue
            
            # Agregar el contenido válido a la lista
            extracted_content.append(text)

    # Combinar el contenido extraído en una sola cadena
    combined_content = "\n".join(extracted_content)

    # Imprimir el contenido combinado
    print(combined_content)

    # Opcionalmente, copiar el contenido al portapapeles
    pyperclip.copy(combined_content)

    # Opcionalmente, guardar el contenido en un archivo
    with open('extracted_content.txt', 'w') as file:
        file.write(combined_content)

    print("Contenido copiado al portapapeles y guardado en 'extracted_content.txt'.")

    # Convertir el contenido copiado a DataFrame
    df = pd.read_clipboard(sep='|')

    # Mostrar el DataFrame
    print(df)

    # Limpiar y procesar el DataFrame
    columnas_drop = ['Unnamed: 0', 'Unnamed: 13']
    df.drop(columns=columnas_drop, inplace=True)
    df.columns = [col.strip() for col in df.columns]

    df_cleaned = df[
        ~df['ITEM CODE'].isna() &
        (df['ITEM CODE'] != '') &
        (df['ITEM CODE'] != 'ITEM CODE') &
        (~df['ITEM CODE'].str.contains('ORDER ITEM CODE', na=False))
    ]

    df_cleaned = df_cleaned[~df_cleaned['BOX NO'].str.contains('ORDER ITEM CODE', na=False)]
    df_cleaned = df_cleaned[~df_cleaned['SNO'].str.contains('SNO')]
    df_cleaned['SNO'] = df_cleaned['SNO'].str.strip()

    df_cleaned['SNO'].replace('', np.nan, inplace=True)

    # Desplazar la columna 'BOX NO' una fila hacia abajo
    df_cleaned['BOX NO shifted'] = df_cleaned['BOX NO'].shift(-1)
    df_cleaned['BOX NO shifted'] = df_cleaned['BOX NO shifted'].str.replace(' ', '').str.replace('(', '').str.replace(')', '')
    df_cleaned['ITEM CODE 2'] = df_cleaned['ITEM CODE']
    # Reemplazar 'ITEM CODE' con 'BOX NO shifted' donde 'SNO' es NaN
    df_cleaned.loc[df_cleaned['BOX NO shifted'].str.len() > 6, 'ITEM CODE'] = df_cleaned['BOX NO shifted']

    df_cleaned = df_cleaned.dropna(subset=['SNO'])
    df_cleaned.drop(columns=['BOX NO shifted'], inplace=True)
    df_cleaned = df_cleaned.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    df_cleaned['ITEM CODE'] = df_cleaned['ITEM CODE'].apply(lambda x: x + '-000' if len(x) < 14 else x)
    df_cleaned['ITEM CODE 2'] = df_cleaned['ITEM CODE 2'].apply(lambda x: x + '-000' if len(x) < 14 else x)
    # Convertir las columnas a tipo numérico
    df_cleaned['VOLUME'] = pd.to_numeric(df_cleaned['VOLUME'], errors='coerce')
    df_cleaned['WEIGHT'] = pd.to_numeric(df_cleaned['WEIGHT'], errors='coerce')
    df_cleaned['AMOUNT'] = pd.to_numeric(df_cleaned['AMOUNT'], errors='coerce')
    df_cleaned['UNIT RATE'] = pd.to_numeric(df_cleaned['UNIT RATE'], errors='coerce')
    df_cleaned['QTY'] = pd.to_numeric(df_cleaned['QTY'], errors='coerce')
    df_cleaned.rename(columns = {'ITEM CODE': 'MAT_PROV_SOLICITADO','ITEM CODE 2': 'MATERIAL_PROV_DESPACHADO','UNIT RATE': 'VALOR_UNITARIO_FACTURA'}, inplace=True)

    df_cleaned['QTY_FACTURADA'] = df_cleaned['QTY']
    df_cleaned['UNIDAD_MEDIDA_FACTURADA'] = 'UN'



    # Añadir una nueva columna 'NRO_ORDEN'
    df_cleaned['NRO_ORDEN_PREFIJO'] = 'CHL' + df_cleaned['ORDER REF NO'] + '-24'
    df_cleaned = df_cleaned[['NRO_ORDEN_PREFIJO','MAT_PROV_SOLICITADO','MATERIAL_PROV_DESPACHADO','QTY','QTY_FACTURADA','VALOR_UNITARIO_FACTURA','UNIDAD_MEDIDA_FACTURADA']]
    # Obtener el nombre del archivo sin extensión
    return df_cleaned
