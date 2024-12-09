import pandas as pd
from docx import Document
import io  # Para manejar archivos en memoria

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

    # Definir patrones a excluir
    exclude_patterns = [
        "PAGE NO :", "|REGISTERED OFFICE:", "BOX ITEM TOTAL", "|MARUTI", 
        "********", "|Plot No.", "|Vasant Kunj", "|Pan", "|DECLARATION"
    ]

    # Iterar sobre los párrafos en el documento
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if start_text in text:
            found_start = True
        if found_start and marker_text in text:
            found_marker = True
            continue
        if found_marker:
            if any(text.startswith(pattern) for pattern in exclude_patterns):
                continue
            extracted_content.append(text)

    # Convertir el contenido extraído en una cadena y luego a un DataFrame
    combined_content = "\n".join(extracted_content)
    df = pd.read_csv(io.StringIO(combined_content), sep="|")

    # Limpiar y procesar el DataFrame
    df.columns = [col.strip() for col in df.columns]
    columnas_drop = ['Unnamed: 0', 'Unnamed: 13']
    df.drop(columns=[col for col in columnas_drop if col in df.columns], inplace=True)

    # Aplicar reglas de limpieza
    df_cleaned = df[
        ~df['ITEM CODE'].isna() &
        (df['ITEM CODE'] != '') &
        (df['ITEM CODE'] != 'ITEM CODE') &
        (~df['ITEM CODE'].str.contains('ORDER ITEM CODE', na=False))
    ]
    df_cleaned['ITEM CODE'] = df_cleaned['ITEM CODE'].apply(lambda x: x + '-000' if len(x) < 14 else x)

    # Renombrar columnas y agregar ajustes
    df_cleaned.rename(columns={'ITEM CODE': 'MAT_PROV_SOLICITADO'}, inplace=True)
    df_cleaned['NRO_ORDEN_PREFIJO'] = 'CHL' + df_cleaned['ORDER REF NO'] + '-24'

    # Seleccionar columnas de interés
    df_cleaned = df_cleaned[
        ['NRO_ORDEN_PREFIJO', 'MAT_PROV_SOLICITADO', 'QTY', 'UNIT RATE']
    ]
    return df_cleaned
