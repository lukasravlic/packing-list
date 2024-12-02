import streamlit as st
import pandas as pd
from io import BytesIO  # Para crear archivos en memoria

# Configuración de la página
st.set_page_config(
    page_title="Proceso de consolidación packing list",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Sidebar para seleccionar la marca
with st.sidebar:
    st.header("Seleccione la Marca")
    brand = st.selectbox(
        "Elija la marca para procesar los datos:",
        options=["Maruti Suzuki"],  # Se pueden agregar más marcas en el futuro
    )

# Título principal
st.title("Packing List - Maruti Suzuki")
st.markdown("---")

# Subir y procesar archivos
st.subheader("Suba sus archivos Excel")

uploaded_files = st.file_uploader(
    "Cargue archivos Excel (.xlsx)",
    type="xlsx",
    accept_multiple_files=True
)

if uploaded_files:
    data_frames = []
    for uploaded_file in uploaded_files:
        try:
            # Leer cada archivo Excel con header=1
            df = pd.read_excel(uploaded_file, engine="openpyxl", header=1)

            # Eliminar la primera fila después del header
            df = df.iloc[1:, :]  # Esto elimina la fila en el índice 0

            data_frames.append(df)

            # Mostrar vista previa
            st.write(f"Archivo cargado: **{uploaded_file.name}**:")
        except Exception as e:
            st.error(f"Error al leer el archivo {uploaded_file.name}: {e}")

    # Combinar todos los archivos cargados
    if data_frames:
        combined_data = pd.concat(data_frames, ignore_index=True)
        st.success("Archivos combinados exitosamente.")

        combined_data = combined_data[
            ['Cod. Material de Proveedor despachado', 'Cantidad solicitada', 'Nro. De Orden – Prefijo']]

        # Inputs para agregar nuevas columnas
        st.subheader("Ingrese datos para agregar nuevas columnas")

        # Campo de texto para nueva columna
        new_column_text = st.text_input("DT", value="Numero de DT")

        # Dropdown para "Tipo de Contenedor"
        container_type = st.selectbox(
            "Seleccione el tipo de contenedor",
            options=["40HC", "4'STD", "Tipo 3", "Tipo 4"]
        )

        # Campo de texto para "Contenedor"
        container = st.text_input("Ingrese el nombre del contenedor", value="Contenedor por defecto")

        if new_column_text and container_type and container:
            # Agregar las nuevas columnas al DataFrame
            combined_data["DT"] = new_column_text
            combined_data["Tipo de Contenedor"] = container_type
            combined_data["Contenedor"] = container

            # Mostrar tabla con las nuevas columnas
            st.subheader("Tabla Combinada con Nuevas Columnas")
            st.dataframe(combined_data)

            # Crear un archivo Excel en memoria
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                combined_data.to_excel(writer, index=False, sheet_name="Datos Combinados")
            excel_data = output.getvalue()

            # Botón para descargar el archivo Excel
            st.download_button(
                label="Descargar tabla combinada como Excel",
                data=excel_data,
                file_name="datos_combinados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.warning("Por favor, suba uno o más archivos Excel para continuar.")