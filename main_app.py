import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import base64

def put_comments(excel_file, tds_file):
    df = pd.read_excel(excel_file, header=None, usecols=[0, 1], names=['caption', 'comment'])
    tree = ET.parse(tds_file)
    root = tree.getroot()

    for _, row in df.iterrows():
        caption = row['caption']
        comment = row['comment']

        for column in root.findall('.//column[@caption="{}"]'.format(caption)):
            existing_comment = column.find('.//desc/formatted-text/run')
            if existing_comment is not None:
                existing_comment.text = comment
            else:
                desc = ET.SubElement(column, 'desc')
                formatted_text = ET.SubElement(desc, 'formatted-text')
                run = ET.SubElement(formatted_text, 'run')
                run.text = comment

    return tree

def put_names(excel_file, tds_file):
    df = pd.read_excel(excel_file, header=None, usecols=[0, 1], names=['current_name', 'new_name'])
    tree = ET.parse(tds_file)
    root = tree.getroot()

    for _, row in df.iterrows():
        current_name = row['current_name']
        new_name = row['new_name']

        # Buscar la columna existente con el nombre actual
        column = root.find('.//column[@caption="{}"]'.format(current_name))

        if column is not None:
            # Si se encuentra una columna con el nombre actual, reemplazar el atributo "caption"
            column.set('caption', new_name)
        else:
            # Si no se encuentra una columna con el nombre actual, crear una nueva columna
            # y agregarla después de la primera columna existente
            first_column = root.find('.//column')
            if first_column is not None:
                new_column = ET.Element('column', {'caption': new_name, 'datatype': 'integer', 'name': '[{}]'.format(current_name), 'role': 'measure', 'type': 'quantitative'})
                root.insert(list(root).index(first_column) + 1, new_column)

    return tree

st.title("Automatización para Tableau")
option = st.selectbox("Seleccione una funcionalidad", ("", "Automatizar Comentarios", "Automatizar Nombres"))

if option == "Automatizar Comentarios":
    excel_file = st.file_uploader("Cargar archivo Excel para Comentarios", type=["xlsx"])
    tds_file = st.file_uploader("Cargar archivo .tds para Comentarios", type=["tds"])
    if excel_file and tds_file:
        if st.button("Aplicar Automatización"):
            new_tds_tree = put_comments(excel_file, tds_file)
            tds_file_name = tds_file.name
            new_tds_bytes = ET.tostring(new_tds_tree.getroot(), encoding="utf-8")
            st.markdown(
                f'<a href="data:application/octet-stream;base64,{base64.b64encode(new_tds_bytes).decode()}" download="{tds_file_name}">Descargar archivo .tds resultante</a>',
                unsafe_allow_html=True
            )

elif option == "Automatizar Nombres":
    excel_file = st.file_uploader("Cargar archivo Excel para Nombres", type=["xlsx"])
    tds_file = st.file_uploader("Cargar archivo .tds para Nombres", type=["tds"])
    if excel_file and tds_file:
        if st.button("Aplicar Automatización"):
            new_tds_tree = put_names(excel_file, tds_file)
            tds_file_name = tds_file.name
            new_tds_bytes = ET.tostring(new_tds_tree.getroot(), encoding="utf-8")
            st.markdown(
                f'<a href="data:application/octet-stream;base64,{base64.b64encode(new_tds_bytes).decode()}" download="{tds_file_name}">Descargar archivo .tds resultante</a>',
                unsafe_allow_html=True
            )