import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import openpyxl
import base64

def put_comments(excel_file, tds_file):
    df = pd.read_excel(excel_file, header=None, usecols=[0, 1], skiprows=1, names=['caption', 'comment'])
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

st.title("Automatización de Comentarios para Tableau")

excel_file = st.file_uploader("Cargar archivo Excel", type=["xlsx"])

tds_file = st.file_uploader("Cargar archivo .tds", type=["tds"])

if excel_file and tds_file:
    if st.button("Aplicar Automatización"):
        new_tds_tree = put_comments(excel_file, tds_file)

        tds_file_name = tds_file.name

        new_tds_bytes = ET.tostring(new_tds_tree.getroot(), encoding="utf-8")

        st.markdown(
            f'<a href="data:application/octet-stream;base64,{base64.b64encode(new_tds_bytes).decode()}" download="{tds_file_name}">Descargar archivo .tds resultante</a>',
            unsafe_allow_html=True
        )
