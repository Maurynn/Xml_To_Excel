import pandas as pd
import os
from time import sleep
import lxml.etree as ET
from typing import List
import streamlit as st
from io import BytesIO
import base64

def extract_info_from_xml(file) -> List[str]:
    try:
        root = ET.parse(file).getroot()

        nota = root.find(".//{http://www.portalfiscal.inf.br/nfe}nNF").text or ""
        emissor = root.find(".//{http://www.portalfiscal.inf.br/nfe}xNome").text or ""
        cliente = root.find(".//{http://www.portalfiscal.inf.br/nfe}dest//{http://www.portalfiscal.inf.br/nfe}xNome").text or ""
        rua = root.find(".//{http://www.portalfiscal.inf.br/nfe}dest//{http://www.portalfiscal.inf.br/nfe}xLgr").text or ""
        numero = root.find(".//{http://www.portalfiscal.inf.br/nfe}dest//{http://www.portalfiscal.inf.br/nfe}nro").text or ""
        municipio = root.find(".//{http://www.portalfiscal.inf.br/nfe}dest//{http://www.portalfiscal.inf.br/nfe}xMun").text or ""

        peso_bruto_element = root.find(".//{http://www.portalfiscal.inf.br/nfe}vol//{http://www.portalfiscal.inf.br/nfe}pesoB")
        if peso_bruto_element is not None:
            peso_bruto = peso_bruto_element.text
        else:
            peso_bruto = ""

        return [nota, emissor, cliente, rua, numero, municipio, peso_bruto]

    except ET.XMLSyntaxError:
        st.error(f"Erro ao analisar o arquivo: {file.name}")
        return [""] * 7  # Retorna uma lista de strings vazias

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.close()
    processed_data = output.getvalue()
    return processed_data

def get_table_download_link(df):
    val = to_excel(df)
    b64 = base64.b64encode(val)  # val looks like b'...'
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="NotaFiscal.xlsx">Download Excel file</a>' 

def main():
    st.image('imagens/IMG_20230725_094141.png')
    st.divider()
    uploaded_files = st.file_uploader("Carregue seus arquivos XML", type=['xml'], accept_multiple_files=True)

    if uploaded_files:
        data = []
        with st.spinner('Processando arquivos...'):
            for uploaded_file in uploaded_files:
                data.append(extract_info_from_xml(uploaded_file))

        st.divider()
        df = pd.DataFrame(data, columns=["Nota", "Emissor", "Cliente", "Rua", "Numero", "Municipio", "Peso Bruto"])
        st.header('Dados Extraídos')
        st.dataframe(df)

        st.divider()

        if st.button('Exportar para Excel'):
            with st.spinner('Convertendo para Excel...'):
                sleep(3)
                st.info('Conversão concluída, Clique no link abaixo para efetuar o download.')
                st.markdown(get_table_download_link(df), unsafe_allow_html=True)

if __name__ == "__main__":
    main()
