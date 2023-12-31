import pandas as pd
import plotly.express as px
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

        nota_element = root.find(".//{http://www.portalfiscal.inf.br/nfe}nNF")
        nota = nota_element.text if nota_element is not None else ""

        # Do the same for the other elements
        emissor_element = root.find(".//{http://www.portalfiscal.inf.br/nfe}xNome")
        emissor = emissor_element.text if emissor_element is not None else ""

        cliente_element = root.find(".//{http://www.portalfiscal.inf.br/nfe}dest//{http://www.portalfiscal.inf.br/nfe}xNome")
        cliente = cliente_element.text if cliente_element is not None else ""

        rua_element = root.find(".//{http://www.portalfiscal.inf.br/nfe}dest//{http://www.portalfiscal.inf.br/nfe}xLgr")
        rua = rua_element.text if rua_element is not None else ""

        numero_element = root.find(".//{http://www.portalfiscal.inf.br/nfe}dest//{http://www.portalfiscal.inf.br/nfe}nro")
        numero = numero_element.text if numero_element is not None else ""

        municipio_element = root.find(".//{http://www.portalfiscal.inf.br/nfe}dest//{http://www.portalfiscal.inf.br/nfe}xMun")
        municipio = municipio_element.text if municipio_element is not None else ""

        peso_bruto_element = root.find(".//{http://www.portalfiscal.inf.br/nfe}vol//{http://www.portalfiscal.inf.br/nfe}pesoB")
        peso_bruto = peso_bruto_element.text if peso_bruto_element is not None else ""

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
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="NotaFiscal.xlsx">Download arquivo Excel</a>'

def main():
    st.image("imagens/IMG_20230726_125320.png")
    st.divider()
    with st.expander('Sobre o NF-e Data Converter'):
        st.write("""
            O NF-e Data Converter é uma ferramenta criada para facilitar o processo de extração de informações relevantes de arquivos de Nota Fiscal Eletrônica (NF-e) em formato XML. 
            A ferramenta foi desenvolvida para simplificar a conversão de arquivos XML de notas fiscais em um formato de dados mais fácil de analisar e manipular.

            FUNCIONALIDADES PRINCIPAIS:
          
           ◽Carregamento de Múltiplos Arquivos: A ferramenta permite o carregamento de vários arquivos XML de uma só vez, tornando o processo mais eficiente.
        
           ◽Extração de Dados: Os dados são extraídos de cada arquivo XML, incluindo informações importantes como número da nota, emissor, cliente, endereço, município e peso bruto.
        
           ◽Tratamento de Erros: A ferramenta inclui tratamento de erros robusto, garantindo que os arquivos sejam analisados corretamente e que os erros sejam informados de forma clara e compreensível.
        
           ◽Exportação de Dados: Os dados extraídos podem ser exportados para um arquivo Excel para análise posterior, tudo com apenas um clique.
        
           ◽Visualização de Dados: Os dados extraídos são exibidos em um DataFrame para uma visualização clara e organizada das informações.""")
    with st.expander('Setores Beneficiados'):
        st.write("""
           ◽Logística: Otimiza a gestão de cargas através da análise rápida de dados das NF-e.
           
           ◽Varejo: Melhora a gestão de vendas e compras ao processar informações de NF-e em grande volume.
           
           ◽Finanças: Proporciona uma coleta precisa e organizada de informações para controle financeiro e tributário.
           
           ◽Jurídico: Facilita a extração de informações para casos de litígios comerciais.
           
           ◽Tecnologia da Informação: Auxilia no desenvolvimento de soluções para gerenciamento de NF-e. """)
    
    st.divider()
    uploaded_files = st.file_uploader("Carregue seus arquivos XML", type=['xml'], accept_multiple_files=True)
    st.divider()

    if uploaded_files:
        data = []
        with st.spinner('Processando arquivos...'):
            for uploaded_file in uploaded_files:
                data.append(extract_info_from_xml(uploaded_file))

        #st.divider()
        df = pd.DataFrame(data, columns=["Nota", "Emissor", "Cliente", "Rua", "Numero", "Municipio", "Peso Bruto"])
        
        # Remove notas fiscais duplicadas
        df = df.drop_duplicates(subset=['Nota'], keep='first')
        
        st.subheader('Dados Extraídos')
        st.dataframe(df)

        # Analysis
        st.subheader("Análise de Dados")

        total_notas = df['Nota'].nunique()
        unique_emissores = df['Emissor'].nunique()
        unique_clientes = df['Cliente'].nunique()
        nota_mais_comum = df['Nota'].mode()[0]
        cliente_mais_comum = df['Cliente'].mode()[0]
            
        # Criar um data frame para o resumo
        summary_df = pd.DataFrame({
                                    "Número total de notas fiscais": [total_notas],
                                    "Número de emissores únicos": [unique_emissores],
                                    "Número de clientes únicos": [unique_clientes],
                                    "Nota fiscal mais comum": [nota_mais_comum],
                                    "Cliente mais comum": [cliente_mais_comum]})
            
        st.table(summary_df)

        if st.button('Exportar para Excel'):
            with st.spinner('Convertendo para Excel...'):
                sleep(3)
                st.info('Conversão concluída, Clique no link abaixo para efetuar o download.')
                st.markdown(get_table_download_link(df), unsafe_allow_html=True)
        st.divider()
    st.markdown("Developed by: Mauro Alves")
    st.markdown("""
        <a href="https://github.com/Maurynn" target="_blank" style="margin-right: 15px; text-decoration: none">
        <img src="https://github.githubassets.com/images/modules/logos_page/GitHub-Mark.png" alt="Github logo" width="25" height="25">
        </a>
        <a href="https://linkedin.com/in/maurosp" target="_blank" style="margin-right: 15px; text-decoration: none">
        <img src="https://content.linkedin.com/content/dam/me/business/en-us/amp/brand-site/v2/bg/LI-Bug.svg.original.svg" alt="LinkedIn logo" width="25" height="25">
        </a>
        <a href="https://instagram.com/maurinn?igshid=ZDc4ODBmNjlmNQ==" target="_blank" style="margin-right: 15px; text-decoration: none">
        <img src="https://upload.wikimedia.org/wikipedia/commons/a/a5/Instagram_icon.png" alt="Instagram logo" width="25" height="25">
        </a>
        <a href="https://wa.me/5511952483074" target="_blank" style="margin-right: 15px; text-decoration: none">
        <img src="https://upload.wikimedia.org/wikipedia/commons/6/6b/WhatsApp.svg" alt="WhatsApp logo" width="25" height="25">
        </a>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
