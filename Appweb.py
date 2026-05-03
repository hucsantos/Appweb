import streamlit as st
import pandas as pd
import requests
from io import BytesIO
import openpyxl
import plotly.express as px
import os
from dotenv import load_dotenv

# Configuração da página
st.set_page_config(page_title="Dashboard Backlog", layout="wide")

@st.cache_data
def load():
    # URL e Headers conforme o seu código original
    url = url2
    headers = {"DeskManager": header2}
    
    try:
        # Faz a requisição para baixar o Excel
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        
        # Lê o Excel (corresponde ao Excel.Workbook e Sheet no Power Query)
        # O pandas já promove os cabeçalhos por padrão (header=0)
        df = pd.read_excel(BytesIO(response.content))

        return df
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return pd.DataFrame()
    
# Título Principal

st.title("Análise de Fechados")

# Botão para carregar/atualizar
if st.button('Atualizar Dados'):
    st.cache_data.clear() # Limpa o cache para forçar nova busca

df = load()

if not df.empty:
    # --- ÁREA DE MÉTRICAS (CARDS) ---
    col1, col2, col3, col4 = st.columns(4)
    
    Fechados = len(df)
    total_grupos = df['Nome do Grupo'].nunique()

# Identifica a quantidadede chamados expirados

    col_sla = 'SLA de Solução Expirado'

    if col_sla in df.columns:
        # Conta quantos chamados estão com SLA "Sim" (Expirado)
        # O .strip() e .upper() garantem que espaços ou letras minúsculas não quebrem o cálculo
        qtd_perdidos = len(df[df[col_sla] == "Expirado"])
        percentual_perdido = (qtd_perdidos / Fechados) * 100
    else:
        qtd_perdidos = 0
        percentual_perdido = 0


# --- TABELA DINÂMICA: GRUPOS (LINHAS) vs MESES (COLUNAS) ---

# --- TABELA DINÂMICA COM REGRAS DE META POR GRUPO ---
st.subheader("% SLA de Solução - Grupo")

if 'Data de Finalização' in df.columns and 'Nome do Grupo' in df.columns:
    # 1. Preparação dos dados
    df_temp = df.copy()
    df_temp['Data de Finalização'] = pd.to_datetime(df_temp['Data de Finalização'])
    df_temp['Mês'] = df_temp['Data de Finalização'].dt.to_period('M').astype(str)
    df_temp['SLA_Aux'] = df_temp['SLA de Solução Expirado'] == "Em Dia"

    # 2. Filtro por Grupo
    grupos_disponiveis = sorted(df_temp['Nome do Grupo'].unique())
    selecao_grupos = st.multiselect("Filtrar Grupos:", options=grupos_disponiveis, default=["Infraestrutura - TI Quality", "Suporte - Nível 1", "Suporte - Nível 2"])

    if selecao_grupos:
        df_temp = df_temp[df_temp['Nome do Grupo'].isin(selecao_grupos)]

    # 3. Agrupamento e Pivot
    df_agrupado = df_temp.groupby(['Mês', 'Nome do Grupo']).agg(
        Total=('SLA_Aux', 'count'),
        Expirados=('SLA_Aux', 'sum')
    ).reset_index()
    
    df_agrupado['% SLA Perdido'] = (df_agrupado['Expirados'] / df_agrupado['Total'] * 100).round(2)
    tabela_pivot = df_agrupado.pivot(index='Nome do Grupo', columns='Mês', values='% SLA Perdido').fillna(0)

    # 4. FUNÇÃO DE FORMATAÇÃO CONDICIONAL ESPECÍFICA
    def aplicar_metas_criticas(row):
        # 'row' aqui é uma Series onde row.name é o nome do Grupo
        grupo = str(row.name).strip()
        estilos = []
        
        for valor in row:
            cor = ""
            # Regra 1: Infraestrutura - TI Quality e Suporte - Nível 2 (Abaixo de 60%)
            if grupo in ["Infraestrutura - TI Quality", "Suporte - Nível 2"]:
                if valor < 60:
                    cor = "color: #FF4B4B;"
                elif valor >= 60:
                    cor = "color: #7fff00;"
            
            # Regra 2: Suporte - Nível 1 (Abaixo de 85%)
            elif grupo == "Suporte - Nível 1":
                if valor < 85:
                    cor = "color: #FF4B4B;"
                elif valor >= 85:
                    cor = "color: #7fff00;"
            
            estilos.append(cor)
        return estilos

    # 5. Exibição com Estilo
    st.dataframe(
        tabela_pivot.style
        .apply(aplicar_metas_criticas, axis=1) # Aplica a regra linha a linha (por grupo)
        .format("{:.2f}%")
        .set_properties(**{
            'text-align': 'center',
            'border': '0.1px solid #ededed'
        })
        .set_table_styles([dict(selector='th', props=[('text-align', 'center')])]),
        use_container_width=True
    )
else:
    st.error("Colunas necessárias não encontradas.")


# --- TABELA DINÂMICA COM REGRAS DE META POR GRUPO POR OPERADOR ---
st.subheader("% SLA de Solução - Operador")

if 'Data de Finalização' in df.columns and 'Nome Completo do Operador' in df.columns:
    # 1. Preparação dos dados
    df_temp = df.copy()
    df_temp['Data de Finalização'] = pd.to_datetime(df_temp['Data de Finalização'])
    df_temp['Mês'] = df_temp['Data de Finalização'].dt.to_period('M').astype(str)
    df_temp['SLA_Aux'] = df_temp['SLA de Solução Expirado'] == "Em Dia"

    # 2. Filtro por Grupo
    grupos_disponiveis = sorted(df_temp['Nome do Grupo'].unique())
    selecao_grupos = st.multiselect("Filtrar Grupos:", options=grupos_disponiveis)

    if selecao_grupos:
        df_temp = df_temp[df_temp['Nome do Grupo'].isin(selecao_grupos)]

    # 3. Agrupamento e Pivot
    df_agrupado = df_temp.groupby(['Mês', 'Nome Completo do Operador', 'Nome do Grupo']).agg(
        Total=('SLA_Aux', 'count'),
        Expirados=('SLA_Aux', 'sum')
    ).reset_index()
    
    df_agrupado['% SLA Perdido'] = (df_agrupado['Expirados'] / df_agrupado['Total'] * 100).round(2)
    tabela_pivot = df_agrupado.pivot(index='Nome Completo do Operador', columns='Mês', values='% SLA Perdido').fillna(0)

    # 4. FUNÇÃO DE FORMATAÇÃO CONDICIONAL ESPECÍFICA
    def aplicar_metas_criticas(row):
        # 'row' aqui é uma Series onde row.name é o nome do Grupo
        estilos = []
        
        for valor in row:
            cor = ""
            if valor < 80:
                    cor = "color: #FF4B4B;"
            elif valor >= 80:
                    cor = "color: #7fff00;"            
            estilos.append(cor)
        return estilos

    # 5. Exibição com Estilo
    st.dataframe(
        tabela_pivot.style
        .apply(aplicar_metas_criticas, axis=1) # Aplica a regra linha a linha (por grupo)
        .format("{:.2f}%")
        .set_properties(**{
            'text-align': 'center',
            'border': '0.1px solid #ededed'
        })
        .set_table_styles([dict(selector='th', props=[('text-align', 'center')])]),
        use_container_width=True
    )
else:
    st.error("Colunas necessárias não encontradas.")

