import streamlit as st
import pandas as pd
import requests
from io import BytesIO
import openpyxl
import plotly.express as px
import os
from dotenv import load_dotenv

st.markdown(
    """
    <style>
    [data-testid="stMetric"] {
        text-align: center;
        margin: auto;
    }
    [data-testid="stMetric"] label {
        text-align: center;
        width: fit-content;
        margin: auto;
    }
    [data-testid="stMetricDelta"] {
        text-align: center;
        width: fit-content;
        margin: auto; 
    }
    [data-testid="stHeadingWithActionElements"] {
        text-align: center;
        width: fit-content;
        margin: auto;         
    }
    # [data-testid="stSidebar"] {
    #     min-width: 200px;
    #     max-width: 200px;
    # }
    </style>
    """,
    unsafe_allow_html=True
)


# Carrega as variáveis do arquivo .env
load_dotenv()

# --- CONFIGURAÇÃO DA INTERFACE ---
st.set_page_config(page_title="Sistema Seguro", layout="wide", initial_sidebar_state=200)


# --- LÓGICA DE AUTENTICAÇÃO ---
def verificar_credenciais(usuario_digitado, senha_digitada):
    """
    Verifica se existe uma variável de ambiente USER_{USUARIO} 
    com a senha correspondente.
    """
    # Transforma o input em maiúsculas para bater com o padrão do .env (USER_ADMIN)
    chave_env = f"USER_{usuario_digitado.upper()}"
    senha_correta = os.getenv(chave_env)
    
    if senha_correta and senha_digitada == senha_correta:
        return True
    return False

# --- CONTROLE DE SESSÃO ---
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.usuario_nome = ""

# --- VIEWS (TELAS) ---

def tela_login():

    # Centralizando o formulário de login
    _, col_central, _ = st.columns([1, 1, 1])
    with col_central:
        st.subheader("Login de Acesso")
    
        with st.form(key="form_login"):
            user = st.text_input("Usuário")
            password = st.text_input("Senha", type="password")
            botao = st.form_submit_button("Entrar", use_container_width=True)
        
        if botao:
            if verificar_credenciais(user, password):
                st.session_state.autenticado = True
                st.session_state.usuario_nome = user
                st.rerun()
            else:
                st.error("Usuário ou senha inválidos.")


def loading():

    with st.sidebar:
            if st.button("Sair", icon=":material/logout:", width="content", type="tertiary"):
                st.session_state.autenticado = False
                st.session_state.usuario_nome = ""
                st.rerun()
            if st.button('Atualizar', icon=":material/sync:", width="content"):
                st.cache_data.clear() # Limpa o cache para forçar nova busca

    # # Cabeçalho com Logout no canto superior
    # col_tit, col_log = st.columns([0.90, 0.1])
    # # with col_tit:
    # #     st.title(f"Bem-vindo, {st.session_state.usuario_nome}!")
    # with col_log:
    #     #st.write("") # Alinhamento
    #     if st.button("Sair", icon=":material/logout:", width="content"):
    #         st.session_state.autenticado = False
    #         st.session_state.usuario_nome = ""
    #         st.rerun()

    #st.divider()

    @st.cache_data
    def load():
            # URL e Headers conforme o seu código original
            url = os.getenv("API_URL")
            headers = {"DeskManager": os.getenv("API_KEY")}
            
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
    # st.title("Análise de Backlog")

        # Botão para carregar/atualizar
    # if st.button('Atualizar Dados'):
    #         st.cache_data.clear() # Limpa o cache para forçar nova busca

    df = load()

# ------------> Indicadores de Performance do Backlo

    st.header("Análise de Backlog")

    if not df.empty:
            # --- ÁREA DE MÉTRICAS (CARDS) ---
            col1, col2, col3, col4, col5, col6  = st.columns(6)
            
            Backlog = len(df)
            total_grupos = df['Nome do Grupo'].nunique()

        # Identifica a quantidadede chamados expirados no Backlog

            col_sla = 'SLA de Solução Expirado'

            if col_sla in df.columns:
                # Conta quantos chamados estão com SLA "Sim" (Expirado)
                # O .strip() e .upper() garantem que espaços ou letras minúsculas não quebrem o cálculo
                qtd_perdidos = len(df[df[col_sla] == "Expirado"])
                percentual_perdido = (qtd_perdidos / Backlog) * 100
            else:
                qtd_perdidos = 0
                percentual_perdido = 0
            
            with col1:
                st.metric(label="Total de Chamados", value=Backlog, border=True)

            with col3:
                # Card de Percentual de Perdidos (SLA Estourado)
                # Usei delta_color="inverse" porque quanto maior o percentual, "pior" é o resultado (fica vermelho)
                st.metric(
                    label="% SLA Perdido", 
                    value=f"{percentual_perdido:.1f}%",
                    delta=f"{qtd_perdidos} chamados",
                    delta_color="inverse",
                    border=True
                )
                
            with col2:
                # Exemplo de chamado mais antigo (Idade do Chamado)
                if "Idade do Chamado - Dias" in df.columns:
                    maior_idade = df["Idade do Chamado - Dias"].max()
                    st.metric(label="Maior Idade (Dias)", value=f"{maior_idade}", border=True)

            ## Metric for Groups

            # Lista dos grupos desejados
    grupos_alvo = ["Suporte - Nível 1", "Suporte - Nível 2", "Infraestrutura - TI Quality"]

        # Criamos as colunas para os cartões
    # col1, col2, col3 = st.columns(3)
    mapa_colunas = {
            "Suporte - Nível 1": col4,
            "Suporte - Nível 2": col5,
            "Infraestrutura - TI Quality": col6
        }

    if 'Nome do Grupo' in df.columns and 'SLA de Solução Expirado' in df.columns:
            # 1. Filtramos apenas o Backlog (Chamados sem data de finalização)
            df_backlog = df[df['Data de Finalização'].isna()].copy()
            
            # 2. Padronização da coluna de SLA
            df_backlog['SLA_Aux'] = df_backlog['SLA de Solução Expirado'] == "Expirado"

            for grupo in grupos_alvo:
                with mapa_colunas[grupo]:
                    # Filtra dados do grupo
                    dados_grupo = df_backlog[df_backlog['Nome do Grupo'] == grupo]
                    
                    total_backlog = len(dados_grupo)
                    total_expirados = dados_grupo['SLA_Aux'].sum()
                    
                    # Cálculo do percentual
                    perc_perdido = (total_expirados / total_backlog * 100) if total_backlog > 0 else 0
                    
                    # 3. Renderização com st.metric
                    # Usamos o delta para exibir o percentual de perda
                    # delta_color="inverse" faz com que o valor apareça em vermelho (ruim) se for positivo
                    st.metric(
                        label=grupo, 
                        value=f"{total_backlog}", 
                        delta=f"{perc_perdido:.2f}% Expirados",
                        delta_color="inverse", 
                        border=True,
                    )
                    
                    # Linha sutil de separação dentro da coluna
                    # st.caption(f"Expirados: {total_expirados}")

    else:
            st.error("Colunas necessárias para os cartões não foram encontradas.")

# ------------> Análise de idade e tempo sem interação nos chamados

    st.subheader("Idade do Chamado vs. Inatividade")

        # Verificação de colunas necessárias
    colunas_obrigatorias = [
            'Nome do Grupo', 
            'Idade do Chamado - Dias',  
            'Idade da Última Ação do Chamado - Dias', 
            'Data de Finalização'
        ]

    if all(col in df.columns for col in colunas_obrigatorias):
            
            # 1. Filtro único por Grupo para ambos os cenários
            grupos_disponiveis = sorted(df['Nome do Grupo'].unique())
            filtro_grupo = st.multiselect(
                "Filtrar grupos (Geral):", 
                options=grupos_disponiveis,
                key="filtro_unificado_aging"
            )

            # 2. Filtragem inicial (Somente Backlog / Não Finalizados)
            df_backlog = df[df['Data de Finalização'].isna()].copy()
            
            if filtro_grupo:
                df_backlog = df_backlog[df_backlog['Nome do Grupo'].isin(filtro_grupo)]

            if not df_backlog.empty:
                # 3. Tratamento e conversão dos dados
                # Idade do Chamado
                df_backlog['Idade do Chamado - Dias'] = pd.to_numeric(df_backlog['Idade do Chamado - Dias'], errors='coerce').fillna(0)
                
                # Inatividade (Limpando o " dias" se houver)
                df_backlog['Idade_Acao_Num'] = (
                    df_backlog['Idade da Última Ação do Chamado - Dias']
                    .astype(str)
                    .str.replace(' dias', '', case=False)
                    .str.strip()
                )
                df_backlog['Idade_Acao_Num'] = pd.to_numeric(df_backlog['Idade_Acao_Num'], errors='coerce').fillna(0)

                # 4. Definição das Faixas (Bins) comuns
                bins = [0, 3, 7, 15, 30, float('inf')]
                labels = ['Até 3 dias', '4 a 7 dias', '8 a 15 dias', '16 a 30 dias', 'Acima de 30 dias']
                
                df_backlog['Faixa_Idade'] = pd.cut(df_backlog['Idade do Chamado - Dias'], bins=bins, labels=labels)
                df_backlog['Faixa_Inatividade'] = pd.cut(df_backlog['Idade_Acao_Num'], bins=bins, labels=labels)

                # 5. Contagem e Agrupamento de ambas as métricas
                contagem_idade = df_backlog['Faixa_Idade'].value_counts().reindex(labels).reset_index()
                contagem_idade.columns = ['Faixa Etária', 'Idade do Chamado']

                contagem_inatividade = df_backlog['Faixa_Inatividade'].value_counts().reindex(labels).reset_index()
                contagem_inatividade.columns = ['Faixa Etária', 'Inatividade (Última Ação)']

                # Unificar as duas contagens em um único DataFrame
                df_unificado = pd.merge(contagem_idade, contagem_inatividade, on='Faixa Etária')

                # Transformar o formato para "Long" (ideal para o Plotly express agrupar as barras)
                df_melted = df_unificado.melt(
                    id_vars=['Faixa Etária'], 
                    value_vars=['Idade do Chamado', 'Inatividade (Última Ação)'],
                    var_name='Métrica', 
                    value_name='Quantidade'
                )

                # 6. Criação do Gráfico de Colunas Agrupadas
                fig_unificado = px.bar(
                    df_melted,
                    x='Faixa Etária',
                    y='Quantidade',
                    color='Métrica',
                    barmode='group', # Garante que as colunas fiquem lado a lado
                    text='Quantidade',
                    color_discrete_sequence=["#D7E4F5", "#86A6F6"], # Laranja Escuro (Idade) e Amarelo/Ouro (Inatividade)
                    labels={'Quantidade': 'Total de Chamados', 'Faixa Etária': 'Período de Tempo'}
                )

                # Ajustes de layout
                fig_unificado.update_traces(textposition='outside')
                fig_unificado.update_layout(
                    xaxis_title="Faixas de Tempo",
                    yaxis_title="Qtd Chamados",
                    legend=dict(
                        orientation='h',        # Corrigido de 'orientational' para 'orientation'
                        yanchor='bottom', 
                        y=1.02, 
                        xanchor='right', 
                        x=1
                    ),
                    margin=dict(t=50) # Espaço para a legenda no topo
                )

                # 7. Exibição do gráfico único
                st.plotly_chart(fig_unificado, use_container_width=True)
                
            else:
                st.info("Não há dados de backlog para os critérios selecionados.")
    else:
            st.error("Verifique se todas as colunas necessárias estão presentes no arquivo.")


# ------------> Tabela de Classificação de Prioridade de SLA

    grupos = st.multiselect("Filtrar por Classificação de SLA", options=df['Status do Tempo de Solução - Nome'].unique())
    df_filtrado = df if not grupos else df[df['Status do Tempo de Solução - Nome'].isin(grupos)].copy()

    if not df_filtrado.empty:
        # 1. Criando a URL completa
        df_filtrado['Link'] = (
            "https://santher.desk.ms/?Ticket#ChamadosSuporte:" + 
            df_filtrado['Nº Chamado'].astype(str) + 
            "RO"
        )

        # 2. Reordenando para que 'Link' seja a primeira coluna
        # Pegamos 'Link' e somamos com todas as outras colunas (exceto o próprio 'Link')
        cols = ['Link'] + [c for c in df_filtrado.columns if c != 'Link']
        df_visualizacao = df_filtrado[cols]

        # 3. Exibição formatada
        st.dataframe(
            df_visualizacao,
            column_config={
                "Link": st.column_config.LinkColumn(
                    "Desk Manager",          # Título da coluna
                    display_text="Consultar"  # Texto amigável
                )
            },
            use_container_width=True,
            hide_index=True
        )
    else:
        st.warning("Nenhum dado encontrado.")

    st.divider()

# ------------> Performance de SLA por Operador

    st.subheader("Performance de SLA por Operador")

        # state_session_user = "Tawany Batista"

    if 'Nome Completo do Operador' in df.columns and 'SLA de Solução Expirado' in df.columns:
            # 1. Preparação dos dados
            # Ajuste: Verificando se o valor é "Expirado" ou "Sim" conforme seus códigos anteriores
            df['SLA_Aux'] = df['SLA de Solução Expirado'] == "Expirado"
            # df[df['Nome Completo do Operador'] == state_session_user]
            
            sla_grupo = df.groupby('Nome Completo do Operador').agg(
                Backlog=('SLA_Aux', 'count'),
                Expirados=('SLA_Aux', 'sum')
            ).reset_index()
            
            # Calculando o percentual
            sla_grupo['% SLA Perdido'] = (sla_grupo['Expirados'] / sla_grupo['Backlog'] * 100).round(2)
            sla_grupo = sla_grupo.sort_values(by='% SLA Perdido', ascending=False)

            # 2. Função para aplicar cor vermelha se menor que 80%
            def colorir_abaixo_meta(valor):
                color = 'red' if valor > 30 else 'green'
                return f'color: {color}' if valor > 30 else f'color: {color}'

            # 3. Aplicação da Estilização
            df_estilizado = (
                sla_grupo.style
                # Aplicando a cor baseada no valor da célula na coluna específica
                .map(colorir_abaixo_meta, subset=['% SLA Perdido']) 
                .format("{:.2f}%", subset=['% SLA Perdido'])
                .set_properties(**{
                    'text-align': 'center',
                    'border': '0.1px solid #ededed'
                })
                .set_table_styles([
                    dict(selector='th', props=[('text-align', 'center')])
                ])
            )

            # 4. Exibição no Streamlit
            st.dataframe(df_estilizado, use_container_width=True, hide_index=True)

    else:
            st.error("Colunas necessárias não encontradas.")

# --- FLUXO DE NAVEGAÇÃO ---
if st.session_state.autenticado:
    loading()
else:
    tela_login()




