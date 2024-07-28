import streamlit as st
import pandas as pd
import openpyxl as xl
import plotly.express as px
import plotly.graph_objects as go

# Projeto de Análise de Dados em Python

# Carregando meu arquivo de Excel
try:
    arquivo = pd.read_excel('mov_fin.xlsx', engine='openpyxl')
except Exception as e:
    st.error(f"Erro ao carregar o arquivo: {e}")

# Carregando meu Streamlit
st.header("Apresentação sobre Ciência de Dados")
st.sidebar.image('LOGO.png')
st.sidebar.text("""Menu de Navegação do Projeto""")
st.sidebar.markdown("""
Condominio Village das Fontes, 770 - Benedito Bentes
Maceió/Al
Tel.(82) 98863-9394
dateanalytics@outlook.com
""")

rotas = "https://app.powerbi.com/view?r=eyJrIjoiZGYyYzcwMDYtYzZmZC00YjlhLWJjYzQtYmE4MmMyOTc5MTdmIiwidCI6ImY1OGYxNjE2LWZkYWEtNGRhZS1hN2ZjLTc1ODI5YzkxOWE2YSJ9"

# Adicionando CSS para diminuir o tamanho das métricas
st.markdown(
    """
    <style>
    .small-font {
        font-size:16px !important;
    }
    </style>
    """, unsafe_allow_html=True
)

cx_selecao = st.sidebar.selectbox("Selecione uma opção:", [
    'Home', 'Apresentação', 'Tabela', 'Gráficos', 'Dashboard Roteirização', 'Tratar Planilha'
])

if cx_selecao == 'Tabela':
    st.dataframe(arquivo)
    
if cx_selecao == 'Home':
    st.header("Projeto de Análise de Dados em Python")
    st.markdown("Este projeto visa demonstrar um projeto de análise de dados referente a uma tabela de movimento financeiro entre receitas e despesas. Observação: Estes dados são fictícios, qualquer necessidade de conhecer um pouco mais sobre a lógica de programação basta entrar em contato.")
    st.image('dados.png')

if cx_selecao == 'Apresentação':
    st.header("Conheça um pouco da minha vivência profissional")
    st.image('foto.jpg')
    st.markdown(""" 
    Me Chamo Williams Rodrigues, sou pós graduando em Ciência de Big Data Analytics e bacharel em administração de empresas,
    ao longo da minha trajetória profissional sempre gostei da área de tecnologia, já iniciei alguns cursos na área porém por questões do destino,
    precisei parar, tudo mudou quando decidi estudar ADMINISTRAÇÃO DE EMPRESAS, e aí foi um divisor de águas na minha trajetória profissional,
    conheci diversos profissionais gabaritados e renomados em diversas áreas, tive o prazer e felicidade de presidir ativamente como diretor de uma delegação para organização de eventos acadêmicos
    que no qual organizamos excursões e palestras dentro do próprio conselho de Administração CRA/AL, até para fora do estado organizamos eventos, e isso me motivava a cada dia.

    Foi aí que tive a brilhante ideia: Organizar, apresentar aulas de Microsoft Excel para graduandos e alunos do ensino médio totalmente de graça, para ajudar no aprendizado e aprimoramento para o mercado de trabalho.
    O projeto foi denominado EXCEL EMPRESARIAL, com aulas voltadas para a prática nas empresas, eu trazia uma forma diferente de treinamento, com uma linguagem mais próxima do aluno e de fácil entendimento.
    Ainda durante a vida acadêmica fui: Palestrante em sustentabilidade, Monitor, Diretor de Delegação e Consultor em Microsoft Excel.

    Com a base acadêmica consolidada decidi recomeçar a estudar fazer a pós-graduação, o que me abriu mais ainda os olhos para a área tecnológica, hoje já desenvolvi diversos projetos em 
    Power BI, em Python, em VBA - Access e diversos materiais que foram comercializados para pessoas e empresas.
    Atuando como consultor ou instrutor nesta área de dados e apresentações de indicadores.
    Atualmente estou trabalhando em diversos projetos de análise de dados, entre eles um específico na área de logística para criar, implementar e desenvolver dados de indicadores referentes aos clientes,

    A ciência de Dados é uma área que não há limites para o conhecimento nela se aprende: automatizar, melhorar os processos a fim de reduzir custos para uma empresa, sempre visando a eficiência, a qualidade e a transparência.
    Quer conhecer um pouco mais sobre ciência de dados acesse: https://www.ibm.com/br-pt/topics/data-science
    """)
    # Saiba mais sobre nossos trabalhos:
    st.header("Saiba mais sobre nossos trabalhos:")
    st.button("Whatsapp", 'https://wa.me/5582988639394')
    st.button("Portfólio", 'https://wrportifolio.streamlit.app')
    st.button("LinkedIn", 'https://www.linkedin.com/in/williams-rodrigues-9b350a6a/')
    st.button("Instagram", 'https://www.instagram.com/williams_rvs85')
   

elif cx_selecao == 'Dashboard Roteirização':
    st.title('Mapeamento e Análise de Dados do Planejamento de Roteirização')
    st.markdown(f'<iframe width="800" height="600" src="{rotas}" frameborder="0" allowfullscreen></iframe>', unsafe_allow_html=True)   

elif cx_selecao == 'Tratar Planilha':
   
    # Função para carregar e exibir a planilha Excel
    def load_excel(uploaded_file):
        try:
            # Ler todas as planilhas do arquivo Excel
            xls = pd.ExcelFile(uploaded_file, engine='openpyxl')

            # Obter nomes das planilhas
            sheet_names = xls.sheet_names
            st.write(f"Planilhas disponíveis: {sheet_names}")

            # Selecionar planilha
            sheet_name = st.bar_chart("Escolha uma planilha para carregar", sheet_names)

            # Ler a planilha selecionada
            df = pd.read_excel(xls, sheet_name=sheet_name)

            # Remover linhas e colunas vazias
            df.dropna(how='all', inplace=True)  # Remove linhas completamente vazias
            df.dropna(axis=1, how='all', inplace=True)  # Remove colunas completamente vazias

            # Exibir o DataFrame
            st.dataframe(df)
        except Exception as e:
            st.error(f"Erro ao carregar o arquivo: {e}")

    # Carregar a planilha Excel
    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type="xlsx")

    if uploaded_file is not None:
        load_excel(uploaded_file)

elif cx_selecao == 'Gráficos':
    # Calculando a soma da coluna "Pago" para Receitas e Despesas
    receita = arquivo[arquivo['Tipo'] == 'Receita']['Pago'].sum()
    despesa = arquivo[arquivo['Tipo'] == 'Despesa']['Pago'].sum()
    total_pago = arquivo['Pago'].sum()
    media_pago = arquivo['Pago'].median()

    # Calculando o saldo
    saldo = receita - despesa

    # Calculando o lucro líquido
    if receita != 0:
        lucro_liquido = saldo / receita
    else:
        lucro_liquido = 0

    # Exibindo as métricas lado a lado com CSS para diminuir o tamanho
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    
    col1.markdown(f'<div class="small-font">Total Receita<br>R$ {receita:,.2f}</div>', unsafe_allow_html=True)
    col2.markdown(f'<div class="small-font">Total Despesa<br>R$ {despesa:,.2f}</div>', unsafe_allow_html=True)
    col3.markdown(f'<div class="small-font">Saldo<br>R$ {saldo:,.2f}</div>', unsafe_allow_html=True)
    col4.markdown(f'<div class="small-font">Total Pago<br>R$ {total_pago:,.2f}</div>', unsafe_allow_html=True)
    col5.markdown(f'<div class="small-font">Média Pago<br>R$ {media_pago:,.2f}</div>', unsafe_allow_html=True)
    col6.markdown(f'<div class="small-font">Lucro Líquido<br>{lucro_liquido:.2%}</div>', unsafe_allow_html=True)

    # Criando um gráfico de barras referente à coluna "Tipo" e coluna "Pago" com cores personalizadas
    fig_tipo_pago = px.bar(
        arquivo, 
        x='Tipo', 
        y='Pago', 
        title='Gráfico de Barras de Tipo e Pago', 
        labels={'Tipo': 'Tipo', 'Pago': 'Pago'},
        color_discrete_sequence=['#636EFA', '#EF553B']  # Personalize as cores aqui
    )
    st.plotly_chart(fig_tipo_pago)

    # Criando um gráfico de barras referente à coluna "Período" e coluna "Pago" com cores personalizadas
    fig_periodo_pago = px.bar(
        arquivo, 
        x='Período', 
        y='Pago', 
        title='Gráfico de Barras de Período e Pago', 
        labels={'Período': 'Período', 'Pago': 'Pago'},
        color_discrete_sequence=['#00CC96', '#AB63FA']  # Personalize as cores aqui
    )
    st.plotly_chart(fig_periodo_pago)

    # Criando um gráfico de rosca 3D para Receita e Despesa com cores personalizadas
    labels = ['Receita', 'Despesa']
    values = [receita, despesa]

    fig_rosca = go.Figure(data=[go.Pie(
        labels=labels, 
        values=values, 
        hole=.3, 
        title="Receita vs Despesa",
        marker=dict(colors=['#FFD700', '#FF6347'])  # Personalize as cores aqui
    )])
    fig_rosca.update_traces(textinfo='percent+label', marker=dict(line=dict(color='#000000', width=2)))
    st.plotly_chart(fig_rosca)



