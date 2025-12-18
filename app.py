# app.py
import streamlit as st
import pandas as pd
from collections import Counter
import plotly.express as px
from io import BytesIO
import ast

# =============================
# 1) Configurações da página
# =============================
st.set_page_config(
    page_title='Dashboard Viagens',
    layout='wide'
)

# =============================
# 2) Função para carregar dados
# =============================
@st.cache_data
def carregar_dados(caminho):
    return pd.read_excel(caminho)

# =============================
# 3) Explode viagens em destinos
# =============================
def explode_viagens(df, coluna='Viagens'):
    df['Viagens_Lista'] = df[coluna].fillna('').astype(str).str.split(',')

    destinos = []
    for lista in df['Viagens_Lista']:
        for v in lista:
            v_limpo = v.strip().title()
            if v_limpo:
                destinos.append(v_limpo)

    return pd.Series(destinos, name='Destino')

# =============================
# 4) Ranking de clientes
# =============================
def ranking_clientes(df, coluna_viagens='Viagens'):
    dados = []

    for _, row in df.iterrows():
        nome = row['Nome']
        viagens = row[coluna_viagens]

        if pd.isna(viagens):
            qtd = 0
        else:
            lista = [v.strip() for v in str(viagens).split(',') if v.strip()]
            qtd = len(lista)

        dados.append({
            'Cliente': nome,
            'Quantidade de Destinos': qtd
        })

    df_clientes = pd.DataFrame(dados)
    df_clientes = df_clientes.sort_values(
        by='Quantidade de Destinos',
        ascending=False
    )

    return df_clientes

# =============================
# 5) Filtrar aposentados
# =============================
def filtrar_aposentados(df):
    return df[df['Profissão'].isin(['Aposentado', 'Aposentada'])]

# =============================
# 6) Input do arquivo
# =============================
caminho = st.text_input(
    'Caminho do arquivo Excel',
    value=r'C:\Users\Joana\Documents\Mananciais 2025\Planilha Clientes 2025.xlsx'
)

if not caminho:
    st.warning('Informe o caminho do arquivo Excel')
    st.stop()

with st.spinner('Carregando dados...'):
    df = carregar_dados(caminho)

# =============================
# 7) Sidebar
# =============================
with st.sidebar:
    st.header('📊 Visualizações')

    modo = st.radio(
        'Escolha o que deseja ver:',
        [
            'Destinos mais viajados',
            'Clientes com mais destinos',
            'Aposentados'
        ]
    )

    st.divider()
    st.header('🔎 Filtros')

    top_n = st.slider('Mostrar top', 5, 50, 10)
    buscar = st.text_input('Buscar destino')

# =============================
# 8) Processamentos
# =============================
serie_destinos = explode_viagens(df)
contagem = serie_destinos.value_counts()

if buscar:
    contagem = contagem[
        contagem.index.str.contains(buscar.strip().title())
    ]

# =============================
# 9) Título
# =============================
st.title('Dashboard de Insights — Viagens')
st.markdown('Visualize destinos mais populares, clientes mais viajados e aposentados.')

# =============================
# 10) Conteúdo central
# =============================
if modo == 'Destinos mais viajados':
    top = contagem.head(top_n)

    df_plot = top[::-1].reset_index()
    df_plot.columns = ['Destino', 'Contagem']

    fig = px.bar(
        df_plot,
        x='Contagem',
        y='Destino',
        orientation='h'
    )

    fig.update_layout(height=500, margin=dict(l=120))
    st.plotly_chart(fig, use_container_width=True)

    df_ranking = top.reset_index()
    df_ranking.columns = ['Destino', 'Contagem']
    st.dataframe(df_ranking, use_container_width=True)

elif modo == 'Clientes com mais destinos':
    st.subheader('🏆 Ranking de Clientes — Mais Destinos Visitados')

    df_clientes = ranking_clientes(df)

    st.dataframe(df_clientes, use_container_width=True)

    fig_clientes = px.bar(
        df_clientes.head(top_n)[::-1],
        x='Quantidade de Destinos',
        y='Cliente',
        orientation='h'
    )

    fig_clientes.update_layout(height=500, margin=dict(l=120))
    st.plotly_chart(fig_clientes, use_container_width=True)

else:  # APOSENTADOS
    st.subheader('👵👴 Lista de Clientes Aposentados')

    df_aposentados = filtrar_aposentados(df)

    if df_aposentados.empty:
        st.info('Nenhum aposentado encontrado.')
    else:
        st.metric(
            label='Total de Aposentados',
            value=len(df_aposentados)
        )

        st.dataframe(
            df_aposentados[['Nome', 'Profissão', 'Celular']],
            use_container_width=True
        )

# =============================
# 11) Exportação
# =============================
def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='dados')
    return output.getvalue()

if modo == 'Clientes com mais destinos':
    export_df = df_clientes
elif modo == 'Aposentados':
    export_df = df_aposentados
else:
    export_df = df_ranking

st.download_button(
    '📥 Baixar Excel',
    data=to_excel_bytes(export_df),
    file_name='dados_exportados.xlsx'
)

# =============================
# 12) Aniversariantes do mês
# =============================
with st.sidebar.expander("🎉 Aniversariantes do Mês", expanded=False):
    meses_pt = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
        5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
        9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }

    mes_atual_num = pd.Timestamp.today().month
    mes_atual_nome = meses_pt[mes_atual_num]

    tabela = r'C:\Users\Joana\Documents\Mananciais 2025\Tabela de Infos 2025.xlsx'
    df_infos = pd.read_excel(tabela)

    linha_mes = df_infos[
        df_infos['Mês'].str.lower() == mes_atual_nome.lower()
    ]

    if linha_mes.empty:
        st.info(f"Nenhum registro encontrado para {mes_atual_nome}.")
    else:
        lista_raw = linha_mes.iloc[0]['lista_aniversariantes']
        lista_niver = ast.literal_eval(lista_raw) if isinstance(lista_raw, str) else lista_raw

        df_aniver = pd.DataFrame(lista_niver)
        df_aniver['Aniversario'] = pd.to_datetime(
            df_aniver['Aniversario'], errors='coerce'
        )

        celulares = []
        for nome in df_aniver['Nome']:
            linha = df[df['Nome'] == nome]
            celulares.append(
                linha['Celular'].values[0]
                if not linha.empty else 'Não encontrado'
            )

        df_aniver['Celular'] = celulares
        df_aniver = df_aniver[
            df_aniver['Aniversario'].dt.month == mes_atual_num
        ]

        if df_aniver.empty:
            st.info(f"Nenhum aniversariante em {mes_atual_nome}.")
        else:
            df_aniver['Dia'] = df_aniver['Aniversario'].dt.day
            df_aniver = df_aniver.sort_values(by='Dia')
            df_aniver['Data Formatada'] = df_aniver['Aniversario'].dt.strftime('%d/%m')

            for _, row in df_aniver.iterrows():
                st.write(
                    f"• {row['Nome']} — {row['Data Formatada']} — {row['Celular']}"
                )

            st.dataframe(
                df_aniver[['Nome', 'Data Formatada', 'Celular']],
                use_container_width=True
            )
