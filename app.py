import streamlit as st
import pandas as pd
import pathlib
import plotly.express as px

# Configuração da página
st.set_page_config(page_title="Gestão de Diárias e Passagens", layout="wide")

def load_and_clean_data():
    # 1. Scan da pasta local por arquivos .xlsx
    current_path = pathlib.Path(".")
    files = list(current_path.glob("*.xlsx"))
    
    if not files:
        st.error("Nenhum arquivo .xlsx encontrado na pasta local.")
        return None

    all_data = []
    
    for file in files:
        # Pula as 6 primeiras linhas conforme a estrutura do seu arquivo anexado
        # O cabeçalho real está na linha 7 (index 6)
        try:
            df_temp = pd.read_excel(file, skiprows=6)
            
            # Limpeza de colunas vazias (comum em arquivos exportados de sistemas antigos)
            df_temp = df_temp.dropna(how='all', axis=1)
            
            # Padronização de nomes de colunas (baseado no seu CSV)
            # Mapeamos as colunas baseado na posição caso os nomes variem levemente
            df_temp.columns = [
                'Processo', 'Ano', 'Data', 'Historico', 'Cargo', 
                'Lotacao', 'Servidor', 'Descricao', 'Valor', 'Fornecedor'
            ] + list(df_temp.columns[10:]) # Mantém o resto se houver
            
            all_data.append(df_temp)
        except Exception as e:
            st.warning(f"Erro ao ler {file.name}: {e}")

    if not all_data:
        return None
        
    df = pd.concat(all_data, ignore_index=True)

    # --- TRATAMENTO DE DADOS ---
    
    # 2. Limpeza do Valor (remover NaN e garantir que seja numérico)
    df['Valor'] = pd.to_numeric(df['Valor'], errors='coerce')
    df = df.dropna(subset=['Valor'])

    # 3. Tratamento de Datas complexas (Ex: "08/02/2017 e 09/02/2017")
    # Pegamos apenas a primeira data mencionada para fins de cronologia
    df['Data_Limpa'] = df['Data'].astype(str).str.extract(r'(\d{2}/\d{2}/\d{4}|\d{4}-\d{2}-\d{2})')[0]
    df['Data_Limpa'] = pd.to_datetime(df['Data_Limpa'], errors='coerce')
    
    return df

# Interface Streamlit
st.title("🏛️ Dashboard de Transparência - Diárias e Passagens")
st.markdown("Análise automatizada de arquivos de prestação de contas.")

df = load_and_clean_data()

if df is not None:
    # Sidebar Filtros
    st.sidebar.header("Filtros de Pesquisa")
    secretarias = st.sidebar.multiselect("Filtrar por Secretaria (Lotação):", options=df['Lotacao'].unique())
    
    if secretarias:
        df = df[df['Lotacao'].isin(secretarias)]

    # --- KPIs ---
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Gasto", f"R$ {df['Valor'].sum():,.2f}")
    with col2:
        st.metric("Nº de Processos", len(df['Processo'].unique()))
    with col3:
        st.metric("Ticket Médio", f"R$ {df['Valor'].mean():,.2f}")

    st.divider()

    # --- GRÁFICOS ---
    c1, c2 = st.columns(2)

    with c1:
        st.subheader("Gastos por Secretaria")
        fig_lotacao = px.bar(
            df.groupby("Lotacao")["Valor"].sum().reset_index().sort_values("Valor", ascending=False),
            x="Lotacao", y="Valor", color="Lotacao",
            labels={'Valor': 'Total (R$)', 'Lotacao': 'Secretaria'}
        )
        st.plotly_chart(fig_lotacao, use_container_width=True)

    with c2:
        st.subheader("Evolução Mensal")
        # Agrupando por mês
        df_mensal = df.set_index('Data_Limpa').resample('M')['Valor'].sum().reset_index()
        fig_tempo = px.line(df_mensal, x='Data_Limpa', y='Valor', markers=True)
        st.plotly_chart(fig_tempo, use_container_width=True)

    # --- RANKING ---
    st.subheader("Top 10 Servidores por Volume de Gastos")
    ranking = df.groupby(["Servidor", "Cargo"])["Valor"].agg(['sum', 'count']).rename(columns={'sum': 'Total Gasto', 'count': 'Qtd Viagens'})
    st.dataframe(ranking.sort_values("Total Gasto", ascending=False).head(10), use_container_width=True)

    # --- BUSCA BRUTA ---
    with st.expander("🔍 Visualizar Dados Brutos"):
        st.dataframe(df)

else:
    st.info("Aguardando arquivos válidos na pasta do projeto.")