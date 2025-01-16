import streamlit as st
import pandas as pd
from itertools import combinations
from io import BytesIO
from openpyxl import Workbook
from datetime import datetime, timedelta
import matplotlib.pyplot as plt

# Configuração inicial do Streamlit
st.set_page_config(
    page_title="Ferramenta de Câmbio",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Função para validar o arquivo
def validar_arquivo(file):
    if not file.name.endswith((".xlsx", ".xls", ".csv")):
        raise ValueError("O arquivo deve estar nos formatos .xlsx, .xls ou .csv.")

# Função para carregar a base de dados
def carregar_base(file):
    validar_arquivo(file)
    try:
        if file.name.endswith(".xlsx") or file.name.endswith(".xls"):
            base = pd.ExcelFile(file, engine="openpyxl")
            if "Base_Unificada" not in base.sheet_names:
                raise ValueError("A aba 'Base_Unificada' não foi encontrada no arquivo.")
            return base.parse(sheet_name="Base_Unificada")
        elif file.name.endswith(".csv"):
            return pd.read_csv(file)
        else:
            raise ValueError("Formato de arquivo não suportado. Use .xlsx, .xls ou .csv.")
    except Exception as e:
        raise ValueError(f"Erro ao carregar o arquivo: {e}")

# Função para listar empresas
def listar_empresas(base):
    return base["Empresa"].dropna().unique()

# Função para verificar processos e dias em aberto
def verificar_processos_dias_aberto(base):
    hoje = datetime.now()
    base["Data"] = pd.to_datetime(base["Data"], errors="coerce")
    base["Dias_Em_Aberto"] = (hoje - base["Data"]).dt.days
    base["Dias_Em_Aberto"] = base["Dias_Em_Aberto"].clip(lower=0)  # Evita valores negativos
    return base

# Função para salvar combinações em um arquivo Excel
def salvar_combinacao_excel(combinacoes):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Sumário de Combinações"

    ws.append(["Processos", "Datas", "Total"])

    for combinacao in combinacoes:
        ws.append([', '.join(map(str, combinacao['Processos'])), 
                   ', '.join(combinacao['Datas']), combinacao['Total']])

    wb.save(output)
    output.seek(0)
    return output

# Função para gerar gráfico de processos por intervalo de dias
def gerar_grafico_dias_aberto(base):
    intervalos = [0, 30, 60, 90, 120, 150, 180, float("inf")]
    labels = ["0-30", "30-60", "60-90", "90-120", "120-150", "150-180", ">180"]
    
    base["Intervalo_Dias"] = pd.cut(
        base["Dias_Em_Aberto"], bins=intervalos, labels=labels, right=False
    )

    processos_por_intervalo = base.groupby(["Empresa", "Intervalo_Dias"]).size().unstack(fill_value=0)

    # Criar o gráfico
    plt.figure(figsize=(12, 8))
    processos_por_intervalo.plot(kind="bar", stacked=True, figsize=(12, 6))
    plt.title("Total de Processos por Intervalo de Dias em Aberto e Empresa")
    plt.xlabel("Empresa")
    plt.ylabel("Total de Processos")
    plt.legend(title="Intervalos de Dias em Aberto")
    plt.xticks(rotation=45)
    plt.tight_layout()
    st.pyplot(plt)

# Função para gerar gráfico de total de processos por empresa
def gerar_grafico_total_processos(base):
    total_por_empresa = base["Empresa"].value_counts()

    # Criar o gráfico
    plt.figure(figsize=(10, 6))
    total_por_empresa.plot(kind="bar", color="skyblue")
    plt.title("Total de Processos por Empresa")
    plt.xlabel("Empresa")
    plt.ylabel("Total de Processos")
    plt.xticks(rotation=45)
    plt.tight_layout()
    st.pyplot(plt)

# Função para encontrar combinações
def encontrar_combinacoes(base, empresa, valor_alvo, margem_fixa=1500, max_combinacoes=5):
    dados_filtrados = base[base["Empresa"] == empresa]
    valores_processos = dados_filtrados[["Processo", "Valor", "Data"]].values

    margem_min = valor_alvo - margem_fixa
    margem_max = valor_alvo + margem_fixa

    combinacoes_possiveis = []
    valor_exato_encontrado = False

    for r in range(1, len(valores_processos) + 1):
        for combinacao in combinations(valores_processos, r):
            soma = sum([item[1] for item in combinacao])
            if soma == valor_alvo:
                valor_exato_encontrado = True
            if margem_min <= soma <= margem_max:
                combinacoes_possiveis.append(combinacao)
                if len(combinacoes_possiveis) >= max_combinacoes:
                    return {"combinacoes": combinacoes_possiveis, "valor_exato": valor_exato_encontrado}

    return {"combinacoes": combinacoes_possiveis, "valor_exato": valor_exato_encontrado}

# Função para formatar resultados de combinações
def formatar_resultados(resultado):
    combinacoes = resultado["combinacoes"]
    valor_exato = resultado["valor_exato"]

    if combinacoes:
        resultado_formatado = []
        for comb in combinacoes:
            processos = [item[0] for item in comb]
            total = sum([item[1] for item in comb])
            datas = [item[2].strftime('%Y-%m-%d') if pd.notna(item[2]) and hasattr(item[2], 'strftime') else 'Data Inválida' for item in comb]
            resultado_formatado.append({'Processos': processos, 'Datas': datas, 'Total': total})

        return pd.DataFrame(resultado_formatado), valor_exato
    return None, valor_exato

# Função principal para exibição de abas no Streamlit
def exibir_abas():
    st.title("Ferramenta de Fechamento de Câmbio")

    # Sistema de Login
    if "autenticado" not in st.session_state:
        st.session_state.autenticado = False

    if not st.session_state.autenticado:
        usuario = st.text_input("Usuário:")
        senha = st.text_input("Senha:", type="password")

        if st.button("Login"):
            if usuario == "icaro" and senha == "gocomexx25":
                st.session_state.autenticado = True
                st.success("Login realizado com sucesso!")
            else:
                st.error("Usuário ou senha incorretos.")
        return

    # Upload do arquivo
    file = st.sidebar.file_uploader("Faça upload do arquivo (.xlsx, .xls, .csv)", type=["xlsx", "xls", "csv"])

    if not file:
        st.warning("Por favor, carregue um arquivo para começar.")
        return

    base = carregar_base(file)
    base = verificar_processos_dias_aberto(base)

    # Estatísticas gerais
    st.sidebar.subheader("Resumo Geral")
    st.sidebar.metric("Total de Empresas", len(listar_empresas(base)))
    st.sidebar.metric("Total de Processos", len(base))
    st.sidebar.metric("Dias Médios em Aberto", int(base["Dias_Em_Aberto"].mean()))

    abas = ["Operações", "Fechamento de Câmbio", "Gráficos", "Notificações", "Sumário de Câmbio"]
    escolha = st.sidebar.radio("Navegar", abas)

    if escolha == "Operações":
        st.header("Operações")
        st.dataframe(base)

    elif escolha == "Gráficos":
        st.header("Gráficos de Processos por Intervalo de Dias")
        gerar_grafico_dias_aberto(base)

        st.header("Gráfico de Total de Processos por Empresa")
        gerar_grafico_total_processos(base)

    elif escolha == "Fechamento de Câmbio":
        st.header("Fechamento de Câmbio")

        empresas = listar_empresas(base)
        empresa_selecionada = st.selectbox("Selecione uma empresa:", empresas)
        valor_alvo = st.number_input("Digite o valor alvo para fechamento:", min_value=0.0, step=0.01)

        if st.button("Buscar Combinações"):
            with st.spinner("Buscando combinações..."):
                resultado = encontrar_combinacoes(base, empresa_selecionada, valor_alvo)

            resultado_df, valor_exato = formatar_resultados(resultado)

            if resultado_df is not None and not resultado_df.empty:
                st.session_state.resultado_df = resultado_df
                st.success("Combinações encontradas com sucesso!")
            else:
                st.warning("Nenhuma combinação encontrada.")

        if "resultado_df" in st.session_state and not st.session_state.resultado_df.empty:
            st.dataframe(st.session_state.resultado_df)

            if "combinacoes_selecionadas" not in st.session_state:
                st.session_state.combinacoes_selecionadas = []

            combinacao_selecionada = st.multiselect("Selecione uma ou mais combinações:", 
                                                     st.session_state.resultado_df.index.tolist(), 
                                                     format_func=lambda x: f"Combinação {x+1}")

            if st.button("Adicionar ao Sumário"):
                for idx in combinacao_selecionada:
                    if st.session_state.resultado_df.iloc[idx].to_dict() not in st.session_state.combinacoes_selecionadas:
                        st.session_state.combinacoes_selecionadas.append(st.session_state.resultado_df.iloc[idx].to_dict())
                        st.success(f"Combinação {idx + 1} adicionada ao sumário!")

    elif escolha == "Sumário de Câmbio":
        st.header("Sumário de Câmbio")

        if "combinacoes_selecionadas" in st.session_state and st.session_state.combinacoes_selecionadas:
            st.subheader("Tabela de Fechamentos Selecionados")
            combinacoes_df = pd.DataFrame(st.session_state.combinacoes_selecionadas)
            st.dataframe(combinacoes_df)

            output = salvar_combinacao_excel(st.session_state.combinacoes_selecionadas)
            file_name = f"sumario_cambio_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
            st.download_button("Baixar Sumário de Câmbio", output, file_name)
        else:
            st.info("Nenhuma combinação foi selecionada ainda.")

    elif escolha == "Notificações":
        st.header("Notificações de Processos")
        processos_pendentes = base[base["Dias_Em_Aberto"] > 180]

        if not processos_pendentes.empty:
            st.warning("Atenção! Existem processos que ultrapassaram 180 dias em aberto:")
            st.dataframe(processos_pendentes)
        else:
            st.info("Todos os processos estão dentro do prazo.")

# Executa o aplicativo
if __name__ == "__main__":
    exibir_abas()
