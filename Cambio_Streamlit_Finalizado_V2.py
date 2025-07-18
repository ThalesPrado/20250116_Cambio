import streamlit as st
import pandas as pd
from itertools import combinations
from io import BytesIO
from openpyxl import Workbook
from datetime import datetime
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
        if file.name.endswith((".xlsx", ".xls")):
            base = pd.read_excel(file, sheet_name=None, engine="openpyxl")
            sheet = st.selectbox("Selecione a aba:", list(base.keys()))
            df = base[sheet]
        elif file.name.endswith(".csv"):
            df = pd.read_csv(file)
        else:
            raise ValueError("Formato de arquivo não suportado. Use .xlsx, .xls ou .csv.")
        
        if "Cambio_Fechado" in df.columns:
            df["Cambio_Fechado"] = df["Cambio_Fechado"].apply(lambda x: True if str(x).strip().lower() == "feito" else False)
        else:
            df["Cambio_Fechado"] = False
        
        return df
    except Exception as e:
        raise ValueError(f"Erro ao carregar o arquivo: {e}")

# Função para listar empresas
def listar_empresas(base):
    return base["Empresa"].dropna().unique()

# Função para listar exportadores
def listar_exportadores(base):
    return base["Exportador"].dropna().unique()

# Função para verificar processos e dias em aberto
def verificar_processos_dias_aberto(base):
    hoje = datetime.now()
    base["Data"] = pd.to_datetime(base["Data"], errors="coerce")
    base["Dias_Em_Aberto"] = (hoje - base["Data"]).dt.days.clip(lower=0)
    return base

# Função para salvar combinações em Excel
def salvar_combinacao_excel(combinacoes):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Sumário de Combinações"

    ws.append(["Empresa", "Exportador", "Processos", "Datas", "Total"])

    for combinacao in combinacoes:
        ws.append([
            combinacao['Empresa'],
            combinacao['Exportador'],
            ', '.join(map(str, combinacao['Processos'])),
            ', '.join(combinacao['Datas']),
            combinacao['Total']
        ])

    wb.save(output)
    output.seek(0)
    return output

# Função para salvar a base atualizada (com flag de câmbio fechado)
def salvar_base_atualizada(base):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        base["Cambio_Fechado"] = base["Cambio_Fechado"].apply(lambda x: "Feito" if x else "Não feito")
        base.to_excel(writer, index=False, sheet_name="Base_Atualizada")
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

    plt.figure(figsize=(12, 6))
    processos_por_intervalo.plot(kind="bar", stacked=True, figsize=(12, 6))
    plt.title("Total de Processos por Intervalo de Dias em Aberto e Empresa")
    plt.xlabel("Empresa")
    plt.ylabel("Total de Processos")
    plt.xticks(rotation=45)
    plt.tight_layout()
    st.pyplot(plt)

# Função para gerar gráfico de total de processos por empresa
def gerar_grafico_total_processos(base):
    total_por_empresa = base["Empresa"].value_counts()

    plt.figure(figsize=(10, 6))
    total_por_empresa.plot(kind="bar", color="skyblue")
    plt.title("Total de Processos por Empresa")
    plt.xlabel("Empresa")
    plt.ylabel("Total de Processos")
    plt.xticks(rotation=45)
    plt.tight_layout()
    st.pyplot(plt)

# Função para encontrar combinações (ignora processos fechados)
def encontrar_combinacoes(base, empresa, exportador, valor_alvo, margem_fixa=1500, max_combinacoes=5):
    dados_filtrados = base[
        (base["Empresa"] == empresa) &
        (base["Exportador"] == exportador) &
        (base["Cambio_Fechado"] == False)
    ]
    valores_processos = dados_filtrados[["Processo", "Valor", "Data"]].values

    margem_min = valor_alvo - margem_fixa
    margem_max = valor_alvo + margem_fixa

    combinacoes_possiveis = []

    for r in range(1, len(valores_processos) + 1):
        for combinacao in combinations(valores_processos, r):
            soma = sum([item[1] for item in combinacao])
            if margem_min <= soma <= margem_max:
                combinacoes_possiveis.append({
                    "Empresa": empresa,
                    "Exportador": exportador,
                    "Processos": [item[0] for item in combinacao],
                    "Datas": [item[2].strftime('%Y-%m-%d') if pd.notna(item[2]) else 'Data Inválida' for item in combinacao],
                    "Total": soma
                })
                if len(combinacoes_possiveis) >= max_combinacoes:
                    return combinacoes_possiveis
    return combinacoes_possiveis

# Função principal para exibição de abas no Streamlit
def exibir_abas():
    st.title("Ferramenta de Fechamento de Câmbio")

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

    file = st.sidebar.file_uploader("Faça upload do arquivo (.xlsx, .xls, .csv)", type=["xlsx", "xls", "csv"])
    if not file:
        st.warning("Por favor, carregue um arquivo para começar.")
        return

    if "base" not in st.session_state:
        base = carregar_base(file)
        base = verificar_processos_dias_aberto(base)
        st.session_state.base = base

    base = st.session_state.base

    st.sidebar.subheader("Resumo Geral")
    st.sidebar.metric("Total de Empresas", len(listar_empresas(base)))
    st.sidebar.metric("Total de Processos", len(base))
    st.sidebar.metric("Dias Médios em Aberto", int(base["Dias_Em_Aberto"].mean()))

    abas = ["Operações", "Fechamento de Câmbio", "Gráficos", "Notificações", "Sumário de Câmbio"]
    escolha = st.sidebar.radio("Navegar", abas)

    if escolha == "Operações":
        st.header("Operações")
        st.dataframe(base.drop(columns=["Cambio_Fechado"], errors='ignore'))

    elif escolha == "Gráficos":
        st.header("Gráficos de Processos")
        gerar_grafico_dias_aberto(base)
        gerar_grafico_total_processos(base)

    elif escolha == "Fechamento de Câmbio":
        st.header("Fechamento de Câmbio")

        empresas = listar_empresas(base)
        exportadores = listar_exportadores(base)

        empresas_opcoes = ["Todas"] + list(empresas)
        exportadores_opcoes = ["Todos"] + list(exportadores)
        status_opcoes = ["Feito", "Não feito"]

        empresas_selecionadas = st.multiselect("Selecione empresa(s):", empresas_opcoes, default="Todas")
        exportadores_selecionados = st.multiselect("Selecione exportador(es):", exportadores_opcoes, default="Todos")
        status_selecionado = st.multiselect("Selecione o status dos processos:", status_opcoes, default="Não feito")

        if "Todas" in empresas_selecionadas:
            empresas_filtradas = empresas
        else:
            empresas_filtradas = empresas_selecionadas

        if "Todos" in exportadores_selecionados:
            exportadores_filtrados = exportadores
        else:
            exportadores_filtrados = exportadores_selecionados

        status_bool_map = {"Feito": True, "Não feito": False}
        status_filtrados = [status_bool_map[s] for s in status_selecionado]

        base_filtrada = base[
            (base["Empresa"].isin(empresas_filtradas)) &
            (base["Exportador"].isin(exportadores_filtrados)) &
            (base["Cambio_Fechado"].isin(status_filtrados))
        ]

        valor_alvo = st.number_input("Digite o valor alvo para fechamento:", min_value=0.0, step=0.01)

        if st.button("Buscar Combinações"):
            st.session_state.resultados = []
            for empresa in empresas_filtradas:
                for exportador in exportadores_filtrados:
                    combinacoes = encontrar_combinacoes(base_filtrada, empresa, exportador, valor_alvo)
                    st.session_state.resultados.extend(combinacoes)

            if st.session_state.resultados:
                st.success(f"{len(st.session_state.resultados)} combinações encontradas.")
                st.session_state.resultado_df = pd.DataFrame(st.session_state.resultados)
                st.dataframe(
                    st.session_state.resultado_df.style.applymap(
                        lambda val: "background-color: #ffcccc" if val in base[base["Cambio_Fechado"]]["Processo"].tolist() else ""
                    )
                )
            else:
                st.warning("Nenhuma combinação encontrada.")

        if "resultado_df" in st.session_state:
            selecionadas = st.multiselect(
                "Selecione as combinações para dar baixa:",
                st.session_state.resultado_df.index.tolist(),
                format_func=lambda x: f"Combinação {x + 1}"
            )

            if st.button("Dar baixa"):
                processos_a_atualizar = []
                for idx in selecionadas:
                    processos_a_atualizar.extend(st.session_state.resultado_df.iloc[idx]["Processos"])

                st.session_state.base.loc[
                    st.session_state.base["Processo"].isin(processos_a_atualizar),
                    "Cambio_Fechado"
                ] = True

                st.success("Processos marcados como fechados. Eles não aparecerão mais em novas combinações.")

                base_atualizada = salvar_base_atualizada(st.session_state.base)
                st.download_button(
                    "Baixar Base Atualizada",
                    base_atualizada,
                    file_name=f"base_atualizada_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
                )

    elif escolha == "Sumário de Câmbio":
        st.header("Sumário de Câmbio")
        if "resultado_df" in st.session_state:
            st.subheader("Combinações Encontradas")
            st.dataframe(
                st.session_state.resultado_df.style.applymap(
                    lambda val: "background-color: #ffcccc" if val in base[base["Cambio_Fechado"]]["Processo"].tolist() else ""
                )
            )

        st.subheader("Processos Já Fechados")
        processos_fechados = base[base["Cambio_Fechado"] == True]
        if not processos_fechados.empty:
            st.dataframe(processos_fechados)
        else:
            st.info("Nenhum processo fechado até o momento.")

    elif escolha == "Notificações":
        st.header("Notificações")
        st.subheader("⚠️ Processos com mais de 180 dias em aberto")
        processos_pendentes = base[(base["Dias_Em_Aberto"] > 180) & (base["Cambio_Fechado"] == False)]
        if not processos_pendentes.empty:
            st.warning("Atenção! Existem processos com mais de 180 dias em aberto:")
            st.dataframe(processos_pendentes)
        else:
            st.info("Nenhum processo acima de 180 dias em aberto.")

        st.subheader("📦 Processos Já Fechados")
        processos_fechados = base[base["Cambio_Fechado"] == True]
        if not processos_fechados.empty:
            st.dataframe(processos_fechados)
        else:
            st.info("Nenhum processo fechado até o momento.")

# Executa o aplicativo
if __name__ == "__main__":
    exibir_abas()

