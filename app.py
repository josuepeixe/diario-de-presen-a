import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import plotly.express as px
from datetime import date
import io

# ==========================================
# 1. CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(page_title="Gestão - Escola Bíblica", page_icon="📖", layout="wide")

# ==========================================
# 2. CONEXÃO COM GOOGLE SHEETS
# ==========================================
# Substitua pela ID real da sua planilha
ID_DA_PLANILHA = "COLE_O_ID_DA_SUA_PLANILHA_AQUI" 

@st.cache_resource
def conectar_google_sheets():
    try:
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets", 
            "https://www.googleapis.com/auth/drive"
        ]
        credenciais = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], 
            scopes=scopes
        )
        cliente = gspread.authorize(credenciais)
        return cliente.open_by_key(ID_DA_PLANILHA)
    except Exception as e:
        st.error("⚠️ Erro ao conectar no Google Sheets. Verifique o arquivo secrets.toml e o ID da planilha.")
        st.stop()

planilha = conectar_google_sheets()

# Função para garantir que a aba base seja carregada e convertida em Pandas DataFrame
def carregar_diario():
    try:
        aba = planilha.worksheet("Diário")
    except gspread.exceptions.WorksheetNotFound:
        # Cria a aba caso não exista com a estrutura básica na Coluna A
        aba = planilha.add_worksheet(title="Diário", rows="100", cols="50")
        aba.update("A1:A4", [["Data"], ["Lição"], ["Valor da Oferta"], ["Total Presente"]])
        
    dados = aba.get_all_values()
    if not dados:
        # Inicializa a estrutura base se a aba estiver vazia
        dados = [["Data"], ["Lição"], ["Valor da Oferta"], ["Total Presente"]]
        aba.update("A1:A4", dados)
        
    return aba, pd.DataFrame(dados)

# ==========================================
# 3. MENU DE NAVEGAÇÃO LATERAL
# ==========================================
st.sidebar.title("📖 Escola Bíblica")
st.sidebar.markdown("---")
menu = st.sidebar.radio("Navegação", ["Dashboard", "Diário de Aula", "Gestão de Alunos"])

# ==========================================
# 4. MÓDULOS DO SISTEMA
# ==========================================

aba_diario, df_diario = carregar_diario()

# Mapeia a estrutura da Coluna A (Índice 0)
rotulos = df_diario.iloc[:, 0].tolist()
idx_data = rotulos.index("Data") if "Data" in rotulos else 0
idx_licao = rotulos.index("Lição") if "Lição" in rotulos else 1
idx_oferta = rotulos.index("Valor da Oferta") if "Valor da Oferta" in rotulos else 2
idx_total = rotulos.index("Total Presente") if "Total Presente" in rotulos else len(rotulos)-1

# Extrai dinamicamente a lista de alunos (linhas entre "Valor da Oferta" e "Total Presente")
lista_alunos = rotulos[idx_oferta + 1 : idx_total]

# ------------------------------------------
# TELA: DASHBOARD
# ------------------------------------------
if menu == "Dashboard":
    st.title("Painel de Indicadores 📊")
    
    # Verifica se há dados lançados além da coluna de rótulos (Coluna A)
    if df_diario.shape[1] > 1:
        datas = df_diario.iloc[idx_data, 1:].tolist()
        
        # Converte valores financeiros e presenças para numérico
        ofertas = pd.to_numeric(df_diario.iloc[idx_oferta, 1:], errors='coerce').fillna(0)
        presencas = pd.to_numeric(df_diario.iloc[idx_total, 1:], errors='coerce').fillna(0)
        
        total_alunos = len(lista_alunos)
        total_arrecadado = ofertas.sum()
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Alunos Cadastrados", total_alunos)
        col2.metric("Aulas Registradas", len(datas))
        col3.metric("Total Arrecadado (R$)", f"R$ {total_arrecadado:.2f}")
        
        st.markdown("---")
        
        df_graficos = pd.DataFrame({
            "Data": datas,
            "Valor da Oferta": ofertas,
            "Total Presentes": presencas
        })
        
        # Converte para datetime para garantir a ordenação no eixo X
        df_graficos['Data_Formatada'] = pd.to_datetime(df_graficos['Data'], format='%d/%m/%Y', errors='coerce')
        df_graficos = df_graficos.sort_values('Data_Formatada')
        
        col_grafico1, col_grafico2 = st.columns(2)
        
        with col_grafico1:
            st.subheader("Evolução da Presença")
            fig_freq = px.line(
                df_graficos, x='Data', y='Total Presentes', markers=True, 
                title="Alunos presentes por aula"
            )
            fig_freq.update_yaxes(rangemode="tozero")
            st.plotly_chart(fig_freq, use_container_width=True)
            
        with col_grafico2:
            st.subheader("Evolução Financeira")
            fig_ofertas = px.bar(
                df_graficos, x='Data', y='Valor da Oferta', 
                title="Ofertas arrecadadas por aula", text_auto='.2f'
            )
            st.plotly_chart(fig_ofertas, use_container_width=True)

            st.markdown("---")
            st.subheader("📥 Exportar Dados")
            
            # 1. Cria um buffer (arquivo em memória)
            buffer = io.BytesIO()
            
            # 2. Escreve os dataframes no arquivo Excel em memória
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                # Planilha 1: Resumo usado nos gráficos
                df_graficos.to_excel(writer, index=False, sheet_name="Resumo")
                
                # Planilha 2: O diário completo bruto
                df_diario.to_excel(writer, index=False, header=False, sheet_name="Diario_Completo")
                
            # 3. Retorna o cursor do buffer para o início
            buffer.seek(0)
            
            # 4. Cria o botão de download nativo do Streamlit
            st.download_button(
                label="📊 Baixar Relatório Completo (Excel)",
                data=buffer,
                file_name=f"relatorio_ebd_{date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Ainda não há aulas registradas no Diário. Acesse o menu lateral para lançar a primeira aula.")



# ------------------------------------------
# TELA: DIÁRIO DE AULA (Frequência e Ofertas)
# ------------------------------------------
elif menu == "Diário de Aula":
    st.title("Diário de Aula 📖")
    
    col1, col2 = st.columns([1, 2])
    with col1:
        data_selecionada = st.date_input("Data da Aula", date.today())
        data_str = data_selecionada.strftime("%d/%m/%Y")
        
    datas_existentes = df_diario.iloc[idx_data, :].tolist()
    
    # Se a data já existir na planilha, resgata os dados daquela coluna
    if data_str in datas_existentes:
        col_idx = datas_existentes.index(data_str)
        st.info(f"Editando registros encontrados para **{data_str}**.")
        licao_atual = str(df_diario.iloc[idx_licao, col_idx])
        oferta_atual = float(pd.to_numeric(df_diario.iloc[idx_oferta, col_idx], errors='coerce'))
        if pd.isna(oferta_atual): oferta_atual = 0.0
        presencas_atuais = df_diario.iloc[idx_oferta + 1 : idx_total, col_idx].tolist()
    # Se for uma data nova, inicializa os campos em branco/zerados
    else:
        col_idx = None
        st.info(f"Abrindo novo registro para **{data_str}**.")
        licao_atual = ""
        oferta_atual = 0.0
        presencas_atuais = ["F"] * len(lista_alunos)
        
    presencas_booleanas = [True if str(p).upper() == "P" else False for p in presencas_atuais]

    with col2:
        licao_input = st.text_input("Tema da Lição", value=licao_atual)
        oferta_input = st.number_input("Valor da Oferta (R$)", value=oferta_atual, min_value=0.0, step=1.0)
        
    st.write("### Lista de Presença")
    
    if len(lista_alunos) == 0:
        st.warning("Nenhum aluno cadastrado. Vá em 'Gestão de Alunos' para adicionar membros.")
    else:
        df_edit = pd.DataFrame({"Aluno": lista_alunos, "Presente": presencas_booleanas})
        
        df_editado = st.data_editor(
            df_edit,
            column_config={
                "Aluno": st.column_config.TextColumn("Nome", disabled=True),
                "Presente": st.column_config.CheckboxColumn("Presente?", default=False)
            },
            hide_index=True,
            use_container_width=True
        )
        
        total_presentes = df_editado["Presente"].sum()
        st.metric("Total de Alunos Presentes", total_presentes)
        
        if st.button("💾 Salvar Diário", type="primary"):
            presencas_formatadas = ["P" if p else "F" for p in df_editado["Presente"]]
            
            nova_coluna = [data_str, licao_input, f"{oferta_input:.2f}"] + presencas_formatadas + [str(total_presentes)]
            
            # Se for edição, sobrescreve a coluna; se for novo, anexa uma coluna nova
            if col_idx is not None:
                df_diario.iloc[:, col_idx] = nova_coluna
            else:
                novo_indice = df_diario.shape[1]
                df_diario[novo_indice] = nova_coluna
                
            aba_diario.clear()
            aba_diario.update(df_diario.values.tolist())
            st.success(f"✅ Diário do dia {data_str} salvo com sucesso!")
            st.rerun()

# ------------------------------------------
# TELA: GESTÃO DE ALUNOS
# ------------------------------------------
elif menu == "Gestão de Alunos":
    st.title("Gestão de Alunos 👥")
    
    st.subheader("Cadastrar Novo Aluno")
    with st.form("form_novo_aluno", clear_on_submit=True):
        nome_aluno = st.text_input("Nome Completo do Aluno")
        submit_button = st.form_submit_button("➕ Adicionar Aluno")
        
        if submit_button and nome_aluno:
            if nome_aluno in lista_alunos:
                st.error("Este aluno já está cadastrado!")
            else:
                # Cria a nova linha (Nome + Faltas retroativas para todas as aulas já cadastradas)
                nova_linha = [nome_aluno] + ["F"] * (df_diario.shape[1] - 1)
                df_nova_linha = pd.DataFrame([nova_linha], columns=df_diario.columns)
                
                # Fatiamento (slicing) para inserir o aluno ANTES da linha de "Total Presente"
                df_topo = df_diario.iloc[:idx_total]
                df_base = df_diario.iloc[idx_total:]
                
                df_atualizado = pd.concat([df_topo, df_nova_linha, df_base]).reset_index(drop=True)
                
                aba_diario.clear()
                aba_diario.update(df_atualizado.values.tolist())
                st.success(f"✅ {nome_aluno} cadastrado com sucesso!")
                st.rerun()
                
    st.markdown("---")
    st.subheader("Lista de Alunos (Ordem da Planilha)")
    
    if len(lista_alunos) > 0:
        for idx, aluno in enumerate(lista_alunos):
            col1, col2 = st.columns([4, 1])
            col1.write(f"- {aluno}")
            
            # Botão de remover (inativar) aluno
            if col2.button("🗑️ Remover", key=f"del_{aluno}"):
                linha_aluno_idx = rotulos.index(aluno)
                df_atualizado = df_diario.drop(linha_aluno_idx).reset_index(drop=True)
                aba_diario.clear()
                aba_diario.update(df_atualizado.values.tolist())
                st.success(f"Aluno {aluno} removido com sucesso!")
                st.rerun()
    else:
        st.write("Nenhum aluno cadastrado no diário.")