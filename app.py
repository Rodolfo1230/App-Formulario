import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(layout="wide", page_title="Formulário de Auditoria")

st.title("📋 Formulário de Auditoria e Recebimento")

# =========================
# SIDEBAR
# =========================
st.sidebar.title("Navegação")
pagina = st.sidebar.radio(
    "Escolha uma página",
    ["Recebimento", "Dashboard", "Configurações"]
)

if pagina == "Recebimento":
    with st.form("meu_formulario"):
        st.subheader("Dados da Avaliação")
        
        col1, col2 = st.columns(2)
        
        with col1:
            avaliador = st.text_input("Nome do avaliador")
            area = st.selectbox("Área", ["Selecione...", "Recebimento", "Expedição", "Estoque", "Shirink", "Picking", "Buffer", "Inventário", "Qualidade"])
            o_que_sera_avaliado = st.selectbox("O que será avaliado?", ["Selecione...", "Padrão do Setor", "Padrão de Processos"])
            qual_operacao_sera_avaliada = st.selectbox("Qual operação?", ["Selecione...", "Multilivros", "Macmillan", "Asurion"])

        with col2:
            responsavel_area_pontuada = st.text_input("Responsável pela área pontuada")
            prazo_melhoria_recorrer = st.text_input("Prazo para a melhoria ocorrer")

        st.divider()
        st.subheader("Critérios de Avaliação")

        # Organizando em colunas para não ficar uma lista infinita
        c1, c2 = st.columns(2)
        
        with c1:
            sinalizacao_iluminacao_adequada = st.selectbox("Sinalização e ILUMINAÇÃO adequada?", ["Selecione...", "Sim", "Não"])
            ambiente_organizado = st.selectbox("Ambiente organizado (sem itens fora de layout)?", ["Selecione...", "Sim", "Não"])
            ambiente_epc_saida_emergencia = st.selectbox("EPC's e SAÍDAS DE EMERGÊNCIA adequadas?", ["Selecione...", "Sim", "Não"])
            ambiente_equipamentos_seguranca_desobstruidos = st.selectbox("EQUIPAMENTOS DE SEGURANÇA desobstruídos?", ["Selecione...", "Sim", "Não"])
            bancadas_conferencia_devidamente_organizada_limpa = st.selectbox("Bancadas ORGANIZADAS e LIMPAS?", ["Selecione...", "Sim", "Não"])
            equipamentos_condicoes_uso = st.selectbox("Equipamentos em boas condições?", ["Selecione...", "Sim", "Não"])
            identificacoes_estao_ok = st.selectbox("Identificações ok? (lixeiras, volumes)", ["Selecione...", "Sim", "Não"])

        with c2:
            sinalizacao_fita_demarcacao_solo = st.selectbox("Fitas de demarcação em boas condições?", ["Selecione...", "Sim", "Não"])
            balanca_funcionando = st.selectbox("A balança está funcionando?", ["Selecione...", "Sim", "Não"])
            desvios_ruptura_atividades_realizadas = st.selectbox("Desvios podem trazer rupturas?", ["Selecione...", "Sim", "Não"])
            situacao_recorrente = st.selectbox("Situação é recorrente?", ["Selecione...", "Sim", "Não"])
            retrabalhos_morosidade_fluxo_trabalho = st.selectbox("Existem retrabalhos/morosidade?", ["Selecione...", "Sim", "Não"])
            existe_item_nao_e_departamento = st.selectbox("Item que não é do departamento?", ["Selecione...", "Sim", "Não", "Não se aplica"])
            colaboradores_uso_correto_epis = st.selectbox("Uso correto de EPI's?", ["Selecione...", "Sim", "Não"])

        st.divider()
        melhoria_depender_outras_areas_acontecer = st.text_area("Melhoria depende de outras áreas? (Manutenção, Orçamentos, etc)")
        retrabalhos_morosidade_fluxo_trabalho_especifique = st.text_area("Se houve retrabalho, especifique:")
        comentario_final = st.text_area("Sugestões de melhoria e Comentários Finais")

        submitted = st.form_submit_button("Gerar Relatório para Exportação")

    if submitted:
        # Criando DataFrame
        data_exportacao = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        dados = {
            "Data Exportacao": data_exportacao,
            "Avaliador": avaliador,
            "Área": area,
            "Operação": qual_operacao_sera_avaliada,
            "Sinalização/Iluminação": sinalizacao_iluminacao_adequada,
            "Organização": ambiente_organizado,
            "Responsável": responsavel_area_pontuada,
            "Comentário Final": comentario_final
            # Adicione as outras variáveis aqui seguindo o mesmo padrão
        }
        
        df = pd.DataFrame([dados])

        st.success("✅ Relatório gerado com sucesso!")
        st.dataframe(df)

        # Gerar Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Auditoria")
        
        st.download_button(
            label="📥 Baixar Planilha Excel (XLSX)",
            data=output.getvalue(),
            file_name=f"auditoria_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif pagina == "Dashboard":
    st.info("Página de Dashboard em desenvolvimento...")
