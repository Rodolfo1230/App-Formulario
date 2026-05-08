import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(layout="wide")

st.title("Formulário + Exportação XLSX")

# =========================
# SIDEBAR
# =========================
st.sidebar.title("Menu")

pagina = st.sidebar.radio(
    "Escolha uma página",
    [
        "Recebimento",
        "Dashboard",
        "Configurações"
    ]
)

# =========================
# PÁGINA FORMULÁRIO
# =========================
if pagina == "Recebimento":

    # =========================
    # FORMULÁRIO
    # =========================
    with st.form("meu_formulario"):

        # Campo de texto
        avaliador = st.text_area(
            "Escreva o nome do avaliador"
        )

        # Lista suspensa
        area = st.selectbox(
            "Área",
            ["Selecione...",
            "Recebimento",
            "Expedição",
            "Estoque",
            "Shirink",
            "Picking",
            "Buffer",
            "Inventário",
            "Qualidade"
            ]
        )

        # Lista suspensa
        o_que_sera_avaliado = st.selectbox(
            "O que será avaliado ?",
            ["Selecione...",
            "Padrão do Setor",
            "Padrão de Processos"
            ]
        )

        # Lista suspensa
        qual_operacao_sera_avaliada = st.selectbox(
            "Qual operação será avaliada ?",
            ["Selecione...",
            "Multilivros",
            "Macmillan",
            "Asurion"
            ]
        )

        # Lista suspensa
        sinalizacao_iluminacao_adequada = st.selectbox(
            "O ambiente apresenta SINALIZAÇÃO e ILUMINAÇÃO adequada?",
            ["Selecione...",
            "Sim",
            "Não"
            ]
        )

        # Lista suspensa
        ambiente_organizado = st.selectbox(
            "O ambiente está organizado, sem Itens fora de Layout? (carrinhos, coletores, pallets...)",
            ["Selecione...",
            "Sim",
            "Não"
            ]
        )

        # Lista suspensa
        ambiente_epc_saida_emergencia = st.selectbox(
            "O ambiente apresenta EPC's e SAÍDAS DE EMERGÊNCIA adequadas?",
            ["Selecione...",
            "Sim",
            "Não"
            ]
        )

        # Lista suspensa
        ambiente_equipamentos_seguranca_desobstruidos = st.selectbox(
            "O ambiente apresenta EQUIPAMENTOS DE SEGURANÇA desobstruídos?",
            ["Selecione...",
            "Sim",
            "Não"
            ]
        )

        # Lista suspensa
        bancadas_conferencia_devidamente_organizada_limpa = st.selectbox(
            "Bancadas de conferência devidamente ORGANIZADAS e LIMPAS?",
            ["Selecione...",
            "Sim",
            "Não"
            ]
        )

        # Lista suspensa
        equipamentos_condicoes_uso = st.selectbox(
            "O ambiente apresenta equipamentos em boas condições de uso? (Carrinhos, Fitas Gomadas, Maquinário...)",
            ["Selecione...",
            "Sim",
            "Não"
            ]
        )

        # Lista suspensa
        identificacoes_estao_ok = st.selectbox(
            "As identificações estão ok? (lixeiras, volumes)",
            ["Selecione...",
            "Sim",
            "Não"
            ]
        )

        # Lista suspensa
        sinalizacao_fita_demarcacao_solo = st.selectbox(
            "As sinalizações com fita demarcação de solo estão em boas condições?",
            ["Selecione...",
            "Sim",
            "Não"
            ]
        )

        # Lista suspensa
        balanca_funcionando = st.selectbox(
            "A balança está funcionando?",
            ["Selecione...",
            "Sim",
            "Não"
            ]
        )

        # Campo de texto
        melhoria_depender_outras_areas_acontecer = st.text_area(
            "Melhoria irá depender de outras áreas para acontecer ? (EX: Manutenção, Orçamentos, etc...)"
        )

        # Lista suspensa
        desvios_ruptura_atividades_realizadas = st.selectbox(
            "Os desvios poderão trazer rupturas nas atividades ali realizadas?",
            ["Selecione...",
            "Sim",
            "Não"
            ]
        )

        # Lista suspensa
        situacao_recorrente = st.selectbox(
            "Situação é recorrente ? (Ou seja, não foi tratada, ou já aconteceu anteriormente)",
            ["Selecione...",
            "Sim",
            "Não"
            ]
        )

        # Lista suspensa
        retrabalhos_morosidade_fluxo_trabalho = st.selectbox(
            "Existem retrabalhos que podem acabar trazendo morosidade ao fluxo de trabalho?",
            ["Selecione...",
            "Sim",
            "Não"
            ]
        )

        # Campo de texto
        retrabalhos_morosidade_fluxo_trabalho_especifique = st.text_area(
            "Se a resposta anterior for 'Sim', por favor, especifique."
        )

        # Lista suspensa
        existe_item_nao_e_departamento = st.selectbox(
            "Existe algum item que está em recebimento que não é do departamento?",
            ["Selecione...",
            "Sim",
            "Não",
            "Não se aplica"
            ]
        )

        # Lista suspensa
        paletes_buffer_recebimento_caixas_devidamente_fechadas = st.selectbox(
            "Os paletes do buffer de recebimento estão com as caixas devidamente fechadas?",
            ["Selecione...",
            "Sim",
            "Não"
            ]
        )

        # Lista suspensa
        tratativas_devolucao_area_demarcada = st.selectbox(
            "As tratativas de devolução estão em sua área demarcada?",
            ["Selecione...",
            "Sim",
            "Não"
            ]
        )

        # Lista suspensa
        colaboradores_uso_correto_epis = st.selectbox(
            "Colaboradores fazem uso correto de EPI's ?",
            ["Selecione...",
            "Sim",
            "Não"
            ]
        )

        # Lista suspensa
        colaborador_entrevistado_apresenta_conhecimento_dds_do_dia = st.selectbox(
            "O colaborador entrevistado apresenta conhecimento sobre o tema do DDS do dia?",
            ["Selecione...",
            "Sim",
            "Não"
            ]
        )

        # Campo de texto
        prazo_melhoria_recorrer = st.text_area(
            "Qual o prazo para a melhoria ocorrer ?"
        )

        # Lista suspensa
        responsavel_area_pontuada = st.selectbox(
            "Qual o nome do responsável pela área pontuada?",
            ["Selecione...",
            "Jorge Santos",
            "Luciano Souza",
            "Eder Silva",
            "Marcelo Marcondes",
            "Elisvaldo Santos",
            "Paulo Cristoveny",
            "Marcio Xavier",
            "Anibal Aguiar"
            ]
        )

        # Campo de texto
        comentario_final = st.text_area(
            "Deixe aqui seu comentário do que foi concluído durante a observação e descreva sua sugestão de melhorias. Assim juntos vamos construir um ambiente cada vez melhor e mais seguro para se trabalhar!!"
        )

        # Multiselect (caixinhas)
#        itens = st.multiselect(
#            "Selecione os itens",
#            ["Notebook", "Monitor", "Mouse", "Teclado", "Headset"]
#        )

        # Botão submit do formulário
        submitted = st.form_submit_button("Salvar")

    # =========================
    # PROCESSAR DADOS
    # =========================
    if submitted:


        # Data e hora atual
        data_exportacao = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        # Criando DataFrame
        df = pd.DataFrame({
            "Avaliador": [avaliador],
            "Área": [area],
            "O que será avaliado ?": [o_que_sera_avaliado],
            "Qual operação será avaliada ?": [qual_operacao_sera_avaliada],
            "O ambiente apresenta SINALIZAÇÃO e ILUMINAÇÃO adequada?": [sinalizacao_iluminacao_adequada],
            "O ambiente está organizado, sem Itens fora de Layout? (carrinhos, coletores, pallets...)": [ambiente_organizado],
            "O ambiente apresenta EPC's e SAÍDAS DE EMERGÊNCIA adequadas?": [ambiente_epc_saida_emergencia],
            "O ambiente apresenta EQUIPAMENTOS DE SEGURANÇA desobstruídos?": [ambiente_equipamentos_seguranca_desobstruidos],
            "Bancadas de conferência devidamente ORGANIZADAS e LIMPAS?": [bancadas_conferencia_devidamente_organizada_limpa],
            "O ambiente apresenta equipamentos em boas condições de uso? (Carrinhos, Fitas Gomadas, Maquinário...)": [equipamentos_condicoes_uso],
            "As identificações estão ok? (lixeiras, volumes)": [identificacoes_estao_ok],
            "As sinalizações com fita demarcação de solo estão em boas condições?": [sinalizacao_fita_demarcacao_solo],
            "A balança está funcionando?": [balanca_funcionando],
            "Melhoria irá depender de outras áreas para acontecer ? (EX: Manutenção, Orçamentos, etc...)": [melhoria_depender_outras_areas_acontecer],
            "Os desvios poderão trazer rupturas nas atividades ali realizadas?": [desvios_ruptura_atividades_realizadas],
            "Situação é recorrente ? (Ou seja, não foi tratada, ou já aconteceu anteriormente)": [situacao_recorrente],
            "Existem retrabalhos que podem acabar trazendo morosidade ao fluxo de trabalho?": [retrabalhos_morosidade_fluxo_trabalho],
            "Se a resposta anterior for 'Sim', por favor, especifique.": [retrabalhos_morosidade_fluxo_trabalho_especifique],
            "Existe algum item que está em recebimento que não é do departamento?": [existe_item_nao_e_departamento],
            "Os paletes do buffer de recebimento estão com as caixas devidamente fechadas?": [paletes_buffer_recebimento_caixas_devidamente_fechadas],
            "As tratativas de devolução estão em sua área demarcada?": [tratativas_devolucao_area_demarcada],
            "Colaboradores fazem uso correto de EPI's ?": [colaboradores_uso_correto_epis],
            "O colaborador entrevistado apresenta conhecimento sobre o tema do DDS do dia?": [colaborador_entrevistado_apresenta_conhecimento_dds_do_dia],
            "Qual o prazo para a melhoria ocorrer ?": [prazo_melhoria_recorrer],
            "Qual o nome do responsável pela área pontuada?": [responsavel_area_pontuada],
            "Comentario Final": [comentario_final],
#            "Itens Selecionados": [", ".join(itens)],
            "Data Exportacao": [data_exportacao]
        })

        st.success("Dados salvos!")

        st.dataframe(df)

        # =========================
        # GERAR XLSX EM MEMÓRIA
        # =========================
        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Formulario")

        output.seek(0)

        # =========================
        # BOTÃO EXPORTAR
        # =========================
        st.download_button(
            label="Exportar XLSX",
            data=output,
            file_name="formulario.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )