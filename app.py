import streamlit as st
import datetime as dt
import io
import traceback

# ── Configuração da página ────────────────────────────────────
st.set_page_config(
    page_title="Relatório Leads / MQL",
    page_icon="📊",
    layout="centered",
)

# ── CSS customizado ───────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #f8f9fb; }
    .stButton > button {
        background-color: #1a6b3c;
        color: white;
        font-size: 16px;
        padding: 0.6rem 2rem;
        border-radius: 8px;
        border: none;
        width: 100%;
    }
    .stButton > button:hover { background-color: #145530; }
    .block-container { padding-top: 2rem; }
    .upload-box {
        background: white;
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
        margin-bottom: 1rem;
    }
    h1 { color: #1a6b3c; }
    h3 { color: #333; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────
st.title("📊 Relatório Leads / MQL")
st.markdown("Faça upload da base de leads, configure os intervalos e baixe o relatório Excel gerado automaticamente.")
st.divider()

# ── Upload ────────────────────────────────────────────────────
st.markdown("### 1. Base de Leads")
uploaded_file = st.file_uploader(
    "Selecione o arquivo Excel (.xlsx) exportado do RD Station / CRM",
    type=["xlsx"],
    help="O arquivo deve conter as colunas: Data da Primeira Conversão, Data da Última Oportunidade, Origem, Tags, E-mail, Telefone e Nome."
)

if uploaded_file:
    st.success(f"✅ Arquivo carregado: **{uploaded_file.name}**")

st.divider()

# ── Intervalos ────────────────────────────────────────────────
st.markdown("### 2. Intervalo Mensal")
st.markdown("Define quais meses aparecerão nas abas **Consolidado** e **Sumário**.")

col1, col2 = st.columns(2)
with col1:
    mes_inicio = st.date_input(
        "Mês de início",
        value=dt.date.today().replace(day=1),
        help="Será usado apenas o mês/ano desta data."
    )
with col2:
    mes_fim = st.date_input(
        "Mês de fim",
        value=dt.date.today().replace(day=1),
        help="Será usado apenas o mês/ano desta data."
    )

st.divider()

st.markdown("### 3. Intervalo Semanal")
st.markdown("Define o período das semanas exibidas nas abas mensais.")

col3, col4 = st.columns(2)
with col3:
    data_inicio_sem = st.date_input(
        "Data de início das semanas",
        value=dt.date.today().replace(day=1),
        key="sem_inicio"
    )
with col4:
    data_fim_sem = st.date_input(
        "Data de fim das semanas",
        value=dt.date.today().replace(day=1),
        key="sem_fim"
    )

st.divider()

# ── Nome do arquivo de saída ──────────────────────────────────
st.markdown("### 4. Nome do Arquivo de Saída")
nome_saida = st.text_input(
    "Nome do arquivo (.xlsx)",
    value="relatorio_leads_mql.xlsx",
    help="O arquivo será disponibilizado para download após a geração."
)
if not nome_saida.endswith(".xlsx"):
    nome_saida += ".xlsx"

st.divider()

# ── Botão de geração ─────────────────────────────────────────
gerar = st.button("🚀 Gerar Relatório")

if gerar:
    # Validações
    if not uploaded_file:
        st.error("❌ Por favor, faça o upload da base de leads antes de continuar.")
        st.stop()

    if mes_inicio > mes_fim:
        st.error("❌ O mês de início não pode ser posterior ao mês de fim.")
        st.stop()

    if data_inicio_sem > data_fim_sem:
        st.error("❌ A data de início das semanas não pode ser posterior à data de fim.")
        st.stop()

    # ── Processamento ─────────────────────────────────────────
    with st.spinner("⏳ Processando a base de leads... Aguarde."):
        try:
            import pandas as pd
            from relatorio_leads_mql import (
                load_base, map_columns, preprocess,
                collect_primary_tags_from_first_conversion,
                classify_primary_from_list,
                collect_all_tags, cl_historico,
                build_month_range,
                gerar_consolidado, gerar_semanais,
                gerar_auditoria, gerar_validacao,
                gerar_sumario, gerar_arquivo_final,
            )

            # Salva o upload num buffer para o pandas conseguir ler
            file_bytes = io.BytesIO(uploaded_file.read())
            df_raw = pd.read_excel(file_bytes)
            df_raw.columns = [str(c) for c in df_raw.columns]

            cols = map_columns(df_raw)
            df = preprocess(df_raw.copy(), cols)

            # Tags
            primary_tags_source = collect_primary_tags_from_first_conversion(df, cols)
            primary_map = {lid: classify_primary_from_list(tags) for lid, tags in primary_tags_source.items()}
            all_tags_source = collect_all_tags(df, cols)
            cl_map = {lid: cl_historico(tags) for lid, tags in all_tags_source.items()}

            # Intervalos — constrói direto com objetos date, sem parsers de string
            import datetime as _dt
            intervalo_mensal  = (
                _dt.date(mes_inicio.year, mes_inicio.month, 1),
                _dt.date(mes_fim.year,    mes_fim.month,    1),
            )
            intervalo_semanal = (data_inicio_sem, data_fim_sem)
            months = build_month_range(*intervalo_mensal)

            # Relatório
            consolidado   = gerar_consolidado(df, primary_map, cl_map, months)
            abas_semanais = gerar_semanais(df, primary_map, cl_map, intervalo_mensal, intervalo_semanal)
            auditoria_df, descartes = gerar_auditoria(df, cols, primary_map, cl_map, intervalo_mensal)
            validacao_df  = gerar_validacao(consolidado, abas_semanais)
            sumario_dict  = gerar_sumario(consolidado)

            # Salva em memória (sem gravar em disco)
            output_buffer = io.BytesIO()
            gerar_arquivo_final(
                consolidado, abas_semanais, auditoria_df,
                descartes, validacao_df, sumario_dict,
                output_buffer  # aceita path ou file-like object
            )
            output_buffer.seek(0)
            xlsx_bytes = output_buffer.read()

            st.success("✅ Relatório gerado com sucesso!")

            # Métricas rápidas
            total_leads = int(consolidado["Leads"].sum())
            total_mqls  = int(consolidado["MQLs"].sum())
            conv_pct    = f"{(total_mqls/total_leads*100):.1f}%" if total_leads > 0 else "—"

            m1, m2, m3 = st.columns(3)
            m1.metric("Total de Leads", f"{total_leads:,}".replace(",", "."))
            m2.metric("Total de MQLs",  f"{total_mqls:,}".replace(",", "."))
            m3.metric("Conversão MQL/Lead", conv_pct)

            st.divider()
            st.download_button(
                label="⬇️ Baixar Relatório Excel",
                data=xlsx_bytes,
                file_name=nome_saida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        except Exception as e:
            st.error(f"❌ Erro durante o processamento:\n\n```\n{traceback.format_exc()}\n```")
