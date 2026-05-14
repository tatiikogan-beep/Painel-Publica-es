"""
app.py — Painel de Divisão de Publicações Diárias
Deploy: Streamlit Cloud  |  github.com/seu-repo/PAINEL_WEB_PUBLICACOES
"""

import re
import streamlit as st
from gerar_relatorio_publicacoes import (
    pre_check, gerar_relatorio, COORDENADORES_CONHECIDOS
)

# ─── CONFIGURAÇÃO DA PÁGINA ───────────────────────────────────────────────────

st.set_page_config(
    page_title="Publicações Diárias — LegalOne",
    page_icon="⚖️",
    layout="centered",
)

st.title("⚖️ Divisão de Publicações Diárias")
st.caption("Controladoria Jurídica · LegalOne")

# ─── ESTADO DA SESSÃO ─────────────────────────────────────────────────────────

if 'extra_mappings' not in st.session_state:
    st.session_state.extra_mappings = {}
if 'check_result' not in st.session_state:
    st.session_state.check_result = None
if 'last_filename' not in st.session_state:
    st.session_state.last_filename = None
if 'output_bytes' not in st.session_state:
    st.session_state.output_bytes = None
if 'resumo' not in st.session_state:
    st.session_state.resumo = None

# ─── UPLOAD ───────────────────────────────────────────────────────────────────

uploaded = st.file_uploader(
    "Envie a planilha exportada do LegalOne (.xlsx)",
    type=['xlsx'],
    help="Arquivo exportado diretamente do LegalOne. Duplicatas serão removidas automaticamente pelo número de CNJ."
)

if not uploaded:
    st.info("Aguardando o arquivo do dia...")
    st.stop()

file_bytes = uploaded.read()
filename   = uploaded.name

# Reseta o estado se o arquivo mudou
if st.session_state.last_filename != filename:
    st.session_state.check_result  = None
    st.session_state.extra_mappings = {}
    st.session_state.output_bytes  = None
    st.session_state.resumo        = None
    st.session_state.last_filename = filename

# ─── PRE-CHECK ────────────────────────────────────────────────────────────────

if st.session_state.check_result is None:
    with st.spinner("Analisando arquivo..."):
        result = pre_check(file_bytes, filename, st.session_state.extra_mappings)
    st.session_state.check_result = result

result = st.session_state.check_result

# ─── MÉTRICAS DO ARQUIVO ──────────────────────────────────────────────────────

st.divider()
col1, col2, col3, col4 = st.columns(4)
col1.metric("📅 Data", result['data_str'])
col2.metric("📆 Dia", result['dia_semana_nome'])
col3.metric("📥 Bruto", result['total_bruto'])
col4.metric("✅ Únicos", result['total_unico'],
            delta=f"-{result['duplicatas']} duplicatas",
            delta_color="off")

# ─── INATIVOS ENCONTRADOS ────────────────────────────────────────────────────

if result['inativos_encontrados']:
    with st.expander(f"🟠 {len(result['inativos_encontrados'])} responsável(is) inativo(s) — serão destacados em laranja"):
        for nome in result['inativos_encontrados']:
            st.write(f"• {nome}")

# ─── COORDENADORES NÃO MAPEADOS ──────────────────────────────────────────────

# Filtra apenas os que ainda não foram mapeados nesta sessão
already_mapped = set(st.session_state.extra_mappings.keys())
unmapped_pending = [r for r in result['unmapped'] if r not in already_mapped]

if unmapped_pending:
    st.divider()
    st.warning(
        f"⚠️ **{len(unmapped_pending)} responsável(is) sem coordenador mapeado.**  \n"
        "Informe o coordenador de cada um antes de gerar o relatório:"
    )

    with st.form("form_coordenadores"):
        selecoes = {}
        for resp in unmapped_pending:
            # Sanitiza chave para o Streamlit (remove caracteres especiais)
            key_safe = re.sub(r'[^a-zA-Z0-9_]', '_', resp)
            opcoes = ["— selecione —"] + COORDENADORES_CONHECIDOS + ["✏️ Outro (digitar abaixo)"]
            sel = st.selectbox(label=resp, options=opcoes, key=f"sel_{key_safe}")
            selecoes[resp] = sel

        outro_texto = st.text_input(
            "Outro coordenador (se selecionou '✏️ Outro' acima):",
            placeholder="Nome completo em maiúsculas"
        )
        submitted = st.form_submit_button("✔ Confirmar mapeamentos", type="primary")

    if submitted:
        algum_confirmado = False
        for resp, sel in selecoes.items():
            if sel == "✏️ Outro (digitar abaixo)" and outro_texto.strip():
                st.session_state.extra_mappings[resp] = outro_texto.strip().upper()
                algum_confirmado = True
            elif sel not in ("— selecione —", "✏️ Outro (digitar abaixo)"):
                st.session_state.extra_mappings[resp] = sel
                algum_confirmado = True
        if algum_confirmado:
            st.session_state.check_result = None   # força re-check com novos mapeamentos
            st.rerun()

    # Bloqueia geração enquanto houver pendentes sem mapear
    ainda_pendentes = [r for r in unmapped_pending if r not in st.session_state.extra_mappings]
    if ainda_pendentes:
        st.stop()

# ─── GERAÇÃO DO RELATÓRIO ─────────────────────────────────────────────────────

st.divider()

if st.button("▶ Gerar Relatório", type="primary", use_container_width=True):
    with st.spinner("Distribuindo publicações e gerando Excel..."):
        output_bytes, resumo = gerar_relatorio(
            file_bytes,
            filename,
            st.session_state.extra_mappings
        )
    st.session_state.output_bytes = output_bytes
    st.session_state.resumo = resumo

# ─── RESULTADO ────────────────────────────────────────────────────────────────

if st.session_state.output_bytes and st.session_state.resumo:
    resumo = st.session_state.resumo
    st.success(f"✅ Relatório gerado — {resumo['data_str']} ({resumo['dia_semana_nome']})")

    # Tabela de distribuição
    st.subheader("Distribuição")

    NOMES_EXIB = {
        'VANESSA': 'Vanessa', 'PALOMA': 'Paloma', 'BARBARA': 'Bárbara',
        'LARA': 'Lara', 'ANNA JULIA': 'Anna Júlia', 'ALANIS': 'Alanis', 'TATIANA': 'Tatiana'
    }
    ORDEM = ['VANESSA', 'PALOMA', 'BARBARA', 'LARA', 'ANNA JULIA', 'ALANIS', 'TATIANA']

    rows_tabela = []
    for a in ORDEM:
        nome = NOMES_EXIB.get(a, a)
        if a == 'ALANIS' and not resumo['alanis_ativa']:
            rows_tabela.append((nome, "—", "⛔ Fora (terça-feira)"))
        elif a == 'ANNA JULIA' and not resumo['anna_julia_ativa']:
            rows_tabela.append((nome, "—", "⛔ Fora (quinta-feira)"))
        elif a not in resumo['cotas']:
            if a == 'TATIANA':
                rows_tabela.append((nome, "—", "⚪ Fora (cota ≤ 40)"))
        else:
            qtd = resumo['alloc_counts'].get(a, 0)
            obs = ""
            if a == 'TATIANA':
                obs = "✅ Incluída"
            elif a in ('ALANIS', 'ANNA JULIA'):
                obs = "50% da cota"
            if a == 'BARBARA':
                obs = "Prioridade GPM"
            if a == 'PALOMA':
                obs = "Prioridade não-trabalhista"
            rows_tabela.append((nome, qtd, obs))

    import pandas as pd
    df_tabela = pd.DataFrame(rows_tabela, columns=["Analista", "Publicações", "Observação"])
    st.dataframe(df_tabela, hide_index=True, use_container_width=True)

    # Download
    st.divider()
    nome_arquivo = f"publicacoes_{resumo['data_str'].replace('/', '_')}.xlsx"
    st.download_button(
        label="⬇ Baixar Excel",
        data=st.session_state.output_bytes,
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

    st.caption(
        f"Abas geradas: RESUMO · analista individual (c/ destaque de tribunal minoritário) · "
        f"POR COORDENADOR · POR RESPONSAVEL"
    )
