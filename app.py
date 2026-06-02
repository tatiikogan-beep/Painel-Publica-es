"""
app.py — Painel de Divisão de Publicações Diárias · Streamlit Cloud
"""

import re
import pandas as pd
import streamlit as st
from gerar_relatorio_publicacoes import (
    pre_check, gerar_relatorio, COORDENADORES_CONHECIDOS
)

st.set_page_config(page_title="Publicações Diárias — LegalOne", page_icon="⚖️", layout="centered")
st.title("⚖️ Divisão de Publicações Diárias")
st.caption("Controladoria Jurídica · LegalOne")

for k, v in [('extra_mappings',{}),('check_result',None),('last_filename',None),
             ('output_bytes',None),('resumo',None),('analistas_excluidos',[])]:
    if k not in st.session_state: st.session_state[k] = v

uploaded = st.file_uploader("Envie a planilha exportada do LegalOne (.xlsx)", type=['xlsx'],
    help="Duplicatas removidas automaticamente pelo número de CNJ.")
if not uploaded: st.info("Aguardando o arquivo do dia..."); st.stop()

file_bytes = uploaded.read()
filename   = uploaded.name
if st.session_state.last_filename != filename:
    for k in ('check_result','extra_mappings','output_bytes','resumo'):
        st.session_state[k] = {} if k=='extra_mappings' else None
    st.session_state.last_filename = filename

if st.session_state.check_result is None:
    with st.spinner("Analisando arquivo..."):
        result = pre_check(file_bytes, filename, st.session_state.extra_mappings)
    st.session_state.check_result = result
result = st.session_state.check_result

st.divider()
c1,c2,c3,c4 = st.columns(4)
c1.metric("📅 Data",   result['data_str'])
c2.metric("📆 Dia",    result['dia_semana_nome'])
c3.metric("📥 Bruto",  result['total_bruto'])
c4.metric("✅ Únicos", result['total_unico'],
          delta=f"-{result['duplicatas']} duplicatas", delta_color="off")

if result['inativos_encontrados']:
    with st.expander(f"🟠 {len(result['inativos_encontrados'])} responsável(is) inativo(s) — serão destacados em laranja"):
        for nome in result['inativos_encontrados']: st.write(f"• {nome}")

already_mapped   = set(st.session_state.extra_mappings.keys())
unmapped_pending = [r for r in result['unmapped'] if r not in already_mapped]

if unmapped_pending:
    st.divider()
    st.error(
        f"### ⚠️ {len(unmapped_pending)} responsável(is) sem coordenador mapeado\n\n"
        "Esses responsáveis **precisam ser informados ao gestor do sistema** para cadastro.\n"
        "Selecione o coordenador para continuar, ou gere o relatório com pendências marcadas."
    )
    st.dataframe(
        pd.DataFrame({"Responsável": unmapped_pending,
                      "Status": ["Sem coordenador"] * len(unmapped_pending),
                      "Ação": ["Informar ao gestor do sistema"] * len(unmapped_pending)}),
        hide_index=True, use_container_width=True,
    )
    st.markdown("**Selecione o coordenador de cada um:**")
    with st.form("form_coordenadores"):
        selecoes = {}
        for resp in unmapped_pending:
            key_safe = re.sub(r'[^a-zA-Z0-9_]', '_', resp)
            opcoes   = ["— selecione —"] + COORDENADORES_CONHECIDOS + ["✏️ Outro (digitar abaixo)"]
            selecoes[resp] = st.selectbox(label=resp, options=opcoes, key=f"sel_{key_safe}")
        outro_texto = st.text_input("Outro coordenador:", placeholder="Nome completo em maiúsculas")
        col_a, col_b = st.columns(2)
        submitted    = col_a.form_submit_button("✔ Confirmar mapeamentos", type="primary")
    if submitted:
        for resp, sel in selecoes.items():
            if sel == "✏️ Outro (digitar abaixo)" and outro_texto.strip():
                st.session_state.extra_mappings[resp] = outro_texto.strip().upper()
            elif sel not in ("— selecione —", "✏️ Outro (digitar abaixo)"):
                st.session_state.extra_mappings[resp] = sel
        st.session_state.check_result = None
        st.rerun()

st.divider()

# ── Painel de gestão de analistas ────────────────────────────────────────────
TODOS_ANALISTAS = ['VANESSA','PALOMA','BARBARA','LARA','ANNA JULIA','ANA CECILIA','ALANIS','TATIANA']
NOMES_ANALISTAS = {
    'VANESSA': 'Vanessa', 'PALOMA': 'Paloma', 'BARBARA': 'Bárbara', 'LARA': 'Lara',
    'ANNA JULIA': 'Anna Júlia', 'ANA CECILIA': 'Ana Cecília', 'ALANIS': 'Alanis', 'TATIANA': 'Tatiana'
}
with st.expander("⚙️ Gerenciar Analistas", expanded=False):
    col_inc, col_exc = st.columns(2)
    with col_inc:
        st.markdown("**✅ Analistas incluídos na análise**")
        analistas_incluidos = st.multiselect(
            "Incluir analistas:",
            options=TODOS_ANALISTAS,
            default=[a for a in TODOS_ANALISTAS if a not in st.session_state.analistas_excluidos],
            format_func=lambda x: NOMES_ANALISTAS.get(x, x),
            key="ms_incluir",
            label_visibility="collapsed"
        )
    with col_exc:
        st.markdown("**❌ Analistas excluídos da análise**")
        analistas_a_excluir = st.multiselect(
            "Excluir analistas:",
            options=TODOS_ANALISTAS,
            default=st.session_state.analistas_excluidos,
            format_func=lambda x: NOMES_ANALISTAS.get(x, x),
            key="ms_excluir",
            label_visibility="collapsed"
        )
    if st.button("🔄 Aplicar seleção de analistas"):
        st.session_state.analistas_excluidos = analistas_a_excluir
        st.session_state.check_result = None
        st.rerun()

especial = st.checkbox("⚠️ Divisão Especial (exceção pontual — RESUMO em vermelho)", value=False)
if st.button("▶ Gerar Relatório", type="primary", use_container_width=True):
    with st.spinner("Distribuindo publicações e gerando Excel..."):
        output_bytes, resumo = gerar_relatorio(file_bytes, filename, st.session_state.extra_mappings, divisao_especial=especial, analistas_excluidos=st.session_state.analistas_excluidos)
    st.session_state.output_bytes = output_bytes
    st.session_state.resumo       = resumo

if st.session_state.output_bytes and st.session_state.resumo:
    resumo = st.session_state.resumo
    st.success(f"✅ Relatório gerado — {resumo['data_str']} ({resumo['dia_semana_nome']})")
    st.subheader("Distribuição")
    NOMES = {'VANESSA':'Vanessa','PALOMA':'Paloma','BARBARA':'Bárbara','LARA':'Lara',
             'ANNA JULIA':'Anna Júlia','ANA CECILIA':'Ana Cecília','ALANIS':'Alanis','TATIANA':'Tatiana'}
    ORDEM = [a for a in ['VANESSA','PALOMA','BARBARA','LARA','ANNA JULIA','ANA CECILIA','ALANIS','TATIANA'] if a not in st.session_state.analistas_excluidos]
    rows = []
    for a in ORDEM:
        nome = NOMES.get(a, a)
        if a=='ALANIS' and not resumo['alanis_ativa']:       rows.append((nome,"—","⛔ Fora (terça-feira)"))
        elif a=='ANNA JULIA' and not resumo['anna_julia_ativa']: rows.append((nome,"—","⛔ Fora (quinta-feira)"))
        elif a not in resumo['cotas']:
            if a=='TATIANA': rows.append((nome,"—","⚪ Fora (cota ≤ 40)"))
        else:
            qtd = resumo['alloc_counts'].get(a,0)
            obs = ('✅ Incluída' if a=='TATIANA' else
                   '50% da cota' if a in ('ALANIS','ANNA JULIA','ANA CECILIA') else
                   'Prioridade GPM' if a=='BARBARA' else
                   'Prioridade não-trabalhista' if a=='PALOMA' else '🟢 Ativa')
            rows.append((nome, str(qtd), obs))
    st.dataframe(pd.DataFrame(rows, columns=["Analista","Publicações","Observação"]),
                 hide_index=True, use_container_width=True)

    pendencias = [r for r in result.get('unmapped',[]) if r not in st.session_state.extra_mappings]
    if pendencias:
        st.divider()
        st.warning(
            f"⚠️ **{len(pendencias)} responsável(is) sem coordenador** constam no relatório "
            "marcados como pendentes.\n\n"
            "**Informe ao gestor do sistema para cadastrar os coordenadores faltantes.**"
        )
        st.dataframe(
            pd.DataFrame({"Responsável": pendencias,
                          "Ação necessária": ["Informar coordenador ao gestor do sistema"]*len(pendencias)}),
            hide_index=True, use_container_width=True,
        )

    st.divider()
    nome_arq = f"publicacoes_{resumo['data_str'].replace('/','_')}.xlsx"
    st.download_button("⬇ Baixar Excel", data=st.session_state.output_bytes,
                       file_name=nome_arq,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True, type="primary")
    st.caption("Abas: RESUMO · analistas individuais · POR COORDENADOR · POR RESPONSÁVEL")
