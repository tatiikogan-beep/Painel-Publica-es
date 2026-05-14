"""
gerar_relatorio_publicacoes.py
Módulo de distribuição de publicações diárias — LegalOne
Uso: from gerar_relatorio_publicacoes import pre_check, gerar_relatorio
"""

import io
import re
import unicodedata
from datetime import datetime, date
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

# ─── NORMALIZAÇÃO ─────────────────────────────────────────────────────────────

def normalizar(s):
    if not s or (isinstance(s, float)):
        return ""
    s = str(s).upper().strip()
    s = unicodedata.normalize('NFD', s)
    return ''.join(c for c in s if unicodedata.category(c) != 'Mn')


# ─── DADOS FIXOS ──────────────────────────────────────────────────────────────

# Responsáveis cujo coordenador já foi confirmado manualmente
COORDENADORES_MAPEADOS = {
    normalizar("RODRIGO RIBEIRO ANTUNES QUARIGUASI"):      "GABRIEL GIORGIO CICCHELERO",
    normalizar("CAMILLA GOES BARBOSA"):                     "CAMILLA GOES BARBOSA",
    normalizar("SUZANA MARIA CAMPOS MARANHAO DE LIMA"):     "SUZANA MARIA CAMPOS MARANHAO DE LIMA",
    normalizar("DALILA DRISANA GOMES GONCALVES"):           "GABRIEL GIORGIO CICCHELERO",
    normalizar("ANA VITORIA SALES DE OLIVEIRA FALCAO"):     "GABRIEL GIORGIO CICCHELERO",
    normalizar("MARCELO LIMA ARRAIS"):                      "HELANZIA DE ARAUJO XAVIER WICHAMNN",
    normalizar("WELLINGTON PEREIRA DA ROCHA FILHO"):        "HELANZIA DE ARAUJO XAVIER WICHAMNN",
    normalizar("ARTUR SARAIVA DE ANDRADE"):                 "HELANZIA DE ARAUJO XAVIER WICHAMNN",
    normalizar("JOHANN DANIEL DE OLIVEIRA INOCENCIO"):      "SUZANA MARIA CAMPOS MARANHAO DE LIMA",
    normalizar("EMERSON DE ALMEIDA MELO JUNIOR"):           "NAYANDERSON LUAN MELLO PINHEIRO",
    normalizar("RAFAEL CAVALCANTE BARSOSA"):                "GABRIEL GIORGIO CICCHELERO",
    normalizar("DALILA CARLOS DE CASTRO"):                  "GABRIEL GIORGIO CICCHELERO",
}

# Lista de coordenadores conhecidos para selectbox no Streamlit
COORDENADORES_CONHECIDOS = [
    "CAMILLA GOES BARBOSA",
    "GABRIEL GIORGIO CICCHELERO",
    "HELANZIA DE ARAUJO XAVIER WICHAMNN",
    "JENIFFER ROSA BARBOSA DE SALES",
    "LUCIANE MODERNEL MENDES",
    "NAYANDERSON LUAN MELLO PINHEIRO",
    "SUZANA MARIA CAMPOS MARANHAO DE LIMA",
    "YURI ALVES BARROS DOS SANTOS",
]

# Responsáveis com Ativo = Não
INATIVOS = {normalizar(n) for n in [
    "ALEXANDRE SOUZA DA SILVA FILHO", "ALICE RIBEIRO MELO CRISOSTOMO",
    "ALLY CHARLYS MUSSE MENDES FILHO", "AMANDA BEZERRA DE MENEZES NETO",
    "AMANDA MOURA DOS SANTOS BRAGA", "ANA CAMILA CIFONI DE VASCONCELOS BARROSO",
    "ANA CAROLINA PAES GALVAO DE MELO SANTOS", "ANA KATHARINE VASCONCELOS DE SOUSA",
    "ANA LIA TERCEIRO ALMEIDA", "ANA LUIZA ROCHA PICANCO", "ANA PAULA MILEO",
    "ANA VITORIA SALES DE OLIVEIRA FALCAO", "ANDRE LUIS SILVA DE SOUZA",
    "ANDRESSA HELENA ANTUNES DE OLIVEIRA", "ANNA KAROLINE CARDOSO AQUINO",
    "ANNE AGUIAR BARBOSA", "APARECIDO JOSE SILVA PLATERO", "ARTHUR LEMOS DE AGUIAR",
    "ARTUR SARAIVA DE ANDRADE", "BRENO LIMA PITTA PINHEIRO", "BRUNA DOS SANTOS CASTRO",
    "BRUNO MACHADO FORTES", "CAIO CESAR PINHEIRO GUERREIRO",
    "CARMEN GEORGIA REBOUCAS DE OLIVEIRA JORGE VIEIRA",
    "CICERO WELLINGTON BATISTA DO NASCIMENTO", "CLAYTON SANTANA LUZ",
    "CRISTIANE SANTOS DOS REIS", "CRISTIANO HOLANDA DA CUNHA",
    "DALILA CARLOS DE CASTRO", "DALILA DRISANA GOMES GONCALVES",
    "DALVA REGINA MARTINS ARGENTAO", "DANIEL ROCHA FERREIRA EUGENIO",
    "DANIEL SARAIVA PINHEIRO", "DANIELA MONDINO CANTORI", "DANIELLE DE QUEIROZ ALVES",
    "DAVID BARROSO PEREIRA", "DEBORA MARIA TEIXEIRA AUGUSTO LIMA",
    "DENISE ARAUJO SILVA DOS SANTOS", "DIEGO LIMA HOLANDA DOS SANTOS",
    "DIOGO DE SOUZA ROSA", "EDUARDO HENRIQUE AGUIAR", "ELANE GERMANO NUNES ALVES",
    "EMANUELLY ARAUJO VIEIRA", "EMERSON DE ALMEIDA MELO JUNIOR",
    "EMILLY TEIXEIRA DE SOUSA", "ERICA VALESCA DE OLIVEIRA FRAGA",
    "FABIO AUGUSTO SILVERIO SILVA", "FABIOLA FARIAS IBIAPINA",
    "FELIPE AGUIAR DE NEGREIROS ANDRADE", "FELIPE BEZERRA RIBEIRO",
    "FELIPE CARVALHO CARNEIRO", "FELIPE DE ALBUQUERQUE BEZERRA",
    "FERNANDA LIMA ALVES", "FERNANDA SOARES RODRIGUES", "FERNANDO GARCIA",
    "FERRUCIO ALISON ALCANTARA AMORIM", "FLAVIA PESSOA MONTEIRO",
    "FRAN COSTA DE CASTRO", "FRANCISCA EDILANDIA FREITAS ARAUJO",
    "FRANCISCA THAYSSE LIMA COSTA", "FRANCISCO ROBERTO DE MATOS",
    "GABRIEL CORREA FURTADO", "GISELE PEREIRA FONTELES",
    "HELIDA ZEDNIK RODRIGUES LIMA", "HELOISA GOMES REBOUCAS", "IAGO AMARAL REIS",
    "IGOR RABELO MAGALHAES", "IPRAZOS", "IRANEIVA ROCHA QUIRINO BARROS",
    "IURY ALVES VIEIRA DE SOUSA", "JAMILE DE GOIS RODRIGUES AMORIM",
    "JANAINA ELIAS CHIARADIA", "JANE DIANE DE RAMOS NUNES GONCALVES",
    "JEAN VICTOR NUNES SARAIVA", "JOAO CARNEIRO MELLO MOREIRA",
    "JOAO MARCOS DE ABREU TEIXEIRA", "JOHANN DANIEL DE OLIVEIRA INOCENCIO",
    "JOSE DE ARIMATEIA GORDIANO OLIVEIRA BARBOSA", "JOSE ZITO RABELO NETO",
    "JOSEPH MICAIAS OLIVEIRA DE CASTRO", "JULHIERME ALEX ZANAQUI",
    "JULIANA MARA VASCONCELOS ANDRADE", "JULIANA OSELAME MACEDO ZANAQUI",
    "JULIANA VECCHI DA SILVA", "KAMILA BARBOSA OLIVEIRA", "KASSIA FAUSTINO COELHO",
    "LAURO LINHARES LEITE", "LAYLA EVELYN NASCIMENTO PINHEIRO",
    "LEANDRO SARUBBI DE CARVALHO ROCHA", "LETICIA DE ALCANTARA MENDES",
    "LETICIA LEMOS BARRETO", "LEVY MOTA DE OLIVEIRA", "LUAN PEDRO CIANFARANI",
    "LUANA PRESTES DE SOUSA MELO", "LUCAS ANDRADE NOBREGA", "LUCIANO PEROBA FILHO",
    "LUIZ GUILHERME GONCALVES GIRAO", "MACSIMUS WALESKO DE CASTRO DUARTE",
    "MANOEL VICTOR BACALHAU", "MARCELO LIMA ARRAIS", "MARIA EDUARDA SANTOS ARAUJO",
    "MARIA LAURA MELO ALMEIDA", "MARIA THAIS RODRIGUES DA COSTA",
    "MARIANA ALMEIDA DE SOUZA", "MARIANA FASANARO DE CARVALHO",
    "MARIANA SILVA DE MORAIS", "MARILIA BARROSO WALRAVEN CUNHA",
    "MARINA NOBRE SIMAO", "MARIO FONSECA GOMES FERREIRA",
    "MATHEUS ARRUDA ALBUQUERQUE", "MATHEUS HAYASAKI",
    "MONICA MARIA QUEIROZ DA SILVA", "MONICA SANTOS MARTINS",
    "OYSTR - ROBO", "PEDRO ELIAS STELMACHUK COSTA",
    "PEDRO HENRIQUE ROOSEVELT GOES BARBOSA", "PHAMELLA SALES DE SOUZA COSTA",
    "RAFAEL PEREIRA DE SOUZA", "RAFAELA JOZIA HOLANDA",
    "RAUL MATIAS DA SILVA PADRAO", "RAYANNE LARI SOUSA ALMEIDA",
    "REBECA RODRIGUES DO NASCIMENTO", "REBECCA ARAUJO ROSA",
    "REGINALDO DE BRITO OLIVEIRA JUNIOR", "RENAN PEREIRA DA SILVA",
    "RENAN SINHORINI FUSATO", "ROBERTA FURTADO DE ARRAES ALENCAR E CASTRO",
    "RODRIGO RIBEIRO ANTUNES QUARIGUASI", "ROGERIO SCARABEL BARBOSA",
    "RUAN PEREIRA DO NASCIMENTO", "RUBIANA APARECIDA BARBIERI",
    "SAMARA DE MOURA FERREIRA", "SAMYA MONTEIRO PAMPLONA DE OLIVEIRA",
    "SANDRA MARA DO NASCIMENTO", "SARAH CHRISTINE ROCHA LOBAO",
    "SCHEILA DE PAULA CORDEIRO", "SUZANA MARIA LIMA BARROSO FELIX",
    "TAMIRIS CAMELO MELO LINO", "TATIANA HAUBERT", "THALYTA MARIA TORQUATO VITOR",
    "THIAGO GURGEL FREIRE LEITE", "VICTOR MARTINS BARBOSA",
    "VICTORIA FERREIRA ALONSO", "VITORIA CAVALCANTE DOS SANTOS",
    "VITORIA DA SILVA FREITAS", "WALESSA DIOGENES PEIXOTO DE ALENCAR",
    "WANDERLUCY CORREIA DE ALMEIDA", "WELLINGTON PEREIRA DA ROCHA FILHO",
    "WELLINGTON RUBENS", "WESLLEY MACEDO DE OLIVEIRA", "WVENDEL SENA OLIVEIRA",
    "YASMIM GORDIANO BARBOSA",
]}


# ─── UTILITÁRIOS ──────────────────────────────────────────────────────────────

def _encontrar_coluna(df, candidatos):
    mapa = {normalizar(c): c for c in df.columns}
    for c in candidatos:
        k = normalizar(c)
        if k in mapa:
            return mapa[k]
    return None


def _carregar_df(input_data):
    if isinstance(input_data, bytes):
        input_data = io.BytesIO(input_data)
    try:
        return pd.read_excel(input_data)
    except Exception:
        input_data.seek(0)
        return pd.read_excel(input_data, engine='openpyxl')


def _extrair_data(df, filename=""):
    col = _encontrar_coluna(df, ['Data cadastro', 'Data Cadastro', 'Data', 'data'])
    if col:
        try:
            datas = pd.to_datetime(df[col], dayfirst=True, errors='coerce').dropna()
            if len(datas):
                d = datas.mode()[0].date()
                return d, d.strftime('%d/%m/%Y')
        except Exception:
            pass
    m = re.search(r'(\d{2})[\.\-/](\d{2})[\.\-/](\d{4})', filename)
    if m:
        try:
            d = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            return d, d.strftime('%d/%m/%Y')
        except Exception:
            pass
    today = date.today()
    return today, today.strftime('%d/%m/%Y')


def _dia_semana_norm(d):
    return ['segunda', 'terca', 'quarta', 'quinta', 'sexta', 'sabado', 'domingo'][d.weekday()]


def _dia_semana_pt(d):
    return ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira',
            'Sexta-feira', 'Sábado', 'Domingo'][d.weekday()]


def classificar_tribunal(cnj):
    if not cnj or pd.isna(cnj):
        return 'OUTROS'
    partes = str(cnj).strip().split('.')
    if len(partes) >= 3:
        j = partes[2].strip()
        if j == '5':
            return 'TRT'
        elif j == '4':
            return 'TRF'
        elif j == '8':
            return 'TJ'
        elif j in ('1', '3'):
            return 'STF/STJ'
    return 'OUTROS'


# ─── CÁLCULO DE COTAS ─────────────────────────────────────────────────────────

def _calcular_cotas(n_total, dia_semana):
    """
    Retorna dict {analista: cota_inteira}.
    Regras:
      - Alanis: 50% da cota; excluída nas terças
      - Anna Júlia: 50% da cota; excluída nas quintas
      - Tatiana: incluída apenas se cota base > 40
      - Sobras (após floor) → Vanessa, depois Paloma, Bárbara, Lara, Tatiana
    """
    alanis_ativa = dia_semana != 'terca'
    aj_ativa = dia_semana != 'quinta'

    full_base = ['VANESSA', 'PALOMA', 'BARBARA', 'LARA']
    half_base = []
    if aj_ativa:
        half_base.append('ANNA JULIA')
    if alanis_ativa:
        half_base.append('ALANIS')

    div_sem_tatiana = len(full_base) + len(half_base) * 0.5
    q_sem_tatiana = n_total / div_sem_tatiana if div_sem_tatiana else 0

    tatiana_ativa = q_sem_tatiana > 40
    full = full_base + (['TATIANA'] if tatiana_ativa else [])
    half = half_base[:]

    div_final = len(full) + len(half) * 0.5
    q = int(n_total / div_final) if div_final else 0
    q_half = q // 2

    distributed = len(full) * q + len(half) * q_half
    leftover = n_total - distributed

    cotas = {a: q for a in full}
    for a in half:
        cotas[a] = q_half

    # Distribui sobras para analistas full na ordem
    for a in full:
        if leftover <= 0:
            break
        cotas[a] += 1
        leftover -= 1

    return cotas


# ─── DISTRIBUIÇÃO DE LINHAS ───────────────────────────────────────────────────

def _distribuir(df, cotas, col_cnj, col_natureza, col_cliente):
    alloc = {a: [] for a in cotas}
    pool = list(df.index)

    # 1. GPM → Bárbara
    if col_cliente and 'BARBARA' in cotas:
        gpm_mask = {i: normalizar(str(df.at[i, col_cliente])).startswith('GPM') for i in pool}
        gpm_rows = [i for i in pool if gpm_mask.get(i)]
        take = gpm_rows[:cotas['BARBARA']]
        alloc['BARBARA'].extend(take)
        pool = [i for i in pool if i not in set(take)]

    # 2. Não-trabalhista → Anna Júlia (até sua cota), depois Paloma
    if col_natureza:
        trab_norms = {normalizar('Trabalhista'), normalizar('TRABALHISTA'), 'TRABALHISTA'}
        non_labor_set = {
            i for i in pool
            if normalizar(str(df.at[i, col_natureza])) not in trab_norms
            and str(df.at[i, col_natureza]).strip() != ''
        }
        non_labor = [i for i in pool if i in non_labor_set]

        if 'ANNA JULIA' in cotas:
            restante_aj = cotas['ANNA JULIA'] - len(alloc['ANNA JULIA'])
            take_aj = non_labor[:restante_aj]
            alloc['ANNA JULIA'].extend(take_aj)
            non_labor = non_labor[restante_aj:]
            pool = [i for i in pool if i not in set(take_aj)]

        if 'PALOMA' in cotas:
            restante_p = cotas['PALOMA'] - len(alloc['PALOMA'])
            take_p = non_labor[:restante_p]
            alloc['PALOMA'].extend(take_p)
            pool = [i for i in pool if i not in set(take_p)]

    # 3. Restante → preenche cotas na ordem
    ordem = ['VANESSA', 'PALOMA', 'BARBARA', 'LARA', 'TATIANA', 'ANNA JULIA', 'ALANIS']
    for a in ordem:
        if a not in cotas or not pool:
            continue
        needed = cotas[a] - len(alloc[a])
        if needed > 0:
            take = pool[:needed]
            alloc[a].extend(take)
            pool = pool[needed:]

    # Sobras → Vanessa
    if pool and 'VANESSA' in alloc:
        alloc['VANESSA'].extend(pool)

    return alloc


# ─── ESTILOS ──────────────────────────────────────────────────────────────────

def _fill(hex_color):
    return PatternFill(fill_type='solid', start_color=hex_color, end_color=hex_color)

FILL_DARK_BLUE       = _fill('1F4E79')
FILL_MED_BLUE        = _fill('2E75B6')
FILL_LIGHT_BLUE      = _fill('D6E4F0')
FILL_ORANGE_INATIVO  = _fill('FCE4D6')   # laranja claro — responsáveis inativos
FILL_ORANGE_TRIB     = _fill('E26B0A')   # laranja escuro — tribunal minoritário
FILL_GRAY       = _fill('F2F2F2')

FONT_WHITE_BOLD = Font(name='Arial', bold=True, color='FFFFFF', size=10)
FONT_BOLD       = Font(name='Arial', bold=True, size=10)
FONT_NORMAL     = Font(name='Arial', size=10)
FONT_SMALL      = Font(name='Arial', size=9)

def _thin():
    s = Side(style='thin', color='BFBFBF')
    return Border(left=s, right=s, top=s, bottom=s)

def _header_row(ws, row, ncols, fill=None, font=None):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill or FILL_DARK_BLUE
        cell.font = font or FONT_WHITE_BOLD
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = _thin()


# ─── ABA RESUMO ───────────────────────────────────────────────────────────────

def _build_resumo(ws, d_str, dia_nome, total_bruto, total_unico, cotas, alloc, dia_semana):
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 28

    # Cabeçalho geral
    ws.merge_cells('A1:C1')
    c = ws['A1']
    c.value = f"RELATÓRIO DE PUBLICAÇÕES — {d_str} ({dia_nome})"
    c.font = Font(name='Arial', bold=True, size=12, color='FFFFFF')
    c.fill = FILL_DARK_BLUE
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    # Métricas
    metricas = [
        ("Total bruto", total_bruto),
        ("Duplicatas removidas", total_bruto - total_unico),
        ("Total único", total_unico),
    ]
    for i, (label, val) in enumerate(metricas, start=2):
        ws.cell(row=i, column=1).value = label
        ws.cell(row=i, column=1).font = FONT_BOLD
        ws.cell(row=i, column=1).fill = FILL_GRAY
        ws.cell(row=i, column=1).border = _thin()
        ws.cell(row=i, column=2).value = val
        ws.cell(row=i, column=2).font = FONT_NORMAL
        ws.cell(row=i, column=2).border = _thin()

    # Título da tabela de distribuição
    row = 6
    ws.cell(row=row, column=1).value = "ANALISTA"
    ws.cell(row=row, column=2).value = "PUBLICAÇÕES"
    ws.cell(row=row, column=3).value = "OBSERVAÇÃO"
    _header_row(ws, row, 3, FILL_MED_BLUE)

    ANALISTAS_ORDEM = ['VANESSA', 'PALOMA', 'BARBARA', 'LARA', 'ANNA JULIA', 'ALANIS', 'TATIANA']
    NOMES_EXIB = {
        'VANESSA': 'VANESSA', 'PALOMA': 'PALOMA', 'BARBARA': 'BÁRBARA',
        'LARA': 'LARA', 'ANNA JULIA': 'ANNA JÚLIA', 'ALANIS': 'ALANIS', 'TATIANA': 'TATIANA'
    }

    row = 7
    for a in ANALISTAS_ORDEM:
        obs = ""
        if a == 'ALANIS' and dia_semana == 'terca':
            qtd, obs = "—", "⛔ Fora (terça-feira)"
        elif a == 'ANNA JULIA' and dia_semana == 'quinta':
            qtd, obs = "—", "⛔ Fora (quinta-feira)"
        elif a not in cotas:
            qtd, obs = "—", "⚪ Fora (cota ≤ 40)" if a == 'TATIANA' else ""
        else:
            qtd = len(alloc.get(a, []))
            if a == 'TATIANA':
                obs = "✅ Incluída (cota > 40)"
            elif a in ('ALANIS', 'ANNA JULIA'):
                obs = "50% da cota"
            if a == 'BARBARA':
                obs = "Prioridade GPM"
            if a == 'PALOMA':
                obs = "Prioridade não-trabalhista"

        fill = FILL_LIGHT_BLUE if row % 2 == 0 else None
        for col, val in enumerate([NOMES_EXIB.get(a, a), qtd, obs], start=1):
            cell = ws.cell(row=row, column=col)
            cell.value = val
            cell.font = FONT_NORMAL
            cell.border = _thin()
            if fill:
                cell.fill = fill
        row += 1


# ─── ABA POR ANALISTA ─────────────────────────────────────────────────────────

def _build_analista_tab(ws, analista, rows_df, all_cols, col_cnj):
    if rows_df.empty:
        return

    # Tribunal de cada linha
    trib_col = col_cnj
    if trib_col and trib_col in rows_df.columns:
        tribunais = rows_df[trib_col].apply(classificar_tribunal)
    else:
        tribunais = pd.Series(['OUTROS'] * len(rows_df), index=rows_df.index)

    contagem = tribunais.value_counts()
    grupos_validos = [g for g in contagem.index if g != 'OUTROS']

    # Identifica minoritário (menor contagem dentre TRT/TRF/TJ)
    minoritario = None
    if len(grupos_validos) > 1:
        # Só destaca se há mais de um segmento
        minoritario = contagem[grupos_validos].idxmin()

    # Cabeçalhos
    cols = [c for c in all_cols if c in rows_df.columns]
    for ci, col in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=ci)
        cell.value = col
        cell.fill = FILL_DARK_BLUE
        cell.font = FONT_WHITE_BOLD
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = _thin()
        ws.column_dimensions[get_column_letter(ci)].width = max(12, len(str(col)) + 2)

    ws.row_dimensions[1].height = 22

    # Dados
    for ri, (idx, row_data) in enumerate(rows_df[cols].iterrows(), start=2):
        trib = tribunais.get(idx, 'OUTROS')
        is_minority = (trib == minoritario)
        for ci, val in enumerate(row_data, start=1):
            cell = ws.cell(row=ri, column=ci)
            cell.value = val
            cell.font = FONT_SMALL
            cell.border = _thin()
            if is_minority:
                cell.fill = FILL_ORANGE_TRIB
                cell.font = Font(name='Arial', size=9, color='FFFFFF')
            elif ri % 2 == 0:
                cell.fill = FILL_GRAY

    # Legenda no rodapé
    if minoritario:
        leg_row = len(rows_df) + 3
        leg_cell = ws.cell(row=leg_row, column=1)
        leg_cell.value = (
            f"🟠 Destaque (laranja escuro): tribunal minoritário — "
            f"{minoritario} ({contagem.get(minoritario, 0)} publ.)"
        )
        leg_cell.font = Font(name='Arial', italic=True, size=9, color='833C00')

    # Auto-fit largura
    for col_cells in ws.columns:
        max_len = 0
        for cell in col_cells:
            try:
                max_len = max(max_len, len(str(cell.value or '')))
            except Exception:
                pass
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 40)


# ─── ABA POR COORDENADOR ──────────────────────────────────────────────────────

def _build_coordenador_tab(ws, df, col_resp, mapeamento, d_str):
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 16

    # Cabeçalho
    ws.merge_cells('A1:B1')
    c = ws['A1']
    c.value = f"TOTAL POR COORDENADOR — {d_str}"
    c.font = Font(name='Arial', bold=True, size=11, color='FFFFFF')
    c.fill = FILL_DARK_BLUE
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 24

    ws.cell(row=2, column=1).value = "COORDENADOR"
    ws.cell(row=2, column=2).value = "PUBLICAÇÕES"
    _header_row(ws, 2, 2, FILL_MED_BLUE)

    # Contar por coordenador
    contagem = {}
    if col_resp:
        for val in df[col_resp].dropna():
            coord = mapeamento.get(normalizar(str(val)), "NÃO MAPEADO")
            contagem[coord] = contagem.get(coord, 0) + 1

    dados = sorted(contagem.items(), key=lambda x: -x[1])

    for i, (coord, qtd) in enumerate(dados, start=3):
        fill = FILL_GRAY if i % 2 == 0 else None
        for col, val in enumerate([coord, qtd], start=1):
            cell = ws.cell(row=i, column=col)
            cell.value = val
            cell.font = FONT_NORMAL
            cell.border = _thin()
            if fill:
                cell.fill = fill

    # Gráfico de barras
    if dados:
        n = len(dados)
        chart = BarChart()
        chart.type = "bar"
        chart.title = "Publicações por Coordenador"
        chart.y_axis.title = "Coordenador"
        chart.x_axis.title = "Quantidade"
        chart.style = 10
        chart.width = 20
        chart.height = max(8, n * 0.7)

        data_ref = Reference(ws, min_col=2, min_row=2, max_row=2 + n)
        cats_ref = Reference(ws, min_col=1, min_row=3, max_row=2 + n)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart.series[0].graphicalProperties.solidFill = "2E75B6"

        ws.add_chart(chart, f"D3")


# ─── ABA POR RESPONSÁVEL ──────────────────────────────────────────────────────

def _build_responsavel_tab(ws, df, col_resp, mapeamento, d_str):
    ws.column_dimensions['A'].width = 44
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 12

    # Cabeçalho
    ws.merge_cells('A1:D1')
    c = ws['A1']
    c.value = f"TOTAL POR RESPONSÁVEL — {d_str}"
    c.font = Font(name='Arial', bold=True, size=11, color='FFFFFF')
    c.fill = FILL_DARK_BLUE
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 24

    for col, titulo in enumerate(['RESPONSÁVEL', 'COORDENADOR', 'PUBLICAÇÕES', 'STATUS'], start=1):
        cell = ws.cell(row=2, column=col)
        cell.value = titulo
        cell.fill = FILL_MED_BLUE
        cell.font = FONT_WHITE_BOLD
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = _thin()

    # Contar por responsável
    contagem = {}
    if col_resp:
        for val in df[col_resp].dropna():
            nome = str(val).strip()
            contagem[nome] = contagem.get(nome, 0) + 1

    dados = sorted(contagem.items(), key=lambda x: -x[1])

    for i, (resp, qtd) in enumerate(dados, start=3):
        resp_norm = normalizar(resp)
        coord = mapeamento.get(resp_norm, "—")
        inativo = resp_norm in INATIVOS
        status = "⚠️ Inativo" if inativo else "Ativo"

        row_fill = FILL_ORANGE_INATIVO if inativo else (FILL_GRAY if i % 2 == 0 else None)

        for col, val in enumerate([resp, coord, qtd, status], start=1):
            cell = ws.cell(row=i, column=col)
            cell.value = val
            cell.font = FONT_NORMAL
            cell.border = _thin()
            if row_fill:
                cell.fill = row_fill

    # Gráfico (top 20)
    n_chart = min(20, len(dados))
    if n_chart > 0:
        chart = BarChart()
        chart.type = "bar"
        chart.title = "Top Responsáveis por Publicações"
        chart.y_axis.title = "Responsável"
        chart.x_axis.title = "Quantidade"
        chart.style = 10
        chart.width = 22
        chart.height = max(8, n_chart * 0.65)

        data_ref = Reference(ws, min_col=3, min_row=2, max_row=2 + n_chart)
        cats_ref = Reference(ws, min_col=1, min_row=3, max_row=2 + n_chart)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart.series[0].graphicalProperties.solidFill = "2E75B6"

        ws.add_chart(chart, f"F3")


# ─── PRE-CHECK (API pública) ──────────────────────────────────────────────────

def pre_check(input_data, filename="", extra_mappings=None):
    """
    Verifica o arquivo sem gerar o relatório.
    Retorna dict com: data_str, dia_semana, dia_semana_nome,
                      total_bruto, total_unico, duplicatas,
                      unmapped (lista), inativos_encontrados (lista)
    """
    if extra_mappings is None:
        extra_mappings = {}
    mapeamento = {**COORDENADORES_MAPEADOS, **{normalizar(k): v for k, v in extra_mappings.items()}}

    df = _carregar_df(input_data)
    col_cnj  = _encontrar_coluna(df, ['Número de CNJ', 'Numero de CNJ', 'CNJ', 'numero cnj'])
    col_resp = _encontrar_coluna(df, ['Responsável', 'Responsavel', 'RESPONSAVEL'])

    total_bruto = len(df)
    if col_cnj:
        df = df.drop_duplicates(subset=[col_cnj])
    total_unico = len(df)

    d, d_str = _extrair_data(df, filename)

    unmapped, inativos = [], []
    if col_resp:
        for val in df[col_resp].dropna().unique():
            norm = normalizar(str(val))
            if norm in INATIVOS:
                inativos.append(str(val).strip())
            if norm not in mapeamento:
                unmapped.append(str(val).strip())

    return {
        "data": d,
        "data_str": d_str,
        "dia_semana": _dia_semana_norm(d),
        "dia_semana_nome": _dia_semana_pt(d),
        "total_bruto": total_bruto,
        "total_unico": total_unico,
        "duplicatas": total_bruto - total_unico,
        "unmapped": sorted(set(unmapped)),
        "inativos_encontrados": sorted(set(inativos)),
    }


# ─── GERAR RELATÓRIO (API pública) ────────────────────────────────────────────

def gerar_relatorio(input_data, filename="", extra_mappings=None):
    """
    Gera o Excel com todas as abas.
    Retorna (output_bytes: bytes, resumo: dict)
    """
    if extra_mappings is None:
        extra_mappings = {}
    mapeamento = {**COORDENADORES_MAPEADOS, **{normalizar(k): v for k, v in extra_mappings.items()}}

    if isinstance(input_data, bytes):
        raw_bytes = input_data
    else:
        raw_bytes = input_data.read()

    df = _carregar_df(raw_bytes)

    col_cnj      = _encontrar_coluna(df, ['Número de CNJ', 'Numero de CNJ', 'CNJ', 'numero cnj'])
    col_natureza = _encontrar_coluna(df, ['Natureza', 'NATUREZA', 'natureza'])
    col_cliente  = _encontrar_coluna(df, ['Cliente', 'CLIENTE', 'cliente'])
    col_resp     = _encontrar_coluna(df, ['Responsável', 'Responsavel', 'RESPONSAVEL'])

    total_bruto = len(df)
    if col_cnj:
        df = df.drop_duplicates(subset=[col_cnj])
    df = df.reset_index(drop=True)
    total_unico = len(df)

    d, d_str = _extrair_data(df, filename)
    dia_semana      = _dia_semana_norm(d)
    dia_semana_nome = _dia_semana_pt(d)

    cotas = _calcular_cotas(total_unico, dia_semana)
    alloc = _distribuir(df, cotas, col_cnj, col_natureza, col_cliente)

    wb = Workbook()
    wb.remove(wb.active)

    # RESUMO
    ws_resumo = wb.create_sheet('RESUMO')
    _build_resumo(ws_resumo, d_str, dia_semana_nome, total_bruto, total_unico, cotas, alloc, dia_semana)

    # Abas individuais
    ORDEM = ['VANESSA', 'PALOMA', 'BARBARA', 'LARA', 'ANNA JULIA', 'ALANIS', 'TATIANA']
    NOMES_ABA = {
        'VANESSA': 'VANESSA', 'PALOMA': 'PALOMA', 'BARBARA': 'BÁRBARA',
        'LARA': 'LARA', 'ANNA JULIA': 'ANNA JÚLIA', 'ALANIS': 'ALANIS', 'TATIANA': 'TATIANA'
    }
    for a in ORDEM:
        if a in alloc and alloc[a]:
            rows_df = df.loc[alloc[a]]
            ws = wb.create_sheet(NOMES_ABA.get(a, a))
            _build_analista_tab(ws, a, rows_df, df.columns.tolist(), col_cnj)

    # POR COORDENADOR
    ws_coord = wb.create_sheet('POR COORDENADOR')
    _build_coordenador_tab(ws_coord, df, col_resp, mapeamento, d_str)

    # POR RESPONSÁVEL
    ws_resp = wb.create_sheet('POR RESPONSAVEL')
    _build_responsavel_tab(ws_resp, df, col_resp, mapeamento, d_str)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    resumo = {
        "data_str": d_str,
        "dia_semana_nome": dia_semana_nome,
        "total_bruto": total_bruto,
        "total_unico": total_unico,
        "duplicatas": total_bruto - total_unico,
        "cotas": cotas,
        "tatiana_incluida": 'TATIANA' in cotas,
        "alanis_ativa": dia_semana != 'terca',
        "anna_julia_ativa": dia_semana != 'quinta',
        "alloc_counts": {a: len(v) for a, v in alloc.items()},
    }
    return output.getvalue(), resumo
