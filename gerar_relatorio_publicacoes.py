"""
gerar_relatorio_publicacoes.py — LegalOne · Publicações Diárias
Versão atualizada com especificações completas de formatação e equipe.
"""

import io, re, unicodedata
from datetime import date
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

# ─── NORMALIZAÇÃO ─────────────────────────────────────────────────────────────

def normalizar(s):
    if not s or isinstance(s, float): return ""
    s = str(s).strip()
    # Substituir qualquer variante de espaço (não-quebrável, zero-width, etc.) por espaço normal
    s = re.sub(r'[ ​‌‍ -   　﻿]', ' ', s)
    s = s.upper()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    s = re.sub(r' +', ' ', s).strip()  # colapsar múltiplos espaços
    return s

# ─── DADOS FIXOS ──────────────────────────────────────────────────────────────

COORDENADORES_MAPEADOS = {
    normalizar("RODRIGO RIBEIRO ANTUNES QUARIGUASI"):      "GABRIEL GIORGIO CICCHELERO",
    normalizar("CAMILLA GOES BARBOSA"):                    "CAMILLA GOES BARBOSA",
    normalizar("SUZANA MARIA CAMPOS MARANHAO DE LIMA"):    "SUZANA MARIA CAMPOS MARANHAO DE LIMA",
    normalizar("DALILA DRISANA GOMES GONCALVES"):          "GABRIEL GIORGIO CICCHELERO",
    normalizar("ANA VITORIA SALES DE OLIVEIRA FALCAO"):    "GABRIEL GIORGIO CICCHELERO",
    normalizar("RAFAEL CAVALCANTE BARSOSA"):               "GABRIEL GIORGIO CICCHELERO",
    normalizar("MATHEUS CAVALCANTI DE ARAUJO"):            "LUCIANE MODERNEL MENDES",
    normalizar("EMERSON DE ALMEIDA MELO JUNIOR"):          "NAYANDERSON LUAN MELLO PINHEIRO",
    normalizar("JOHANN DANIEL DE OLIVEIRA INOCENCIO"):     "SUZANA MARIA CAMPOS MARANHAO DE LIMA",
    normalizar("MARCELO LIMA ARRAIS"):                     "HELANZIA DE ARAUJO XAVIER WICHAMNN",
    normalizar("WELLINGTON PEREIRA DA ROCHA FILHO"):       "HELANZIA DE ARAUJO XAVIER WICHAMNN",
    normalizar("RENAN PEREIRA DA SILVA"):                  "HELANZIA DE ARAUJO XAVIER WICHAMNN",
    normalizar("JEAN VICTOR NUNES SARAIVA"):               "NAYANDERSON LUAN MELLO PINHEIRO",
    normalizar("YASMIM GORDIANO BARBOSA"):                 "MARCELLE LEITE RENTROIA",
    normalizar("ARTUR SARAIVA DE ANDRADE"):                "HELANZIA DE ARAUJO XAVIER WICHAMNN",
    normalizar("DALILA CARLOS DE CASTRO"):                 "GABRIEL GIORGIO CICCHELERO",
    normalizar("GABRIEL GIORGIO CICCHELERO"):              "GABRIEL GIORGIO CICCHELERO",
    normalizar("HELANZIA DE ARAUJO XAVIER WICHAMNN"):      "HELANZIA DE ARAUJO XAVIER WICHAMNN",
    normalizar("JENIFFER ROSA BARBOSA DE SALES"):          "JENIFFER ROSA BARBOSA DE SALES",
    normalizar("LUCIANE MODERNEL MENDES"):                 "LUCIANE MODERNEL MENDES",
    normalizar("NAYANDERSON LUAN MELLO PINHEIRO"):         "NAYANDERSON LUAN MELLO PINHEIRO",
    normalizar("YURI ALVES BARROS DOS SANTOS"):            "YURI ALVES BARROS DOS SANTOS",
    normalizar("MARCELLE LEITE RENTROIA"):                 "MARCELLE LEITE RENTROIA",
}

COORDENADORES_CONHECIDOS = [
    "CAMILLA GOES BARBOSA", "GABRIEL GIORGIO CICCHELERO",
    "HELANZIA DE ARAUJO XAVIER WICHAMNN", "JENIFFER ROSA BARBOSA DE SALES",
    "LUCIANE MODERNEL MENDES", "MARCELLE LEITE RENTROIA",
    "NAYANDERSON LUAN MELLO PINHEIRO", "SUZANA MARIA CAMPOS MARANHAO DE LIMA",
    "YURI ALVES BARROS DOS SANTOS",
]

INATIVOS = {normalizar(n) for n in [
    "ALEXANDRE SOUZA DA SILVA FILHO","ALICE RIBEIRO MELO CRISOSTOMO",
    "ALLY CHARLYS MUSSE MENDES FILHO","AMANDA BEZERRA DE MENEZES NETO",
    "AMANDA MOURA DOS SANTOS BRAGA","ANA CAMILA CIFONI DE VASCONCELOS BARROSO",
    "ANA CAROLINA PAES GALVAO DE MELO SANTOS","ANA KATHARINE VASCONCELOS DE SOUSA",
    "ANA LIA TERCEIRO ALMEIDA","ANA LUIZA ROCHA PICANCO","ANA PAULA MILEO",
    "ANA VITORIA SALES DE OLIVEIRA FALCAO","ANDRE LUIS SILVA DE SOUZA",
    "ANDRESSA HELENA ANTUNES DE OLIVEIRA","ANNA KAROLINE CARDOSO AQUINO",
    "ANNE AGUIAR BARBOSA","APARECIDO JOSE SILVA PLATERO","ARTHUR LEMOS DE AGUIAR",
    "ARTUR SARAIVA DE ANDRADE","BRENO LIMA PITTA PINHEIRO","BRUNA DOS SANTOS CASTRO",
    "BRUNO MACHADO FORTES","CAIO CESAR PINHEIRO GUERREIRO",
    "CARMEN GEORGIA REBOUCAS DE OLIVEIRA JORGE VIEIRA",
    "CICERO WELLINGTON BATISTA DO NASCIMENTO","CLAYTON SANTANA LUZ",
    "CRISTIANE SANTOS DOS REIS","CRISTIANO HOLANDA DA CUNHA",
    "DALILA CARLOS DE CASTRO","DALILA DRISANA GOMES GONCALVES",
    "DALVA REGINA MARTINS ARGENTAO","DANIEL ROCHA FERREIRA EUGENIO",
    "DANIEL SARAIVA PINHEIRO","DANIELA MONDINO CANTORI","DANIELLE DE QUEIROZ ALVES",
    "DAVID BARROSO PEREIRA","DEBORA MARIA TEIXEIRA AUGUSTO LIMA",
    "DENISE ARAUJO SILVA DOS SANTOS","DIEGO LIMA HOLANDA DOS SANTOS",
    "DIOGO DE SOUZA ROSA","EDUARDO HENRIQUE AGUIAR","ELANE GERMANO NUNES ALVES",
    "EMANUELLY ARAUJO VIEIRA","EMERSON DE ALMEIDA MELO JUNIOR",
    "EMILLY TEIXEIRA DE SOUSA","ERICA VALESCA DE OLIVEIRA FRAGA",
    "FABIO AUGUSTO SILVERIO SILVA","FABIOLA FARIAS IBIAPINA",
    "FELIPE AGUIAR DE NEGREIROS ANDRADE","FELIPE BEZERRA RIBEIRO",
    "FELIPE CARVALHO CARNEIRO","FELIPE DE ALBUQUERQUE BEZERRA",
    "FERNANDA LIMA ALVES","FERNANDA SOARES RODRIGUES","FERNANDO GARCIA",
    "FERRUCIO ALISON ALCANTARA AMORIM","FLAVIA PESSOA MONTEIRO",
    "FRAN COSTA DE CASTRO","FRANCISCA EDILANDIA FREITAS ARAUJO",
    "FRANCISCA THAYSSE LIMA COSTA","FRANCISCO ROBERTO DE MATOS",
    "GABRIEL CORREA FURTADO","GISELE PEREIRA FONTELES",
    "HELIDA ZEDNIK RODRIGUES LIMA","HELOISA GOMES REBOUCAS","IAGO AMARAL REIS",
    "IGOR RABELO MAGALHAES","IPRAZOS","IRANEIVA ROCHA QUIRINO BARROS",
    "IURY ALVES VIEIRA DE SOUSA","JAMILE DE GOIS RODRIGUES AMORIM",
    "JANAINA ELIAS CHIARADIA","JANE DIANE DE RAMOS NUNES GONCALVES",
    "JEAN VICTOR NUNES SARAIVA","JOAO CARNEIRO MELLO MOREIRA",
    "JOAO MARCOS DE ABREU TEIXEIRA","JOHANN DANIEL DE OLIVEIRA INOCENCIO",
    "JOSE DE ARIMATEIA GORDIANO OLIVEIRA BARBOSA","JOSE ZITO RABELO NETO",
    "JOSEPH MICAIAS OLIVEIRA DE CASTRO","JULHIERME ALEX ZANAQUI",
    "JULIANA MARA VASCONCELOS ANDRADE","JULIANA OSELAME MACEDO ZANAQUI",
    "JULIANA VECCHI DA SILVA","KAMILA BARBOSA OLIVEIRA","KASSIA FAUSTINO COELHO",
    "LAURO LINHARES LEITE","LAYLA EVELYN NASCIMENTO PINHEIRO",
    "LEANDRO SARUBBI DE CARVALHO ROCHA","LETICIA DE ALCANTARA MENDES",
    "LETICIA LEMOS BARRETO","LEVY MOTA DE OLIVEIRA","LUAN PEDRO CIANFARANI",
    "LUANA PRESTES DE SOUSA MELO","LUCAS ANDRADE NOBREGA","LUCIANO PEROBA FILHO",
    "LUIZ GUILHERME GONCALVES GIRAO","MACSIMUS WALESKO DE CASTRO DUARTE",
    "MANOEL VICTOR BACALHAU","MARCELO LIMA ARRAIS","MARIA EDUARDA SANTOS ARAUJO",
    "MARIA LAURA MELO ALMEIDA","MARIA THAIS RODRIGUES DA COSTA",
    "MARIANA ALMEIDA DE SOUZA","MARIANA FASANARO DE CARVALHO",
    "MARIANA SILVA DE MORAIS","MARILIA BARROSO WALRAVEN CUNHA",
    "MARINA NOBRE SIMAO","MARIO FONSECA GOMES FERREIRA",
    "MATHEUS ARRUDA ALBUQUERQUE","MATHEUS HAYASAKI",
    "MONICA MARIA QUEIROZ DA SILVA","MONICA SANTOS MARTINS",
    "OYSTR - ROBO","PEDRO ELIAS STELMACHUK COSTA",
    "PEDRO HENRIQUE ROOSEVELT GOES BARBOSA","PHAMELLA SALES DE SOUZA COSTA",
    "RAFAEL PEREIRA DE SOUZA","RAFAELA JOZIA HOLANDA",
    "RAUL MATIAS DA SILVA PADRAO","RAYANNE LARI SOUSA ALMEIDA",
    "REBECA RODRIGUES DO NASCIMENTO","REBECCA ARAUJO ROSA",
    "REGINALDO DE BRITO OLIVEIRA JUNIOR","RENAN PEREIRA DA SILVA",
    "RENAN SINHORINI FUSATO","ROBERTA FURTADO DE ARRAES ALENCAR E CASTRO",
    "RODRIGO RIBEIRO ANTUNES QUARIGUASI","ROGERIO SCARABEL BARBOSA",
    "RUAN PEREIRA DO NASCIMENTO","RUBIANA APARECIDA BARBIERI",
    "SAMARA DE MOURA FERREIRA","SAMYA MONTEIRO PAMPLONA DE OLIVEIRA",
    "SANDRA MARA DO NASCIMENTO","SARAH CHRISTINE ROCHA LOBAO",
    "SCHEILA DE PAULA CORDEIRO","SUZANA MARIA LIMA BARROSO FELIX",
    "TAMIRIS CAMELO MELO LINO","TATIANA HAUBERT","THALYTA MARIA TORQUATO VITOR",
    "THIAGO GURGEL FREIRE LEITE","VICTOR MARTINS BARBOSA",
    "VICTORIA FERREIRA ALONSO","VITORIA CAVALCANTE DOS SANTOS",
    "VITORIA DA SILVA FREITAS","WALESSA DIOGENES PEIXOTO DE ALENCAR",
    "WANDERLUCY CORREIA DE ALMEIDA","WELLINGTON PEREIRA DA ROCHA FILHO",
    "WELLINGTON RUBENS","WESLLEY MACEDO DE OLIVEIRA","WVENDEL SENA OLIVEIRA",
    "YASMIM GORDIANO BARBOSA",
]}

# ─── COLUNAS DE SAÍDA — abas individuais ──────────────────────────────────────
# (nome_exibição, largura, candidatos_busca)
OUTPUT_COLS = [
    ('Data cadastro', 20, ['Data cadastro','Data Cadastro','Data']),
    ('Pasta',         16, ['Pasta','PASTA']),
    ('Natureza',      14, ['Natureza','NATUREZA','natureza']),
    ('Responsável',   32, ['Responsável','Responsavel','RESPONSAVEL']),
    ('Status',        12, ['Status','STATUS','status']),
    ('Cliente',       30, ['Cliente','CLIENTE','cliente']),
    ('Número de CNJ', 28, ['Número de CNJ','Numero de CNJ','CNJ']),
]

# ─── CORES ────────────────────────────────────────────────────────────────────
def _fill(hex8): return PatternFill(fill_type='solid', start_color=hex8, end_color=hex8)
def _font(bold=False, size=9, color='FF000000', italic=False):
    return Font(name='Arial', bold=bold, size=size, color=color, italic=italic)
def _border():
    s = Side(style='thin', color='FFCCCCCC')
    return Border(left=s, right=s, top=s, bottom=s)
def _align(h='left', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

FILL_TITLE    = _fill('FF1F3864')   # título (fundo escuro)
FILL_HEADER   = _fill('FF2E4D8A')   # cabeçalho de coluna
FILL_TRIB_MIN = _fill('FFFF8C00')   # tribunal minoritário
FILL_INATIVO  = _fill('FFF4A460')   # responsável inativo
FILL_DATA     = _fill('FFF0F4FF')   # dados gerais fundo claro
FILL_WHITE    = _fill('FFFFFFFF')   # linha par alternada
FILL_GRAY     = _fill('FFEEEEEE')   # ausente / sem lote
FILL_RED_SPEC = _fill('FF8B0000')   # divisão especial

# Cores por analista (linhas ímpares nas abas individuais)
ANALISTA_FILL = {
    'VANESSA':    _fill('FFB3D9FF'),
    'PALOMA':     _fill('FFC8F7C5'),
    'BARBARA':    _fill('FFFFE4B5'),
    'LARA':       _fill('FFE8D5FF'),
    'ANNA JULIA': _fill('FFFFD6D6'),
    'ANA CECILIA':_fill('FFFFE4F4'),
    'ALANIS':     _fill('FFD6F0FF'),
    'TATIANA':    _fill('FFFFF0C0'),
}

# Gradiente POR COORDENADOR (9 posições)
GRAD_COORD = [
    'FF1F3864','FF2E4D8A','FF3A5FA8','FF4A72C4','FF5A85D8',
    'FF6B98EC','FF7BABFF','FF8CBEFD','FF9DCFFE',
]

def _gradient_colors(n, start='FF1F3864', end='FFA2D4FF'):
    """Interpola n cores entre start e end (ARGB 8 hex)."""
    def h2rgb(h): h=h[2:]; return int(h[:2],16),int(h[2:4],16),int(h[4:],16)
    def rgb2h(r,g,b): return f'FF{r:02X}{g:02X}{b:02X}'
    if n<=1: return [start]
    sr,sg,sb = h2rgb(start); er,eg,eb = h2rgb(end)
    return [rgb2h(int(sr+(er-sr)*i/(n-1)),int(sg+(eg-sg)*i/(n-1)),int(sb+(eb-sb)*i/(n-1)))
            for i in range(n)]

# ─── UTILITÁRIOS ──────────────────────────────────────────────────────────────

def _encontrar_coluna(df, candidatos):
    mapa = {normalizar(c): c for c in df.columns}
    for c in candidatos:
        k = normalizar(c)
        if k in mapa: return mapa[k]
    return None

def _carregar_df(data):
    if isinstance(data, bytes): data = io.BytesIO(data)
    try: return pd.read_excel(data)
    except: data.seek(0); return pd.read_excel(data, engine='openpyxl')

def _extrair_data(df, filename=""):
    # 1. Nome do arquivo (mais confiável — Data cadastro pode ser D-1)
    m = re.search(r'(\d{2})[_\.\-/](\d{2})[_\.\-/](\d{2,4})', filename)
    if m:
        try:
            y = int(m.group(3)); y = y+2000 if y<100 else y
            d = date(y, int(m.group(2)), int(m.group(1)))
            return d, d.strftime('%d/%m/%Y')
        except: pass
    # 2. Coluna Data cadastro
    col = _encontrar_coluna(df, ['Data cadastro','Data Cadastro','Data'])
    if col:
        try:
            ds = pd.to_datetime(df[col], errors='coerce').dropna()
            if len(ds):
                d = ds.mode()[0].date()
                return d, d.strftime('%d/%m/%Y')
        except: pass
    today = date.today()
    return today, today.strftime('%d/%m/%Y')

def _dia_semana_norm(d):
    return ['segunda','terca','quarta','quinta','sexta','sabado','domingo'][d.weekday()]

def _dia_semana_pt(d):
    return ['Segunda-feira','Terça-feira','Quarta-feira','Quinta-feira',
            'Sexta-feira','Sábado','Domingo'][d.weekday()]

def classificar_tribunal(cnj):
    if not cnj or pd.isna(cnj): return 'OUTROS'
    p = str(cnj).strip().split('.')
    if len(p)>=3:
        j = p[2].strip()
        if j=='5': return 'TRT'
        if j=='4': return 'TRF'
        if j=='8': return 'TJ'
        if j in('1','3'): return 'STF/STJ'
    return 'OUTROS'

def _set_cell(cell, value, font=None, fill=None, align=None, border=True):
    cell.value = value
    if font:   cell.font = font
    if fill:   cell.fill = fill
    if align:  cell.alignment = align
    if border: cell.border = _border()

# ─── CÁLCULO DE COTAS ─────────────────────────────────────────────────────────

def _calcular_cotas(n_total, dia_semana):
    alanis_ativa = dia_semana != 'terca'
    aj_ativa     = dia_semana != 'quinta'

    full_base = ['VANESSA','PALOMA','BARBARA','LARA']
    half_base = ['ANA CECILIA']           # sempre ativa
    if aj_ativa:    half_base.append('ANNA JULIA')
    if alanis_ativa:half_base.append('ALANIS')

    div_st = len(full_base) + len(half_base) * 0.5
    q_st   = n_total / div_st if div_st else 0
    tatiana = q_st > 40

    full = full_base + (['TATIANA'] if tatiana else [])
    half = half_base[:]

    div = len(full) + len(half) * 0.5
    q   = int(n_total / div) if div else 0
    qh  = q // 2

    cotas = {a: q for a in full}
    for a in half: cotas[a] = qh

    leftover = n_total - (len(full)*q + len(half)*qh)
    for a in full:
        if leftover <= 0: break
        cotas[a] += 1; leftover -= 1

    return cotas

# ─── DISTRIBUIÇÃO ─────────────────────────────────────────────────────────────

def _distribuir(df, cotas, col_cnj, col_natureza, col_cliente):
    alloc = {a: [] for a in cotas}
    pool  = list(df.index)

    # 1. GPM → Bárbara
    if col_cliente and 'BARBARA' in cotas:
        gpm = [i for i in pool if normalizar(str(df.at[i,col_cliente])).startswith('GPM')]
        take = gpm[:cotas['BARBARA']]
        alloc['BARBARA'].extend(take)
        pool = [i for i in pool if i not in set(take)]

    # 2. Não-trabalhista → Anna Júlia, Ana Cecília, Paloma (nessa ordem)
    if col_natureza:
        nlt = {i for i in pool
               if normalizar(str(df.at[i,col_natureza])) not in {'TRABALHISTA'}
               and str(df.at[i,col_natureza]).strip() != ''}
        nl = [i for i in pool if i in nlt]

        for analista in ['ANNA JULIA','ANA CECILIA','PALOMA']:
            if analista not in cotas: continue
            need = cotas[analista] - len(alloc[analista])
            take = nl[:need]
            alloc[analista].extend(take)
            nl   = nl[need:]
            pool = [i for i in pool if i not in set(take)]

    # 3. Restante → preenche cotas na ordem
    for a in ['VANESSA','PALOMA','BARBARA','LARA','TATIANA','ANNA JULIA','ANA CECILIA','ALANIS']:
        if a not in cotas or not pool: continue
        need = cotas[a] - len(alloc[a])
        if need > 0:
            take = pool[:need]
            alloc[a].extend(take)
            pool = pool[need:]

    # Sobras → Vanessa
    if pool and 'VANESSA' in alloc:
        alloc['VANESSA'].extend(pool)

    return alloc

# ─── ABA RESUMO ───────────────────────────────────────────────────────────────

def _build_resumo(ws, d_str, dia_nome, total_bruto, total_unico, cotas, alloc,
                  dia_semana, df, col_natureza, col_cliente, divisao_especial=False):

    NOMES_EXIB = {'VANESSA':'VANESSA','PALOMA':'PALOMA','BARBARA':'BÁRBARA',
                  'LARA':'LARA','ANNA JULIA':'ANNA JÚLIA','ANA CECILIA':'ANA CECÍLIA',
                  'ALANIS':'ALANIS','TATIANA':'TATIANA'}

    # Larguras
    for col,w in [(1,20),(2,14),(3,12),(4,38),(5,16)]:
        ws.column_dimensions[get_column_letter(col)].width = w

    # Título
    titulo = f"{'⚠️ DIVISÃO ESPECIAL — ' if divisao_especial else ''}RELATÓRIO DE PUBLICAÇÕES — {d_str} ({dia_nome})"
    ws.merge_cells('A1:E1')
    c = ws['A1']
    _set_cell(c, titulo,
              font=_font(bold=True, size=14, color='FFFFFFFF'),
              fill=FILL_RED_SPEC if divisao_especial else FILL_TITLE,
              align=_align('center','center'))
    ws.row_dimensions[1].height = 32

    # Estatísticas
    n_gpm = sum(1 for i in df.index
                if col_cliente and normalizar(str(df.at[i,col_cliente])).startswith('GPM')) if col_cliente else 0
    n_ntrab = sum(1 for i in df.index
                  if col_natureza and normalizar(str(df.at[i,col_natureza])) not in {'TRABALHISTA'}
                  and str(df.at[i,col_natureza]).strip()) if col_natureza else 0

    stats = [
        ('Total bruto', total_bruto),   ('Publicações GPM', n_gpm),
        ('Duplicatas removidas', total_bruto-total_unico), ('Não-trabalhistas', n_ntrab),
        ('Total único', total_unico),
    ]
    for i,(lbl,val) in enumerate(stats):
        r = 2 + i//2; c_col = 1 if i%2==0 else 3
        _set_cell(ws.cell(row=r,column=c_col), lbl,
                  font=_font(bold=True,size=9), fill=FILL_DATA, align=_align('left','center'))
        _set_cell(ws.cell(row=r,column=c_col+1), val,
                  font=_font(size=9), align=_align('center','center'))

    # Distribuição por natureza
    r_nat = 5
    if col_natureza:
        nat_counts = df[col_natureza].value_counts()
        _set_cell(ws.cell(row=r_nat,column=1), 'DISTRIBUIÇÃO POR NATUREZA',
                  font=_font(bold=True,size=9,color='FFFFFFFF'), fill=FILL_HEADER,
                  align=_align('left','center'))
        _set_cell(ws.cell(row=r_nat,column=2), 'TOTAL',
                  font=_font(bold=True,size=9,color='FFFFFFFF'), fill=FILL_HEADER,
                  align=_align('center','center'))
        _set_cell(ws.cell(row=r_nat,column=3), '%',
                  font=_font(bold=True,size=9,color='FFFFFFFF'), fill=FILL_HEADER,
                  align=_align('center','center'))
        for j,(nat,cnt) in enumerate(nat_counts.items(),start=1):
            rr = r_nat + j
            fill = FILL_DATA if j%2==1 else FILL_WHITE
            _set_cell(ws.cell(row=rr,column=1), str(nat), font=_font(size=9), fill=fill, align=_align())
            _set_cell(ws.cell(row=rr,column=2), int(cnt), font=_font(size=9), fill=fill, align=_align('center'))
            pct = f"{cnt/total_unico*100:.1f}%" if total_unico else "0%"
            _set_cell(ws.cell(row=rr,column=3), pct, font=_font(size=9), fill=fill, align=_align('center'))
        r_tab = r_nat + len(nat_counts) + 2
    else:
        r_tab = 7

    # Tabela de analistas
    for col_h,lbl in enumerate(['ANALISTA','PUBLICAÇÕES','%','OBSERVAÇÃO','STATUS'],start=1):
        _set_cell(ws.cell(row=r_tab,column=col_h), lbl,
                  font=_font(bold=True,size=10,color='FFFFFFFF'), fill=FILL_HEADER,
                  align=_align('center','center'))
    ws.row_dimensions[r_tab].height = 22

    ORDEM = ['VANESSA','PALOMA','BARBARA','LARA','ANNA JULIA','ANA CECILIA','ALANIS','TATIANA']
    for i,a in enumerate(ORDEM):
        r = r_tab + 1 + i
        fill = ANALISTA_FILL.get(a, FILL_GRAY)
        nome = NOMES_EXIB.get(a, a)
        # Determinar qtd e obs
        if a=='ALANIS' and dia_semana=='terca':
            qtd,pct,obs,status = '—','—','','⛔ Fora (terça-feira)'
            fill = FILL_GRAY
        elif a=='ANNA JULIA' and dia_semana=='quinta':
            qtd,pct,obs,status = '—','—','','⛔ Fora (quinta-feira)'
            fill = FILL_GRAY
        elif a not in cotas:
            qtd,pct,obs,status = '—','—','','⚪ Fora (cota ≤ 40)' if a=='TATIANA' else ''
            fill = FILL_GRAY
        else:
            qtd = len(alloc.get(a,[]))
            pct = f"{qtd/total_unico*100:.1f}%" if total_unico else '0%'
            obs,status = '',''
            if a=='TATIANA': status='✅ Incluída'
            elif a in('ALANIS','ANNA JULIA','ANA CECILIA'): status='50% da cota'
            else: status='Ativa'
            if a=='BARBARA': obs='Prioridade GPM'
            if a=='PALOMA':  obs='Prioridade não-trabalhista'
            if a in('ANNA JULIA','ANA CECILIA'): obs='Prioridade não-trabalhista (50%)'

        for ci,val in enumerate([nome,qtd,pct,obs,status],start=1):
            _set_cell(ws.cell(row=r,column=ci), val, font=_font(size=9), fill=fill,
                      align=_align('center' if ci>1 else 'left','center'))

# ─── ABA INDIVIDUAL POR ANALISTA ──────────────────────────────────────────────

def _build_analista_tab(ws, analista, nome_exib, rows_df, df_cols, d_str):
    if rows_df.empty: return

    # Mapear colunas de saída
    out_cols = []
    col_map  = {normalizar(c): c for c in df_cols}
    for display, width, cands in OUTPUT_COLS:
        src = None
        for c in cands:
            k = normalizar(c)
            if k in col_map: src = col_map[k]; break
        out_cols.append((display, width, src))

    n_cols = len(out_cols)

    # --- Linha 1: Título da aba ---
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    _set_cell(ws.cell(row=1,column=1), f"{nome_exib} — {d_str}",
              font=_font(bold=True,size=11,color='FFFFFFFF'),
              fill=FILL_TITLE, align=_align('center','center'))
    ws.row_dimensions[1].height = 24

    # --- Linha 2: Cabeçalhos de coluna ---
    for ci,(display,width,_) in enumerate(out_cols,start=1):
        _set_cell(ws.cell(row=2,column=ci), display,
                  font=_font(bold=True,size=10,color='FFFFFFFF'),
                  fill=FILL_HEADER, align=_align('center','center'))
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.row_dimensions[2].height = 20

    # Freeze panes em A3 (rows 1 e 2 fixas)
    ws.freeze_panes = 'A3'

    # --- Classificar tribunais ---
    col_cnj_src = None
    for _,_,src in out_cols:
        if src and normalizar(src) == 'NUMERO DE CNJ': col_cnj_src = src; break
    if col_cnj_src and col_cnj_src in rows_df.columns:
        tribunais = rows_df[col_cnj_src].apply(classificar_tribunal)
    else:
        tribunais = pd.Series(['OUTROS']*len(rows_df), index=rows_df.index)

    contagem = tribunais.value_counts()
    grupos   = [g for g in contagem.index if g != 'OUTROS']
    minoritario = contagem.loc[grupos].idxmin() if len(grupos)>1 else None

    fill_base = ANALISTA_FILL.get(analista, FILL_DATA)

    # --- Dados (a partir da linha 3) ---
    for ri,(idx,_) in enumerate(rows_df.iterrows(), start=3):
        data_row_idx = ri - 2              # 1-based dentro dos dados
        trib = tribunais[idx] if idx in tribunais.index else 'OUTROS'
        is_min = (trib == minoritario)

        if is_min:
            row_fill = FILL_TRIB_MIN
            row_font_color = 'FFFFFFFF'
        elif data_row_idx % 2 == 0:
            row_fill = FILL_WHITE
            row_font_color = 'FF000000'
        else:
            row_fill = fill_base
            row_font_color = 'FF000000'

        for ci,(_,_,src) in enumerate(out_cols,start=1):
            val = rows_df.at[idx, src] if src and src in rows_df.columns else ''
            _set_cell(ws.cell(row=ri,column=ci), val,
                      font=_font(size=9,color=row_font_color),
                      fill=row_fill, align=_align('left','center'))

    # --- Legenda ---
    if minoritario:
        leg_row = len(rows_df) + 6
        ws.merge_cells(start_row=leg_row, start_column=1, end_row=leg_row, end_column=n_cols)
        _set_cell(ws.cell(row=leg_row,column=1),
                  f"🟠 Laranja = tribunal minoritário — {minoritario} ({contagem.get(minoritario,0)} publ.)",
                  font=_font(italic=True,size=8,color='FF7B2D00'),
                  fill=None, border=False, align=_align('left','center'))

# ─── ABA POR COORDENADOR ──────────────────────────────────────────────────────

def _build_coordenador_tab(ws, df, col_resp, mapeamento, d_str, total_unico):
    for col,w in [(1,42),(2,16),(3,14)]:
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.merge_cells('A1:C1')
    _set_cell(ws['A1'], f"TOTAL POR COORDENADOR — {d_str}",
              font=_font(bold=True,size=13,color='FFFFFFFF'),
              fill=FILL_TITLE, align=_align('center','center'))
    ws.row_dimensions[1].height = 30

    for ci,lbl in enumerate(['COORDENADOR','PUBLICAÇÕES','% DO TOTAL'],start=1):
        _set_cell(ws.cell(row=2,column=ci), lbl,
                  font=_font(bold=True,size=10,color='FFFFFFFF'),
                  fill=FILL_HEADER, align=_align('center','center'))

    contagem = {}
    if col_resp:
        for val in df[col_resp].dropna():
            coord = mapeamento.get(normalizar(str(val)), '—')
            contagem[coord] = contagem.get(coord,0)+1

    dados = sorted(contagem.items(), key=lambda x:-x[1])
    cores  = GRAD_COORD
    for i,(coord,qtd) in enumerate(dados,start=3):
        hex_color = cores[min(i-3, len(cores)-1)]
        fill = _fill(hex_color)
        pct  = f"{qtd/total_unico*100:.1f}%" if total_unico else '0%'
        for ci,val in enumerate([coord,qtd,pct],start=1):
            _set_cell(ws.cell(row=i,column=ci), val,
                      font=_font(size=9,color='FFFFFFFF'),
                      fill=fill, align=_align('center' if ci>1 else 'left','center'))

    if dados:
        n = len(dados)
        chart = BarChart(); chart.type="bar"
        chart.title="Publicações por Coordenador"
        chart.style=10; chart.width=22; chart.height=14
        data_ref = Reference(ws, min_col=2, min_row=2, max_row=2+n)
        cats_ref = Reference(ws, min_col=1, min_row=3, max_row=2+n)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart.series[0].graphicalProperties.solidFill = "2E4D8A"
        ws.add_chart(chart, "D3")

# ─── ABA POR RESPONSÁVEL ──────────────────────────────────────────────────────

def _build_responsavel_tab(ws, df, col_resp, mapeamento, d_str, total_unico):
    for col,w in [(1,38),(2,36),(3,14),(4,12),(5,14)]:
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.merge_cells('A1:E1')
    _set_cell(ws['A1'], f"TOTAL POR RESPONSÁVEL — {d_str}",
              font=_font(bold=True,size=13,color='FFFFFFFF'),
              fill=FILL_TITLE, align=_align('center','center'))
    ws.row_dimensions[1].height = 30

    for ci,lbl in enumerate(['RESPONSÁVEL','COORDENADOR','STATUS','TOTAL','%'],start=1):
        _set_cell(ws.cell(row=2,column=ci), lbl,
                  font=_font(bold=True,size=10,color='FFFFFFFF'),
                  fill=FILL_HEADER, align=_align('center','center'))

    contagem = {}
    if col_resp:
        for val in df[col_resp].dropna():
            nome = str(val).strip()
            contagem[nome] = contagem.get(nome,0)+1

    dados  = sorted(contagem.items(), key=lambda x:-x[1])
    cores  = _gradient_colors(max(len(dados),1))
    for i,(resp,qtd) in enumerate(dados,start=3):
        resp_norm = normalizar(resp)
        coord     = mapeamento.get(resp_norm,'—')
        inativo   = resp_norm in INATIVOS
        status    = "⚠️ Inativo" if inativo else "Ativo"
        pct       = f"{qtd/total_unico*100:.1f}%" if total_unico else '0%'

        if inativo:
            row_fill = FILL_INATIVO
            font_col = 'FF000000'
        else:
            row_fill = _fill(cores[min(i-3,len(cores)-1)])
            font_col = 'FFFFFFFF'

        for ci,val in enumerate([resp,coord,status,qtd,pct],start=1):
            _set_cell(ws.cell(row=i,column=ci), val,
                      font=_font(size=9,color=font_col if not inativo else 'FF000000'),
                      fill=row_fill, align=_align('center' if ci>2 else 'left','center'))

    # Legenda
    leg = len(dados) + 4
    ws.merge_cells(start_row=leg,start_column=1,end_row=leg,end_column=5)
    _set_cell(ws.cell(row=leg,column=1),
              "⚠️ Laranja claro = Responsável inativo no sistema",
              font=_font(italic=True,size=8,color='FF7B2D00'),
              fill=None, border=False, align=_align('left','center'))

    # Gráfico Top 15
    n_chart = min(15, len(dados))
    if n_chart > 0:
        chart2 = BarChart(); chart2.type="bar"
        chart2.title="Top 15 Responsáveis"; chart2.style=10
        chart2.width=24; chart2.height=16
        dr = Reference(ws, min_col=4, min_row=2, max_row=2+n_chart)
        cr = Reference(ws, min_col=1, min_row=3, max_row=2+n_chart)
        chart2.add_data(dr, titles_from_data=True)
        chart2.set_categories(cr)
        chart2.series[0].graphicalProperties.solidFill = "3A5FA8"
        ws.add_chart(chart2, "G3")

# ─── PRÉ-CHECK ────────────────────────────────────────────────────────────────

def pre_check(input_data, filename="", extra_mappings=None):
    if extra_mappings is None: extra_mappings = {}
    mapeamento = {**COORDENADORES_MAPEADOS,
                  **{normalizar(k):v for k,v in extra_mappings.items()}}

    df = _carregar_df(input_data)
    col_cnj  = _encontrar_coluna(df, ['Número de CNJ','Numero de CNJ','CNJ'])
    col_resp = _encontrar_coluna(df, ['Responsável','Responsavel','RESPONSAVEL'])

    total_bruto = len(df)
    if col_cnj: df = df.drop_duplicates(subset=[col_cnj])
    total_unico = len(df)

    d, d_str = _extrair_data(df, filename)

    unmapped, inativos = [], []
    if col_resp:
        for val in df[col_resp].dropna().unique():
            norm = normalizar(str(val))
            if norm in INATIVOS: inativos.append(str(val).strip())
            if norm not in mapeamento: unmapped.append(str(val).strip())

    return {
        "data": d, "data_str": d_str,
        "dia_semana": _dia_semana_norm(d), "dia_semana_nome": _dia_semana_pt(d),
        "total_bruto": total_bruto, "total_unico": total_unico,
        "duplicatas": total_bruto-total_unico,
        "unmapped": sorted(set(unmapped)),
        "inativos_encontrados": sorted(set(inativos)),
    }

# ─── GERAR RELATÓRIO ──────────────────────────────────────────────────────────

def gerar_relatorio(input_data, filename="", extra_mappings=None, divisao_especial=False):
    if extra_mappings is None: extra_mappings = {}
    mapeamento = {**COORDENADORES_MAPEADOS,
                  **{normalizar(k):v for k,v in extra_mappings.items()}}

    raw = input_data if isinstance(input_data,bytes) else input_data.read()
    df  = _carregar_df(raw)

    col_cnj  = _encontrar_coluna(df, ['Número de CNJ','Numero de CNJ','CNJ'])
    col_nat  = _encontrar_coluna(df, ['Natureza','NATUREZA','natureza'])
    col_cli  = _encontrar_coluna(df, ['Cliente','CLIENTE','cliente'])
    col_resp = _encontrar_coluna(df, ['Responsável','Responsavel','RESPONSAVEL'])

    total_bruto = len(df)
    if col_cnj: df = df.drop_duplicates(subset=[col_cnj])
    df = df.reset_index(drop=True)
    total_unico = len(df)

    d, d_str        = _extrair_data(df, filename)
    dia_semana      = _dia_semana_norm(d)
    dia_semana_nome = _dia_semana_pt(d)

    cotas = _calcular_cotas(total_unico, dia_semana)
    alloc = _distribuir(df, cotas, col_cnj, col_nat, col_cli)

    wb = Workbook(); wb.remove(wb.active)

    NOMES_ABA = {'VANESSA':'VANESSA','PALOMA':'PALOMA','BARBARA':'BÁRBARA',
                 'LARA':'LARA','ANNA JULIA':'ANNA JÚLIA','ANA CECILIA':'ANA CECÍLIA',
                 'ALANIS':'ALANIS','TATIANA':'TATIANA'}
    ORDEM = ['VANESSA','PALOMA','BARBARA','LARA','ANNA JULIA','ANA CECILIA','ALANIS','TATIANA']

    # RESUMO
    ws_res = wb.create_sheet('RESUMO')
    _build_resumo(ws_res, d_str, dia_semana_nome, total_bruto, total_unico,
                  cotas, alloc, dia_semana, df, col_nat, col_cli, divisao_especial)

    # Abas individuais
    for a in ORDEM:
        if a in alloc and alloc[a]:
            rows_df = df.loc[alloc[a]]
            ws = wb.create_sheet(NOMES_ABA.get(a,a))
            _build_analista_tab(ws, a, NOMES_ABA.get(a,a), rows_df,
                                df.columns.tolist(), d_str)

    # POR COORDENADOR
    ws_coord = wb.create_sheet('POR COORDENADOR')
    _build_coordenador_tab(ws_coord, df, col_resp, mapeamento, d_str, total_unico)

    # POR RESPONSÁVEL
    ws_resp = wb.create_sheet('POR RESPONSÁVEL')
    _build_responsavel_tab(ws_resp, df, col_resp, mapeamento, d_str, total_unico)

    out = io.BytesIO(); wb.save(out); out.seek(0)

    resumo = {
        "data_str": d_str, "dia_semana_nome": dia_semana_nome,
        "total_bruto": total_bruto, "total_unico": total_unico,
        "duplicatas": total_bruto-total_unico,
        "cotas": cotas, "tatiana_incluida": 'TATIANA' in cotas,
        "alanis_ativa": dia_semana!='terca',
        "anna_julia_ativa": dia_semana!='quinta',
        "alloc_counts": {a: len(v) for a,v in alloc.items()},
    }
    return out.getvalue(), resumo
