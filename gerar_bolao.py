#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gera a planilha do bolão da Copa do Mundo 2026 (fase de grupos).

Saída: bolao_copa_2026.xlsx  (pronta para subir no Google Sheets)

Estrutura de cada linha (um jogo):
    Data | Grupo | Time 1 | Bandeira | Placar Real T1 | Palpite T1 |
    Palpite T2 | Placar Real T2 | Bandeira | Time 2

As colunas de "Placar Real" e "Palpite" ficam em branco para preencher.
Os nomes dos times estão em português; as bandeiras são emoji.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ANO = "2026"

# --- Times: nome em português -> código de bandeira (ISO, usado no flagcdn) --
# As bandeiras são exibidas como IMAGEM real (função =IMAGE do Google Sheets),
# e não como emoji — assim aparecem em qualquer sistema, sem virar a sigla de
# 2 letras. Escócia/Inglaterra usam os códigos de subdivisão gb-sct / gb-eng.
CODIGO_BANDEIRA = {
    "México": "mx",
    "África do Sul": "za",
    "Coreia do Sul": "kr",
    "Rep. Tcheca": "cz",
    "Canadá": "ca",
    "Bósnia e Herzegovina": "ba",
    "Catar": "qa",
    "Suíça": "ch",
    "Brasil": "br",
    "Marrocos": "ma",
    "Haiti": "ht",
    "Escócia": "gb-sct",
    "Estados Unidos": "us",
    "Paraguai": "py",
    "Austrália": "au",
    "Turquia": "tr",
    "Alemanha": "de",
    "Curaçao": "cw",
    "Costa do Marfim": "ci",
    "Equador": "ec",
    "Países Baixos": "nl",
    "Japão": "jp",
    "Suécia": "se",
    "Tunísia": "tn",
    "Bélgica": "be",
    "Egito": "eg",
    "Irã": "ir",
    "Nova Zelândia": "nz",
    "Espanha": "es",
    "Cabo Verde": "cv",
    "Arábia Saudita": "sa",
    "Uruguai": "uy",
    "França": "fr",
    "Senegal": "sn",
    "Iraque": "iq",
    "Noruega": "no",
    "Argentina": "ar",
    "Argélia": "dz",
    "Áustria": "at",
    "Jordânia": "jo",
    "Portugal": "pt",
    "RD Congo": "cd",
    "Uzbequistão": "uz",
    "Colômbia": "co",
    "Inglaterra": "gb-eng",
    "Croácia": "hr",
    "Gana": "gh",
    "Panamá": "pa",
}

# URL das imagens de bandeira (w80 = 80px de largura, redimensionada na célula).
FLAG_URL = "https://flagcdn.com/w80/{code}.png"


def bandeira(time):
    """Fórmula =IMAGE(...) que exibe a bandeira real do time no Google Sheets."""
    return f'=IMAGE("{FLAG_URL.format(code=CODIGO_BANDEIRA[time])}")'

# --- Os 72 jogos da fase de grupos: (dia/mês, grupo, Time 1, Time 2) ---------
# Fonte: ESPN / Wikipedia (sorteio de 05/12/2025). Ordem cronológica.
JOGOS = [
    # 11/06
    ("11/06", "A", "México", "África do Sul"),
    ("11/06", "A", "Coreia do Sul", "Rep. Tcheca"),
    # 12/06
    ("12/06", "B", "Canadá", "Bósnia e Herzegovina"),
    ("12/06", "D", "Estados Unidos", "Paraguai"),
    # 13/06
    ("13/06", "B", "Catar", "Suíça"),
    ("13/06", "C", "Brasil", "Marrocos"),
    ("13/06", "C", "Haiti", "Escócia"),
    ("13/06", "D", "Austrália", "Turquia"),
    # 14/06
    ("14/06", "E", "Alemanha", "Curaçao"),
    ("14/06", "F", "Países Baixos", "Japão"),
    ("14/06", "E", "Costa do Marfim", "Equador"),
    ("14/06", "F", "Suécia", "Tunísia"),
    # 15/06
    ("15/06", "H", "Espanha", "Cabo Verde"),
    ("15/06", "G", "Bélgica", "Egito"),
    ("15/06", "H", "Arábia Saudita", "Uruguai"),
    ("15/06", "G", "Irã", "Nova Zelândia"),
    # 16/06
    ("16/06", "I", "França", "Senegal"),
    ("16/06", "I", "Iraque", "Noruega"),
    ("16/06", "J", "Argentina", "Argélia"),
    ("16/06", "J", "Áustria", "Jordânia"),
    # 17/06
    ("17/06", "K", "Portugal", "RD Congo"),
    ("17/06", "L", "Inglaterra", "Croácia"),
    ("17/06", "L", "Gana", "Panamá"),
    ("17/06", "K", "Uzbequistão", "Colômbia"),
    # 18/06
    ("18/06", "A", "Rep. Tcheca", "África do Sul"),
    ("18/06", "B", "Suíça", "Bósnia e Herzegovina"),
    ("18/06", "B", "Canadá", "Catar"),
    ("18/06", "A", "México", "Coreia do Sul"),
    # 19/06
    ("19/06", "D", "Estados Unidos", "Austrália"),
    ("19/06", "C", "Escócia", "Marrocos"),
    ("19/06", "C", "Brasil", "Haiti"),
    ("19/06", "D", "Turquia", "Paraguai"),
    # 20/06
    ("20/06", "F", "Países Baixos", "Suécia"),
    ("20/06", "E", "Alemanha", "Costa do Marfim"),
    ("20/06", "E", "Equador", "Curaçao"),
    ("20/06", "F", "Tunísia", "Japão"),
    # 21/06
    ("21/06", "H", "Espanha", "Arábia Saudita"),
    ("21/06", "G", "Bélgica", "Irã"),
    ("21/06", "H", "Uruguai", "Cabo Verde"),
    ("21/06", "G", "Nova Zelândia", "Egito"),
    # 22/06
    ("22/06", "J", "Argentina", "Áustria"),
    ("22/06", "I", "França", "Iraque"),
    ("22/06", "I", "Noruega", "Senegal"),
    ("22/06", "J", "Jordânia", "Argélia"),
    # 23/06
    ("23/06", "K", "Portugal", "Uzbequistão"),
    ("23/06", "L", "Inglaterra", "Gana"),
    ("23/06", "L", "Panamá", "Croácia"),
    ("23/06", "K", "Colômbia", "RD Congo"),
    # 24/06
    ("24/06", "B", "Suíça", "Canadá"),
    ("24/06", "B", "Bósnia e Herzegovina", "Catar"),
    ("24/06", "C", "Escócia", "Brasil"),
    ("24/06", "C", "Marrocos", "Haiti"),
    ("24/06", "A", "Rep. Tcheca", "México"),
    ("24/06", "A", "África do Sul", "Coreia do Sul"),
    # 25/06
    ("25/06", "E", "Equador", "Alemanha"),
    ("25/06", "E", "Curaçao", "Costa do Marfim"),
    ("25/06", "F", "Japão", "Suécia"),
    ("25/06", "F", "Tunísia", "Países Baixos"),
    ("25/06", "D", "Turquia", "Estados Unidos"),
    ("25/06", "D", "Paraguai", "Austrália"),
    # 26/06
    ("26/06", "I", "Noruega", "França"),
    ("26/06", "I", "Senegal", "Iraque"),
    ("26/06", "H", "Cabo Verde", "Arábia Saudita"),
    ("26/06", "H", "Uruguai", "Espanha"),
    ("26/06", "G", "Egito", "Irã"),
    ("26/06", "G", "Nova Zelândia", "Bélgica"),
    # 27/06
    ("27/06", "L", "Panamá", "Inglaterra"),
    ("27/06", "L", "Croácia", "Gana"),
    ("27/06", "K", "Colômbia", "Portugal"),
    ("27/06", "K", "RD Congo", "Uzbequistão"),
    ("27/06", "J", "Argélia", "Áustria"),
    ("27/06", "J", "Jordânia", "Argentina"),
]


def checar_integridade():
    """Garante que os dados batem antes de gerar o arquivo."""
    assert len(JOGOS) == 72, f"Esperado 72 jogos, achei {len(JOGOS)}"

    # 6 jogos por grupo (A..L)
    por_grupo = {}
    aparicoes = {}
    grupo_do_time = {}
    for _, grupo, t1, t2 in JOGOS:
        por_grupo[grupo] = por_grupo.get(grupo, 0) + 1
        for t in (t1, t2):
            aparicoes[t] = aparicoes.get(t, 0) + 1
            grupo_do_time.setdefault(t, set()).add(grupo)
            assert t in CODIGO_BANDEIRA, f"Sem bandeira para: {t}"

    grupos = sorted(por_grupo)
    assert grupos == list("ABCDEFGHIJKL"), f"Grupos inesperados: {grupos}"
    for g, n in por_grupo.items():
        assert n == 6, f"Grupo {g} tem {n} jogos (esperado 6)"

    # Cada seleção joga exatamente 3 vezes, sempre no mesmo grupo
    assert len(aparicoes) == 48, f"Esperado 48 times, achei {len(aparicoes)}"
    for t, n in aparicoes.items():
        assert n == 3, f"{t} aparece {n}x (esperado 3)"
        assert len(grupo_do_time[t]) == 1, f"{t} aparece em mais de um grupo"

    print("OK: 72 jogos, 12 grupos x 6 jogos, 48 times x 3 jogos cada.")


def gerar_planilha(caminho="bolao_copa_2026.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Fase de Grupos"

    cabecalhos = [
        "Data", "Grupo", "Time 1", "Bandeira", "Placar Real T1",
        "Palpite T1", "Palpite T2", "Placar Real T2", "Bandeira", "Time 2",
    ]
    larguras = [10, 8, 24, 10, 14, 12, 12, 14, 10, 24]
    ncols = len(cabecalhos)

    # Estilos reutilizáveis
    azul = "1F4E78"
    fonte_titulo = Font(bold=True, size=14, color="FFFFFF")
    fonte_cab = Font(bold=True, size=11, color="FFFFFF")
    fill_titulo = PatternFill("solid", fgColor=azul)
    fill_cab = PatternFill("solid", fgColor="2E75B6")
    fill_par = PatternFill("solid", fgColor="FFFFFF")
    fill_impar = PatternFill("solid", fgColor="EAF1F8")
    centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    esq = Alignment(horizontal="left", vertical="center")
    fina = Side(style="thin", color="BFBFBF")
    borda = Border(left=fina, right=fina, top=fina, bottom=fina)

    # Larguras das colunas
    for i, w in enumerate(larguras, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Linha 1: título
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value="BOLÃO COPA DO MUNDO 2026 — Fase de Grupos")
    c.font = fonte_titulo
    c.fill = fill_titulo
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    # Linha 2: cabeçalho
    for i, texto in enumerate(cabecalhos, start=1):
        cel = ws.cell(row=2, column=i, value=texto)
        cel.font = fonte_cab
        cel.fill = fill_cab
        cel.alignment = centro
        cel.border = borda
    ws.row_dimensions[2].height = 30

    # Linhas 3+: jogos (sombreamento alterna por dia)
    cols_centro = {2, 4, 5, 6, 7, 8, 9}  # Grupo, bandeiras, placares, palpites
    data_anterior = None
    listrado = False
    linha = 3
    for data, grupo, t1, t2 in JOGOS:
        if data != data_anterior:
            listrado = not listrado
            data_anterior = data
        fill = fill_impar if listrado else fill_par

        valores = [
            f"{data}/{ANO}", grupo, t1, bandeira(t1),
            "", "", "", "", bandeira(t2), t2,
        ]
        for i, v in enumerate(valores, start=1):
            cel = ws.cell(row=linha, column=i, value=v)
            cel.fill = fill
            cel.border = borda
            cel.alignment = centro if i in cols_centro else esq
        ws.row_dimensions[linha].height = 30
        linha += 1

    # Congelar título+cabeçalho e ligar o filtro no cabeçalho
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(ncols)}2"

    wb.save(caminho)
    print(f"OK: planilha gerada em {caminho} ({linha - 3} jogos).")


if __name__ == "__main__":
    checar_integridade()
    gerar_planilha()
