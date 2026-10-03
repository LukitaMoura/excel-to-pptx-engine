# -*- coding: utf-8 -*-
"""Especificacao unica dos slides do relatorio de RH.

Esta estrutura alimenta DOIS programas:
  * gerar_modelo_excel.py  -> cria a planilha que a gestora preenche
  * gerar_apresentacao.py  -> le a planilha e monta o PowerPoint

Manter os dados aqui garante que planilha e PPT nunca fiquem dessincronizados.

Modelo de cada bloco (tabela/grafico de um slide):
    {
      "key":     identificador unico dentro do slide (vira marcador na planilha)
      "title":   titulo exibido acima do bloco
      "render":  como desenhar -> kpi | table | column | bar | line |
                 pie | doughnut | column_stacked | bar_stacked | cota
      "columns": cabecalhos das colunas
      "rows":    linhas de exemplo (a gestora substitui pelos dados do mes)
      "pos":     (left, top, width, height) em polegadas no slide
    }

O bloco "cota" desenha um selo de conformidade (verde "dentro" / vermelho "fora").
A 1a linha do bloco e o marcador ["__status__", "dentro"|"fora"]; as demais sao
pares (rotulo, valor) exibidos como legenda ao lado do selo.

Convencao p/ graficos: 1a coluna = categoria (eixo X / fatias);
demais colunas numericas = series.
Convencao p/ kpi: colunas ["Indicador", "Valor"]; cada linha vira um cartao.
"""

# ---------------------------------------------------------------------------
# Metadados da capa (slide 1) - editaveis na aba "Capa" da planilha
# ---------------------------------------------------------------------------
META = {
    "titulo":         "Indicadores de Recursos Humanos",
    "subtitulo":      "Suporte ao Negocio",
    "periodo":        "Fevereiro / 2022",
    "area":           "Recursos Humanos - Empresa Exemplo",
    "data_referencia": "05/03/2022",
}

# ---------------------------------------------------------------------------
# Posicoes padrao (polegadas) - slide 16:9 = 13.333 x 7.5
# ---------------------------------------------------------------------------
FULL   = (0.40, 1.05, 12.53, 6.10)          # bloco unico ocupando o slide
KPI    = (0.40, 1.05, 12.53, 1.05)          # faixa de cartoes KPI no topo

# duas colunas abaixo de uma faixa KPI
L_HALF = (0.40, 2.30, 6.06, 4.85)
R_HALF = (6.86, 2.30, 6.06, 4.85)

# duas colunas sem faixa KPI
L_TOP  = (0.40, 1.05, 6.06, 6.10)
R_TOP  = (6.86, 1.05, 6.06, 6.10)

# quadrantes (sem KPI)
Q_TL = (0.40, 1.05, 6.06, 2.95)
Q_TR = (6.86, 1.05, 6.06, 2.95)
Q_BL = (0.40, 4.20, 6.06, 2.95)
Q_BR = (6.86, 4.20, 6.06, 2.95)

# quadrantes abaixo de uma faixa KPI
Q_TL2 = (0.40, 2.30, 6.06, 2.45)
Q_TR2 = (6.86, 2.30, 6.06, 2.45)
Q_BL2 = (0.40, 4.85, 6.06, 2.30)
Q_BR2 = (6.86, 4.85, 6.06, 2.30)

MESES_12 = ["mar/21", "abr/21", "mai/21", "jun/21", "jul/21", "ago/21",
            "set/21", "out/21", "nov/21", "dez/21", "jan/22", "fev/22"]


def b(key, title, render, columns, rows, pos):
    return {"key": key, "title": title, "render": render,
            "columns": columns, "rows": rows, "pos": pos}


# ===========================================================================
# SLIDES  (capa/fechamento a parte). Ordem ALINHADA a COMPUTERS em dados.py.
# Revisao da devolutiva do RH: Sobreaviso e Horas Extras removidos;
# Turnover/Desligamento/Span consolidados; PCD dividido em PCD + Aprendiz;
# novos slides de Absenteismo e Terceiros.
# ===========================================================================
SLIDES = [

    # ----- 1: Visao Geral por Negocio (sem Lideranca/Tx Deslig; com Aprendiz) -----
    {"title": "Indicadores de RH - por Negocio", "blocks": [
        b("geral", "Indicadores por Negocio / Diretoria", "table",
          ["Negocio", "Colab.", "Admitidos", "Desligados", "PCD", "Aprendiz"],
          [["Suporte ao Negocio", 3972, 18, 45, 268, 8],
           ["Total",              3972, 18, 45, 268, 8]],
          FULL)]},

    # ----- 2: Head Count -----
    {"title": "Head Count (HC)", "blocks": [
        b("kpis", "", "kpi", ["Indicador", "Valor"],
          [["Total Colaboradores", 3972], ["Masculino", 2233], ["Feminino", 1739],
           ["Ativos", 3414], ["Inativos", 558], ["PCD", 268], ["Aprendiz", 8]],
          KPI),
        b("piramide", "Piramide por Classe Funcional (Ativos)", "bar",
          ["Classe", "Qtd"],
          [["13 - PRE", 1], ["11 - VP", 5], ["10 - DIR", 46], ["09 - GS", 56],
           ["08 - G", 185], ["07 - CS", 17], ["06 - C", 374], ["05 - S", 3],
           ["04 - E", 1132], ["03 - P", 1328], ["02 - A", 260], ["01 - APR/EST", 7]],
          L_HALF),
        b("tempo_casa", "Tempo de Casa (% colab.)", "column",
          ["Faixa", "% Colab"],
          [["0-05", 21.93], ["06-10", 29.15], ["11-15", 26.76], ["16-20", 11.15],
           ["21-25", 8.18], ["26-30", 1.51], ["31-35", 0.70], ["36-40", 0.40],
           ["41+", 0.20]],
          Q_TR2),
        b("idade_genero", "Idade e Genero (% colab.)", "column_stacked",
          ["Faixa", "Feminino", "Masculino"],
          [["0-20", 0.28, 0.00], ["21-30", 3.20, 2.84], ["31-40", 16.44, 18.63],
           ["41-50", 16.59, 21.55], ["51+", 7.43, 13.04]],
          Q_BR2)]},

    # ----- 3: HC Evolucao (HC Realizado x Projetado) -----
    {"title": "Head Count - Evolucao", "blocks": [
        b("kpis", "", "kpi", ["Indicador", "Valor"],
          [["Admitidos (12m)", 183], ["Desligados (12m)", -515], ["Diferenca", -332]],
          KPI),
        b("evolucao_hc", "Evolucao de HC (Realizado x Projetado)", "line",
          ["Mes", "HC Realizado", "HC Projetado"],
          [[MESES_12[0], 4297, 4300], [MESES_12[1], 4278, 4280],
           [MESES_12[2], 4238, 4250], [MESES_12[3], 4207, 4220],
           [MESES_12[4], 4182, 4190], [MESES_12[5], 4127, 4160],
           [MESES_12[6], 4083, 4120], [MESES_12[7], 4036, 4080],
           [MESES_12[8], 3981, 4040], [MESES_12[9], 4004, 4010],
           [MESES_12[10], 3983, 3990], [MESES_12[11], 3972, 3970]],
          L_HALF),
        b("adm_dem", "Admitidos x Demitidos", "column",
          ["Mes", "Admitidos", "Demitidos"],
          [[MESES_12[0], 0, 24], [MESES_12[1], 0, 29], [MESES_12[2], 0, 44],
           [MESES_12[3], 0, 34], [MESES_12[4], 0, 32], [MESES_12[5], 0, 66],
           [MESES_12[6], 18, 60], [MESES_12[7], 27, 51], [MESES_12[8], 0, 58],
           [MESES_12[9], 31, 31], [MESES_12[10], 19, 41], [MESES_12[11], 18, 45]],
          R_HALF)]},

    # ----- 4: Admitidos -----
    {"title": "Admitidos", "blocks": [
        b("kpis", "", "kpi", ["Indicador", "Valor"],
          [["Admitidos no Mes", 18], ["Presencial (CLT)", 18], ["Estagio/Aprendiz", 0]],
          KPI),
        b("por_genero", "Admitidos por Genero", "pie",
          ["Genero", "Qtd"],
          [["Feminino", 5], ["Masculino", 13]],
          L_HALF),
        b("por_classe", "Admitidos por Classe Funcional", "bar",
          ["Classe", "Qtd"],
          [["03 - P", 9], ["04 - E", 7], ["06 - C", 1], ["08 - G", 1]],
          Q_TR2),
        b("pareto", "Top Gerencias - Admissoes", "table",
          ["Gerencia", "Admissoes", "% Ranking"],
          [["DIR Sistemas OSS BSS", 6, "33%"],
           ["GER Novas Plataformas Digitais", 2, "44%"],
           ["DIR Infraestrutura", 1, "50%"],
           ["DIR Faturamento", 1, "56%"]],
          Q_BR2)]},

    # ----- 5: Turnover (tempo de casa em 3 faixas) -----
    {"title": "Turnover", "blocks": [
        b("kpis", "", "kpi", ["Indicador", "Valor"],
          [["Desligados", 45], ["Voluntario", 30], ["Involuntario", 15],
           ["Turnover Geral", "0,79%"], ["Tx Deslig", "1,13%"],
           ["Tx Vol", "0,76%"], ["Tx Invol", "0,38%"]],
          KPI),
        b("tempo_casa", "Desligados por Tempo de Casa", "bar",
          ["Faixa", "Qtd"],
          [["Ate 6 Meses", 6], ["6 Meses a 2 Anos", 0], ["Acima de 2 Anos", 39]],
          L_HALF),
        b("geracao", "Desligados por Geracao", "column",
          ["Geracao", "Qtd"],
          [["Geracao X", 23], ["Geracao Y", 21], ["Geracao Z", 1]],
          Q_TR2),
        b("genero", "Desligados por Genero", "doughnut",
          ["Genero", "%"],
          [["Feminino", 42.22], ["Masculino", 57.78]],
          Q_BR2)]},

    # ----- 6: Turnover Evolucao (consolidado: evolucao + ano corrente) -----
    {"title": "Turnover - Evolucao", "blocks": [
        b("kpis", "", "kpi", ["Indicador", "Valor"],
          [["Turnover Acumulado", "7,50%"], ["Tx Deslig Acumulado", "11,38%"],
           ["Turnover Ano Corrente", "1,55%"]],
          KPI),
        b("serie", "Evolucao Turnover x Taxa de Desligamento", "line",
          ["Mes", "Turnover %", "Tx Deslig %"],
          [["jan/21", 0.43, 0.65], ["fev/21", 0.46, 0.49], ["mar/21", 0.56, 0.63],
           ["abr/21", 0.35, 0.63], ["mai/21", 0.58, 1.01], ["jun/21", 0.45, 0.81],
           ["jul/21", 0.53, 0.96], ["ago/21", 0.98, 1.69], ["set/21", 0.95, 1.46],
           ["out/21", 0.96, 1.26], ["nov/21", 0.82, 1.45], ["dez/21", 0.77, 0.98]],
          (0.40, 2.30, 12.53, 4.85))]},

    # ----- 7: Desligados (1/2) -----
    {"title": "Desligados", "blocks": [
        b("kpis", "", "kpi", ["Indicador", "Valor"],
          [["Demitidos", 45], ["Voluntario", 30], ["Involuntario", 15]],
          KPI),
        b("por_genero", "Desligados por Genero", "table",
          ["Genero", "Qtd"], [["Feminino", 19], ["Masculino", 26], ["Total", 45]],
          Q_TL2),
        b("por_vinculo", "Desligados por Tipo de Vinculo", "table",
          ["Vinculo", "Qtd"], [["Presencial (CLT)", 45], ["Total", 45]],
          Q_TR2),
        b("por_classe", "Desligados por Classe Funcional", "column",
          ["Classe", "Qtd"],
          [["02 - A", 7], ["03 - P", 16], ["04 - E", 16], ["06 - C", 4],
           ["08 - G", 1], ["09 - GS", 1]],
          Q_BL2),
        b("por_tipo", "Desligados por Tipo", "table",
          ["Tipo", "Qtd"], [["Voluntario", 30], ["Involuntario", 15]],
          Q_BR2)]},

    # ----- 8: Desligados Evolucao (2/2) -----
    {"title": "Desligados - Evolucao", "blocks": [
        b("serie", "Desligados Voluntario x Involuntario", "line",
          ["Mes", "Voluntario", "Involuntario"],
          [["mar/21", 19, 13], ["abr/21", 14, 8], ["mai/21", 30, 14],
           ["jun/21", 18, 16], ["jul/21", 13, 4], ["ago/21", 38, 28],
           ["set/21", 34, 17], ["out/21", 32, 19], ["nov/21", 26, 5],
           ["dez/21", 35, 6], ["jan/22", 30, 11], ["fev/22", 30, 15]],
          L_TOP),
        b("maiores_vol", "Maiores % Voluntaria por Gerencia", "table",
          ["Gerencia", "% Vol"],
          [["DIR Eng. e Tec. da Informacao", "100%"],
           ["GER Novas Plataformas Digitais", "100%"],
           ["DIR Sistemas e Analytics", "94%"]],
          Q_TR),
        b("maiores_tx", "Maiores Taxas de Deslig. por Gerencia", "table",
          ["Gerencia", "% Tx"],
          [["GER Novas Plataformas Digitais", "36,2%"],
           ["GER Responsabilidade Social", "35,3%"],
           ["GER Gestao Indicadores Juridicos", "29,2%"]],
          Q_BR)]},

    # ----- 9: Span de Controle (consolidado em 1 slide, 1 grafico) -----
    {"title": "Span de Controle", "blocks": [
        b("kpis", "", "kpi", ["Indicador", "Valor"],
          [["Colaboradores Ativos", 3414], ["Lideranca", 671],
           ["Base", 2743], ["Span", "5,09"]],
          KPI),
        b("categoria", "% por Categoria", "doughnut",
          ["Categoria", "%"], [["Lideranca", 19.65], ["Base", 80.35]],
          L_HALF),
        b("por_diretoria", "Span por Diretoria", "table",
          ["Diretoria", "Ativos", "Lideranca", "Span"],
          [["DIR Auditoria", 47, 13, "3,62"],
           ["DIR Compras", 95, 18, "5,28"],
           ["DIR Recursos Humanos", 227, 42, "5,43"],
           ["Presidencia", 14, 11, "1,27"],
           ["VP Eng. e Tec. da Informacao", 227, 42, "5,40"],
           ["VP Financeira", 1221, 208, "5,87"]],
          R_HALF)]},

    # ----- 10: PCD (com selo de conformidade da cota legal) -----
    {"title": "PCD - Pessoas com Deficiencia", "blocks": [
        b("kpis", "", "kpi", ["Indicador", "Valor"],
          [["Total Colaboradores", 3972], ["PCD Total", 268],
           ["% do Quadro", "6,75%"]],
          KPI),
        b("conformidade", "Cota Legal (Lei 8.213/91 - 2% a 5%)", "cota",
          ["Indicador", "Valor"],
          [["__status__", "dentro"],
           ["Cota legal (2%)", 80],
           ["PCD ativos", 268],
           ["Situacao", "Dentro da cota"]],
          L_HALF),
        b("por_negocio", "PCD por Negocio", "table",
          ["Negocio", "PCD", "% do Total"],
          [["Suporte ao Negocio", 268, "100%"]],
          Q_TR2),
        b("raca_genero", "Raca e Genero PCD (%)", "column_stacked",
          ["Raca", "Feminino", "Masculino"],
          [["Branca", 33.0, 30.0], ["Parda", 13.0, 13.0], ["Preta", 7.0, 1.4],
           ["Nao Informado", 1.1, 1.1]],
          Q_BR2)]},

    # ----- 11: Aprendiz (NOVO; cota "de 1 a 3") -----
    {"title": "Programa de Aprendizes", "blocks": [
        b("kpis", "", "kpi", ["Indicador", "Valor"],
          [["Total Colaboradores", 3972], ["Aprendizes Ativos", 8],
           ["Cota Legal", "de 1 a 3"]],
          KPI),
        b("conformidade", "Cota de Aprendizes (de 1 a 3)", "cota",
          ["Indicador", "Valor"],
          [["__status__", "dentro"],
           ["Cota legal", "de 1 a 3"],
           ["Aprendizes ativos", 8],
           ["Situacao", "Dentro da cota"]],
          L_HALF),
        b("por_negocio", "Aprendizes por Diretoria", "table",
          ["Diretoria", "Aprendizes"],
          [["DIR Tecnologia", 3], ["DIR Recursos Humanos", 2],
           ["DIR Financeira", 2], ["DIR Compras", 1]],
          Q_TR2),
        b("por_genero", "Aprendizes por Genero", "bar",
          ["Genero", "Qtd"],
          [["Feminino", 5], ["Masculino", 3]],
          Q_BR2)]},

    # ----- 12: Genero (com % de lideranca por genero) -----
    {"title": "Genero", "blocks": [
        b("kpis", "", "kpi", ["Indicador", "Valor"],
          [["Total Colaboradores", 3972], ["Masculino", 2233],
           ["Feminino", 1739], ["% Mulheres Lideranca", "36,51%"],
           ["% Homens Lideranca", "63,49%"]],
          KPI),
        b("piramide", "Piramide por Genero", "table",
          ["Nivel", "Feminino", "Masculino", "Total"],
          [["1 - Presidente", 0, 1, 1], ["2 - CEO/VP", 0, 5, 5],
           ["3 - Diretor", 9, 37, 46], ["4 - Gerente", 95, 163, 258],
           ["5 - Coordenador", 148, 232, 380], ["6 - Demais Niveis", 1482, 1792, 3274],
           ["7 - APR/EST", 5, 3, 8], ["Total", 1739, 2233, 3972]],
          L_HALF),
        b("raca_genero", "Raca e Genero", "column_stacked",
          ["Raca", "Feminino", "Masculino"],
          [["Branca", 31.29, 37.61], ["Parda", 8.26, 12.31], ["Preta", 6.17, 0],
           ["Amarela", 2.57, 0], ["Nao Inf.", 1.64, 0]],
          Q_TR2),
        b("lideranca", "Lideranca por Genero", "table",
          ["Grupo", "Lider F", "Lider M", "% F", "% M"],
          [["Lideranca", 245, 426, "36,51%", "63,49%"]],
          Q_BR2)]},

    # ----- 13: Absenteismo (NOVO) -----
    {"title": "Absenteismo", "blocks": [
        b("kpis", "", "kpi", ["Indicador", "Valor"],
          [["Absenteismo Medio", "1,11%"], ["Horas Previstas", "536"],
           ["Horas Perdidas", "5,9"]],
          KPI),
        b("serie", "Evolucao do Absenteismo (%)", "line",
          ["Mes", "Absenteismo %"],
          [["jan", 0.78], ["fev", 1.85], ["mar", 1.27], ["abr", 0.55],
           ["mai", 1.10], ["jun", 0.95]],
          (0.40, 2.30, 12.53, 4.85))]},

    # ----- 14: Custo Rescisorio -----
    {"title": "Custo Rescisorio", "blocks": [
        b("kpis", "", "kpi", ["Indicador", "Valor"],
          [["Custo Resc. Ano", "1,10 Mi"], ["Custo Resc. Mes", "720,4 Mil"],
           ["Desligados (Mes)", 14]],
          KPI),
        b("por_tempo", "Resumo por Tempo de Casa", "table",
          ["Tempo de Casa", "Qtd", "Custo Resc. (R$)", "Media (R$)"],
          [["0-05", 4, "5.460", "1.365"], ["06-10", 5, "61.451", "12.290"],
           ["11-15", 3, "115.976", "38.659"], ["16-20", 2, "196.347", "98.174"],
           ["Total", 14, "720.444", "51.460"]],
          L_HALF),
        b("evolucao", "Evolucao do Custo Rescisorio", "line",
          ["Mes", "Custo (R$ Mi)"],
          [["jan/21", 0.19], ["fev/21", 0.62], ["mar/21", 0.29], ["abr/21", 0.28],
           ["mai/21", 0.54], ["jun/21", 1.37], ["jul/21", 1.40], ["ago/21", 0.76],
           ["set/21", 0.40], ["out/21", 0.20], ["nov/21", 0.30], ["dez/21", 0.45],
           ["jan/22", 0.38], ["fev/22", 0.72]],
          R_HALF)]},

    # ----- 15: Terceiros (NOVO; grau de terceirizacao) -----
    {"title": "Terceiros", "blocks": [
        b("kpis", "", "kpi", ["Indicador", "Valor"],
          [["Terceiros Ativos", 32], ["Entradas", 7], ["Saidas", 25],
           ["Grau Terceirizacao", "20,1%"]],
          KPI),
        b("por_parceiro", "Terceiros por Parceiro", "bar",
          ["Parceiro", "Qtd"],
          [["Parceiro A", 7], ["Parceiro B", 7], ["Parceiro C", 6], ["Parceiro D", 6]],
          L_HALF),
        b("por_diretoria", "Grau de Terceirizacao por Area", "table",
          ["Diretoria", "Proprios", "Terceiros", "GTER %"],
          [["Operacoes", 70, 40, "36,4%"], ["Real Estate", 16, 6, "27,3%"],
           ["Comercial", 11, 3, "21,4%"], ["Financas", 19, 5, "20,8%"]],
          Q_TR2),
        b("por_orcamento", "Terceiros por Orcamento", "doughnut",
          ["Orcamento", "Qtd"],
          [["CAPEX", 17], ["OPEX", 15]],
          Q_BR2)]},
]
