# -*- coding: utf-8 -*-
"""Resolucao do periodo escolhido na aba Config da planilha.

Converte (Ano, Mes, TipoPeriodo) em um intervalo de datas + a lista de meses
que compoem o periodo e a janela de evolucao usada nos graficos de tendencia.
"""
import datetime as _dt

MESES_PT = ["jan", "fev", "mar", "abr", "mai", "jun",
            "jul", "ago", "set", "out", "nov", "dez"]
MESES_PT_FULL = ["janeiro", "fevereiro", "marco", "abril", "maio", "junho",
                 "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]


def _ultimo_dia(ano, mes):
    if mes == 12:
        return _dt.date(ano, 12, 31)
    return _dt.date(ano, mes + 1, 1) - _dt.timedelta(days=1)


def label_mes(ano, mes):
    return f"{MESES_PT[mes - 1]}/{str(ano)[2:]}"


def parse_mes(valor):
    """Aceita numero (1-12) ou nome ('fevereiro', 'fev') -> int 1..12."""
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return int(valor)
    s = str(valor).strip().lower()
    if s.isdigit():
        return int(s)
    for i, nome in enumerate(MESES_PT_FULL):
        if s.startswith(nome[:3]):
            return i + 1
    return None


def meses_anteriores(ano, mes, n):
    """Lista dos n meses terminando em (ano, mes), do mais antigo ao mais novo."""
    out = []
    a, m = ano, mes
    for _ in range(n):
        out.append((a, m))
        m -= 1
        if m == 0:
            m = 12
            a -= 1
    return list(reversed(out))


class Periodo:
    def __init__(self, tipo, ano, mes, janela_evo=12):
        self.tipo = (tipo or "Mensal").strip().capitalize()
        self.ano = int(ano)
        self.mes = int(mes)
        self.janela_evo = int(janela_evo or 12)

        if self.tipo.startswith("Anual"):
            self.inicio = _dt.date(self.ano, 1, 1)
            self.fim = _dt.date(self.ano, 12, 31)
            self.meses = [(self.ano, m) for m in range(1, 13)]
        elif self.tipo.startswith("Semestr"):
            if self.mes <= 6:
                ini_m, fim_m = 1, 6
            else:
                ini_m, fim_m = 7, 12
            self.inicio = _dt.date(self.ano, ini_m, 1)
            self.fim = _ultimo_dia(self.ano, fim_m)
            self.meses = [(self.ano, m) for m in range(ini_m, fim_m + 1)]
        else:  # Mensal
            self.tipo = "Mensal"
            self.inicio = _dt.date(self.ano, self.mes, 1)
            self.fim = _ultimo_dia(self.ano, self.mes)
            self.meses = [(self.ano, self.mes)]

        # janela de evolucao (linhas de tendencia)
        if self.tipo == "Mensal":
            self.meses_evo = meses_anteriores(self.ano, self.mes, self.janela_evo)
        else:
            self.meses_evo = list(self.meses)

    # ---- helpers de filtro ----
    def contem(self, data):
        return data is not None and self.inicio <= data <= self.fim

    def ativo_no_fim(self, admissao, desligamento):
        if admissao is None or admissao > self.fim:
            return False
        return desligamento is None or desligamento > self.fim

    @property
    def label(self):
        if self.tipo == "Anual":
            return f"Ano {self.ano}"
        if self.tipo == "Semestral":
            sem = "1o Semestre" if self.mes <= 6 else "2o Semestre"
            return f"{sem} / {self.ano}"
        return f"{MESES_PT_FULL[self.mes - 1].capitalize()} / {self.ano}"

    @property
    def labels_evo(self):
        return [label_mes(a, m) for (a, m) in self.meses_evo]
