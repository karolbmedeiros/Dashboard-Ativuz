# -*- coding: utf-8 -*-
"""
Cálculo de TIR (XIRR) por sócio para o Capital Investido.

Monta, para cada sócio, a série completa de fluxos de caixa — aportes como
saídas, devoluções como entradas e, na data de hoje, a fatia do valor estimado
dos veículos que caberia a ele — e resolve a taxa que zera o valor presente
líquido com datas irregulares (XIRR).

Sem dependências externas: Newton-Raphson com bisseção como fallback.
"""

from __future__ import annotations

import csv
import io
import re
import unicodedata
from datetime import date, datetime

# ── Configuração ──────────────────────────────────────────────────────────────

# Grupos de nomes que são somados e divididos igualmente entre os membros:
# cada membro "pesa" o mesmo, independente de quanto colocou de fato.
# Nomes normalizados (minúsculos, sem acento).
GRUPOS_IGUALITARIOS: list[list[str]] = [
    ["joel araujo lopes", "lucas trindade vieira veras diniz"],
]

# Nomes que aparecem nos aportes mas não são sócios (terceiros, empréstimos).
EXCLUIR_NOMES: set[str] = {
    "adalgiza - mae joel",
}

# Descrições que representam despesa interna do caixa da empresa — não são
# movimentação de/para o bolso de um sócio e ficam fora da TIR individual.
PADROES_DESPESA_INTERNA: list[str] = [
    "saque para pagamento",
    "saque p/ pagamento",
    "mao de obra",
    "pagamento de fornecedor",
    "pagamento fornecedor",
    "despesa interna",
]

DIAS_NO_ANO = 365.0

# Limites do solver
_MAX_ITER      = 200
_TOL           = 1e-7
_RATE_MIN      = -0.9999999
_RATE_MAX      = 1e6


# ── Helpers ───────────────────────────────────────────────────────────────────

def normalizar_nome(s: str) -> str:
    """Minúsculo, sem acento, espaços colapsados — a chave de comparação."""
    s = (s or "").strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s)


def eh_despesa_interna(descricao: str) -> bool:
    d = normalizar_nome(descricao)
    return any(p in d for p in PADROES_DESPESA_INTERNA)


def _parse_data(valor) -> date | None:
    """Aceita date, datetime, 'YYYY-MM-DD', 'DD/MM/YYYY' e ISO com hora."""
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    s = str(valor or "").strip()
    if not s:
        return None
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
    if m:
        try:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
    return None


def parse_brl(s) -> float | None:
    """'R$ 6 429,10' → 6429.10 · '-5 000,00' → -5000.0 · '1234.56' → 1234.56"""
    if isinstance(s, (int, float)):
        return float(s)
    txt = re.sub(r"[^0-9,.\-]", "", str(s or "").strip())
    if not txt or txt in ("-", ".", ","):
        return None
    if re.search(r",\d{1,2}$", txt):          # formato BR: 1.234,56
        txt = txt.replace(".", "").replace(",", ".")
    else:
        txt = txt.replace(",", "")
    try:
        return float(txt)
    except ValueError:
        return None


# ── XIRR ──────────────────────────────────────────────────────────────────────

class TIRError(Exception):
    """Não foi possível calcular a TIR — a mensagem explica o porquê."""


def _vpl(taxa: float, fluxos: list[tuple[date, float]], t0: date) -> float:
    total = 0.0
    for dt, valor in fluxos:
        expo = (dt - t0).days / DIAS_NO_ANO
        total += valor / ((1.0 + taxa) ** expo)
    return total


def _dvpl(taxa: float, fluxos: list[tuple[date, float]], t0: date) -> float:
    total = 0.0
    for dt, valor in fluxos:
        expo = (dt - t0).days / DIAS_NO_ANO
        total += -expo * valor / ((1.0 + taxa) ** (expo + 1.0))
    return total


def xirr(fluxos: list[tuple[date, float]], guess: float = 0.1) -> float:
    """
    Taxa anualizada que zera o VPL de fluxos em datas irregulares.

    Newton-Raphson a partir de `guess`; se não convergir (ou sair do domínio),
    cai para bisseção sobre um intervalo com troca de sinal.
    Levanta TIRError quando o cálculo não é possível ou não converge.
    """
    if len(fluxos) < 2:
        raise TIRError("São necessários ao menos dois fluxos de caixa.")
    if not any(v < 0 for _, v in fluxos):
        raise TIRError("Não há aporte (fluxo negativo) — TIR indefinida.")
    if not any(v > 0 for _, v in fluxos):
        raise TIRError("Não há retorno (fluxo positivo) — TIR indefinida.")

    fluxos = sorted(fluxos, key=lambda f: f[0])
    t0 = fluxos[0][0]

    # 1) Newton-Raphson
    taxa = guess
    for _ in range(_MAX_ITER):
        try:
            f = _vpl(taxa, fluxos, t0)
        except (OverflowError, ZeroDivisionError, ValueError):
            break
        if abs(f) < _TOL:
            return taxa
        try:
            d = _dvpl(taxa, fluxos, t0)
        except (OverflowError, ZeroDivisionError, ValueError):
            break
        if d == 0 or not _finito(d):
            break
        nova = taxa - f / d
        if not _finito(nova) or nova <= _RATE_MIN:
            break
        if abs(nova - taxa) < _TOL:
            return nova
        taxa = nova

    # 2) Bisseção — varre o domínio atrás de uma troca de sinal
    lo = _RATE_MIN + 1e-9
    try:
        f_lo = _vpl(lo, fluxos, t0)
    except (OverflowError, ZeroDivisionError, ValueError):
        raise TIRError("Fluxos de caixa fora do domínio numérico do cálculo.")

    hi = None
    passo = 0.25
    cur = lo
    while cur < _RATE_MAX:
        cur = cur + passo if cur < 10 else cur * 2.0
        try:
            f_cur = _vpl(cur, fluxos, t0)
        except (OverflowError, ZeroDivisionError, ValueError):
            break
        if not _finito(f_cur):
            break
        if f_lo * f_cur <= 0:
            hi = cur
            break
        lo, f_lo = cur, f_cur

    if hi is None:
        raise TIRError(
            "O cálculo não convergiu em %d iterações — verifique os fluxos."
            % _MAX_ITER
        )

    for _ in range(_MAX_ITER):
        meio = (lo + hi) / 2.0
        f_meio = _vpl(meio, fluxos, t0)
        if abs(f_meio) < _TOL or (hi - lo) / 2.0 < _TOL:
            return meio
        if f_lo * f_meio <= 0:
            hi = meio
        else:
            lo, f_lo = meio, f_meio

    raise TIRError(
        "O cálculo não convergiu em %d iterações — verifique os fluxos."
        % _MAX_ITER
    )


def _finito(x: float) -> bool:
    return x == x and x not in (float("inf"), float("-inf"))


# ── CSV da planilha APORTES ───────────────────────────────────────────────────

_COLS = {
    "data":          ["data", "date", "competencia", "lancamento"],
    "investidor":    ["quem", "investidor", "investor", "socio", "nome",
                      "parceiro", "acionista", "cotista"],
    "descricao":     ["descri", "historico", "memo", "observ", "obs"],
    "banco_destino": ["banco destino", "destino", "dest"],
    "valor":         ["valor", "value", "amount", "montante", "quantia"],
}


def parse_csv_aportes(texto: str) -> list[dict]:
    """
    Lê o CSV publicado da planilha APORTES no mesmo formato que a página usa:
    pula linhas de título vazias, detecta colunas pelo cabeçalho (sem acento) e
    descarta linhas sem data válida ou sem valor numérico.
    """
    if not texto or not texto.strip():
        return []

    linhas = [l for l in csv.reader(io.StringIO(texto))]
    idx_cab = 0
    while idx_cab < len(linhas) and not any(c.strip() for c in linhas[idx_cab]):
        idx_cab += 1
    if idx_cab >= len(linhas) - 1:
        return []

    cab = [normalizar_nome(c) for c in linhas[idx_cab]]

    def achar(chaves, padrao):
        for i, c in enumerate(cab):
            for k in chaves:
                if k in c:
                    return i
        return padrao

    idx = {
        "data":          achar(_COLS["data"], 0),
        "investidor":    achar(_COLS["investidor"], 1),
        "descricao":     achar(_COLS["descricao"], 2),
        "banco_destino": achar(_COLS["banco_destino"], 4),
        "valor":         achar(_COLS["valor"], 5),
    }
    necessario = max(idx.values())

    out = []
    for linha in linhas[idx_cab + 1:]:
        if len(linha) <= necessario:
            continue
        dt = _parse_data(linha[idx["data"]])
        if dt is None:
            continue
        valor = parse_brl(linha[idx["valor"]])
        if valor is None:
            continue
        out.append({
            "data":          dt,
            "investidor":    linha[idx["investidor"]].strip(),
            "descricao":     linha[idx["descricao"]].strip(),
            "banco_destino": linha[idx["banco_destino"]].strip(),
            "valor":         valor,
        })
    return out


# ── Agregação por sócio ───────────────────────────────────────────────────────

def _grupo_de(nome_norm: str) -> list[str] | None:
    for grupo in GRUPOS_IGUALITARIOS:
        if nome_norm in grupo:
            return grupo
    return None


def calcular_tir_por_socio(
    aportes: list[dict],
    valor_veiculos: float,
    hoje: date | None = None,
) -> dict:
    """
    aportes: dicts com data (date|str), investidor, descricao, valor.
             Convenção de sinal da origem: valor > 0 é aporte (dinheiro saindo
             do bolso do sócio), valor < 0 é devolução para o sócio.
    valor_veiculos: valor total estimado dos veículos hoje, rateado entre os
             sócios conforme o peso de cada um no total aportado líquido.

    Retorna {"socios": [...], "total_aportado": float, "valor_veiculos": float,
             "descartados": {...}}
    """
    hoje = hoje or date.today()
    valor_veiculos = float(valor_veiculos or 0.0)

    descartados = {"despesa_interna": 0, "nao_socio": 0, "sem_data": 0}

    # 1) Filtra e normaliza
    validos = []
    for a in aportes:
        dt = _parse_data(a.get("data"))
        if dt is None:
            descartados["sem_data"] += 1
            continue
        if eh_despesa_interna(a.get("descricao") or ""):
            descartados["despesa_interna"] += 1
            continue
        nome_norm = normalizar_nome(a.get("investidor") or "")
        if not nome_norm or nome_norm in EXCLUIR_NOMES:
            descartados["nao_socio"] += 1
            continue
        valor = a.get("valor")
        valor = float(valor) if valor is not None else None
        if valor is None:
            continue
        validos.append({"data": dt, "nome_norm": nome_norm,
                        "nome": (a.get("investidor") or "").strip(),
                        "valor": valor})

    # 2) Agrupa: membros de um grupo igualitário compartilham os fluxos,
    #    cada um ficando com uma fração igual (1/n) do total combinado.
    chaves: dict[str, dict] = {}   # chave de agregação → dados
    for v in validos:
        grupo = _grupo_de(v["nome_norm"])
        chave = "grupo:" + "|".join(grupo) if grupo else v["nome_norm"]
        alvo = chaves.setdefault(chave, {
            "grupo":   grupo,
            "membros": {},
            "fluxos":  [],
        })
        alvo["fluxos"].append((v["data"], v["valor"]))
        alvo["membros"].setdefault(v["nome_norm"], v["nome"])

    # Um membro de grupo que não tem nenhum lançamento ainda assim participa
    for grupo in GRUPOS_IGUALITARIOS:
        chave = "grupo:" + "|".join(grupo)
        if chave in chaves:
            for m in grupo:
                chaves[chave]["membros"].setdefault(m, m.title())

    # 3) Expande em sócios individuais, aplicando a divisão igualitária
    socios = []
    for chave, dados in chaves.items():
        grupo = dados["grupo"]
        if grupo:
            nomes = [(m, dados["membros"].get(m, m.title())) for m in grupo]
            fracao = 1.0 / len(grupo)
        else:
            nomes = list(dados["membros"].items())
            fracao = 1.0
        for nome_norm, nome in nomes:
            fluxos = [(d, v * fracao) for d, v in dados["fluxos"]]
            socios.append({
                "nome":           nome,
                "nome_norm":      nome_norm,
                "compartilhado":  bool(grupo),
                "fluxos_aportes": fluxos,
                "total_aportado": sum(v for _, v in fluxos),
                "n_lancamentos":  len(fluxos),
            })

    # 4) Rateio do valor dos veículos pelo peso no total aportado líquido
    total_geral = sum(s["total_aportado"] for s in socios)
    for s in socios:
        peso = (s["total_aportado"] / total_geral) if total_geral else 0.0
        s["peso"] = peso
        s["valor_veiculos"] = valor_veiculos * peso

        # Convenção XIRR: aporte (valor > 0 na origem) é saída de caixa para o
        # sócio → negativo. Devolução (valor < 0 na origem) → positivo.
        fluxos = [(d, -v) for d, v in s["fluxos_aportes"]]
        if s["valor_veiculos"]:
            fluxos.append((hoje, s["valor_veiculos"]))

        try:
            s["tir"] = xirr(fluxos)
            s["tir_erro"] = None
        except TIRError as e:
            s["tir"] = None
            s["tir_erro"] = str(e)

        # Quanto o dinheiro dele vale hoje, o ganho nominal e por quantas vezes
        # o aporte se multiplicou — a leitura em reais, ao lado da taxa.
        s["valor_hoje"] = s["valor_veiculos"]
        s["ganho"]      = s["valor_veiculos"] - s["total_aportado"]
        s["multiplo"]   = (s["valor_veiculos"] / s["total_aportado"]
                           if s["total_aportado"] > 0 else None)
        del s["fluxos_aportes"]

    socios.sort(key=lambda s: s["total_aportado"], reverse=True)

    total_ganho = sum(s["ganho"] for s in socios)
    return {
        "socios":         socios,
        "total_aportado": total_geral,
        "total_ganho":    total_ganho,
        "valor_veiculos": valor_veiculos,
        "data_calculo":   hoje.isoformat(),
        "descartados":    descartados,
    }
