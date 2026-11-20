#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Monitor de sessões — Estudo "From Control to Memory" (Digital Tourism)
=====================================================================

Lê, em tempo real, o Google Sheets onde as páginas da tarefa gravam as respostas
(folhas "Folha1" = Encoding, "T1_Responses" e "T2_Responses") e mostra:

  1. Quem está a fazer a tarefa AGORA (fase, progresso, último registo).
  2. Sessões que pararam a meio (sem registos há mais de X minutos).
  3. Quem terminou hoje.
  4. Janela da Sessão 2 (T2): quem pode fazer agora, quem ainda tem de esperar
     e quem já passou das 48 h.

NOTA: o Google Sheets só sabe o que a página da tarefa grava. Se o upload para o
iMotions falhou, isso NÃO aparece aqui — confirmar no iMotions no fim da sessão.

-----------------------------------------------------------------------------
INSTALAÇÃO (uma vez)
    pip install pandas requests rich

COMO DAR ACESSO À FOLHA (escolher UMA das opções)

  A) Link de leitura (mais simples)
     No Google Sheets: Partilhar -> "Qualquer pessoa com o link" -> Leitor.
     Depois basta correr:
         python monitor_sessoes.py

  B) Conta de serviço Google (a folha não fica pública)
     pip install gspread google-auth
     Criar uma conta de serviço no Google Cloud, descarregar o JSON e partilhar
     a folha com o email da conta de serviço (como Leitor). Depois:
         python monitor_sessoes.py --credenciais conta_servico.json

OUTRAS OPÇÕES
    --intervalo 20          atualizar a cada 20 s (por omissão 30)
    --uma-vez               mostra uma vez e sai
    --ativo-min 3           minutos sem registos até deixar de contar como "a fazer"
    --parado-min 10         minutos sem registos para marcar como "parou a meio"
    --dia 2026-11-19        mostrar outro dia em vez de hoje
    --xlsx ficheiro.xlsx    ler de um Excel descarregado (teste / sem internet)
    --html painel.html      grava também um painel HTML (abre no browser e
                            atualiza sozinho)
-----------------------------------------------------------------------------
"""

import argparse
import io
import sys
import time
from datetime import datetime, timedelta, timezone

import pandas as pd
import requests

try:
    from zoneinfo import ZoneInfo
    TZ = ZoneInfo("Europe/Lisbon")
except Exception:  # pragma: no cover
    TZ = timezone.utc

try:
    from rich.console import Console, Group
    from rich.live import Live
    from rich.panel import Panel
    from rich.table import Table
    from rich.text import Text
    RICH = True
except ImportError:
    RICH = False

# --------------------------------------------------------------------------- #
# Configuração do estudo
# --------------------------------------------------------------------------- #
SHEET_ID = "1hGbuA8eG8SO1Spdb9ReaPDGbDJDGEyYUqsVo7CzFe2Y"
FOLHAS = {"enc": "Folha1", "t1": "T1_Responses", "t2": "T2_Responses"}
N_ENC, N_T1, N_T2 = 18, 36, 36
PREFIXO_ID = "P5_"          # só a versão final do protocolo
JANELA_T2 = (24, 48)        # horas após o fim do T1

COL_ENC = ["ts", "pid", "item", "bloco", "pos", "cor", "categoria", "controlo",
           "atraso", "rating", "trial_ms"]
COL_TEST = ["ts", "pid", "ensaio", "estimulo", "tipo", "resposta", "posicao",
            "cor", "categoria", "agencia", "confianca", "rt", "rt2"]


# --------------------------------------------------------------------------- #
# Leitura dos dados
# --------------------------------------------------------------------------- #
def _normalizar(df: pd.DataFrame, colunas) -> pd.DataFrame:
    df = df.iloc[:, : len(colunas)].copy()
    df.columns = colunas[: df.shape[1]]
    df["ts"] = pd.to_datetime(df["ts"], errors="coerce", utc=True)
    df = df[df["ts"].notna() & df["pid"].astype(str).str.startswith(PREFIXO_ID)]
    df["pid"] = df["pid"].astype(str).str.strip()
    df["resp"] = df["pid"].str.extract(r"_R(\d+)$")[0]
    return df


def ler_link(sheet_id: str):
    dados = {}
    for chave, folha in FOLHAS.items():
        url = (f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq"
               f"?tqx=out:csv&headers=0&sheet={requests.utils.quote(folha)}")
        r = requests.get(url, timeout=30)
        if r.status_code != 200 or r.text.lstrip().startswith("<"):
            raise RuntimeError(
                "Não foi possível ler a folha pelo link. Confirme que está partilhada "
                "como 'Qualquer pessoa com o link: Leitor', ou use --credenciais.")
        df = pd.read_csv(io.StringIO(r.text), header=None, dtype=str)
        dados[chave] = df
    return dados


def ler_credenciais(sheet_id: str, ficheiro_json: str):
    import gspread  # noqa: import tardio, só se for usado
    gc = gspread.service_account(filename=ficheiro_json)
    sh = gc.open_by_key(sheet_id)
    dados = {}
    for chave, folha in FOLHAS.items():
        valores = sh.worksheet(folha).get_all_values()
        dados[chave] = pd.DataFrame(valores)
    return dados


def ler_xlsx(caminho: str):
    xl = pd.read_excel(caminho, sheet_name=None, header=None, dtype=str)
    return {chave: xl[folha] for chave, folha in FOLHAS.items()}


def carregar(args):
    if args.xlsx:
        brutos = ler_xlsx(args.xlsx)
    elif args.credenciais:
        brutos = ler_credenciais(args.sheet_id, args.credenciais)
    else:
        brutos = ler_link(args.sheet_id)
    enc = _normalizar(brutos["enc"], COL_ENC)
    t1 = _normalizar(brutos["t1"], COL_TEST)
    t2 = _normalizar(brutos["t2"], COL_TEST)
    return enc, t1, t2


# --------------------------------------------------------------------------- #
# Estado de cada participante
# --------------------------------------------------------------------------- #
def resumo_fase(df, col_item, total):
    if df.empty:
        return pd.DataFrame(columns=["n", "max_ensaio", "inicio", "fim"])
    g = df.groupby("resp")
    out = pd.DataFrame({
        "n": g[col_item].nunique(),
        "inicio": g["ts"].min(),
        "fim": g["ts"].max(),
    })
    if "ensaio" in df.columns:
        out["max_ensaio"] = pd.to_numeric(df["ensaio"], errors="coerce").groupby(df["resp"]).max()
    else:
        out["max_ensaio"] = out["n"]
    out["completo"] = (out["n"] >= total) | (out["max_ensaio"] >= total)
    return out


def estado(enc, t1, t2, agora, args):
    e = resumo_fase(enc, "item", N_ENC).add_prefix("enc_")
    a = resumo_fase(t1, "estimulo", N_T1).add_prefix("t1_")
    b = resumo_fase(t2, "estimulo", N_T2).add_prefix("t2_")
    m = e.join(a, how="outer").join(b, how="outer")
    m.index.name = "resp"
    m = m.reset_index()

    linhas = []
    for _, r in m.iterrows():
        fases = []
        for f, tot in (("enc", N_ENC), ("t1", N_T1), ("t2", N_T2)):
            if pd.notna(r.get(f"{f}_fim")):
                fases.append((f, r[f"{f}_fim"], r[f"{f}_n"], tot, bool(r[f"{f}_completo"])))
        if not fases:
            continue
        fase, ultimo, n, tot, completo = max(fases, key=lambda x: x[1])
        minutos = (agora - ultimo).total_seconds() / 60
        # Os últimos 1–3 ensaios do T1/T2 às vezes não chegam à folha: se já tem
        # quase todos e parou de gravar, conta como concluído.
        if fase in ("t1", "t2") and not completo and n >= args.min_teste and minutos > args.ativo_min:
            completo = True
        nome = {"enc": "Encoding", "t1": "T1 (memória)", "t2": "T2 (sessão 2)"}[fase]

        if fase == "enc" and completo:
            situacao = "Encoding concluído → deve começar o T1"
        elif fase == "t1" and completo:
            situacao = "Sessão 1 concluída"
        elif fase == "t2" and completo:
            situacao = "Sessão 2 concluída"
        elif minutos <= args.ativo_min:
            situacao = "A fazer"
        elif minutos <= args.parado_min:
            situacao = "Sem registos há alguns minutos"
        else:
            situacao = "Parou a meio"

        # transição Encoding -> T1 que não arrancou
        if fase == "enc" and completo and minutos > args.parado_min:
            situacao = "Encoding concluído, T1 não começou"

        linhas.append({
            "resp": r["resp"], "fase": nome, "fase_id": fase, "n": int(n), "total": tot,
            "completo": completo, "ultimo": ultimo, "min": minutos, "situacao": situacao,
            "t1_fim": r.get("t1_fim"),
            "t1_completo": (bool(r.get("t1_completo")) if pd.notna(r.get("t1_completo")) else False)
                           or (pd.notna(r.get("t1_n")) and r.get("t1_n") >= args.min_teste
                               and (agora - r.get("t1_fim")).total_seconds() / 60 > args.ativo_min),
            "t2_n": r.get("t2_n"),
        })
    return pd.DataFrame(linhas)


def janela_t2(est, agora):
    if est.empty:
        return est
    j = est[est["t1_completo"] & est["t2_n"].isna()].copy()
    if j.empty:
        return j
    j["abre"] = j["t1_fim"] + pd.Timedelta(hours=JANELA_T2[0])
    j["fecha"] = j["t1_fim"] + pd.Timedelta(hours=JANELA_T2[1])

    def s(r):
        if agora < r["abre"]:
            return "Ainda não pode"
        if agora <= r["fecha"]:
            return "PODE FAZER AGORA"
        return "Fora da janela (>48 h)"
    j["estado_t2"] = j.apply(s, axis=1)
    return j.sort_values("abre")


# --------------------------------------------------------------------------- #
# Apresentação
# --------------------------------------------------------------------------- #
def hora(ts):
    return ts.tz_convert(TZ).strftime("%H:%M:%S") if pd.notna(ts) else ""


def dia_hora(ts):
    return ts.tz_convert(TZ).strftime("%d/%m %H:%M") if pd.notna(ts) else ""


def barra(n, total, largura=18):
    """Barra de progresso que se lê bem em qualquer terminal (incl. Windows)."""
    if RICH:
        from rich.progress_bar import ProgressBar
        return ProgressBar(total=total, completed=min(n, total), width=largura,
                           complete_style="green" if n >= total else "cyan",
                           finished_style="green")
    cheio = int(round(largura * min(n, total) / total))
    return "#" * cheio + "." * (largura - cheio)


def tempo(minutos):
    """'4 min', '2 h 05 min', '1 d 3 h'."""
    m = int(round(max(minutos, 0)))
    if m < 60:
        return f"{m} min"
    h, m = divmod(m, 60)
    if h < 24:
        return f"{h} h {m:02d} min"
    d, h = divmod(h, 24)
    return f"{d} d {h} h"


def selecionar_dia(est, dia):
    if est.empty:
        return est
    d = est["ultimo"].dt.tz_convert(TZ).dt.date
    return est[d == dia]


def montar_rich(est, jan, agora, dia, args, erro=None):
    hoje = selecionar_dia(est, dia)
    ativos = hoje[hoje["situacao"].isin(["A fazer", "Sem registos há alguns minutos",
                                          "Encoding concluído → deve começar o T1"])]
    parados = hoje[hoje["situacao"].isin(["Parou a meio", "Encoding concluído, T1 não começou"])]
    feitos = hoje[hoje["situacao"].isin(["Sessão 1 concluída", "Sessão 2 concluída"])]

    cab = Text()
    cab.append("Estudo Memória — monitor de sessões   ", style="bold")
    cab.append(f"{agora.astimezone(TZ):%d/%m/%Y %H:%M:%S}   ", style="cyan")
    cab.append(f"A fazer: {len(ativos)}  ", style="bold green")
    cab.append(f"Pararam: {len(parados)}  ", style="bold red")
    cab.append(f"Concluídos hoje: {len(feitos)}", style="bold blue")
    if erro:
        cab.append(f"\n⚠ {erro}", style="bold red")

    def tabela(df, titulo, estilo):
        t = Table(title=titulo, title_style=estilo, expand=True, show_lines=False)
        for c in ["Resp.", "Fase", "Progresso", "Ensaios", "Último registo", "Há", "Situação"]:
            t.add_column(c)
        for _, r in df.sort_values("ultimo", ascending=False).iterrows():
            cor = {"A fazer": "green", "Parou a meio": "red",
                   "Encoding concluído, T1 não começou": "red",
                   "Sem registos há alguns minutos": "yellow"}.get(r["situacao"], "white")
            t.add_row(f"R{r['resp']}", r["fase"], barra(r["n"], r["total"]),
                      f"{r['n']}/{r['total']}", hora(r["ultimo"]),
                      tempo(r["min"]), Text(r["situacao"], style=cor))
        if df.empty:
            t.add_row("—", "", "", "", "", "", "")
        return t

    jt = Table(title="Janela da Sessão 2 (T2: 24–48 h após o fim do T1)", title_style="bold magenta", expand=True)
    for c in ["Resp.", "Fim do T1", "Abre", "Fecha", "Estado"]:
        jt.add_column(c)
    if jan is not None and not jan.empty:
        for _, r in jan.iterrows():
            cor = {"PODE FAZER AGORA": "bold green", "Ainda não pode": "yellow"}.get(r["estado_t2"], "dim")
            if r["estado_t2"].startswith("Fora") and not args.mostrar_fora:
                continue
            jt.add_row(f"R{r['resp']}", dia_hora(r["t1_fim"]), dia_hora(r["abre"]),
                       dia_hora(r["fecha"]), Text(r["estado_t2"], style=cor))
    else:
        jt.add_row("—", "", "", "", "")

    return Group(Panel(cab), tabela(ativos, "A fazer agora", "bold green"),
                 tabela(parados, "Pararam a meio (hoje)", "bold red"),
                 tabela(feitos, "Concluídos hoje", "bold blue"), jt)


def texto_simples(est, jan, agora, dia, args):
    hoje = selecionar_dia(est, dia)
    linhas = [f"=== Monitor de sessões — {agora.astimezone(TZ):%d/%m/%Y %H:%M:%S} ==="]
    for _, r in hoje.sort_values("ultimo", ascending=False).iterrows():
        linhas.append(f"R{r['resp']:>4}  {r['fase']:<14} {r['n']:>2}/{r['total']}  "
                      f"último {hora(r['ultimo'])} (há {tempo(r['min'])})  {r['situacao']}")
    if jan is not None and not jan.empty:
        linhas.append("--- Janela T2 ---")
        for _, r in jan.iterrows():
            if r["estado_t2"].startswith("Fora") and not args.mostrar_fora:
                continue
            linhas.append(f"R{r['resp']:>4}  abre {dia_hora(r['abre'])}  fecha {dia_hora(r['fecha'])}  {r['estado_t2']}")
    return "\n".join(linhas)


def gravar_html(caminho, est, jan, agora, dia, args):
    hoje = selecionar_dia(est, dia)
    cores = {"A fazer": "#d4f7d4", "Parou a meio": "#ffd6d6",
             "Encoding concluído, T1 não começou": "#ffd6d6",
             "Sem registos há alguns minutos": "#fff3c4",
             "Sessão 1 concluída": "#dbe9ff", "Sessão 2 concluída": "#dbe9ff"}
    linhas = ""
    for _, r in hoje.sort_values("ultimo", ascending=False).iterrows():
        pct = 100 * min(r["n"], r["total"]) / r["total"]
        linhas += (f"<tr style='background:{cores.get(r['situacao'], '#fff')}'>"
                   f"<td>R{r['resp']}</td><td>{r['fase']}</td>"
                   f"<td><div class='b'><div style='width:{pct:.0f}%'></div></div> {r['n']}/{r['total']}</td>"
                   f"<td>{hora(r['ultimo'])}</td><td>{tempo(r['min'])}</td><td>{r['situacao']}</td></tr>")
    jl = ""
    if jan is not None and not jan.empty:
        for _, r in jan.iterrows():
            if r["estado_t2"].startswith("Fora") and not args.mostrar_fora:
                continue
            jl += (f"<tr><td>R{r['resp']}</td><td>{dia_hora(r['t1_fim'])}</td><td>{dia_hora(r['abre'])}</td>"
                   f"<td>{dia_hora(r['fecha'])}</td><td>{r['estado_t2']}</td></tr>")
    html = f"""<!doctype html><html lang="pt"><head><meta charset="utf-8">
<meta http-equiv="refresh" content="{args.intervalo}"><title>Monitor de sessões</title>
<style>body{{font-family:system-ui,sans-serif;margin:24px;color:#222}}table{{border-collapse:collapse;width:100%;margin-bottom:28px}}
td,th{{border-bottom:1px solid #ddd;padding:6px 8px;text-align:left;font-size:14px}}th{{background:#f3f3f3}}
.b{{display:inline-block;width:120px;height:10px;background:#eee;border-radius:5px;vertical-align:middle}}
.b div{{height:10px;background:#3a7bd5;border-radius:5px}}</style></head><body>
<h2>Estudo Memória — sessões de {dia:%d/%m/%Y}</h2><p>Atualizado às {agora.astimezone(TZ):%H:%M:%S} (atualiza a cada {args.intervalo} s)</p>
<table><tr><th>Resp.</th><th>Fase</th><th>Progresso</th><th>Último registo</th><th>Há</th><th>Situação</th></tr>{linhas or '<tr><td colspan=6>Sem registos hoje</td></tr>'}</table>
<h3>Janela da Sessão 2 (24–48 h após o fim do T1)</h3>
<table><tr><th>Resp.</th><th>Fim do T1</th><th>Abre</th><th>Fecha</th><th>Estado</th></tr>{jl or '<tr><td colspan=5>—</td></tr>'}</table>
<p style="color:#777">O upload para o iMotions não aparece aqui: confirme no iMotions no fim de cada sessão.</p>
</body></html>"""
    with open(caminho, "w", encoding="utf-8") as f:
        f.write(html)


# --------------------------------------------------------------------------- #
def main():
    p = argparse.ArgumentParser(description="Monitor de sessões do Estudo Memória (Google Sheets).")
    p.add_argument("--sheet-id", default=SHEET_ID)
    p.add_argument("--credenciais", help="JSON da conta de serviço Google (opção B)")
    p.add_argument("--xlsx", help="ler de um ficheiro Excel em vez do Google Sheets")
    p.add_argument("--intervalo", type=int, default=30, help="segundos entre atualizações")
    p.add_argument("--ativo-min", type=float, default=3)
    p.add_argument("--parado-min", type=float, default=10)
    p.add_argument("--min-teste", type=int, default=33,
                   help="n.º mínimo de ensaios (de 36) para dar o T1/T2 como concluído quando deixa de gravar")
    p.add_argument("--dia", help="AAAA-MM-DD (por omissão: hoje)")
    p.add_argument("--agora", help="simular a hora atual, ex. '2026-11-19 21:15' (testes)")
    p.add_argument("--uma-vez", action="store_true")
    p.add_argument("--mostrar-fora", action="store_true", help="mostrar também quem já passou das 48 h")
    p.add_argument("--html", help="gravar também um painel HTML neste ficheiro")
    args = p.parse_args()

    def ciclo():
        agora = (pd.Timestamp(args.agora, tz=TZ).tz_convert("UTC") if args.agora
                 else pd.Timestamp.now(tz="UTC"))
        dia = (datetime.strptime(args.dia, "%Y-%m-%d").date() if args.dia
               else agora.tz_convert(TZ).date())
        enc, t1, t2 = carregar(args)
        if args.agora:  # simulação: ignorar registos posteriores à hora simulada
            enc, t1, t2 = (d[d["ts"] <= agora] for d in (enc, t1, t2))
        est = estado(enc, t1, t2, agora, args)
        jan = janela_t2(est, agora)
        if args.html:
            gravar_html(args.html, est, jan, agora, dia, args)
        return est, jan, agora, dia

    if not RICH or args.uma_vez:
        while True:
            try:
                est, jan, agora, dia = ciclo()
                if RICH:
                    Console().print(montar_rich(est, jan, agora, dia, args))
                else:
                    print(texto_simples(est, jan, agora, dia, args))
            except Exception as ex:
                print(f"Erro: {ex}", file=sys.stderr)
            if args.uma_vez:
                return
            time.sleep(args.intervalo)

    console = Console()
    ultimo = None
    with Live(console=console, refresh_per_second=1, screen=False) as live:
        while True:
            try:
                ultimo = ciclo()
                live.update(montar_rich(*ultimo, args))
            except Exception as ex:
                if ultimo:
                    live.update(montar_rich(*ultimo, args, erro=f"Falha na atualização: {ex}"))
                else:
                    live.update(Panel(Text(f"Erro: {ex}", style="bold red")))
            time.sleep(args.intervalo)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
