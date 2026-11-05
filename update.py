#!/usr/bin/env python3
"""Predictor Real Madrid (LaLiga y Champions) - actualización diaria."""

import argparse
import io
import math
import os
import sys
import datetime as dt
from pathlib import Path

import numpy as np
import pandas as pd

try:
    import requests
except ImportError:
    requests = None

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
OUT = BASE / "predictor_real_madrid.xlsx"
MADRID = "Real Madrid"
URL_SEASON = "https://www.football-data.co.uk/mmz4281/{code}/SP1.csv"
URL_FIXTURES = "https://www.football-data.co.uk/fixtures.csv"

HFA = 65.0
ELO_K = 20.0
ELO_NEW = 1450.0
ELO_REG = 0.20
DRAW_D = 0.27
DRAW_S = 400.0
HALF_LIFE = 300.0
PRIOR_GAMES = 3.0
BAJAS_K = 0.03
N_HIST = 10
N_BACKTEST = 4

NOMBRES = {
    "Ath Madrid": "Atlético de Madrid", "Ath Bilbao": "Athletic Club", "Sociedad": "Real Sociedad",
    "Betis": "Real Betis", "Vallecano": "Rayo Vallecano", "Celta": "Celta de Vigo",
    "Barcelona": "FC Barcelona", "Espanol": "Espanyol", "Alaves": "Alavés", "Cadiz": "Cádiz",
    "Almeria": "Almería", "Santander": "Racing de Santander", "La Coruna": "Deportivo La Coruña",
    "Leganes": "Leganés", "Malaga": "Málaga", "Sp Gijon": "Sporting de Gijón",
    "Oviedo": "Real Oviedo", "Mallorca": "Mallorca", "Valladolid": "Real Valladolid",
}

SEED_HIST = [
    ("2016-17", 29, 6, 3), ("2017-18", 22, 10, 6), ("2018-19", 21, 5, 12),
    ("2019-20", 26, 9, 3), ("2020-21", 25, 9, 4), ("2021-22", 26, 8, 4),
    ("2022-23", 24, 5, 9), ("2023-24", 29, 8, 1), ("2024-25", 26, 6, 6),
    ("2025-26", 27, 5, 6),
]
SEED_2627 = [
    (1, None, "Espanyol", None, None, None, "V"),
    (2, None, "Real Sociedad", None, None, None, "V"),
    (3, None, "Málaga", "L", None, None, "V"),
    (4, None, "Real Betis", "V", None, None, "D"),
    (5, None, "Rayo Vallecano", "L", 4, 1, "V"),
    (6, None, "Elche", "V", None, None, "V"),
    (7, None, "Atlético de Madrid", "V", 1, 2, "D"),
    (8, None, "Villarreal", "L", None, None, "V"),
    (9, None, "Sevilla", "L", 4, 1, "V"),
    (10, None, "FC Barcelona", "V", 2, 1, "V"),
    (11, "2026-11-01", "Racing de Santander", "V", 3, 0, "V"),
]
SEED_BAJAS = [
    ("Raúl Asencio", "Baja", "Fractura por estrés de tibia derecha (operado 30-sep-2026)",
     "varios meses", 0.3, "Prensa 30-sep-2026; importancia estimada por defecto"),
]

def nombre(t):
    return NOMBRES.get(t, t)

def season_code(y):
    return f"{y % 100:02d}{(y + 1) % 100:02d}"

def parse_dates(s):
    try:
        return pd.to_datetime(s, dayfirst=True, errors="coerce", format="mixed")
    except (TypeError, ValueError):
        return pd.to_datetime(s, dayfirst=True, errors="coerce")

def fetch(url, path, refresh):
    if refresh and requests is not None:
        try:
            r = requests.get(url, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
            r.raise_for_status()
            path.write_bytes(r.content)
            return True
        except Exception as e:
            print(f"  aviso: no se pudo descargar {url}: {e}")
    return path.exists()

def load_season(y, download):
    path = DATA / f"SP1_{season_code(y)}.csv"
    if download:
        fetch(URL_SEASON.format(code=season_code(y)), path, True)
    if not path.exists():
        return None
    df = pd.read_csv(path, encoding="latin-1", on_bad_lines="skip")
    need = ["Date", "HomeTeam", "AwayTeam", "FTHG", "FTAG"]
    if any(c not in df.columns for c in need):
        return None
    df = df[need].dropna()
    df["Date"] = parse_dates(df["Date"])
    df = df.dropna(subset=["Date"])
    df["FTHG"] = df["FTHG"].astype(int)
    df["FTAG"] = df["FTAG"].astype(int)
    df["Season"] = y
    return df

def load_all(today, refresh):
    DATA.mkdir(exist_ok=True)
    cur = today.year if today.month >= 7 else today.year - 1
    frames = []
    for y in range(cur - N_HIST, cur + 1):
        need = refresh and (y == cur or not (DATA / f"SP1_{season_code(y)}.csv").exists())
        df = load_season(y, need)
        if df is not None and len(df):
            frames.append(df)
    if not frames:
        return None, cur
    return pd.concat(frames).sort_values("Date").reset_index(drop=True), cur

def load_fixtures(refresh):
    path = DATA / "fixtures.csv"
    if refresh:
        fetch(URL_FIXTURES, path, True)
    if not path.exists():
        return pd.DataFrame()
    try:
        fx = pd.read_csv(path, encoding="latin-1", on_bad_lines="skip")
        fx = fx[fx["Div"] == "SP1"].copy()
        fx["Date"] = parse_dates(fx["Date"])
        return fx
    except Exception:
        return pd.DataFrame()

def elo_probs(dr):
    e = 1.0 / (1.0 + 10 ** (-dr / 400.0))
    pdw = DRAW_D * math.exp(-((dr / DRAW_S) ** 2))
    pw = max(e - pdw / 2, 0.01)
    pl = max(1 - e - pdw / 2, 0.01)
    s = pw + pdw + pl
    return pw / s, pdw / s, pl / s

def run_elo(df):
    df = df.sort_values("Date").reset_index(drop=True)
    ratings, eh, ea, last = {}, [], [], None
    for r in df.itertuples():
        if last is not None and r.Season != last:
            for t in ratings:
                ratings[t] = 1500 + (1 - ELO_REG) * (ratings[t] - 1500)
        last = r.Season
        rh, ra = ratings.get(r.HomeTeam, ELO_NEW), ratings.get(r.AwayTeam, ELO_NEW)
        eh.append(rh)
        ea.append(ra)
        dr = rh + HFA - ra
        exp = 1.0 / (1.0 + 10 ** (-dr / 400.0))
        s = 1.0 if r.FTHG > r.FTAG else 0.5 if r.FTHG == r.FTAG else 0.0
        gd = abs(r.FTHG - r.FTAG)
        mult = 1.0 if gd <= 1 else 1.5 if gd == 2 else (11 + gd) / 8
        d = ELO_K * mult * (s - exp)
        ratings[r.HomeTeam] = rh + d
        ratings[r.AwayTeam] = ra - d
    df["EloH"], df["EloA"] = eh, ea
    return df, ratings

def fit_strengths(df, ref_date, years=3):
    d = df[(df.Date < ref_date) & (df.Date >= ref_date - pd.Timedelta(days=365 * years))]
    if len(d) < 80:
        return None
    w = 0.5 ** ((ref_date - d.Date).dt.days.values / HALF_LIFE)
    teams = sorted(set(d.HomeTeam) | set(d.AwayTeam))
    idx = {t: i for i, t in enumerate(teams)}
    hi = np.array([idx[t] for t in d.HomeTeam])
    ai = np.array([idx[t] for t in d.AwayTeam])
    hg, ag = d.FTHG.values.astype(float), d.FTAG.values.astype(float)
    bh, ba = (w * hg).sum() / w.sum(), (w * ag).sum() / w.sum()
    c = (bh + ba) / 2
    n = len(teams)
    att, dfn = np.ones(n), np.ones(n)
    for _ in range(30):
        num = np.bincount(hi, w * hg, n) + np.bincount(ai, w * ag, n)
        den = np.bincount(hi, w * bh * dfn[ai], n) + np.bincount(ai, w * ba * dfn[hi], n)
        att = (num + PRIOR_GAMES * c) / (den + PRIOR_GAMES * c)
        att /= att.mean()
        num = np.bincount(hi, w * ag, n) + np.bincount(ai, w * hg, n)
        den = np.bincount(hi, w * ba * att[ai], n) + np.bincount(ai, w * bh * att[hi], n)
        dfn = (num + PRIOR_GAMES * c) / (den + PRIOR_GAMES * c)
        dfn /= dfn.mean()
    return {"bh": bh, "ba": ba, "att": dict(zip(teams, att)), "dfn": dict(zip(teams, dfn))}

def poisson_probs(model, home, away, adj_att=1.0, adj_def=1.0):
    mu_h = model["bh"] * model["att"].get(home, 0.85) * model["dfn"].get(away, 1.15)
    mu_a = model["ba"] * model["att"].get(away, 0.85) * model["dfn"].get(home, 1.15)
    if home == MADRID:
        mu_h *= adj_att
        mu_a *= adj_def
    elif away == MADRID:
        mu_a *= adj_att
        mu_h *= adj_def
    k = np.arange(10)
    fact = np.array([math.factorial(int(i)) for i in k], dtype=float)
    ph = np.exp(-mu_h) * mu_h ** k / fact
    pa = np.exp(-mu_a) * mu_a ** k / fact
    m = np.outer(ph, pa)
    m /= m.sum()
    return np.tril(m, -1).sum(), np.trace(m), np.triu(m, 1).sum(), mu_h, mu_a

def madrid_matches(df):
    m = df[(df.HomeTeam == MADRID) | (df.AwayTeam == MADRID)].copy()
    m["Home"] = m.HomeTeam == MADRID
    m["Rival"] = np.where(m.Home, m.AwayTeam, m.HomeTeam)
    m["GF"] = np.where(m.Home, m.FTHG, m.FTAG)
    m["GC"] = np.where(m.Home, m.FTAG, m.FTHG)
    m["Res"] = np.where(m.GF > m.GC, "V", np.where(m.GF == m.GC, "E", "D"))
    return m

def to_madrid(p3, home):
    return (p3[0], p3[1], p3[2]) if home else (p3[2], p3[1], p3[0])

def score(P, y):
    P = np.clip(np.asarray(P, float), 1e-6, 1.0)
    P = P / P.sum(1, keepdims=True)
    oh = np.eye(3)[y]
    return {
        "acc": float((P.argmax(1) == y).mean()),
        "brier": float(((P - oh) ** 2).sum(1).mean()),
        "ll": float(-np.log(P[np.arange(len(y)), y]).mean()),
        "n": int(len(y)),
    }

def backtest(df, cur):
    test_years = list(range(cur - N_BACKTEST, cur))
    mm = madrid_matches(df)
    rows = []
    for r in mm[mm.Season.isin(test_years)].itertuples():
        model = fit_strengths(df, r.Date)
        if model is None:
            continue
        pp = poisson_probs(model, r.HomeTeam, r.AwayTeam)[:3]
        pe = elo_probs(r.EloH + HFA - r.EloA)
        yidx = {"V": 0, "E": 1, "D": 2}[r.Res]
        prev = mm[mm.Season < r.Season]
        base = [(prev.Res == k).mean() for k in ("V", "E", "D")] if len(prev) else [1 / 3] * 3
        rows.append({
            "Season": r.Season, "Fecha": r.Date, "Rival": r.Rival, "Home": r.Home, "Res": r.Res,
            "y": yidx, "pois": to_madrid(pp, r.Home), "elo": to_madrid(pe, r.Home), "base": tuple(base),
        })
    if not rows:
        return None
    y = np.array([r["y"] for r in rows])
    Pp = np.array([r["pois"] for r in rows])
    Pe = np.array([r["elo"] for r in rows])
    Pb = np.array([r["base"] for r in rows])
    best_w, best_ll = 0.5, 9e9
    for w in np.linspace(0, 1, 11):
        ll = score(w * Pp + (1 - w) * Pe, y)["ll"]
        if ll < best_ll:
            best_w, best_ll = float(w), ll
    Pm = best_w * Pp + (1 - best_w) * Pe
    res = {
        "w": best_w, "rows": rows, "Pm": Pm,
        "mixto": score(Pm, y), "poisson": score(Pp, y), "elo": score(Pe, y), "base": score(Pb, y),
    }
    seasons = []
    for s in test_years:
        idx = [i for i, r in enumerate(rows) if r["Season"] == s]
        if not idx:
            continue
        real = int(sum(1 for i in idx if rows[i]["y"] == 0))
        esp = float(Pm[idx, 0].sum())
        base = float(Pb[idx, 0].sum())
        seasons.append((f"{s}-{(s + 1) % 100:02d}", len(idx), real, esp, base))
    res["seasons"] = seasons
    res["mae_model"] = float(np.mean([abs(s[3] - s[2]) for s in seasons]))
    res["mae_base"] = float(np.mean([abs(s[4] - s[2]) for s in seasons]))
    return res

def read_bajas(path):
    csvp = BASE / "bajas.csv"
    if csvp.exists():
        try:
            d = pd.read_csv(csvp, encoding="utf-8").fillna("")
            rows = []
            for _, r in d.iterrows():
                v = list(r.values) + [""] * 6
                if str(v[0]).strip():
                    rows.append(v[:6])
            if rows:
                return rows
        except Exception as e:
            print("  aviso: no se pudo leer bajas.csv:", e)
    rows = []
    if path.exists():
        try:
            ws = load_workbook(path)["Bajas"]
            for r in ws.iter_rows(min_row=2, values_only=True):
                if r and r[0]:
                    rows.append(list(r[:6]) + [None] * (6 - len(r[:6])))
        except Exception:
            pass
    return rows or [list(x) for x in SEED_BAJAS]

def api_football(today, bajas):
    key = os.environ.get("API_FOOTBALL_KEY")
    if not key or requests is None:
        return bajas, []
    h = {"x-apisports-key": key}
    y = today.year if today.month >= 7 else today.year - 1
    jugadores = []
    try:
        r = requests.get("https://v3.football.api-sports.io/injuries",
                         params={"team": 541}, headers=h, timeout=30).json()
        if r.get("errors"):
            print("  API-Football (lesiones):", r["errors"])
        latest = {}
        for it in r.get("response", []):
            f = pd.to_datetime(it["fixture"]["date"]).tz_localize(None)
            if f >= pd.Timestamp(today) - pd.Timedelta(days=7):
                latest[it["player"]["name"]] = (f, it["player"].get("type", ""), it["player"].get("reason", ""))
        by_name = {b[0]: b for b in bajas}
        for name, (_, tipo, motivo) in latest.items():
            estado = "Duda" if "Questionable" in str(tipo) else "Baja"
            if name in by_name:
                by_name[name][1], by_name[name][2], by_name[name][5] = estado, motivo, "API-Football"
            else:
                bajas.append([name, estado, motivo, "", 0.5, "API-Football; importancia por defecto"])
        page, total = 1, 1
        while page <= total and page <= 3:
            r = requests.get("https://v3.football.api-sports.io/players",
                             params={"team": 541, "season": y, "page": page},
                             headers=h, timeout=30).json()
            total = r.get("paging", {}).get("total", 1)
            for it in r.get("response", []):
                st = it["statistics"][0]
                mins = st["games"].get("minutes") or 0
                jugadores.append([
                    it["player"]["name"], mins, st["games"].get("appearences") or 0,
                    float(st["games"].get("rating") or 0), st["goals"].get("total") or 0,
                    st["goals"].get("assists") or 0,
                ])
            page += 1
        if jugadores:
            mx = max(j[1] for j in jugadores) or 1
            imp = {j[0]: round(min(1.0, j[1] / mx), 2) for j in jugadores}
            for b in bajas:
                if b[0] in imp and "por defecto" in str(b[5]):
                    b[4] = imp[b[0]]
    except Exception as e:
        print("  aviso: API-Football no disponible:", e)
    return bajas, jugadores

PESO_ESTADO = {"baja": 1.0, "sancion": 1.0, "sanción": 1.0, "duda": 0.5}

def indice_bajas(bajas):
    total = 0.0
    for b in bajas:
        w = PESO_ESTADO.get(str(b[1]).strip().lower(), 0.0)
        try:
            total += w * float(b[4] or 0)
        except (TypeError, ValueError):
            pass
    return total

def predecir(df, cur, today, ratings, w, L, fixtures):
    now = pd.Timestamp(today) + pd.Timedelta(days=1)
    model = fit_strengths(df, now)
    if model is None:
        return None
    cd = df[df.Season == cur]
    played_h = set(cd[cd.HomeTeam == MADRID].AwayTeam)
    played_a = set(cd[cd.AwayTeam == MADRID].HomeTeam)
    adj_att, adj_def = math.exp(-BAJAS_K * L), math.exp(BAJAS_K * 0.5 * L)

    scheduled_matches = []
    if len(fixtures):
        for r in fixtures.itertuples():
            h_team = nombre(str(getattr(r, "HomeTeam", "")))
            a_team = nombre(str(getattr(r, "AwayTeam", "")))
            r_date = getattr(r, "Date", pd.NaT)
            if h_team == MADRID or a_team == MADRID:
                is_home = (h_team == MADRID)
                rival = a_team if is_home else h_team
                already_played = (is_home and rival in played_h) or (not is_home and rival in played_a)
                if not already_played:
                    scheduled_matches.append({
                        "rival": rival,
                        "home": is_home,
                        "fecha": r_date if pd.notna(r_date) else None
                    })

    teams = sorted((set(cd.HomeTeam) | set(cd.AwayTeam)) - {MADRID})
    existing_pairs = {(m["rival"], m["home"]) for m in scheduled_matches}
    for t in teams:
        for home in (True, False):
            if (home and t in played_h) or (not home and t in played_a):
                continue
            if (nombre(t), home) not in existing_pairs:
                scheduled_matches.append({
                    "rival": nombre(t),
                    "home": home,
                    "fecha": None
                })

    rows = []
    for match in scheduled_matches:
        t_name = match["rival"]
        home = match["home"]
        t_orig = next((k for k, v in NOMBRES.items() if v == t_name), t_name)

        h, a = (MADRID, t_orig) if home else (t_orig, MADRID)
        pp = poisson_probs(model, h, a, adj_att, adj_def)
        re, rr = ratings.get(MADRID, 1500.0), ratings.get(t_orig, ELO_NEW)
        pe = elo_probs((re + HFA - rr) if home else (rr + HFA - re))
        pm = to_madrid(pp[:3], home)
        pel = to_madrid(pe, home)
        p = [w * pm[i] + (1 - w) * pel[i] for i in range(3)]
        gm, gr = (pp[3], pp[4]) if home else (pp[4], pp[3])

        rows.append({
            "rival": t_name,
            "home": home,
            "fecha": match["fecha"],
            "p": p,
            "gm": gm,
            "gr": gr,
            "elo": rr
        })

    return rows

def distribucion(pw):
    d = np.array([1.0])
    for p in pw:
        d = np.convolve(d, [1 - p, p])
    return d

ARIAL = "Arial"
HEAD = PatternFill("solid", fgColor="1F3864")
SOFT = PatternFill("solid", fgColor="DDEBF7")
INPUT = PatternFill("solid", fgColor="FFFF99")

def hdr(ws, row, labels, col=1):
    for i, t in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=t)
        c.font = Font(name=ARIAL, bold=True, color="FFFFFF")
        c.fill = HEAD
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def style_all(ws):
    for row in ws.iter_rows():
        for c in row:
            if c.font is None or c.font.name != ARIAL:
                c.font = Font(name=ARIAL, size=c.font.size or 10, bold=c.font.bold, color=c.font.color,
                              italic=c.font.italic)

def widths(ws, ws_w):
    for i, w in enumerate(ws_w, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def write_workbook(path, hist, temporada, bajas, jugadores, pred, bt, meta):
    wb = Workbook()
    wsH = wb.active
    wsH.title = "Historico"
    hdr(wsH, 1, ["Temporada", "PJ", "Victorias", "Empates", "Derrotas", "Puntos", "% Victorias", "Peso (previsión)"])
    pesos = [0.025] * 4 + [0.10] * 2 + [0.15] * 2 + [0.20] * 2
    for i, (t, v, e, d) in enumerate(hist):
        r = i + 2
        wsH.cell(r, 1, t)
        wsH.cell(r, 2, f"=C{r}+D{r}+E{r}")
        for j, val in enumerate((v, e, d), 3):
            c = wsH.cell(r, j, val)
            c.font = Font(name=ARIAL, color="0000FF")
        wsH.cell(r, 6, f"=3*C{r}+D{r}")
        wsH.cell(r, 7, f"=C{r}/B{r}").number_format = "0.0%"
        c = wsH.cell(r, 8, pesos[i] if len(hist) == 10 else round(1 / len(hist), 3))
        c.font = Font(name=ARIAL, color="0000FF")
        c.number_format = "0.0%"
    n = len(hist)
    last = n + 1
    wsH.cell(last + 1, 1, "Media / total").font = Font(name=ARIAL, bold=True)
    wsH.cell(last + 1, 2, f"=SUM(B2:B{last})")
    wsH.cell(last + 1, 3, f"=AVERAGE(C2:C{last})").number_format = "0.0"
    wsH.cell(last + 1, 7, f"=SUM(C2:C{last})/SUM(B2:B{last})").number_format = "0.0%"
    wsH.cell(last + 1, 8, f"=SUM(H2:H{last})").number_format = "0.0%"
    wsH.cell(last + 2, 1, "Tasa de victoria ponderada").font = Font(name=ARIAL, bold=True)
    wsH.cell(last + 2, 7, f"=SUMPRODUCT(G2:G{last},H2:H{last})").number_format = "0.0%"
    widths(wsH, [16, 8, 11, 10, 10, 9, 12, 16])
    ch = BarChart()
    ch.type = "col"
    ch.title = "Victorias del Real Madrid en LaLiga por temporada"
    ch.add_data(Reference(wsH, min_col=3, min_row=1, max_row=last), titles_from_data=True)
    ch.set_categories(Reference(wsH, min_col=1, min_row=2, max_row=last))
    ch.height, ch.width, ch.legend = 8, 16, None
    wsH.add_chart(ch, "J2")
    HROW = last + 2

    wsT = wb.create_sheet("Temporada_actual")
    hdr(wsT, 1, ["Jornada", "Fecha", "Rival", "Condición (L/V)", "GF", "GC", "Resultado (V/E/D)", "Nota"])
    for i, row in enumerate(temporada):
        r = i + 2
        for j, v in enumerate(row[:7], 1):
            if v is not None:
                wsT.cell(r, j, v)
        wsT.cell(r, 8, row[7] if len(row) > 7 else None)
    widths(wsT, [9, 12, 24, 16, 6, 6, 18, 40])

    wsP = wb.create_sheet("Predicciones")
    hdr(wsP, 1, ["Rival", "Condición", "Fecha", "P victoria", "P empate", "P derrota", "Goles esp. Madrid", "Goles esp. rival", "Elo rival"])
    labels = ["P10 victorias finales", "P50 victorias finales", "P90 victorias finales", "Puntos esperados finales",
              "Índice de bajas (L)", "Peso Poisson (w)", "Elo Real Madrid", "Actualizado"]
    for i, t in enumerate(labels):
        wsP.cell(1 + i, 11, t).font = Font(name=ARIAL, bold=True)
    nlabels = ["Próximo partido", "Rival", "Condición", "Fecha", "P victoria", "P empate", "P derrota"]
    for i, t in enumerate(nlabels):
        wsP.cell(1 + i, 14, t).font = Font(name=ARIAL, bold=True)
    if pred is not None:
        for i, r in enumerate(pred["rows"]):
            rr = i + 2
            wsP.cell(rr, 1, r["rival"])
            wsP.cell(rr, 2, "L" if r["home"] else "V")
            if r["fecha"] is not None and pd.notna(r["fecha"]):
                try:
                    wsP.cell(rr, 3, r["fecha"].to_pydatetime()).number_format = "dd/mm/yyyy"
                except Exception:
                    wsP.cell(rr, 3, str(r["fecha"]))
            for j in range(3):
                wsP.cell(rr, 4 + j, round(float(r["p"][j]), 4)).number_format = "0.0%"
            wsP.cell(rr, 7, round(float(r["gm"]), 2))
            wsP.cell(rr, 8, round(float(r["gr"]), 2))
            wsP.cell(rr, 9, round(float(r["elo"]), 0))
        for i, v in enumerate(pred["resumen"]):
            c = wsP.cell(1 + i, 12, v)
            if i in (3, 4, 5, 6):
                c.number_format = "0.00"
        nx = pred.get("proximo")
        if nx:
            wsP.cell(2, 15, nx["rival"])
            wsP.cell(3, 15, "L" if nx["home"] else "V")
            if nx["fecha"] is not None and pd.notna(nx["fecha"]):
                try:
                    wsP.cell(4, 15, nx["fecha"].to_pydatetime()).number_format = "dd/mm/yyyy"
                except Exception:
                    wsP.cell(4, 15, str(nx["fecha"]))
            for j in range(3):
                wsP.cell(5 + j, 15, round(float(nx["p"][j]), 4)).number_format = "0.0%"
    widths(wsP, [24, 10, 12, 11, 11, 11, 16, 15, 10, 3, 26, 18, 3, 18, 22])

    wsB = wb.create_sheet("Backtest")
    hdr(wsB, 1, ["Modelo", "Precisión", "Brier", "LogLoss", "N partidos"])
    names = [("Mixto Poisson+Elo", "mixto"), ("Solo Poisson", "poisson"), ("Solo Elo", "elo"), ("Base histórica constante", "base")]
    for i, (nm, k) in enumerate(names):
        wsB.cell(2 + i, 1, nm)
        if bt:
            s = bt[k]
            wsB.cell(2 + i, 2, s["acc"]).number_format = "0.0%"
            wsB.cell(2 + i, 3, s["brier"]).number_format = "0.000"
            wsB.cell(2 + i, 4, s["ll"]).number_format = "0.000"
            wsB.cell(2 + i, 5, s["n"])
    widths(wsB, [28, 12, 10, 24, 10, 16, 12, 10])

    wsX = wb.create_sheet("Bajas")
    hdr(wsX, 1, ["Jugador", "Estado (Disponible/Duda/Baja/Sanción)", "Motivo", "Vuelta estimada", "Importancia (0-1)", "Fuente"])
    for i, b in enumerate(bajas):
        for j, v in enumerate(b[:6], 1):
            c = wsX.cell(i + 2, j, v)
            if j in (2, 5):
                c.font = Font(name=ARIAL, color="0000FF")
                c.fill = INPUT
    widths(wsX, [24, 22, 52, 16, 16, 40])

    wsJ = wb.create_sheet("Jugadores")
    hdr(wsJ, 1, ["Jugador", "Minutos", "Partidos", "Valoración media", "Goles", "Asistencias", "G+A por 90'"])
    if jugadores:
        for i, j in enumerate(sorted(jugadores, key=lambda x: -x[1])):
            r = i + 2
            for k, v in enumerate(j, 1):
                wsJ.cell(r, k, v)
            wsJ.cell(r, 7, f"=IF(B{r}>0,(E{r}+F{r})/B{r}*90,0)").number_format = "0.00"
    widths(wsJ, [26, 10, 10, 16, 8, 12, 12])

    wsR = wb.create_sheet("Parametros")
    wsR.cell(1, 1, "Parámetro").font = Font(name=ARIAL, bold=True)
    wsR.cell(1, 2, "Valor").font = Font(name=ARIAL, bold=True)
    wsR.cell(1, 3, "Notas").font = Font(name=ARIAL, bold=True)
    params = [
        ("Fuerza del prior histórico (partidos equivalentes)", 10, "Solo para la estimación simple del Dashboard", True),
        ("Ventaja de campo Elo", HFA, "Fijo en update.py", False),
        ("K de Elo", ELO_K, "Fijo en update.py", False),
        ("Semivida de los partidos (días)", HALF_LIFE, "Fijo en update.py", False),
        ("Sensibilidad a bajas (K)", BAJAS_K, "SIN CALIBRAR", False),
    ]
    for i, (a, b, c, inp) in enumerate(params):
        wsR.cell(2 + i, 1, a)
        cc = wsR.cell(2 + i, 2, b)
        if inp:
            cc.font = Font(name=ARIAL, color="0000FF")
            cc.fill = INPUT
        wsR.cell(2 + i, 3, c)
    widths(wsR, [50, 12, 60])

    ws = wb.create_sheet("Dashboard", 0)
    ws["A1"] = "Predictor Real Madrid - LaLiga"
    ws["A1"].font = Font(name=ARIAL, size=16, bold=True, color="1F3864")
    ws["A2"] = meta["estado"]
    ws["A2"].font = Font(name=ARIAL, italic=True, size=9)
    rows = [
        (4, "TEMPORADA ACTUAL", None, None),
        (5, "Partidos jugados", "=COUNTA(Temporada_actual!G2:G39)", "0"),
        (6, "Victorias", '=COUNTIF(Temporada_actual!G2:G39,"V")', "0"),
        (7, "Empates", '=COUNTIF(Temporada_actual!G2:G39,"E")', "0"),
        (8, "Derrotas", '=COUNTIF(Temporada_actual!G2:G39,"D")', "0"),
        (9, "Puntos", "=3*B6+B7", "0"),
        (10, "% victorias", "=IF(B5>0,B6/B5,0)", "0.0%"),
    ]
    for r, lab, f, fmt in rows:
        a = ws.cell(r, 1, lab)
        if f is None:
            a.font = Font(name=ARIAL, bold=True, color="FFFFFF")
            a.fill = HEAD
            ws.cell(r, 2).fill = HEAD
        else:
            c = ws.cell(r, 2, f)
            c.alignment = Alignment(horizontal="right")
            if fmt:
                c.number_format = fmt
    widths(ws, [52, 30])
    wb.save(path)

import html as _html
import json as _json

def _fmt_date(d):
    if d is None or d == "" or pd.isna(d):
        return "Pendiente de fecha"
    if hasattr(d, "strftime"):
        return d.strftime("%d/%m/%Y")
    try:
        return dt.date.fromisoformat(str(d)[:10]).strftime("%d/%m/%Y")
    except ValueError:
        return str(d)

def _pct(x):
    return f"{100 * x:.0f} %"

def write_site(site_dir, ctx):
    e = _html.escape
    hist, temporada, pred, bt, bajas, meta = (ctx[k] for k in ("hist", "temporada", "pred", "bt", "bajas", "meta"))
    res = [r[6] for r in temporada if r[6]]
    pj, v, em, d = len(res), res.count("V"), res.count("E"), res.count("D")
    pts = 3 * v + em
    rem = 38 - pj
    pesos = [0.025] * 4 + [0.10] * 2 + [0.15] * 2 + [0.20] * 2
    if len(hist) != 10:
        pesos = [1 / len(hist)] * len(hist)
    prior = sum(w * h[1] / 38 for w, h in zip(pesos, hist))
    p = (prior * 10 + v) / (10 + pj)
    if pred:
        fut = sum(r["p"][0] for r in pred["rows"])
        p10, p50, p90 = pred["resumen"][:3]
        metodo = "Modelo Poisson + Elo con ajuste por bajas"
    else:
        fut = rem * p
        sd = math.sqrt(rem * p * (1 - p))
        p10, p50, p90 = round(v + fut - 1.28 * sd), round(v + fut), round(v + fut + 1.28 * sd)
        metodo = "Estimación simple"
    total = v + fut

    nx = pred.get("proximo") if pred else None
    if nx:
        nxt = (nx["rival"], "local" if nx["home"] else "visitante", _fmt_date(nx["fecha"]), nx["p"])
    else:
        nxt = (meta["next_rival"], "visitante" if meta["next_cond"] == "V" else "local" if meta["next_cond"] else "—", "—", None)

    def prob_bar(pp):
        if not pp:
            return "<p class='muted'>Probabilidades disponibles tras la primera actualización con datos.</p>"
        return ("<div class='bar'>"
                f"<span class='w' style='width:{100*pp[0]:.1f}%'>{_pct(pp[0])}</span>"
                f"<span class='d' style='width:{100*pp[1]:.1f}%'>{_pct(pp[1])}</span>"
                f"<span class='l' style='width:{100*pp[2]:.1f}%'>{_pct(pp[2])}</span></div>"
                "<p class='muted'>Victoria · Empate · Derrota</p>")

    played = "".join(
        f"<tr><td>{r[0]}</td><td>{_fmt_date(r[1])}</td><td>{e(str(r[2]))}</td><td>{e(str(r[3] or '—'))}</td>"
        f"<td>{'—' if r[4] is None else f'{r[4]}-{r[5]}'}</td><td class='r{r[6]}'>{r[6]}</td></tr>" for r in temporada if r[6])
    future = ""
    if pred:
        future = "".join(
            f"<tr><td>{e(r['rival'])}</td><td>{'L' if r['home'] else 'V'}</td><td>{_fmt_date(r['fecha'])}</td>"
            f"<td>{_pct(r['p'][0])}</td><td>{_pct(r['p'][1])}</td><td>{_pct(r['p'][2])}</td></tr>" for r in pred["rows"])
    bj = "".join(
        f"<tr><td>{e(str(b[0]))}</td><td>{e(str(b[1]))}</td><td>{e(str(b[2]))}</td><td>{e(str(b[3]))}</td><td>{e(str(b[4]))}</td></tr>"
        for b in bajas if str(b[1]).strip().lower() in PESO_ESTADO)

    page = f"""<!doctype html>
<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Predictor Real Madrid</title>
<style>
:root{{--bg:#f6f7fb;--card:#fff;--text:#1b2133;--muted:#667;--line:#e3e6ef;--accent:#1f3864;--w:#2e8b57;--d:#d4a017;--l:#c0392b;--bar:#3a5ba0}}
@media (prefers-color-scheme:dark){{:root{{--bg:#12151f;--card:#1b2030;--text:#e8ebf5;--muted:#9aa3bd;--line:#2a3148;--accent:#8fb0ff;--bar:#6f93e0}}}}
*{{box-sizing:border-box}}body{{margin:0;font-family:system-ui,-apple-system,Segoe UI,Arial,sans-serif;background:var(--bg);color:var(--text);line-height:1.45}}
main{{max-width:980px;margin:0 auto;padding:20px 16px 60px}}h1{{font-size:1.6rem;margin:.2rem 0;color:var(--accent)}}h2{{font-size:1.1rem;margin:0 0 .6rem;color:var(--accent)}}
.muted{{color:var(--muted);font-size:.85rem}}.grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin:14px 0}}
.card{{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px;margin-bottom:14px}}.grid .card{{margin:0}}
.kpi{{font-size:1.8rem;font-weight:700}}table{{width:100%;border-collapse:collapse;font-size:.9rem}}th,td{{text-align:left;padding:6px 8px;border-bottom:1px solid var(--line)}}th{{color:var(--muted);font-weight:600}}
.wrap{{overflow-x:auto}}.bar{{display:flex;height:30px;border-radius:8px;overflow:hidden;font-size:.8rem;color:#fff;font-weight:600}}.bar span{{display:flex;align-items:center;justify-content:center;min-width:0}}
.bar .w{{background:var(--w)}}.bar .d{{background:var(--d)}}.bar .l{{background:var(--l)}}td.rV{{color:var(--w);font-weight:700}}td.rD{{color:var(--l);font-weight:700}}td.rE{{color:var(--d);font-weight:700}}
</style></head><body><main>
<h1>Predictor Real Madrid</h1>
<p class="muted">{e(meta['estado'])}</p>
<div class="grid">
<div class="card"><div class="muted">Partidos jugados</div><div class="kpi">{pj}</div><div class="muted">{v}V · {em}E · {d}D · {pts} pts</div></div>
<div class="card"><div class="muted">Victorias finales previstas</div><div class="kpi">{total:.1f}</div><div class="muted">Rango P10–P90: {p10} – {p90} (central {p50})</div></div>
<div class="card"><div class="muted">Victorias restantes esperadas</div><div class="kpi">{fut:.1f}</div><div class="muted">de {rem} partidos</div></div>
</div>
<div class="card"><h2>Próximo partido: {e(nxt[0])} ({nxt[1]}) · {nxt[2]}</h2>{prob_bar(nxt[3])}</div>
<div class="card"><h2>Próximos encuentros en calendario</h2>{'<div class="wrap"><table><tr><th>Rival</th><th>L/V</th><th>Fecha</th><th>Victoria</th><th>Empate</th><th>Derrota</th></tr>'+future+'</table></div>' if future else "<p class='muted'>Se rellena tras la primera actualización con datos.</p>"}</div>
<div class="card"><h2>Bajas y dudas</h2>{'<div class="wrap"><table><tr><th>Jugador</th><th>Estado</th><th>Motivo</th><th>Vuelta</th><th>Importancia</th></tr>'+bj+'</table></div>' if bj else "<p class='muted'>Sin bajas registradas en bajas.csv.</p>"}</div>
</main></body></html>"""
    (site_dir / "index.html").write_text(page, encoding="utf-8")
    data = {
        "actualizado": meta["estado"], "jugados": pj, "V": v, "E": em, "D": d, "puntos": pts,
        "victorias_finales_previstas": round(total, 2), "p10": int(p10), "p50": int(p50), "p90": int(p90),
        "metodo": metodo,
        "proximo": {"rival": nxt[0], "condicion": nxt[1], "fecha": nxt[2],
                    "p": [round(float(x), 4) for x in nxt[3]] if nxt[3] else None},
        "restantes": [{"rival": r["rival"], "local": bool(r["home"]), "p": [round(float(x), 4) for x in r["p"]]}
                      for r in (pred["rows"] if pred else [])],
    }
    (site_dir / "data.json").write_text(_json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

LEEME = ["PREDICTOR REAL MADRID"]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-download", action="store_true")
    ap.add_argument("--today", default=None)
    ap.add_argument("--out", default=str(OUT))
    ap.add_argument("--site", action="store_true")
    a = ap.parse_args()
    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    out = Path(a.out)
    if a.site:
        out = BASE / "site" / "predictor_real_madrid.xlsx"
        out.parent.mkdir(exist_ok=True)
    refresh = not a.no_download

    df, cur = load_all(today, refresh)
    fixtures = load_fixtures(refresh)
    bajas = read_bajas(out)
    bajas, jugadores = api_football(today, bajas) if refresh else (bajas, [])

    hist, temporada, pred, bt = SEED_HIST, [list(x) + [None] for x in SEED_2627], None, None
    meta = {"leeme": LEEME, "next_rival": "Valencia", "next_cond": "V", "next_date": "DATE(2026,11,7)",
            "estado": f"Actualizado el {today:%d/%m/%Y}."}

    if df is not None:
        df, ratings = run_elo(df)
        mm = madrid_matches(df)
        h = []
        for s in range(cur - N_HIST, cur):
            x = mm[mm.Season == s]
            if len(x) >= 38:
                h.append((f"{s}-{(s + 1) % 100:02d}", int((x.Res == "V").sum()), int((x.Res == "E").sum()), int((x.Res == "D").sum())))
        if len(h) == N_HIST:
            hist = h
        cm = mm[mm.Season == cur].sort_values("Date")
        if len(cm):
            temporada = []
            for i, r in enumerate(cm.itertuples(), 1):
                temporada.append([i, r.Date.to_pydatetime(), nombre(r.Rival), "L" if r.Home else "V",
                                  int(r.GF), int(r.GC), r.Res, None])
        bt = backtest(df, cur)
        w = bt["w"] if bt else 0.5
        L = indice_bajas(bajas)
        rows = predecir(df, cur, today, ratings, w, L, fixtures)
        if rows:
            ps = [r["p"][0] for r in rows]
            dist = distribucion(ps)
            wins_now = int((cm.Res == "V").sum()) if len(cm) else 0
            cdf = np.cumsum(dist)
            pct = lambda q: int(np.searchsorted(cdf, q)) + wins_now
            pts_now = 3 * wins_now + int((cm.Res == "E").sum()) if len(cm) else 0
            pts = pts_now + sum(3 * r["p"][0] + r["p"][1] for r in rows)
            pred = {
                "rows": rows,
                "resumen": [pct(0.10), pct(0.50), pct(0.90), round(float(pts), 1), round(L, 2), w,
                            round(float(ratings.get(MADRID, 1500)), 0), dt.datetime.now().strftime("%d/%m/%Y %H:%M")],
                "proximo": rows[0] if len(rows) > 0 else None,
            }
            meta["estado"] = f"Actualizado {dt.datetime.now():%d/%m/%Y %H:%M}."

    write_workbook(out, hist, temporada, bajas, jugadores, pred, bt, meta)
    if a.site:
        write_site(out.parent, {"hist": hist, "temporada": temporada, "pred": pred, "bt": bt, "bajas": bajas, "meta": meta})
    print("Listo:", out)

if __name__ == "__main__":
    sys.exit(main())
