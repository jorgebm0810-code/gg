#!/usr/bin/env python3
"""Predictor Real Madrid (LaLiga y Champions) - v4 con Escudo, Logos, Resumen y Probabilidad de Liga."""

import argparse
import io
import math
import os
import sys
import datetime as dt
from pathlib import Path

import numpy as np
import pandas as pd
from scipy.stats import norm

try:
    import requests
except ImportError:
    requests = None

from openpyxl import Workbook, load_workbook
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
    "Oviedo": "Real Oviedo", "Mallorca": "Mallorca", "Valladolid": "Real Valladolid", "Girona": "Girona FC",
    "Osasuna": "Osasuna", "Getafe": "Getafe", "Villarreal": "Villarreal", "Sevilla": "Sevilla FC",
    "Las Palmas": "UD Las Palmas", "Valencia": "Valencia CF", "Elche": "Elche CF"
}

TEAM_LOGOS = {
    "Real Madrid": "https://upload.wikimedia.org/wikipedia/commons/5/56/Real_Madrid_CF.svg",
    "FC Barcelona": "https://upload.wikimedia.org/wikipedia/commons/4/47/FC_Barcelona_%28crest%29.svg",
    "Atlético de Madrid": "https://upload.wikimedia.org/wikipedia/commons/c/c1/Atletico_Madrid_2024_logo.svg",
    "Athletic Club": "https://upload.wikimedia.org/wikipedia/commons/9/98/Athletic_Bilbao_logo.svg",
    "Real Sociedad": "https://upload.wikimedia.org/wikipedia/commons/f/f1/Real_Sociedad_logo.svg",
    "Real Betis": "https://upload.wikimedia.org/wikipedia/commons/1/13/Real_Betis_Balompi%C3%A9.svg",
    "Villarreal": "https://upload.wikimedia.org/wikipedia/commons/a/a2/Villarreal_CF_logo.svg",
    "Sevilla FC": "https://upload.wikimedia.org/wikipedia/commons/b/b4/Sevilla_FC_logo.svg",
    "Valencia CF": "https://upload.wikimedia.org/wikipedia/commons/c/ce/Cf.valencia.svg",
    "Girona FC": "https://upload.wikimedia.org/wikipedia/commons/e/e2/Girona_FC_Logo.svg",
    "Osasuna": "https://upload.wikimedia.org/wikipedia/commons/7/7b/CA_Osasuna_logo.svg",
    "Celta de Vigo": "https://upload.wikimedia.org/wikipedia/commons/a/a1/RC_Celta_de_Vigo_logo.svg",
    "Mallorca": "https://upload.wikimedia.org/wikipedia/commons/1/1a/RCD_Mallorca_logo.svg",
    "Rayo Vallecano": "https://upload.wikimedia.org/wikipedia/commons/7/71/Rayo_Vallecano_logo.svg",
    "Getafe": "https://upload.wikimedia.org/wikipedia/commons/3/3a/Getafe_CF_logo.svg",
    "Alavés": "https://upload.wikimedia.org/wikipedia/commons/a/a9/Deportivo_Alav%C3%A9s_logo.svg",
    "Leganés": "https://upload.wikimedia.org/wikipedia/commons/0/05/CD_Legan%C3%A9s_logo.svg",
    "Espanyol": "https://upload.wikimedia.org/wikipedia/commons/f/f6/RCD_Espanyol_logo.svg",
    "Real Valladolid": "https://upload.wikimedia.org/wikipedia/commons/f/f9/Real_Valladolid_logo.svg",
    "UD Las Palmas": "https://upload.wikimedia.org/wikipedia/commons/8/87/UD_Las_Palmas_logo.svg",
    "Cádiz": "https://upload.wikimedia.org/wikipedia/commons/4/44/C%C3%A1diz_CF_logo.svg",
    "Almería": "https://upload.wikimedia.org/wikipedia/commons/8/8b/UD_Almer%C3%ADa_logo.svg",
    "Elche CF": "https://upload.wikimedia.org/wikipedia/commons/e/e0/Elche_CF_logo.svg",
    "Málaga": "https://upload.wikimedia.org/wikipedia/commons/7/7e/M%C3%A1lage_CF_logo.svg",
    "Racing de Santander": "https://upload.wikimedia.org/wikipedia/commons/4/4b/Real_Racing_Club_de_Santander_logo.svg",
    "Deportivo La Coruña": "https://upload.wikimedia.org/wikipedia/commons/0/09/Real_Club_Deportivo_La_Coru%C3%B1a.svg",
    "Sporting de Gijón": "https://upload.wikimedia.org/wikipedia/commons/1/18/Real_Sporting_de_Gij%C3%B3n_logo.svg",
    "Real Oviedo": "https://upload.wikimedia.org/wikipedia/commons/b/bc/Real_Oviedo_logo.svg"
}

def get_logo(team_name):
    url = TEAM_LOGOS.get(team_name, "https://upload.wikimedia.org/wikipedia/commons/5/56/Real_Madrid_CF.svg")
    return f"<img src='{url}' alt='{team_name}' class='team-logo' onerror=\"this.style.display='none'\">"

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
    (9, None, "Sevilla FC", "L", 4, 1, "V"),
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

def poisson_matrix(mu_h, mu_a, max_goals=6):
    k = np.arange(max_goals + 1)
    fact = np.array([math.factorial(int(i)) for i in k], dtype=float)
    ph = np.exp(-mu_h) * (mu_h ** k) / fact
    pa = np.exp(-mu_a) * (mu_a ** k) / fact
    m = np.outer(ph, pa)
    m /= m.sum()
    scores = []
    for h in range(max_goals + 1):
        for a in range(max_goals + 1):
            scores.append((h, a, m[h, a]))
    scores.sort(key=lambda x: x[2], reverse=True)
    return scores[:3]

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
    top_scores = poisson_matrix(mu_h, mu_a)
    return np.tril(m, -1).sum(), np.trace(m), np.triu(m, 1).sum(), mu_h, mu_a, top_scores

def madrid_matches(df):
    m = df[(df.HomeTeam == MADRID) | (df.AwayTeam == MADRID)].copy()
    m["Home"] = m.HomeTeam == MADRID
    m["Rival"] = np.where(m.Home, m.AwayTeam, m.HomeTeam)
    m["GF"] = np.where(m.Home, m.FTHG, m.FTAG)
    m["GC"] = np.where(m.Home, m.FTAG, m.FTHG)
    m["Res"] = np.where(m.GF > m.GC, "V", np.where(m.GF == m.GC, "E", "D"))
    return m

def get_standings(df, cur):
    cd = df[df.Season == cur].copy()
    teams = sorted(set(cd.HomeTeam) | set(cd.AwayTeam))
    stats = {nombre(t): {"PJ": 0, "V": 0, "E": 0, "D": 0, "GF": 0, "GC": 0, "DG": 0, "Pts": 0} for t in teams}
    
    for r in cd.itertuples():
        h = nombre(r.HomeTeam)
        a = nombre(r.AwayTeam)
        gh, ga = int(r.FTHG), int(r.FTAG)
        
        stats[h]["PJ"] += 1
        stats[a]["PJ"] += 1
        stats[h]["GF"] += gh
        stats[h]["GC"] += ga
        stats[a]["GF"] += ga
        stats[a]["GC"] += gh
        
        if gh > ga:
            stats[h]["V"] += 1
            stats[h]["Pts"] += 3
            stats[a]["D"] += 1
        elif gh < ga:
            stats[a]["V"] += 1
            stats[a]["Pts"] += 3
            stats[h]["D"] += 1
        else:
            stats[h]["E"] += 1
            stats[a]["E"] += 1
            stats[h]["Pts"] += 1
            stats[a]["Pts"] += 1

    for t in stats:
        stats[t]["DG"] = stats[t]["GF"] - stats[t]["GC"]

    tabla = []
    for team, data in stats.items():
        data["Equipo"] = team
        tabla.append(data)

    tabla.sort(key=lambda x: (x["Pts"], x["DG"], x["GF"]), reverse=True)
    for pos, item in enumerate(tabla, 1):
        item["Pos"] = pos
    return tabla

def to_madrid(p3, home):
    return (p3[0], p3[1], p3[2]) if home else (p3[2], p3[1], p3[0])

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
        pp_res = poisson_probs(model, h, a, adj_att, adj_def)
        pp = pp_res[:3]
        top_scores = pp_res[5]
        
        re, rr = ratings.get(MADRID, 1500.0), ratings.get(t_orig, ELO_NEW)
        pe = elo_probs((re + HFA - rr) if home else (rr + HFA - re))
        pm = to_madrid(pp, home)
        pel = to_madrid(pe, home)
        p = [w * pm[i] + (1 - w) * pel[i] for i in range(3)]
        gm, gr = (pp_res[3], pp_res[4]) if home else (pp_res[4], pp_res[3])

        rows.append({
            "rival": t_name,
            "home": home,
            "fecha": match["fecha"],
            "p": p,
            "gm": gm,
            "gr": gr,
            "elo": rr,
            "top_scores": top_scores
        })

    return rows

def distribucion(pw):
    d = np.array([1.0])
    for p in pw:
        d = np.convolve(d, [1 - p, p])
    return d

ARIAL = "Arial"
HEAD = PatternFill("solid", fgColor="1F3864")

def hdr(ws, row, labels, col=1):
    for i, t in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=t)
        c.font = Font(name=ARIAL, bold=True, color="FFFFFF")
        c.fill = HEAD
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

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
            wsH.cell(r, j, val)
        wsH.cell(r, 6, f"=3*C{r}+D{r}")
        wsH.cell(r, 7, f"=C{r}/B{r}").number_format = "0.0%"
        wsH.cell(r, 8, pesos[i] if len(hist) == 10 else round(1 / len(hist), 3)).number_format = "0.0%"
    widths(wsH, [16, 8, 11, 10, 10, 9, 12, 16])
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
    hist, temporada, pred, bt, bajas, meta, standings = (ctx[k] for k in ("hist", "temporada", "pred", "bt", "bajas", "meta", "standings"))
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
        fut_wins = sum(r["p"][0] for r in pred["rows"])
        fut_pts = sum(3 * r["p"][0] + r["p"][1] for r in pred["rows"])
        p10, p50, p90 = pred["resumen"][:3]
        total_wins = v + fut_wins
        total_pts_esperados = pts + fut_pts
        prob_liga = float(np.clip((total_pts_esperados - 75) / (92 - 75) * 100, 2.0, 98.0))
        # O usar distribución normal con scipy
        incertidumbre = np.sqrt(max(1, rem)) * 1.6
        prob_liga = float(np.clip((1.0 - norm.cdf(85.0, loc=total_pts_esperados, scale=incertidumbre)) * 100, 2.0, 98.0))
        metodo = "Modelo Poisson + Elo"
    else:
        fut_wins = rem * p
        sd = math.sqrt(rem * p * (1 - p))
        p10, p50, p90 = round(v + fut_wins - 1.28 * sd), round(v + fut_wins), round(v + fut_wins + 1.28 * sd)
        total_wins = v + fut_wins
        total_pts_esperados = pts + (fut_wins * 3)
        prob_liga = 35.0
        metodo = "Estimación simple"
    
    # Buscar posición actual del Madrid en standings
    madrid_pos = "-"
    madrid_dg = 0
    for row in standings:
        if row["Equipo"] == MADRID:
            madrid_pos = row["Pos"]
            madrid_dg = row["DG"]
            break

    nx = pred.get("proximo") if pred else None
    if nx:
        nxt = (nx["rival"], "local" if nx["home"] else "visitante", _fmt_date(nx["fecha"]), nx["p"], nx.get("top_scores", []))
    else:
        nxt = (meta["next_rival"], "visitante" if meta["next_cond"] == "V" else "local" if meta["next_cond"] else "—", "—", None, [])

    def prob_bar(pp):
        if not pp:
            return "<p class='muted'>Probabilidades disponibles tras la primera actualización con datos.</p>"
        return ("<div class='bar'>"
                f"<span class='w' style='width:{100*pp[0]:.1f}%'>{_pct(pp[0])}</span>"
                f"<span class='d' style='width:{100*pp[1]:.1f}%'>{_pct(pp[1])}</span>"
                f"<span class='l' style='width:{100*pp[2]:.1f}%'>{_pct(pp[2])}</span></div>"
                "<p class='muted'>Victoria · Empate · Derrota</p>")

    def format_top_scores(scores):
        if not scores:
            return ""
        items = [f"<b>{h}-{a}</b> ({p*100:.1f}%)" for h, a, p in scores]
        return "<div class='top-scores'>Marcadores más probables: " + " · ".join(items) + "</div>"

    future = ""
    if pred:
        for r in pred["rows"]:
            scores_str = ", ".join([f"{h}-{a}" for h, a, _ in r.get("top_scores", [])[:2]])
            logo_rival = get_logo(r['rival'])
            future += (f"<tr><td><div class='team-cell'>{logo_rival} <span>{e(r['rival'])}</span></div></td>"
                       f"<td>{'L' if r['home'] else 'V'}</td><td>{_fmt_date(r['fecha'])}</td>"
                       f"<td>{_pct(r['p'][0])}</td><td>{_pct(r['p'][1])}</td><td>{_pct(r['p'][2])}</td>"
                       f"<td><span class='score-tag'>{scores_str}</span></td></tr>")

    bj = "".join(
        f"<tr><td>{e(str(b[0]))}</td><td>{e(str(b[1]))}</td><td>{e(str(b[2]))}</td><td>{e(str(b[3]))}</td><td>{e(str(b[4]))}</td></tr>"
        for b in bajas if str(b[1]).strip().lower() in PESO_ESTADO)

    std_rows = ""
    if standings:
        for row in standings:
            highlight = "class='madrid-row'" if row["Equipo"] == MADRID else ""
            logo_eq = get_logo(row["Equipo"])
            std_rows += f"<tr {highlight}><td><b>{row['Pos']}</b></td><td><div class='team-cell'>{logo_eq} <span>{e(row['Equipo'])}</span></div></td><td>{row['PJ']}</td><td>{row['V']}</td><td>{row['E']}</td><td>{row['D']}</td><td>{row['GF']}</td><td>{row['GC']}</td><td>{row['DG']}</td><td><b>{row['Pts']}</b></td></tr>"

    logo_rm = TEAM_LOGOS[MADRID]
    logo_nx = get_logo(nxt[0])

    page = f"""<!doctype html>
<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Predictor Real Madrid</title>
<style>
:root{{--bg:#f6f7fb;--card:#fff;--text:#1b2133;--muted:#667;--line:#e3e6ef;--accent:#1f3864;--w:#2e8b57;--d:#d4a017;--l:#c0392b;--bar:#3a5ba0}}
@media (prefers-color-scheme:dark){{:root{{--bg:#12151f;--card:#1b2030;--text:#e8ebf5;--muted:#9aa3bd;--line:#2a3148;--accent:#8fb0ff;--bar:#6f93e0}}}}
*{{box-sizing:border-box}}body{{margin:0;font-family:system-ui,-apple-system,Segoe UI,Arial,sans-serif;background:var(--bg);color:var(--text);line-height:1.45}}
main{{max-width:980px;margin:0 auto;padding:20px 16px 60px}}
.header-container{{display:flex;align-items:center;gap:16px;margin-bottom:16px;background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px 20px}}
.header-shield{{width:60px;height:60px;object-fit:contain}}
h1{{font-size:1.5rem;margin:0;color:var(--accent)}}h2{{font-size:1.1rem;margin:0 0 .6rem;color:var(--accent)}}
.muted{{color:var(--muted);font-size:.85rem}}.grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin:14px 0}}
.card{{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px;margin-bottom:14px}}.grid .card{{margin:0}}
.kpi{{font-size:1.8rem;font-weight:700}}table{{width:100%;border-collapse:collapse;font-size:.9rem}}th,td{{text-align:left;padding:6px 8px;border-bottom:1px solid var(--line)}}th{{color:var(--muted);font-weight:600}}
.wrap{{overflow-x:auto}}.bar{{display:flex;height:30px;border-radius:8px;overflow:hidden;font-size:.8rem;color:#fff;font-weight:600}}.bar span{{display:flex;align-items:center;justify-content:center;min-width:0}}
.bar .w{{background:var(--w)}}.bar .d{{background:var(--d)}}.bar .l{{background:var(--l)}}
.team-cell{{display:flex;align-items:center;gap:8px}}
.team-logo{{width:22px;height:22px;object-fit:contain}}
.top-scores{{margin-top:10px;font-size:.9rem;background:var(--bg);padding:8px 12px;border-radius:6px;border:1px solid var(--line)}}
.score-tag{{background:var(--bg);padding:2px 6px;border-radius:4px;font-weight:600;font-size:.8rem}}
.madrid-row{{background:rgba(31,56,100,0.12);font-weight:bold}}
.badge-prob{{display:inline-block;background:var(--accent);color:#fff;padding:4px 10px;border-radius:6px;font-weight:700;font-size:1.1rem}}
</style></head><body><main>

<div class="header-container">
<img src="{logo_rm}" alt="Real Madrid" class="header-shield">
<div>
<h1>Predictor Real Madrid &middot; La Liga 2026/27</h1>
<p class="muted">{e(meta['estado'])}</p>
</div>
</div>

<div class="card" style="border-left:5px solid var(--accent);">
<h2>Resumen de Temporada y Probabilidad de Liga</h2>
<p>El Real Madrid marcha actualmente en la <b>posición #{madrid_pos}</b> de La Liga tras disputar {pj} jornadas, acumulando <b>{pts} puntos</b> ({v}V, {em}E, {d}D) y una diferencia de goles de {madrid_dg:+.1f}.</p>
<p>La probabilidad en vivo estimada de que el equipo conquiste el título de La Liga esta temporada es del <span class="badge-prob">{prob_liga:.1f}%</span>.</p>
</div>

<div class="grid">
<div class="card"><div class="muted">Partidos jugados</div><div class="kpi">{pj}</div><div class="muted">{v}V · {em}E · {d}D · {pts} pts</div></div>
<div class="card"><div class="muted">Victorias finales previstas</div><div class="kpi">{total_wins:.1f}</div><div class="muted">Rango P10–P90: {p10} – {p90} (central {p50})</div></div>
<div class="card"><div class="muted">Victorias restantes esperadas</div><div class="kpi">{fut_wins:.1f}</div><div class="muted">de {rem} partidos</div></div>
</div>

<div class="card">
<h2>Próximo partido: <div class="team-cell" style="display:inline-flex;vertical-align:middle;">{logo_nx} <span>{e(nxt[0])}</span></div> ({nxt[1]}) · {nxt[2]}</h2>
{prob_bar(nxt[3])}
{format_top_scores(nxt[4])}
</div>

<div class="card">
<h2>Clasificación Actual de La Liga</h2>
{'<div class="wrap"><table><tr><th>Pos</th><th>Equipo</th><th>PJ</th><th>V</th><th>E</th><th>D</th><th>GF</th><th>GC</th><th>DG</th><th>Pts</th></tr>'+std_rows+'</table></div>' if std_rows else "<p class='muted'>Clasificación disponible tras cargar los datos.</p>"}
</div>

<div class="card">
<h2>Próximos encuentros y marcadores probables</h2>
{'<div class="wrap"><table><tr><th>Rival</th><th>L/V</th><th>Fecha</th><th>Victoria</th><th>Empate</th><th>Derrota</th><th>Marcadores probables</th></tr>'+future+'</table></div>' if future else "<p class='muted'>Se rellena tras la primera actualización con datos.</p>"}
</div>

<div class="card">
<h2>Bajas y dudas</h2>
{'<div class="wrap"><table><tr><th>Jugador</th><th>Estado</th><th>Motivo</th><th>Vuelta</th><th>Importancia</th></tr>'+bj+'</table></div>' if bj else "<p class='muted'>Sin bajas registradas en bajas.csv.</p>"}
</div>
</main></body></html>"""
    (site_dir / "index.html").write_text(page, encoding="utf-8")
    data = {
        "actualizado": meta["estado"], "jugados": pj, "V": v, "E": em, "D": d, "puntos": pts,
        "probabilidad_liga": round(prob_liga, 1),
        "victorias_finales_previstas": round(total_wins, 2), "p10": int(p10), "p50": int(p50), "p90": int(p90),
        "proximo": {"rival": nxt[0], "condicion": nxt[1], "fecha": nxt[2],
                    "p": [round(float(x), 4) for x in nxt[3]] if nxt[3] else None,
                    "top_scores": nxt[4]}
    }
    (site_dir / "data.json").write_text(_json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

LEEME = ["PREDICTOR REAL MADRID v4"]

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
    standings = []
    hist, temporada, pred, bt = SEED_HIST, [list(x) + [None] for x in SEED_2627], None, None
    meta = {"leeme": LEEME, "next_rival": "Valencia", "next_cond": "V", "next_date": "DATE(2026,11,7)",
            "estado": f"Actualizado el {today:%d/%m/%Y}."}

    if df is not None:
        df, ratings = run_elo(df)
        standings = get_standings(df, cur)
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
        bt = backtest(df, cur) if len(mm) > 10 else None
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

    write_workbook(out, hist, temporada, bajas, [], pred, bt, meta)
    if a.site:
        write_site(out.parent, {"hist": hist, "temporada": temporada, "pred": pred, "bt": bt, "bajas": bajas, "meta": meta, "standings": standings})
    print("Listo:", out)

if __name__ == "__main__":
    sys.exit(main())
