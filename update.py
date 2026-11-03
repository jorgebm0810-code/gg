#!/usr/bin/env python3
"""Predictor Real Madrid (LaLiga) - actualización diaria.

Qué hace:
  1. Descarga resultados de LaLiga (football-data.co.uk, gratis, sin clave) de las últimas 10
     temporadas + la actual, y el calendario próximo (fixtures.csv).
  2. Calcula Elo y fuerzas de ataque/defensa (Poisson con decaimiento temporal).
  3. Backtest walk-forward (cada partido se predice solo con datos anteriores) en las últimas
     4 temporadas completas y elige el peso Poisson/Elo.
  4. Lee las bajas (hoja "Bajas" del Excel; opcionalmente API-Football si hay clave) y ajusta
     las predicciones futuras.
  5. Escribe predictor_real_madrid.xlsx (dashboard con fórmulas).

Uso:  python update.py            (descarga y actualiza)
      python update.py --no-download   (usa solo la caché de ./data)
"""
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
except ImportError:  # sin requests solo funciona con --no-download
    requests = None

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------- configuración
BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
OUT = BASE / "predictor_real_madrid.xlsx"
MADRID = "Real Madrid"
URL_SEASON = "https://www.football-data.co.uk/mmz4281/{code}/SP1.csv"
URL_FIXTURES = "https://www.football-data.co.uk/fixtures.csv"

HFA = 65.0            # ventaja de campo en puntos Elo
ELO_K = 20.0
ELO_NEW = 1450.0      # equipo sin historial (ascendido)
ELO_REG = 0.20        # regresión a 1500 al empezar temporada
DRAW_D = 0.27         # prob. de empate con equipos igualados
DRAW_S = 400.0
HALF_LIFE = 300.0     # días (decaimiento del peso de partidos antiguos)
PRIOR_GAMES = 3.0     # encogimiento hacia la media en ataque/defensa
BAJAS_K = 0.03        # sensibilidad al índice de bajas (NO calibrado: no hay histórico de lesiones)
N_HIST = 10           # temporadas de histórico
N_BACKTEST = 4        # temporadas de prueba

NOMBRES = {
    "Ath Madrid": "Atlético de Madrid", "Ath Bilbao": "Athletic Club", "Sociedad": "Real Sociedad",
    "Betis": "Real Betis", "Vallecano": "Rayo Vallecano", "Celta": "Celta de Vigo",
    "Barcelona": "FC Barcelona", "Espanol": "Espanyol", "Alaves": "Alavés", "Cadiz": "Cádiz",
    "Almeria": "Almería", "Santander": "Racing de Santander", "La Coruna": "Deportivo La Coruña",
    "Leganes": "Leganés", "Malaga": "Málaga", "Sp Gijon": "Sporting de Gijón",
    "Oviedo": "Real Oviedo", "Mallorca": "Mallorca", "Valladolid": "Real Valladolid",
}

# Verificado en prensa (nov-2026). Se usa si no hay datos descargados.
SEED_HIST = [  # temporada, V, E, D
    ("2016-17", 29, 6, 3), ("2017-18", 22, 10, 6), ("2018-19", 21, 5, 12),
    ("2019-20", 26, 9, 3), ("2020-21", 25, 9, 4), ("2021-22", 26, 8, 4),
    ("2022-23", 24, 5, 9), ("2023-24", 29, 8, 1), ("2024-25", 26, 6, 6),
    ("2025-26", 27, 5, 6),
]
SEED_2627 = [  # jornada, fecha, rival, cond, GF, GC, res  (None = dato no verificado)
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
SEED_BAJAS = [  # jugador, estado, motivo, vuelta, importancia, fuente
    ("Raúl Asencio", "Baja", "Fractura por estrés de tibia derecha (operado 30-sep-2026)",
     "varios meses", 0.3, "Prensa 30-sep-2026; importancia estimada por defecto"),
]


def nombre(t):
    return NOMBRES.get(t, t)


# ----------------------------------------------------------------------------- datos
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
        except Exception as e:  # noqa: BLE001
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
        # solo se refresca la temporada en curso; las pasadas se descargan una vez
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
        return fx.dropna(subset=["Date"])
    except Exception:  # noqa: BLE001
        return pd.DataFrame()


# ----------------------------------------------------------------------------- Elo
def elo_probs(dr):
    """Prob. (local, empate, visitante) a partir de la diferencia Elo (con ventaja de campo)."""
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


# ----------------------------------------------------------------------------- Poisson
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
    """(P local, P empate, P visitante, mu_local, mu_visitante). adj_* se aplican al Real Madrid."""
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


# ----------------------------------------------------------------------------- backtest
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
        # frecuencia histórica previa (base ingenua)
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


# ----------------------------------------------------------------------------- bajas / jugadores
def read_bajas(path):
    """Fuente principal: bajas.csv (editable, versionable en Git). Respaldo: hoja Bajas del Excel."""
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
        except Exception as e:  # noqa: BLE001
            print("  aviso: no se pudo leer bajas.csv:", e)
    rows = []
    if path.exists():
        try:
            ws = load_workbook(path)["Bajas"]
            for r in ws.iter_rows(min_row=2, values_only=True):
                if r and r[0]:
                    rows.append(list(r[:6]) + [None] * (6 - len(r[:6])))
        except Exception:  # noqa: BLE001
            pass
    return rows or [list(x) for x in SEED_BAJAS]


def api_football(today, bajas):
    """Opcional: API_FOOTBALL_KEY (plan gratuito: la cobertura de temporadas puede ser limitada)."""
    key = os.environ.get("API_FOOTBALL_KEY")
    if not key or requests is None:
        return bajas, []
    h = {"x-apisports-key": key}
    y = today.year if today.month >= 7 else today.year - 1
    jugadores = []
    try:
        r = requests.get("https://v3.football.api-sports.io/injuries",
                         params={"league": 140, "season": y, "team": 541}, headers=h, timeout=30).json()
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
                             params={"team": 541, "season": y, "league": 140, "page": page},
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
    except Exception as e:  # noqa: BLE001
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


# ----------------------------------------------------------------------------- predicción
def predecir(df, cur, today, ratings, w, L, fixtures):
    now = pd.Timestamp(today) + pd.Timedelta(days=1)
    model = fit_strengths(df, now)
    if model is None:
        return None
    cd = df[df.Season == cur]
    teams = sorted((set(cd.HomeTeam) | set(cd.AwayTeam)) - {MADRID})
    played_h = set(cd[cd.HomeTeam == MADRID].AwayTeam)
    played_a = set(cd[cd.AwayTeam == MADRID].HomeTeam)
    adj_att, adj_def = math.exp(-BAJAS_K * L), math.exp(BAJAS_K * 0.5 * L)
    fx = {}
    if len(fixtures):
        for r in fixtures.itertuples():
            if r.HomeTeam == MADRID:
                fx[(r.AwayTeam, True)] = r.Date
            elif r.AwayTeam == MADRID:
                fx[(r.HomeTeam, False)] = r.Date
    rows = []
    for t in teams:
        for home in (True, False):
            if (home and t in played_h) or (not home and t in played_a):
                continue
            h, a = (MADRID, t) if home else (t, MADRID)
            pp = poisson_probs(model, h, a, adj_att, adj_def)
            re, rr = ratings.get(MADRID, 1500.0), ratings.get(t, ELO_NEW)
            pe = elo_probs((re + HFA - rr) if home else (rr + HFA - re))
            pm = to_madrid(pp[:3], home)
            pel = to_madrid(pe, home)
            p = [w * pm[i] + (1 - w) * pel[i] for i in range(3)]
            gm, gr = (pp[3], pp[4]) if home else (pp[4], pp[3])
            rows.append({"rival": nombre(t), "home": home, "fecha": fx.get((t, home)), "p": p,
                         "gm": gm, "gr": gr, "elo": rr})
    rows.sort(key=lambda r: (r["fecha"] is None, r["fecha"] if r["fecha"] is not None else pd.Timestamp.max, r["rival"]))
    return rows


def distribucion(pw):
    d = np.array([1.0])
    for p in pw:
        d = np.convolve(d, [1 - p, p])
    return d


# ----------------------------------------------------------------------------- Excel
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

    # ---- Historico
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
    wsH.cell(last + 4, 1, "Fuente: clasificaciones finales de LaLiga (prensa/estadísticas oficiales); la de "
                          "2025-26 contrastada con prensa (86 pts = 27V-5E-6D). Azul = dato introducido.").font = Font(name=ARIAL, italic=True, size=9)
    widths(wsH, [16, 8, 11, 10, 10, 9, 12, 16])
    ch = BarChart()
    ch.type = "col"
    ch.title = "Victorias del Real Madrid en LaLiga por temporada"
    ch.add_data(Reference(wsH, min_col=3, min_row=1, max_row=last), titles_from_data=True)
    ch.set_categories(Reference(wsH, min_col=1, min_row=2, max_row=last))
    ch.height, ch.width, ch.legend = 8, 16, None
    wsH.add_chart(ch, "J2")
    HROW = last + 2  # fila de la tasa ponderada

    # ---- Temporada
    wsT = wb.create_sheet("Temporada_actual")
    hdr(wsT, 1, ["Jornada", "Fecha", "Rival", "Condición (L/V)", "GF", "GC", "Resultado (V/E/D)", "Nota"])
    for i, row in enumerate(temporada):
        r = i + 2
        for j, v in enumerate(row[:7], 1):
            if v is not None:
                wsT.cell(r, j, v)
        wsT.cell(r, 8, row[7] if len(row) > 7 else None)
    wsT.cell(41, 1, "L = local, V = visitante. Celdas vacías en GF/GC = marcador no verificado (se rellena solo con update.py).").font = Font(name=ARIAL, italic=True, size=9)
    widths(wsT, [9, 12, 24, 16, 6, 6, 18, 40])

    # ---- Predicciones
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
            if r["fecha"] is not None:
                wsP.cell(rr, 3, r["fecha"].to_pydatetime()).number_format = "dd/mm/yyyy"
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
            wsP.cell(4, 15, nx["fecha"].to_pydatetime()).number_format = "dd/mm/yyyy"
            for j in range(3):
                wsP.cell(5 + j, 15, round(float(nx["p"][j]), 4)).number_format = "0.0%"
    widths(wsP, [24, 10, 12, 11, 11, 11, 16, 15, 10, 3, 26, 18, 3, 18, 22])

    # ---- Backtest
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
    wsB.cell(2, 7, "MAE victorias/temporada (modelo)")
    wsB.cell(3, 7, "MAE victorias/temporada (base)")
    wsB.cell(4, 7, "Peso Poisson óptimo (w)")
    if bt:
        wsB.cell(2, 8, bt["mae_model"]).number_format = "0.00"
        wsB.cell(3, 8, bt["mae_base"]).number_format = "0.00"
        wsB.cell(4, 8, bt["w"]).number_format = "0.0"
    hdr(wsB, 8, ["Temporada", "PJ", "Victorias reales", "Victorias previstas (modelo)", "Error", "Previstas (base)", "Error base"])
    if bt:
        for i, (t, pj, real, esp, base) in enumerate(bt["seasons"]):
            r = 9 + i
            wsB.cell(r, 1, t)
            wsB.cell(r, 2, pj)
            wsB.cell(r, 3, real)
            wsB.cell(r, 4, round(esp, 2))
            wsB.cell(r, 5, f"=D{r}-C{r}").number_format = "0.00"
            wsB.cell(r, 6, round(base, 2))
            wsB.cell(r, 7, f"=F{r}-C{r}").number_format = "0.00"
        r0 = 16
        hdr(wsB, r0, ["Fecha", "Rival", "Condición", "Resultado", "P victoria", "P empate", "P derrota"])
        for i, row in enumerate(bt["rows"]):
            r = r0 + 1 + i
            wsB.cell(r, 1, row["Fecha"].to_pydatetime()).number_format = "dd/mm/yyyy"
            wsB.cell(r, 2, nombre(row["Rival"]))
            wsB.cell(r, 3, "L" if row["Home"] else "V")
            wsB.cell(r, 4, row["Res"])
            for j in range(3):
                wsB.cell(r, 5 + j, round(float(bt["Pm"][i][j]), 4)).number_format = "0.0%"
    else:
        wsB.cell(9, 1, "Pendiente: se rellena al ejecutar update.py (necesita descargar los resultados de LaLiga).")
    widths(wsB, [28, 12, 10, 24, 10, 16, 12, 10])

    # ---- Bajas
    wsX = wb.create_sheet("Bajas")
    hdr(wsX, 1, ["Jugador", "Estado (Disponible/Duda/Baja/Sanción)", "Motivo", "Vuelta estimada", "Importancia (0-1)", "Fuente"])
    for i, b in enumerate(bajas):
        for j, v in enumerate(b[:6], 1):
            c = wsX.cell(i + 2, j, v)
            if j in (2, 5):
                c.font = Font(name=ARIAL, color="0000FF")
                c.fill = INPUT
    wsX.cell(30, 1, "Edita Estado e Importancia (amarillo). Importancia: 1 = indispensable, 0.5 = rotación, 0.1 = marginal. "
                    "Baja/Sanción pesa 1, Duda pesa 0,5. Fuente real: bajas.csv (esta hoja se regenera con cada ejecución).").font = Font(name=ARIAL, italic=True, size=9)
    widths(wsX, [24, 22, 52, 16, 16, 40])

    # ---- Jugadores
    wsJ = wb.create_sheet("Jugadores")
    hdr(wsJ, 1, ["Jugador", "Minutos", "Partidos", "Valoración media", "Goles", "Asistencias", "G+A por 90'"])
    if jugadores:
        for i, j in enumerate(sorted(jugadores, key=lambda x: -x[1])):
            r = i + 2
            for k, v in enumerate(j, 1):
                wsJ.cell(r, k, v)
            wsJ.cell(r, 7, f"=IF(B{r}>0,(E{r}+F{r})/B{r}*90,0)").number_format = "0.00"
    else:
        wsJ.cell(2, 1, "Sin datos: define la variable de entorno API_FOOTBALL_KEY (ver LEEME.md) para rellenar estadísticas individuales.")
    widths(wsJ, [26, 10, 10, 16, 8, 12, 12])

    # ---- Parametros
    wsR = wb.create_sheet("Parametros")
    wsR.cell(1, 1, "Parámetro").font = Font(name=ARIAL, bold=True)
    wsR.cell(1, 2, "Valor").font = Font(name=ARIAL, bold=True)
    wsR.cell(1, 3, "Notas").font = Font(name=ARIAL, bold=True)
    params = [
        ("Fuerza del prior histórico (partidos equivalentes)", 10, "Solo para la estimación simple del Dashboard", True),
        ("Ventaja de campo Elo", HFA, "Fijo en update.py", False),
        ("K de Elo", ELO_K, "Fijo en update.py", False),
        ("Semivida de los partidos (días)", HALF_LIFE, "Fijo en update.py", False),
        ("Sensibilidad a bajas (K)", BAJAS_K, "SIN CALIBRAR: no hay histórico de lesiones gratuito", False),
    ]
    for i, (a, b, c, inp) in enumerate(params):
        wsR.cell(2 + i, 1, a)
        cc = wsR.cell(2 + i, 2, b)
        if inp:
            cc.font = Font(name=ARIAL, color="0000FF")
            cc.fill = INPUT
        wsR.cell(2 + i, 3, c)
    widths(wsR, [50, 12, 60])

    # ---- Dashboard
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
        (12, "PREVISIÓN DE FIN DE TEMPORADA", None, None),
        (13, "Partidos restantes", "=38-B5", "0"),
        (14, "Tasa histórica ponderada de victoria", f"=Historico!G{HROW}", "0.0%"),
        (15, "Tasa estimada (prior + temporada actual)", "=(B14*Parametros!B2+B6)/(Parametros!B2+B5)", "0.0%"),
        (16, "Victorias futuras - estimación simple", "=B13*B15", "0.0"),
        (17, "Victorias futuras - modelo Poisson+Elo", '=IF(COUNT(Predicciones!D2:D40)>0,SUM(Predicciones!D2:D40),"sin ejecutar update.py")', "0.0"),
        (18, "Victorias finales previstas", "=B6+IF(ISNUMBER(B17),B17,B16)", "0.0"),
        (19, "Escenario conservador (P10)", "=IF(ISNUMBER(Predicciones!L1),Predicciones!L1,ROUND(B6+B16-1.28*SQRT(B13*B15*(1-B15)),0))", "0"),
        (20, "Escenario central", "=IF(ISNUMBER(Predicciones!L2),Predicciones!L2,ROUND(B18,0))", "0"),
        (21, "Escenario optimista (P90)", "=IF(ISNUMBER(Predicciones!L3),Predicciones!L3,ROUND(B6+B16+1.28*SQRT(B13*B15*(1-B15)),0))", "0"),
        (22, "Método", '=IF(ISNUMBER(B17),"Modelo Poisson+Elo con ajuste por bajas","Estimación simple (ejecuta update.py)")', None),
        (24, "PRÓXIMO PARTIDO", None, None),
        (25, "Rival", f'=IF(Predicciones!O2<>"",Predicciones!O2,"{meta["next_rival"]}")', None),
        (26, "Condición (L/V)", f'=IF(Predicciones!O3<>"",Predicciones!O3,"{meta["next_cond"]}")', None),
        (27, "Fecha", f'=IF(Predicciones!O4<>"",Predicciones!O4,{meta["next_date"]})', "dd/mm/yyyy"),
        (28, "P victoria", '=IF(ISNUMBER(Predicciones!O5),Predicciones!O5,"sin ejecutar update.py")', "0.0%"),
        (29, "P empate", '=IF(ISNUMBER(Predicciones!O6),Predicciones!O6,"sin ejecutar update.py")', "0.0%"),
        (30, "P derrota", '=IF(ISNUMBER(Predicciones!O7),Predicciones!O7,"sin ejecutar update.py")', "0.0%"),
        (32, "PRECISIÓN DEL MODELO (backtest)", None, None),
        (33, "Precisión del resultado más probable", '=IF(ISNUMBER(Backtest!B2),Backtest!B2,"pendiente")', "0.0%"),
        (34, "Error medio en victorias/temporada (modelo)", '=IF(ISNUMBER(Backtest!H2),Backtest!H2,"pendiente")', "0.00"),
        (35, "Error medio en victorias/temporada (base histórica)", '=IF(ISNUMBER(Backtest!H3),Backtest!H3,"pendiente")', "0.00"),
        (37, "BAJAS", None, None),
        (38, "Índice de bajas (L)", '=IF(ISNUMBER(Predicciones!L5),Predicciones!L5,"pendiente")', "0.00"),
        (39, "Jugadores baja/duda", '=COUNTIF(Bajas!B2:B29,"Baja")+COUNTIF(Bajas!B2:B29,"Duda")+COUNTIF(Bajas!B2:B29,"Sanción")', "0"),
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
    for r in (18, 20):
        ws.cell(r, 1).font = Font(name=ARIAL, bold=True)
        ws.cell(r, 2).font = Font(name=ARIAL, bold=True)
    widths(ws, [52, 30])

    # ---- LEEME
    wl = wb.create_sheet("LEEME")
    for i, t in enumerate(meta["leeme"], 1):
        wl.cell(i, 1, t)
    wl.column_dimensions["A"].width = 130

    for sh in wb.worksheets:
        for row in sh.iter_rows():
            for c in row:
                f = c.font
                if f.name != ARIAL:
                    c.font = Font(name=ARIAL, size=f.size if f.size and f.size != 11 else 10, bold=f.bold,
                                  italic=f.italic, color=f.color)
    wb.save(path)


# ----------------------------------------------------------------------------- web estática (Netlify)
import html as _html
import json as _json


def _fmt_date(d):
    if d is None or d == "":
        return "—"
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
        metodo = "Estimación simple (sin datos descargados)"
    total = v + fut

    # próximo partido
    nx = pred.get("proximo") if pred else None
    if nx:
        nxt = (nx["rival"], "local" if nx["home"] else "visitante", _fmt_date(nx["fecha"]), nx["p"])
    else:
        nxt = (meta["next_rival"], "visitante" if meta["next_cond"] == "V" else "local" if meta["next_cond"] else "—", "07/11/2026" if not pred else "—", None)

    def prob_bar(pp):
        if not pp:
            return "<p class='muted'>Probabilidades disponibles tras la primera actualización con datos.</p>"
        return ("<div class='bar'>"
                f"<span class='w' style='width:{100*pp[0]:.1f}%'>{_pct(pp[0])}</span>"
                f"<span class='d' style='width:{100*pp[1]:.1f}%'>{_pct(pp[1])}</span>"
                f"<span class='l' style='width:{100*pp[2]:.1f}%'>{_pct(pp[2])}</span></div>"
                "<p class='muted'>Victoria · Empate · Derrota</p>")

    # tablas
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
    if bt:
        btx = ("<table><tr><th>Modelo</th><th>Precisión</th><th>Brier</th><th>LogLoss</th></tr>" + "".join(
            f"<tr><td>{n}</td><td>{100*bt[k]['acc']:.1f} %</td><td>{bt[k]['brier']:.3f}</td><td>{bt[k]['ll']:.3f}</td></tr>"
            for n, k in (("Mixto Poisson+Elo", "mixto"), ("Solo Poisson", "poisson"), ("Solo Elo", "elo"), ("Base histórica", "base")))
               + "</table>"
               f"<p class='muted'>Error medio en victorias por temporada: modelo {bt['mae_model']:.2f} frente a {bt['mae_base']:.2f} "
               f"de la media histórica. Peso Poisson elegido: {bt['w']:.1f}. Pruebas partido a partido sobre {bt['mixto']['n']} partidos "
               "(cada uno predicho solo con datos anteriores). El ajuste por bajas no entra en el backtest.</p>")
    else:
        btx = "<p class='muted'>Pendiente: se calcula cuando el despliegue descarga los resultados de LaLiga.</p>"

    # gráfico histórico (SVG inline)
    bars, bw, gap = "", 40, 18
    for i, (t, w_, _, _) in enumerate(hist):
        h_ = 170 * w_ / 38
        x = 10 + i * (bw + gap)
        bars += (f"<rect x='{x}' y='{190-h_:.1f}' width='{bw}' height='{h_:.1f}' class='bar-svg'/>"
                 f"<text x='{x+bw/2}' y='{184-h_:.1f}' text-anchor='middle' class='lbl'>{w_}</text>"
                 f"<text x='{x+bw/2}' y='208' text-anchor='middle' class='lbl small'>{t[2:]}</text>")
    svg = f"<svg viewBox='0 0 {10+len(hist)*(bw+gap)} 215' role='img' aria-label='Victorias por temporada'>{bars}</svg>"
    media = sum(h[1] for h in hist) / len(hist)

    page = f"""<!doctype html>
<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Predictor Real Madrid · LaLiga</title>
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
.bar-svg{{fill:var(--bar)}}.lbl{{fill:var(--text);font-size:12px}}.small{{font-size:10px;fill:var(--muted)}}a{{color:var(--accent)}}
</style></head><body><main>
<h1>Predictor Real Madrid · LaLiga 2026-27</h1>
<p class="muted">{e(meta['estado'])}</p>
<div class="grid">
<div class="card"><div class="muted">Partidos jugados</div><div class="kpi">{pj}</div><div class="muted">{v}V · {em}E · {d}D · {pts} pts</div></div>
<div class="card"><div class="muted">Victorias finales previstas</div><div class="kpi">{total:.1f}</div><div class="muted">Rango P10–P90: {p10} – {p90} (central {p50})</div></div>
<div class="card"><div class="muted">Victorias restantes esperadas</div><div class="kpi">{fut:.1f}</div><div class="muted">de {rem} partidos</div></div>
</div>
<p class="muted">Método: {e(metodo)}.</p>
<div class="card"><h2>Próximo partido: {e(nxt[0])} ({nxt[1]}) · {nxt[2]}</h2>{prob_bar(nxt[3])}</div>
<div class="card"><h2>Partidos que quedan</h2>{'<div class="wrap"><table><tr><th>Rival</th><th>L/V</th><th>Fecha</th><th>Victoria</th><th>Empate</th><th>Derrota</th></tr>'+future+'</table></div>' if future else "<p class='muted'>Se rellena tras la primera actualización con datos.</p>"}</div>
<div class="card"><h2>Bajas y dudas</h2>{'<div class="wrap"><table><tr><th>Jugador</th><th>Estado</th><th>Motivo</th><th>Vuelta</th><th>Importancia</th></tr>'+bj+'</table></div>' if bj else "<p class='muted'>Sin bajas registradas en bajas.csv.</p>"}<p class="muted">Importancia: 1 = indispensable. El ajuste por bajas no está calibrado (no hay histórico gratuito de lesiones).</p></div>
<div class="card"><h2>Precisión del modelo (backtest)</h2>{btx}</div>
<div class="card"><h2>Resultados de la temporada</h2><div class="wrap"><table><tr><th>J</th><th>Fecha</th><th>Rival</th><th>L/V</th><th>Marcador</th><th>Res.</th></tr>{played}</table></div></div>
<div class="card"><h2>Histórico: victorias por temporada (media {media:.1f})</h2>{svg}</div>
<p class="muted">Datos de resultados: football-data.co.uk. Descarga el <a href="predictor_real_madrid.xlsx">Excel completo</a> o los <a href="data.json">datos en JSON</a>. No es una herramienta de apuestas: las probabilidades son estimaciones.</p>
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


LEEME = [
    "PREDICTOR REAL MADRID - guía rápida",
    "",
    "1) Dashboard: resumen de temporada, previsión de victorias, próximo partido, precisión y bajas.",
    "2) Bajas: se editan en bajas.csv (Estado e Importancia); esta hoja se regenera.",
    "3) Para actualizar: ejecuta 'python update.py' o despliega en Netlify (ver LEEME.md). Las bajas se editan en bajas.csv.",
    "",
    "Qué está verificado y qué no:",
    "- Histórico de victorias 2016-17 a 2025-26 y estado de la temporada 2026-27 hasta la jornada 11: contrastado con prensa.",
    "- Predicción partido a partido (Poisson + Elo), backtest y calendario: se calculan al ejecutar update.py con internet.",
    "- El ajuste por bajas NO está calibrado (no existe histórico gratuito de lesiones). Es un ajuste prudente, no una verdad.",
    "- xG: no incluido (la fuente gratuita no lo ofrece).",
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-download", action="store_true")
    ap.add_argument("--today", default=None)
    ap.add_argument("--out", default=str(OUT))
    ap.add_argument("--site", action="store_true", help="genera ./site (index.html, data.json y el Excel) para Netlify")
    a = ap.parse_args()
    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    out = Path(a.out)
    if a.site:
        out = BASE / "site" / "predictor_real_madrid.xlsx"
        out.parent.mkdir(exist_ok=True)
    refresh = not a.no_download

    print("1/5 Cargando datos...")
    df, cur = load_all(today, refresh)
    fixtures = load_fixtures(refresh)
    bajas = read_bajas(out)
    bajas, jugadores = api_football(today, bajas) if refresh else (bajas, [])

    hist, temporada, pred, bt = SEED_HIST, [list(x) + [None] for x in SEED_2627], None, None
    meta = {"leeme": LEEME, "next_rival": "Valencia", "next_cond": "V", "next_date": "DATE(2026,11,7)",
            "estado": f"Sin datos descargados: mostrando datos verificados en prensa a 1-nov-2026. Generado {today:%d/%m/%Y}."}

    if df is not None:
        print(f"   {len(df)} partidos cargados.")
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
        print("2/5 Backtest walk-forward...")
        bt = backtest(df, cur)
        w = bt["w"] if bt else 0.5
        print("3/5 Bajas y ajuste...")
        L = indice_bajas(bajas)
        print("4/5 Predicciones...")
        rows = predecir(df, cur, today, ratings, w, L, fixtures)
        if rows:
            ps = [r["p"][0] for r in rows]
            dist = distribucion(ps)
            wins_now = int((cm.Res == "V").sum()) if len(cm) else 0
            cdf = np.cumsum(dist)
            pct = lambda q: int(np.searchsorted(cdf, q)) + wins_now  # noqa: E731
            pts_now = 3 * wins_now + int((cm.Res == "E").sum()) if len(cm) else 0
            pts = pts_now + sum(3 * r["p"][0] + r["p"][1] for r in rows)
            dated = [r for r in rows if r["fecha"] is not None]
            pred = {
                "rows": rows,
                "resumen": [pct(0.10), pct(0.50), pct(0.90), round(float(pts), 1), round(L, 2), w,
                            round(float(ratings.get(MADRID, 1500)), 0), dt.datetime.now().strftime("%d/%m/%Y %H:%M")],
                "proximo": dated[0] if dated else None,
            }
            meta["estado"] = f"Actualizado {dt.datetime.now():%d/%m/%Y %H:%M} con {len(df)} partidos de LaLiga (football-data.co.uk)."
            if pred["proximo"] is not None:
                meta["next_rival"] = pred["proximo"]["rival"]
                meta["next_cond"] = "L" if pred["proximo"]["home"] else "V"
            else:
                meta["next_rival"], meta["next_cond"], meta["next_date"] = "(sin fecha en fixtures.csv)", "", '""'
    else:
        print("   AVISO: no hay datos (sin internet ni caché). Se genera el Excel con los datos verificados.")

    print("5/5 Escribiendo Excel...")
    write_workbook(out, hist, temporada, bajas, jugadores, pred, bt, meta)
    if a.site:
        write_site(out.parent, {"hist": hist, "temporada": temporada, "pred": pred, "bt": bt, "bajas": bajas, "meta": meta})
        print("Web generada en:", out.parent)
    print("Listo:", out)
    if bt:
        print(f"Backtest: precisión {bt['mixto']['acc']:.1%}, logloss {bt['mixto']['ll']:.3f} "
              f"(base {bt['base']['ll']:.3f}), MAE victorias {bt['mae_model']:.2f} (base {bt['mae_base']:.2f}), w={bt['w']:.1f}")


if __name__ == "__main__":
    sys.exit(main())
