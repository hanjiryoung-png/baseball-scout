

import os, re, time, shutil, sqlite3
from pathlib import Path
from datetime import datetime

import pandas as pd
import requests
import streamlit as st

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("DATA_DIR", str(APP_DIR / "data")))
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = DATA_DIR / "baseball_scout.db"
SNAPSHOT_PATH = DATA_DIR / "presentation_snapshot.db"

# 2026 Play-by-Play 기본 데이터
# - GitHub 저장소 루트에 파일이 있으면 그 파일을 우선 사용
# - 없으면 데이터 관리에서 업로드한 파일을 DATA_DIR에 보관하여 사용
REPO_BASE_PARQUET = APP_DIR / "kbo_pbp_2026.parquet"
UPLOADED_BASE_PARQUET = DATA_DIR / "kbo_pbp_2026.parquet"

# KBO 공식 시즌 기록 Excel
REPO_KBO_RECORDS = APP_DIR / "kbo_2026_records.xlsx"

TEAM = {
    "KT":"KT","LG":"LG","SS":"삼성","OB":"두산","HT":"KIA",
    "SK":"SSG","WO":"키움","LT":"롯데","NC":"NC","HH":"한화"
}
PRIORITY_TEAMS = ["KT","LG","삼성","두산","KIA"]


@st.cache_data(show_spinner=False)
def load_kbo_records(path_text, modified_ns):
    path = Path(path_text)

    sheets = pd.read_excel(
        path,
        sheet_name=["타자_기본","타자_세부","투수_기본","투수_세부"],
        engine="openpyxl"
    )

    def clean(df):
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]
        for c in ["선수명","팀명"]:
            if c in df.columns:
                df[c] = df[c].astype(str).str.strip()
        return df

    hb = clean(sheets["타자_기본"])
    hd = clean(sheets["타자_세부"])
    pb = clean(sheets["투수_기본"])
    pd2 = clean(sheets["투수_세부"])

    # 기본기록 + 세부기록을 선수명/팀명 기준으로 결합
    hitter = hb.merge(
        hd,
        on=["선수명","팀명"],
        how="outer",
        suffixes=("","_세부")
    )

    pitcher = pb.merge(
        pd2,
        on=["선수명","팀명"],
        how="outer",
        suffixes=("","_세부")
    )

    return hitter, pitcher

def kbo_records():
    if not REPO_KBO_RECORDS.exists():
        return pd.DataFrame(), pd.DataFrame()
    try:
        stat = REPO_KBO_RECORDS.stat()
        return load_kbo_records(str(REPO_KBO_RECORDS), stat.st_mtime_ns)
    except Exception:
        return pd.DataFrame(), pd.DataFrame()

def find_kbo_player(df, player_name, team_name):
    if df.empty:
        return None

    d = df[df["선수명"].astype(str).str.strip() == str(player_name).strip()].copy()

    # 이름이 같고 팀도 일치하면 우선 사용
    if team_name and "팀명" in d.columns:
        same_team = d[d["팀명"].astype(str).str.strip() == str(team_name).strip()]
        if not same_team.empty:
            return same_team.iloc[0]

    # 이름이 유일하면 팀 매칭이 안 되어도 사용
    if len(d) == 1:
        return d.iloc[0]

    return None

def metric_value(row, key):
    if row is None or key not in row.index:
        return "-"
    v = row[key]
    if pd.isna(v):
        return "-"
    if isinstance(v, float):
        if v.is_integer():
            return f"{int(v)}"
        return f"{v:.3f}".rstrip("0").rstrip(".")
    return str(v)


def kbo_avg_value(row, key="AVG"):
    """KBO 타율은 공식 표기처럼 항상 소수점 셋째 자리까지 표시."""
    if row is None or key not in row.index:
        return "-"
    v = row[key]
    if pd.isna(v):
        return "-"
    try:
        return f"{float(v):.3f}"
    except Exception:
        return str(v)

def kbo_ip_value(row, key="IP"):
    """KBO 이닝의 .333/.667 값을 야구식 1/3, 2/3 표기로 표시."""
    if row is None or key not in row.index:
        return "-"
    v = row[key]
    if pd.isna(v):
        return "-"

    # Excel에서 문자열(예: '111 1/3')로 들어온 경우 그대로 사용
    if isinstance(v, str):
        s = v.strip()
        if s:
            return s
        return "-"

    try:
        x = float(v)
    except Exception:
        return str(v)

    whole = int(x)
    frac = x - whole

    if abs(frac) < 0.01:
        return str(whole)
    if abs(frac - 1/3) < 0.03 or abs(frac - 0.333) < 0.03:
        return f"{whole} 1/3"
    if abs(frac - 2/3) < 0.03 or abs(frac - 0.667) < 0.03:
        return f"{whole} 2/3"

    return f"{x:.3f}".rstrip("0").rstrip(".")

def db():
    c = sqlite3.connect(DB_PATH, timeout=30)
    c.execute("PRAGMA journal_mode=WAL")
    c.execute("PRAGMA foreign_keys=ON")
    return c

def init_db():
    c = db()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS games(
        game_id TEXT PRIMARY KEY,
        game_date TEXT NOT NULL,
        away_team TEXT,
        home_team TEXT,
        source TEXT,
        saved_at TEXT,
        innings INTEGER,
        demo_ready INTEGER DEFAULT 1
    );

    CREATE TABLE IF NOT EXISTS favorites(
        player_id TEXT PRIMARY KEY,
        player_name TEXT NOT NULL,
        saved_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS pitches(
        pitch_id TEXT PRIMARY KEY,
        game_id TEXT NOT NULL,
        inning INTEGER,
        half TEXT,
        offense_team TEXT,
        defense_team TEXT,
        pitcher_id TEXT,
        pitcher_name TEXT,
        batter_id TEXT,
        batter_name TEXT,
        pitch_num INTEGER,
        pitch_type TEXT,
        speed REAL,
        result_code TEXT,
        pitch_text TEXT,
        balls INTEGER,
        strikes INTEGER,
        outs INTEGER,
        plate_x REAL,
        plate_y REAL,
        pa_result TEXT,
        FOREIGN KEY(game_id) REFERENCES games(game_id)
    );

    CREATE INDEX IF NOT EXISTS ix_pitches_game ON pitches(game_id);
    CREATE INDEX IF NOT EXISTS ix_pitches_pitcher ON pitches(pitcher_id);
    CREATE INDEX IF NOT EXISTS ix_pitches_batter ON pitches(batter_id);
    CREATE INDEX IF NOT EXISTS ix_pitches_offense ON pitches(offense_team);
    CREATE INDEX IF NOT EXISTS ix_pitches_defense ON pitches(defense_team);
    """)
    c.commit()
    c.close()

def snapshot_db():
    # 발표용 백업: 새 경기 저장이 성공할 때마다 마지막 정상 DB를 별도 파일로 보관
    c = db()
    b = sqlite3.connect(SNAPSHOT_PATH)
    with b:
        c.backup(b)
    b.close()
    c.close()

def restore_snapshot():
    if not SNAPSHOT_PATH.exists():
        return False
    # main DB를 직접 덮기 전에 잠깐 연결 해제 상태에서 복원
    temp = DATA_DIR / "restore_tmp.db"
    shutil.copy2(SNAPSHOT_PATH, temp)
    shutil.copy2(temp, DB_PATH)
    temp.unlink(missing_ok=True)
    return True

def qdf(sql, params=()):
    c = db()
    d = pd.read_sql_query(sql, c, params=params)
    c.close()
    return d

def gid_from_text(s):
    m = re.search(r'([0-9]{8}[A-Z]{4}[0-9]{5})', s or "")
    return m.group(1) if m else None

def teams_from_gid(gid):
    return TEAM.get(gid[8:10], gid[8:10]), TEAM.get(gid[10:12], gid[10:12])

def date_from_gid(gid):
    try:
        return datetime.strptime(gid[:8], "%Y%m%d").strftime("%Y-%m-%d")
    except:
        return gid[:8]

def game_exists(gid):
    c = db()
    row = c.execute("SELECT 1 FROM games WHERE game_id=?", (gid,)).fetchone()
    c.close()
    return bool(row)

def request_inning(gid, inning):
    url = f"https://api-gw.sports.naver.com/schedule/games/{gid}/relay?inning={inning}"
    r = requests.get(
        url,
        timeout=15,
        headers={
            "User-Agent":"Mozilla/5.0",
            "Accept":"application/json,text/plain,*/*"
        }
    )
    if r.status_code in (403, 429):
        raise RuntimeError("데이터 제공 서버가 요청을 제한했습니다. 자동 재시도 없이 중단했습니다.")
    r.raise_for_status()
    return r.json()

def infer_max_inning(payload):
    tr = payload.get("result",{}).get("textRelayData",{})
    score = tr.get("inningScore",{}) or {}
    nums = []
    for side in ("home","away"):
        for k in (score.get(side,{}) or {}).keys():
            try:
                nums.append(int(k))
            except:
                pass
    return max(nums) if nums else 9

def player_map(tr):
    names = {}
    for k in ("homeLineup","awayLineup"):
        b = tr.get(k,{}) or {}
        for role in ("batter","pitcher"):
            for p in b.get(role,[]) or []:
                if p.get("pcode"):
                    names[str(p["pcode"])] = p.get("name")
    for k in ("homeEntry","awayEntry"):
        b = tr.get(k,{}) or {}
        for role in ("batter","pitcher"):
            for p in b.get(role,[]) or []:
                pid = str(p.get("pcode",""))
                if pid and pid not in names:
                    names[pid] = p.get("name")
    return names

def extract(payload, gid, away, home):
    tr = payload.get("result",{}).get("textRelayData",{})
    names = player_map(tr)
    rows = {}

    for rel in tr.get("textRelays",[]) or []:
        inning = rel.get("inn")
        half = "말" if str(rel.get("homeOrAway","")) == "1" else "초"
        offense = home if half == "말" else away
        defense = away if half == "말" else home
        options = rel.get("textOptions",[]) or []
        tracks = {
            str(x.get("pitchId")):x
            for x in (rel.get("ptsOptions",[]) or [])
            if x.get("pitchId")
        }
        pitch_options = [
            x for x in options
            if x.get("pitchNum") is not None or x.get("ptsPitchId") or x.get("type")==1
        ]
        if not pitch_options:
            continue

        finals = [x.get("text") for x in options if x.get("type")==13 and x.get("text")]
        final_text = finals[-1] if finals else None

        for t in pitch_options:
            state = t.get("currentGameState",{}) or {}
            track_id = str(t.get("ptsPitchId") or "")
            pitch_id = track_id or f"{gid}|{rel.get('no')}|{t.get('pitchNum')}|{t.get('seqno')}"
            track = tracks.get(track_id,{})
            pitcher_id = str(state.get("pitcher",""))
            batter_id = str(state.get("batter",""))

            try:
                speed = float(t.get("speed")) if t.get("speed") not in (None,"") else None
            except:
                speed = None

            rows[pitch_id] = (
                pitch_id, gid, inning, half, offense, defense,
                pitcher_id, names.get(pitcher_id),
                batter_id, names.get(batter_id),
                t.get("pitchNum"), t.get("stuff"), speed,
                t.get("pitchResult"), t.get("text"),
                int(state.get("ball")) if str(state.get("ball","")).isdigit() else None,
                int(state.get("strike")) if str(state.get("strike","")).isdigit() else None,
                int(state.get("out")) if str(state.get("out","")).isdigit() else None,
                track.get("crossPlateX"), track.get("crossPlateY"),
                final_text
            )
    return list(rows.values())

def collect_game(text):
    gid = gid_from_text(text)
    if not gid:
        return "error", "경기 주소를 확인해 주세요."

    # 핵심 원칙: DB 확인이 외부 요청보다 먼저
    if game_exists(gid):
        return "cached", "이미 추가된 경기입니다."

    away, home = teams_from_gid(gid)
    first = request_inning(gid, 1)
    max_inn = infer_max_inning(first)
    payloads = [first]

    for inning in range(2, max_inn + 1):
        time.sleep(1.0)
        payloads.append(request_inning(gid, inning))

    all_rows = {}
    for payload in payloads:
        for row in extract(payload, gid, away, home):
            all_rows[row[0]] = row

    c = db()
    try:
        c.execute("BEGIN")
        c.execute(
            "INSERT INTO games(game_id,game_date,away_team,home_team,source,saved_at,innings,demo_ready) VALUES(?,?,?,?,?,?,?,1)",
            (gid, date_from_gid(gid), away, home, text, datetime.now().isoformat(timespec="seconds"), max_inn)
        )
        c.executemany(
            "INSERT OR IGNORE INTO pitches VALUES(" + ",".join(["?"]*21) + ")",
            list(all_rows.values())
        )
        c.commit()
    except:
        c.rollback()
        raise
    finally:
        c.close()

    # 성공한 순간 발표용 백업 갱신
    snapshot_db()
    return "new", f"경기가 추가되었습니다. ({away} vs {home})"

def action(text):
    s = str(text or "")
    if "헛스윙" in s: return "헛스윙"
    if "파울" in s: return "파울"
    if "타격" in s: return "인플레이"
    if "스트라이크" in s: return "스트라이크"
    if "볼" in s: return "볼"
    return "기타"

def is_favorite(player_id):
    c = db()
    row = c.execute("SELECT 1 FROM favorites WHERE player_id=?", (player_id,)).fetchone()
    c.close()
    return bool(row)

def toggle_favorite(player_id, player_name):
    c = db()
    row = c.execute("SELECT 1 FROM favorites WHERE player_id=?", (player_id,)).fetchone()
    if row:
        c.execute("DELETE FROM favorites WHERE player_id=?", (player_id,))
        action = "removed"
    else:
        c.execute(
            "INSERT OR REPLACE INTO favorites(player_id,player_name,saved_at) VALUES(?,?,?)",
            (player_id, player_name, datetime.now().isoformat(timespec="seconds"))
        )
        action = "added"
    c.commit()
    c.close()
    snapshot_db()
    return action

def favorite_players():
    return qdf("SELECT player_id, player_name FROM favorites ORDER BY saved_at DESC")



BASE_COLUMNS = [
    "game_pk","game_date","home_team","away_team","inning","inning_topbot",
    "at_bat_number","pitch_number","batter","pitcher","batter_name","pitcher_name",
    "balls","strikes","outs_when_up","pitch_result","type",
    "pitch_name","_naver_pitch_name","release_speed_kmh","plate_x","plate_z","events"
]

def get_base_parquet_path():
    if REPO_BASE_PARQUET.exists():
        return REPO_BASE_PARQUET
    if UPLOADED_BASE_PARQUET.exists():
        return UPLOADED_BASE_PARQUET
    return None

@st.cache_data(show_spinner=False)
def load_base_parquet(path_text, modified_ns):
    path = Path(path_text)
    raw = pd.read_parquet(path, columns=BASE_COLUMNS)

    # 현재 NAVER/SQLite 분석 화면에서 쓰는 공통 형식으로 변환
    home = raw["home_team"].map(TEAM).fillna(raw["home_team"])
    away = raw["away_team"].map(TEAM).fillna(raw["away_team"])
    is_bottom = raw["inning_topbot"].astype(str).str.lower().isin(["bot","bottom","말"])

    result_map = {
        "T":"스트라이크",
        "B":"볼",
        "F":"파울",
        "S":"헛스윙",
        "H":"타격",
    }

    pitch_type = raw["_naver_pitch_name"].copy()
    pitch_type = pitch_type.where(pitch_type.notna() & (pitch_type.astype(str) != ""), raw["pitch_name"])

    d = pd.DataFrame({
        "pitch_id": (
            raw["game_pk"].astype(str) + "|" +
            raw["at_bat_number"].astype(str) + "|" +
            raw["pitch_number"].astype(str)
        ),
        "game_id": raw["game_pk"].astype(str),
        "game_date": raw["game_date"].astype(str).str[:10],
        "inning": raw["inning"],
        "half": is_bottom.map({True:"말", False:"초"}),
        "offense_team": home.where(is_bottom, away),
        "defense_team": away.where(is_bottom, home),
        "pitcher_id": raw["pitcher"].astype(str),
        "pitcher_name": raw["pitcher_name"],
        "batter_id": raw["batter"].astype(str),
        "batter_name": raw["batter_name"],
        "pitch_num": raw["pitch_number"],
        "pitch_type": pitch_type,
        "speed": raw["release_speed_kmh"],
        "result_code": raw["pitch_result"],
        "pitch_text": raw["pitch_result"].map(result_map).fillna(raw["pitch_result"]),
        "balls": raw["balls"],
        "strikes": raw["strikes"],
        "outs": raw["outs_when_up"],
        "plate_x": raw["plate_x"],
        "plate_y": raw["plate_z"],
        "pa_result": raw["events"],
    })

    games = (
        pd.DataFrame({
            "game_id": raw["game_pk"].astype(str),
            "game_date": raw["game_date"].astype(str).str[:10],
            "away_team": away,
            "home_team": home,
        })
        .drop_duplicates("game_id")
        .reset_index(drop=True)
    )
    return d, games

def base_data():
    path = get_base_parquet_path()
    if path is None:
        return pd.DataFrame(), pd.DataFrame()
    try:
        stat = path.stat()
        return load_base_parquet(str(path), stat.st_mtime_ns)
    except Exception:
        return pd.DataFrame(), pd.DataFrame()

def extra_pitches():
    # NAVER 최신 자료에도 경기일을 붙여 통합 분석 기간이 실제 최신 경기까지 확장되도록 함
    return qdf("""
        SELECT p.*, g.game_date
        FROM pitches p
        LEFT JOIN games g ON p.game_id = g.game_id
    """)

def extra_games():
    return qdf("SELECT * FROM games")

def all_pitches():
    base_p, base_g = base_data()
    extra = extra_pitches()

    if base_p.empty:
        return extra
    if extra.empty:
        return base_p

    # 같은 경기 ID가 기본 Parquet에 이미 있으면 SQLite 쪽 중복 경기 제외
    extra = extra[~extra["game_id"].astype(str).isin(set(base_p["game_id"].astype(str)))]
    cols = list(base_p.columns)
    for col in cols:
        if col not in extra.columns:
            extra[col] = None
    return pd.concat([base_p, extra[cols]], ignore_index=True)

def all_games():
    base_p, base_g = base_data()
    extra = extra_games()

    if base_g.empty:
        return extra
    if extra.empty:
        g = base_g.copy()
        g["source"] = "Play-by-Play"
        g["saved_at"] = ""
        g["innings"] = None
        g["demo_ready"] = 1
        return g

    extra = extra[~extra["game_id"].astype(str).isin(set(base_g["game_id"].astype(str)))]
    g = base_g.copy()
    g["source"] = "Play-by-Play"
    g["saved_at"] = ""
    g["innings"] = None
    g["demo_ready"] = 1
    return pd.concat([g[extra.columns], extra], ignore_index=True)


def source_pitches():
    """Parquet 기본 데이터와 NAVER 추가 데이터를 분리해서 반환."""
    base_p, _ = base_data()
    extra = extra_pitches()

    if not base_p.empty and not extra.empty:
        extra = extra[
            ~extra["game_id"].astype(str).isin(set(base_p["game_id"].astype(str)))
        ].copy()

    return base_p.copy(), extra.copy()

def source_games():
    """Parquet 기본 경기와 NAVER 추가 경기를 분리해서 반환."""
    _, base_g = base_data()
    extra = extra_games()

    if not base_g.empty and not extra.empty:
        extra = extra[
            ~extra["game_id"].astype(str).isin(set(base_g["game_id"].astype(str)))
        ].copy()

    return base_g.copy(), extra.copy()

def analysis_period(df):
    """실제 분석에 사용된 데이터의 최소/최대 경기일을 자동 계산."""
    if df is None or df.empty or "game_date" not in df.columns:
        return None

    dates = pd.to_datetime(df["game_date"], errors="coerce").dropna()
    if dates.empty:
        return None

    start = dates.min().strftime("%Y.%m.%d")
    end = dates.max().strftime("%Y.%m.%d")
    return start if start == end else f"{start} ~ {end}"

def source_caption(text):
    st.caption(text)

def top_pitch_info(df):
    if df is None or df.empty or "pitch_type" not in df.columns:
        return "-", 0, 0.0
    s = df["pitch_type"].dropna().astype(str)
    s = s[s.str.strip() != ""]
    if s.empty:
        return "-", 0, 0.0
    counts = s.value_counts()
    name = counts.index[0]
    count = int(counts.iloc[0])
    share = count / len(df) * 100 if len(df) else 0.0
    return name, count, share


def avg_speed_value(df):
    if df is None or df.empty or "speed" not in df.columns:
        return "-"
    s = pd.to_numeric(df["speed"], errors="coerce").dropna()
    return f"{s.mean():.1f} km/h" if not s.empty else "-"


def data_source_status():
    base_p, base_g = base_data()
    _, naver_g = source_games()
    kbo_h, kbo_p = kbo_records()
    return {
        "kbo": (not kbo_h.empty) or (not kbo_p.empty),
        "naver_games": int(naver_g["game_id"].nunique()) if not naver_g.empty else 0,
        "base_games": int(base_g["game_id"].nunique()) if not base_g.empty else 0,
    }


def analysis_header(title, period=None, source_note=None):
    html = (
        '<div class="dugout-analysis-head">'
        '<div class="dugout-analysis-kicker">DATA DUGOUT ANALYSIS</div>'
        f'<div class="dugout-analysis-title">{title}</div>'
        '</div>'
    )
    st.markdown(html, unsafe_allow_html=True)
    if period:
        st.caption(f"분석 기간: {period}")
    if source_note:
        st.caption(source_note)


def evidence_header():
    st.markdown('<div class="evidence-label">분석 근거 자료</div>', unsafe_allow_html=True)

def combined_counts():
    g = all_games()
    p = all_pitches()
    if p.empty:
        players = 0
    else:
        ids = pd.concat([
            p["pitcher_id"].dropna().astype(str),
            p["batter_id"].dropna().astype(str)
        ])
        ids = ids[ids != ""]
        players = ids.nunique()
    return {
        "games": int(g["game_id"].nunique()) if not g.empty else 0,
        "pitches": int(len(p)),
        "players": int(players)
    }

def import_parquet_to_db(raw):
    required = {
        "game_pk","game_date","home_team","away_team","inning","inning_topbot",
        "at_bat_number","pitch_number","batter","pitcher",
        "batter_name","pitcher_name","balls","strikes","outs_when_up"
    }
    missing = sorted(required - set(raw.columns))
    if missing:
        raise ValueError("필수 컬럼이 없습니다: " + ", ".join(missing))

    work = raw.copy()
    work["game_pk"] = work["game_pk"].astype(str)

    # 이미 DB에 있는 경기는 통째로 제외:
    # 같은 경기를 NAVER/Parquet에서 중복 저장하지 않음.
    c = db()
    existing = {
        str(r[0]) for r in c.execute("SELECT game_id FROM games").fetchall()
    }
    new_game_ids = [g for g in work["game_pk"].dropna().unique().tolist() if g not in existing]

    if not new_game_ids:
        c.close()
        return 0, 0, len(existing.intersection(set(work["game_pk"].unique())))

    work = work[work["game_pk"].isin(new_game_ids)].copy()

    # 경기 테이블
    game_rows = []
    for gid, g in work.groupby("game_pk", sort=False):
        home_code = str(g["home_team"].iloc[0])
        away_code = str(g["away_team"].iloc[0])
        home = TEAM.get(home_code, home_code)
        away = TEAM.get(away_code, away_code)
        game_date = str(g["game_date"].iloc[0])[:10]
        try:
            innings = int(pd.to_numeric(g["inning"], errors="coerce").max())
        except Exception:
            innings = None
        game_rows.append((
            gid, game_date, away, home,
            "Hugging Face · kbo_playbyplay v0",
            datetime.now().isoformat(timespec="seconds"),
            innings, 1
        ))

    c.executemany(
        """INSERT OR IGNORE INTO games
        (game_id,game_date,away_team,home_team,source,saved_at,innings,demo_ready)
        VALUES(?,?,?,?,?,?,?,?)""",
        game_rows
    )

    def val(row, col, default=None):
        if col not in work.columns:
            return default
        v = row.get(col, default)
        if pd.isna(v):
            return default
        return v

    pitch_rows = []
    for _, r in work.iterrows():
        gid = str(r["game_pk"])
        home_code = str(r["home_team"])
        away_code = str(r["away_team"])
        home = TEAM.get(home_code, home_code)
        away = TEAM.get(away_code, away_code)

        topbot = str(r["inning_topbot"]).lower()
        is_bottom = topbot in ("bot", "bottom", "말")
        half = "말" if is_bottom else "초"
        offense = home if is_bottom else away
        defense = away if is_bottom else home

        at_bat = int(r["at_bat_number"]) if not pd.isna(r["at_bat_number"]) else 0
        pitch_no = int(r["pitch_number"]) if not pd.isna(r["pitch_number"]) else 0
        pitch_id = f"{gid}|{at_bat}|{pitch_no}"

        pitch_type = (
            val(r, "_naver_pitch_name")
            or val(r, "pitch_name")
            or val(r, "pitch_type")
        )

        result_code = val(r, "pitch_result") or val(r, "type")
        pitch_text = val(r, "pitch_result")
        pa_result = val(r, "events")

        def as_int(x):
            try:
                return int(x) if x is not None and not pd.isna(x) else None
            except Exception:
                return None

        def as_float(x):
            try:
                return float(x) if x is not None and not pd.isna(x) else None
            except Exception:
                return None

        pitch_rows.append((
            pitch_id,
            gid,
            as_int(val(r, "inning")),
            half,
            offense,
            defense,
            str(val(r, "pitcher", "") or ""),
            str(val(r, "pitcher_name", "") or ""),
            str(val(r, "batter", "") or ""),
            str(val(r, "batter_name", "") or ""),
            pitch_no,
            str(pitch_type) if pitch_type is not None else None,
            as_float(val(r, "release_speed_kmh")),
            str(result_code) if result_code is not None else None,
            str(pitch_text) if pitch_text is not None else None,
            as_int(val(r, "balls")),
            as_int(val(r, "strikes")),
            as_int(val(r, "outs_when_up")),
            as_float(val(r, "plate_x")),
            as_float(val(r, "plate_z")),
            str(pa_result) if pa_result is not None else None
        ))

    c.executemany(
        "INSERT OR IGNORE INTO pitches VALUES(" + ",".join(["?"] * 21) + ")",
        pitch_rows
    )
    c.commit()

    inserted_games = len(new_game_ids)
    inserted_pitches = c.execute(
        "SELECT COUNT(*) FROM pitches WHERE game_id IN ({})".format(
            ",".join(["?"] * len(new_game_ids))
        ),
        new_game_ids
    ).fetchone()[0]
    c.close()

    snapshot_db()
    return inserted_games, inserted_pitches, len(existing.intersection(set(raw["game_pk"].astype(str).unique())))

def current_counts():
    c = combined_counts()
    return pd.Series(c)


st.set_page_config(page_title="K-BASEBALL DATA DUGOUT", page_icon="⚾", layout="wide")
init_db()

# 최초 실행 시 스냅샷이 없으면 빈 DB라도 백업 파일 생성
if not SNAPSHOT_PATH.exists():
    snapshot_db()

st.markdown("""
<style>
#MainMenu{visibility:hidden}
footer{visibility:hidden}
header{visibility:hidden}
.block-container{max-width:1180px;padding-top:1.5rem;padding-bottom:3rem}
.brand{font-size:2.2rem;font-weight:850;letter-spacing:-1.2px;margin-bottom:.1rem}
.tagline{font-size:1rem;color:#707070;margin-bottom:1.35rem}
.mode{font-size:.85rem;padding:.35rem .65rem;border-radius:999px;border:1px solid #ddd;display:inline-block}
div[data-testid="stMetric"]{border:1px solid #ececec;border-radius:18px;padding:16px;background:#fff}
div[data-testid="stDataFrame"]{border-radius:14px;overflow:hidden}
.dugout-analysis-head{margin-top:1rem;margin-bottom:.35rem;padding:20px 22px;border:2px solid #1f2937;border-radius:18px;background:linear-gradient(135deg,#fafafa 0%,#f1f5f9 100%)}
.dugout-analysis-kicker{font-size:.72rem;font-weight:800;letter-spacing:.14em;color:#64748b;margin-bottom:.25rem}
.dugout-analysis-title{font-size:1.55rem;font-weight:850;color:#111827;letter-spacing:-.035em}
.evidence-label{margin-top:2rem;padding-top:1rem;border-top:1px solid #e5e7eb;font-size:.82rem;font-weight:800;color:#64748b;letter-spacing:.08em}
.home-intro{padding:22px 24px;border:1px solid #e5e7eb;border-radius:18px;background:#fafafa;margin:.6rem 0 1.2rem 0}
.home-intro-title{font-size:1.2rem;font-weight:800;margin-bottom:.35rem}
.home-intro-text{color:#5f6368;line-height:1.55}

/* 앱 전체 제목 옆 Streamlit 자동 링크(앵커) 아이콘 숨김 */
[data-testid="stHeaderActionElements"]{
    display:none !important;
}

/* Streamlit 버전에 따라 heading 내부에 직접 생성되는 anchor까지 숨김 */
[data-testid="stMarkdownContainer"] h1 a,
[data-testid="stMarkdownContainer"] h2 a,
[data-testid="stMarkdownContainer"] h3 a,
[data-testid="stMarkdownContainer"] h4 a,
[data-testid="stMarkdownContainer"] h5 a,
[data-testid="stMarkdownContainer"] h6 a{
    display:none !important;
}
</style>
""", unsafe_allow_html=True)

# 발표 모드: 외부 요청 기능을 완전히 숨기고 저장 데이터만 사용
if "presentation_mode" not in st.session_state:
    st.session_state.presentation_mode = False

top1, top2 = st.columns([5,1])
with top1:
    st.markdown('<div class="brand">⚾ K-BASEBALL DATA DUGOUT</div>', unsafe_allow_html=True)

with top2:
    st.session_state.presentation_mode = st.toggle("발표 모드", value=st.session_state.presentation_mode,
                                                   help="저장된 데이터만 사용합니다. NAVER에 새 요청을 보내지 않습니다.")

if st.session_state.presentation_mode:
    st.caption("발표 모드 · 저장된 데이터만 사용 중")

nav_items = ["홈","팀","선수"] if st.session_state.presentation_mode else ["홈","팀","선수","데이터"]
if "main_nav" not in st.session_state or st.session_state.main_nav not in nav_items:
    st.session_state.main_nav = "홈"
nav = st.radio("", nav_items, horizontal=True, label_visibility="collapsed", key="main_nav")

def go_to(page, player_name=None):
    st.session_state.main_nav = page
    if player_name is not None:
        st.session_state.player_search = player_name


games = all_games()
if not games.empty:
    games = games.sort_values(["game_date","saved_at"], ascending=[False,False], na_position="last")
counts = current_counts()

if nav == "홈":
    allg = all_games()
    status = data_source_status()
    period = analysis_period(allg)

    st.markdown(
        '''<div class="home-intro">
            <div class="home-intro-title">흩어진 야구 자료를 한곳에서 검색하고 분석합니다.</div>
            <div class="home-intro-text">KBO 공식 기록, NAVER Sports 문자중계, 공개 Play-by-Play 자료를 연결해 팀과 선수 단위로 확인하고 DATA DUGOUT 분석으로 재구성합니다.</div>
        </div>''', unsafe_allow_html=True)

    h1,h2,h3 = st.columns(3)
    h1.metric("분석 경기", int(allg["game_id"].nunique()) if not allg.empty else 0)
    h2.metric("등록 선수", int(counts.players))
    h3.metric("분석 기간", period or "-")

    st.markdown("### 데이터 연결 상태")
    s1,s2,s3 = st.columns(3)
    s1.metric("KBO 공식 기록", "연결됨" if status["kbo"] else "미연결")
    s2.metric("NAVER 최신 자료", f"{status['naver_games']}경기")
    s3.metric("Play-by-Play 자료", f"{status['base_games']}경기")

    st.markdown("### 바로가기")
    b1,b2 = st.columns(2)
    with b1:
        st.button("팀 분석 보기", use_container_width=True, on_click=go_to, args=("팀",))
    with b2:
        st.button("선수 분석 보기", use_container_width=True, on_click=go_to, args=("선수",))

    favs = favorite_players()
    if not favs.empty:
        st.markdown("### 관심 선수")
        cols = st.columns(min(4, len(favs)))
        for pos, (_, row) in enumerate(favs.iterrows()):
            with cols[pos % len(cols)]:
                st.button(f"★ {row['player_name']}", key=f"home_fav_{row['player_id']}", use_container_width=True,
                          on_click=go_to, args=("선수", row["player_name"]))

elif nav == "데이터":
    st.markdown("## 데이터")
    st.caption("최신 경기 자료 추가와 Play-by-Play 자료 관리를 한 화면에서 구분해 처리합니다.")

    kbo_h, kbo_p = kbo_records()
    if (not kbo_h.empty) or (not kbo_p.empty):
        st.success("KBO 공식 기록 · 2026 KBO 정규시즌 · 연결됨")
    else:
        st.warning("KBO 공식 기록 파일을 확인해 주세요.")

    st.markdown("### ① NAVER 최신 자료 추가")
    st.caption("출처: NAVER Sports 문자중계")
    entries=[]
    for i in range(1,4):
        entries.append(st.text_input(f"경기 {i}", key=f"game_{i}", placeholder="NAVER Sports 경기 주소 붙여넣기"))
    if "extra_games" not in st.session_state: st.session_state.extra_games=0
    if st.button("+ 경기 입력칸 추가"):
        st.session_state.extra_games=min(2,st.session_state.extra_games+1); st.rerun()
    for j in range(st.session_state.extra_games):
        idx=j+4
        entries.append(st.text_input(f"경기 {idx}", key=f"game_{idx}", placeholder="NAVER Sports 경기 주소 붙여넣기"))
    if st.button("NAVER 최신 자료 가져오기", type="primary", use_container_width=True):
        values=[x.strip() for x in entries if x.strip()]
        if not values: st.warning("경기 주소를 입력해 주세요.")
        else:
            progress=st.progress(0)
            for i,x in enumerate(values[:5],start=1):
                try:
                    status_code,msg=collect_game(x)
                    if status_code=="new": st.success(msg)
                    elif status_code=="cached": st.info(msg)
                    else: st.error(msg)
                except Exception as e:
                    st.error(f"수집을 중단했습니다: {e}"); break
                progress.progress(i/len(values[:5]))

    st.divider()
    st.markdown("### ② Play-by-Play 자료 관리")
    st.caption("출처: 공개 Play-by-Play 데이터셋")
    base_path=get_base_parquet_path(); base_p,base_g=base_data()
    if base_path is not None and not base_p.empty:
        p1,p2,p3=st.columns(3)
        p1.metric("투구 자료",f"{len(base_p):,}")
        p2.metric("경기",f"{base_g['game_id'].nunique():,}")
        player_ids=pd.concat([base_p["pitcher_id"].dropna().astype(str),base_p["batter_id"].dropna().astype(str)])
        p3.metric("선수",f"{player_ids[player_ids!=''].nunique():,}")
        st.caption(f"자료 기간: {analysis_period(base_g) or '-'}")
    uploaded=st.file_uploader("Play-by-Play Parquet 파일 선택",type=["parquet"],accept_multiple_files=False,key="parquet_uploader")
    if uploaded is not None:
        try:
            temp_path=DATA_DIR/"kbo_pbp_2026.uploading"
            with open(temp_path,"wb") as f: f.write(uploaded.getbuffer())
            temp_path.replace(UPLOADED_BASE_PARQUET); load_base_parquet.clear(); base_p,base_g=base_data()
            if base_p.empty: st.error("파일을 읽지 못했습니다.")
            else:
                st.success("Play-by-Play 자료 업데이트가 완료되었습니다.")
                st.caption(f"자료 기간: {analysis_period(base_g) or '-'}")
        except Exception as e: st.error(f"Parquet 파일을 읽지 못했습니다: {e}")

elif nav == "팀":
    st.markdown("## 팀")
    allp=all_pitches(); allg=all_games(); base_p,naver_p=source_pitches(); base_g,naver_g=source_games()
    present=sorted(set(allg.away_team.dropna()).union(set(allg.home_team.dropna()))) if not allg.empty else []
    priority=[t for t in PRIORITY_TEAMS if t in present]; options=priority+[t for t in present if t not in priority]
    if not options: st.info("데이터를 먼저 추가해 주세요.")
    else:
        team=st.selectbox("팀 선택",options)
        team_games=allg[(allg["away_team"]==team)|(allg["home_team"]==team)].copy()
        thrown=allp[allp["defense_team"]==team].copy() if not allp.empty else pd.DataFrame()
        seen=allp[allp["offense_team"]==team].copy() if not allp.empty else pd.DataFrame()
        analysis_header(f"{team} 팀 분석",analysis_period(team_games),"현재 반영 자료: NAVER Sports + 공개 Play-by-Play 자료 · KBO 팀 공식 기록은 연결 전")
        with st.container(border=True):
            top_throw,_,top_throw_share=top_pitch_info(thrown); top_seen,_,top_seen_share=top_pitch_info(seen)
            a,b,c,d=st.columns(4)
            a.metric("통합 경기",team_games["game_id"].nunique() if not team_games.empty else 0); b.metric("투수진 투구",len(thrown)); c.metric("주 사용 구종",top_throw); d.metric("주 사용 비율",f"{top_throw_share:.1f}%" if top_throw!="-" else "-")
            a,b,c,d=st.columns(4)
            a.metric("타선이 본 투구",len(seen)); b.metric("가장 많이 본 구종",top_seen); c.metric("상대 구종 비율",f"{top_seen_share:.1f}%" if top_seen!="-" else "-"); d.metric("투수진 평균 구속",avg_speed_value(thrown))
            st.markdown("#### 투수진 구종 구성")
            if not thrown.empty:
                mix=thrown.groupby("pitch_type",dropna=True).agg(투구수=("pitch_id","count"),평균구속=("speed","mean")).reset_index(); mix["평균구속"]=mix["평균구속"].round(1)
                st.bar_chart(mix.set_index("pitch_type")["투구수"]); st.dataframe(mix.sort_values("투구수",ascending=False),hide_index=True,use_container_width=True)
            st.markdown("#### 타선이 상대해 온 구종")
            if not seen.empty:
                mix2=seen.groupby("pitch_type",dropna=True).size().reset_index(name="투구수"); st.bar_chart(mix2.set_index("pitch_type")["투구수"]); st.dataframe(mix2.sort_values("투구수",ascending=False),hide_index=True,use_container_width=True)
        evidence_header()
        st.markdown("### KBO 공식 기록"); st.caption("출처: KBO 공식 홈페이지"); st.info("KBO 팀 공식 기록 자료는 아직 연결되지 않았습니다. 팀 기록 확보 후 DATA DUGOUT 팀 분석에도 반영합니다.")
        st.markdown("### NAVER 최신 자료")
        nv_team_games=naver_g[(naver_g["away_team"]==team)|(naver_g["home_team"]==team)].copy() if not naver_g.empty else pd.DataFrame(); nv_thrown=naver_p[naver_p["defense_team"]==team].copy() if not naver_p.empty else pd.DataFrame(); nv_seen=naver_p[naver_p["offense_team"]==team].copy() if not naver_p.empty else pd.DataFrame()
        n1,n2,n3=st.columns(3); n1.metric("최신 경기",nv_team_games["game_id"].nunique() if not nv_team_games.empty else 0); n2.metric("투수진 최신 투구",len(nv_thrown)); n3.metric("타선 최신 상대 투구",len(nv_seen)); st.caption("출처: NAVER Sports 문자중계")
        st.markdown("### Play-by-Play 자료")
        bp_team_games=base_g[(base_g["away_team"]==team)|(base_g["home_team"]==team)].copy() if not base_g.empty else pd.DataFrame(); bp_thrown=base_p[base_p["defense_team"]==team].copy() if not base_p.empty else pd.DataFrame(); bp_seen=base_p[base_p["offense_team"]==team].copy() if not base_p.empty else pd.DataFrame()
        p1,p2,p3=st.columns(3); p1.metric("경기",bp_team_games["game_id"].nunique() if not bp_team_games.empty else 0); p2.metric("투수진 투구",len(bp_thrown)); p3.metric("타선이 본 투구",len(bp_seen)); st.caption(f"자료 기간: {analysis_period(bp_team_games) or '-'}"); st.caption("출처: 공개 Play-by-Play 데이터셋")

elif nav == "선수":
    allp=all_pitches()
    if allp.empty: st.info("데이터를 먼저 추가해 주세요.")
    else:
        batter_rows=allp[["batter_id","batter_name","offense_team","game_date"]].rename(columns={"batter_id":"id","batter_name":"name","offense_team":"team","game_date":"game_date"}); batter_rows["role"]="타자"
        pitcher_rows=allp[["pitcher_id","pitcher_name","defense_team","game_date"]].rename(columns={"pitcher_id":"id","pitcher_name":"name","defense_team":"team","game_date":"game_date"}); pitcher_rows["role"]="투수"
        role_rows=pd.concat([batter_rows,pitcher_rows],ignore_index=True).dropna(subset=["id","name"]); role_rows["id"]=role_rows["id"].astype(str); role_rows["name"]=role_rows["name"].astype(str); role_rows["team"]=role_rows["team"].fillna("").astype(str); role_rows["game_date"]=pd.to_datetime(role_rows["game_date"],errors="coerce")
        role_rows=role_rows[(role_rows["id"]!="")&(role_rows["name"]!="")&(role_rows["id"]!="nan")&(role_rows["name"]!="nan")]
        latest_team=role_rows.sort_values("game_date").drop_duplicates(["id","role"],keep="last")[["id","role","team"]]; player_master=role_rows[["id","name","role"]].drop_duplicates().merge(latest_team,on=["id","role"],how="left")
        grouped=[]
        for (pid,name),g in player_master.groupby(["id","name"],sort=False):
            roles=set(g["role"].dropna()); role_label="투타" if roles=={"타자","투수"} else ("투수" if "투수" in roles else "타자"); teams0=[t for t in g["team"].tolist() if t]; grouped.append({"id":str(pid),"name":name,"team":teams0[-1] if teams0 else "","role":role_label})
        players=pd.DataFrame(grouped).sort_values(["team","role","name"]).reset_index(drop=True)
        st.markdown("### 선수 찾기")
        if "player_search" not in st.session_state: st.session_state.player_search=""
        search_name=st.text_input("🔍 이름 검색",placeholder="선수 이름을 입력하세요",key="player_search").strip(); st.caption("이름으로 바로 검색하거나, 아래 필터로 선수 목록을 좁혀볼 수 있습니다.")
        f1,f2=st.columns(2); teams=["전체"]+[t for t in ["KT","LG","삼성","두산","KIA","NC","SSG","롯데","키움","한화"] if t in set(players["team"])]
        with f1: team_filter=st.selectbox("구단",teams)
        with f2: role_filter=st.selectbox("구분",["전체","투수","타자"])
        if search_name: filtered=players[players["name"].str.contains(search_name,case=False,na=False)].copy()
        else:
            filtered=players.copy()
            if team_filter!="전체": filtered=filtered[filtered["team"]==team_filter]
            if role_filter=="투수": filtered=filtered[filtered["role"].isin(["투수","투타"])]
            elif role_filter=="타자": filtered=filtered[filtered["role"].isin(["타자","투타"])]
        if filtered.empty: st.info("조건에 맞는 선수가 없습니다.")
        else:
            direct=bool(search_name and len(filtered)==1)
            if not direct:
                st.caption(f"선수 {len(filtered):,}명"); lv=filtered[["team","role","name"]].copy(); lv.columns=["구단","구분","선수"]; st.dataframe(lv,hide_index=True,use_container_width=True,height=min(420,38*(len(lv)+1)))
            player_id=player_name=player_team=player_role=None
            if direct:
                only=filtered.iloc[0]; player_id,player_name,player_team,player_role=str(only["id"]),only["name"],only["team"],only["role"]
            else:
                lm={f"{r['team']} · {r['role']} · {r['name']} ({r['id']})":(str(r['id']),r['name']) for _,r in filtered.iterrows()}; sel=st.selectbox("분석할 선수 선택",list(lm.keys()),index=None,placeholder="선수를 선택하세요")
                if sel is not None:
                    player_id,player_name=lm[sel]; picked=filtered[(filtered["id"].astype(str)==str(player_id))&(filtered["name"]==player_name)]
                    if not picked.empty: player_team,player_role=picked.iloc[0]["team"],picked.iloc[0]["role"]
            if player_id is not None:
                st.markdown(f"### {player_name}"); st.caption(" · ".join([x for x in [player_team,player_role] if x])); fav_now=is_favorite(player_id)
                if st.button("★ 관심 선수 해제" if fav_now else "☆ 관심 선수 등록",key=f"fav_{player_id}"): toggle_favorite(player_id,player_name); st.rerun()
                kbo_hitters,kbo_pitchers=kbo_records(); kbo_hitter=find_kbo_player(kbo_hitters,player_name,player_team); kbo_pitcher=find_kbo_player(kbo_pitchers,player_name,player_team)
                batter_data=allp[allp["batter_id"].astype(str)==str(player_id)].copy(); pitcher_data=allp[allp["pitcher_id"].astype(str)==str(player_id)].copy(); has_batter,has_pitcher=not batter_data.empty,not pitcher_data.empty
                base_player_p,naver_player_p=source_pitches(); bp_batter=base_player_p[base_player_p["batter_id"].astype(str)==str(player_id)].copy() if not base_player_p.empty else pd.DataFrame(); bp_pitcher=base_player_p[base_player_p["pitcher_id"].astype(str)==str(player_id)].copy() if not base_player_p.empty else pd.DataFrame(); nv_batter=naver_player_p[naver_player_p["batter_id"].astype(str)==str(player_id)].copy() if not naver_player_p.empty else pd.DataFrame(); nv_pitcher=naver_player_p[naver_player_p["pitcher_id"].astype(str)==str(player_id)].copy() if not naver_player_p.empty else pd.DataFrame()
                player_all=pd.concat([batter_data,pitcher_data],ignore_index=True) if (has_batter or has_pitcher) else pd.DataFrame(); analysis_header(f"{player_name} 선수 분석",analysis_period(player_all),"KBO 공식 기록 + NAVER Sports + 공개 Play-by-Play 자료를 근거로 구성")
                with st.container(border=True):
                    if has_pitcher:
                        top_pitch,_,top_share=top_pitch_info(pitcher_data); r1,r2,r3,r4=st.columns(4); r1.metric("KBO ERA",metric_value(kbo_pitcher,"ERA") if kbo_pitcher is not None else "-"); r2.metric("KBO WHIP",metric_value(kbo_pitcher,"WHIP") if kbo_pitcher is not None else "-"); r3.metric("평균 구속",avg_speed_value(pitcher_data)); r4.metric("주 사용 구종",top_pitch); st.caption(f"주 사용 구종 비율: {top_share:.1f}%" if top_pitch!="-" else "")
                        mix=pitcher_data.groupby("pitch_type",dropna=True).agg(투구수=("pitch_id","count"),평균구속=("speed","mean")).reset_index()
                        if not mix.empty: mix["평균구속"]=mix["평균구속"].round(1); st.markdown("#### 구종 구성"); st.bar_chart(mix.set_index("pitch_type")["투구수"]); st.dataframe(mix.sort_values("투구수",ascending=False),hide_index=True,use_container_width=True)
                        loc=pitcher_data[["plate_x","plate_y","pitch_type"]].dropna(subset=["plate_x","plate_y"])
                        if not loc.empty: st.markdown("#### 투구 위치"); st.scatter_chart(loc,x="plate_x",y="plate_y",color="pitch_type")
                    if has_batter:
                        top_seen,_,top_seen_share=top_pitch_info(batter_data); r1,r2,r3,r4=st.columns(4); r1.metric("KBO AVG",kbo_avg_value(kbo_hitter,"AVG") if kbo_hitter is not None else "-"); r2.metric("KBO HR",metric_value(kbo_hitter,"HR") if kbo_hitter is not None else "-"); r3.metric("상대 평균 구속",avg_speed_value(batter_data)); r4.metric("가장 많이 본 구종",top_seen); st.caption(f"가장 많이 본 구종 비율: {top_seen_share:.1f}%" if top_seen!="-" else "")
                        d=batter_data.copy(); d["행동"]=d.pitch_text.apply(action); acts=d["행동"].value_counts().rename_axis("결과").reset_index(name="횟수")
                        if not acts.empty: st.markdown("#### 투구 결과"); st.bar_chart(acts.set_index("결과")["횟수"])
                        m=d.groupby("pitch_type",dropna=True).size().reset_index(name="투구수")
                        if not m.empty: st.markdown("#### 상대 구종"); st.bar_chart(m.set_index("pitch_type")["투구수"])
                    nvg=pd.concat([nv_batter[["game_id"]] if not nv_batter.empty else pd.DataFrame(columns=["game_id"]),nv_pitcher[["game_id"]] if not nv_pitcher.empty else pd.DataFrame(columns=["game_id"])],ignore_index=True); st.caption(f"NAVER 최신 자료 반영: {nvg['game_id'].nunique() if not nvg.empty else 0}경기")
                evidence_header()
                st.markdown("### KBO 공식 기록"); st.caption("출처: KBO 공식 홈페이지")
                if kbo_hitter is None and kbo_pitcher is None: st.info("연결된 KBO 공식 기록이 없습니다.")
                else:
                    if kbo_hitter is not None:
                        st.markdown("#### 타자 공식 기록"); r1=st.columns(6); r1[0].metric("AVG",kbo_avg_value(kbo_hitter,"AVG")); r1[1].metric("경기",metric_value(kbo_hitter,"G")); r1[2].metric("타석",metric_value(kbo_hitter,"PA")); r1[3].metric("안타",metric_value(kbo_hitter,"H")); r1[4].metric("홈런",metric_value(kbo_hitter,"HR")); r1[5].metric("타점",metric_value(kbo_hitter,"RBI")); r2=st.columns(6); r2[0].metric("XBH",metric_value(kbo_hitter,"XBH")); r2[1].metric("BB/K",metric_value(kbo_hitter,"BB/K")); r2[2].metric("P/PA",metric_value(kbo_hitter,"P/PA")); r2[3].metric("ISOP",metric_value(kbo_hitter,"ISOP")); r2[4].metric("XR",metric_value(kbo_hitter,"XR")); r2[5].metric("GPA",metric_value(kbo_hitter,"GPA"))
                    if kbo_pitcher is not None:
                        st.markdown("#### 투수 공식 기록"); r1=st.columns(6); r1[0].metric("ERA",metric_value(kbo_pitcher,"ERA")); r1[1].metric("경기",metric_value(kbo_pitcher,"G")); r1[2].metric("승",metric_value(kbo_pitcher,"W")); r1[3].metric("패",metric_value(kbo_pitcher,"L")); r1[4].metric("이닝",kbo_ip_value(kbo_pitcher,"IP")); r1[5].metric("탈삼진",metric_value(kbo_pitcher,"SO")); r2=st.columns(6); r2[0].metric("WHIP",metric_value(kbo_pitcher,"WHIP")); r2[1].metric("세이브",metric_value(kbo_pitcher,"SV")); r2[2].metric("홀드",metric_value(kbo_pitcher,"HLD")); r2[3].metric("볼넷",metric_value(kbo_pitcher,"BB")); r2[4].metric("선발",metric_value(kbo_pitcher,"GS")); r2[5].metric("GO/AO",metric_value(kbo_pitcher,"GO/AO"))
                st.markdown("### NAVER 최신 자료")
                nv_games=pd.concat([nv_batter[["game_id","game_date"]] if not nv_batter.empty else pd.DataFrame(columns=["game_id","game_date"]),nv_pitcher[["game_id","game_date"]] if not nv_pitcher.empty else pd.DataFrame(columns=["game_id","game_date"])],ignore_index=True); n1,n2,n3=st.columns(3); n1.metric("최신 경기",nv_games["game_id"].nunique() if not nv_games.empty else 0); n2.metric("타자로 본 최신 투구",len(nv_batter)); n3.metric("투수로 던진 최신 투구",len(nv_pitcher));
                if not nv_games.empty: st.caption(f"자료 기간: {analysis_period(nv_games) or '-'}")
                st.caption("출처: NAVER Sports 문자중계")
                st.markdown("### Play-by-Play 자료")
                bp_games=pd.concat([bp_batter[["game_id","game_date"]] if not bp_batter.empty else pd.DataFrame(columns=["game_id","game_date"]),bp_pitcher[["game_id","game_date"]] if not bp_pitcher.empty else pd.DataFrame(columns=["game_id","game_date"])],ignore_index=True); p1,p2,p3=st.columns(3); p1.metric("경기",bp_games["game_id"].nunique() if not bp_games.empty else 0); p2.metric("타자로 본 투구",len(bp_batter)); p3.metric("투수로 던진 투구",len(bp_pitcher));
                if not bp_games.empty: st.caption(f"자료 기간: {analysis_period(bp_games) or '-'}")
                st.caption("출처: 공개 Play-by-Play 데이터셋")
