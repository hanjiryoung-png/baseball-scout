
import os, re, time, shutil, sqlite3, gc, json
from pathlib import Path
from datetime import datetime

import pandas as pd
import requests
import streamlit as st

APP_DIR = Path(__file__).resolve().parent

# Render Persistent Disk
# Render에서 설정한 Mount path(/var/data)가 있으면 그곳을 영구 저장소로 사용합니다.
# 로컬 실행처럼 /var/data가 없는 환경에서는 기존처럼 앱 폴더의 data를 사용합니다.
RENDER_DISK_DIR = Path("/var/data")
if RENDER_DISK_DIR.exists() and os.access(RENDER_DISK_DIR, os.W_OK):
    DATA_DIR = RENDER_DISK_DIR
else:
    DATA_DIR = Path(os.environ.get("DATA_DIR", str(APP_DIR / "data")))

DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = DATA_DIR / "baseball_scout.db"
SNAPSHOT_PATH = DATA_DIR / "presentation_snapshot.db"

# 2026 Play-by-Play 기본 데이터
# - GitHub 저장소 루트에 파일이 있으면 그 파일을 우선 사용
# - 없으면 데이터 관리에서 업로드한 파일을 DATA_DIR에 보관하여 사용
REPO_BASE_PARQUET = APP_DIR / "kbo_pbp_2026.parquet"
UPLOADED_BASE_PARQUET = DATA_DIR / "kbo_pbp_2026.parquet"

# KBO 공식 기록
# - 선수 기록: GitHub 기본 파일 또는 데이터 메뉴에서 업로드한 최신 파일
# - 팀 기록: 데이터 메뉴에서 업로드한 파일을 Persistent Disk에 저장
REPO_KBO_RECORDS = APP_DIR / "kbo_2026_records.xlsx"
UPLOADED_KBO_PLAYER_RECORDS = DATA_DIR / "kbo_player_records.xlsx"
UPLOADED_KBO_TEAM_RECORDS = DATA_DIR / "kbo_team_records.xlsx"
KBO_META_PATH = DATA_DIR / "kbo_metadata.json"

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


@st.cache_data(show_spinner=False)
def load_kbo_team_records(path_text, modified_ns):
    path = Path(path_text)
    sheets = pd.read_excel(
        path,
        sheet_name=["팀_타자","팀_투수","팀_수비","팀_주루"],
        engine="openpyxl"
    )

    out = {}
    for name, df in sheets.items():
        d = df.copy()
        d.columns = [str(c).strip() for c in d.columns]
        if "팀명" in d.columns:
            d["팀명"] = d["팀명"].fillna("").astype(str).str.strip()
            d = d[(d["팀명"] != "") & (d["팀명"].str.lower() != "nan")]
        out[name] = d.reset_index(drop=True)

    return out


def _read_kbo_meta():
    default = {
        "player_as_of": "2026-08-10",
        "team_as_of": None,
        "player_uploaded_at": None,
        "team_uploaded_at": None,
    }
    if not KBO_META_PATH.exists():
        return default
    try:
        saved = json.loads(KBO_META_PATH.read_text(encoding="utf-8"))
        default.update(saved if isinstance(saved, dict) else {})
    except Exception:
        pass
    return default


def _write_kbo_meta(**updates):
    meta = _read_kbo_meta()
    meta.update(updates)
    KBO_META_PATH.write_text(
        json.dumps(meta, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )


def _display_date(date_text):
    if not date_text:
        return "-"
    try:
        return pd.to_datetime(date_text).strftime("%Y.%m.%d")
    except Exception:
        return str(date_text)


def get_kbo_player_path():
    if UPLOADED_KBO_PLAYER_RECORDS.exists():
        return UPLOADED_KBO_PLAYER_RECORDS
    if REPO_KBO_RECORDS.exists():
        return REPO_KBO_RECORDS
    return None


def get_kbo_team_path():
    if UPLOADED_KBO_TEAM_RECORDS.exists():
        return UPLOADED_KBO_TEAM_RECORDS
    repo_team = APP_DIR / "kbo_2026_team_records.xlsx"
    if repo_team.exists():
        return repo_team
    return None


def kbo_records():
    path = get_kbo_player_path()
    if path is None:
        return pd.DataFrame(), pd.DataFrame()
    try:
        stat = path.stat()
        return load_kbo_records(str(path), stat.st_mtime_ns)
    except Exception:
        return pd.DataFrame(), pd.DataFrame()


def kbo_team_records():
    path = get_kbo_team_path()
    if path is None:
        return {
            "팀_타자": pd.DataFrame(),
            "팀_투수": pd.DataFrame(),
            "팀_수비": pd.DataFrame(),
            "팀_주루": pd.DataFrame(),
        }
    try:
        stat = path.stat()
        return load_kbo_team_records(str(path), stat.st_mtime_ns)
    except Exception:
        return {
            "팀_타자": pd.DataFrame(),
            "팀_투수": pd.DataFrame(),
            "팀_수비": pd.DataFrame(),
            "팀_주루": pd.DataFrame(),
        }


def kbo_player_caption():
    meta = _read_kbo_meta()
    return f"2026 정규시즌 · {_display_date(meta.get('player_as_of') or '2026-08-10')}일 기준"


def kbo_team_caption():
    meta = _read_kbo_meta()
    d = meta.get("team_as_of")
    return f"2026 정규시즌 · {_display_date(d)}일 기준" if d else "2026 정규시즌"


def validate_player_workbook(path):
    required = {"타자_기본","타자_세부","투수_기본","투수_세부"}
    xl = pd.ExcelFile(path, engine="openpyxl")
    missing = sorted(required - set(xl.sheet_names))
    if missing:
        raise ValueError("선수 기록 시트가 없습니다: " + ", ".join(missing))

    # 실제 로더까지 실행하여 형식 검증
    stat = Path(path).stat()
    h, p = load_kbo_records(str(path), stat.st_mtime_ns)
    if h.empty and p.empty:
        raise ValueError("선수 기록을 읽지 못했습니다.")
    return h, p


def validate_team_workbook(path):
    required = {"팀_타자","팀_투수","팀_수비","팀_주루"}
    xl = pd.ExcelFile(path, engine="openpyxl")
    missing = sorted(required - set(xl.sheet_names))
    if missing:
        raise ValueError("팀 기록 시트가 없습니다: " + ", ".join(missing))

    stat = Path(path).stat()
    data = load_kbo_team_records(str(path), stat.st_mtime_ns)
    for name in required:
        d = data.get(name, pd.DataFrame())
        if d.empty or "팀명" not in d.columns:
            raise ValueError(f"{name} 시트를 확인해 주세요.")
        teams = set(d["팀명"].astype(str).str.strip())
        if len(teams) < 10:
            raise ValueError(f"{name}: 10개 구단이 모두 들어 있지 않습니다.")
    return data


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

    CREATE TABLE IF NOT EXISTS favorite_teams(
        team_name TEXT PRIMARY KEY,
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
    return qdf("SELECT player_id, player_name FROM favorites ORDER BY player_name COLLATE NOCASE ASC")


def is_favorite_team(team_name):
    c = db()
    row = c.execute(
        "SELECT 1 FROM favorite_teams WHERE team_name=?",
        (str(team_name),)
    ).fetchone()
    c.close()
    return bool(row)


def toggle_favorite_team(team_name):
    c = db()
    row = c.execute(
        "SELECT 1 FROM favorite_teams WHERE team_name=?",
        (str(team_name),)
    ).fetchone()

    if row:
        c.execute("DELETE FROM favorite_teams WHERE team_name=?", (str(team_name),))
        action = "removed"
    else:
        c.execute(
            "INSERT OR REPLACE INTO favorite_teams(team_name,saved_at) VALUES(?,?)",
            (str(team_name), datetime.now().isoformat(timespec="seconds"))
        )
        action = "added"

    c.commit()
    c.close()
    snapshot_db()
    return action


def favorite_teams():
    return qdf("SELECT team_name FROM favorite_teams ORDER BY team_name COLLATE NOCASE ASC")



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


def _normalize_base_subset(raw):
    """Play-by-Play 원자료의 일부 행만 공통 분석 형식으로 변환."""
    if raw is None or raw.empty:
        return pd.DataFrame(), pd.DataFrame()

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
    pitch_type = pitch_type.where(
        pitch_type.notna() & (pitch_type.astype(str) != ""),
        raw["pitch_name"]
    )

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


def _player_filter_value(player_id):
    s = str(player_id).strip()
    if s.isdigit():
        try:
            return int(s)
        except Exception:
            return s
    return s


@st.cache_data(show_spinner=False, max_entries=8)
def load_base_player_subset(path_text, modified_ns, player_id):
    """선택한 선수의 행만 Parquet에서 읽음."""
    if not path_text:
        return pd.DataFrame(), pd.DataFrame()

    val = _player_filter_value(player_id)
    filters = [
        [("pitcher", "==", val)],
        [("batter", "==", val)],
    ]

    try:
        raw = pd.read_parquet(
            path_text,
            columns=BASE_COLUMNS,
            filters=filters
        )
    except Exception:
        # parquet 스키마가 문자열 ID인 경우 재시도
        sval = str(player_id)
        filters = [
            [("pitcher", "==", sval)],
            [("batter", "==", sval)],
        ]
        raw = pd.read_parquet(
            path_text,
            columns=BASE_COLUMNS,
            filters=filters
        )

    return _normalize_base_subset(raw)


@st.cache_data(show_spinner=False, max_entries=6)
def load_base_team_subset(path_text, modified_ns, team):
    """선택한 팀이 출전한 경기 행만 Parquet에서 읽음."""
    if not path_text:
        return pd.DataFrame(), pd.DataFrame()

    raw_codes = sorted({k for k,v in TEAM.items() if v == team} | {team})
    filters = []
    for code in raw_codes:
        filters.append([("home_team", "==", code)])
        filters.append([("away_team", "==", code)])

    raw = pd.read_parquet(
        path_text,
        columns=BASE_COLUMNS,
        filters=filters
    )
    return _normalize_base_subset(raw)


def base_player_data(player_id):
    path = get_base_parquet_path()
    if path is None:
        return pd.DataFrame(), pd.DataFrame()
    stat = path.stat()
    return load_base_player_subset(str(path), stat.st_mtime_ns, str(player_id))


def base_team_data(team):
    path = get_base_parquet_path()
    if path is None:
        return pd.DataFrame(), pd.DataFrame()
    stat = path.stat()
    return load_base_team_subset(str(path), stat.st_mtime_ns, str(team))

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

def kbo_record_year():
    """KBO 공식 기록 파일명에서 시즌 연도를 읽어 표시."""
    m = re.search(r"(20\d{2})", REPO_KBO_RECORDS.name)
    return m.group(1) if m else None


def show_center_loader():
    st.markdown(
        '<div class="center-loader"><div class="center-loader-ring"></div></div>',
        unsafe_allow_html=True
    )

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


def percentile_value(df, key, value, higher_is_better=True):
    """KBO 공식 기록 내 해당 지표의 백분위. 100에 가까울수록 우수."""
    if df is None or df.empty or key not in df.columns or value is None:
        return None
    try:
        v = float(value)
    except Exception:
        return None

    s = pd.to_numeric(df[key], errors="coerce").dropna()
    if s.empty:
        return None

    if higher_is_better:
        pct = (s <= v).mean() * 100
    else:
        pct = (s >= v).mean() * 100
    return round(float(pct), 1)


def action_rate(df, target_actions):
    if df is None or df.empty or "pitch_text" not in df.columns:
        return None
    acts = df["pitch_text"].apply(action)
    if len(acts) == 0:
        return None
    return round(float(acts.isin(target_actions).mean() * 100), 1)


def first_pitch_rate(df, target_actions):
    if df is None or df.empty or "pitch_num" not in df.columns:
        return None
    first = df[pd.to_numeric(df["pitch_num"], errors="coerce") == 1].copy()
    if first.empty:
        return None
    acts = first["pitch_text"].apply(action)
    return round(float(acts.isin(target_actions).mean() * 100), 1)


def speed_delta(latest_df, base_df):
    """NAVER 최신 자료 평균구속 - Play-by-Play 기본자료 평균구속."""
    if latest_df is None or latest_df.empty or base_df is None or base_df.empty:
        return None
    if "speed" not in latest_df.columns or "speed" not in base_df.columns:
        return None
    latest = pd.to_numeric(latest_df["speed"], errors="coerce").dropna()
    base = pd.to_numeric(base_df["speed"], errors="coerce").dropna()
    if latest.empty or base.empty:
        return None
    return round(float(latest.mean() - base.mean()), 1)


def pct_text(v):
    return f"{v:.1f}%" if v is not None else "-"


def delta_text(v):
    if v is None:
        return "-"
    return f"{v:+.1f} km/h"


def data_source_status():
    base_p, base_g = base_data()
    _, naver_g = source_games()
    kbo_h, kbo_p = kbo_records()
    kbo_team = kbo_team_records()
    team_ready = any(
        isinstance(df, pd.DataFrame) and not df.empty
        for df in kbo_team.values()
    )
    return {
        "kbo": (not kbo_h.empty) or (not kbo_p.empty),
        "kbo_team": team_ready,
        "naver_games": int(naver_g["game_id"].nunique()) if not naver_g.empty else 0,
        "base_games": int(base_g["game_id"].nunique()) if not base_g.empty else 0,
    }



def _base_file_signature():
    path = get_base_parquet_path()
    if path is None:
        return None, 0
    try:
        stat = path.stat()
        return str(path), stat.st_mtime_ns
    except Exception:
        return str(path), 0


@st.cache_data(show_spinner=False)
def load_home_base_summary(path_text, modified_ns):
    """홈 화면에 필요한 최소 정보만 Parquet에서 읽습니다."""
    if not path_text:
        return {
            "game_ids": set(),
            "recent_games": pd.DataFrame(columns=["game_date","away_team","home_team"]),
            "players": pd.DataFrame(columns=["player_id","player_name","team"]),
        }

    path = Path(path_text)
    if not path.exists():
        return {
            "game_ids": set(),
            "recent_games": pd.DataFrame(columns=["game_date","away_team","home_team"]),
            "players": pd.DataFrame(columns=["player_id","player_name","team"]),
        }

    cols = [
        "game_pk","game_date","home_team","away_team","inning_topbot",
        "pitcher","pitcher_name","batter","batter_name"
    ]
    raw = pd.read_parquet(path, columns=cols)

    home = raw["home_team"].map(TEAM).fillna(raw["home_team"])
    away = raw["away_team"].map(TEAM).fillna(raw["away_team"])
    bottom = raw["inning_topbot"].astype(str).str.lower().isin(["bot","bottom","말"])

    defense = away.where(bottom, home)
    offense = home.where(bottom, away)

    # 경기
    games = pd.DataFrame({
        "game_id": raw["game_pk"].astype(str),
        "game_date": raw["game_date"].astype(str).str[:10],
        "away_team": away,
        "home_team": home,
    }).drop_duplicates("game_id")

    # 선수는 필요한 3개 열만 남기고 즉시 중복 제거
    pitchers = pd.DataFrame({
        "player_id": raw["pitcher"].fillna("").astype(str).str.strip(),
        "player_name": raw["pitcher_name"].fillna("").astype(str).str.strip(),
        "team": defense.fillna("").astype(str).str.strip(),
        "role": "투수",
    })
    batters = pd.DataFrame({
        "player_id": raw["batter"].fillna("").astype(str).str.strip(),
        "player_name": raw["batter_name"].fillna("").astype(str).str.strip(),
        "team": offense.fillna("").astype(str).str.strip(),
        "role": "타자",
    })

    players = pd.concat([pitchers, batters], ignore_index=True)
    players = players[
        (players["player_name"] != "") &
        (players["player_name"].str.lower() != "nan")
    ].drop_duplicates(["player_id","player_name","team","role"])

    return {
        "game_ids": set(games["game_id"].astype(str)),
        "recent_games": games[["game_date","away_team","home_team"]],
        "players": players,
    }



def player_master_light():
    """선수 검색 목록 = Play-by-Play + NAVER + KBO 공식 선수 기록 합집합."""
    path_text, modified_ns = _base_file_signature()
    base = load_home_base_summary(path_text, modified_ns)
    players = base["players"].copy()

    # NAVER 문자중계/저장 DB 선수
    nav = qdf("""
        SELECT pitcher_id AS player_id, pitcher_name AS player_name,
               defense_team AS team, '투수' AS role
        FROM pitches
        UNION ALL
        SELECT batter_id AS player_id, batter_name AS player_name,
               offense_team AS team, '타자' AS role
        FROM pitches
    """)

    if nav is not None and not nav.empty:
        nav["player_id"] = nav["player_id"].fillna("").astype(str).str.strip()
        nav["player_name"] = nav["player_name"].fillna("").astype(str).str.strip()
        nav["team"] = nav["team"].fillna("").astype(str).str.strip()
        nav["role"] = nav["role"].fillna("").astype(str).str.strip()
        nav = nav[
            (nav["player_name"] != "") &
            (nav["player_name"].str.lower() != "nan")
        ].drop_duplicates(["player_id","player_name","team","role"])
        players = pd.concat([players, nav], ignore_index=True)

    # KBO 공식 선수 기록도 검색 목록에 포함
    kbo_hitters, kbo_pitchers = kbo_records()
    kbo_parts = []

    if kbo_hitters is not None and not kbo_hitters.empty and "선수명" in kbo_hitters.columns:
        kh = pd.DataFrame({
            "player_id": "",
            "player_name": kbo_hitters["선수명"].fillna("").astype(str).str.strip(),
            "team": kbo_hitters["팀명"].fillna("").astype(str).str.strip()
                    if "팀명" in kbo_hitters.columns else "",
            "role": "타자",
        })
        kbo_parts.append(kh)

    if kbo_pitchers is not None and not kbo_pitchers.empty and "선수명" in kbo_pitchers.columns:
        kp = pd.DataFrame({
            "player_id": "",
            "player_name": kbo_pitchers["선수명"].fillna("").astype(str).str.strip(),
            "team": kbo_pitchers["팀명"].fillna("").astype(str).str.strip()
                    if "팀명" in kbo_pitchers.columns else "",
            "role": "투수",
        })
        kbo_parts.append(kp)

    if kbo_parts:
        kbo_players = pd.concat(kbo_parts, ignore_index=True)
        kbo_players = kbo_players[
            (kbo_players["player_name"] != "") &
            (kbo_players["player_name"].str.lower() != "nan")
        ].drop_duplicates(["player_name","team","role"])

        # 같은 이름+팀이 PBP/NAVER에 있으면 그 ID를 KBO 행에도 연결
        if not players.empty:
            id_map = (
                players.assign(
                    player_name=players["player_name"].fillna("").astype(str).str.strip(),
                    team=players["team"].fillna("").astype(str).str.strip(),
                    player_id=players["player_id"].fillna("").astype(str).str.strip(),
                )
                .query("player_id != ''")
                .drop_duplicates(["player_name","team"])
                .set_index(["player_name","team"])["player_id"]
                .to_dict()
            )
            kbo_players["player_id"] = [
                id_map.get((n, t), "")
                for n, t in zip(kbo_players["player_name"], kbo_players["team"])
            ]

        players = pd.concat([players, kbo_players], ignore_index=True)

    if players.empty:
        return pd.DataFrame(columns=["id","name","team","role"])

    players = players.rename(columns={
        "player_id":"id", "player_name":"name"
    })
    players["id"] = players["id"].fillna("").astype(str).str.strip()
    players["name"] = players["name"].fillna("").astype(str).str.strip()
    players["team"] = players["team"].fillna("").astype(str).str.strip()
    players["role"] = players["role"].fillna("").astype(str).str.strip()
    players = players[
        (players["name"] != "") &
        (players["name"].str.lower() != "nan")
    ]

    # KBO에만 있는 선수는 ID가 없으므로 이름+팀 기반의 안정적인 내부 검색키 생성
    missing_id = (players["id"] == "") | (players["id"].str.lower() == "nan")
    players.loc[missing_id, "id"] = (
        "KBO:" +
        players.loc[missing_id, "team"].astype(str) + ":" +
        players.loc[missing_id, "name"].astype(str)
    )

    players = players.drop_duplicates(["id","name","team","role"])

    rows = []
    for (pid, name), g in players.groupby(["id","name"], sort=False):
        roles = set(g["role"].dropna())
        role_label = "투타" if roles == {"타자","투수"} else ("투수" if "투수" in roles else "타자")
        teams = [t for t in g["team"].tolist() if t]
        rows.append({
            "id": str(pid),
            "name": name,
            "team": teams[-1] if teams else "",
            "role": role_label,
        })

    result = pd.DataFrame(rows)

    # 같은 선수(이름+팀)가 KBO 내부키와 실제 PBP/NAVER ID로 중복된 경우 실제 ID 우선
    if not result.empty:
        result["_real_id"] = ~result["id"].astype(str).str.startswith("KBO:")
        result = (
            result.sort_values(["name","team","_real_id"], ascending=[True, True, False])
                  .drop_duplicates(["name","team"], keep="first")
                  .drop(columns="_real_id")
        )

    return result.sort_values(["team","role","name"]).reset_index(drop=True)

def _dedup_extra_against_base(extra, base_games):
    if extra is None or extra.empty:
        return pd.DataFrame()
    if base_games is None or base_games.empty:
        return extra.copy()
    base_ids = set(base_games["game_id"].astype(str))
    return extra[~extra["game_id"].astype(str).isin(base_ids)].copy()

def overview_counts():
    """홈 화면용 경량 집계.

    전체 Play-by-Play DataFrame을 만들지 않고 필요한 열만 읽어서
    경기 수와 등록 선수 수를 계산합니다.
    """
    path_text, modified_ns = _base_file_signature()
    base = load_home_base_summary(path_text, modified_ns)

    # NAVER 추가 경기/선수는 SQLite에서 필요한 열만 읽음
    nav_games = qdf("SELECT game_id, game_date, away_team, home_team FROM games")
    nav_players = qdf("""
        SELECT pitcher_id AS player_id, pitcher_name AS player_name, defense_team AS team
        FROM pitches
        UNION ALL
        SELECT batter_id AS player_id, batter_name AS player_name, offense_team AS team
        FROM pitches
    """)

    base_game_ids = base["game_ids"]

    if nav_games is not None and not nav_games.empty:
        nav_games = nav_games[
            ~nav_games["game_id"].astype(str).isin(base_game_ids)
        ].copy()
        game_count = len(base_game_ids) + int(nav_games["game_id"].astype(str).nunique())
    else:
        game_count = len(base_game_ids)

    players = base["players"].copy()

    if nav_players is not None and not nav_players.empty:
        nav_players["player_id"] = nav_players["player_id"].fillna("").astype(str).str.strip()
        nav_players["player_name"] = nav_players["player_name"].fillna("").astype(str).str.strip()
        nav_players["team"] = nav_players["team"].fillna("").astype(str).str.strip()
        nav_players = nav_players[
            (nav_players["player_name"] != "") &
            (nav_players["player_name"].str.lower() != "nan")
        ].drop_duplicates(["player_id","player_name","team"])
        players = pd.concat([players, nav_players], ignore_index=True)

    players["player_id"] = players["player_id"].fillna("").astype(str).str.strip()
    players["player_name"] = players["player_name"].fillna("").astype(str).str.strip()
    players["team"] = players["team"].fillna("").astype(str).str.strip()

    # ID가 있는 선수는 ID 기준. ID 없는 경우 이름+팀 기준.
    valid_id = (players["player_id"] != "") & (players["player_id"].str.lower() != "nan")
    id_keys = set(("id:" + players.loc[valid_id, "player_id"]).tolist())
    fallback_keys = set(
        ("name_team:" + players.loc[~valid_id, "player_name"] + "|" + players.loc[~valid_id, "team"]).tolist()
    )
    player_keys = id_keys | fallback_keys

    # KBO 공식 기록에만 존재하는 선수가 있는 경우만 추가
    kbo_h, kbo_p = kbo_records()
    name_team_pairs = set(zip(players["player_name"], players["team"]))
    name_counts = players.groupby("player_name")["player_id"].nunique(dropna=False).to_dict()

    kbo_only = set()
    for df in (kbo_h, kbo_p):
        if df is None or df.empty or "선수명" not in df.columns:
            continue
        names = df["선수명"].fillna("").astype(str).str.strip()
        teams = (
            df["팀명"].fillna("").astype(str).str.strip()
            if "팀명" in df.columns
            else pd.Series([""] * len(df), index=df.index)
        )
        for name, team in zip(names, teams):
            if not name or name.lower() == "nan":
                continue
            if (name, team) in name_team_pairs:
                continue
            if name_counts.get(name, 0) == 1:
                continue
            kbo_only.add(f"kbo:{name}|{team}")

    return {
        "games": int(game_count),
        "players": len(player_keys | kbo_only),
    }


def recent_games_light(limit=5):
    """홈의 최근 경기 표 전용. 전체 투구 데이터는 읽지 않습니다."""
    path_text, modified_ns = _base_file_signature()
    base = load_home_base_summary(path_text, modified_ns)
    base_games = base["recent_games"].copy()

    nav = qdf("SELECT game_id, game_date, away_team, home_team FROM games")
    if nav is not None and not nav.empty:
        if base["game_ids"]:
            nav = nav[~nav["game_id"].astype(str).isin(base["game_ids"])].copy()
        nav = nav[["game_date","away_team","home_team"]]
        games = pd.concat([base_games, nav], ignore_index=True)
    else:
        games = base_games

    if games.empty:
        return games

    games["game_date"] = games["game_date"].astype(str)
    return (
        games.drop_duplicates()
        .sort_values("game_date", ascending=False)
        .head(limit)
        .reset_index(drop=True)
    )


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


st.set_page_config(page_title="찐팬의 데이터 덕아웃", page_icon="⚾", layout="wide")
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

/* 중앙 로딩 표시 */
.center-loader{
    position:fixed;
    inset:0;
    display:flex;
    align-items:center;
    justify-content:center;
    background:rgba(255,255,255,.72);
    z-index:99999;
}
.center-loader-ring{
    width:34px;
    height:34px;
    border:4px solid #d9dde3;
    border-top-color:#555b66;
    border-radius:50%;
    animation:dugout-spin .8s linear infinite;
}
@keyframes dugout-spin{
    to{transform:rotate(360deg)}
}
.dugout-analysis-head{margin-top:1rem;margin-bottom:.35rem;padding:20px 22px;border:2px solid #1f2937;border-radius:18px;background:linear-gradient(135deg,#fafafa 0%,#f1f5f9 100%)}
.dugout-analysis-kicker{font-size:.72rem;font-weight:800;letter-spacing:.14em;color:#64748b;margin-bottom:.25rem}
.dugout-analysis-title{font-size:1.55rem;font-weight:850;color:#111827;letter-spacing:-.035em}
.evidence-label{margin-top:2rem;padding-top:1rem;border-top:1px solid #e5e7eb;font-size:.82rem;font-weight:800;color:#64748b;letter-spacing:.08em}
.home-intro{padding:22px 24px;border:1px solid #e5e7eb;border-radius:18px;background:#fafafa;margin:.6rem 0 1.2rem 0}
.home-intro-title{font-size:1.2rem;font-weight:800;margin-bottom:.35rem}
.home-intro-text{color:#5f6368;line-height:1.55}

/* 홈 관심 선수 버튼을 작고 촘촘하게 표시 */
div[data-testid="stButton"] > button {
    border-radius:999px;
}

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
    st.markdown(
        '''
        <div class="brand">⚾ 찐팬의 데이터 덕아웃</div>
        <div style="font-size:17px;font-weight:500;color:#7a7f8c;margin-top:2px;">
            나만의 KBO 스카우팅 리포트
        </div>
        ''',
        unsafe_allow_html=True
    )

with top2:
    st.session_state.presentation_mode = st.toggle("발표 모드", value=st.session_state.presentation_mode,
                                                   help="저장된 데이터만 사용합니다. NAVER에 새 요청을 보내지 않습니다.")

if st.session_state.presentation_mode:
    st.caption("발표 모드 · 저장된 데이터만 사용 중")

nav_items = ["홈","팀","선수"] if st.session_state.presentation_mode else ["홈","팀","선수","데이터"]
if "main_nav" not in st.session_state or st.session_state.main_nav not in nav_items:
    st.session_state.main_nav = "홈"
nav = st.radio("메인 메뉴", nav_items, horizontal=True, label_visibility="collapsed", key="main_nav")

def go_to(page, player_name=None, team_name=None, player_id=None):
    st.session_state.main_nav = page
    if player_name is not None:
        st.session_state.player_search = player_name
    if team_name is not None:
        st.session_state.team_select = team_name
    if player_id is not None:
        st.session_state.selected_player_id = str(player_id)


gc.collect()

if nav == "홈":

    loader = st.empty()
    with loader.container():
        show_center_loader()

    summary = overview_counts()
    favs = favorite_players()
    fav_teams = favorite_teams()
    games = recent_games_light(5)

    loader.empty()

    st.markdown("## DATA DUGOUT 현황")
    h1,h2 = st.columns(2)
    h1.metric("분석 경기", f"{summary['games']:,}")
    h2.metric("등록 선수", f"{summary['players']:,}")

    st.markdown("### ⭐ 관심 선수")
    if favs.empty:
        st.caption("등록된 관심 선수가 없습니다. 선수 페이지에서 ☆ 관심 선수 등록을 눌러 추가할 수 있습니다.")
    else:
        # 가나다순으로 정렬된 관심 선수를 작고 촘촘한 버튼으로 표시
        fav_cols = st.columns(6)
        for pos, (_, row) in enumerate(favs.iterrows()):
            with fav_cols[pos % 6]:
                st.button(
                    f"★ {row['player_name']}",
                    key=f"home_fav_{row['player_id']}",
                    use_container_width=False,
                    on_click=go_to,
                    args=("선수", row["player_name"], None, row["player_id"])
                )

    st.markdown("### ⚾ 관심 팀")
    if fav_teams.empty:
        st.caption("등록된 관심 팀이 없습니다. 팀 페이지에서 ☆ 관심 팀 등록을 눌러 추가할 수 있습니다.")
    else:
        team_cols = st.columns(6)
        for pos, (_, row) in enumerate(fav_teams.iterrows()):
            with team_cols[pos % 6]:
                st.button(
                    f"★ {row['team_name']}",
                    key=f"home_fav_team_{row['team_name']}",
                    use_container_width=False,
                    on_click=go_to,
                    args=("팀", None, row["team_name"])
                )

    st.markdown("### 최근 경기")
    if games.empty:
        st.caption("표시할 경기 자료가 없습니다.")
    else:
        recent = games[["game_date","away_team","home_team"]].copy()
        recent.columns = ["날짜","원정","홈"]
        st.dataframe(
            recent,
            hide_index=True,
            use_container_width=True,
            height=min(250, 38 * (len(recent) + 1))
        )

elif nav == "데이터":
    st.markdown("## 데이터")

    status = data_source_status()
    st.markdown("### 데이터 연결 상태")
    meta = _read_kbo_meta()
    s1, s2, s3, s4 = st.columns(4)

    player_date = (
        f"{_display_date(meta.get('player_as_of') or '2026-08-10')} 기준"
        if status["kbo"]
        else "자료 없음"
    )

    team_date = (
        f"{_display_date(meta.get('team_as_of'))} 기준"
        if status["kbo_team"] and meta.get("team_as_of")
        else ("자료 있음" if status["kbo_team"] else "자료 없음")
    )

    with s1:
        st.markdown(
            f"""
<div style="border:1px solid #e5e7eb; border-radius:18px; padding:22px 24px; min-height:125px; background:white;">
    <div style="font-size:24px; font-weight:700; color:#30323d; margin-bottom:14px;">
        KBO 선수 기록
    </div>
    <div style="font-size:18px; font-weight:400; color:#6b7280;">
        {player_date}
    </div>
</div>
""",
            unsafe_allow_html=True,
        )

    with s2:
        st.markdown(
            f"""
<div style="border:1px solid #e5e7eb; border-radius:18px; padding:22px 24px; min-height:125px; background:white;">
    <div style="font-size:24px; font-weight:700; color:#30323d; margin-bottom:14px;">
        KBO 팀 기록
    </div>
    <div style="font-size:18px; font-weight:400; color:#6b7280;">
        {team_date}
    </div>
</div>
""",
            unsafe_allow_html=True,
        )

    with s3:
        st.markdown(
            f"""
<div style="border:1px solid #e5e7eb; border-radius:18px; padding:22px 24px; min-height:125px; background:white;">
    <div style="font-size:24px; font-weight:700; color:#30323d; margin-bottom:14px;">
        NAVER 최신 자료
    </div>
    <div style="font-size:18px; font-weight:400; color:#6b7280;">
        {status['naver_games']}경기
    </div>
</div>
""",
            unsafe_allow_html=True,
        )

    with s4:
        st.markdown(
            f"""
<div style="border:1px solid #e5e7eb; border-radius:18px; padding:22px 24px; min-height:125px; background:white;">
    <div style="font-size:24px; font-weight:700; color:#30323d; margin-bottom:14px;">
        Play-by-Play 자료
    </div>
    <div style="font-size:18px; font-weight:400; color:#6b7280;">
        {status['base_games']}경기
    </div>
</div>
""",
            unsafe_allow_html=True,
        )

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
            st.rerun()

    st.divider()
    st.markdown("### ② KBO 공식 기록 업데이트")
    st.caption("KBO 선수 기록과 팀 기록은 NAVER 문자중계와 별도로 저장됩니다.")

    k1, k2 = st.columns(2)

    with k1:
        st.markdown("#### KBO 선수 기록 업데이트")
        st.caption(f"현재 기준: {_display_date(_read_kbo_meta().get('player_as_of') or '2026-08-10')}")
        player_as_of = st.date_input(
            "선수 기록 기준일",
            value=pd.Timestamp(_read_kbo_meta().get("player_as_of") or "2026-08-10").date(),
            key="kbo_player_as_of"
        )
        player_upload = st.file_uploader(
            "KBO 선수 기록 Excel",
            type=["xlsx"],
            key="kbo_player_records_uploader",
            help="타자_기본 / 타자_세부 / 투수_기본 / 투수_세부 시트가 필요합니다."
        )
        if st.button("KBO 선수 기록 업데이트", use_container_width=True):
            if player_upload is None:
                st.warning("선수 기록 Excel 파일을 선택해 주세요.")
            else:
                try:
                    temp = DATA_DIR / "kbo_player_records.uploading.xlsx"
                    with open(temp, "wb") as f:
                        f.write(player_upload.getbuffer())
                    validate_player_workbook(temp)
                    temp.replace(UPLOADED_KBO_PLAYER_RECORDS)
                    load_kbo_records.clear()
                    _write_kbo_meta(
                        player_as_of=str(player_as_of),
                        player_uploaded_at=datetime.now().isoformat(timespec="seconds")
                    )
                    st.success(f"KBO 선수 기록을 {_display_date(str(player_as_of))} 기준으로 업데이트했습니다.")
                    st.rerun()
                except Exception as e:
                    st.error(f"KBO 선수 기록을 업데이트하지 못했습니다: {e}")

    with k2:
        st.markdown("#### KBO 팀 기록 업데이트")
        team_meta_date = _read_kbo_meta().get("team_as_of")
        default_team_date = pd.Timestamp(team_meta_date or "2026-08-14").date()
        st.caption(
            f"현재 기준: {_display_date(team_meta_date)}"
            if team_meta_date else
            "현재 기준: 팀 기록 파일 미등록"
        )
        team_as_of = st.date_input(
            "팀 기록 기준일",
            value=default_team_date,
            key="kbo_team_as_of"
        )
        team_upload = st.file_uploader(
            "KBO 팀 기록 Excel",
            type=["xlsx"],
            key="kbo_team_records_uploader",
            help="팀_타자 / 팀_투수 / 팀_수비 / 팀_주루 시트가 필요합니다."
        )
        if st.button("KBO 팀 기록 업데이트", use_container_width=True):
            if team_upload is None:
                st.warning("팀 기록 Excel 파일을 선택해 주세요.")
            else:
                try:
                    temp = DATA_DIR / "kbo_team_records.uploading.xlsx"
                    with open(temp, "wb") as f:
                        f.write(team_upload.getbuffer())
                    validate_team_workbook(temp)
                    temp.replace(UPLOADED_KBO_TEAM_RECORDS)
                    load_kbo_team_records.clear()
                    _write_kbo_meta(
                        team_as_of=str(team_as_of),
                        team_uploaded_at=datetime.now().isoformat(timespec="seconds")
                    )
                    st.success(f"KBO 팀 기록을 {_display_date(str(team_as_of))} 기준으로 업데이트했습니다.")
                    st.rerun()
                except Exception as e:
                    st.error(f"KBO 팀 기록을 업데이트하지 못했습니다: {e}")

    st.divider()
    st.markdown("### ③ Play-by-Play 자료 관리")
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
            temp_path.replace(UPLOADED_BASE_PARQUET); base_p,base_g=base_data()
            if base_p.empty: st.error("파일을 읽지 못했습니다.")
            else:
                st.success("Play-by-Play 자료 업데이트가 완료되었습니다.")
                st.caption(f"자료 기간: {analysis_period(base_g) or '-'}")
        except Exception as e: st.error(f"Parquet 파일을 읽지 못했습니다: {e}")

elif nav == "팀":
    st.markdown("## 팀")

    # 팀 목록은 고정 10개 구단. 전체 시즌 Parquet를 먼저 열 필요가 없습니다.
    options = ["KT","LG","삼성","두산","KIA","NC","SSG","롯데","키움","한화"]
    priority=[t for t in PRIORITY_TEAMS if t in options]
    options=priority+[t for t in options if t not in priority]

    if "team_select" not in st.session_state or st.session_state.team_select not in options:
        st.session_state.team_select = options[0]

    team=st.selectbox("팀 선택", options, key="team_select")

    team_fav_now = is_favorite_team(team)
    if st.button(
        "★ 관심 팀 해제" if team_fav_now else "☆ 관심 팀 등록",
        key=f"fav_team_{team}"
    ):
        toggle_favorite_team(team)
        st.rerun()

    loader = st.empty()
    with loader.container():
        show_center_loader()

    # 선택한 팀 경기만 Parquet에서 읽음
    bp_team_p, bp_team_games = base_team_data(team)

    naver_p = extra_pitches()
    naver_g = extra_games()

    # 중복 NAVER 경기 제거에는 가벼운 base game id 목록만 사용
    path_text, modified_ns = _base_file_signature()
    base_summary = load_home_base_summary(path_text, modified_ns)
    base_ids = base_summary["game_ids"]

    if not naver_p.empty:
        naver_p = naver_p[~naver_p["game_id"].astype(str).isin(base_ids)].copy()
    if not naver_g.empty:
        naver_g = naver_g[~naver_g["game_id"].astype(str).isin(base_ids)].copy()

    nv_team_games = naver_g[
        (naver_g["away_team"]==team)|(naver_g["home_team"]==team)
    ].copy() if not naver_g.empty else pd.DataFrame()

    nv_thrown = naver_p[naver_p["defense_team"]==team].copy() if not naver_p.empty else pd.DataFrame()
    nv_seen = naver_p[naver_p["offense_team"]==team].copy() if not naver_p.empty else pd.DataFrame()

    bp_thrown = bp_team_p[bp_team_p["defense_team"]==team].copy() if not bp_team_p.empty else pd.DataFrame()
    bp_seen = bp_team_p[bp_team_p["offense_team"]==team].copy() if not bp_team_p.empty else pd.DataFrame()

    thrown = pd.concat([bp_thrown,nv_thrown],ignore_index=True)
    seen = pd.concat([bp_seen,nv_seen],ignore_index=True)
    team_games = pd.concat([bp_team_games,nv_team_games],ignore_index=True)

    loader.empty()

    analysis_header(f"{team} 팀 분석")
    with st.container(border=True):
        top_throw,_,top_throw_share=top_pitch_info(thrown)
        top_seen,_,top_seen_share=top_pitch_info(seen)
        # KBO 공식 팀 경기 수를 DATA DUGOUT 대표 경기 수로 사용합니다.
        # Play-by-Play/NAVER 경기 수는 세부 분석 커버리지로 별도 표시합니다.
        team_official = kbo_team_records()
        _th = team_official.get("팀_타자", pd.DataFrame())
        _tp = team_official.get("팀_투수", pd.DataFrame())
        _td = team_official.get("팀_수비", pd.DataFrame())
        _tr = team_official.get("팀_주루", pd.DataFrame())

        def _official_row(df0):
            if df0 is None or df0.empty or "팀명" not in df0.columns:
                return None
            q = df0[df0["팀명"].astype(str).str.strip() == team]
            return q.iloc[0] if not q.empty else None

        _hr = _official_row(_th)
        _pr = _official_row(_tp)
        _dr = _official_row(_td)
        _rr = _official_row(_tr)

        official_games = (
            metric_value(_hr, "G")
            if _hr is not None
            else (metric_value(_pr, "G") if _pr is not None else "-")
        )
        analysis_games = team_games["game_id"].nunique() if not team_games.empty else 0

        a,b,c,d=st.columns(4)
        a.metric("공식 경기", official_games)
        b.metric("투수진 투구",len(thrown))
        c.metric("주 사용 구종",top_throw)
        d.metric("주 사용 비율",f"{top_throw_share:.1f}%" if top_throw!="-" else "-")
        a,b,c,d=st.columns(4)
        a.metric("타선이 본 투구",len(seen))
        b.metric("가장 많이 본 구종",top_seen)
        c.metric("상대 구종 비율",f"{top_seen_share:.1f}%" if top_seen!="-" else "-")
        d.metric("투수진 평균 구속",avg_speed_value(thrown))

        # KBO 공식 팀 기록을 단순 반복하지 않고 리그 내 상대 위치로 재구성
        if any(x is not None for x in [_hr,_pr,_dr,_rr]):
            def _num(v):
                try:
                    return float(v)
                except Exception:
                    return None

            def _rank_of(df0, team_name, col, higher_is_better=True):
                """10개 구단 순위를 계산하되 같은 값은 공동 순위로 처리."""
                if df0 is None or df0.empty or "팀명" not in df0.columns or col not in df0.columns:
                    return "-"

                d = df0[["팀명", col]].copy()
                d["팀명"] = d["팀명"].fillna("").astype(str).str.strip()
                d[col] = pd.to_numeric(d[col], errors="coerce")
                d = d[(d["팀명"] != "") & d[col].notna()].copy()

                if d.empty:
                    return "-"

                # pandas rank(method="min"):
                # 같은 값은 같은 순위를 부여하고, 다음 순위는 건너뜀.
                # 예: 1위, 공동 2위, 공동 2위, 4위
                d["_rank"] = d[col].rank(
                    method="min",
                    ascending=not higher_is_better
                ).astype(int)

                row = d[d["팀명"] == str(team_name).strip()]
                if row.empty:
                    return "-"

                return int(row.iloc[0]["_rank"])

            def _per_game(row, num_col):
                if row is None:
                    return None
                g = _num(row.get("G"))
                n = _num(row.get(num_col))
                if not g or n is None:
                    return None
                return n / g

            avg_rank = _rank_of(_th, team, "AVG", True)
            runs_rank = _rank_of(_th.assign(
                R_G=pd.to_numeric(_th.get("R"), errors="coerce") / pd.to_numeric(_th.get("G"), errors="coerce")
            ) if not _th.empty else _th, team, "R_G", True)
            hr_rank = _rank_of(_th.assign(
                HR_G=pd.to_numeric(_th.get("HR"), errors="coerce") / pd.to_numeric(_th.get("G"), errors="coerce")
            ) if not _th.empty else _th, team, "HR_G", True)

            era_rank = _rank_of(_tp, team, "ERA", False)
            whip_rank = _rank_of(_tp, team, "WHIP", False)
            fpct_rank = _rank_of(_td, team, "FPCT", True)
            sbpct_rank = _rank_of(_tr, team, "SB%", True)

            runs_pg = _per_game(_hr, "R")
            hr_pg = _per_game(_hr, "HR")

            def _rank_label(df0, team_name, col, rank_value):
                """같은 기록의 팀이 2개 이상이면 '공동 n위'로 표시."""
                if rank_value == "-" or df0 is None or df0.empty or col not in df0.columns or "팀명" not in df0.columns:
                    return None

                d = df0[["팀명", col]].copy()
                d["팀명"] = d["팀명"].fillna("").astype(str).str.strip()
                d[col] = pd.to_numeric(d[col], errors="coerce")
                d = d[(d["팀명"] != "") & d[col].notna()].copy()

                row = d[d["팀명"] == str(team_name).strip()]
                if row.empty:
                    return f"리그 {rank_value}위"

                value = row.iloc[0][col]
                tied = int((d[col] == value).sum()) > 1
                return f"리그 공동 {rank_value}위" if tied else f"리그 {rank_value}위"

            # 경기당 지표는 임시 계산 컬럼을 포함한 데이터프레임을 다시 준비
            _th_runs = _th.assign(
                R_G=pd.to_numeric(_th.get("R"), errors="coerce") / pd.to_numeric(_th.get("G"), errors="coerce")
            ) if not _th.empty else _th
            _th_hr = _th.assign(
                HR_G=pd.to_numeric(_th.get("HR"), errors="coerce") / pd.to_numeric(_th.get("G"), errors="coerce")
            ) if not _th.empty else _th

            avg_rank_label = _rank_label(_th, team, "AVG", avg_rank)
            runs_rank_label = _rank_label(_th_runs, team, "R_G", runs_rank)
            hr_rank_label = _rank_label(_th_hr, team, "HR_G", hr_rank)
            era_rank_label = _rank_label(_tp, team, "ERA", era_rank)
            whip_rank_label = _rank_label(_tp, team, "WHIP", whip_rank)
            fpct_rank_label = _rank_label(_td, team, "FPCT", fpct_rank)
            sbpct_rank_label = _rank_label(_tr, team, "SB%", sbpct_rank)

            st.markdown("#### DATA DUGOUT 팀 프로필")
            p1,p2,p3,p4 = st.columns(4)
            p1.metric(
                "공격 정확도",
                f"타율 {kbo_avg_value(_hr, 'AVG')}" if _hr is not None else "-",
                avg_rank_label
            )
            p2.metric(
                "득점 생산",
                f"{runs_pg:.2f}점/경기" if runs_pg is not None else "-",
                runs_rank_label
            )
            p3.metric(
                "장타 생산",
                f"{hr_pg:.2f}HR/경기" if hr_pg is not None else "-",
                hr_rank_label
            )
            p4.metric(
                "실점 억제",
                f"ERA {metric_value(_pr, 'ERA')}" if _pr is not None else "-",
                era_rank_label
            )

            p1,p2,p3,p4 = st.columns(4)
            p1.metric(
                "출루 허용",
                f"WHIP {metric_value(_pr, 'WHIP')}" if _pr is not None else "-",
                whip_rank_label
            )
            p2.metric(
                "수비 안정성",
                f"FPCT {metric_value(_dr, 'FPCT')}" if _dr is not None else "-",
                fpct_rank_label
            )
            _sb = metric_value(_rr, "SB%") if _rr is not None else "-"
            p3.metric(
                "주루 성공률",
                f"{_sb}%" if _sb != "-" else "-",
                sbpct_rank_label
            )
            p4.metric(
                "투구 성향",
                top_throw,
                f"{top_throw_share:.1f}% · 평균 {avg_speed_value(thrown)}" if top_throw != "-" else None
            )

            # 원자료를 조합한 짧은 해석: 사실 기반 순위/비율만 사용
            insights = []
            if avg_rank != "-":
                insights.append(f"팀 타율은 {avg_rank_label}")
            if era_rank != "-":
                insights.append(f"팀 ERA는 {era_rank_label}")
            if fpct_rank != "-":
                insights.append(f"수비율은 {fpct_rank_label}")
            if sbpct_rank != "-":
                insights.append(f"도루 성공률은 {sbpct_rank_label}")
            if top_throw != "-":
                insights.append(f"투수진은 {top_throw}을 가장 많이 사용({top_throw_share:.1f}%)")
            if top_seen != "-":
                insights.append(f"타선은 상대 {top_seen}을 가장 많이 상대({top_seen_share:.1f}%)")

            if insights:
                st.markdown("#### 핵심 해석")
                st.write(" · ".join(insights))

        st.markdown("#### 투수진 구종 구성")
        if not thrown.empty:
            mix=thrown.groupby("pitch_type",dropna=True).agg(
                투구수=("pitch_id","count"),평균구속=("speed","mean")
            ).reset_index()
            mix["평균구속"]=mix["평균구속"].round(1)
            st.bar_chart(mix.set_index("pitch_type")["투구수"])
            st.dataframe(mix.sort_values("투구수",ascending=False),hide_index=True,use_container_width=True)

        st.markdown("#### 타선이 상대해 온 구종")
        if not seen.empty:
            mix2=seen.groupby("pitch_type",dropna=True).size().reset_index(name="투구수")
            st.bar_chart(mix2.set_index("pitch_type")["투구수"])
            st.dataframe(mix2.sort_values("투구수",ascending=False),hide_index=True,use_container_width=True)

    evidence_header()

    st.markdown("### KBO 공식 팀 기록")
    team_records = kbo_team_records()
    t_hit = team_records.get("팀_타자", pd.DataFrame())
    t_pit = team_records.get("팀_투수", pd.DataFrame())
    t_def = team_records.get("팀_수비", pd.DataFrame())
    t_run = team_records.get("팀_주루", pd.DataFrame())

    def team_row(df, team_name):
        if df is None or df.empty or "팀명" not in df.columns:
            return None
        d = df[df["팀명"].astype(str).str.strip() == team_name]
        return d.iloc[0] if not d.empty else None

    hr = team_row(t_hit, team)
    pr = team_row(t_pit, team)
    dr = team_row(t_def, team)
    rr = team_row(t_run, team)

    if all(x is None for x in [hr, pr, dr, rr]):
        st.info("KBO 팀 공식 기록이 아직 등록되지 않았습니다. 데이터 메뉴에서 팀 기록 Excel을 업데이트해 주세요.")
    else:
        st.markdown("#### 타격")
        a,b,c,d = st.columns(4)
        a.metric("경기", metric_value(hr, "G") if hr is not None else "-")
        b.metric("타율", kbo_avg_value(hr, "AVG") if hr is not None else "-")
        c.metric("홈런", metric_value(hr, "HR") if hr is not None else "-")
        d.metric("타점", metric_value(hr, "RBI") if hr is not None else "-")

        a,b,c,d = st.columns(4)
        a.metric("득점", metric_value(hr, "R") if hr is not None else "-")
        b.metric("안타", metric_value(hr, "H") if hr is not None else "-")
        c.metric("2루타", metric_value(hr, "2B") if hr is not None else "-")
        d.metric("3루타", metric_value(hr, "3B") if hr is not None else "-")

        st.markdown("#### 투수")
        a,b,c,d = st.columns(4)
        a.metric("ERA", metric_value(pr, "ERA") if pr is not None else "-")
        b.metric("승", metric_value(pr, "W") if pr is not None else "-")
        c.metric("패", metric_value(pr, "L") if pr is not None else "-")
        d.metric("WHIP", metric_value(pr, "WHIP") if pr is not None else "-")

        a,b,c,d = st.columns(4)
        a.metric("세이브", metric_value(pr, "SV") if pr is not None else "-")
        b.metric("홀드", metric_value(pr, "HLD") if pr is not None else "-")
        c.metric("탈삼진", metric_value(pr, "SO") if pr is not None else "-")
        d.metric("이닝", kbo_ip_value(pr, "IP") if pr is not None else "-")

        st.markdown("#### 수비 · 주루")
        a,b,c,d = st.columns(4)
        a.metric("실책", metric_value(dr, "E") if dr is not None else "-")
        a2 = metric_value(dr, "FPCT") if dr is not None else "-"
        b.metric("수비율", a2)
        c.metric("도루", metric_value(rr, "SB") if rr is not None else "-")
        sbpct = metric_value(rr, "SB%") if rr is not None else "-"
        d.metric("도루 성공률", f"{sbpct}%" if sbpct != "-" else "-")

        with st.expander("KBO 팀 기록 전체 보기"):
            tabs = st.tabs(["타격","투수","수비","주루"])
            datasets = [
                (tabs[0], t_hit),
                (tabs[1], t_pit),
                (tabs[2], t_def),
                (tabs[3], t_run),
            ]
            for tab, df0 in datasets:
                with tab:
                    if df0 is not None and not df0.empty:
                        st.dataframe(df0, hide_index=True, use_container_width=True)

    st.caption(kbo_team_caption())
    st.caption("출처: KBO 공식 홈페이지")

    st.markdown("### NAVER 최신 자료")
    n1,n2,n3=st.columns(3)
    n1.metric("최신 경기",nv_team_games["game_id"].nunique() if not nv_team_games.empty else 0)
    n2.metric("투수진 최신 투구",len(nv_thrown))
    n3.metric("타선 최신 상대 투구",len(nv_seen))
    st.caption("출처: NAVER Sports 문자중계")

    st.markdown("### Play-by-Play 자료")
    p1,p2,p3=st.columns(3)
    p1.metric("경기",bp_team_games["game_id"].nunique() if not bp_team_games.empty else 0)
    p2.metric("투수진 투구",len(bp_thrown))
    p3.metric("타선이 본 투구",len(bp_seen))
    st.caption(f"자료 기간: {analysis_period(bp_team_games) or '-'}")
    st.caption("출처: 공개 Play-by-Play 데이터셋")

elif nav == "선수":
    loader = st.empty()
    with loader.container():
        show_center_loader()

    # 검색/필터 단계에서는 전체 투구 데이터를 읽지 않음
    players = player_master_light()

    loader.empty()
    if players.empty:
        st.info("데이터를 먼저 추가해 주세요.")
    else:
        player_favs = favorite_players()

        st.markdown("### ⭐ 관심 선수")
        if player_favs.empty:
            st.caption("등록된 관심 선수가 없습니다.")
        else:
            fav_cols = st.columns(6)
            for pos, (_, row) in enumerate(player_favs.iterrows()):
                with fav_cols[pos % 6]:
                    st.button(
                        f"★ {row['player_name']}",
                        key=f"player_page_fav_{row['player_id']}",
                        use_container_width=False,
                        on_click=go_to,
                        args=("선수", row["player_name"], None, row["player_id"])
                    )

        st.markdown("### 선수 찾기")
        if "player_search" not in st.session_state: st.session_state.player_search=""
        search_name=st.text_input("🔍 이름 검색",placeholder="선수 이름을 입력하세요",key="player_search").strip()
        st.caption("이름으로 바로 검색하거나, 아래 필터로 선수 목록을 좁혀볼 수 있습니다.")

        # 사용자가 직접 검색어를 바꾸면 이전에 클릭했던 특정 선수 ID 선택은 해제
        prev_search = st.session_state.get("_last_player_search", "")
        if search_name != prev_search:
            if "_last_player_search" in st.session_state:
                st.session_state.selected_player_id = ""
            st.session_state._last_player_search = search_name
        f1,f2=st.columns(2); teams=["전체"]+[t for t in ["KT","LG","삼성","두산","KIA","NC","SSG","롯데","키움","한화"] if t in set(players["team"])]
        with f1: team_filter=st.selectbox("구단",teams)
        with f2: role_filter=st.selectbox("구분",["전체","투수","타자"])
        if search_name: filtered=players[players["name"].str.contains(search_name,case=False,na=False)].copy()
        else:
            filtered=players.copy()
            if team_filter!="전체": filtered=filtered[filtered["team"]==team_filter]
            if role_filter=="투수": filtered=filtered[filtered["role"].isin(["투수","투타"])]
            elif role_filter=="타자": filtered=filtered[filtered["role"].isin(["타자","투타"])]
        if filtered.empty:
            st.info("조건에 맞는 선수가 없습니다.")
        else:
            direct = bool(search_name and len(filtered) == 1)
            filter_used = bool(search_name or team_filter != "전체" or role_filter != "전체")

            player_id=player_name=player_team=player_role=None

            selected_pid = str(st.session_state.get("selected_player_id", "") or "").strip()

            # 동일 이름 선수가 여러 명이어도 클릭한 선수 ID를 우선 사용
            if selected_pid:
                picked = players[players["id"].astype(str) == selected_pid]
                if not picked.empty:
                    only = picked.iloc[0]
                    player_id,player_name,player_team,player_role = (
                        str(only["id"]), only["name"], only["team"], only["role"]
                    )

            if player_id is None and direct:
                only=filtered.iloc[0]
                player_id,player_name,player_team,player_role = (
                    str(only["id"]), only["name"], only["team"], only["role"]
                )

            elif player_id is None and filter_used:
                st.caption(f"선수 {len(filtered):,}명 · 이름을 클릭하면 바로 분석합니다.")

                # 기존 '분석할 선수 선택' 드롭다운을 없애고
                # 검색 결과에서 선수 이름을 직접 클릭하도록 변경
                result_cols = st.columns(4)
                for pos, (_, row) in enumerate(filtered.iterrows()):
                    with result_cols[pos % 4]:
                        st.button(
                            f"{row['name']} · {row['team']} · {row['role']}",
                            key=f"player_result_{row['id']}_{pos}",
                            use_container_width=True,
                            on_click=go_to,
                            args=("선수", row["name"], None, row["id"])
                        )
            if player_id is not None:
                st.markdown(f"### {player_name}"); st.caption(" · ".join([x for x in [player_team,player_role] if x])); fav_now=is_favorite(player_id)
                if st.button("★ 관심 선수 해제" if fav_now else "☆ 관심 선수 등록",key=f"fav_{player_id}"): toggle_favorite(player_id,player_name); st.rerun()
                with st.spinner("선수 데이터를 불러오는 중입니다..."):
                    kbo_hitters,kbo_pitchers=kbo_records()
                    kbo_hitter=find_kbo_player(kbo_hitters,player_name,player_team)
                    kbo_pitcher=find_kbo_player(kbo_pitchers,player_name,player_team)

                    base_player_p, base_player_g = base_player_data(player_id)

                    path_text, modified_ns = _base_file_signature()
                    base_ids = load_home_base_summary(path_text, modified_ns)["game_ids"]
                    naver_player_p = extra_pitches()
                    if not naver_player_p.empty:
                        naver_player_p = naver_player_p[
                            ~naver_player_p["game_id"].astype(str).isin(base_ids)
                        ].copy()

                    bp_batter=base_player_p[
                        base_player_p["batter_id"].astype(str)==str(player_id)
                    ].copy() if not base_player_p.empty else pd.DataFrame()
                    bp_pitcher=base_player_p[
                        base_player_p["pitcher_id"].astype(str)==str(player_id)
                    ].copy() if not base_player_p.empty else pd.DataFrame()
                    nv_batter=naver_player_p[
                        naver_player_p["batter_id"].astype(str)==str(player_id)
                    ].copy() if not naver_player_p.empty else pd.DataFrame()
                    nv_pitcher=naver_player_p[
                        naver_player_p["pitcher_id"].astype(str)==str(player_id)
                    ].copy() if not naver_player_p.empty else pd.DataFrame()

                    batter_data=pd.concat([bp_batter,nv_batter],ignore_index=True)
                    pitcher_data=pd.concat([bp_pitcher,nv_pitcher],ignore_index=True)
                    has_batter,has_pitcher=not batter_data.empty,not pitcher_data.empty
                player_all=pd.concat([batter_data,pitcher_data],ignore_index=True) if (has_batter or has_pitcher) else pd.DataFrame(); analysis_header(f"{player_name} 선수 분석")
                with st.container(border=True):
                    if has_pitcher:
                        top_pitch,_,top_share = top_pitch_info(pitcher_data)
                        whiff_rate = action_rate(pitcher_data, ["헛스윙"])
                        first_strike_rate = first_pitch_rate(
                            pitcher_data,
                            ["스트라이크","헛스윙","파울","인플레이"]
                        )

                        era_pct = None
                        whip_pct = None
                        if kbo_pitcher is not None:
                            era_pct = percentile_value(
                                kbo_pitchers, "ERA", kbo_pitcher.get("ERA"), higher_is_better=False
                            )
                            whip_pct = percentile_value(
                                kbo_pitchers, "WHIP", kbo_pitcher.get("WHIP"), higher_is_better=False
                            )

                        recent_delta = speed_delta(nv_pitcher, bp_pitcher)

                        st.markdown("#### 투수 분석 지표")
                        r1,r2,r3 = st.columns(3)
                        r1.metric("ERA 백분위", pct_text(era_pct))
                        r2.metric("WHIP 백분위", pct_text(whip_pct))
                        r3.metric("주력 구종", top_pitch)

                        r4,r5,r6 = st.columns(3)
                        r4.metric("주력 구종 비율", f"{top_share:.1f}%" if top_pitch != "-" else "-")
                        r5.metric("헛스윙 유도율", pct_text(whiff_rate))
                        r6.metric("초구 스트라이크 비율", pct_text(first_strike_rate))

                        if recent_delta is not None:
                            st.metric("NAVER 최신 구속 변화", delta_text(recent_delta))

                        summary_parts = []
                        if era_pct is not None and whip_pct is not None:
                            summary_parts.append(
                                f"KBO 공식 기록 기준 ERA 백분위 {era_pct:.1f}%, WHIP 백분위 {whip_pct:.1f}%"
                            )
                        if top_pitch != "-":
                            summary_parts.append(f"전체 투구자료에서 {top_pitch} 사용 비율 {top_share:.1f}%")
                        if whiff_rate is not None:
                            summary_parts.append(f"헛스윙 유도율 {whiff_rate:.1f}%")
                        if recent_delta is not None:
                            summary_parts.append(
                                f"NAVER 최신 자료의 평균 구속은 기본 Play-by-Play 대비 {recent_delta:+.1f} km/h"
                            )
                        if summary_parts:
                            st.markdown("#### 분석 요약")
                            st.write(" · ".join(summary_parts))

                        mix=pitcher_data.groupby("pitch_type",dropna=True).agg(
                            투구수=("pitch_id","count"),평균구속=("speed","mean")
                        ).reset_index()
                        if not mix.empty:
                            mix["평균구속"]=mix["평균구속"].round(1)
                            st.markdown("#### 구종 구성")
                            st.bar_chart(mix.set_index("pitch_type")["투구수"])
                            st.dataframe(
                                mix.sort_values("투구수",ascending=False),
                                hide_index=True,
                                use_container_width=True
                            )

                        loc=pitcher_data[["plate_x","plate_y","pitch_type"]].dropna(
                            subset=["plate_x","plate_y"]
                        )
                        if not loc.empty:
                            st.markdown("#### 투구 위치")
                            st.scatter_chart(loc,x="plate_x",y="plate_y",color="pitch_type")

                    if has_batter:
                        top_seen,_,top_seen_share = top_pitch_info(batter_data)
                        batter_whiff = action_rate(batter_data, ["헛스윙"])
                        first_swing = first_pitch_rate(
                            batter_data,
                            ["헛스윙","파울","인플레이"]
                        )

                        avg_pct = None
                        gpa_pct = None
                        if kbo_hitter is not None:
                            avg_pct = percentile_value(
                                kbo_hitters, "AVG", kbo_hitter.get("AVG"), higher_is_better=True
                            )
                            gpa_pct = percentile_value(
                                kbo_hitters, "GPA", kbo_hitter.get("GPA"), higher_is_better=True
                            )

                        recent_seen_delta = speed_delta(nv_batter, bp_batter)

                        st.markdown("#### 타자 분석 지표")
                        r1,r2,r3 = st.columns(3)
                        r1.metric("AVG 백분위", pct_text(avg_pct))
                        r2.metric("GPA 백분위", pct_text(gpa_pct))
                        r3.metric("가장 많이 본 구종", top_seen)

                        r4,r5,r6 = st.columns(3)
                        r4.metric("최다 상대 구종 비율", f"{top_seen_share:.1f}%" if top_seen != "-" else "-")
                        r5.metric("헛스윙 비율", pct_text(batter_whiff))
                        r6.metric("초구 스윙 비율", pct_text(first_swing))

                        if recent_seen_delta is not None:
                            st.metric("NAVER 최신 상대구속 변화", delta_text(recent_seen_delta))

                        summary_parts = []
                        if avg_pct is not None and gpa_pct is not None:
                            summary_parts.append(
                                f"KBO 공식 기록 기준 AVG 백분위 {avg_pct:.1f}%, GPA 백분위 {gpa_pct:.1f}%"
                            )
                        if top_seen != "-":
                            summary_parts.append(
                                f"전체 투구자료에서 가장 많이 상대한 구종은 {top_seen} ({top_seen_share:.1f}%)"
                            )
                        if batter_whiff is not None:
                            summary_parts.append(f"헛스윙 비율 {batter_whiff:.1f}%")
                        if recent_seen_delta is not None:
                            summary_parts.append(
                                f"NAVER 최신 자료의 상대 평균 구속은 기본 Play-by-Play 대비 {recent_seen_delta:+.1f} km/h"
                            )
                        if summary_parts:
                            st.markdown("#### 분석 요약")
                            st.write(" · ".join(summary_parts))

                        d=batter_data.copy()
                        d["행동"]=d.pitch_text.apply(action)
                        acts=d["행동"].value_counts().rename_axis("결과").reset_index(name="횟수")
                        if not acts.empty:
                            st.markdown("#### 투구 결과")
                            st.bar_chart(acts.set_index("결과")["횟수"])

                        m=d.groupby("pitch_type",dropna=True).size().reset_index(name="투구수")
                        if not m.empty:
                            st.markdown("#### 상대 구종")
                            st.bar_chart(m.set_index("pitch_type")["투구수"])
                evidence_header()
                st.markdown("### KBO 공식 기록"); st.caption(kbo_player_caption()); st.caption("출처: KBO 공식 홈페이지")
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
