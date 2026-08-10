

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

TEAM = {
    "KT":"KT","LG":"LG","SS":"삼성","OB":"두산","HT":"KIA",
    "SK":"SSG","WO":"키움","LT":"롯데","NC":"NC","HH":"한화"
}
PRIORITY_TEAMS = ["KT","LG","삼성","두산","KIA"]

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

def current_counts():
    return qdf("""
    SELECT
      (SELECT COUNT(*) FROM games) games,
      (SELECT COUNT(*) FROM pitches) pitches,
      (SELECT COUNT(DISTINCT id) FROM (
        SELECT pitcher_id id FROM pitches WHERE COALESCE(pitcher_id,'')<>''
        UNION
        SELECT batter_id FROM pitches WHERE COALESCE(batter_id,'')<>''
      )) players
    """).iloc[0]

st.set_page_config(page_title="BASEBALL SCOUT", page_icon="⚾", layout="wide")
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
</style>
""", unsafe_allow_html=True)

# 발표 모드: 외부 요청 기능을 완전히 숨기고 저장 데이터만 사용
if "presentation_mode" not in st.session_state:
    st.session_state.presentation_mode = False

top1, top2 = st.columns([5,1])
with top1:
    st.markdown('<div class="brand">⚾ BASEBALL SCOUT</div>', unsafe_allow_html=True)

with top2:
    st.session_state.presentation_mode = st.toggle("발표 모드", value=st.session_state.presentation_mode,
                                                   help="저장된 데이터만 사용합니다. NAVER에 새 요청을 보내지 않습니다.")

if st.session_state.presentation_mode:
    st.caption("발표 모드 · 저장된 데이터만 사용 중")

nav_items = ["홈","팀","선수"] if st.session_state.presentation_mode else ["홈","경기 추가","팀","선수"]
nav = st.radio("", nav_items, horizontal=True, label_visibility="collapsed")

games = qdf("SELECT * FROM games ORDER BY game_date DESC, saved_at DESC")
counts = current_counts()

if nav == "홈":
    a,b = st.columns(2)
    a.metric("저장 경기", int(counts.games))
    b.metric("등록 선수", int(counts.players))

    st.markdown("### ⭐ 관심 선수")
    favs = favorite_players()
    if favs.empty:
        st.caption("선수 페이지에서 ☆를 눌러 관심 선수를 등록할 수 있습니다.")
    else:
        fav_cols = st.columns(min(4, len(favs)))
        for pos, (_, row) in enumerate(favs.iterrows()):
            with fav_cols[pos % len(fav_cols)]:
                st.markdown(f"**★ {row['player_name']}**")

    st.markdown("### 최근 경기")
    if games.empty:
        st.info("아직 저장된 경기가 없습니다.")
    else:
        show = games[["game_date","away_team","home_team"]].head(10).copy()
        show.columns = ["날짜","원정","홈"]
        st.dataframe(show, use_container_width=True, hide_index=True)

    st.markdown("### 분석 바로가기")
    c1,c2 = st.columns(2)
    with c1:
        st.markdown("**팀 분석**")
        st.caption("투수진의 구종 구성과 타선이 상대해 온 구종 패턴을 봅니다.")
    with c2:
        st.markdown("**선수 분석**")
        st.caption("타자와 투수의 누적 투구 데이터를 선수별로 봅니다.")

elif nav == "경기 추가":
    st.markdown("## 경기 추가")
    st.caption("경기 주소를 붙여넣고 데이터를 추가하세요.")

    entries = []
    for i in range(1,4):
        entries.append(st.text_input(f"경기 {i}", key=f"game_{i}", placeholder="경기 주소 붙여넣기"))

    if "extra_games" not in st.session_state:
        st.session_state.extra_games = 0

    if st.button("+ 경기 추가", use_container_width=False):
        st.session_state.extra_games = min(2, st.session_state.extra_games + 1)
        st.rerun()

    for j in range(st.session_state.extra_games):
        idx = j + 4
        entries.append(st.text_input(f"경기 {idx}", key=f"game_{idx}", placeholder="경기 주소 붙여넣기"))

    if st.button("데이터 가져오기", type="primary", use_container_width=True):
        values = [x.strip() for x in entries if x.strip()]
        if not values:
            st.warning("경기를 입력해 주세요.")
        else:
            progress = st.progress(0)
            for i, x in enumerate(values[:5], start=1):
                try:
                    status, msg = collect_game(x)
                    if status == "new":
                        st.success(msg)
                    elif status == "cached":
                        st.info(msg)
                    else:
                        st.error(msg)
                except Exception as e:
                    st.error(f"수집을 중단했습니다: {e}")
                    break
                progress.progress(i / len(values[:5]))

elif nav == "팀":
    st.markdown("## 팀")
    present = sorted(set(games.away_team.dropna()).union(set(games.home_team.dropna()))) if not games.empty else []
    priority = [t for t in PRIORITY_TEAMS if t in present]
    options = priority + [t for t in present if t not in priority]

    if not options:
        st.info("먼저 경기를 추가해 주세요.")
    else:
        team = st.selectbox("팀 선택", options)
        n_games = len(qdf("SELECT game_id FROM games WHERE away_team=? OR home_team=?", (team,team)))
        thrown = qdf("SELECT * FROM pitches WHERE defense_team=?", (team,))
        seen = qdf("SELECT * FROM pitches WHERE offense_team=?", (team,))

        a,b,c = st.columns(3)
        a.metric("저장 경기", n_games)
        b.metric("투수진 투구", len(thrown))
        c.metric("타선이 본 투구", len(seen))

        st.markdown("### 투수진 구종 구성")
        if thrown.empty:
            st.info("데이터가 없습니다.")
        else:
            mix = thrown.groupby("pitch_type", dropna=True).agg(
                투구수=("pitch_id","count"),
                평균구속=("speed","mean")
            ).reset_index()
            mix["평균구속"] = mix["평균구속"].round(1)
            st.bar_chart(mix.set_index("pitch_type")["투구수"])
            st.dataframe(mix.sort_values("투구수",ascending=False), hide_index=True, use_container_width=True)

        st.markdown("### 타선이 상대해 온 구종")
        if not seen.empty:
            mix2 = seen.groupby("pitch_type", dropna=True).size().reset_index(name="투구수")
            st.bar_chart(mix2.set_index("pitch_type")["투구수"])
            st.dataframe(mix2.sort_values("투구수",ascending=False), hide_index=True, use_container_width=True)

elif nav == "선수":
    st.markdown("## 선수")
    players = qdf("""
    SELECT DISTINCT id,name FROM (
        SELECT pitcher_id id,pitcher_name name FROM pitches
        UNION
        SELECT batter_id,batter_name FROM pitches
    )
    WHERE COALESCE(name,'')<>'' ORDER BY name
    """)

    if players.empty:
        st.info("먼저 경기를 추가해 주세요.")
    else:
        label_map = {f"{r['name']} ({r['id']})": (r["id"],r["name"]) for _,r in players.iterrows()}
        selected = st.selectbox("선수 검색", list(label_map))
        player_id, player_name = label_map[selected]

        fav_now = is_favorite(player_id)
        fav_label = "★ 관심 선수 해제" if fav_now else "☆ 관심 선수 등록"
        if st.button(fav_label, key=f"fav_{player_id}"):
            toggle_favorite(player_id, player_name)
            st.rerun()

        role = st.radio("", ["타자","투수"], horizontal=True, label_visibility="collapsed")

        if role == "타자":
            d = qdf("SELECT * FROM pitches WHERE batter_id=?", (player_id,))
            a,b,c = st.columns(3)
            a.metric("본 투구", len(d))
            b.metric("경기", d.game_id.nunique() if not d.empty else 0)
            c.metric("평균 구속", f"{d.speed.dropna().mean():.1f} km/h" if not d.empty and d.speed.notna().any() else "-")

            if not d.empty:
                d["행동"] = d.pitch_text.apply(action)
                acts = d["행동"].value_counts().rename_axis("결과").reset_index(name="횟수")
                st.markdown("### 투구 결과")
                st.bar_chart(acts.set_index("결과")["횟수"])

                first = d[d.pitch_num == 1].copy()
                if not first.empty:
                    first["결과"] = first.pitch_text.apply(action)
                    f = first["결과"].value_counts().rename_axis("결과").reset_index(name="횟수")
                    st.markdown("### 초구")
                    st.dataframe(f, hide_index=True, use_container_width=True)

                st.markdown("### 상대 구종")
                m = d.groupby("pitch_type",dropna=True).size().reset_index(name="투구수")
                st.bar_chart(m.set_index("pitch_type")["투구수"])

        else:
            d = qdf("SELECT * FROM pitches WHERE pitcher_id=?", (player_id,))
            a,b,c = st.columns(3)
            a.metric("투구", len(d))
            b.metric("경기", d.game_id.nunique() if not d.empty else 0)
            c.metric("평균 구속", f"{d.speed.dropna().mean():.1f} km/h" if not d.empty and d.speed.notna().any() else "-")

            if not d.empty:
                mix = d.groupby("pitch_type",dropna=True).agg(
                    투구수=("pitch_id","count"),
                    평균구속=("speed","mean")
                ).reset_index()
                mix["평균구속"] = mix["평균구속"].round(1)
                st.markdown("### 구종 구성")
                st.bar_chart(mix.set_index("pitch_type")["투구수"])
                st.dataframe(mix.sort_values("투구수",ascending=False), hide_index=True, use_container_width=True)

                st.markdown("### 투구 위치")
                loc = d[["plate_x","plate_y","pitch_type"]].dropna(subset=["plate_x","plate_y"])
                if not loc.empty:
                    st.scatter_chart(loc, x="plate_x", y="plate_y", color="pitch_type")
