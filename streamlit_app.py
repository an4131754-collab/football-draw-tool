from __future__ import annotations

import hashlib
import io
from pathlib import Path, PureWindowsPath
from typing import Any

import streamlit as st

import tournament_tools
from tournament_tools import (
    BASE_DIR,
    clear_latest_artifacts,
    create_draw_artifacts,
    get_artifact_filenames,
    get_latest_draw_data,
    load_config,
    load_teams,
    normalize_download_options,
    resolve_registration_path,
)

ALLOWED_UPLOAD_EXTENSIONS = {".xlsx", ".xlsm"}


st.set_page_config(
    page_title="Interdept Cup Draw Studio",
    page_icon="⚽",
    layout="wide",
    initial_sidebar_state="collapsed",
)


def inject_styles() -> None:
    st.markdown(
        """
        <style>
          :root {
            --bg: #050505;
            --ink: #f5f2ea;
            --muted: #9b978f;
            --line: rgba(245, 242, 234, 0.16);
            --accent: #d6ff63;
            --accent-2: #9ae6ff;
            --danger: #ff765f;
            --panel: rgba(15, 15, 15, 0.78);
          }

          .stApp {
            color: var(--ink);
            background:
              radial-gradient(circle at 14% 7%, rgba(214,255,99,0.18), transparent 22%),
              radial-gradient(circle at 86% 3%, rgba(154,230,255,0.14), transparent 24%),
              linear-gradient(135deg, #050505 0%, #0e0e0e 48%, #020202 100%);
          }

          .stApp::before {
            content: "";
            position: fixed;
            inset: 0;
            pointer-events: none;
            opacity: .22;
            background-image:
              linear-gradient(rgba(255,255,255,.08) 1px, transparent 1px),
              linear-gradient(90deg, rgba(255,255,255,.07) 1px, transparent 1px);
            background-size: 84px 84px;
            transform: perspective(820px) rotateX(64deg) translateY(-14%);
            transform-origin: top;
            animation: gridDrift 18s linear infinite;
          }

          .stApp::after {
            content: "";
            position: fixed;
            inset: auto -10% 0;
            height: 34vh;
            pointer-events: none;
            background:
              linear-gradient(90deg, transparent 0 48%, rgba(214,255,99,.17) 49%, transparent 51% 100%),
              linear-gradient(to top, rgba(214,255,99,.12), transparent 72%);
            clip-path: polygon(40% 0, 60% 0, 100% 100%, 0 100%);
          }

          header[data-testid="stHeader"] { background: transparent; }
          [data-testid="stToolbar"],
          [data-testid="stDecoration"],
          #MainMenu,
          footer { visibility: hidden; }

          .block-container {
            max-width: 1180px;
            padding-top: 1.65rem;
            padding-bottom: 5rem;
          }

          .brand-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 1rem;
            margin-bottom: clamp(1.8rem, 4.2vw, 3.5rem);
          }

          .brand {
            display: grid;
            gap: .25rem;
            font: 900 .78rem/1.1 "Bahnschrift", "Arial Narrow", "Microsoft JhengHei", sans-serif;
            letter-spacing: .24em;
            text-transform: uppercase;
          }

          .brand span:last-child {
            color: var(--muted);
            letter-spacing: .34em;
          }

          .nav {
            display: flex;
            gap: 1.5rem;
            color: var(--muted);
            font-size: .75rem;
            letter-spacing: .18em;
            text-transform: uppercase;
          }

          .hero {
            display: grid;
            grid-template-columns: minmax(0, 1.14fr) minmax(240px, .56fr);
            gap: clamp(1.6rem, 4vw, 3.6rem);
            align-items: center;
            margin-bottom: 1.65rem;
          }

          .hero h1 {
            margin: 0;
            max-width: 830px;
            color: var(--ink);
            font: 900 clamp(3.25rem, 9.2vw, 7.35rem)/.84 "Bahnschrift", "Arial Narrow", "Microsoft JhengHei", sans-serif;
            letter-spacing: -.08em;
            text-transform: uppercase;
          }

          .hero p {
            max-width: 630px;
            color: var(--muted);
            font-size: clamp(.96rem, 1.3vw, 1.08rem);
            line-height: 1.78;
          }

          .hero code,
          code {
            color: var(--accent);
            font-weight: 800;
          }

          .hero-art {
            position: relative;
            width: min(100%, 360px);
            min-height: 320px;
            justify-self: center;
            border: 1px solid var(--line);
            border-radius: 28px;
            overflow: hidden;
            background:
              linear-gradient(145deg, rgba(255,255,255,.12), rgba(255,255,255,.025)),
              radial-gradient(circle at 50% 20%, rgba(214,255,99,.20), transparent 28%),
              #101010;
            box-shadow: 0 26px 84px rgba(0,0,0,.48);
            animation: floatCard 7s ease-in-out infinite;
          }

          .hero-art::before {
            content: "";
            position: absolute;
            inset: 8%;
            border-radius: 999px;
            border: 1px solid rgba(255,255,255,.16);
            background:
              radial-gradient(circle, transparent 0 36%, rgba(214,255,99,.08) 37% 38%, transparent 39%),
              conic-gradient(from 180deg, transparent, rgba(214,255,99,.22), transparent, rgba(154,230,255,.18), transparent);
            animation: spinSlow 18s linear infinite;
          }

          .hero-art::after {
            content: "SYSTEM RANDOM";
            position: absolute;
            left: 20px;
            right: 20px;
            bottom: 20px;
            padding: 16px 18px;
            border: 1px solid rgba(255,255,255,.18);
            border-radius: 22px;
            background: rgba(0,0,0,.54);
            backdrop-filter: blur(18px);
            font: 900 1.18rem/1 "Bahnschrift", "Arial Narrow", sans-serif;
            letter-spacing: -.04em;
          }

          .hero-art .spec {
            position: absolute;
            top: 20px;
            left: 20px;
            right: 20px;
            display: flex;
            justify-content: space-between;
            color: var(--muted);
            font-size: .64rem;
            letter-spacing: .18em;
            text-transform: uppercase;
          }

          .ball {
            position: absolute;
            left: 50%;
            top: 48%;
            width: 112px;
            height: 112px;
            transform: translate(-50%, -50%);
            border-radius: 50%;
            background:
              radial-gradient(circle at 38% 30%, #fff 0 7%, transparent 8%),
              radial-gradient(circle at 60% 62%, rgba(214,255,99,.92) 0 9%, transparent 10%),
              linear-gradient(145deg, #f7f3e7, #676767 58%, #111);
            box-shadow: 0 28px 82px rgba(0,0,0,.54), 0 0 70px rgba(214,255,99,.28);
          }

          .ticker {
            overflow: hidden;
            border-block: 1px solid var(--line);
            margin: 1rem 0 1.1rem;
            padding: .72rem 0;
            color: var(--muted);
            white-space: nowrap;
            font: 900 clamp(1.25rem, 3vw, 2.85rem)/1 "Bahnschrift", "Arial Narrow", sans-serif;
            letter-spacing: -.05em;
            text-transform: uppercase;
          }

          .ticker-track {
            display: inline-flex;
            gap: 2rem;
            animation: marquee 24s linear infinite;
          }

          .lux-panel {
            border: 1px solid var(--line);
            border-radius: 30px;
            background: var(--panel);
            box-shadow: 0 34px 110px rgba(0,0,0,.5);
            backdrop-filter: blur(22px);
            padding: clamp(1.2rem, 3vw, 1.8rem);
            min-height: 100%;
          }

          .kicker {
            margin: 0 0 .45rem;
            color: var(--accent);
            font-size: .72rem;
            letter-spacing: .2em;
            text-transform: uppercase;
            font-weight: 900;
          }

          .section-title {
            margin: 0 0 1rem;
            color: var(--ink);
            font: 900 clamp(1.9rem, 3vw, 3.1rem)/.95 "Bahnschrift", "Arial Narrow", "Microsoft JhengHei", sans-serif;
            letter-spacing: -.06em;
            text-transform: uppercase;
          }

          .muted {
            color: var(--muted);
            line-height: 1.75;
          }

          div[data-testid="stForm"] {
            border: 1px solid var(--line);
            border-radius: 24px;
            background: rgba(255,255,255,.035);
            padding: 1rem;
          }

          label,
          [data-testid="stWidgetLabel"] p {
            color: var(--ink) !important;
            font-size: .76rem !important;
            font-weight: 900 !important;
            letter-spacing: .14em;
            text-transform: uppercase;
          }

          div[data-baseweb="input"] > div,
          div[data-baseweb="select"] > div,
          [data-testid="stFileUploaderDropzone"] {
            border: 1px solid var(--line) !important;
            border-radius: 18px !important;
            background: rgba(255,255,255,.055) !important;
            color: var(--ink) !important;
          }

          [data-testid="stFileUploaderDropzone"] small,
          [data-testid="stFileUploaderDropzone"] span {
            color: var(--muted) !important;
          }

          .stButton > button,
          .stDownloadButton > button,
          button[kind="primary"] {
            min-height: 48px;
            width: 100%;
            border: 1px solid transparent;
            border-radius: 999px;
            background: var(--accent);
            color: #080808;
            font-weight: 900;
            letter-spacing: .18em;
            text-transform: uppercase;
            box-shadow: 0 18px 46px rgba(214,255,99,.20);
            transition: transform .18s ease, box-shadow .18s ease;
          }

          .stButton > button:hover,
          .stDownloadButton > button:hover {
            transform: translateY(-2px);
            color: #080808;
            border-color: transparent;
            box-shadow: 0 24px 54px rgba(214,255,99,.28);
          }

          .group-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: .85rem;
            margin-top: 1rem;
          }

          .group-card {
            position: relative;
            min-height: 210px;
            border: 1px solid var(--line);
            border-radius: 24px;
            padding: 1.1rem;
            overflow: hidden;
            background:
              linear-gradient(145deg, rgba(255,255,255,.09), rgba(255,255,255,.03)),
              #101010;
          }

          .group-card h3 {
            margin: 0 0 1rem;
            color: var(--ink);
            font: 900 clamp(2rem, 4vw, 4rem)/.8 "Bahnschrift", "Arial Narrow", sans-serif;
            letter-spacing: -.08em;
          }

          .group-card ol {
            display: grid;
            gap: .55rem;
            margin: 0;
            padding: 0;
            list-style: none;
            counter-reset: teams;
          }

          .group-card li {
            counter-increment: teams;
            display: grid;
            grid-template-columns: 2rem 1fr;
            color: var(--ink);
          }

          .group-card li::before {
            content: counter(teams, decimal-leading-zero);
            color: var(--muted);
            font-size: .74rem;
            letter-spacing: .12em;
          }

          .team-list {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: .55rem;
            margin-top: 1rem;
          }

          .team-chip,
          .metric {
            border: 1px solid var(--line);
            border-radius: 16px;
            background: rgba(255,255,255,.045);
            padding: .75rem .8rem;
          }

          .metric strong {
            display: block;
            color: var(--ink);
            font: 900 2rem/.9 "Bahnschrift", "Arial Narrow", sans-serif;
            letter-spacing: -.06em;
          }

          .status-box {
            border: 1px solid rgba(214,255,99,.28);
            border-radius: 20px;
            padding: 1rem;
            background: linear-gradient(135deg, rgba(214,255,99,.11), rgba(255,255,255,.035));
            color: var(--muted);
            line-height: 1.7;
            margin: 1rem 0;
          }

          .status-box.infeasible {
            border-color: rgba(255,118,95,.34);
            background: rgba(255,118,95,.1);
          }

          @keyframes gridDrift { to { background-position: 0 168px; } }
          @keyframes floatCard { 0%,100% { transform: translateY(0) rotate(-1deg); } 50% { transform: translateY(-12px) rotate(1deg); } }
          @keyframes spinSlow { to { transform: rotate(360deg); } }
          @keyframes marquee { to { transform: translateX(-50%); } }

          @media (prefers-reduced-motion: reduce) {
            *, *::before, *::after {
              animation-duration: 1ms !important;
              animation-iteration-count: 1 !important;
              transition-duration: 1ms !important;
            }
          }

          @media (max-width: 760px) {
            .brand-row, .hero { display: block; }
            .nav { display: none; }
            .hero h1 { font-size: clamp(3.1rem, 18vw, 5.4rem); }
            .hero-art { min-height: 280px; margin: 1.5rem auto 0; }
            .group-grid, .team-list { grid-template-columns: 1fr; }
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    inject_styles()
    config = load_config(BASE_DIR)
    latest_draw = get_latest_draw_data(BASE_DIR)

    render_header()
    render_messages()

    st.markdown(
        '<div class="ticker"><div class="ticker-track">'
        "<span>Groups</span><span>/</span><span>Schedule</span><span>/</span><span>PDF Proof</span><span>/</span><span>Excel Output</span><span>/</span>"
        "<span>Groups</span><span>/</span><span>Schedule</span><span>/</span><span>PDF Proof</span><span>/</span><span>Excel Output</span><span>/</span>"
        "</div></div>",
        unsafe_allow_html=True,
    )

    setup_col, result_col = st.columns([0.86, 1.14], gap="large")
    with setup_col:
        render_setup_panel(config, latest_draw)

    with result_col:
        render_results_panel(latest_draw)


def render_header() -> None:
    st.markdown(
        """
        <div class="brand-row">
          <div class="brand"><span>Interdept Cup</span><span>Draw Studio</span></div>
          <div class="nav"><span>Draw</span><span>Results</span><span>Outputs</span></div>
        </div>
        <section class="hero">
          <div>
            <h1>Draw the cup. Keep the trust.</h1>
            <p>
              一個給足球系際盃使用的抽籤與賽程工作室。上傳同格式 Google 表單 Excel，
              選擇組數、晉級規則與輸出格式，再用 <code>secrets.SystemRandom().shuffle()</code>
              完成公開、可追溯的抽籤。
            </p>
          </div>
          <div class="hero-art">
            <div class="spec"><span>OS entropy</span><span>Offline ready</span></div>
            <div class="ball"></div>
          </div>
        </section>
        """,
        unsafe_allow_html=True,
    )


def render_messages() -> None:
    message = st.session_state.pop("message", None)
    if not message:
        return

    level, text = message
    if level == "success":
        st.success(text)
    elif level == "warning":
        st.warning(text)
    else:
        st.error(text)


def render_setup_panel(config: dict[str, Any], latest_draw: dict[str, Any] | None) -> None:
    teams, teams_error = load_fallback_teams()
    latest_schedule = latest_draw.get("schedule", {}) if latest_draw else {}
    latest_constraints = latest_schedule.get("constraints", {})
    latest_download_options = normalize_download_options(latest_draw.get("download_options") if latest_draw else None)

    # Existing-schedule referee mode stores group_count=0 because no draw is created.
    # Do not reuse that metadata as defaults for the normal draw form.
    reuse_draw_defaults = bool(latest_draw and latest_draw.get("schedule_mode") != "existing_schedule")
    defaults = {
        "group_count": int(latest_draw.get("group_count", config["default_group_count"])) if reuse_draw_defaults else int(config["default_group_count"]),
        "advance_per_group": int(latest_draw.get("advancement", {}).get("advance_per_group", config["default_advance_per_group"])) if reuse_draw_defaults else int(config["default_advance_per_group"]),
        "wildcard_count": int(latest_draw.get("advancement", {}).get("wildcard_count", config["default_wildcard_count"])) if reuse_draw_defaults else int(config["default_wildcard_count"]),
        "knockout_format": latest_draw.get("knockout_format", config.get("default_knockout_format", "semifinal")) if reuse_draw_defaults else config.get("default_knockout_format", "semifinal"),
        "day1_latest_end": latest_constraints.get("day1_latest_end", config["default_day1_latest_end"]) if reuse_draw_defaults else config["default_day1_latest_end"],
        "day2_latest_end": latest_constraints.get("day2_latest_end", config["default_day2_latest_end"]) if reuse_draw_defaults else config["default_day2_latest_end"],
    }
    defaults["group_count"] = min(max(defaults["group_count"], 2), int(config["max_group_count"]))
    defaults["advance_per_group"] = max(defaults["advance_per_group"], 1)
    defaults["wildcard_count"] = max(defaults["wildcard_count"], 0)

    st.markdown('<div class="lux-panel">', unsafe_allow_html=True)
    st.markdown(
        '<p class="kicker">Configure</p>'
        '<h2 class="section-title">抽籤設定</h2>'
        '<p class="muted">上傳報名表 Excel，設定組數、晉級規則與輸出格式；如果沒有上傳，會使用本機 fallback 報名表。</p>',
        unsafe_allow_html=True,
    )

    with st.form("draw_form"):
        uploaded_file = st.file_uploader("上傳報名表 Excel", type=["xlsx", "xlsm"], help=f"隊名欄位預設讀取「{config['team_column']}」。")
        group_count = st.number_input("組數", min_value=2, max_value=int(config["max_group_count"]), value=defaults["group_count"], step=1)
        advance_per_group = st.number_input("每組晉級幾隊", min_value=1, value=defaults["advance_per_group"], step=1)
        wildcard_count = st.number_input("外卡晉級隊數", min_value=0, value=defaults["wildcard_count"], step=1)

        knockout_label_to_value = {
            "直接四強 - 4 隊晉級": "semifinal",
            "八強賽 - 8 隊晉級": "quarterfinal",
        }
        knockout_labels = list(knockout_label_to_value)
        default_knockout_label = label_for_value(knockout_label_to_value, defaults["knockout_format"])
        knockout_format_label = st.selectbox("淘汰賽格式", knockout_labels, index=knockout_labels.index(default_knockout_label))

        day1_options = list(config["latest_end_options"]["DAY1"])
        day2_options = list(config["latest_end_options"]["DAY2"])
        day1_latest_end = st.selectbox("DAY1 最晚結束時間", day1_options, index=option_index(day1_options, defaults["day1_latest_end"]))
        day2_latest_end = st.selectbox("DAY2 最晚結束時間", day2_options, index=option_index(day2_options, defaults["day2_latest_end"]))

        st.markdown('<p class="kicker" style="margin-top: .4rem;">輸出檔案</p>', unsafe_allow_html=True)
        output_col_1, output_col_2, output_col_3 = st.columns(3)
        with output_col_1:
            generate_json = st.checkbox("JSON", value=latest_download_options["json"])
        with output_col_2:
            generate_excel = st.checkbox("Excel", value=latest_download_options["schedule"])
        with output_col_3:
            generate_pdf = st.checkbox("PDF", value=latest_download_options["pdf"])

        submitted = st.form_submit_button("開始抽籤")

    if submitted:
        run_draw(
            uploaded_file=uploaded_file,
            group_count=int(group_count),
            advance_per_group=int(advance_per_group),
            wildcard_count=int(wildcard_count),
            knockout_format=knockout_label_to_value[knockout_format_label],
            download_options={
                "json": bool(generate_json),
                "schedule": bool(generate_excel),
                "pdf": bool(generate_pdf),
            },
            day1_latest_end=day1_latest_end,
            day2_latest_end=day2_latest_end,
        )

    st.divider()
    render_existing_schedule_panel(config)

    st.divider()
    st.markdown('<p class="kicker">Local Fallback</p><h2 class="section-title" style="font-size:2.2rem;">本機隊伍名單</h2>', unsafe_allow_html=True)
    if teams:
        st.markdown(f'<p class="muted">目前本機報名表讀到 {len(teams)} 隊。上傳檔案時，會以上傳檔案為準。</p>', unsafe_allow_html=True)
        render_team_chips(teams)
    else:
        st.markdown(f'<p class="muted">{teams_error or "尚未讀到本機報名表，請直接上傳 Excel 後抽籤。"}</p>', unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


def render_existing_schedule_panel(config: dict[str, Any]) -> None:
    st.markdown(
        '<p class="kicker">Referee Only</p>'
        '<h2 class="section-title" style="font-size:2.2rem;">既有賽程只排裁判</h2>'
        '<p class="muted">如果你已經自己排好賽程，可以在這裡上傳 Excel。系統只會新增或更新「裁判」工作表，不會重新抽籤或改你的賽程。</p>',
        unsafe_allow_html=True,
    )

    with st.form("existing_schedule_form"):
        schedule_file = st.file_uploader("上傳已排好的賽程 Excel", type=["xlsx", "xlsm"], key="existing_schedule_file")
        referee_names = st.text_area("裁判名單", help="一行一位裁判；產生後可在右側結果區繼續設定所屬隊伍與不可排場次。")
        submitted = st.form_submit_button("讀取賽程並建立裁判設定")

    if submitted:
        run_existing_schedule_referee_mode(schedule_file, referee_names, config)


def render_results_panel(latest_draw: dict[str, Any] | None) -> None:
    st.markdown('<div class="lux-panel">', unsafe_allow_html=True)
    if latest_draw is None:
        st.markdown(
            """
            <div style="min-height: 520px; display:grid; place-items:center; text-align:center;">
              <div>
                <p class="kicker">No Result Yet</p>
                <h2 class="section-title">等待開始</h2>
                <p class="muted">左側可以抽籤產生新賽程，也可以上傳已排好的賽程只安排裁判。</p>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)
        return

    st.markdown(
        f"""
        <p class="kicker">Latest Result</p>
        <h2 class="section-title">目前結果</h2>
        <p class="muted">建立時間：{latest_draw["drawn_at"]}<br>模式：<code>{latest_draw["random_function"]}</code></p>
        """,
        unsafe_allow_html=True,
    )

    metric_1, metric_2, metric_3 = st.columns(3)
    with metric_1:
        st.markdown(f'<div class="metric"><strong>{latest_draw.get("team_count", len(latest_draw["teams"]))}</strong>隊伍</div>', unsafe_allow_html=True)
    with metric_2:
        st.markdown(f'<div class="metric"><strong>{latest_draw.get("group_count", len(latest_draw["groups"]))}</strong>組數</div>', unsafe_allow_html=True)
    with metric_3:
        advancement = latest_draw.get("advancement", {})
        st.markdown(f'<div class="metric"><strong>{advancement.get("total_advancers", "TBD")}</strong>晉級</div>', unsafe_allow_html=True)

    if advancement:
        if advancement.get("summary"):
            st.markdown(f'<p class="muted">晉級規則：{advancement["summary"]}</p>', unsafe_allow_html=True)
        if latest_draw.get("schedule_mode") != "existing_schedule":
            st.markdown(f'<p class="muted">淘汰賽格式：{advancement.get("knockout_stage") or latest_draw.get("knockout_format")}</p>', unsafe_allow_html=True)

    schedule = latest_draw.get("schedule")
    if schedule:
        status_class = "infeasible" if schedule.get("status") != "scheduled" else ""
        status_title = "賽程已就緒" if schedule.get("status") == "scheduled" else "賽程需要調整"
        messages = "".join(f"<div>{message}</div>" for message in schedule.get("messages", []))
        extra = ""
        if schedule.get("status") != "scheduled":
            extra = "<div>可以放寬 DAY1/DAY2 最晚結束時間，或調整晉級隊數後重新產生。</div>"
        st.markdown(
            f'<div class="status-box {status_class}"><strong>{status_title}</strong>{messages}{extra}</div>',
            unsafe_allow_html=True,
        )

    if latest_draw.get("groups"):
        render_groups(latest_draw["groups"])
    render_referee_panel(latest_draw, load_config(BASE_DIR))

    items = build_download_items(latest_draw)
    if items:
        st.markdown('<p class="kicker" style="margin-top:1rem;">Outputs</p>', unsafe_allow_html=True)
        columns = st.columns(len(items))
        for column, item in zip(columns, items):
            with column:
                st.download_button(
                    item["label"],
                    data=item["path"].read_bytes(),
                    file_name=item["path"].name,
                    mime=item["mime"],
                    use_container_width=True,
                )
    else:
        st.info("這次沒有選擇下載輸出。系統仍會保留內部抽籤紀錄。")

    if st.button("清空目前結果", use_container_width=True):
        completed = clear_latest_artifacts(BASE_DIR)
        if completed:
            st.session_state["message"] = ("success", "已清空目前結果。")
        else:
            st.session_state["message"] = ("error", "部分 latest 檔案可能正被開啟，無法完全清空；請稍後再試。")
        st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)


def render_referee_panel(latest_draw: dict[str, Any], config: dict[str, Any]) -> None:
    schedule = latest_draw.get("schedule", {})
    if schedule.get("status") != "scheduled":
        return

    st.markdown('<div class="referee-panel">', unsafe_allow_html=True)
    st.markdown(
        '<p class="kicker" style="margin-top:1rem;">Referee Setup</p>'
        '<h2 class="section-title" style="font-size:2.2rem;">裁判排班</h2>'
        '<p class="muted">抽籤與賽程完成後，在這裡輸入裁判；每場會排 3 位，同時段不會重複安排同一人。</p>',
        unsafe_allow_html=True,
    )

    existing_referees = {item.get("name", ""): item for item in latest_draw.get("referees", [])}
    draw_key = latest_draw.get("drawn_at", "")
    if st.session_state.get("referee_names_drawn_at") != draw_key:
        st.session_state["referee_names_text"] = "\n".join(existing_referees)
        st.session_state["referee_names_drawn_at"] = draw_key
    elif "referee_names_text" not in st.session_state:
        st.session_state["referee_names_text"] = "\n".join(existing_referees)

    names_text = st.text_area(
        "裁判名單",
        key="referee_names_text",
        help="一行一位裁判。輸入後頁面會重新整理出每位裁判的設定。",
    )
    names = parse_referee_names(names_text)
    if not names:
        st.info("先輸入至少一位裁判姓名，再設定所屬隊伍與不可排場次。")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    match_options = build_match_options(latest_draw)
    no_affiliation_label = "無所屬隊伍"
    team_options = [no_affiliation_label] + list(latest_draw.get("teams", []))
    editor_rows = build_referee_editor_rows(names, existing_referees)
    editor_signature = hashlib.sha1((draw_key + "|" + "|".join(names)).encode("utf-8")).hexdigest()[:12]

    st.caption("在表格的「所屬隊伍」欄位選擇裁判屬於哪隊；沒有隊伍就保持「無所屬隊伍」。不可排場次請輸入場次號，例如 1,2,13。")
    edited_rows = st.data_editor(
        editor_rows,
        key=f"referee_editor_{editor_signature}",
        hide_index=True,
        use_container_width=True,
        disabled=["name"],
        column_config={
            "name": st.column_config.TextColumn("姓名", width="medium"),
            "affiliated_team": st.column_config.SelectboxColumn(
                "所屬隊伍",
                options=team_options,
                width="medium",
            ),
            "unavailable_match_nos": st.column_config.TextColumn(
                "不可排場次",
                help="輸入場次號，逗號分隔，例如 1,2,13。",
                width="large",
            ),
        },
    )

    with st.expander("場次號對照表", expanded=False):
        st.dataframe(build_match_reference_rows(match_options), hide_index=True, use_container_width=True)

    submitted = st.button("產生裁判表", use_container_width=True)

    warnings = latest_draw.get("referee_warnings", [])
    if warnings:
        for warning in warnings:
            st.warning(warning)
    elif latest_draw.get("referee_assignments"):
        st.success("裁判表已產生，Excel 下載檔已包含「裁判」工作表。")

    if submitted:
        run_referee_assignment(parse_referee_editor_rows(edited_rows), config)

    st.markdown("</div>", unsafe_allow_html=True)


def parse_referee_names(names_text: str) -> list[str]:
    names: list[str] = []
    seen: set[str] = set()
    for line in names_text.splitlines():
        name = line.strip()
        if not name or name in seen:
            continue
        names.append(name)
        seen.add(name)
    return names


def build_match_options(draw_data: dict[str, Any]) -> dict[str, int]:
    options: dict[str, int] = {}
    for match in draw_data.get("schedule", {}).get("matches", []):
        match_no = int(match.get("match_no", 0))
        home = match.get("home", match.get("home_label", ""))
        away = match.get("away", match.get("away_label", ""))
        label = f"#{match_no} {match.get('day', '')} {match.get('time', '')} {match.get('field', '')} - {home} vs {away}"
        options[label] = match_no
    return options


def build_referee_editor_rows(names: list[str], existing_referees: dict[str, dict[str, Any]]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for name in names:
        existing = existing_referees.get(name, {})
        unavailable = existing.get("unavailable_match_nos", [])
        rows.append(
            {
                "name": name,
                "affiliated_team": existing.get("affiliated_team") or "無所屬隊伍",
                "unavailable_match_nos": ",".join(str(match_no) for match_no in unavailable),
            }
        )
    return rows


def parse_referee_editor_rows(rows: Any) -> list[dict[str, Any]]:
    if hasattr(rows, "to_dict"):
        rows = rows.to_dict("records")

    referees: list[dict[str, Any]] = []
    for row in rows:
        name = str(row.get("name", "")).strip()
        if not name:
            continue
        affiliated_team = str(row.get("affiliated_team", "")).strip()
        if affiliated_team in {"No affiliation", "無所屬隊伍"}:
            affiliated_team = ""
        referees.append(
            {
                "name": name,
                "affiliated_team": affiliated_team,
                "unavailable_match_nos": row.get("unavailable_match_nos", ""),
            }
        )
    return referees


def build_match_reference_rows(match_options: dict[str, int]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for label, match_no in match_options.items():
        rows.append({"場次": match_no, "比賽": label})
    return rows


def run_referee_assignment(referees: list[dict[str, Any]], config: dict[str, Any]) -> None:
    try:
        latest_draw = get_latest_draw_data(BASE_DIR)
        if latest_draw is None:
            raise ValueError("尚未有抽籤結果，請先完成抽籤。")

        update_draw_referees = getattr(tournament_tools, "update_draw_referees", None)
        generate_artifacts = getattr(tournament_tools, "generate_artifacts", None)
        if update_draw_referees is None or generate_artifacts is None:
            raise ValueError("線上版本尚未同步到最新裁判排班程式，請先在 Streamlit 按 Reboot / Redeploy。")

        updated_draw = update_draw_referees(latest_draw, referees, config=config)
        if latest_draw.get("schedule_mode") == "existing_schedule":
            cached_key = st.session_state.get("existing_schedule_template_drawn_at")
            cached_bytes = st.session_state.get("existing_schedule_template_bytes")
            if cached_key == latest_draw.get("drawn_at") and cached_bytes:
                updated_draw["_existing_schedule_template_bytes"] = cached_bytes
        artifacts = generate_artifacts(updated_draw, base_dir=BASE_DIR)
        if artifacts.latest_sync_complete:
            st.session_state["message"] = ("success", "裁判排班已完成，Excel 已更新「裁判」工作表。")
        else:
            st.session_state["message"] = ("warning", "裁判排班已完成，但 latest 檔案同步可能被 Excel 開啟中擋住。")
        st.rerun()
    except Exception as exc:
        st.session_state["message"] = ("error", str(exc))
        st.rerun()


def run_existing_schedule_referee_mode(schedule_file: Any, referee_names: str, config: dict[str, Any]) -> None:
    try:
        if schedule_file is None:
            raise ValueError("請先上傳已排好的賽程 Excel。")

        source_name = Path(PureWindowsPath(schedule_file.name).name).name
        extension = Path(source_name).suffix.lower()
        if extension not in ALLOWED_UPLOAD_EXTENSIONS:
            raise ValueError("請上傳 .xlsx 或 .xlsm Excel 檔。")

        payload = schedule_file.getvalue()
        if not payload:
            raise ValueError("上傳檔案是空的，請重新選擇 Excel。")

        referees = [{"name": name} for name in parse_referee_names(referee_names)]
        if not referees:
            raise ValueError("請至少輸入一位裁判。")

        create_referee_only_artifacts = getattr(tournament_tools, "create_referee_only_artifacts", None)
        if create_referee_only_artifacts is None:
            raise ValueError("線上版本尚未同步到既有賽程裁判模式，請先在 Streamlit 按 Reboot / Redeploy。")

        draw_data, artifacts = create_referee_only_artifacts(
            io.BytesIO(payload),
            referees,
            source_file=source_name,
            config=config,
            base_dir=BASE_DIR,
        )
        st.session_state["existing_schedule_template_bytes"] = payload
        st.session_state["existing_schedule_template_drawn_at"] = draw_data.get("drawn_at", "")
        st.session_state["referee_names_text"] = "\n".join(referee["name"] for referee in draw_data.get("referees", []))
        st.session_state["referee_names_drawn_at"] = draw_data.get("drawn_at", "")
        if artifacts.latest_sync_complete:
            st.session_state["message"] = ("success", "已讀取既有賽程並產生裁判表，右側可繼續微調裁判設定。")
        else:
            st.session_state["message"] = ("warning", "已讀取既有賽程，但 latest 檔案同步可能被 Excel 開啟中擋住。")
        st.rerun()
    except Exception as exc:
        st.session_state["message"] = ("error", str(exc))
        st.rerun()


def run_draw(
    *,
    uploaded_file: Any,
    group_count: int,
    advance_per_group: int,
    wildcard_count: int,
    knockout_format: str,
    download_options: dict[str, bool],
    day1_latest_end: str,
    day2_latest_end: str,
) -> None:
    try:
        registration_source, source_file = get_registration_source(uploaded_file)
        draw_data, artifacts = create_draw_artifacts(
            BASE_DIR,
            registration_source=registration_source,
            source_file=source_file,
            group_count=group_count,
            advance_per_group=advance_per_group,
            wildcard_count=wildcard_count,
            knockout_format=knockout_format,
            download_options=download_options,
            day1_latest_end=day1_latest_end,
            day2_latest_end=day2_latest_end,
        )

        outputs = selected_output_labels(download_options)
        if draw_data.get("schedule", {}).get("status") == "scheduled":
            st.session_state["message"] = ("success", f"抽籤完成，賽程可排入目前限制。已產生：{outputs}。")
        else:
            st.session_state["message"] = ("warning", f"抽籤完成，但賽程需要調整限制。已產生：{outputs}。")

        if not artifacts.latest_sync_complete:
            st.session_state["message"] = ("warning", "抽籤完成，但 latest 資料夾有檔案未能更新；請確認沒有檔案被開啟。")
        st.rerun()
    except Exception as exc:
        st.session_state["message"] = ("error", str(exc))
        st.rerun()


def get_registration_source(uploaded_file: Any) -> tuple[Path | io.BytesIO, str]:
    if uploaded_file is None:
        registration_path = resolve_registration_path(BASE_DIR)
        return registration_path, registration_path.name

    source_name = Path(PureWindowsPath(uploaded_file.name).name).name
    extension = Path(source_name).suffix.lower()
    if extension not in ALLOWED_UPLOAD_EXTENSIONS:
        raise ValueError("請上傳 .xlsx 或 .xlsm Excel 檔。")

    payload = uploaded_file.getvalue()
    if not payload:
        raise ValueError("上傳檔案是空的，請重新選擇 Excel。")
    return io.BytesIO(payload), source_name


def render_groups(groups: dict[str, list[str]]) -> None:
    cards = []
    for group_name, members in groups.items():
        members_html = "".join(f"<li>{team}</li>" for team in members)
        cards.append(f'<div class="group-card"><h3>{group_name}</h3><ol>{members_html}</ol></div>')
    st.markdown(f'<div class="group-grid">{"".join(cards)}</div>', unsafe_allow_html=True)


def render_team_chips(teams: list[str]) -> None:
    chips = "".join(f'<div class="team-chip">{team}</div>' for team in teams)
    st.markdown(f'<div class="team-list">{chips}</div>', unsafe_allow_html=True)


def load_fallback_teams() -> tuple[list[str], str | None]:
    try:
        registration_path = resolve_registration_path(BASE_DIR)
        return load_teams(registration_path), None
    except Exception as exc:
        return [], str(exc)


def build_download_items(latest_draw: dict[str, Any]) -> list[dict[str, Any]]:
    labels = {
        "json": "下載 JSON",
        "schedule": "下載 Excel",
        "pdf": "下載 PDF",
    }
    mime_types = {
        "json": "application/json",
        "schedule": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "pdf": "application/pdf",
    }
    artifact_filenames = get_artifact_filenames(BASE_DIR)
    download_options = normalize_download_options(latest_draw.get("download_options"))
    latest_dir = BASE_DIR / "outputs" / "latest"
    items: list[dict[str, Any]] = []
    for kind in ("json", "schedule", "pdf"):
        path = latest_dir / artifact_filenames[kind]
        if download_options[kind] and path.exists():
            items.append({"kind": kind, "label": labels[kind], "path": path, "mime": mime_types[kind]})
    return items


def selected_output_labels(download_options: dict[str, bool]) -> str:
    labels = []
    if download_options["json"]:
        labels.append("JSON")
    if download_options["schedule"]:
        labels.append("Excel")
    if download_options["pdf"]:
        labels.append("PDF")
    return "、".join(labels) if labels else "不產生下載檔"


def option_index(options: list[str], selected: str) -> int:
    try:
        return options.index(selected)
    except ValueError:
        return 0


def label_for_value(label_to_value: dict[str, str], selected_value: str) -> str:
    for label, value in label_to_value.items():
        if value == selected_value:
            return label
    return next(iter(label_to_value))


if __name__ == "__main__":
    main()
