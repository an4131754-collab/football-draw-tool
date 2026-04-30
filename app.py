from __future__ import annotations

import io
import os
import threading
from pathlib import Path, PureWindowsPath
from typing import Any

from flask import Flask, flash, redirect, render_template_string, request, send_file, url_for

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

HOST = "127.0.0.1"
PORT = 8000
ALLOWED_UPLOAD_EXTENSIONS = {".xlsx", ".xlsm"}
STATE_LOCK = threading.Lock()

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "local-tournament-draw-tool")
app.config["MAX_CONTENT_LENGTH"] = 12 * 1024 * 1024


PAGE_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>足球系際盃抽籤工具</title>
  <style>
    :root {
      --bg: #f4efe4;
      --panel: #fffdf8;
      --ink: #173026;
      --muted: #62756d;
      --line: #d9ceb8;
      --accent: #206d50;
      --accent-dark: #134935;
      --danger: #8a4738;
      --gold: #c38a2e;
      --shadow: rgba(31, 59, 45, 0.10);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Microsoft JhengHei", "Noto Sans TC", sans-serif;
      color: var(--ink);
      background:
        radial-gradient(circle at 8% 4%, rgba(255,255,255,0.9), transparent 26%),
        radial-gradient(circle at 92% 10%, rgba(32,109,80,0.14), transparent 24%),
        linear-gradient(140deg, #eee4d1 0%, #fbf7ed 46%, #e7f0e8 100%);
      min-height: 100vh;
    }
    .shell {
      max-width: 1160px;
      margin: 0 auto;
      padding: 42px 20px 58px;
    }
    h1 {
      margin: 0 0 10px;
      font-size: clamp(1.8rem, 4vw, 3rem);
      letter-spacing: 0.04em;
    }
    .lead {
      max-width: 880px;
      margin: 0 0 24px;
      color: var(--muted);
      line-height: 1.75;
    }
    .layout {
      display: grid;
      grid-template-columns: minmax(300px, 390px) 1fr;
      gap: 20px;
      align-items: start;
    }
    .panel {
      background: rgba(255, 253, 248, 0.94);
      border: 1px solid rgba(217, 206, 184, 0.85);
      border-radius: 22px;
      padding: 22px;
      box-shadow: 0 18px 44px var(--shadow);
      backdrop-filter: blur(10px);
    }
    .panel h2 {
      margin: 0 0 14px;
      font-size: 1.25rem;
    }
    .form-grid {
      display: grid;
      gap: 14px;
    }
    label {
      display: grid;
      gap: 7px;
      color: var(--muted);
      font-weight: 700;
      font-size: 0.95rem;
    }
    input[type="file"],
    input[type="number"],
    select {
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 14px;
      background: #fffaf0;
      color: var(--ink);
      padding: 11px 12px;
      font: inherit;
    }
    input[type="file"] { cursor: pointer; }
    .checkbox-row {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
    }
    .checkbox-row label {
      display: inline-flex;
      flex-direction: row;
      align-items: center;
      gap: 7px;
      border: 1px solid var(--line);
      border-radius: 999px;
      padding: 9px 13px;
      background: #fffaf0;
      color: var(--ink);
      font-weight: 700;
    }
    .hint,
    .note {
      color: var(--muted);
      font-size: 0.92rem;
      line-height: 1.6;
    }
    .actions {
      margin-top: 18px;
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      align-items: center;
    }
    button,
    .button-link {
      border: 0;
      border-radius: 999px;
      background: linear-gradient(135deg, var(--accent) 0%, var(--accent-dark) 100%);
      color: white;
      padding: 12px 22px;
      font-size: 1rem;
      font-weight: 800;
      cursor: pointer;
      box-shadow: 0 13px 25px rgba(32, 109, 80, 0.18);
      text-decoration: none;
      display: inline-flex;
      align-items: center;
      justify-content: center;
    }
    button:hover,
    .button-link:hover { transform: translateY(-1px); }
    .button-secondary {
      background: linear-gradient(135deg, var(--danger) 0%, #6d3026 100%);
      box-shadow: 0 13px 25px rgba(109, 48, 38, 0.16);
    }
    .messages {
      display: grid;
      gap: 10px;
      margin: 0 0 18px;
    }
    .message {
      padding: 12px 14px;
      border-radius: 14px;
      background: #fff7dd;
      border: 1px solid rgba(195, 138, 46, 0.35);
      color: #60461b;
      line-height: 1.55;
    }
    .message.error {
      background: #fff0eb;
      border-color: rgba(138, 71, 56, 0.35);
      color: #6d3026;
    }
    .team-list {
      margin: 12px 0 0;
      padding-left: 1.35rem;
      line-height: 1.8;
    }
    .meta {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 12px;
      margin: 14px 0 20px;
    }
    .meta-card {
      border: 1px solid var(--line);
      border-radius: 16px;
      padding: 12px;
      background: #fffaf0;
    }
    .meta-card strong {
      display: block;
      font-size: 1.2rem;
      color: var(--accent-dark);
    }
    .status {
      border-left: 4px solid var(--accent);
      padding: 12px 14px;
      background: #f6f4e9;
      border-radius: 14px;
      margin: 14px 0;
      line-height: 1.65;
    }
    .status.infeasible {
      border-left-color: var(--danger);
      background: #fff0eb;
    }
    .groups-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(180px, 1fr));
      gap: 16px;
      margin: 18px 0;
    }
    .group-card {
      border: 1px solid var(--line);
      border-radius: 18px;
      padding: 16px;
      background: linear-gradient(180deg, rgba(255,255,255,0.94), rgba(248,244,236,0.98));
    }
    .group-card h3 { margin: 0 0 10px; }
    .group-card ol {
      margin: 0;
      padding-left: 1.2rem;
      line-height: 1.8;
    }
    .downloads {
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      margin-top: 16px;
    }
    code {
      font-family: Consolas, "Courier New", monospace;
      font-size: 0.95em;
    }
    @media (max-width: 860px) {
      .layout,
      .meta,
      .groups-grid { grid-template-columns: 1fr; }
      .shell { padding-top: 28px; }
    }
  </style>
</head>
<body>
  <main class="shell">
    <h1>足球系際盃抽籤工具</h1>
    <p class="lead">
      上傳同格式 Google 表單回覆 Excel，設定組數、晉級規則、加開時段與下載項目後抽籤。
      亂數使用 <code>secrets.SystemRandom().shuffle()</code>，非 12 隊 4 組也會嘗試自動排小組賽與淘汰賽。
    </p>

    {% with messages = get_flashed_messages(with_categories=true) %}
      {% if messages %}
        <div class="messages">
          {% for category, message in messages %}
            <div class="message {{ category }}">{{ message }}</div>
          {% endfor %}
        </div>
      {% endif %}
    {% endwith %}

    <div class="layout">
      <section class="panel">
        <h2>建立抽籤</h2>
        <form method="post" action="{{ url_for('draw') }}" enctype="multipart/form-data">
          <div class="form-grid">
            <label>
              上傳報名表 Excel
              <input type="file" name="registration_file" accept=".xlsx,.xlsm">
              <span class="hint">不選檔案時，會使用資料夾內可找到的本機回覆表 fallback。</span>
            </label>
            <label>
              組數
              <input type="number" name="group_count" min="2" max="{{ config.max_group_count }}" value="{{ defaults.group_count }}" required>
            </label>
            <label>
              每組前 N 名晉級
              <input type="number" name="advance_per_group" min="1" value="{{ defaults.advance_per_group }}" required>
            </label>
            <label>
              最佳名次補位幾隊
              <input type="number" name="wildcard_count" min="0" value="{{ defaults.wildcard_count }}" required>
            </label>
            <label>
              DAY1 最晚踢到
              <select name="day1_latest_end">
                {% for value in day1_latest_end_options %}
                  <option value="{{ value }}" {% if value == defaults.day1_latest_end %}selected{% endif %}>{{ value }}</option>
                {% endfor %}
              </select>
            </label>
            <label>
              DAY2 最晚踢到
              <select name="day2_latest_end">
                {% for value in day2_latest_end_options %}
                  <option value="{{ value }}" {% if value == defaults.day2_latest_end %}selected{% endif %}>{{ value }}</option>
                {% endfor %}
              </select>
            </label>
            <div>
              <div class="note" style="font-weight:700;margin-bottom:8px;">本次提供下載</div>
              <div class="checkbox-row">
                <label><input type="checkbox" name="generate_json" value="1" {% if defaults.generate_json %}checked{% endif %}> JSON</label>
                <label><input type="checkbox" name="generate_excel" value="1" {% if defaults.generate_excel %}checked{% endif %}> Excel</label>
                <label><input type="checkbox" name="generate_pdf" value="1" {% if defaults.generate_pdf %}checked{% endif %}> PDF</label>
              </div>
              <span class="hint">抽籤後不會自動下載；只會顯示你勾選項目的下載按鈕。</span>
            </div>
          </div>
          <div class="actions">
            <button type="submit">開始抽籤</button>
            <span class="note">每次抽籤會更新 <code>outputs/latest</code>，舊結果仍保留在 <code>outputs/archive</code>。</span>
          </div>
        </form>

        <hr style="border:0;border-top:1px solid var(--line);margin:22px 0;">
        <h2>本機 fallback 名單</h2>
        {% if teams %}
          <p class="note">目前可從本機報名表讀到 {{ teams|length }} 隊。上傳檔案時會以你上傳的檔案為準。</p>
          <ol class="team-list">
            {% for team in teams %}
              <li>{{ team }}</li>
            {% endfor %}
          </ol>
        {% else %}
          <p class="note">{{ teams_error or "尚未讀到本機報名表，請直接上傳 Excel 後抽籤。" }}</p>
        {% endif %}
      </section>

      <section class="panel">
        {% if latest_draw %}
          <h2>目前抽籤結果</h2>
          <p class="note">抽籤時間：{{ latest_draw.drawn_at }}</p>
          <p class="note">亂數函數：<code>{{ latest_draw.random_function }}</code></p>
          <div class="meta">
            <div class="meta-card"><strong>{{ latest_draw.team_count or latest_draw.teams|length }}</strong>隊伍數</div>
            <div class="meta-card"><strong>{{ latest_draw.group_count or latest_draw.groups|length }}</strong>組數</div>
            <div class="meta-card"><strong>{{ latest_draw.advancement.total_advancers if latest_draw.advancement else "待定" }}</strong>晉級隊數</div>
          </div>
          {% if latest_draw.advancement %}
            <p>晉級規則：{{ latest_draw.advancement.summary }}</p>
          {% endif %}

          {% if latest_draw.schedule %}
            <div class="status {{ latest_draw.schedule.status }}">
              <strong>排程狀態：{{ "已排定" if latest_draw.schedule.status == "scheduled" else "排不下" }}</strong>
              {% for message in latest_draw.schedule.messages %}
                <div>{{ message }}</div>
              {% endfor %}
              {% if latest_draw.schedule.status != "scheduled" %}
                <div>可以把左邊 DAY1/DAY2 最晚結束時間往後選，再重新抽籤或調整組數/晉級隊數。</div>
              {% endif %}
            </div>
          {% endif %}

          <div class="groups-grid">
            {% for group_name, members in latest_draw.groups.items() %}
              <div class="group-card">
                <h3>{{ group_name }} 組</h3>
                <ol>
                  {% for team in members %}
                    <li>{{ team }}</li>
                  {% endfor %}
                </ol>
              </div>
            {% endfor %}
          </div>

          {% if download_items %}
            <div class="downloads">
              {% for item in download_items %}
                <a class="button-link" href="{{ url_for('download', kind=item.kind) }}">{{ item.label }}</a>
              {% endfor %}
            </div>
          {% else %}
            <p class="note">本次沒有勾選公開下載項目；結果仍已保存在系統內部 JSON 供目前頁面顯示。</p>
          {% endif %}

          <form class="actions" method="post" action="{{ url_for('clear') }}">
            <button class="button-secondary" type="submit">清空目前結果</button>
            <span class="note">只清空 <code>outputs/latest</code>，不刪 archive 備份。</span>
          </form>
        {% else %}
          <h2>尚未抽籤</h2>
          <p class="note">
            左邊上傳 Excel 並按下抽籤後，這裡會顯示分組、排程狀態與你勾選的下載按鈕。
          </p>
        {% endif %}
      </section>
    </div>
  </main>
</body>
</html>
"""


@app.get("/")
def index() -> str:
    config = load_config(BASE_DIR)
    teams, teams_error = load_fallback_teams()
    latest_draw = get_latest_draw_data(BASE_DIR)
    latest_schedule = latest_draw.get("schedule", {}) if latest_draw else {}
    latest_constraints = latest_schedule.get("constraints", {})
    latest_download_options = normalize_download_options(latest_draw.get("download_options") if latest_draw else None)
    defaults = {
        "group_count": latest_draw.get("group_count", config["default_group_count"]) if latest_draw else config["default_group_count"],
        "advance_per_group": (
            latest_draw.get("advancement", {}).get("advance_per_group", config["default_advance_per_group"])
            if latest_draw
            else config["default_advance_per_group"]
        ),
        "wildcard_count": (
            latest_draw.get("advancement", {}).get("wildcard_count", config["default_wildcard_count"])
            if latest_draw
            else config["default_wildcard_count"]
        ),
        "day1_latest_end": latest_constraints.get("day1_latest_end", config["default_day1_latest_end"]),
        "day2_latest_end": latest_constraints.get("day2_latest_end", config["default_day2_latest_end"]),
        "generate_json": latest_download_options["json"],
        "generate_excel": latest_download_options["schedule"],
        "generate_pdf": latest_download_options["pdf"],
    }
    download_items = build_download_items(latest_draw)

    return render_template_string(
        PAGE_TEMPLATE,
        config=config,
        defaults=defaults,
        latest_draw=latest_draw,
        download_items=download_items,
        day1_latest_end_options=config["latest_end_options"]["DAY1"],
        day2_latest_end_options=config["latest_end_options"]["DAY2"],
        teams=teams,
        teams_error=teams_error,
    )


@app.post("/draw")
def draw() -> Any:
    try:
        group_count = parse_int_field("group_count", "組數")
        advance_per_group = parse_int_field("advance_per_group", "每組晉級名額")
        wildcard_count = parse_int_field("wildcard_count", "最佳名次補位數")
        day1_latest_end = request.form.get("day1_latest_end", "").strip() or None
        day2_latest_end = request.form.get("day2_latest_end", "").strip() or None
        download_options = parse_download_options()
        registration_source, source_file = get_registration_source_from_request()

        with STATE_LOCK:
            draw_data, artifacts = create_draw_artifacts(
                BASE_DIR,
                registration_source=registration_source,
                source_file=source_file,
                group_count=group_count,
                advance_per_group=advance_per_group,
                wildcard_count=wildcard_count,
                download_options=download_options,
                day1_latest_end=day1_latest_end,
                day2_latest_end=day2_latest_end,
            )

        selected_outputs = selected_output_labels(download_options)
        if draw_data.get("schedule", {}).get("status") == "scheduled":
            flash(f"抽籤完成，賽程已排定。本次下載項目：{selected_outputs}。", "success")
        else:
            flash(f"抽籤完成，但目前時段排不下完整賽程。本次下載項目：{selected_outputs}。", "error")

        if not artifacts.latest_sync_complete:
            flash("提醒：latest 資料夾有檔案被開啟中，部分檔案可能未能更新；請關閉 Excel/PDF 後再抽一次。", "error")
    except Exception as exc:
        flash(str(exc), "error")

    return redirect(url_for("index"))


@app.post("/clear")
def clear() -> Any:
    with STATE_LOCK:
        completed = clear_latest_artifacts(BASE_DIR)

    if completed:
        flash("已清空目前結果。", "success")
    else:
        flash("部分 latest 檔案可能正被開啟，無法完全清空；請關閉後再試一次。", "error")
    return redirect(url_for("index"))


@app.get("/download")
def download() -> Any:
    kind = request.args.get("kind", "")
    latest_draw = get_latest_draw_data(BASE_DIR)
    if latest_draw is None:
        flash("找不到目前抽籤結果，請先完成一次抽籤。", "error")
        return redirect(url_for("index"))

    download_options = normalize_download_options(latest_draw.get("download_options"))
    if kind not in download_options or not download_options[kind]:
        flash("本次抽籤沒有勾選這個下載項目。", "error")
        return redirect(url_for("index"))

    artifact_filenames = get_artifact_filenames(BASE_DIR)
    filename = artifact_filenames.get(kind)
    if filename is None:
        flash("不支援的下載類型。", "error")
        return redirect(url_for("index"))

    file_path = BASE_DIR / "outputs" / "latest" / filename
    if not file_path.exists():
        flash("找不到下載檔案，請重新產生或勾選該輸出項目。", "error")
        return redirect(url_for("index"))

    return send_file(file_path, as_attachment=True, download_name=file_path.name)


def build_download_items(latest_draw: dict[str, Any] | None) -> list[dict[str, str]]:
    if latest_draw is None:
        return []

    labels = {
        "json": "下載 JSON",
        "schedule": "下載賽程 Excel",
        "pdf": "下載 PDF 說明",
    }
    artifact_filenames = get_artifact_filenames(BASE_DIR)
    download_options = normalize_download_options(latest_draw.get("download_options"))
    latest_dir = BASE_DIR / "outputs" / "latest"
    items: list[dict[str, str]] = []
    for kind in ("json", "schedule", "pdf"):
        if download_options[kind] and (latest_dir / artifact_filenames[kind]).exists():
            items.append({"kind": kind, "label": labels[kind]})
    return items


def load_fallback_teams() -> tuple[list[str], str | None]:
    try:
        registration_path = resolve_registration_path(BASE_DIR)
        return load_teams(registration_path), None
    except Exception as exc:
        return [], str(exc)


def parse_int_field(field_name: str, label: str) -> int:
    raw_value = request.form.get(field_name, "").strip()
    if raw_value == "":
        raise ValueError(f"請填寫{label}。")
    try:
        return int(raw_value)
    except ValueError as exc:
        raise ValueError(f"{label}必須是整數。") from exc


def parse_download_options() -> dict[str, bool]:
    return {
        "json": request.form.get("generate_json") == "1",
        "schedule": request.form.get("generate_excel") == "1",
        "pdf": request.form.get("generate_pdf") == "1",
    }


def selected_output_labels(download_options: dict[str, bool]) -> str:
    labels = []
    if download_options["json"]:
        labels.append("JSON")
    if download_options["schedule"]:
        labels.append("Excel")
    if download_options["pdf"]:
        labels.append("PDF")
    return "、".join(labels) if labels else "無公開下載"


def get_registration_source_from_request() -> tuple[Path | io.BytesIO | None, str | None]:
    uploaded_file = request.files.get("registration_file")
    if uploaded_file is None or uploaded_file.filename == "":
        registration_path = resolve_registration_path(BASE_DIR)
        return registration_path, registration_path.name

    source_name = Path(PureWindowsPath(uploaded_file.filename).name).name
    extension = Path(source_name).suffix.lower()
    if extension not in ALLOWED_UPLOAD_EXTENSIONS:
        raise ValueError("請上傳 .xlsx 或 .xlsm Excel 檔。")

    payload = uploaded_file.read()
    if not payload:
        raise ValueError("上傳檔案是空的，請重新選擇 Excel。")

    return io.BytesIO(payload), source_name


def main() -> int:
    port = int(os.environ.get("PORT", PORT))
    host = "0.0.0.0" if os.environ.get("PORT") else HOST
    print(f"抽籤網站已啟動：http://{HOST}:{port}")
    print("按 Ctrl+C 可結束網站。")
    app.run(host=host, port=port, debug=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
