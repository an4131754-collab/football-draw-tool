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
    generate_artifacts,
    get_artifact_filenames,
    get_latest_draw_data,
    load_config,
    load_teams,
    normalize_download_options,
    resolve_registration_path,
    update_draw_referees,
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
  <title>Interdept Cup Draw Studio</title>
  <style>
    :root {
      --bg: #050505;
      --bg-soft: #0d0d0d;
      --ink: #f5f2ea;
      --muted: #9b978f;
      --line: rgba(245, 242, 234, 0.14);
      --line-strong: rgba(245, 242, 234, 0.26);
      --panel: rgba(15, 15, 15, 0.72);
      --panel-solid: #101010;
      --accent: #d6ff63;
      --accent-2: #9ae6ff;
      --danger: #ff765f;
      --shadow: 0 34px 110px rgba(0, 0, 0, 0.5);
      --font-display: "Bahnschrift", "Arial Narrow", "Microsoft JhengHei", sans-serif;
      --font-ui: "Microsoft JhengHei", "Noto Sans TC", sans-serif;
    }

    * { box-sizing: border-box; }

    html { scroll-behavior: smooth; }

    body {
      margin: 0;
      min-height: 100vh;
      color: var(--ink);
      font-family: var(--font-ui);
      background: var(--bg);
      overflow-x: hidden;
    }

    body::before,
    body::after {
      content: "";
      position: fixed;
      inset: -20%;
      z-index: -3;
      pointer-events: none;
    }

    body::before {
      background:
        radial-gradient(circle at 16% 12%, rgba(214, 255, 99, 0.18), transparent 22%),
        radial-gradient(circle at 86% 6%, rgba(154, 230, 255, 0.13), transparent 24%),
        radial-gradient(circle at 72% 78%, rgba(255, 255, 255, 0.08), transparent 28%),
        linear-gradient(135deg, #050505 0%, #0e0e0e 46%, #020202 100%);
      animation: ambientShift 16s ease-in-out infinite alternate;
    }

    body::after {
      opacity: 0.22;
      background-image:
        linear-gradient(rgba(255,255,255,0.08) 1px, transparent 1px),
        linear-gradient(90deg, rgba(255,255,255,0.07) 1px, transparent 1px);
      background-size: 84px 84px;
      transform: perspective(820px) rotateX(64deg) translateY(-16%);
      transform-origin: top;
      animation: gridDrift 18s linear infinite;
    }

    a { color: inherit; }

    .grain {
      position: fixed;
      inset: 0;
      z-index: -1;
      pointer-events: none;
      opacity: 0.23;
      background-image:
        repeating-radial-gradient(circle at 8% 18%, rgba(255,255,255,0.15) 0 1px, transparent 1px 4px);
      mix-blend-mode: overlay;
    }

    .runway {
      position: fixed;
      inset: auto -10% 0;
      height: 36vh;
      z-index: -2;
      pointer-events: none;
      background:
        linear-gradient(90deg, transparent 0 48%, rgba(214,255,99,0.18) 49%, transparent 51% 100%),
        linear-gradient(to top, rgba(214,255,99,0.14), transparent 72%);
      clip-path: polygon(40% 0, 60% 0, 100% 100%, 0 100%);
      filter: blur(0.2px);
      opacity: 0.9;
    }

    .shell {
      width: min(1180px, calc(100% - 34px));
      margin: 0 auto;
      padding: 22px 0 72px;
    }

    .nav {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 18px;
      min-height: 72px;
    }

    .brand {
      display: inline-grid;
      gap: 2px;
      text-decoration: none;
      letter-spacing: 0.24em;
      text-transform: uppercase;
      font-family: var(--font-display);
      font-size: 0.82rem;
      font-weight: 800;
    }

    .brand span:last-child {
      color: var(--muted);
      font-size: 0.72rem;
      letter-spacing: 0.34em;
    }

    .nav-links {
      display: flex;
      align-items: center;
      gap: 22px;
      color: var(--muted);
      font-size: 0.76rem;
      letter-spacing: 0.18em;
      text-transform: uppercase;
    }

    .nav-links a {
      text-decoration: none;
      transition: color 180ms ease;
    }

    .nav-links a:hover { color: var(--ink); }

    .hero {
      min-height: 72vh;
      display: grid;
      grid-template-columns: minmax(0, 1.05fr) minmax(300px, 0.78fr);
      gap: clamp(28px, 5vw, 74px);
      align-items: center;
      padding: clamp(42px, 7vw, 92px) 0;
    }

    .hero-copy {
      display: grid;
      gap: 24px;
      animation: riseIn 760ms ease both;
    }

    h1 {
      margin: 0;
      max-width: 820px;
      font-family: var(--font-display);
      font-size: clamp(3.8rem, 12vw, 9.8rem);
      line-height: 0.84;
      letter-spacing: -0.08em;
      text-transform: uppercase;
    }

    .lead {
      max-width: 620px;
      margin: 0;
      color: var(--muted);
      font-size: clamp(1rem, 1.7vw, 1.2rem);
      line-height: 1.9;
    }

    .hero-actions,
    .actions,
    .downloads {
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      align-items: center;
    }

    .button,
    button,
    .button-link {
      position: relative;
      isolation: isolate;
      border: 1px solid var(--line-strong);
      border-radius: 999px;
      min-height: 48px;
      padding: 13px 22px;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 10px;
      color: var(--ink);
      background: rgba(255, 255, 255, 0.055);
      font: 800 0.78rem/1 var(--font-ui);
      letter-spacing: 0.18em;
      text-transform: uppercase;
      text-decoration: none;
      cursor: pointer;
      overflow: hidden;
      transition: transform 180ms ease, border-color 180ms ease, background 180ms ease;
    }

    .button::before,
    button::before,
    .button-link::before {
      content: "";
      position: absolute;
      inset: 0;
      z-index: -1;
      background: linear-gradient(110deg, transparent, rgba(255,255,255,0.18), transparent);
      transform: translateX(-120%);
      transition: transform 520ms ease;
    }

    .button:hover,
    button:hover,
    .button-link:hover {
      transform: translateY(-2px);
      border-color: rgba(214,255,99,0.64);
      background: rgba(255,255,255,0.09);
    }

    .button:hover::before,
    button:hover::before,
    .button-link:hover::before {
      transform: translateX(120%);
    }

    .button-primary,
    button[type="submit"] {
      color: #080808;
      border-color: transparent;
      background: var(--accent);
      box-shadow: 0 18px 46px rgba(214,255,99,0.22);
    }

    .button-secondary {
      color: #fff;
      background: rgba(255, 118, 95, 0.16);
      border-color: rgba(255, 118, 95, 0.36);
    }

    .hero-card {
      position: relative;
      min-height: 520px;
      border: 1px solid var(--line);
      border-radius: 34px;
      overflow: hidden;
      background:
        linear-gradient(145deg, rgba(255,255,255,0.12), rgba(255,255,255,0.02)),
        radial-gradient(circle at 50% 20%, rgba(214,255,99,0.22), transparent 28%),
        #101010;
      box-shadow: var(--shadow);
      animation: floatCard 7s ease-in-out infinite;
    }

    .hero-card::before {
      content: "";
      position: absolute;
      inset: 8%;
      border-radius: 999px;
      border: 1px solid rgba(255,255,255,0.16);
      background:
        radial-gradient(circle, transparent 0 36%, rgba(214,255,99,0.08) 37% 38%, transparent 39%),
        conic-gradient(from 180deg, transparent, rgba(214,255,99,0.22), transparent, rgba(154,230,255,0.18), transparent);
      filter: blur(0.2px);
      animation: spinSlow 18s linear infinite;
    }

    .hero-card::after {
      content: "SYSTEM RANDOM";
      position: absolute;
      left: 28px;
      right: 28px;
      bottom: 28px;
      padding: 22px;
      border: 1px solid rgba(255,255,255,0.18);
      border-radius: 22px;
      background: rgba(0,0,0,0.52);
      backdrop-filter: blur(18px);
      font: 800 1.65rem/1 var(--font-display);
      letter-spacing: -0.04em;
    }

    .spec-strip {
      position: absolute;
      top: 28px;
      left: 28px;
      right: 28px;
      display: flex;
      justify-content: space-between;
      gap: 10px;
      color: var(--muted);
      font-size: 0.72rem;
      letter-spacing: 0.18em;
      text-transform: uppercase;
    }

    .ball-orbit {
      position: absolute;
      left: 50%;
      top: 48%;
      width: 164px;
      height: 164px;
      transform: translate(-50%, -50%);
      border-radius: 50%;
      background:
        radial-gradient(circle at 38% 30%, #fff 0 7%, transparent 8%),
        radial-gradient(circle at 60% 62%, rgba(214,255,99,0.92) 0 9%, transparent 10%),
        linear-gradient(145deg, #f7f3e7, #676767 58%, #111);
      box-shadow: 0 28px 82px rgba(0,0,0,0.54), 0 0 70px rgba(214,255,99,0.28);
    }

    .ticker {
      border-block: 1px solid var(--line);
      overflow: hidden;
      color: var(--muted);
      font-family: var(--font-display);
      font-size: clamp(1.4rem, 4vw, 3.8rem);
      line-height: 1;
      letter-spacing: -0.05em;
      text-transform: uppercase;
      white-space: nowrap;
      padding: 18px 0;
    }

    .ticker-track {
      display: inline-flex;
      gap: 34px;
      min-width: max-content;
      animation: marquee 24s linear infinite;
    }

    .workspace {
      display: grid;
      grid-template-columns: minmax(320px, 0.86fr) minmax(0, 1.14fr);
      gap: 18px;
      padding-top: 24px;
    }

    .panel {
      border: 1px solid var(--line);
      border-radius: 30px;
      background: var(--panel);
      box-shadow: var(--shadow);
      backdrop-filter: blur(22px);
      overflow: hidden;
    }

    .panel-inner { padding: clamp(20px, 3vw, 30px); }

    .panel-head {
      display: flex;
      justify-content: space-between;
      gap: 18px;
      align-items: flex-start;
      margin-bottom: 22px;
    }

    .panel-kicker {
      margin: 0 0 8px;
      color: var(--accent);
      font-size: 0.72rem;
      letter-spacing: 0.2em;
      text-transform: uppercase;
      font-weight: 900;
    }

    h2 {
      margin: 0;
      font-family: var(--font-display);
      font-size: clamp(1.9rem, 3vw, 3.2rem);
      letter-spacing: -0.055em;
      line-height: 0.96;
      text-transform: uppercase;
    }

    .note,
    .hint {
      margin: 0;
      color: var(--muted);
      line-height: 1.7;
      font-size: 0.92rem;
    }

    .form-grid {
      display: grid;
      gap: 14px;
    }

    label,
    .field-label {
      display: grid;
      gap: 8px;
      color: var(--ink);
      font-size: 0.76rem;
      font-weight: 900;
      letter-spacing: 0.14em;
      text-transform: uppercase;
    }

    input[type="file"],
    input[type="number"],
    textarea,
    select {
      width: 100%;
      min-height: 52px;
      border: 1px solid var(--line);
      border-radius: 18px;
      background: rgba(255,255,255,0.055);
      color: var(--ink);
      padding: 13px 15px;
      font: 700 1rem/1.2 var(--font-ui);
      outline: none;
      transition: border-color 180ms ease, background 180ms ease, box-shadow 180ms ease;
    }

    select option {
      color: #101010;
      background: #f5f2ea;
    }

    input[type="file"] { cursor: pointer; }

    input:focus,
    textarea:focus,
    select:focus {
      border-color: rgba(214,255,99,0.66);
      background: rgba(255,255,255,0.08);
      box-shadow: 0 0 0 4px rgba(214,255,99,0.1);
    }

    textarea {
      min-height: 132px;
      resize: vertical;
      line-height: 1.55;
      text-transform: none;
      letter-spacing: 0;
    }

    .checkbox-row {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 8px;
    }

    .checkbox-row label {
      display: flex;
      min-height: 48px;
      align-items: center;
      justify-content: center;
      gap: 8px;
      border: 1px solid var(--line);
      border-radius: 999px;
      background: rgba(255,255,255,0.05);
      cursor: pointer;
      transition: border-color 180ms ease, background 180ms ease, transform 180ms ease;
    }

    .checkbox-row label:hover {
      transform: translateY(-1px);
      border-color: rgba(214,255,99,0.54);
      background: rgba(214,255,99,0.1);
    }

    .checkbox-row input { accent-color: var(--accent); }

    .referee-list {
      display: grid;
      gap: 14px;
      margin-top: 18px;
    }

    .referee-card {
      border: 1px solid var(--line);
      border-radius: 20px;
      padding: 14px;
      background: rgba(255,255,255,0.045);
    }

    .referee-card h3 {
      margin: 0 0 12px;
      font-size: 1rem;
      text-transform: uppercase;
    }

    .match-checks {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 8px;
      max-height: 190px;
      overflow: auto;
      padding-right: 4px;
    }

    .match-checks label {
      min-height: auto;
      padding: 8px 10px;
      border: 1px solid var(--line);
      border-radius: 14px;
      background: rgba(255,255,255,0.035);
      font-size: 0.68rem;
      line-height: 1.35;
      letter-spacing: 0.06em;
      text-transform: none;
    }

    .messages {
      display: grid;
      gap: 10px;
      margin: 0 0 18px;
    }

    .message {
      border: 1px solid rgba(214,255,99,0.28);
      border-radius: 18px;
      padding: 13px 16px;
      background: rgba(214,255,99,0.09);
      color: var(--ink);
      line-height: 1.6;
      backdrop-filter: blur(16px);
    }

    .message.error {
      border-color: rgba(255,118,95,0.36);
      background: rgba(255,118,95,0.11);
    }

    .message.success {
      border-color: rgba(214,255,99,0.36);
    }

    .team-list {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 8px;
      list-style: none;
      margin: 16px 0 0;
      padding: 0;
    }

    .team-list li,
    .mini-stat,
    .meta-card {
      border: 1px solid var(--line);
      border-radius: 16px;
      background: rgba(255,255,255,0.045);
      padding: 11px 12px;
    }

    .meta {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 10px;
      margin: 18px 0;
    }

    .meta-card {
      color: var(--muted);
      min-height: 84px;
    }

    .meta-card strong {
      display: block;
      margin-bottom: 8px;
      color: var(--ink);
      font: 900 2rem/0.9 var(--font-display);
      letter-spacing: -0.06em;
    }

    .status {
      border: 1px solid rgba(214,255,99,0.26);
      border-radius: 20px;
      padding: 16px;
      background:
        linear-gradient(135deg, rgba(214,255,99,0.12), rgba(255,255,255,0.035));
      color: var(--muted);
      line-height: 1.7;
      margin: 16px 0;
    }

    .status strong {
      display: block;
      color: var(--ink);
      margin-bottom: 6px;
    }

    .status.infeasible {
      border-color: rgba(255,118,95,0.34);
      background: rgba(255,118,95,0.1);
    }

    .groups-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 12px;
      margin: 18px 0;
    }

    .group-card {
      position: relative;
      min-height: 220px;
      border: 1px solid var(--line);
      border-radius: 24px;
      padding: 18px;
      overflow: hidden;
      background:
        linear-gradient(145deg, rgba(255,255,255,0.09), rgba(255,255,255,0.03)),
        var(--panel-solid);
      transition: transform 220ms ease, border-color 220ms ease, background 220ms ease;
    }

    .group-card::after {
      content: "";
      position: absolute;
      right: -42px;
      bottom: -42px;
      width: 150px;
      height: 150px;
      border-radius: 50%;
      background: radial-gradient(circle, rgba(214,255,99,0.18), transparent 66%);
      opacity: 0;
      transition: opacity 220ms ease;
    }

    .group-card:hover {
      transform: translateY(-4px);
      border-color: rgba(214,255,99,0.48);
      background: linear-gradient(145deg, rgba(255,255,255,0.12), rgba(255,255,255,0.04));
    }

    .group-card:hover::after { opacity: 1; }

    .group-card h3 {
      margin: 0 0 18px;
      font: 900 clamp(2rem, 5vw, 4.4rem)/0.8 var(--font-display);
      letter-spacing: -0.08em;
      text-transform: uppercase;
    }

    .group-card ol {
      position: relative;
      z-index: 1;
      display: grid;
      gap: 9px;
      margin: 0;
      padding: 0;
      list-style: none;
      counter-reset: teams;
    }

    .group-card li {
      counter-increment: teams;
      display: grid;
      grid-template-columns: 32px 1fr;
      gap: 10px;
      align-items: center;
      color: var(--ink);
    }

    .group-card li::before {
      content: counter(teams, decimal-leading-zero);
      color: var(--muted);
      font-size: 0.74rem;
      letter-spacing: 0.12em;
    }

    .empty-state {
      min-height: 420px;
      display: grid;
      place-items: center;
      text-align: center;
      padding: 34px;
      background:
        radial-gradient(circle at 50% 24%, rgba(214,255,99,0.16), transparent 26%),
        linear-gradient(145deg, rgba(255,255,255,0.08), rgba(255,255,255,0.02));
    }

    .empty-state p { max-width: 520px; }

    code {
      color: var(--accent);
      font-family: "Cascadia Mono", Consolas, monospace;
      font-size: 0.9em;
    }

    hr {
      border: 0;
      border-top: 1px solid var(--line);
      margin: 24px 0;
    }

    @keyframes ambientShift {
      from { transform: translate3d(-1%, -1%, 0) scale(1); }
      to { transform: translate3d(1.5%, 1%, 0) scale(1.05); }
    }

    @keyframes gridDrift {
      from { background-position: 0 0; }
      to { background-position: 0 168px; }
    }

    @keyframes riseIn {
      from { opacity: 0; transform: translateY(22px); }
      to { opacity: 1; transform: translateY(0); }
    }

    @keyframes floatCard {
      0%, 100% { transform: translateY(0) rotate(-1deg); }
      50% { transform: translateY(-12px) rotate(1deg); }
    }

    @keyframes spinSlow {
      to { transform: rotate(360deg); }
    }

    @keyframes marquee {
      to { transform: translateX(-50%); }
    }

    @media (prefers-reduced-motion: reduce) {
      *, *::before, *::after {
        animation-duration: 1ms !important;
        animation-iteration-count: 1 !important;
        scroll-behavior: auto !important;
        transition-duration: 1ms !important;
      }
    }

    @media (max-width: 940px) {
      .hero,
      .workspace {
        grid-template-columns: 1fr;
      }

      .hero-card {
        min-height: 420px;
      }
    }

    @media (max-width: 680px) {
      .shell { width: min(100% - 22px, 1180px); padding-bottom: 42px; }
      .nav { align-items: flex-start; }
      .nav-links { display: none; }
      .hero { min-height: auto; padding: 42px 0; }
      h1 { font-size: clamp(3.2rem, 18vw, 5.2rem); }
      .panel-head,
      .meta,
      .groups-grid,
      .team-list,
      .checkbox-row {
        grid-template-columns: 1fr;
      }
      .panel-head { display: grid; }
      .button,
      button,
      .button-link {
        width: 100%;
      }
      .hero-card { min-height: 360px; border-radius: 26px; }
    }
  </style>
</head>
<body>
  <div class="grain"></div>
  <div class="runway"></div>
  <main class="shell">
    <nav class="nav" aria-label="Main navigation">
      <a class="brand" href="#top">
        <span>Interdept Cup</span>
        <span>Draw Studio</span>
      </a>
      <div class="nav-links">
        <a href="#draw">Draw</a>
        <a href="#results">Results</a>
        <a href="#outputs">Outputs</a>
      </div>
    </nav>

    <section class="hero" id="top">
      <div class="hero-copy">
        <h1>Draw the cup. Keep the trust.</h1>
        <p class="lead">
          一個給足球系際盃使用的抽籤與賽程工作室。上傳同格式 Google 表單 Excel，
          選擇組數、晉級規則與輸出格式，再用 <code>secrets.SystemRandom().shuffle()</code>
          完成公開、可追溯的抽籤。
        </p>
        <div class="hero-actions">
          <a class="button button-primary" href="#draw">Start Draw</a>
          <a class="button" href="#results">Explore Results</a>
        </div>
      </div>
      <div class="hero-card" aria-hidden="true">
        <div class="spec-strip">
          <span>OS entropy</span>
          <span>Offline ready</span>
        </div>
        <div class="ball-orbit"></div>
      </div>
    </section>

    <section class="ticker" aria-hidden="true">
      <div class="ticker-track">
        <span>Groups</span><span>/</span><span>Schedule</span><span>/</span><span>PDF Proof</span><span>/</span><span>Excel Output</span><span>/</span>
        <span>Groups</span><span>/</span><span>Schedule</span><span>/</span><span>PDF Proof</span><span>/</span><span>Excel Output</span><span>/</span>
      </div>
    </section>

    {% with messages = get_flashed_messages(with_categories=true) %}
      {% if messages %}
        <div class="messages" style="margin-top: 24px;">
          {% for category, message in messages %}
            <div class="message {{ category }}">{{ message }}</div>
          {% endfor %}
        </div>
      {% endif %}
    {% endwith %}

    <section class="workspace">
      <div class="panel" id="draw">
        <div class="panel-inner">
          <div class="panel-head">
            <div>
              <p class="panel-kicker">Configure</p>
              <h2>Draw Setup</h2>
            </div>
            <p class="note">像選商品尺寸一樣設定抽籤規格。沒有上傳檔案時，會使用本機 fallback 報名表。</p>
          </div>

          <form method="post" action="{{ url_for('draw') }}" enctype="multipart/form-data">
            <div class="form-grid">
              <label>
                Upload Excel
                <input type="file" name="registration_file" accept=".xlsx,.xlsm">
                <span class="hint">支援 .xlsx / .xlsm，隊名欄位預設讀取「{{ config.team_column }}」。</span>
              </label>

              <label>
                Group Count
                <input type="number" name="group_count" min="2" max="{{ config.max_group_count }}" value="{{ defaults.group_count }}" required>
              </label>

              <label>
                Advance Per Group
                <input type="number" name="advance_per_group" min="1" value="{{ defaults.advance_per_group }}" required>
              </label>

              <label>
                Wildcard Teams
                <input type="number" name="wildcard_count" min="0" value="{{ defaults.wildcard_count }}" required>
                <span class="hint">例如最佳第二名 2 隊，就填 2。</span>
              </label>

              <label>
                Knockout Format
                <select name="knockout_format">
                  <option value="semifinal" {% if defaults.knockout_format == "semifinal" %}selected{% endif %}>Semifinal - 4 teams advance</option>
                  <option value="quarterfinal" {% if defaults.knockout_format == "quarterfinal" %}selected{% endif %}>Quarterfinal - 8 teams advance</option>
                </select>
              </label>

              <label>
                DAY1 Latest Finish
                <select name="day1_latest_end">
                  {% for value in day1_latest_end_options %}
                    <option value="{{ value }}" {% if value == defaults.day1_latest_end %}selected{% endif %}>{{ value }}</option>
                  {% endfor %}
                </select>
              </label>

              <label>
                DAY2 Latest Finish
                <select name="day2_latest_end">
                  {% for value in day2_latest_end_options %}
                    <option value="{{ value }}" {% if value == defaults.day2_latest_end %}selected{% endif %}>{{ value }}</option>
                  {% endfor %}
                </select>
              </label>

              <div>
                <div class="field-label" style="margin-bottom: 8px;">Output Pack</div>
                <div class="checkbox-row">
                  <label><input type="checkbox" name="generate_json" value="1" {% if defaults.generate_json %}checked{% endif %}> JSON</label>
                  <label><input type="checkbox" name="generate_excel" value="1" {% if defaults.generate_excel %}checked{% endif %}> Excel</label>
                  <label><input type="checkbox" name="generate_pdf" value="1" {% if defaults.generate_pdf %}checked{% endif %}> PDF</label>
                </div>
                <p class="hint" style="margin-top: 8px;">JSON 會作為抽籤紀錄；PDF 用來向參賽者說明隨機方式。</p>
              </div>
            </div>

            <div class="actions" style="margin-top: 22px;">
              <button type="submit">Start Draw</button>
              <span class="note">結果會更新 <code>outputs/latest</code>，舊版保留在 <code>outputs/archive</code>。</span>
            </div>
          </form>

          <hr>
          <div>
            <p class="panel-kicker">Local Fallback</p>
            <h2 style="font-size: clamp(1.5rem, 2.4vw, 2.3rem);">Team List</h2>
            {% if teams %}
              <p class="note" style="margin-top: 10px;">目前本機報名表讀到 {{ teams|length }} 隊。上傳檔案時，會以上傳檔案為準。</p>
              <ol class="team-list">
                {% for team in teams %}
                  <li>{{ team }}</li>
                {% endfor %}
              </ol>
            {% else %}
              <p class="note" style="margin-top: 10px;">{{ teams_error or "尚未讀到本機報名表，請直接上傳 Excel 後抽籤。" }}</p>
            {% endif %}
          </div>
        </div>
      </div>

      <div class="panel" id="results">
        {% if latest_draw %}
          <div class="panel-inner">
            <div class="panel-head">
              <div>
                <p class="panel-kicker">Latest Drop</p>
                <h2>Draw Results</h2>
              </div>
              <p class="note">抽籤時間：{{ latest_draw.drawn_at }}<br>亂數函數：<code>{{ latest_draw.random_function }}</code></p>
            </div>

            <div class="meta">
              <div class="meta-card"><strong>{{ latest_draw.team_count or latest_draw.teams|length }}</strong>Teams</div>
              <div class="meta-card"><strong>{{ latest_draw.group_count or latest_draw.groups|length }}</strong>Groups</div>
              <div class="meta-card"><strong>{{ latest_draw.advancement.total_advancers if latest_draw.advancement else "TBD" }}</strong>Advance</div>
            </div>

            {% if latest_draw.advancement %}
              <p class="note">晉級規則：{{ latest_draw.advancement.summary }}</p>
              <p class="note">淘汰賽格式：{{ latest_draw.advancement.knockout_stage or latest_draw.knockout_format }}</p>
            {% endif %}

            {% if latest_draw.schedule %}
              <div class="status {{ latest_draw.schedule.status }}">
                <strong>{{ "Schedule ready" if latest_draw.schedule.status == "scheduled" else "Schedule needs adjustment" }}</strong>
                {% for message in latest_draw.schedule.messages %}
                  <div>{{ message }}</div>
                {% endfor %}
                {% if latest_draw.schedule.status != "scheduled" %}
                  <div>可以放寬 DAY1/DAY2 最晚結束時間，或調整晉級隊數後重新抽籤。</div>
                {% endif %}
              </div>
            {% endif %}

            <div class="groups-grid">
              {% for group_name, members in latest_draw.groups.items() %}
                <div class="group-card">
                  <h3>{{ group_name }}</h3>
                  <ol>
                    {% for team in members %}
                      <li>{{ team }}</li>
                    {% endfor %}
                  </ol>
                </div>
              {% endfor %}
            </div>

            {% if latest_draw.schedule and latest_draw.schedule.status == "scheduled" %}
              <hr>
              <div id="referees">
                <p class="panel-kicker">Referee Setup</p>
                <h2 style="font-size: clamp(1.5rem, 2.4vw, 2.3rem);">Assign Officials</h2>
                <p class="note" style="margin-top: 10px;">每場安排 3 位裁判；同一時間不會把同一位裁判排到甲、乙兩場。若有隊伍歸屬，會避開該隊相關比賽。</p>
                {% if latest_draw.referee_warnings %}
                  <div class="status infeasible" style="margin-top: 14px;">
                    <strong>Referee warnings</strong>
                    {% for warning in latest_draw.referee_warnings %}
                      <div>{{ warning }}</div>
                    {% endfor %}
                  </div>
                {% elif latest_draw.referee_assignments %}
                  <div class="status" style="margin-top: 14px;">
                    <strong>Referee schedule ready</strong>
                    <div>Excel 下載檔已包含「裁判」工作表。</div>
                  </div>
                {% endif %}

                <form method="post" action="{{ url_for('referees') }}" style="margin-top: 18px;">
                  <label>
                    Referee Names
                    <textarea id="refereeNames" name="referee_names" placeholder="一行一位裁判">{{ referee_names_text }}</textarea>
                    <span class="hint">輸入姓名後，下方會自動產生每位裁判的所屬隊伍與不可排場次設定。</span>
                  </label>
                  <input type="hidden" id="refereeCount" name="referee_count" value="0">
                  <div class="referee-list" id="refereeRows"></div>
                  <div class="actions" style="margin-top: 18px;">
                    <button type="submit">Generate Referee Schedule</button>
                  </div>
                </form>
              </div>
            {% endif %}

            <div id="outputs">
              {% if download_items %}
                <div class="downloads">
                  {% for item in download_items %}
                    <a class="button-link" href="{{ url_for('download', kind=item.kind) }}">{{ item.label }}</a>
                  {% endfor %}
                </div>
              {% else %}
                <p class="note">這次沒有選擇下載輸出。系統仍會保留內部抽籤紀錄。</p>
              {% endif %}
            </div>

            <form class="actions" method="post" action="{{ url_for('clear') }}" style="margin-top: 18px;">
              <button class="button-secondary" type="submit">Clear Current Result</button>
              <span class="note">只清空 <code>outputs/latest</code>，不刪 archive。</span>
            </form>
          </div>
        {% else %}
          <div class="empty-state">
            <div>
              <p class="panel-kicker">No Draw Yet</p>
              <h2>Ready when the teams are.</h2>
              <p class="note" style="margin: 16px auto 0;">左側上傳 Excel 並開始抽籤後，這裡會顯示分組結果、排程狀態與下載連結。</p>
            </div>
          </div>
        {% endif %}
      </div>
    </section>
  </main>
  <script>
    const refereeTeams = {{ (latest_draw.teams if latest_draw else []) | tojson }};
    const refereeMatches = {{ referee_match_options | tojson }};
    const existingReferees = {{ existing_referees | tojson }};

    function escapeHtml(value) {
      return String(value).replace(/[&<>"']/g, char => ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;"
      })[char]);
    }

    function renderRefereeRows() {
      const namesInput = document.getElementById("refereeNames");
      const rows = document.getElementById("refereeRows");
      const countInput = document.getElementById("refereeCount");
      if (!namesInput || !rows || !countInput) return;

      const seen = new Set();
      const names = namesInput.value.split(/\r?\n/)
        .map(name => name.trim())
        .filter(name => {
          if (!name || seen.has(name)) return false;
          seen.add(name);
          return true;
        });
      countInput.value = names.length;

      rows.innerHTML = names.map((name, index) => {
        const existing = existingReferees[name] || {};
        const teamOptions = ["", ...refereeTeams].map(team => {
          const label = team || "No affiliation";
          const selected = (existing.affiliated_team || "") === team ? "selected" : "";
          return `<option value="${escapeHtml(team)}" ${selected}>${escapeHtml(label)}</option>`;
        }).join("");
        const unavailable = new Set((existing.unavailable_match_nos || []).map(String));
        const checks = refereeMatches.map(match => {
          const checked = unavailable.has(String(match.match_no)) ? "checked" : "";
          return `<label><input type="checkbox" name="unavailable_${index}" value="${match.match_no}" ${checked}> ${escapeHtml(match.label)}</label>`;
        }).join("");

        return `
          <div class="referee-card">
            <input type="hidden" name="referee_name_${index}" value="${escapeHtml(name)}">
            <h3>${escapeHtml(name)}</h3>
            <label>
              Affiliated Team
              <select name="affiliated_team_${index}">${teamOptions}</select>
            </label>
            <div class="field-label" style="margin-top: 12px;">Unavailable Matches</div>
            <div class="match-checks">${checks}</div>
          </div>
        `;
      }).join("");
    }

    document.addEventListener("DOMContentLoaded", () => {
      renderRefereeRows();
      const namesInput = document.getElementById("refereeNames");
      if (namesInput) namesInput.addEventListener("input", renderRefereeRows);
    });
  </script>
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
        "knockout_format": (
            latest_draw.get("knockout_format", config.get("default_knockout_format", "semifinal"))
            if latest_draw
            else config.get("default_knockout_format", "semifinal")
        ),
        "day1_latest_end": latest_constraints.get("day1_latest_end", config["default_day1_latest_end"]),
        "day2_latest_end": latest_constraints.get("day2_latest_end", config["default_day2_latest_end"]),
        "generate_json": latest_download_options["json"],
        "generate_excel": latest_download_options["schedule"],
        "generate_pdf": latest_download_options["pdf"],
    }

    return render_template_string(
        PAGE_TEMPLATE,
        config=config,
        defaults=defaults,
        latest_draw=latest_draw,
        download_items=build_download_items(latest_draw),
        referee_match_options=build_referee_match_options(latest_draw),
        existing_referees=build_existing_referees(latest_draw),
        referee_names_text=build_referee_names_text(latest_draw),
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
        knockout_format = request.form.get("knockout_format", "semifinal").strip() or "semifinal"
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
                knockout_format=knockout_format,
                download_options=download_options,
                day1_latest_end=day1_latest_end,
                day2_latest_end=day2_latest_end,
            )

        selected_outputs = selected_output_labels(download_options)
        if draw_data.get("schedule", {}).get("status") == "scheduled":
            flash(f"抽籤完成，賽程可排入目前限制。已產生：{selected_outputs}。", "success")
        else:
            flash(f"抽籤完成，但賽程需要調整限制。已產生：{selected_outputs}。", "error")

        if not artifacts.latest_sync_complete:
            flash("提醒：latest 資料夾有檔案可能正被開啟，部分檔案未能更新；請關閉 Excel/PDF 後再試一次。", "error")
    except Exception as exc:
        flash(str(exc), "error")

    return redirect(url_for("index"))


@app.post("/referees")
def referees() -> Any:
    try:
        latest_draw = get_latest_draw_data(BASE_DIR)
        if latest_draw is None:
            raise ValueError("尚未有抽籤結果，請先完成抽籤。")

        referee_payload = parse_referees_from_request()
        with STATE_LOCK:
            updated_draw = update_draw_referees(
                latest_draw,
                referee_payload,
                config=load_config(BASE_DIR),
            )
            artifacts = generate_artifacts(updated_draw, base_dir=BASE_DIR)

        if artifacts.latest_sync_complete:
            flash("裁判排班已完成，Excel 已更新「裁判」工作表。", "success")
        else:
            flash("裁判排班已完成，但 latest 檔案可能因 Excel 開啟中而無法同步。", "error")
    except Exception as exc:
        flash(str(exc), "error")

    return redirect(url_for("index") + "#referees")


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
        flash("這次抽籤沒有選擇產生這個下載檔。", "error")
        return redirect(url_for("index"))

    artifact_filenames = get_artifact_filenames(BASE_DIR)
    filename = artifact_filenames.get(kind)
    if filename is None:
        flash("不支援的下載類型。", "error")
        return redirect(url_for("index"))

    file_path = BASE_DIR / "outputs" / "latest" / filename
    if not file_path.exists():
        flash("找不到下載檔案，請先完成一次有勾選該輸出的抽籤。", "error")
        return redirect(url_for("index"))

    return send_file(file_path, as_attachment=True, download_name=file_path.name)


def build_download_items(latest_draw: dict[str, Any] | None) -> list[dict[str, str]]:
    if latest_draw is None:
        return []

    labels = {
        "json": "Download JSON",
        "schedule": "Download Excel",
        "pdf": "Download PDF",
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


def parse_referees_from_request() -> list[dict[str, Any]]:
    try:
        referee_count = int(request.form.get("referee_count", "0"))
    except ValueError as exc:
        raise ValueError("裁判數量格式錯誤。") from exc

    referees: list[dict[str, Any]] = []
    for index in range(referee_count):
        name = request.form.get(f"referee_name_{index}", "").strip()
        if not name:
            continue
        unavailable_values = request.form.getlist(f"unavailable_{index}")
        referees.append(
            {
                "name": name,
                "affiliated_team": request.form.get(f"affiliated_team_{index}", "").strip(),
                "unavailable_match_nos": unavailable_values,
            }
        )

    if not referees:
        names_text = request.form.get("referee_names", "")
        referees = [{"name": line.strip()} for line in names_text.splitlines() if line.strip()]
    return referees


def build_referee_match_options(latest_draw: dict[str, Any] | None) -> list[dict[str, Any]]:
    if latest_draw is None:
        return []
    options: list[dict[str, Any]] = []
    for match in latest_draw.get("schedule", {}).get("matches", []):
        match_no = int(match.get("match_no", 0))
        home = match.get("home", match.get("home_label", ""))
        away = match.get("away", match.get("away_label", ""))
        options.append(
            {
                "match_no": match_no,
                "label": f"#{match_no} {match.get('day', '')} {match.get('time', '')} {match.get('field', '')} - {home} vs {away}",
            }
        )
    return options


def build_existing_referees(latest_draw: dict[str, Any] | None) -> dict[str, dict[str, Any]]:
    if latest_draw is None:
        return {}
    return {item.get("name", ""): item for item in latest_draw.get("referees", []) if item.get("name")}


def build_referee_names_text(latest_draw: dict[str, Any] | None) -> str:
    if latest_draw is None:
        return ""
    return "\n".join(item.get("name", "") for item in latest_draw.get("referees", []) if item.get("name"))


def selected_output_labels(download_options: dict[str, bool]) -> str:
    labels = []
    if download_options["json"]:
        labels.append("JSON")
    if download_options["schedule"]:
        labels.append("Excel")
    if download_options["pdf"]:
        labels.append("PDF")
    return "、".join(labels) if labels else "不產生下載檔"


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
