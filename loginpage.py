import streamlit as st

# ─────────────────────────────────────────
#  PAGE CONFIG — MUST BE FIRST STREAMLIT CALL
# ─────────────────────────────────────────
st.set_page_config(
    page_title="FaceTrack",
    page_icon="🎭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────
#  ALL OTHER IMPORTS AFTER set_page_config
# ─────────────────────────────────────────
import streamlit.components.v1 as components
import datetime
import os
import math
import numpy as np
import requests
import pandas as pd
import time
import random
import math
import threading
# Optional heavy imports — wrapped so app doesn't crash if missing
try:
    import cv2
    import face_recognition
    FACE_RECOGNITION_AVAILABLE = True
except ImportError:
    FACE_RECOGNITION_AVAILABLE = False

try:
    from streamlit_webrtc import webrtc_streamer, VideoTransformerBase, RTCConfiguration
    import av
    WEBRTC_AVAILABLE = True
except ImportError:
    WEBRTC_AVAILABLE = False

try:
    from supabase import create_client, Client
    SUPABASE_AVAILABLE = True
    SUPABASE_IMPORT_ERROR = ""
except Exception as e:
    SUPABASE_AVAILABLE = False
    SUPABASE_IMPORT_ERROR = str(e)

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# ─────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────
# Try to load credentials from Streamlit Secrets (safe for deployment), otherwise use hardcoded defaults
try:
    SUPABASE_URL = st.secrets["SUPABASE_URL"]
    SUPABASE_KEY = st.secrets["SUPABASE_KEY"]
except (KeyError, FileNotFoundError, Exception):
    # No fallback — secrets must be configured in Streamlit Cloud or .streamlit/secrets.toml
    SUPABASE_URL = "https://jmjdbrqoilxkrtfhlmuw.supabase.co"
    SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImptamRicnFvaWx4a3J0ZmhsbXV3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQzNjY5NTgsImV4cCI6MjA4OTk0Mjk1OH0.ccQUa3UCk32QA992ZhNbNb3Rk0c_J22bOwIuhCQdvl0"

EXCEL_FOLDER = "attendance_exports"
EXCEL_FILE   = os.path.join(EXCEL_FOLDER, "attendance.xlsx")
ATTENDANCE_START = datetime.time(9, 0)
ATTENDANCE_ENDS  = datetime.time(23, 42)
LATE_TIME        = datetime.time(10, 30)
ABSENT_TIME      = datetime.time(19, 0)

INDIA_HOLIDAYS = {
datetime.date(2026, 1, 26),   # Republic Day
    datetime.date(2026, 3, 6),    # Holi
    datetime.date(2026, 4, 3),    # Good Friday
    datetime.date(2026, 4, 14),   # Ambedkar Jayanti
    datetime.date(2026, 8, 15),   # Independence Day
    datetime.date(2026, 10, 2),   # Gandhi Jayanti
    datetime.date(2026, 11, 8),   # Diwali
    datetime.date(2026, 11, 25),  # Guru Nanak Jayanti
    datetime.date(2026, 12, 25),  # Christmas
}

# ─────────────────────────────────────────
#  LANDING PAGE HTML
# ─────────────────────────────────────────
LANDING_PAGE_HTML = r'''
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>FaceTrack</title>
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:ital,wght@0,300;0,400;0,500;1,300&display=swap" rel="stylesheet"/>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  :root {
    --bg:#0d1210;--bg2:#111915;--surface:#161e19;--border:rgba(255,255,255,0.07);
    --accent:#e8603c;--accent2:#d4f0c2;--green:#3ecf6f;--text:#e8ede9;--muted:#7a9080;--card:#192118;
  }
  html{scroll-behavior:smooth;}
  body{background:var(--bg);color:var(--text);font-family:'DM Sans',sans-serif;font-size:15px;line-height:1.6;overflow-x:hidden;}
  nav{position:fixed;top:0;left:0;right:0;z-index:100;display:flex;align-items:center;justify-content:space-between;padding:0 40px;height:64px;background:rgba(13,18,16,0.85);backdrop-filter:blur(18px);border-bottom:1px solid var(--border);}
  .logo{display:flex;align-items:center;gap:10px;font-family:'Syne',sans-serif;font-weight:800;font-size:20px;}
  .logo-icon{width:36px;height:36px;border-radius:10px;background:linear-gradient(135deg,#e8603c,#c94020);display:flex;align-items:center;justify-content:center;font-size:18px;}
  .nav-links{display:flex;gap:32px;list-style:none;}
  .nav-links a{color:var(--muted);text-decoration:none;font-size:14px;transition:color .2s;}
  .nav-links a:hover{color:var(--text);}
  .nav-cta{background:var(--accent);color:#fff;border:none;padding:9px 22px;border-radius:8px;font-size:14px;font-weight:500;cursor:pointer;transition:opacity .2s;}
  .nav-cta:hover{opacity:.85;}
  .hero{min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;padding:120px 24px 80px;position:relative;overflow:hidden;}
  .hero-bg{position:absolute;inset:0;pointer-events:none;background:radial-gradient(ellipse 700px 500px at 50% 30%,rgba(232,96,60,0.12),transparent),radial-gradient(ellipse 500px 400px at 80% 80%,rgba(62,207,111,0.07),transparent);}
  .hero-badge{display:inline-flex;align-items:center;gap:8px;background:rgba(62,207,111,0.1);border:1px solid rgba(62,207,111,0.25);color:var(--green);padding:6px 16px;border-radius:100px;font-size:12px;font-weight:500;letter-spacing:.08em;text-transform:uppercase;margin-bottom:28px;}
  .hero-badge span{width:6px;height:6px;background:var(--green);border-radius:50%;animation:blink 1.5s ease-in-out infinite;}
  @keyframes blink{0%,100%{opacity:1}50%{opacity:.3}}
  h1{font-family:'Syne',sans-serif;font-size:clamp(42px,7vw,88px);font-weight:800;line-height:1.05;letter-spacing:-.03em;max-width:900px;margin-bottom:24px;}
  h1 em{color:var(--accent);font-style:normal;}
  .hero-sub{max-width:560px;color:var(--muted);font-size:17px;margin-bottom:44px;line-height:1.7;}
  .hero-btns{display:flex;gap:14px;flex-wrap:wrap;justify-content:center;}
  .btn-primary{background:var(--accent);color:#fff;border:none;padding:14px 32px;border-radius:10px;font-size:15px;font-weight:500;cursor:pointer;transition:transform .15s,box-shadow .15s;display:flex;align-items:center;gap:8px;}
  .btn-primary:hover{transform:translateY(-2px);box-shadow:0 8px 30px rgba(232,96,60,0.35);}
  .btn-outline{background:transparent;color:var(--text);border:1px solid var(--border);padding:14px 32px;border-radius:10px;font-size:15px;cursor:pointer;transition:border-color .2s,background .2s;}
  .btn-outline:hover{border-color:var(--muted);background:rgba(255,255,255,0.04);}
  .stats-bar{display:flex;gap:0;border-top:1px solid var(--border);border-bottom:1px solid var(--border);background:var(--surface);}
  .stat-item{flex:1;padding:32px 24px;text-align:center;border-right:1px solid var(--border);}
  .stat-item:last-child{border-right:none;}
  .stat-num{font-family:'Syne',sans-serif;font-size:40px;font-weight:800;color:var(--accent2);display:block;line-height:1.1;}
  .stat-label{color:var(--muted);font-size:12px;text-transform:uppercase;letter-spacing:.08em;margin-top:4px;}
  .camera-section{padding:100px 40px;max-width:1200px;margin:0 auto;}
  .section-label{font-size:11px;letter-spacing:.14em;text-transform:uppercase;color:var(--accent);margin-bottom:16px;display:flex;align-items:center;gap:8px;}
  .section-label::before{content:'';display:block;width:24px;height:1px;background:var(--accent);}
  .camera-grid{display:grid;grid-template-columns:1fr 1fr;gap:40px;align-items:center;}
  .camera-preview{background:var(--card);border:1px solid var(--border);border-radius:20px;overflow:hidden;position:relative;aspect-ratio:4/3;}
  .camera-topbar{position:absolute;top:0;left:0;right:0;display:flex;justify-content:space-between;align-items:center;padding:14px 20px;background:rgba(13,18,16,0.7);backdrop-filter:blur(8px);z-index:2;}
  .live-dot{display:flex;align-items:center;gap:8px;font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:var(--accent);}
  .live-dot span{width:7px;height:7px;background:var(--accent);border-radius:50%;animation:blink 1.5s infinite;}
  .enc-badge{font-size:11px;color:var(--green);display:flex;align-items:center;gap:6px;}
  .face-box-wrap{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#111915 0%,#0d1210 100%);}
  .face-box{width:180px;height:180px;border:2px solid rgba(62,207,111,0.6);border-radius:12px;position:relative;animation:facePulse 2.5s ease-in-out infinite;}
  @keyframes facePulse{0%,100%{border-color:rgba(62,207,111,0.6);box-shadow:0 0 0 0 rgba(62,207,111,0)}50%{border-color:rgba(62,207,111,1);box-shadow:0 0 0 12px rgba(62,207,111,0)}}
  .face-corner{position:absolute;width:20px;height:20px;border-color:var(--green);border-style:solid;}
  .fc-tl{top:-2px;left:-2px;border-width:3px 0 0 3px;border-radius:4px 0 0 0;}
  .fc-tr{top:-2px;right:-2px;border-width:3px 3px 0 0;border-radius:0 4px 0 0;}
  .fc-bl{bottom:-2px;left:-2px;border-width:0 0 3px 3px;border-radius:0 0 0 4px;}
  .fc-br{bottom:-2px;right:-2px;border-width:0 3px 3px 0;border-radius:0 0 4px 0;}
  .face-icon{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;}
  .face-icon svg{opacity:.5;}
  .match-badge{position:absolute;top:-28px;left:50%;transform:translateX(-50%);background:rgba(62,207,111,0.15);border:1px solid rgba(62,207,111,0.4);color:var(--green);font-size:10px;letter-spacing:.08em;text-transform:uppercase;padding:4px 12px;border-radius:6px;white-space:nowrap;}
  .camera-bottombar{position:absolute;bottom:0;left:0;right:0;display:flex;align-items:center;justify-content:space-between;padding:14px 20px;background:rgba(13,18,16,0.8);backdrop-filter:blur(8px);}
  .verified-tag{display:flex;align-items:center;gap:6px;font-size:12px;color:var(--green);}
  .scan-line{position:absolute;left:0;right:0;height:1px;background:linear-gradient(90deg,transparent,rgba(62,207,111,0.6),transparent);animation:scan 2.5s ease-in-out infinite;top:40%;}
  @keyframes scan{0%{top:20%;opacity:0}10%{opacity:1}90%{opacity:1}100%{top:80%;opacity:0}}
  .camera-text h2{font-family:'Syne',sans-serif;font-size:clamp(28px,4vw,44px);font-weight:800;line-height:1.1;margin-bottom:20px;}
  .camera-text h2 em{color:var(--accent);font-style:normal;}
  .camera-text p{color:var(--muted);line-height:1.8;margin-bottom:32px;}
  .feature-chips{display:flex;flex-wrap:wrap;gap:10px;}
  .chip{background:var(--surface);border:1px solid var(--border);padding:7px 14px;border-radius:8px;font-size:13px;color:var(--text);display:flex;align-items:center;gap:6px;}
  .chip-dot{width:6px;height:6px;border-radius:50%;}
  .features-section{padding:80px 40px;background:var(--surface);border-top:1px solid var(--border);}
  .section-header{text-align:center;margin-bottom:64px;}
  .section-header small{font-size:11px;text-transform:uppercase;letter-spacing:.12em;color:var(--accent);}
  .section-header h2{font-family:'Syne',sans-serif;font-size:clamp(26px,4vw,42px);font-weight:800;margin-top:12px;line-height:1.15;}
  .section-header h2 em{color:var(--accent2);font-style:normal;}
  .section-header p{color:var(--muted);max-width:480px;margin:14px auto 0;}
  .features-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:1px;background:var(--border);max-width:1200px;margin:0 auto;border:1px solid var(--border);border-radius:16px;overflow:hidden;}
  .feat-card{background:var(--card);padding:36px 32px;transition:background .2s;}
  .feat-card:hover{background:var(--surface);}
  .feat-icon{width:48px;height:48px;border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:22px;margin-bottom:20px;}
  .feat-card h3{font-family:'Syne',sans-serif;font-size:18px;font-weight:700;margin-bottom:10px;}
  .feat-card p{color:var(--muted);font-size:14px;line-height:1.7;}
  .dashboard-section{padding:100px 40px;max-width:1200px;margin:0 auto;}
  .dash-grid{display:grid;grid-template-columns:1fr 1fr;gap:40px;align-items:center;}
  .dash-preview{background:var(--card);border:1px solid var(--border);border-radius:20px;padding:24px;overflow:hidden;}
  .dash-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:24px;padding-bottom:16px;border-bottom:1px solid var(--border);}
  .dash-title{font-family:'Syne',sans-serif;font-weight:700;font-size:15px;}
  .dash-stats{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:20px;}
  .dash-stat{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:16px;}
  .ds-val{font-family:'Syne',sans-serif;font-size:26px;font-weight:800;color:var(--accent2);}
  .ds-lbl{font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);margin-top:2px;}
  .donut-wrap{display:flex;justify-content:center;margin:8px 0 16px;}
  svg.donut-chart{filter:drop-shadow(0 4px 20px rgba(232,96,60,0.15));}
  .bar-chart-wrap{margin-top:8px;}
  .bar-row{display:flex;align-items:center;gap:10px;margin-bottom:10px;}
  .bar-label{font-size:11px;color:var(--muted);width:30px;text-align:right;}
  .bar-track{flex:1;height:8px;background:rgba(255,255,255,0.05);border-radius:4px;overflow:hidden;}
  .bar-fill{height:100%;border-radius:4px;background:linear-gradient(90deg,var(--green),var(--accent2));transition:width .4s;}
  .bar-val{font-size:11px;color:var(--muted);width:28px;}
  .legend{display:flex;gap:20px;justify-content:center;}
  .leg-item{display:flex;align-items:center;gap:6px;font-size:12px;color:var(--muted);}
  .leg-dot{width:8px;height:8px;border-radius:50%;}
  .menu-section{padding:100px 40px;background:var(--surface);border-top:1px solid var(--border);}
  .menu-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:16px;max-width:1100px;margin:0 auto;}
  .menu-card{background:var(--card);border:1px solid var(--border);border-radius:16px;padding:28px 20px;text-align:center;transition:transform .2s,border-color .2s,box-shadow .2s;cursor:pointer;}
  .menu-card:hover{transform:translateY(-4px);border-color:rgba(232,96,60,0.4);box-shadow:0 12px 40px rgba(232,96,60,0.1);}
  .menu-card-icon{font-size:32px;margin-bottom:16px;display:block;}
  .menu-card h3{font-family:'Syne',sans-serif;font-size:14px;font-weight:700;margin-bottom:6px;}
  .menu-card p{font-size:12px;color:var(--muted);line-height:1.6;}
  .cta-section{padding:100px 40px;text-align:center;position:relative;overflow:hidden;}
  .cta-section::before{content:'';position:absolute;inset:0;background:radial-gradient(ellipse 700px 400px at 50% 50%,rgba(232,96,60,0.08),transparent);pointer-events:none;}
  .cta-section h2{font-family:'Syne',sans-serif;font-size:clamp(30px,5vw,56px);font-weight:800;margin-bottom:20px;line-height:1.1;}
  .cta-section h2 em{color:var(--accent);font-style:normal;}
  .cta-section p{color:var(--muted);max-width:480px;margin:0 auto 40px;font-size:16px;}
  footer{border-top:1px solid var(--border);padding:32px 40px;display:flex;justify-content:space-between;align-items:center;}
  footer .logo{font-size:16px;}
  footer p{color:var(--muted);font-size:13px;}
  @media(max-width:900px){.camera-grid,.dash-grid{grid-template-columns:1fr;}.features-grid{grid-template-columns:1fr 1fr;}.menu-grid{grid-template-columns:repeat(3,1fr);}nav{padding:0 20px;}.nav-links{display:none;}}
  @media(max-width:600px){.stats-bar{flex-direction:column;}.stat-item{border-right:none;border-bottom:1px solid var(--border);}.features-grid{grid-template-columns:1fr;}.menu-grid{grid-template-columns:1fr 1fr;}}
  .fade-up{opacity:0;transform:translateY(30px);animation:fadeUp .7s ease forwards;}
  .fade-up-2{animation-delay:.15s;}.fade-up-3{animation-delay:.3s;}.fade-up-4{animation-delay:.45s;}
  @keyframes fadeUp{to{opacity:1;transform:translateY(0)}}
</style>
</head>
<body>
<nav>
  <div class="logo"><div class="logo-icon">&#127917;</div>FaceTrack</div>
  <ul class="nav-links">
    <li><a href="#features">Features</a></li>
    <li><a href="#dashboard">Dashboard</a></li>
    <li><a href="#menu">Navigation</a></li>
  </ul>
  <button class="nav-cta">Get Started &#8594;</button>
</nav>
<section class="hero">
  <div class="hero-bg"></div>
  <div class="hero-badge fade-up"><span></span> Real-Time Face Recognition</div>
  <h1 class="fade-up fade-up-2">Attendance <em>reimagined</em> with face recognition</h1>
  <p class="hero-sub fade-up fade-up-3">A Streamlit-powered system that uses OpenCV and face_recognition to automate attendance tracking.</p>
  <div class="hero-btns fade-up fade-up-4">
    <button class="btn-primary">&#128065; View Live Dashboard</button>
    <button class="btn-outline">Explore Features</button>
  </div>
</section>
<div class="stats-bar">
  <div class="stat-item"><span class="stat-num">98.7%</span><div class="stat-label">Accuracy</div></div>
  <div class="stat-item"><span class="stat-num">150+</span><div class="stat-label">Users</div></div>
  <div class="stat-item"><span class="stat-num">2s</span><div class="stat-label">Detection</div></div>
  <div class="stat-item"><span class="stat-num">0ms</span><div class="stat-label">Manual Entry</div></div>
</div>
<section class="camera-section">
  <div class="camera-grid">
    <div class="camera-preview">
      <div class="camera-topbar">
        <div class="live-dot"><span></span> LIVE CAMERA</div>
        <div class="enc-badge">&#128274; ENCRYPTED</div>
      </div>
      <div class="face-box-wrap">
        <div class="scan-line"></div>
        <div class="face-box">
          <div class="face-corner fc-tl"></div><div class="face-corner fc-tr"></div>
          <div class="face-corner fc-bl"></div><div class="face-corner fc-br"></div>
          <div class="match-badge">MATCH FOUND</div>
          <div class="face-icon">
            <svg width="80" height="80" viewBox="0 0 80 80" fill="none">
              <circle cx="27" cy="32" r="5" fill="#3ecf6f" opacity=".7"/>
              <circle cx="53" cy="32" r="5" fill="#3ecf6f" opacity=".7"/>
              <path d="M22 54 Q40 66 58 54" stroke="#3ecf6f" stroke-width="3" fill="none" stroke-linecap="round" opacity=".7"/>
            </svg>
          </div>
        </div>
      </div>
      <div class="camera-bottombar">
        <div class="verified-tag">&#9989; Identity verified</div>
        <div style="font-size:11px;color:var(--muted)">09:14 AM</div>
      </div>
    </div>
    <div class="camera-text">
      <div class="section-label">LIVE RECOGNITION</div>
      <h2>Face-powered <em>instant</em> check-in</h2>
      <p>Using your device's browser camera via WebRTC, FaceTrack recognises registered users in real-time.</p>
      <div class="feature-chips">
        <div class="chip"><div class="chip-dot" style="background:#3ecf6f"></div>Browser camera</div>
        <div class="chip"><div class="chip-dot" style="background:#e8603c"></div>OpenCV processing</div>
        <div class="chip"><div class="chip-dot" style="background:#378add"></div>Supabase storage</div>
        <div class="chip"><div class="chip-dot" style="background:#d4f0c2"></div>Auto absent marking</div>
        <div class="chip"><div class="chip-dot" style="background:#ef9f27"></div>Excel export</div>
      </div>
    </div>
  </div>
</section>
<section class="features-section" id="features">
  <div class="section-header">
    <small>CONTROL ROOM</small>
    <h2>System <em>Features</em></h2>
    <p>Every component of the attendance system, built with Python and powered by real-time AI.</p>
  </div>
  <div class="features-grid">
    <div class="feat-card"><div class="feat-icon" style="background:rgba(62,207,111,0.12)">&#127917;</div><h3>Face Recognition</h3><p>OpenCV + face_recognition library for real-time biometric identification with 98.7% accuracy.</p></div>
    <div class="feat-card"><div class="feat-icon" style="background:rgba(232,96,60,0.12)">&#9729;&#65039;</div><h3>Supabase Backend</h3><p>PostgreSQL database with row-level security. Records and encodings stored and streamed live.</p></div>
    <div class="feat-card"><div class="feat-icon" style="background:rgba(55,138,221,0.12)">&#128202;</div><h3>Smart Dashboard</h3><p>Personal stats with SVG donut charts, monthly distribution, year-wise summaries and trends.</p></div>
    <div class="feat-card"><div class="feat-icon" style="background:rgba(212,240,194,0.12)">&#128197;</div><h3>Holiday Calendar</h3><p>India public holidays pre-configured. Working-day calculations exclude Sundays automatically.</p></div>
    <div class="feat-card"><div class="feat-icon" style="background:rgba(239,159,39,0.12)">&#129302;</div><h3>Auto-Absent System</h3><p>After 7 PM, the system marks all users without check-in as absent via bot — zero manual overhead.</p></div>
    <div class="feat-card"><div class="feat-icon" style="background:rgba(127,119,221,0.12)">&#128209;</div><h3>Excel Export</h3><p>Styled .xlsx reports with headers, freeze panes, total formulas — one-click export.</p></div>
  </div>
</section>
<section class="dashboard-section" id="dashboard">
  <div class="dash-grid">
    <div>
      <div class="section-label">ANALYTICS</div>
      <h2 style="font-family:'Syne',sans-serif;font-size:clamp(26px,4vw,40px);font-weight:800;margin-bottom:16px;line-height:1.15;">Live attendance <span style="color:var(--accent)">analytics</span> at a glance</h2>
      <p style="color:var(--muted);line-height:1.8;margin-bottom:28px;">Track present/absent ratios, monthly distributions, and year-on-year trends from your personal dashboard.</p>
      <div class="feature-chips">
        <div class="chip">&#128200; Donut charts</div>
        <div class="chip">&#128197; Monthly pie</div>
        <div class="chip">&#128202; Year-wise bar</div>
        <div class="chip">&#127884; Holiday-aware</div>
      </div>
    </div>
    <div class="dash-preview">
      <div class="dash-header"><div class="dash-title">Attendance Overview</div><div style="font-size:12px;color:var(--muted)">2026</div></div>
      <div class="donut-wrap">
        <svg class="donut-chart" width="200" height="200" viewBox="0 0 200 200">
          <path d="M 100 15 A 85 85 0 1 1 89.6 14.4" stroke="#3ecf6f" stroke-width="26" fill="none" stroke-linecap="round"/>
          <path d="M 89.6 14.4 A 85 85 0 0 1 100 15" stroke="#e8603c" stroke-width="26" fill="none" stroke-linecap="round"/>
          <text x="100" y="95" text-anchor="middle" font-family="Syne,sans-serif" font-size="28" font-weight="800" fill="#d4f0c2">88%</text>
          <text x="100" y="115" text-anchor="middle" font-family="DM Sans,sans-serif" font-size="11" fill="#7a9080">ATTENDANCE</text>
        </svg>
      </div>
      <div class="legend" style="margin-bottom:20px;">
        <div class="leg-item"><div class="leg-dot" style="background:#3ecf6f"></div>Present (142)</div>
        <div class="leg-item"><div class="leg-dot" style="background:#e8603c"></div>Absent (18)</div>
      </div>
      <div style="font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.08em;margin-bottom:12px;">Weekly Attendance</div>
      <div class="bar-chart-wrap">
        <div class="bar-row"><span class="bar-label">Mon</span><div class="bar-track"><div class="bar-fill" style="width:92%"></div></div><span class="bar-val">92%</span></div>
        <div class="bar-row"><span class="bar-label">Tue</span><div class="bar-track"><div class="bar-fill" style="width:88%"></div></div><span class="bar-val">88%</span></div>
        <div class="bar-row"><span class="bar-label">Wed</span><div class="bar-track"><div class="bar-fill" style="width:95%"></div></div><span class="bar-val">95%</span></div>
        <div class="bar-row"><span class="bar-label">Thu</span><div class="bar-track"><div class="bar-fill" style="width:78%"></div></div><span class="bar-val">78%</span></div>
        <div class="bar-row"><span class="bar-label">Fri</span><div class="bar-track"><div class="bar-fill" style="width:85%"></div></div><span class="bar-val">85%</span></div>
        <div class="bar-row"><span class="bar-label">Sat</span><div class="bar-track"><div class="bar-fill" style="width:60%"></div></div><span class="bar-val">60%</span></div>
      </div>
      <div class="dash-stats" style="margin-top:20px;">
        <div class="dash-stat"><div class="ds-val">142</div><div class="ds-lbl">Present Days</div></div>
        <div class="dash-stat"><div class="ds-val">18</div><div class="ds-lbl">Absent Days</div></div>
        <div class="dash-stat"><div class="ds-val" style="color:var(--accent)">88%</div><div class="ds-lbl">Attendance</div></div>
        <div class="dash-stat"><div class="ds-val">12</div><div class="ds-lbl">Holidays</div></div>
      </div>
    </div>
  </div>
</section>
<section class="menu-section" id="menu">
  <div class="section-header">
    <small>NAVIGATION</small>
    <h2>Everything you <em>need</em></h2>
    <p>Five focused screens, zero bloat — built for speed and clarity.</p>
  </div>
  <div class="menu-grid">
    <div class="menu-card"><span class="menu-card-icon">&#128100;</span><h3>Personal Dashboard</h3><p>Your stats, charts, profile editor, and year-wise breakdown</p></div>
    <div class="menu-card"><span class="menu-card-icon">&#128248;</span><h3>Take Photo</h3><p>Capture and upload face photos to Supabase storage bucket</p></div>
    <div class="menu-card"><span class="menu-card-icon">&#128204;</span><h3>Mark Attendance</h3><p>Face-recognition camera check-in within the 9 AM–7 PM window</p></div>
    <div class="menu-card"><span class="menu-card-icon">&#128202;</span><h3>Database</h3><p>Filter, search, and export attendance records to styled Excel</p></div>
    <div class="menu-card"><span class="menu-card-icon">&#128682;</span><h3>Logout</h3><p>Secure session logout with instant state clear and rerun</p></div>
  </div>
</section>
<section class="cta-section" id="cta">
  <h2>Ready to go <em>hands-free</em>?</h2>
  <p>Sign up, upload your face, and let FaceTrack handle the rest — automatically, every day.</p>
  <div class="hero-btns" style="justify-content:center">
    <button class="btn-primary">&#128640; Start Free</button>
    <button class="btn-outline">&#128196; View Docs</button>
  </div>
</section>
<footer>
  <div class="logo"><div class="logo-icon">&#127917;</div>FaceTrack</div>
  <p>Built with Streamlit · OpenCV · Supabase</p>
  <p style="color:var(--muted);font-size:13px;">© 2026 FaceTrack</p>
</footer>
</body>
</html>
'''

# ─────────────────────────────────────────
#  GLOBAL STYLES
# ─────────────────────────────────────────
st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;500;600;700;800;900&family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500;600&family=Space+Grotesk:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root {
  --void:#0a0a0c;--obsidian:#111114;--slate:#1a1a1f;--graphite:#242429;--steel:#2e2e35;
  --silver:#6e6e7a;--pewter:#8e8e9a;--bone:#c8c6c0;--ivory:#eae8e2;--cream:#f5f3ee;
  --snow:#faf9f6;--white:#ffffff;
  --emerald:#10b981;--emerald-dk:#059669;--emerald-glow:rgba(16,185,129,0.15);--emerald-deep:rgba(16,185,129,0.06);
  --amber:#f59e0b;--amber-lt:#fbbf24;--amber-glow:rgba(245,158,11,0.18);
  --rose:#f43f5e;--rose-glow:rgba(244,63,94,0.14);
  --sapphire:#3b82f6;--sapphire-glow:rgba(59,130,246,0.14);
  --r-xs:4px;--r-sm:8px;--r-md:12px;--r-lg:16px;--r-xl:24px;--r-2xl:32px;--r-full:9999px;
  --ease-out:cubic-bezier(0.16,1,0.3,1);--ease-spring:cubic-bezier(0.34,1.56,0.64,1);--ease-smooth:cubic-bezier(0.4,0,0.2,1);
  --font-display:'Playfair Display',Georgia,serif;--font-body:'Inter',-apple-system,sans-serif;
  --font-mono:'JetBrains Mono','Fira Code',monospace;--font-heading:'Space Grotesk','Inter',sans-serif;
}
*,*::before,*::after{box-sizing:border-box;}
.stApp{background:var(--void)!important;font-family:var(--font-body)!important;color:var(--ivory)!important;-webkit-font-smoothing:antialiased;}
.stApp::before{content:'';position:fixed;inset:0;background:radial-gradient(ellipse 900px 700px at 15% 5%,rgba(16,185,129,0.07) 0%,transparent 60%),radial-gradient(ellipse 600px 500px at 85% 90%,rgba(245,158,11,0.05) 0%,transparent 55%);pointer-events:none;z-index:0;}
footer,#MainMenu[data-testid="stHeader"]{visibility:hidden!important;height:0!important;overflow:hidden!important;}
section[data-testid="stSidebar"]{background:rgba(17,17,20,0.92)!important;backdrop-filter:blur(40px) saturate(150%)!important;border-right:1px solid rgba(255,255,255,0.06)!important;box-shadow:4px 0 40px rgba(0,0,0,0.5)!important;}
section[data-testid="stSidebar"] *{color:var(--bone)!important;}
section[data-testid="stSidebar"] .stRadio label,section[data-testid="stSidebar"] [data-testid="stWidgetLabel"]{font-size:0.72rem!important;font-weight:600!important;letter-spacing:0.14em!important;text-transform:uppercase!important;color:var(--silver)!important;}
#sidebar-brand{padding:2rem 1.4rem 1.6rem;border-bottom:1px solid rgba(255,255,255,0.06);margin-bottom:1rem;position:relative;}
#sidebar-brand .brand-icon{font-size:2rem;line-height:1;display:inline-block;animation:iconFloat 4s ease-in-out infinite;}
@keyframes iconFloat{0%,100%{transform:translateY(0) rotate(0deg);}25%{transform:translateY(-3px) rotate(-2deg);}75%{transform:translateY(2px) rotate(1deg);}}
#sidebar-brand .brand-title{font-family:var(--font-display)!important;font-size:1.4rem!important;font-weight:700!important;color:var(--white)!important;letter-spacing:-0.02em!important;margin:0.5rem 0 0.15rem!important;}
#sidebar-brand .brand-sub{font-family:var(--font-heading)!important;font-size:0.65rem!important;font-weight:500!important;color:var(--emerald)!important;letter-spacing:0.18em!important;text-transform:uppercase!important;}
h1{font-family:var(--font-display)!important;font-weight:800!important;font-size:clamp(1.8rem,3.5vw,2.8rem)!important;letter-spacing:-0.03em!important;color:var(--white)!important;line-height:1.1!important;background:linear-gradient(135deg,var(--white) 0%,var(--emerald) 50%,var(--amber) 100%)!important;-webkit-background-clip:text!important;-webkit-text-fill-color:transparent!important;background-clip:text!important;margin-bottom:0.15em!important;}
h2{font-family:var(--font-heading)!important;font-weight:600!important;font-size:1.35rem!important;color:var(--white)!important;letter-spacing:-0.01em!important;}
h3{font-family:var(--font-heading)!important;font-weight:600!important;font-size:0.82rem!important;color:var(--emerald)!important;letter-spacing:0.12em!important;text-transform:uppercase!important;}
p,li{font-family:var(--font-body)!important;font-size:0.9rem!important;line-height:1.75!important;color:#ffffff !important;}
strong{color:var(--emerald)!important;font-weight:600!important;}
code{font-family:var(--font-mono)!important;font-size:0.8rem!important;background:rgba(16,185,129,0.1)!important;color:var(--emerald)!important;padding:2px 8px!important;border-radius:var(--r-xs)!important;border:1px solid rgba(16,185,129,0.15)!important;}
#page-header{display:flex;align-items:center;gap:1rem;padding:0.6rem 0 1.8rem;border-bottom:1px solid rgba(255,255,255,0.05);margin-bottom:2rem;position:relative;}
#page-header::after{content:'';position:absolute;bottom:-1px;left:0;width:80px;height:2px;background:linear-gradient(90deg,var(--emerald),transparent);border-radius:2px;}
#page-header .header-badge{width:50px;height:50px;background:linear-gradient(145deg,var(--graphite),var(--slate));border:1px solid rgba(255,255,255,0.08);border-radius:var(--r-lg);display:flex;align-items:center;justify-content:center;font-size:1.5rem;box-shadow:0 8px 24px rgba(0,0,0,0.4),inset 0 1px 0 rgba(255,255,255,0.06);flex-shrink:0;}
div[data-testid="stForm"]{background:rgba(26,26,31,0.6)!important;backdrop-filter:blur(24px) saturate(120%)!important;border:1px solid rgba(255,255,255,0.07)!important;border-radius:var(--r-xl)!important;padding:2.2rem 2.5rem!important;box-shadow:0 24px 80px rgba(0,0,0,0.4),inset 0 1px 0 rgba(255,255,255,0.04)!important;position:relative;overflow:hidden;}
input[type="text"],input[type="password"]{font-family:var(--font-body)!important;font-size:0.92rem!important;color:var(--white)!important;background:rgba(17,17,20,0.8)!important;border:1.5px solid rgba(255,255,255,0.08)!important;border-radius:var(--r-md)!important;padding:0.75rem 1.1rem!important;transition:all 0.3s var(--ease-out)!important;box-shadow:inset 0 2px 6px rgba(0,0,0,0.3)!important;}
input[type="text"]:focus,input[type="password"]:focus{border-color:var(--emerald)!important;box-shadow:0 0 0 4px var(--emerald-glow),inset 0 2px 6px rgba(0,0,0,0.3)!important;background:rgba(17,17,20,1)!important;outline:none!important;}
.stTextInput label,.stPasswordInput label{font-family:var(--font-heading)!important;font-size:0.72rem!important;font-weight:600!important;letter-spacing:0.12em!important;text-transform:uppercase!important;color:var(--silver)!important;margin-bottom:6px!important;}
.stButton>button{font-family:var(--font-heading)!important;font-size:0.82rem!important;font-weight:600!important;letter-spacing:0.14em!important;text-transform:uppercase!important;color:var(--white)!important;background:linear-gradient(135deg,#051396  0%,#10b92d 100%)!important;border:1px solid rgba(16,185,129,0.3)!important;border-radius:var(--r-md)!important;padding:0.72rem 2rem!important;cursor:pointer!important;overflow:hidden!important;transition:all 0.3s var(--ease-out)!important;box-shadow:0 6px 24px rgba(16,185,129,0.25),inset 0 1px 0 rgba(255,255,255,0.15)!important;}
.stButton>button:hover{transform:translateY(-2px)!important;box-shadow:0 12px 40px rgba(16,185,129,0.35),inset 0 1px 0 rgba(255,255,255,0.2)!important;}
.stButton>button:active{transform:translateY(0) scale(0.98)!important;}
.stDownloadButton>button{background:linear-gradient(135deg,#b45309,var(--amber))!important;border-color:rgba(245,158,11,0.3)!important;box-shadow:0 6px 24px rgba(245,158,11,0.2),inset 0 1px 0 rgba(255,255,255,0.15)!important;color:var(--void)!important;}
[data-testid="metric-container"]{background:rgba(26,26,31,0.5)!important;backdrop-filter:blur(16px)!important;border:1px solid rgba(255,255,255,0.06)!important;border-radius:var(--r-lg)!important;padding:1.4rem 1.6rem!important;box-shadow:0 12px 40px rgba(0,0,0,0.3),inset 0 1px 0 rgba(255,255,255,0.04)!important;transition:all 0.3s var(--ease-out)!important;position:relative;overflow:hidden;}
[data-testid="metric-container"]::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,var(--emerald),var(--amber));opacity:0.7;}
[data-testid="metric-container"]:hover{transform:translateY(-4px)!important;border-color:rgba(255,255,255,0.1)!important;}
[data-testid="metric-container"] [data-testid="stMetricLabel"]{font-family:var(--font-heading)!important;font-size:0.68rem!important;font-weight:600!important;letter-spacing:0.14em!important;text-transform:uppercase!important;color:var(--silver)!important;}
[data-testid="metric-container"] [data-testid="stMetricValue"]{font-family:var(--font-display)!important;font-size:2.4rem!important;font-weight:700!important;color:var(--white)!important;line-height:1.1!important;}
[data-testid="stDataFrame"]{border-radius:var(--r-lg)!important;overflow:hidden!important;box-shadow:0 12px 40px rgba(0,0,0,0.3)!important;border:1px solid rgba(255,255,255,0.06)!important;background:rgba(26,26,31,0.4)!important;}
.stAlert,[data-testid="stAlert"]{border-radius:var(--r-lg)!important;font-family:var(--font-body)!important;font-size:0.88rem!important;border:1px solid rgba(255,255,255,0.06)!important;border-left-width:3px!important;backdrop-filter:blur(16px)!important;box-shadow:0 8px 32px rgba(0,0,0,0.25)!important;}
[data-baseweb="select"]>div{background:rgba(17,17,20,0.8)!important;border:1.5px solid rgba(255,255,255,0.08)!important;border-radius:var(--r-md)!important;font-family:var(--font-body)!important;color:var(--ivory)!important;transition:all 0.25s var(--ease-out)!important;}
[data-baseweb="select"]>div:hover,[data-baseweb="select"]>div:focus-within{border-color:var(--emerald)!important;box-shadow:0 0 0 4px var(--emerald-glow)!important;}
hr,.stDivider{border:none!important;height:1px!important;background:linear-gradient(90deg,transparent 0%,rgba(255,255,255,0.08) 20%,rgba(16,185,129,0.2) 50%,rgba(255,255,255,0.08) 80%,transparent 100%)!important;margin:2.5rem 0!important;}
[data-testid="stCameraInput"]{border-radius:var(--r-xl)!important;overflow:hidden!important;border:2px solid rgba(255,255,255,0.08)!important;transition:all 0.3s var(--ease-out)!important;box-shadow:0 12px 40px rgba(0,0,0,0.3)!important;}
[data-testid="stCameraInput"]:hover{border-color:var(--emerald)!important;}
::-webkit-scrollbar{width:6px;height:6px;}
::-webkit-scrollbar-track{background:transparent;}
::-webkit-scrollbar-thumb{background:var(--steel);border-radius:10px;}
::-webkit-scrollbar-thumb:hover{background:var(--emerald);}
.info-row{display:flex;flex-wrap:wrap;gap:2rem;padding:1.2rem 1.6rem;background:rgba(16,185,129,0.04);border-radius:var(--r-lg);border:1px solid rgba(16,185,129,0.1);margin:1rem 0;backdrop-filter:blur(8px);}
.info-item{display:flex;flex-direction:column;}
.info-label{font-family:var(--font-heading)!important;font-size:0.65rem!important;font-weight:600!important;letter-spacing:0.14em!important;text-transform:uppercase!important;color:var(--silver)!important;margin-bottom:2px!important;}
.info-value{font-family:var(--font-display)!important;font-size:1.15rem!important;font-weight:600!important;color:var(--white)!important;}
.datetime-strip{display:inline-flex;align-items:center;gap:8px;padding:6px 16px;background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.06);border-radius:var(--r-full);font-family:var(--font-mono)!important;font-size:0.75rem!important;color:var(--silver)!important;letter-spacing:0.04em!important;}
.datetime-strip::before{content:'';width:6px;height:6px;border-radius:50%;background:var(--emerald);animation:dotPulse 2s ease-in-out infinite;}
@keyframes dotPulse{0%,100%{opacity:1;transform:scale(1);}50%{opacity:0.4;transform:scale(0.7);}}
::selection{background:rgba(16,185,129,0.3);color:var(--white);}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
#  SIDEBAR BRAND BLOCK
# ─────────────────────────────────────────
st.sidebar.markdown("""
<div id="sidebar-brand">
  <div class="brand-icon">🎭</div>
  <div class="brand-title">FaceTrack</div>
  <div class="brand-sub">Real-Time Recognition</div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
#  SUPABASE — only init if available
# ─────────────────────────────────────────
def init_supabase():
    if not SUPABASE_AVAILABLE:
        err_msg = f"Supabase package is not available. Import error: {SUPABASE_IMPORT_ERROR}"
        st.error(err_msg)
        print(err_msg)
        return None
    if not SUPABASE_URL or not SUPABASE_KEY:
        st.error(
            "⚠️ Supabase credentials are missing.\n\n"
            "**For Streamlit Cloud deployment:** Go to your app settings → Secrets and add:\n"
            "```\nSUPABASE_URL = \"https://your-project.supabase.co\"\n"
            "SUPABASE_KEY = \"your-anon-key\"\n```\n\n"
            "**For local development:** Create `.streamlit/secrets.toml` with the same keys."
        )
        return None
    try:
        c = create_client(SUPABASE_URL, SUPABASE_KEY)
        print("Supabase client initialized successfully")
        return c
    except Exception as e:
        err_msg = f"Supabase connection error: {e}"
        st.error(err_msg)
        print(err_msg)
        return None
# ROLE CHECK
def ensure_admin_user(client):
    if client is None:
        return
    try:
        res = client.table("users").select("*").eq("username", "admin").execute()
        if not res.data:
            client.table("users").insert({
                "username": "admin",
                "password": "admin",  # Default admin password
                "role": "admin",
                "department": "Administration",
                "emp_id": "ADM-001"
            }).execute()
            print("Special admin account created.")
        else:
            # If the user exists but role column is not set or not 'admin'
            if res.data[0].get("role") != "admin":
                client.table("users").update({"role": "admin"}).eq("username", "admin").execute()
                print("Admin account role updated to admin.")
    except Exception as e:
        print("Note: Could not ensure admin user (make sure you ran the SQL script to add the 'role' column):", e)

supabase = init_supabase()
ensure_admin_user(supabase)

# ─────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────
def working_days_in_range(start: datetime.date, end: datetime.date) -> int:
    count, d = 0, start
    while d <= end:
        if d.weekday() != 6 and d not in INDIA_HOLIDAYS:
            count += 1
        d += datetime.timedelta(days=1)
    return count

def attendance_pct(present: int, working: int) -> float:
    return round((present / working) * 100, 1) if working else 0.0

def arc_path(cx, cy, r, start_deg, end_deg):
    s  = math.radians(start_deg - 90)
    e  = math.radians(end_deg   - 90)
    x1, y1 = cx + r * math.cos(s), cy + r * math.sin(s)
    x2, y2 = cx + r * math.cos(e), cy + r * math.sin(e)
    large  = 1 if (end_deg - start_deg) > 180 else 0
    return f"M {x1:.2f} {y1:.2f} A {r} {r} 0 {large} 1 {x2:.2f} {y2:.2f}"

def mark_auto_absent():
    """Mark absent — silently skips if Supabase unavailable."""
    if supabase is None:
        return
    today     = datetime.date.today()
    today_str = str(today)
    now_time  = datetime.datetime.now().time()
    if today.weekday() == 6:
        return
    if today in INDIA_HOLIDAYS:
        return
    if now_time < ABSENT_TIME:
        return
    try:
        users = supabase.table("users").select("username", "department").execute()
        if not users.data:
            return
        for user in users.data:
            name = user.get("username")
            dept = user.get("department")
            existing = supabase.table("attendance")\
                .select("*").eq("name", name).eq("date", today_str).execute()
            if not existing.data:
                supabase.table("attendance").insert({
                    "name": name, "date": today_str, "time": "00:00:00",
                    "marked_by": "system", "department": dept, "status": "absent"
                }).execute()
    except Exception as e:
        print("Auto absent check error:", e)

# SECURITY LOGS
def log_security_event_async(username, event_type, details):
    """Log a security event to Supabase in a background thread to avoid blocking."""
    if supabase is None:
        return
    def run():
        try:
            supabase.table("security_logs").insert({
                "username": username,
                "event_type": event_type,
                "details": details
            }).execute()
        except Exception as e:
            print(f"Failed to insert security log: {e}")
    threading.Thread(target=run, daemon=True).start()

# ─────────────────────────────────────────
#  LIVENESS PROCESSOR (WebRTC)
# ─────────────────────────────────────────
if WEBRTC_AVAILABLE:
    class FaceLivenessProcessor(VideoTransformerBase):
        def __init__(self):
            self.liveness_step = 0
            self.action_challenge = None
            self.initial_encoding = None
            self.initial_landmarks = None
            self.eye_dist_init = 1.0
            self.blink_detected_closed = False
            self.verified_name = None
            self.verification_success = False
            self.frames_passed = 0
            self.known_encodings = []
            self.known_names = []
            self.supabase = None
            self.user_dept = ""
            self.username = ""
            self.today_date = ""
            self.lock = threading.Lock()
            self.frame_count = 0
            self.last_img_annotated = None
            self.challenge_eval_count = 0  # To track timeouts
            
        def recv(self, frame: av.VideoFrame) -> av.VideoFrame:
            img = frame.to_ndarray(format="bgr24")

            with self.lock:
                # ── Already verified: just annotate and return ──────────────
                if self.verification_success:
                    cv2.putText(img, f"Verified: {self.verified_name}! Attendance marked.",
                                (20, 40), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2)
                    return av.VideoFrame.from_ndarray(img, format="bgr24")

                self.frame_count += 1

                # ── Always overlay the current challenge hint ────────────────
                # (shown on every frame so user sees it even on skipped frames)
                if self.liveness_step == 1 and self.action_challenge:
                    challenge_text = f"Challenge: {self.action_challenge}  ({self.frames_passed}/2)"
                    cv2.putText(img, challenge_text, (20, 40),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 255, 255), 2)
                elif self.liveness_step == 0:
                    cv2.putText(img, "Scanning for face — please look at the camera",
                                (20, 40), cv2.FONT_HERSHEY_SIMPLEX, 0.65, (200, 200, 200), 2)

                # ── Throttle heavy processing to every 10th frame ───────────
                if self.frame_count % 10 != 0:
                    return av.VideoFrame.from_ndarray(img, format="bgr24")

                # ── Downscale for speed ─────────────────────────────────────
                small = cv2.resize(img, (0, 0), fx=0.25, fy=0.25)
                rgb_s = cv2.cvtColor(small, cv2.COLOR_BGR2RGB)
                locs  = face_recognition.face_locations(rgb_s, model="hog")

                # SECURITY LOGS: Multiple faces check
                if not locs or len(locs) > 1:
                    if locs and len(locs) > 1:
                        log_security_event_async(
                            self.username or "unknown",
                            "multiple_faces",
                            f"Multiple faces ({len(locs)}) detected in camera frame during check-in."
                        )
                    cv2.putText(img, "Ensure exactly ONE face is visible.",
                                (20, 40), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)
                    return av.VideoFrame.from_ndarray(img, format="bgr24")

                landmarks = face_recognition.face_landmarks(rgb_s, locs)
                if not landmarks:
                    return av.VideoFrame.from_ndarray(img, format="bgr24")

                lm = landmarks[0]

                # ── Step 0: Baseline Capture ────────────────────────────────
                if self.liveness_step == 0:
                    encs = face_recognition.face_encodings(rgb_s, locs)
                    if not encs:
                        return av.VideoFrame.from_ndarray(img, format="bgr24")
                    self.initial_encoding  = encs[0]
                    self.initial_landmarks = lm
                    self.eye_dist_init     = math.hypot(
                        lm["left_eye"][0][0] - lm["right_eye"][3][0],
                        lm["left_eye"][0][1] - lm["right_eye"][3][1])
                    if self.eye_dist_init < 1e-6:
                        self.eye_dist_init = 1.0
                    # Only Smile, Raise Eyebrows, Blink (no turn challenges)
                    self.action_challenge      = random.choice(
                        ["Smile", "Raise Eyebrows", "Blink"])
                    self.liveness_step         = 1
                    self.frames_passed         = 0
                    self.blink_detected_closed = False
                    self.challenge_eval_count  = 0
                    cv2.putText(img, f"Baseline OK! Now: {self.action_challenge}",
                                (20, 40), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 255, 0), 2)
                    return av.VideoFrame.from_ndarray(img, format="bgr24")

                # ── Step 1: Challenge check ────────────────────────────────
                if self.liveness_step == 1:
                    self.challenge_eval_count += 1
                    if self.challenge_eval_count > 20:
                        log_security_event_async(
                            self.username or "unknown",
                            "failed_liveness",
                            f"Liveness challenge '{self.action_challenge}' failed/timed out after 20 evaluations."
                        )
                        self.liveness_step = 0
                        self.frames_passed = 0
                        self.challenge_eval_count = 0
                        cv2.putText(img, "Liveness challenge timed out. Restarting.",
                                    (20, 40), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)
                        return av.VideoFrame.from_ndarray(img, format="bgr24")

                    eye_dist = math.hypot(
                        lm["left_eye"][0][0] - lm["right_eye"][3][0],
                        lm["left_eye"][0][1] - lm["right_eye"][3][1])
                    if eye_dist < 1e-6:
                        eye_dist = 1.0

                    action_passed = False

                    if self.action_challenge == "Smile":
                        w_init = math.hypot(
                            self.initial_landmarks["bottom_lip"][6][0] - self.initial_landmarks["bottom_lip"][0][0],
                            self.initial_landmarks["bottom_lip"][6][1] - self.initial_landmarks["bottom_lip"][0][1]
                        ) / self.eye_dist_init
                        w_curr = math.hypot(
                            lm["bottom_lip"][6][0] - lm["bottom_lip"][0][0],
                            lm["bottom_lip"][6][1] - lm["bottom_lip"][0][1]
                        ) / eye_dist
                        if w_curr > w_init * 1.08:
                            action_passed = True

                    

                    elif self.action_challenge == "Raise Eyebrows":
                        init_dist = math.hypot(
                            self.initial_landmarks["left_eye"][1][0] - self.initial_landmarks["left_eyebrow"][2][0],
                            self.initial_landmarks["left_eye"][1][1] - self.initial_landmarks["left_eyebrow"][2][1]
                        ) / self.eye_dist_init
                        curr_dist = math.hypot(
                            lm["left_eye"][1][0] - lm["left_eyebrow"][2][0],
                            lm["left_eye"][1][1] - lm["left_eyebrow"][2][1]
                        ) / eye_dist
                        if curr_dist > init_dist * 1.12:
                            action_passed = True

                    elif self.action_challenge == "Blink":
                        def ear(eye_pts):
                            A = math.hypot(eye_pts[1][0]-eye_pts[5][0], eye_pts[1][1]-eye_pts[5][1])
                            B = math.hypot(eye_pts[2][0]-eye_pts[4][0], eye_pts[2][1]-eye_pts[4][1])
                            C = math.hypot(eye_pts[0][0]-eye_pts[3][0], eye_pts[0][1]-eye_pts[3][1]) + 1e-6
                            return (A + B) / (2.0 * C)
                        EAR = (ear(lm["left_eye"]) + ear(lm["right_eye"])) / 2.0
                        if EAR < 0.20:
                            self.blink_detected_closed = True
                        elif EAR > 0.25 and self.blink_detected_closed:
                            action_passed = True

                    if action_passed:
                        self.frames_passed += 1

                    # ── Enough challenge frames: run face match + DB insert ──
                    if self.frames_passed >= 2:
                        encs = face_recognition.face_encodings(rgb_s, locs)
                        if encs:
                            enc     = encs[0]
                            matches = face_recognition.compare_faces(
                                [self.initial_encoding], enc, tolerance=0.5)
                            if not matches[0]:
                                # Face changed — reset and restart
                                log_security_event_async(
                                    self.username or "unknown",
                                    "photo_spoofing",
                                    "Face mismatch between baseline and verification check. Possible photo spoofing or mid-session swap."
                                )
                                self.liveness_step = 0
                                self.frames_passed = 0
                                cv2.putText(img, "Face mismatch — restarting.",
                                            (20, 40), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)
                                return av.VideoFrame.from_ndarray(img, format="bgr24")

                            if self.known_encodings:
                                db_matches = face_recognition.compare_faces(
                                    self.known_encodings, enc, tolerance=0.5)
                                if True in db_matches:
                                    idx = int(np.argmin(
                                        face_recognition.face_distance(self.known_encodings, enc)))
                                    self.verified_name = self.known_names[idx]
                                    # ── Mark attendance in Supabase ───────────
                                    if self.supabase:
                                        try:
                                            existing = self.supabase.table("attendance") \
                                                .select("*") \
                                                .eq("name", self.verified_name) \
                                                .eq("date", self.today_date) \
                                                .execute()
                                            already_present = any(
                                                r.get("status") == "present" and
                                                r.get("marked_by") != "system"
                                                for r in (existing.data or []))
                                            if not already_present:
                                                now_dt = datetime.datetime.now()
                                                self.supabase.table("attendance").insert({
                                                    "name":       self.verified_name,
                                                    "date":       str(now_dt.date()),
                                                    "time":       str(now_dt.time()),
                                                    "marked_by":  self.username,
                                                    "department": self.user_dept,
                                                    "status":     "present"
                                                }).execute()
                                            # Mark success regardless (already marked = fine)
                                            self.verification_success = True
                                        except Exception as db_err:
                                            print("DB insert error:", db_err)
                                            # Still mark verified so UI shows success
                                            self.verification_success = True
                                    else:
                                        # No supabase reference — still verify locally
                                        self.verification_success = True
                                else:
                                    # Not in DB — reset
                                    log_security_event_async(
                                        self.username or "unknown",
                                        "unknown_face",
                                        f"Face verification failed: Captured face does not match registered biometric data for @{self.username}."
                                    )
                                    self.liveness_step = 0
                                    self.frames_passed = 0
                            else:
                                # No known encodings loaded — reset
                                log_security_event_async(
                                    self.username or "unknown",
                                    "unknown_face",
                                    f"No registered face datasets loaded for @{self.username} check-in attempt."
                                )
                                self.liveness_step = 0
                                self.frames_passed = 0

                    # Annotate progress on the current frame
                    if self.verification_success:
                        cv2.putText(img, f"Verified: {self.verified_name}! Attendance marked.",
                                    (20, 40), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2)
                    elif self.liveness_step == 1:
                        cv2.putText(
                            img,
                            f"Challenge: {self.action_challenge}  ({self.frames_passed}/2)",
                            (20, 40), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 255, 255), 2)

            return av.VideoFrame.from_ndarray(img, format="bgr24")

# ─────────────────────────────────────────
#  EXCEL AUTO-SAVE
# ─────────────────────────────────────────
def save_attendance_excel():
    if not OPENPYXL_AVAILABLE:
        return False, "openpyxl not installed"
    if supabase is None:
        return False, "Supabase not available"
    try:
        os.makedirs(EXCEL_FOLDER, exist_ok=True)
        data    = supabase.table("attendance").select("*").order("date", desc=True).order("time", desc=True).execute()
        records = data.data or []
        wb = Workbook(); ws = wb.active; ws.title = "Attendance"
        hdr_font  = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        hdr_fill  = PatternFill("solid", start_color="111114")
        ca        = Alignment(horizontal="center", vertical="center")
        la        = Alignment(horizontal="left",   vertical="center")
        bdr       = Border(
            left=Side(style="thin", color="2E2E35"), right=Side(style="thin", color="2E2E35"),
            top=Side(style="thin",  color="2E2E35"), bottom=Side(style="thin", color="2E2E35")
        )
        alt_fill     = PatternFill("solid", start_color="1A1A1F")
        present_fill = PatternFill("solid", start_color="0D4F2E")
        absent_fill  = PatternFill("solid", start_color="4F1010")
        body_font = Font(name="Arial", size=10, color="C8C6C0")
        ws.merge_cells("A1:G1"); ws["A1"] = "Attendance Record"
        ws["A1"].font = Font(name="Arial", bold=True, size=14, color="10B981"); ws["A1"].alignment = ca
        ws.row_dimensions[1].height = 30
        ws.merge_cells("A2:G2")
        ws["A2"] = f"Generated: {datetime.datetime.now().strftime('%d %b %Y, %I:%M %p')}"
        ws["A2"].font = Font(name="Arial", size=9, color="6E6E7A"); ws["A2"].alignment = ca
        ws.row_dimensions[2].height = 16; ws.append([])
        headers = ["#", "Name", "Date", "Time", "Marked By", "Department", "Status"]
        ws.append(headers)
        for ci in range(1, 8):
            c = ws.cell(row=4, column=ci)
            c.font = hdr_font; c.fill = hdr_fill; c.alignment = ca; c.border = bdr
        ws.row_dimensions[4].height = 22
        for i, r in enumerate(records, 1):
            st_val = str(r.get("status", "present")).title()
            if r.get("marked_by") == "system":
                st_val = "Absent"
            ws.append([i, r.get("name", ""), r.get("date", ""),
                       str(r.get("time", ""))[:8], r.get("marked_by", ""),
                       r.get("department", ""), st_val])
            rn = 4 + i
            for ci in range(1, 8):
                c = ws.cell(row=rn, column=ci)
                c.font = body_font; c.border = bdr
                c.alignment = la if ci == 2 else ca
                if st_val == "Absent":
                    c.fill = absent_fill
                elif st_val == "Present":
                    c.fill = present_fill
                elif i % 2 == 0:
                    c.fill = alt_fill
            ws.row_dimensions[rn].height = 18
        sr = 4 + len(records) + 1
        ws.cell(row=sr, column=1).value = "Total"
        ws.cell(row=sr, column=1).font  = Font(name="Arial", bold=True, size=10, color="10B981")
        ws.cell(row=sr, column=2).value = f"=COUNTA(B5:B{4+len(records)})"
        ws.cell(row=sr, column=2).font  = Font(name="Arial", bold=True, size=10, color="10B981")
        ws.cell(row=sr, column=2).alignment = ca
        for ci, w in enumerate([5, 22, 14, 12, 16, 18, 12], 1):
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.freeze_panes = "A5"
        wb.save(EXCEL_FILE)
        return True, EXCEL_FILE
    except Exception as e:
        return False, str(e)
# ─────────────────────────────────────────
#  ADMIN PANEL RENDER FUNCTIONS
# ─────────────────────────────────────────

# ADMIN PANEL
def render_admin_dashboard():
    # ROLE CHECK
    if st.session_state.role != "admin":
        st.error("Access Denied: You do not have permission to access this page.")
        st.stop()

    st.markdown("""
    <div id="page-header">
      <div class="header-badge">👑</div>
      <div>
        <h2 style="margin:0;font-size:1.5rem;">Admin Dashboard</h2>
        <p style="margin:0;color:var(--silver);font-size:0.82rem;">FaceTrack system analytics & real-time summary</p>
      </div>
    </div>
    """, unsafe_allow_html=True)

    today_str = str(datetime.date.today())

    # Get data
    try:
        # Students count
        users_resp = supabase.table("users").select("*").eq("role", "student").execute()
        students = users_resp.data or []
        total_students = len(students)

        # Attendance today
        att_resp = supabase.table("attendance").select("*").eq("date", today_str).execute()
        att_records = att_resp.data or []
        present_names = set(r.get("name") for r in att_records if r.get("status") == "present" and r.get("marked_by") != "system")
        present_today = len(present_names)
        absent_today = max(0, total_students - present_today)

        # Security logs count
        logs_resp = supabase.table("security_logs").select("*", count="exact").execute()
        total_alerts = logs_resp.count if logs_resp.count is not None else len(logs_resp.data or [])
    except Exception as e:
        st.error(f"Failed to query database statistics: {e}")
        total_students, present_today, absent_today, total_alerts = 0, 0, 0, 0
        students, att_records = [], []

    # KPI Metrics
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("👤 Total Students", total_students)
    m2.metric("✅ Present Today", present_today)
    m3.metric("❌ Absent Today", absent_today)
    m4.metric("🚨 Security Alerts", total_alerts)

    st.divider()

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### 📊 Today's Attendance Overview")
        if total_students > 0:
            present_pct = round((present_today / total_students) * 100, 1)
            p_end = (present_today / total_students) * 360
            cx, cy, ro, ri = 110, 110, 85, 54
            
            if present_today > 0 and absent_today > 0:
                slices_svg = (
                    f"<path d='{arc_path(cx,cy,ro,0,p_end)} "
                    f"L {cx+ri*math.cos(math.radians(p_end-90)):.2f} {cy+ri*math.sin(math.radians(p_end-90)):.2f} "
                    f"{arc_path(cx,cy,ri,p_end,0)[2:]} Z' fill='#10b981' opacity='0.92'/>"
                    f"<path d='{arc_path(cx,cy,ro,p_end,360)} "
                    f"L {cx+ri*math.cos(math.radians(360-90)):.2f} {cy+ri*math.sin(math.radians(360-90)):.2f} "
                    f"{arc_path(cx,cy,ri,360,p_end)[2:]} Z' fill='#f43f5e' opacity='0.85'/>"
                )
            elif present_today > 0:
                slices_svg = f"<circle cx='{cx}' cy='{cy}' r='{ro}' fill='#10b981' opacity='0.92'/>"
            else:
                slices_svg = f"<circle cx='{cx}' cy='{cy}' r='{ro}' fill='#f43f5e' opacity='0.85'/>"

            donut_svg = f"""
            <div style="display:flex; justify-content:center; margin:1rem 0;">
            <svg viewBox="0 0 220 220" xmlns="http://www.w3.org/2000/svg" width="220" height="220">
              <defs><filter id="shadow" x="-20%" y="-20%" width="140%" height="140%">
                <feDropShadow dx="0" dy="4" stdDeviation="8" flood-color="rgba(0,0,0,0.4)"/>
              </filter></defs>
              <g filter="url(#shadow)">{slices_svg}<circle cx="{cx}" cy="{cy}" r="{ri}" fill="#0a0a0c"/></g>
              <text x="{cx}" y="{cy-8}" text-anchor="middle" font-size="22" font-weight="700"
                    font-family="Playfair Display,serif" fill="#ffffff">{present_pct}%</text>
              <text x="{cx}" y="{cy+13}" text-anchor="middle" font-size="10"
                    font-family="Space Grotesk,sans-serif" fill="#6e6e7a" letter-spacing="2">ATTENDANCE</text>
            </svg>
            </div>
            """
            st.markdown(donut_svg, unsafe_allow_html=True)
        else:
            st.info("No students registered yet.")

    with c2:
        st.markdown("### 🏢 Department-wise Attendance Status")
        dept_data = []
        for s in students:
            username = s.get("username")
            dept = s.get("department") or "General"
            is_present = username in present_names
            dept_data.append({"username": username, "department": dept, "status": "Present" if is_present else "Absent"})
            
        if dept_data:
            df_dept = pd.DataFrame(dept_data)
            df_grouped = df_dept.groupby(["department", "status"]).size().unstack(fill_value=0)
            if "Present" not in df_grouped.columns:
                df_grouped["Present"] = 0
            if "Absent" not in df_grouped.columns:
                df_grouped["Absent"] = 0
            st.bar_chart(df_grouped[["Present", "Absent"]])
        else:
            st.info("No student data available to display chart.")


# USER MANAGEMENT
def render_user_management():
    # ROLE CHECK
    if st.session_state.role != "admin":
        st.error("Access Denied: You do not have permission to access this page.")
        st.stop()

    st.markdown("""
    <div id="page-header">
      <div class="header-badge">👤</div>
      <div>
        <h2 style="margin:0;font-size:1.5rem;">User Management</h2>
        <p style="margin:0;color:var(--silver);font-size:0.82rem;">Create, edit, reset biometric records, or delete users</p>
      </div>
    </div>
    """, unsafe_allow_html=True)

    try:
        users_resp = supabase.table("users").select("*").order("username").execute()
        users_list = users_resp.data or []
    except Exception as e:
        st.error(f"Failed to fetch users: {e}")
        return

    # User search
    search_q = st.text_input("🔍 Search Users by Username", placeholder="Enter query...").strip().lower()
    filtered_users = users_list
    if search_q:
        filtered_users = [u for u in users_list if search_q in u.get("username", "").lower()]

    # Render User List
    df_users = pd.DataFrame(filtered_users)
    if not df_users.empty:
        # Hide password column for display
        display_cols = ["id", "username", "role", "emp_id", "department", "created_at"]
        st.dataframe(df_users[[c for c in display_cols if c in df_users.columns]], use_container_width=True, hide_index=True)
    else:
        st.info("No users match your search criteria.")

    st.divider()

    st.markdown("### 🛠️ User Actions")
    usernames = [u.get("username") for u in users_list if u.get("username")]
    selected_username = st.selectbox("Select User to Manage", ["-- Select User --"] + usernames)

    if selected_username != "-- Select User --":
        u_info = next((u for u in users_list if u.get("username") == selected_username), {})
        
        tab_edit, tab_reset, tab_delete = st.tabs(["📝 Edit Info", "🔄 Reset Face Data", "🗑️ Delete User"])
        
        with tab_edit:
            st.markdown(f"#### 📝 Edit User Details: @{selected_username}")
            with st.form("admin_edit_user"):
                new_emp_id = st.text_input("Employee ID", value=u_info.get("emp_id") or "")
                new_dept = st.text_input("Department", value=u_info.get("department") or "")
                new_role = st.selectbox("Role", ["student", "admin"], index=0 if u_info.get("role") == "student" else 1)
                new_pwd = st.text_input("New Password (leave blank to keep current)", type="password")
                
                submitted = st.form_submit_button("Save Changes")
                if submitted:
                    upd_data = {
                        "emp_id": new_emp_id.strip() or None,
                        "department": new_dept.strip() or None,
                        "role": new_role
                    }
                    if new_pwd:
                        upd_data["password"] = new_pwd
                        
                    try:
                        supabase.table("users").update(upd_data).eq("username", selected_username).execute()
                        st.success(f"Successfully updated @{selected_username}!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Failed to update user: {e}")

        with tab_reset:
            st.markdown("#### 🔄 Reset Face Data")
            st.warning("This will delete all registered face images for this user. They will have to re-register their face.")
            if st.button("Delete Face Dataset", key="reset_face_btn"):
                try:
                    files = supabase.storage.from_("faces").list(selected_username)
                    if files:
                        file_paths = [f"{selected_username}/{f['name']}" for f in files if f.get("name")]
                        supabase.storage.from_("faces").remove(file_paths)
                    # Clear session state cache
                    for k in [f"encs_{selected_username}", f"names_{selected_username}"]:
                        if k in st.session_state:
                            del st.session_state[k]
                    st.success(f"Face data reset completed for @{selected_username}.")
                except Exception as e:
                    st.error(f"Failed to reset face data: {e}")

        with tab_delete:
            st.markdown("#### ❌ Delete User")
            st.error("WARNING: This will delete this user and all associated face/attendance data. This action is IRREVERSIBLE.")
            if st.button("Delete User Account completely", key="delete_user_btn"):
                try:
                    # 1. Reset face storage
                    files = supabase.storage.from_("faces").list(selected_username)
                    if files:
                        file_paths = [f"{selected_username}/{f['name']}" for f in files if f.get("name")]
                        supabase.storage.from_("faces").remove(file_paths)
                    # 2. Delete attendance records
                    supabase.table("attendance").delete().eq("name", selected_username).execute()
                    # 3. Delete user
                    supabase.table("users").delete().eq("username", selected_username).execute()
                    # 4. Clear cache
                    for k in [f"encs_{selected_username}", f"names_{selected_username}"]:
                        if k in st.session_state:
                            del st.session_state[k]
                    st.success(f"User @{selected_username} has been completely deleted.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Failed to delete user: {e}")


# USER MANAGEMENT
def render_attendance_management():
    # ROLE CHECK
    if st.session_state.role != "admin":
        st.error("Access Denied: You do not have permission to access this page.")
        st.stop()

    st.markdown("""
    <div id="page-header">
      <div class="header-badge">📌</div>
      <div>
        <h2 style="margin:0;font-size:1.5rem;">Attendance Management</h2>
        <p style="margin:0;color:var(--silver);font-size:0.82rem;">View, correct, and manually log attendance</p>
      </div>
    </div>
    """, unsafe_allow_html=True)

    try:
        users_resp = supabase.table("users").select("username", "department").eq("role", "student").execute()
        students = users_resp.data or []
        student_names = [s["username"] for s in students]
        
        att_resp = supabase.table("attendance").select("*").order("date", desc=True).order("time", desc=True).execute()
        att_list = att_resp.data or []
    except Exception as e:
        st.error(f"Failed to load attendance details: {e}")
        return

    tab_view, tab_manual, tab_correct = st.tabs([
        "📋 View Attendance", "✍️ Mark Manually", "✏️ Correct / Delete"
    ])

    with tab_view:
        st.markdown("### 📋 Attendance Records")
        if not att_list:
            st.info("No attendance records found.")
        else:
            df_att = pd.DataFrame(att_list)
            c1, c2, c3 = st.columns(3)
            with c1:
                f_user = st.selectbox("Filter by Student", ["All"] + student_names)
            with c2:
                f_date = st.date_input("Filter by Date", value=None)
            with c3:
                f_status = st.selectbox("Filter by Status", ["All", "Present", "Absent"])
            
            df_filtered = df_att.copy()
            if f_user != "All":
                df_filtered = df_filtered[df_filtered["name"] == f_user]
            if f_date:
                df_filtered = df_filtered[df_filtered["date"] == str(f_date)]
            if f_status != "All":
                df_filtered = df_filtered[df_filtered["status"] == f_status.lower()]
                
            st.dataframe(df_filtered, use_container_width=True, hide_index=True)

    with tab_manual:
        st.markdown("### ✍️ Mark Attendance Manually")
        with st.form("manual_attendance_form"):
            selected_student = st.selectbox("Select Student", student_names)
            m_date = st.date_input("Date", value=datetime.date.today())
            m_time = st.time_input("Time", value=datetime.datetime.now().time())
            m_status = st.selectbox("Status", ["Present", "Absent"])
            submit_manual = st.form_submit_button("Mark Attendance")
            
            if submit_manual:
                dept = next((s["department"] for s in students if s["username"] == selected_student), "")
                try:
                    existing = supabase.table("attendance").select("*").eq("name", selected_student).eq("date", str(m_date)).execute()
                    if existing.data:
                        st.warning(f"Record already exists for {selected_student} on {m_date}. Use correction tab to modify.")
                    else:
                        supabase.table("attendance").insert({
                            "name": selected_student,
                            "date": str(m_date),
                            "time": str(m_time),
                            "status": m_status.lower(),
                            "marked_by": f"admin ({st.session_state.username})",
                            "department": dept
                        }).execute()
                        st.success(f"Successfully marked attendance for {selected_student}!")
                        st.rerun()
                except Exception as e:
                    st.error(f"Error marking attendance: {e}")

    with tab_correct:
        st.markdown("### ✏️ Correct or Delete Attendance")
        if not att_list:
            st.info("No records to correct.")
        else:
            record_options = [
                f"{r.get('id')} - {r.get('name')} ({r.get('date')} {r.get('time')[:8]}) [{r.get('status')}]"
                for r in att_list
            ]
            selected_record_desc = st.selectbox("Select Record to Modify/Delete", record_options)
            record_id = int(selected_record_desc.split(" - ")[0])
            
            rec = next((r for r in att_list if r["id"] == record_id), None)
            if rec:
                with st.form("correct_form"):
                    st.write(f"Modifying record for **{rec.get('name')}** on **{rec.get('date')}**")
                    c_date = st.date_input("Date", value=datetime.datetime.strptime(rec.get('date'), "%Y-%m-%d").date())
                    c_time = st.time_input("Time", value=datetime.datetime.strptime(rec.get('time')[:8], "%H:%M:%S").time())
                    c_status = st.selectbox("Status", ["Present", "Absent"], index=0 if rec.get('status') == 'present' else 1)
                    
                    col_save, col_del = st.columns(2)
                    with col_save:
                        save_btn = st.form_submit_button("💾 Save Changes")
                    with col_del:
                        delete_btn = st.form_submit_button("🗑️ Delete Record")
                        
                    if save_btn:
                        try:
                            supabase.table("attendance").update({
                                "date": str(c_date),
                                "time": str(c_time),
                                "status": c_status.lower(),
                                "marked_by": f"admin_edit ({st.session_state.username})"
                            }).eq("id", record_id).execute()
                            st.success("Record updated successfully!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Update failed: {e}")
                            
                    if delete_btn:
                        try:
                            supabase.table("attendance").delete().eq("id", record_id).execute()
                            st.success("Record deleted successfully!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Delete failed: {e}")


# SECURITY LOGS
def render_security_logs():
    # ROLE CHECK
    if st.session_state.role != "admin":
        st.error("Access Denied: You do not have permission to access this page.")
        st.stop()

    st.markdown("""
    <div id="page-header">
      <div class="header-badge">🚨</div>
      <div>
        <h2 style="margin:0;font-size:1.5rem;">Security Logs</h2>
        <p style="margin:0;color:var(--silver);font-size:0.82rem;">Audit trail of access violations, spoofing, and liveness failures</p>
      </div>
    </div>
    """, unsafe_allow_html=True)
    
    try:
        logs_resp = supabase.table("security_logs").select("*").order("timestamp", desc=True).execute()
        logs_list = logs_resp.data or []
    except Exception as e:
        st.error(f"Failed to fetch security logs: {e}")
        return
        
    if not logs_list:
        st.info("No security logs recorded yet.")
        return
        
    df_logs = pd.DataFrame(logs_list)
    df_logs["timestamp"] = pd.to_datetime(df_logs["timestamp"])
    
    c1, c2 = st.columns(2)
    with c1:
        f_type = st.selectbox("Event Type Filter", ["All", "multiple_faces", "failed_liveness", "unknown_face", "photo_spoofing"])
    with c2:
        f_user = st.text_input("Search Username", placeholder="All users...")
        
    df_filtered = df_logs.copy()
    if f_type != "All":
        df_filtered = df_filtered[df_filtered["event_type"] == f_type]
    if f_user:
        df_filtered = df_filtered[df_filtered["username"].str.contains(f_user, case=False, na=False)]
        
    st.dataframe(df_filtered, use_container_width=True, hide_index=True)
    
    if st.button("🗑️ Clear All Logs"):
        try:
            supabase.table("security_logs").delete().neq("id", 0).execute()
            st.success("All security logs cleared successfully!")
            st.rerun()
        except Exception as e:
            st.error(f"Failed to clear logs: {e}")


# ADMIN PANEL
def render_system_settings():
    # ROLE CHECK
    if st.session_state.role != "admin":
        st.error("Access Denied: You do not have permission to access this page.")
        st.stop()

    st.markdown("""
    <div id="page-header">
      <div class="header-badge">⚙️</div>
      <div>
        <h2 style="margin:0;font-size:1.5rem;">System Settings</h2>
        <p style="margin:0;color:var(--silver);font-size:0.82rem;">Adjust system rules and schedules</p>
      </div>
    </div>
    """, unsafe_allow_html=True)
    
    with st.form("sys_settings_form"):
        st.markdown("### ⏰ Attendance Window & Rules")
        new_start = st.time_input("Attendance Start Time", value=st.session_state.attendance_start)
        new_ends = st.time_input("Attendance End Time", value=st.session_state.attendance_ends)
        new_late = st.time_input("Late Warning Time", value=st.session_state.late_time)
        new_absent = st.time_input("Auto-Absent Cutoff Time (Bot)", value=st.session_state.absent_time)
        
        save_settings = st.form_submit_button("💾 Save System Settings")
        if save_settings:
            st.session_state.attendance_start = new_start
            st.session_state.attendance_ends = new_ends
            st.session_state.late_time = new_late
            st.session_state.absent_time = new_absent
            st.success("System configurations updated successfully for this session!")
            st.rerun()

# ─────────────────────────────────────────
#  SESSION STATE INIT
# ─────────────────────────────────────────
for k, v in [
    ("logged_in", False),
    ("username", ""),
    ("role", "student"),
    ("edit_mode", False),
    ("confirm_logout", False),
    ("attendance_start", datetime.time(9, 0)),
    ("attendance_ends", datetime.time(23, 0)),
    ("late_time", datetime.time(10, 30)),
    ("absent_time", datetime.time(19, 0))
]:
    if k not in st.session_state:
        st.session_state[k] = v

# Run auto-absent check (after session state is ready)
mark_auto_absent()

# ─────────────────────────────────────────
#  MENU
# ─────────────────────────────────────────
if st.session_state.logged_in:
    if st.session_state.role == "admin":
        st.sidebar.markdown("### 👑 Admin Panel")
        menu = st.sidebar.radio("Navigation", [
            "Dashboard", "User Management", "Attendance Management",
            "Security Logs", "System Settings", "🚪 Logout"
        ])
    else:
        menu = st.sidebar.radio("Navigation", [
            "👤 Personal Dashboard", "📸 Take Photo",
            "📌 Mark Attendance", "📊 Database", "🚪 Logout"
        ])
else:
    menu = st.sidebar.selectbox("Access", [ "Login", "Signup"])

# ─────────────────────────────────────────
#  PAGE TITLE
# ─────────────────────────────────────────
st.title("REAL-TIME FACE ATTENDANCE SYSTEM")

now_str = datetime.datetime.now().strftime("%A, %d %B %Y  ·  %I:%M %p")
st.markdown(f'<div class="datetime-strip">{now_str}</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────
#  SIGNUP
# ─────────────────────────────────────────
if menu == "Signup":
    if supabase is None:
        st.error("Database not available. Check Supabase configuration.")
        st.stop()

    col_c, col_form, col_d = st.columns([1, 2, 1])
    with col_form:
        st.markdown("""
        <div id="page-header">
          <div class="header-badge">✨</div>
          <div>
            <h2 style="margin:0;font-size:1.5rem;">Create Account</h2>
            <p style="margin:0;color:var(--silver);font-size:0.82rem;">Join the attendance system</p>
          </div>
        </div>
        """, unsafe_allow_html=True)

        with st.form("signup_form"):
            user = st.text_input("Username", placeholder="e.g. john.doe")
            pwd  = st.text_input("Password", type="password", placeholder="Choose a strong password")
            submitted = st.form_submit_button("Create Account →", use_container_width=True)

        if submitted:
            user = user.strip().lower()
            if user and pwd:
                try:
                    ex = supabase.table("users").select("username").eq("username", user).execute()
                    if ex.data:
                        st.warning("⚠️ Username already exists. Try a different one.")
                    else:
                        supabase.table("users").insert({"username": user, "password": pwd, "role": "student"}).execute()
                        st.success("🎉 Account created! You can now login.")
                except Exception as e:
                    st.error(f"Signup error: {e}")
            else:
                st.warning("Please fill in all fields.")

# ─────────────────────────────────────────
#  LOGIN
# ─────────────────────────────────────────
elif menu == "Login":
    if supabase is None:
        st.error("Database not available. Check Supabase configuration.")
        st.stop()

    col_c, col_form, col_d = st.columns([1, 2, 1])
    with col_form:
        st.markdown("""
        <div id="page-header">
          <div class="header-badge">🔐</div>
          <div>
            <h2 style="margin:0;font-size:1.5rem;">Welcome Back</h2>
            <p style="margin:0;color:var(--silver);font-size:0.82rem;">Sign in to your account</p>
          </div>
        </div>
        """, unsafe_allow_html=True)

        with st.form("login_form"):
            user = st.text_input("Username", placeholder="your username")
            pwd  = st.text_input("Password", type="password", placeholder="your password")
            submitted = st.form_submit_button("Sign In →", use_container_width=True)

        if submitted:
            user = user.strip().lower()
            if user and pwd:
                try:
                    res = supabase.table("users").select("*").eq("username", user).execute()
                    if res.data and res.data[0]["password"] == pwd:
                        st.session_state.logged_in = True
                        st.session_state.username  = user
                        st.session_state.role      = res.data[0].get("role", "student")
                        st.rerun()
                    else:
                        st.error("❌ Invalid credentials. Please try again.")
                except Exception as e:
                    st.error(f"Login error: {e}")
            else:
                st.warning("Enter both fields.")

# ─────────────────────────────────────────
#  PERSONAL DASHBOARD
# ─────────────────────────────────────────
elif menu == "👤 Personal Dashboard":
    if supabase is None:
        st.error("Database not available.")
        st.stop()

    username  = st.session_state.username
    u_resp    = supabase.table("users").select("*").eq("username", username).execute()
    user_info = u_resp.data[0] if u_resp.data else {}
    emp_id    = user_info.get("emp_id", "") or ""
    dept      = user_info.get("department", "") or ""

    st.markdown("""
    <div id="page-header">
      <div class="header-badge">👤</div>
      <div>
        <h2 style="margin:0;font-size:1.5rem;">My Profile</h2>
        <p style="margin:0;color:var(--silver);font-size:0.82rem;">Personal attendance dashboard</p>
      </div>
    </div>
    """, unsafe_allow_html=True)

    if "profile_success" in st.session_state:
        st.success(st.session_state.profile_success)
        del st.session_state["profile_success"]

    st.markdown(f"""
    <div class="info-row">
      <div class="info-item" style="margin-right:2rem;">
        <span class="info-label">Username</span>
        <span class="info-value">@{username}</span>
      </div>
      <div class="info-item" style="margin-right:2rem;">
        <span class="info-label">Employee ID</span>
        <span class="info-value">{emp_id if emp_id else '—'}</span>
      </div>
      <div class="info-item">
        <span class="info-label">Department</span>
        <span class="info-value">{dept if dept else '—'}</span>
      </div>
    </div>
    """, unsafe_allow_html=True)

    if st.button("✏️  Edit Profile"):
        st.session_state.edit_mode = not st.session_state.edit_mode
  
    if st.session_state.edit_mode:
        with st.form("edit_profile"):
            st.markdown("##### ✏️ Update Profile Details")
            new_emp  = st.text_input("Employee ID", value=emp_id, placeholder="e.g. EMP-001")
            new_dept = st.text_input("Department",  value=dept,   placeholder="e.g. Engineering")
            new_pwd  = st.text_input("New Password (leave blank to keep current)", type="password")
            s_col, c_col = st.columns(2)
            with s_col: save   = st.form_submit_button("💾 Save Changes", use_container_width=True)
            with c_col: cancel = st.form_submit_button("Cancel",          use_container_width=True)

            if save:
                upd = {"emp_id": new_emp, "department": new_dept}
                if new_pwd: upd["password"] = new_pwd
                try:
                    res = supabase.table("users").update(upd).eq("username", username).execute()
                    if not res.data:
                        st.error("Update failed! Check RLS policy.")
                    else:
                        st.session_state.profile_success = "✅ Profile updated successfully!"
                        st.session_state.edit_mode = False
                        st.rerun()
                except Exception as e:
                    st.error(f"Database error: {e}")
            if cancel:
                st.session_state.edit_mode = False
                st.rerun()

    st.divider()

    att_resp = supabase.table("attendance").select("*").eq("name", username).execute()
    if not att_resp.data:
        st.info("No attendance records found yet.")
        st.stop()

    df = pd.DataFrame(att_resp.data)
    df["date"] = pd.to_datetime(df["date"]).dt.date
    if "status" not in df.columns:
        df["status"] = "present"
    else:
        df["status"] = df["status"].fillna("present")
    df.loc[df["marked_by"] == "system", "status"] = "absent"
    df["_priority"] = df["status"].apply(lambda s: 0 if s == "present" else 1)
    df = df.sort_values("_priority").drop_duplicates(subset="date", keep="first").drop(columns="_priority")
    df["year"] = pd.to_datetime(df["date"]).dt.year

    today = datetime.date.today()
    years = sorted(df["year"].unique(), reverse=True)

    st.markdown("""
    <div id="page-header" style="border:none;padding-bottom:0.5rem;margin-bottom:1rem;">
      <div class="header-badge" style="width:40px;height:40px;font-size:1.2rem;">📊</div>
      <h2 style="margin:0;font-size:1.4rem;">Attendance Statistics</h2>
    </div>
    """, unsafe_allow_html=True)

    selected_year = st.selectbox("Select Year", years, index=0)
    df_year       = df[df["year"] == selected_year]
    year_start    = datetime.date(selected_year, 1, 1)
    year_end      = min(datetime.date(selected_year, 12, 31), today)
    total_working = working_days_in_range(year_start, year_end)
    present_df    = df_year[df_year["status"] != "absent"]
    present_days  = len(present_df)
    absent_days   = max(total_working - present_days, 0)
    holidays_cnt  = sum(1 for d in INDIA_HOLIDAYS if d.year == selected_year and year_start <= d <= year_end)
    pct           = attendance_pct(present_days, total_working)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("✅ Present Days",  present_days)
    m2.metric("❌ Absent Days",   absent_days)
    m3.metric("🎌 Holidays",      holidays_cnt)
    m4.metric("📈 Attendance %",  f"{pct}%",
              delta="Good ↑" if pct >= 75 else "Low ↓",
              delta_color="normal" if pct >= 75 else "inverse")

    st.divider()

    chart_col1, chart_col2 = st.columns(2)

    with chart_col1:
        st.markdown("**Attendance Overview**")
        total_d      = present_days + absent_days
        present_frac = (present_days / total_d) if total_d > 0 else 0
        p_end        = present_frac * 360
        cx, cy, ro, ri = 110, 110, 85, 54

        if present_days > 0 and absent_days > 0:
            slices_svg = (
                f"<path d='{arc_path(cx,cy,ro,0,p_end)} "
                f"L {cx+ri*math.cos(math.radians(p_end-90)):.2f} {cy+ri*math.sin(math.radians(p_end-90)):.2f} "
                f"{arc_path(cx,cy,ri,p_end,0)[2:]} Z' fill='#10b981' opacity='0.92'/>"
                f"<path d='{arc_path(cx,cy,ro,p_end,360)} "
                f"L {cx+ri*math.cos(math.radians(360-90)):.2f} {cy+ri*math.sin(math.radians(360-90)):.2f} "
                f"{arc_path(cx,cy,ri,360,p_end)[2:]} Z' fill='#f43f5e' opacity='0.85'/>"
            )
        elif present_days > 0:
            slices_svg = f"<circle cx='{cx}' cy='{cy}' r='{ro}' fill='#10b981' opacity='0.92'/>"
        else:
            slices_svg = f"<circle cx='{cx}' cy='{cy}' r='{ro}' fill='#f43f5e' opacity='0.85'/>"

        donut_svg = f"""
<svg viewBox="0 0 220 220" xmlns="http://www.w3.org/2000/svg" width="220" height="220">
  <defs><filter id="shadow" x="-20%" y="-20%" width="140%" height="140%">
    <feDropShadow dx="0" dy="4" stdDeviation="8" flood-color="rgba(0,0,0,0.4)"/>
  </filter></defs>
  <g filter="url(#shadow)">{slices_svg}<circle cx="{cx}" cy="{cy}" r="{ri}" fill="#0a0a0c"/></g>
  <text x="{cx}" y="{cy-8}" text-anchor="middle" font-size="22" font-weight="700"
        font-family="Playfair Display,serif" fill="#ffffff">{pct}%</text>
  <text x="{cx}" y="{cy+13}" text-anchor="middle" font-size="10"
        font-family="Space Grotesk,sans-serif" fill="#6e6e7a" letter-spacing="2">ATTENDANCE</text>
  <rect x="24" y="200" width="10" height="10" fill="#10b981" rx="3"/>
  <text x="40" y="209" font-size="10" font-family="Inter,sans-serif" fill="#8e8e9a">Present ({present_days}d)</text>
  <rect x="128" y="200" width="10" height="10" fill="#f43f5e" rx="3"/>
  <text x="144" y="209" font-size="10" font-family="Inter,sans-serif" fill="#8e8e9a">Absent ({absent_days}d)</text>
</svg>"""
        st.markdown(donut_svg, unsafe_allow_html=True)

    with chart_col2:
        st.markdown("**Monthly Distribution (Present Days)**")
        df_m = present_df.copy()
        if not df_m.empty:
            df_m["month"] = pd.to_datetime(df_m["date"]).dt.month
            monthly = df_m.groupby("month").size().reindex(range(1, 13), fill_value=0)
        else:
            monthly = pd.Series([0]*12, index=range(1, 13))

        mnames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
        active  = [(mnames[i], int(v)) for i, v in enumerate(monthly) if v > 0]
        total_m = sum(v for _, v in active)
        pie_colors = ["#10b981","#f59e0b","#3b82f6","#f43f5e",
                      "#059669","#fbbf24","#2563eb","#e11d48",
                      "#34d399","#d97706","#60a5fa","#fb7185"]
        pcx, pcy, pr = 110, 100, 78

        if total_m > 0:
            pie_paths = ""
            leg_items = ""
            angle     = 0
            for idx, (mname, val) in enumerate(active):
                sweep    = (val / total_m) * 360
                end_ang  = angle + sweep
                col      = pie_colors[idx % len(pie_colors)]
                pie_paths += (
                    f"<path d='{arc_path(pcx,pcy,pr,angle,end_ang)} "
                    f"L {pcx} {pcy} Z' fill='{col}' stroke='#0a0a0c' stroke-width='1.5' opacity='0.90'/>"
                )
                lx = 10 + (idx % 2) * 105
                ly = 195 + (idx // 2) * 14
                leg_items += (
                    f"<rect x='{lx}' y='{ly-9}' width='9' height='9' fill='{col}' rx='2'/>"
                    f"<text x='{lx+12}' y='{ly}' font-size='9' font-family='Inter,sans-serif' fill='#8e8e9a'>{mname}: {val}d</text>"
                )
                angle = end_ang

            pie_svg = f"""
<svg viewBox="0 0 220 {195 + ((len(active)-1)//2 + 1)*14 + 10}"
     xmlns="http://www.w3.org/2000/svg" width="220">
  <defs><filter id="pieShadow" x="-20%" y="-20%" width="140%" height="140%">
    <feDropShadow dx="0" dy="4" stdDeviation="6" flood-color="rgba(0,0,0,0.4)"/>
  </filter></defs>
  <g filter="url(#pieShadow)">{pie_paths}</g>{leg_items}
</svg>"""
            st.markdown(pie_svg, unsafe_allow_html=True)
        else:
            st.info("No data for selected year.")

    st.divider()

    st.markdown("""
    <div id="page-header" style="border:none;padding-bottom:0.5rem;margin-bottom:1rem;">
      <div class="header-badge" style="width:40px;height:40px;font-size:1.2rem;">📅</div>
      <h2 style="margin:0;font-size:1.4rem;">Year-wise Summary</h2>
    </div>
    """, unsafe_allow_html=True)

    yearly = []
    for yr in years:
        df_yr = df[df["year"] == yr]
        ys    = datetime.date(yr, 1, 1)
        ye    = min(datetime.date(yr, 12, 31), today)
        wd    = working_days_in_range(ys, ye)
        pr_df = df_yr[df_yr["status"] != "absent"]
        pr    = len(pr_df)
        yearly.append({"Year": str(yr), "Present": pr, "Working Days": wd, "Attendance %": attendance_pct(pr, wd)})

    stats_df = pd.DataFrame(yearly).set_index("Year")
    st.bar_chart(stats_df[["Present", "Working Days"]])
    st.dataframe(stats_df.style.format({"Attendance %": "{:.1f}%"}), use_container_width=True)

    st.divider()
    st.markdown("**📋 Attendance Records**")
    show = df_year[["name", "date", "time", "status"]].sort_values("date", ascending=False)
    show["status"] = show["status"].str.upper()
    st.dataframe(show, use_container_width=True, hide_index=True)

# ─────────────────────────────────────────
#  TAKE PHOTO
# ─────────────────────────────────────────
elif menu == "📸 Take Photo":
    if supabase is None:
        st.error("Database not available.")
        st.stop()

    st.markdown("""
    <div id="page-header">
      <div class="header-badge">📸</div>
      <div>
        <h2 style="margin:0;font-size:1.5rem;">Capture Face</h2>
        <p style="margin:0;color:var(--silver);font-size:0.82rem;">Upload your photo for face recognition</p>
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div style="background: rgba(16, 185, 129, 0.05); border: 1px solid rgba(16, 185, 129, 0.2); padding: 1.5rem; border-radius: 12px; margin-bottom: 1.5rem;">
        <h4 style="color: var(--emerald); margin-top: 0; margin-bottom: 0.8rem; display: flex; align-items: center; gap: 8px; font-family: var(--font-heading); font-size: 0.95rem; text-transform: uppercase; letter-spacing: 0.05em;">
            📋 Guidelines for a Perfect Profile Photo
        </h4>
        <p style="margin: 0 0 0.8rem 0; font-size: 0.85rem; color: var(--silver); line-height: 1.6;">
            To ensure the real-time liveness verification system recognizes you and prevents spoofing via phone/printed photos, please make sure your registration photo meets these criteria:
        </p>
        <ul style="margin: 0; padding-left: 1.2rem; font-size: 0.85rem; color: var(--bone); line-height: 1.6;">
            <li><strong>Proper Lighting:</strong> Stand in a well-lit area. Avoid harsh backlighting or deep shadows on your face.</li>
            <li><strong>Look Straight:</strong> Face the camera directly. Do not tilt your head up, down, or sideways.</li>
            <li><strong>Neutral Expression:</strong> Keep a relaxed, neutral face expression (lips closed, eyes fully open).</li>
            <li><strong>Clear Visibility:</strong> Ensure your eyes, eyebrows, nose, and mouth are fully visible. Remove sunglasses, masks, or hats.</li>
            <li><strong>No Screens/Photos:</strong> Always take a direct photo of your real face. Uploading a picture of a phone screen or physical photo will fail our secure liveness verification during attendance marking.</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

    img = st.camera_input("Point your face at the camera and click 📸")
    if img is not None:
        fn = f"{st.session_state.username}/{st.session_state.username}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
        try:
            supabase.storage.from_("faces").upload(fn, img.getvalue(), {"content-type": "image/jpeg"})
            # Clear face cache to force reload on next attendance check
            cache_key_encs = f"encs_{st.session_state.username}"
            cache_key_names = f"names_{st.session_state.username}"
            if cache_key_encs in st.session_state:
                del st.session_state[cache_key_encs]
            if cache_key_names in st.session_state:
                del st.session_state[cache_key_names]
            st.success("✅ Photo uploaded successfully!")
        except Exception as e:
            st.error(f"Upload failed: {e}")

# ─────────────────────────────────────────
#  MARK ATTENDANCE
# ─────────────────────────────────────────
elif menu == "📌 Mark Attendance":
    if supabase is None:
        st.error("Database not available.")
        st.stop()

    if not FACE_RECOGNITION_AVAILABLE:
        st.error("⚠️ `face_recognition` and `opencv-python` are required. Install them with:\n```\npip install face_recognition opencv-python\n```")
        st.stop()

    st.markdown("""
    <div id="page-header">
      <div class="header-badge">📌</div>
      <div>
        <h2 style="margin:0;font-size:1.5rem;">Mark Attendance</h2>
        <p style="margin:0;color:var(--silver);font-size:0.82rem;">Face recognition attendance system</p>
      </div>
    </div>
    """, unsafe_allow_html=True)

    try:
        files = supabase.storage.from_("faces").list(st.session_state.username)
    except Exception as e:
        st.error(f"Storage error: {e}")
        st.stop()

    if not files:
        st.warning("No face images found. Upload photos first in **📸 Take Photo**.")
        st.stop()

    cache_key_encs = f"encs_{st.session_state.username}"
    cache_key_names = f"names_{st.session_state.username}"

    if cache_key_encs not in st.session_state or cache_key_names not in st.session_state:
        known_encodings, known_names = [], []
        with st.spinner("Loading face dataset..."):
            for file in files:
                fname = file.get("name")
                if not fname: continue
                url = supabase.storage.from_("faces").get_public_url(f"{st.session_state.username}/{fname}")
                try:
                    r   = requests.get(url, timeout=10); r.raise_for_status()
                    arr = np.asarray(bytearray(r.content), dtype=np.uint8)
                    img = cv2.imdecode(arr, cv2.IMREAD_COLOR)
                    if img is None: continue
                    enc = face_recognition.face_encodings(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
                    if enc:
                        known_encodings.append(enc[0])
                        known_names.append(st.session_state.username)
                except Exception:
                    continue
        st.session_state[cache_key_encs] = known_encodings
        st.session_state[cache_key_names] = known_names
    else:
        known_encodings = st.session_state[cache_key_encs]
        known_names = st.session_state[cache_key_names]

    if not known_encodings:
        st.error("No valid face encodings found.")
        st.stop()

    st.success(f"✅ Loaded {len(known_encodings)} face(s).")
    u_resp    = supabase.table("users").select("department").eq("username", st.session_state.username).execute()
    user_dept = (u_resp.data[0].get("department", "") if u_resp.data else "") or ""

    now_time   = datetime.datetime.now().time()
    in_window  = ATTENDANCE_START <= now_time <= ATTENDANCE_ENDS
    now        = datetime.datetime.now()
    today_date = str(now.date())

    if not in_window:
        st.warning(
            f"⏰ Attendance camera is only available between "
            f"{ATTENDANCE_START.strftime('%I:%M %p')} and "
            f"{ATTENDANCE_ENDS.strftime('%I:%M %p')}. "
            f"Current time: {now_time.strftime('%I:%M %p')}"
        )
        st.stop()

    if not WEBRTC_AVAILABLE:
        st.error("⚠️ `streamlit-webrtc` is required for continuous liveness verification.")
        st.stop()
        
    st.markdown("""
    <div style="background: rgba(16, 185, 129, 0.05); border: 1px solid rgba(16, 185, 129, 0.2); padding: 1rem 1.4rem; border-radius: 12px; margin-bottom: 1rem;">
        <p style="margin: 0; font-size: 0.85rem; color: var(--bone); line-height: 1.6;">
            📷 <strong style="color: var(--emerald);">Camera is active.</strong> Watch the video feed for on-screen prompts. A random liveness challenge (Smile, Blink, Turn, etc.) will appear — follow it to verify you're a real person. Once you see <strong style="color: #10b981;">"Verified!"</strong> on screen, click the button below.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # ── WebRTC config: works over tunnels (ngrok / Cloudflare / etc.) ──────
    # media_stream_constraints forces the BROWSER (remote device) to open
    # its own camera — NOT the server camera. This is the key fix for tunneling.
    # Multiple public STUN servers improve ICE negotiation over tunnels.
    webrtc_ctx = webrtc_streamer(
        key="attendance-liveness",
        video_processor_factory=FaceLivenessProcessor,
        rtc_configuration=RTCConfiguration({
            "iceServers": [
                {"urls": ["stun:stun.l.google.com:19302"]},
                {"urls": ["stun:stun1.l.google.com:19302"]},
                {"urls": ["stun:stun2.l.google.com:19302"]},
                {"urls": ["stun:openrelay.metered.ca:80"]},
                {
                    "urls":       [
                        "turn:openrelay.metered.ca:80",
                        "turn:openrelay.metered.ca:443",
                        "turns:openrelay.metered.ca:443?transport=tcp"
                    ],
                    "username":   "openrelayproject",
                    "credential": "openrelayproject",
                },
            ]
        }),
        media_stream_constraints={
            "video": {
                "width":      {"ideal": 640},
                "height":     {"ideal": 480},
                "facingMode": "user",   # front/selfie cam on phones
            },
            "audio": False,
        },
        async_processing=True,
    )

    # ── Always push attributes into the processor (it may start mid-run) ────
    if webrtc_ctx.video_processor:
        webrtc_ctx.video_processor.known_encodings = known_encodings
        webrtc_ctx.video_processor.known_names     = known_names
        webrtc_ctx.video_processor.supabase        = supabase
        webrtc_ctx.video_processor.user_dept       = user_dept
        webrtc_ctx.video_processor.username        = st.session_state.username
        webrtc_ctx.video_processor.today_date      = today_date

        # ── Persist verification result in session state ─────────────────────
        if webrtc_ctx.video_processor:
            with webrtc_ctx.video_processor.lock:
                if webrtc_ctx.video_processor.verification_success:
                    st.session_state["att_verified"]      = True
                    st.session_state["att_verified_name"] = webrtc_ctx.video_processor.verified_name

        # Show persistent success banner if already verified this session
        if st.session_state.get("att_verified"):
            verified_name_display = st.session_state.get("att_verified_name", "")
            st.success(f"🎉 **Attendance already marked for {verified_name_display} today!**")
            # if st.button("🔄 Mark Again (different user)", use_container_width=True):
                # st.session_state.pop("att_verified", None)
                # st.session_state.pop("att_verified_name", None)
                # st.rerun()
        else:
            if st.button("✅ Check Verification Status", use_container_width=True):
                if webrtc_ctx.video_processor:
                    with webrtc_ctx.video_processor.lock:
                        success = webrtc_ctx.video_processor.verification_success
                        name    = webrtc_ctx.video_processor.verified_name
                    if success:
                        st.session_state["att_verified"]      = True
                        st.session_state["att_verified_name"] = name
                        save_attendance_excel()
                        st.success(f"🎉 **Verified and marked attendance for {name}!**")
                        st.balloons()
                    else:
                        st.warning("⏳ Not verified yet. Keep following the on-screen challenge and try again.")
                else:
                    st.warning("⏳ Camera not started yet — click START above first.")

# ─────────────────────────────────────────
#  DATABASE
# ─────────────────────────────────────────
elif menu == "📊 Database":
    if supabase is None:
        st.error("Database not available.")
        st.stop()

    st.markdown("""
    <div id="page-header">
      <div class="header-badge">📊</div>
      <div>
        <h2 style="margin:0;font-size:1.5rem;">Attendance Records</h2>
        <p style="margin:0;color:var(--silver);font-size:0.82rem;">Full database view with filters</p>
      </div>
    </div>
    """, unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1: fd = st.date_input("Filter by date", value=None)
    with c2: fn = st.text_input("Filter by name", placeholder="All employees")

    try:
        q = supabase.table("attendance").select("*").order("date", desc=True).order("time", desc=True)
        if fd: q = q.eq("date", str(fd))
        if fn: q = q.ilike("name", f"%{fn}%")
        data = q.execute()

        if data.data:
            df_db = pd.DataFrame(data.data)
            if "status" not in df_db.columns:
                df_db["status"] = "present"
            else:
                df_db["status"] = df_db["status"].fillna("present")
            df_db.loc[df_db["marked_by"] == "system", "status"] = "absent"
            df_db["status"] = df_db["status"].str.upper()
            cols  = ["id", "name", "date", "time", "status", "marked_by", "department"]
            df_db = df_db[[c for c in cols if c in df_db.columns]]
            st.dataframe(df_db, use_container_width=True)
            st.markdown(f"""<p style="font-family:var(--font-mono);font-size:0.78rem;color:var(--silver);letter-spacing:0.05em;margin-top:0.3rem;">
              Total records: <strong style="color:var(--emerald);">{len(data.data)}</strong></p>""",
              unsafe_allow_html=True)

            if st.button("⬇  Save / Refresh Excel"):
                ok, res = save_attendance_excel()
                st.success(f"Saved to `{res}`") if ok else st.error(f"Failed: {res}")
        else:
            st.info("No records found.")
    except Exception as e:
        st.error(f"Error: {e}")

# ADMIN PANEL
elif menu == "Dashboard":
    # ROLE CHECK
    if st.session_state.role != "admin":
        st.error("Access Denied: You do not have permission to access this page.")
        st.stop()
    render_admin_dashboard()

elif menu == "User Management":
    # ROLE CHECK
    if st.session_state.role != "admin":
        st.error("Access Denied: You do not have permission to access this page.")
        st.stop()
    render_user_management()

elif menu == "Attendance Management":
    # ROLE CHECK
    if st.session_state.role != "admin":
        st.error("Access Denied: You do not have permission to access this page.")
        st.stop()
    render_attendance_management()

elif menu == "Security Logs":
    # ROLE CHECK
    if st.session_state.role != "admin":
        st.error("Access Denied: You do not have permission to access this page.")
        st.stop()
    render_security_logs()

elif menu == "System Settings":
    # ROLE CHECK
    if st.session_state.role != "admin":
        st.error("Access Denied: You do not have permission to access this page.")
        st.stop()
    render_system_settings()

# ─────────────────────────────────────────
#  LOGOUT CONFIRMATION PAGE
# ─────────────────────────────────────────
elif menu == "🚪 Logout":
    st.markdown("""
    <style>
    .logout-overlay{
        display:flex;align-items:center;justify-content:center;
        min-height:60vh;padding:2rem 0;
    }
    .logout-card{
        background:rgba(22,22,28,0.85);
        backdrop-filter:blur(40px) saturate(150%);
        border:1px solid rgba(255,255,255,0.09);
        border-radius:28px;
        padding:3.5rem 4rem 3rem;
        max-width:480px;width:100%;
        text-align:center;
        box-shadow:0 40px 100px rgba(0,0,0,0.6),inset 0 1px 0 rgba(255,255,255,0.06);
        position:relative;overflow:hidden;
    }
    .logout-card::before{
        content:'';position:absolute;top:0;left:0;right:0;height:3px;
        background:linear-gradient(90deg,var(--rose,#f43f5e),var(--amber,#f59e0b),var(--emerald,#10b981));
        border-radius:3px 3px 0 0;
    }
    .logout-icon{
        font-size:4rem;margin-bottom:1.2rem;
        display:inline-block;
        animation:logoutWave 2.5s ease-in-out infinite;
    }
    @keyframes logoutWave{
        0%,100%{transform:rotate(0deg);}
        20%{transform:rotate(-15deg);}
        40%{transform:rotate(12deg);}
        60%{transform:rotate(-8deg);}
        80%{transform:rotate(5deg);}
    }
    .logout-title{
        font-family:var(--font-display,'Playfair Display',serif)!important;
        font-size:1.9rem!important;font-weight:700!important;
        color:var(--white,#fff)!important;
        margin:0 0 0.5rem!important;letter-spacing:-0.02em!important;
    }
    .logout-subtitle{
        font-size:0.95rem!important;
        color:var(--pewter,#8e8e9a)!important;
        line-height:1.65!important;margin-bottom:0.8rem!important;
    }
    .logout-user-badge{
        display:inline-flex;align-items:center;gap:8px;
        padding:7px 20px;margin:0.9rem 0 2rem;
        background:rgba(16,185,129,0.08);
        border:1px solid rgba(16,185,129,0.2);
        border-radius:9999px;
        font-size:0.82rem!important;
        color:var(--emerald,#10b981)!important;
        letter-spacing:0.04em;
    }
    .logout-user-badge .dot{
        width:7px;height:7px;border-radius:50%;
        background:#10b981;
        animation:dotPulse 2s ease-in-out infinite;
    }
    </style>

    <div class="logout-overlay">
      <div class="logout-card">
        <div class="logout-icon">👋</div>
        <div class="logout-title">Leaving so soon?</div>
        <div class="logout-subtitle">
          Are you sure you want to log out?<br>
          Your session and progress are safe — you can always come back.
        </div>
        <div class="logout-user-badge">
          <span class="dot"></span>
          """ + f"Logged in as &nbsp;<strong>{st.session_state.username}</strong>" + """
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # Centered buttons
    _, col1, col2, _ = st.columns([1.5, 1, 1, 1.5])

    with col1:
        if st.button("✅  Stay Here", use_container_width=True, type="secondary"):
            st.session_state.confirm_logout = False
            st.rerun()

    with col2:
        if st.button("🚪  Yes, Logout", use_container_width=True):
            st.session_state.logged_in      = False
            st.session_state.username       = ""
            st.session_state.confirm_logout = False
            st.session_state.edit_mode      = False
            st.rerun()