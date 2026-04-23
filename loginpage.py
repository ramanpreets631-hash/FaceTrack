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

# Optional heavy imports — wrapped so app doesn't crash if missing
try:
    import cv2
    import face_recognition
    FACE_RECOGNITION_AVAILABLE = True
except ImportError:
    FACE_RECOGNITION_AVAILABLE = False

try:
    from supabase import create_client, Client
    SUPABASE_AVAILABLE = True
except ImportError:
    SUPABASE_AVAILABLE = False

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
SUPABASE_URL = "https://jmjdbrqoilxkrtfhlmuw.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImptamRicnFvaWx4a3J0ZmhsbXV3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQzNjY5NTgsImV4cCI6MjA4OTk0Mjk1OH0.ccQUa3UCk32QA992ZhNbNb3Rk0c_J22bOwIuhCQdvl0"
EXCEL_FOLDER = "attendance_exports"
EXCEL_FILE   = os.path.join(EXCEL_FOLDER, "attendance.xlsx")
ATTENDANCE_START = datetime.time(9, 0)
ATTENDANCE_ENDS  = datetime.time(20, 0)
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
@st.cache_resource
def init_supabase():
    if not SUPABASE_AVAILABLE:
        return None
    try:
        return create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception as e:
        st.error(f"Supabase connection error: {e}")
        return None

supabase = init_supabase()

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
#  SESSION STATE INIT
# ─────────────────────────────────────────
for k, v in [("logged_in", False), ("username", ""), ("edit_mode", False), ("confirm_logout", False)]:
    if k not in st.session_state:
        st.session_state[k] = v

# Run auto-absent check (after session state is ready)
mark_auto_absent()

# ─────────────────────────────────────────
#  MENU
# ─────────────────────────────────────────
if st.session_state.logged_in:
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
                        supabase.table("users").insert({"username": user, "password": pwd}).execute()
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

    img = st.camera_input("Point your face at the camera and click 📸")
    if img is not None:
        fn = f"{st.session_state.username}/{st.session_state.username}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
        try:
            supabase.storage.from_("faces").upload(fn, img.getvalue(), {"content-type": "image/jpeg"})
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

    st.info("📷 Your **browser camera** will be used.")
    snap = st.camera_input("Point your face at the camera and click 📸 to mark attendance")

    if snap is not None:
        arr   = np.asarray(bytearray(snap.getvalue()), dtype=np.uint8)
        frame = cv2.imdecode(arr, cv2.IMREAD_COLOR)

        if frame is None:
            st.error("Could not decode the captured image. Please try again.")
            st.stop()

        small = cv2.resize(frame, (0, 0), fx=0.5, fy=0.5)
        rgb_s = cv2.cvtColor(small, cv2.COLOR_BGR2RGB)
        locs  = face_recognition.face_locations(rgb_s)
        encs  = face_recognition.face_encodings(rgb_s, locs)

        if not locs:
            st.warning("⚠️ No face detected. Please try again with better lighting.")
        else:
            marked = []
            for encode, loc in zip(encs, locs):
                matches = face_recognition.compare_faces(known_encodings, encode, tolerance=0.5)
                dists   = face_recognition.face_distance(known_encodings, encode)
                name    = "UNKNOWN"
                color   = (94, 63, 244)

                if True in matches:
                    idx   = int(np.argmin(dists))
                    name  = known_names[idx]
                    color = (129, 185, 16)

                    existing = supabase.table("attendance")\
                        .select("*").eq("name", name).eq("date", today_date).execute()

                    existing_rows   = existing.data or []
                    already_present = any(
                        r.get("status") == "present" and r.get("marked_by") != "system"
                        for r in existing_rows
                    )

                    if name not in marked and not already_present:
                        now_dt = datetime.datetime.now()
                        db_ok  = False
                        try:
                            res = supabase.table("attendance").insert({
                                "name":       name,
                                "date":       str(now_dt.date()),
                                "time":       str(now_dt.time()),
                                "marked_by":  st.session_state.username,
                                "department": user_dept,
                                "status":     "present"
                            }).execute()
                            if res.data:
                                db_ok = True
                            else:
                                st.error(f"❌ DB INSERT returned no rows for **{name}**. Check Supabase RLS.")
                        except Exception as e:
                            st.error(f"❌ Database error for **{name}**: {e}")

                        if db_ok:
                            marked.append(name)
                            ok, path = save_attendance_excel()
                            st.toast(f"✅ {name} marked present — Excel {'saved' if ok else 'failed'}")

                    elif already_present:
                        st.info(f"ℹ️ {name} already marked present for today.")

                y1, x2, y2, x1 = [v * 2 for v in loc]
                cv2.rectangle(frame, (x1, y1), (x2, y2), color, 2)
                cv2.rectangle(frame, (x1, y2 - 30), (x2, y2), color, cv2.FILLED)
                cv2.putText(frame, name, (x1 + 6, y2 - 6),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 255), 2)

            st.image(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB),
                     channels="RGB", use_container_width=True,
                     caption="Recognition result")

            if marked:
                st.success(f"✅ Marked present: {', '.join(marked)}")
            else:
                all_unknown = all(
                    not any(face_recognition.compare_faces(known_encodings, e, tolerance=0.5))
                    for e in encs
                )
                if all_unknown:
                    st.error("❌ Face not recognised. Make sure you've uploaded your photo first.")

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