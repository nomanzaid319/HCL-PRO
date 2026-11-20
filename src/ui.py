"""
Insurance Claim Processing Analysis - UI Framework & Tailwind Component Library
Author: Noman Zaid (Roll No: 2500320140103, MCA 2nd Year Data Analytics, ABES)
Tech: Streamlit Native CSS Injection, Tailwind CSS CDN via iframe, Interactive Components
"""

import streamlit as st
import streamlit.components.v1 as components

def inject_custom_theme():
    """
    Injects executive, modern CSS directly into the Streamlit DOM for a premium,
    attention-grabbing InsurTech look (custom fonts, glassmorphism, glowing accents, and smooth hover effects).
    """
    custom_css = """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    /* Main background subtle gradient */
    .stApp {
        background: linear-gradient(180deg, #F8FAFC 0%, #EDF2F7 100%);
    }

    /* Modern Card Container */
    div[data-testid="stVerticalBlock"] > div[style*="border"] {
        border-radius: 16px !important;
        box-shadow: 0 4px 20px -2px rgba(0, 0, 0, 0.05) !important;
        transition: all 0.25s ease-in-out !important;
    }
    div[data-testid="stVerticalBlock"] > div[style*="border"]:hover {
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1) !important;
        transform: translateY(-2px) !important;
    }

    /* Streamlit Sidebar Custom Styling */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0F172A 0%, #1E293B 100%) !important;
        color: #F8FAFC !important;
        border-right: 1px solid rgba(255, 255, 255, 0.08) !important;
    }
    section[data-testid="stSidebar"] h1, 
    section[data-testid="stSidebar"] h2, 
    section[data-testid="stSidebar"] h3, 
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] .stMarkdown p {
        color: #E2E8F0 !important;
    }
    section[data-testid="stSidebar"] hr {
        border-color: rgba(255, 255, 255, 0.12) !important;
    }

    /* Sidebar Navigation Links */
    div[data-testid="stSidebarNav"] li div a {
        border-radius: 10px !important;
        transition: all 0.2s ease !important;
        color: #CBD5E1 !important;
        font-weight: 500 !important;
    }
    div[data-testid="stSidebarNav"] li div a:hover {
        background: rgba(59, 130, 246, 0.15) !important;
        color: #60A5FA !important;
        transform: translateX(4px) !important;
    }
    div[data-testid="stSidebarNav"] li div[aria-selected="true"] a {
        background: linear-gradient(90deg, rgba(37, 99, 235, 0.3) 0%, rgba(37, 99, 235, 0.1) 100%) !important;
        color: #93C5FD !important;
        border-left: 3px solid #3B82F6 !important;
        font-weight: 700 !important;
    }

    /* Primary Buttons Styling */
    .stButton > button {
        border-radius: 12px !important;
        font-weight: 600 !important;
        letter-spacing: 0.02em !important;
        transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1) !important;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.2) !important;
    }
    .stButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 8px 20px rgba(37, 99, 235, 0.35) !important;
    }

    /* Download Buttons */
    .stDownloadButton > button {
        border-radius: 12px !important;
        font-weight: 600 !important;
        background: linear-gradient(135deg, #10B981 0%, #059669 100%) !important;
        color: #FFFFFF !important;
        border: none !important;
        box-shadow: 0 4px 14px rgba(16, 185, 129, 0.3) !important;
        transition: all 0.2s ease !important;
    }
    .stDownloadButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(16, 185, 129, 0.45) !important;
    }

    /* Metric cards styling */
    [data-testid="stMetric"] {
        background: #FFFFFF !important;
        padding: 16px 20px !important;
        border-radius: 14px !important;
        border: 1px solid #E2E8F0 !important;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04) !important;
        transition: all 0.2s ease !important;
    }
    [data-testid="stMetric"]:hover {
        border-color: #3B82F6 !important;
        box-shadow: 0 8px 16px rgba(59, 130, 246, 0.08) !important;
    }

    /* Custom Scrollbars */
    ::-webkit-scrollbar {
        width: 8px;
        height: 8px;
    }
    ::-webkit-scrollbar-track {
        background: #F1F5F9;
    }
    ::-webkit-scrollbar-thumb {
        background: #CBD5E1;
        border-radius: 4px;
    }
    ::-webkit-scrollbar-thumb:hover {
        background: #94A3B8;
    }
    </style>
    """
    st.markdown(custom_css, unsafe_allow_html=True)

def render_sidebar_profile():
    """
    Renders an executive author and academic affiliation card inside the sidebar.
    """
    st.sidebar.markdown("""
    <div style="background: rgba(255, 255, 255, 0.05); border: 1px solid rgba(255, 255, 255, 0.1); border-radius: 14px; padding: 14px; margin-bottom: 16px;">
        <div style="display: flex; align-items: center; gap: 10px;">
            <div style="width: 38px; height: 38px; border-radius: 10px; background: linear-gradient(135deg, #3B82F6 0%, #1D4ED8 100%); display: flex; align-items: center; justify-content: center; font-weight: 800; color: #FFF; font-size: 14px; box-shadow: 0 4px 10px rgba(59, 130, 246, 0.4);">
                NZ
            </div>
            <div>
                <div style="font-weight: 700; font-size: 13px; color: #FFFFFF; line-height: 1.2;">Noman Zaid</div>
                <div style="font-size: 11px; color: #94A3B8;">Roll: 2500320140103</div>
            </div>
        </div>
        <div style="margin-top: 10px; padding-top: 8px; border-top: 1px solid rgba(255, 255, 255, 0.08); font-size: 11px; color: #CBD5E1; display: flex; flex-direction: column; gap: 2px;">
            <div>🏫 <b>ABES Engineering College</b></div>
            <div>🎓 MCA 2nd Year (Data Analytics)</div>
            <div>👩‍🏫 Mentor: <span style="color: #93C5FD; font-weight: 600;">Ms. Shilpa</span></div>
        </div>
    </div>
    """, unsafe_allow_html=True)

def render_header(title: str, subtitle: str, badge_text: str = "PRODUCTION ENTERPRISE APP"):
    """
    Renders an executive, modern Tailwind CSS page header with responsive typography,
    glowing badge, and academic metadata.
    """
    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <script src="https://cdn.tailwindcss.com"></script>
        <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@500;700;800&display=swap" rel="stylesheet">
        <style>
            body {{ font-family: 'Plus Jakarta Sans', sans-serif; }}
            @keyframes pulse-slow {{
                0%, 100% {{ opacity: 1; transform: scale(1); }}
                50% {{ opacity: 0.7; transform: scale(1.15); }}
            }}
            .pulse-dot {{ animation: pulse-slow 2s infinite ease-in-out; }}
            .glow-effect {{
                box-shadow: 0 10px 30px -5px rgba(30, 58, 138, 0.4), 0 0 0 1px rgba(255, 255, 255, 0.1);
            }}
        </style>
    </head>
    <body class="bg-transparent m-0 p-0 antialiased">
        <div class="bg-gradient-to-r from-slate-950 via-slate-900 to-blue-950 p-6 rounded-2xl glow-effect text-white flex flex-col md:flex-row justify-between items-start md:items-center gap-4 relative overflow-hidden">
            <!-- Background accent glow -->
            <div class="absolute -right-16 -top-16 w-64 h-64 rounded-full bg-blue-500/10 blur-3xl pointer-events-none"></div>
            
            <div class="relative z-10">
                <div class="inline-flex items-center gap-2 px-3 py-1 rounded-full text-xs font-bold bg-blue-500/20 text-blue-300 border border-blue-400/30 mb-2.5">
                    <span class="w-2 h-2 rounded-full bg-emerald-400 pulse-dot"></span>
                    <span class="tracking-wide">{badge_text}</span>
                </div>
                <h1 class="text-2xl md:text-3xl font-extrabold tracking-tight text-white m-0 bg-gradient-to-r from-white via-slate-100 to-slate-300 bg-clip-text text-transparent">{title}</h1>
                <p class="text-slate-300 text-xs md:text-sm mt-1.5 mb-0 max-w-2xl font-normal leading-relaxed">{subtitle}</p>
            </div>
            
            <div class="relative z-10 text-left md:text-right text-xs text-slate-300 bg-white/5 backdrop-blur-md p-3.5 rounded-xl border border-white/10 shadow-inner flex flex-col gap-1 min-w-[210px]">
                <div class="flex items-center justify-between md:justify-end gap-2">
                    <span class="w-2 h-2 rounded-full bg-blue-400"></span>
                    <span class="font-bold text-slate-100">ABES Engineering College</span>
                </div>
                <div class="text-slate-400">MCA 2nd Yr · Data Analytics</div>
                <div class="text-blue-300 font-semibold pt-1 border-t border-white/10 flex items-center justify-between md:justify-end gap-1">
                    <span>Lead:</span>
                    <span>Noman Zaid (2500320140103)</span>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    components.html(html_code, height=135)

def render_ticker_ribbon():
    """
    Renders an eye-catching live InsurTech telemetry ticker ribbon with key enterprise metrics.
    """
    html_code = """
    <!DOCTYPE html>
    <html>
    <head>
        <script src="https://cdn.tailwindcss.com"></script>
        <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@600;700&display=swap" rel="stylesheet">
        <style>body { font-family: 'Plus Jakarta Sans', sans-serif; }</style>
    </head>
    <body class="bg-transparent m-0 p-0">
        <div class="grid grid-cols-2 md:grid-cols-4 gap-3 py-1">
            <div class="flex items-center gap-2.5 px-3.5 py-2 rounded-xl bg-white border border-slate-200/80 shadow-xs">
                <div class="w-7 h-7 rounded-lg bg-blue-50 text-blue-600 flex items-center justify-center font-bold text-xs">⚡</div>
                <div>
                    <div class="text-[10px] uppercase font-bold text-slate-400 tracking-wider">Engine Status</div>
                    <div class="text-xs font-extrabold text-slate-800">Operational (100%)</div>
                </div>
            </div>
            <div class="flex items-center gap-2.5 px-3.5 py-2 rounded-xl bg-white border border-slate-200/80 shadow-xs">
                <div class="w-7 h-7 rounded-lg bg-emerald-50 text-emerald-600 flex items-center justify-center font-bold text-xs">₹</div>
                <div>
                    <div class="text-[10px] uppercase font-bold text-slate-400 tracking-wider">Disbursement</div>
                    <div class="text-xs font-extrabold text-emerald-700">₹ 4.34 Cr Disbursed</div>
                </div>
            </div>
            <div class="flex items-center gap-2.5 px-3.5 py-2 rounded-xl bg-white border border-slate-200/80 shadow-xs">
                <div class="w-7 h-7 rounded-lg bg-amber-50 text-amber-600 flex items-center justify-center font-bold text-xs">⚖️</div>
                <div>
                    <div class="text-[10px] uppercase font-bold text-slate-400 tracking-wider">IRDAI SLA Compliance</div>
                    <div class="text-xs font-extrabold text-slate-800">9.0d Median TAT</div>
                </div>
            </div>
            <div class="flex items-center gap-2.5 px-3.5 py-2 rounded-xl bg-white border border-slate-200/80 shadow-xs">
                <div class="w-7 h-7 rounded-lg bg-purple-50 text-purple-600 flex items-center justify-center font-bold text-xs">🛡️</div>
                <div>
                    <div class="text-[10px] uppercase font-bold text-slate-400 tracking-wider">SIU Fraud Gate</div>
                    <div class="text-xs font-extrabold text-purple-700">Active Rule Triage</div>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    components.html(html_code, height=55)

def render_kpi_row(kpi_list):
    """
    Renders a responsive 4-column row of Tailwind CSS KPI metric cards with glassmorphism,
    hover micro-animations, bold value badges, and trend indicators.
    """
    cards_html = ""
    for k in kpi_list:
        color = k.get("color", "blue")
        if color == "emerald":
            accent_border = "border-t-4 border-t-emerald-500"
            border_cls = "border-slate-200"
            bg_cls = "bg-gradient-to-br from-white to-emerald-50/40"
            title_cls = "text-emerald-700"
            val_cls = "text-emerald-950"
            badge_cls = "bg-emerald-100 text-emerald-800 border border-emerald-200"
        elif color == "rose":
            accent_border = "border-t-4 border-t-rose-500"
            border_cls = "border-slate-200"
            bg_cls = "bg-gradient-to-br from-white to-rose-50/40"
            title_cls = "text-rose-700"
            val_cls = "text-rose-950"
            badge_cls = "bg-rose-100 text-rose-800 border border-rose-200"
        elif color == "amber":
            accent_border = "border-t-4 border-t-amber-500"
            border_cls = "border-slate-200"
            bg_cls = "bg-gradient-to-br from-white to-amber-50/40"
            title_cls = "text-amber-700"
            val_cls = "text-amber-950"
            badge_cls = "bg-amber-100 text-amber-800 border border-amber-200"
        else: # blue / default
            accent_border = "border-t-4 border-t-blue-600"
            border_cls = "border-slate-200"
            bg_cls = "bg-gradient-to-br from-white to-blue-50/40"
            title_cls = "text-blue-700"
            val_cls = "text-slate-900"
            badge_cls = "bg-blue-100 text-blue-800 border border-blue-200"

        trend_html = f'<span class="text-[11px] px-2.5 py-0.5 rounded-full font-bold {badge_cls}">{k.get("trend", "")}</span>' if k.get("trend") else ""

        cards_html += f"""
        <div class="flex-1 min-w-[210px] p-5 rounded-2xl border {border_cls} {accent_border} {bg_cls} shadow-sm transition-all duration-300 hover:shadow-lg hover:-translate-y-1">
            <div class="flex justify-between items-center mb-1.5">
                <span class="text-[11px] font-bold uppercase tracking-wider {title_cls}">{k['title']}</span>
                {trend_html}
            </div>
            <div class="text-3xl font-black {val_cls} tracking-tight my-1">{k['value']}</div>
            <div class="text-xs text-slate-500 font-medium mt-1">{k.get('subtext', '')}</div>
        </div>
        """

    full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <script src="https://cdn.tailwindcss.com"></script>
        <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@600;700;800;900&display=swap" rel="stylesheet">
        <style>body {{ font-family: 'Plus Jakarta Sans', sans-serif; }}</style>
    </head>
    <body class="bg-transparent m-0 p-0 antialiased">
        <div class="flex flex-wrap gap-4 w-full">
            {cards_html}
        </div>
    </body>
    </html>
    """
    components.html(full_html, height=130)

def render_insight_box(title: str, text: str, category: str = "Operational Insight"):
    """
    Renders an executive analytical callout / insight box with modern gradient borders.
    """
    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <script src="https://cdn.tailwindcss.com"></script>
        <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@500;700;800&display=swap" rel="stylesheet">
        <style>body {{ font-family: 'Plus Jakarta Sans', sans-serif; }}</style>
    </head>
    <body class="bg-transparent m-0 p-0 antialiased">
        <div class="p-4 rounded-xl border border-blue-200/80 bg-gradient-to-r from-blue-50/90 via-indigo-50/50 to-white shadow-xs text-slate-800 flex items-start gap-3">
            <div class="w-8 h-8 rounded-lg bg-blue-600 text-white flex items-center justify-center font-bold text-sm shrink-0 shadow-sm">
                💡
            </div>
            <div class="flex-1">
                <div class="flex items-center gap-2 mb-1">
                    <span class="text-[10px] font-extrabold uppercase tracking-wider text-blue-700 bg-blue-100/80 px-2 py-0.5 rounded border border-blue-200">{category}</span>
                    <span class="text-sm font-extrabold text-slate-900">{title}</span>
                </div>
                <p class="text-xs md:text-sm text-slate-700 m-0 leading-relaxed font-normal">{text}</p>
            </div>
        </div>
    </body>
    </html>
    """
    components.html(html_code, height=98)
