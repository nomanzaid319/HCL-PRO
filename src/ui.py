"""
Insurance Claim Processing Analysis - UI Framework & Tailwind Component Library
Author: Noman Zaid (Roll No: 2500320140103, MCA 2nd Year Data Analytics, ABES)
Tech: Streamlit Components, Tailwind CSS CDN via iframe
"""

import streamlit as st
import streamlit.components.v1 as components

def render_header(title: str, subtitle: str, badge_text: str = "PRODUCTION ENTERPRISE APP"):
    """
    Renders an executive, modern Tailwind CSS page header with responsive typography.
    """
    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-transparent m-0 p-0 font-sans antialiased">
        <div class="bg-gradient-to-r from-slate-900 via-blue-950 to-slate-900 p-6 rounded-2xl shadow-md border border-slate-800 text-white flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
            <div>
                <div class="inline-flex items-center gap-2 px-3 py-1 rounded-full text-xs font-semibold bg-blue-500/20 text-blue-300 border border-blue-400/30 mb-2">
                    <span class="w-2 h-2 rounded-full bg-blue-400 animate-pulse"></span>
                    {badge_text}
                </div>
                <h1 class="text-2xl md:text-3xl font-extrabold tracking-tight text-white m-0">{title}</h1>
                <p class="text-slate-300 text-xs md:text-sm mt-1 mb-0">{subtitle}</p>
            </div>
            <div class="text-left md:text-right text-xs text-slate-400 bg-white/5 p-3 rounded-xl border border-white/10">
                <p class="font-bold text-slate-200 m-0">ABES Engineering College</p>
                <p class="m-0 text-slate-300">MCA 2nd Yr · Data Analytics Project</p>
                <p class="m-0 text-blue-300 font-medium">Noman Zaid (2500320140103)</p>
            </div>
        </div>
    </body>
    </html>
    """
    components.html(html_code, height=130)

def render_kpi_row(kpi_list):
    """
    Renders a responsive 4-column row of Tailwind CSS KPI metric cards.
    Each item in kpi_list is a dict:
    {'title': str, 'value': str, 'subtext': str, 'trend': str, 'color': 'blue'|'emerald'|'amber'|'rose'}
    """
    cards_html = ""
    for k in kpi_list:
        color = k.get("color", "blue")
        if color == "emerald":
            border_cls = "border-emerald-200"
            bg_cls = "bg-gradient-to-br from-white to-emerald-50/50"
            title_cls = "text-emerald-700"
            val_cls = "text-emerald-900"
            badge_cls = "bg-emerald-100 text-emerald-800"
        elif color == "rose":
            border_cls = "border-rose-200"
            bg_cls = "bg-gradient-to-br from-white to-rose-50/50"
            title_cls = "text-rose-700"
            val_cls = "text-rose-900"
            badge_cls = "bg-rose-100 text-rose-800"
        elif color == "amber":
            border_cls = "border-amber-200"
            bg_cls = "bg-gradient-to-br from-white to-amber-50/50"
            title_cls = "text-amber-700"
            val_cls = "text-amber-900"
            badge_cls = "bg-amber-100 text-amber-800"
        else: # blue / default
            border_cls = "border-blue-200"
            bg_cls = "bg-gradient-to-br from-white to-blue-50/50"
            title_cls = "text-blue-700"
            val_cls = "text-slate-900"
            badge_cls = "bg-blue-100 text-blue-800"

        trend_html = f'<span class="text-xs px-2 py-0.5 rounded font-semibold {badge_cls}">{k.get("trend", "")}</span>' if k.get("trend") else ""

        cards_html += f"""
        <div class="flex-1 min-w-[200px] p-4 rounded-xl border {border_cls} {bg_cls} shadow-sm transition hover:shadow-md">
            <div class="flex justify-between items-center mb-1">
                <span class="text-xs font-semibold uppercase tracking-wider {title_cls}">{k['title']}</span>
                {trend_html}
            </div>
            <div class="text-2xl font-black {val_cls} my-1">{k['value']}</div>
            <div class="text-xs text-slate-500 font-medium">{k.get('subtext', '')}</div>
        </div>
        """

    full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-transparent m-0 p-0 font-sans antialiased">
        <div class="flex flex-wrap gap-4 w-full">
            {cards_html}
        </div>
    </body>
    </html>
    """
    components.html(full_html, height=120)

def render_insight_box(title: str, text: str, category: str = "Operational Insight"):
    """
    Renders an executive analytical callout / insight box.
    """
    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <script src="https://cdn.tailwindcss.com"></script>
    </head>
    <body class="bg-transparent m-0 p-0 font-sans antialiased">
        <div class="p-4 rounded-xl border-l-4 border-blue-600 bg-blue-50/80 shadow-sm text-slate-800">
            <div class="flex items-center gap-2 mb-1">
                <span class="text-xs font-bold uppercase tracking-wider text-blue-700 bg-blue-100 px-2 py-0.5 rounded">{category}</span>
                <span class="text-sm font-bold text-slate-900">{title}</span>
            </div>
            <p class="text-xs md:text-sm text-slate-700 m-0 leading-relaxed">{text}</p>
        </div>
    </body>
    </html>
    """
    components.html(html_code, height=95)
