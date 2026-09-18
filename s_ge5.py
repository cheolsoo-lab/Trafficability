# -*- coding: utf-8 -*-
"""
장비주행성 검토 프로그램 (Equipment Trafficability Review)
============================================================
- 하중분산 응력법: 복토 두께에 따른 원지반 작용응력(σ) 산정
- Meyerhof and Hanna (1978): 층상지반 허용지지력(qa) 산정
- 복토 두께 0.0m ~ 2.0m (0.1m 간격) 개별 체크박스 선택 방식 적용
- 인쇄용 양식 기반 PDF, Word, Excel, 바로인쇄 원클릭 다운로드 기능 탑재
"""

import math
import io
import os
import base64
import tempfile
from pathlib import Path
import pandas as pd
import streamlit as st
import plotly.graph_objects as go

# 실제 파일 생성용 라이브러리
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
                                PageBreak, Image as RLImage, KeepTogether)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Mm, Pt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.drawing.image import Image as XLImage

# theory_images 모듈 예외 처리
try:
    from theory_images import THEORY_IMG_A, THEORY_IMG_B
except ImportError:
    DUMMY_IMG_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORTH5CYII="
    THEORY_IMG_A = DUMMY_IMG_BASE64
    THEORY_IMG_B = DUMMY_IMG_BASE64

st.set_page_config(page_title="장비주행성 검토 프로그램", layout="wide")

# ------------------------------------------------------------------
# 1. 계산 함수
# ------------------------------------------------------------------

def nq2_func(phi2):
    return math.exp(math.pi * math.tan(math.radians(phi2))) * (math.tan(math.radians(45 + phi2 / 2))) ** 2

def nc2_func(phi2):
    if phi2 == 0:
        return 5.14
    return (nq2_func(phi2) - 1) / math.tan(math.radians(phi2))

KS_TABLE_X = [20, 25, 30, 35, 40, 45, 50]
KS_TABLE_Y = [1.89, 2.22, 3.06, 4.45, 6.95, 11.12, 19.15]

def ks_lookup(phi1):
    """펀칭전단계수 Ks : 복토층 내부마찰각(phi1) 기준 선형보간"""
    if phi1 <= 0:
        return 0.0
    xs, ys = KS_TABLE_X, KS_TABLE_Y
    if phi1 <= xs[0]:
        return ys[0]
    if phi1 >= xs[-1]:
        return ys[-1]
    for i in range(len(xs) - 1):
        if xs[i] <= phi1 <= xs[i + 1]:
            t = (phi1 - xs[i]) / (xs[i + 1] - xs[i])
            return ys[i] + t * (ys[i + 1] - ys[i])
    return ys[-1]

def sigma_dispersion(P, b, L, H, gamma1, theta_deg, impact=0.0):
    """하중분산 응력법에 따른 원지반 작용응력 σ (kPa)"""
    denom = (b + 2 * H * math.tan(math.radians(theta_deg))) * (L + 2 * H * math.tan(math.radians(theta_deg)))
    if denom == 0:
        return 0.0
    return (P * b * L * (1 + impact)) / denom + gamma1 * H

def qa_meyerhof_hanna(b, L, H, c2, phi1, phi2, gamma1, Fs, T_allow, theta_deg, Df=0.0):
    """Meyerhof and Hanna (1978) 모델에 따른 허용지지력 qa (kPa)"""
    Ks = ks_lookup(phi1)
    Nc2 = nc2_func(phi2)
    Fcs2 = 1 + 0.2 * (b / L)
    Fcd2 = 1 + 0.4 * (Df / b)
    
    term1 = c2 * Nc2 * Fcs2 * Fcd2
    term2 = gamma1 * (H ** 2) * (1 + (b / L)) * (Ks * math.tan(math.radians(phi1)) / b)
    term3 = (2 * T_allow * math.sin(math.radians(theta_deg))) / (b + H)
    
    qu = term1 + term2 + term3
    return qu / Fs

def compute_contact_pressure(W, b, L, is_dump=False):
    """접지압 P (kPa) 자동 산정식"""
    if b <= 0 or L <= 0:
        return 0.0
    if is_dump:
        return round(W * 0.4 / (b * L), 2)
    return round(W / (2 * b * L), 2)

# ------------------------------------------------------------------
# 2. 기본 장비 데이터베이스
# ------------------------------------------------------------------

DEFAULT_EQUIPMENT = [
    {"검토포함": True, "장비명": "도저", "규격": "6t", "중량W(kN)": 64.1, "폭b(m)": 0.71, "길이L(m)": 2.03, "덤프특수식": False},
    {"검토포함": True, "장비명": "도저", "규격": "13t", "중량W(kN)": 131.0, "폭b(m)": 0.77, "길이L(m)": 2.61, "덤프특수식": False},
    {"검토포함": True, "장비명": "도저", "규격": "19t", "중량W(kN)": 190.0, "폭b(m)": 0.76, "길이L(m)": 2.80, "덤프특수식": False},
    {"검토포함": True, "장비명": "도저", "규격": "24t", "중량W(kN)": 240.0, "폭b(m)": 0.55, "길이L(m)": 2.84, "덤프특수식": False},
    {"검토포함": True, "장비명": "도저", "규격": "30t", "중량W(kN)": 348.9, "폭b(m)": 0.63, "길이L(m)": 3.21, "덤프특수식": False},
    {"검토포함": True, "장비명": "덤프트럭", "규격": "15t", "중량W(kN)": 260.0, "폭b(m)": 0.20, "길이L(m)": 0.50, "덤프특수식": True},
    {"검토포함": True, "장비명": "덤프트럭", "규격": "24t", "중량W(kN)": 432.0, "폭b(m)": 0.23, "길이L(m)": 0.58, "덤프특수식": True},
    {"검토포함": False, "장비명": "백호(Back hoe)", "규격": "25t", "중량W(kN)": 268.0, "폭b(m)": 0.60, "길이L(m)": 3.61, "덤프특수식": False},
    {"검토포함": False, "장비명": "PBD", "규격": "50t", "중량W(kN)": 855.9, "폭b(m)": 1.20, "길이L(m)": 7.67, "덤프특수식": False},
    {"검토포함": False, "장비명": "PBD", "규격": "106t", "중량W(kN)": 1056.0, "폭b(m)": 1.40, "길이L(m)": 9.09, "덤프특수식": False},
    {"검토포함": False, "장비명": "GCP", "규격": "-", "중량W(kN)": 879.0, "폭b(m)": 0.76, "길이L(m)": 5.11, "덤프특수식": False},
    {"검토포함": False, "장비명": "DCM", "규격": "120t", "중량W(kN)": 1177.2, "폭b(m)": 0.86, "길이L(m)": 5.23, "덤프특수식": False},
]

# ------------------------------------------------------------------
# 3. 사이드바 설정
# ------------------------------------------------------------------

st.sidebar.header("📋 계산서 표지 정보")
proj_name = st.sidebar.text_input("PROJECT (프로젝트명)", "")

st.sidebar.divider()
st.sidebar.header("🌍 지반 및 설계 매개변수 입력")
gamma1 = st.sidebar.number_input("복토층 단위중량 γ1 (kN/m³)", value=18.0, step=0.5)
c2 = st.sidebar.number_input("원지반 점착력 c2 (kPa)", value=7.0, step=0.5)
phi1 = st.sidebar.number_input("복토층 내부마찰각 φ1 (°)", value=25.0, step=1.0)
phi2 = st.sidebar.number_input("원지반 내부마찰각 φ2 (°)", value=30.0, step=1.0)
Fs = st.sidebar.number_input("설계 안전율 Fs", value=1.5, step=0.1)
theta_geo = st.sidebar.number_input("토목섬유/하중분산 각도 θ (°)", value=30.0, step=1.0)
T_allow = st.sidebar.number_input("토목섬유 허용인장력 T (kN/m)", value=0.0, step=1.0)
impact = st.sidebar.number_input("충격계수 ε", value=0.0, step=0.05)

st.sidebar.divider()
st.sidebar.header("📏 복토 두께 검토 조건 (0.0m ~ 2.0m)")

thickness_list = [round(i * 0.1, 1) for i in range(21)]
default_selected_set = {0.1, 0.2, 0.3, 0.4, 0.5, 0.6}

for h in thickness_list:
    key = f"chk_h_{h}"
    if key not in st.session_state:
        st.session_state[key] = (h in default_selected_set)

select_all = st.sidebar.checkbox("전체 선택 / 해제", value=False, key="chk_select_all")
if st.sidebar.button("선택 상태 적용"):
    for h in thickness_list:
        st.session_state[f"chk_h_{h}"] = select_all
    st.rerun()

selected_thicknesses = []
with st.sidebar.expander("복토두께 항목 체크 (클릭하여 열기)", expanded=True):
    col1, col2 = st.columns(2)
    for idx, h in enumerate(thickness_list):
        target_col = col1 if idx % 2 == 0 else col2
        key = f"chk_h_{h}"
        checked = target_col.checkbox(f"{h:.1f} m", value=st.session_state[key], key=key)
        if checked:
            selected_thicknesses.append(h)

if not selected_thicknesses:
    st.sidebar.warning("최소 1개 이상의 복토두께를 선택해야 합니다. (기본값 0.1m 적용)")
    selected_thicknesses = [0.1]

selected_thicknesses.sort()

# ------------------------------------------------------------------
# 4. 메인 - 이론 및 산정식
# ------------------------------------------------------------------

st.title("🚜 연약지반 장비주행성 검토 프로그램")
st.caption("하중분산 응력법 (작용응력) & Meyerhof and Hanna 층상지반 지지력 모델 (허용지지력)")

with st.expander("📖 적용 이론 및 산정식 상세 보기", expanded=False):
    st.markdown("#### 1. 장비 접지압 (P) 자동 산정식")
    st.latex(r"P = \frac{W}{2 \cdot b \cdot L} \quad \left(\text{덤프트럭: } P = \frac{0.4 \cdot W}{b \cdot L}\right)")
    
    st.markdown("#### 2. 원지반상 작용응력 ($\sigma$) — 하중분산 응력법")
    st.latex(r"\sigma = \frac{P \cdot b \cdot L \cdot (1+\varepsilon)}{(b+2H\tan\theta)(L+2H\tan\theta)} + \gamma_1 H")
    st.image(f"data:image/png;base64,{THEORY_IMG_B}", caption="하중분산 응력 모델 개념도", use_container_width=True)

    st.markdown("#### 3. 허용지지력 ($q_a$) — Meyerhof and Hanna(1978) 층상지반 모델")
    st.latex(r"q_a = \frac{1}{F_s} \left[ \left(1+0.2\frac{b}{L}\right)c_2 N_{c(2)} F_{cs(2)} F_{cd(2)} + \gamma_1 H^2 \left(1+\frac{b}{L}\right)\frac{K_s\tan\phi_1}{b} + \frac{2T\sin\theta}{b+H} \right]")
    st.image(f"data:image/png;base64,{THEORY_IMG_A}", caption="Meyerhof-Hanna 층상지반 파괴 메커니즘", use_container_width=True)

st.divider()

# ------------------------------------------------------------------
# 5. 장비 데이터 입력 / 편집
# ------------------------------------------------------------------

st.subheader("🚛 검토 장비 입력 및 제원 설정")
st.caption("기본 제공 장비 외에 새로운 장비를 하단 `+` 버튼을 눌러 자유롭게 추가할 수 있습니다.")

if "equip_df" not in st.session_state:
    st.session_state.equip_df = pd.DataFrame(DEFAULT_EQUIPMENT)

edited_df = st.data_editor(
    st.session_state.equip_df,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "검토포함": st.column_config.CheckboxColumn("검토포함", default=True),
        "장비명": st.column_config.TextColumn("장비명", required=True),
        "규격": st.column_config.TextColumn("규격"),
        "중량W(kN)": st.column_config.NumberColumn("중량 W (kN)", min_value=0.0, step=1.0),
        "폭b(m)": st.column_config.NumberColumn("폭 b (m)", min_value=0.01, step=0.05),
        "길이L(m)": st.column_config.NumberColumn("길이 L (m)", min_value=0.01, step=0.05),
        "덤프특수식": st.column_config.CheckboxColumn("덤프 특수식 (P=0.4W/bL)", default=False),
    },
    key="equip_editor"
)

st.session_state.equip_df = edited_df

# ------------------------------------------------------------------
# 6. 검토 실행 및 연산
# ------------------------------------------------------------------

if st.button("▶ 검토 실행", type="primary") or "last_summary" in st.session_state:
    selected_target = edited_df[edited_df["검토포함"] == True].copy()
    
    summary_results = []
    matrix_rows = []

    for _, row in selected_target.iterrows():
        try:
            name = str(row["장비명"])
            spec = str(row["규격"])
            W = float(row["중량W(kN)"])
            b = float(row["폭b(m)"])
            L = float(row["길이L(m)"])
            is_dump = bool(row["덤프특수식"])
        except (ValueError, TypeError):
            continue

        P = compute_contact_pressure(W, b, L, is_dump)
        
        target_H = None
        final_sigma = 0.0
        final_qa = 0.0
        status = "N.G"

        mat_row = {
            "장비명": name, "규격": spec, "중량W(kN)": W,
            "폭b(m)": b, "길이L(m)": L, "접지압P(kPa)": P
        }

        for H in selected_thicknesses:
            sig = sigma_dispersion(P, b, L, H, gamma1, theta_geo, impact)
            qa = qa_meyerhof_hanna(b, L, H, c2, phi1, phi2, gamma1, Fs, T_allow, theta_geo)
            
            mat_row[f"σ({H:.1f}m)"] = round(sig, 2)
            mat_row[f"qa({H:.1f}m)"] = round(qa, 2)

            if qa >= sig and target_H is None:
                target_H = H
                final_sigma = sig
                final_qa = qa
                status = "O.K"

        if target_H is None:
            target_H = selected_thicknesses[-1]
            final_sigma = sigma_dispersion(P, b, L, target_H, gamma1, theta_geo, impact)
            final_qa = qa_meyerhof_hanna(b, L, target_H, c2, phi1, phi2, gamma1, Fs, T_allow, theta_geo)
            status = "N.G"

        safety_factor = round(final_qa / final_sigma, 2) if final_sigma > 0 else 0.0

        summary_results.append({
            "장비명": name, "규격": spec, "중량W(kN)": W,
            "폭b(m)": b, "길이L(m)": L, "접지압P(kPa)": P,
            "복토두께H(m)": target_H,
            "작용응력σ(kPa)": round(final_sigma, 2),
            "허용지지력qa(kPa)": round(final_qa, 2),
            "안전율(qa/σ)": safety_factor,
            "판정": status
        })
        matrix_rows.append(mat_row)

    summary_df = pd.DataFrame(summary_results)
    matrix_df = pd.DataFrame(matrix_rows)
    
    st.session_state.last_summary = summary_df
    st.session_state.last_matrix = matrix_df

    st.subheader("📊 검토 결과 요약 (최소 요구 복토두께 판정)")
    
    def highlight_status(val):
        if val == "O.K":
            return "background-color: #d4edda; color: #155724; font-weight: bold;"
        return "background-color: #f8d7da; color: #721c24; font-weight: bold;"

    st.dataframe(
        summary_df.style.map(highlight_status, subset=["판정"]),
        use_container_width=True, hide_index=True
    )

    with st.expander("🔍 선택 복토두께별 전체 작용응력 및 허용지지력 매트릭스 보기"):
        st.dataframe(matrix_df, use_container_width=True, hide_index=True)

    fig = go.Figure()
    chart_colors = ["#2e7d32" if s == "O.K" else "#c62828" for s in summary_df["판정"]]
    x_labels = [f"{n}({s})" for n, s in zip(summary_df["장비명"], summary_df["규격"])]
    
    fig.add_trace(go.Bar(x=x_labels, y=summary_df["작용응력σ(kPa)"], name="작용응력 σ", marker_color="#90a4ae"))
    fig.add_trace(go.Scatter(x=x_labels, y=summary_df["허용지지력qa(kPa)"], name="허용지지력 qa",
                              mode="lines+markers", marker=dict(color=chart_colors, size=10)))
    fig.update_layout(title="장비별 작용응력(σ) vs 허용지지력(qa)", yaxis_title="kPa", legend=dict(x=1.02))
    
    st.session_state.last_fig = fig
    st.plotly_chart(fig, use_container_width=True, key="main_result_chart")

# ------------------------------------------------------------------
# 7. 인쇄용 구조계산서 - 파일 생성 유틸리티 및 렌더링
# ------------------------------------------------------------------

st.divider()
st.subheader("🖨️ 인쇄용 구조계산서 보기 및 출력/다운로드")
st.caption("화면에 표시되는 구조계산서와 동일한 구성으로 A4 페이지에 맞춰 PDF·Word·Excel을 생성합니다.")

def _img_from_b64(b64, suffix='.png'):
    """Base64 데이터를 디스크 상의 임시 파일로 안전하게 기록합니다."""
    img_bytes = base64.b64decode(b64)
    f = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    f.write(img_bytes)
    f.flush()
    os.fsync(f.fileno())
    f.close()
    return f.name

def _make_chart_png(summary_df):
    """Plotly 그래프를 PNG 바이너리 데이터로 변환합니다."""
    try:
        fig = go.Figure()
        x = [f"{n}({sp})" for n, sp in zip(summary_df['장비명'], summary_df['규격'])]
        fig.add_trace(go.Bar(x=x, y=summary_df['작용응력σ(kPa)'], name='작용응력 σ'))
        fig.add_trace(go.Scatter(x=x, y=summary_df['허용지지력qa(kPa)'], name='허용지지력 qa', mode='lines+markers'))
        fig.update_layout(title='장비별 작용응력(σ) vs 허용지지력(qa)', xaxis_title='장비', yaxis_title='kPa',
                          autosize=False, width=1000, height=500, margin=dict(l=70, r=30, t=70, b=120))
        return fig.to_image(format='png', width=1000, height=500, scale=2)
    except Exception:
        import matplotlib.pyplot as plt
        labels = [f"{n}({sp})" for n, sp in zip(summary_df['장비명'], summary_df['규격'])]
        fig, ax = plt.subplots(figsize=(10, 5.0))
        x_range = range(len(labels))
        ax.bar(list(x_range), summary_df['작용응력σ(kPa)'], label='작용응력 σ')
        ax.plot(list(x_range), summary_df['허용지지력qa(kPa)'], marker='o', label='허용지지력 qa')
        ax.set_xticks(list(x_range))
        ax.set_xticklabels(labels, rotation=45, ha='right', fontsize=8)
        ax.set_ylabel('kPa')
        ax.set_title('장비별 작용응력(σ) vs 허용지지력(qa)')
        ax.grid(axis='y', alpha=.25)
        ax.legend()
        fig.tight_layout()
        bio = io.BytesIO()
        fig.savefig(bio, format='png', dpi=180)
        plt.close(fig)
        return bio.getvalue()

def _make_formula_images():
    """Matplotlib을 사용해 수식 이미지를 바이트 데이터로 생성합니다."""
    import matplotlib.pyplot as plt
    formulas = {
      'P': r'$P=\frac{W}{2\,b\,L}$   (덤프트럭: $P=\frac{0.4W}{bL}$)',
      'sigma': r'$\sigma=\frac{P\,b\,L\,(1+\varepsilon)}{(b+2H\tan\theta)(L+2H\tan\theta)}+\gamma_1H$',
      'qa': r'$q_a=\frac{1}{F_s}\left[\left(1+0.2\frac{b}{L}\right)c_2N_{c(2)}F_{cs(2)}F_{cd(2)}+\gamma_1H^2\left(1+\frac{b}{L}\right)\frac{K_s\tan\phi_1}{b}+\frac{2T\sin\theta}{b+H}\right]$'
    }
    out = {}
    for k, formula in formulas.items():
        fig = plt.figure(figsize=(11, 0.85))
        fig.patch.set_alpha(0)
        fig.text(.02, .48, formula, fontsize=17, va='center')
        plt.axis('off')
        bio = io.BytesIO()
        fig.savefig(bio, format='png', dpi=220, bbox_inches='tight', pad_inches=.08, transparent=True)
        plt.close(fig)
        out[k] = bio.getvalue()
    return out

def _safe_pdf_font():
    candidates = [
      r'C:\Windows\Fonts\malgun.ttf', r'C:\Windows\Fonts\malgunbd.ttf',
      '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
      '/usr/share/fonts/truetype/nanum/NanumGothic.ttf'
    ]
    for fp in candidates:
        if os.path.exists(fp):
            try:
                pdfmetrics.registerFont(TTFont('Korean', fp))
                return 'Korean'
            except Exception:
                pass
    return 'Helvetica'

def _pdf_table(data, widths, header=True, font='Helvetica', fontsize=7.2):
    t = Table(data, colWidths=widths, repeatRows=1 if header else 0, hAlign='CENTER')
    cmds = [
        ('GRID', (0, 0), (-1, -1), 0.35, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), font),
        ('FONTSIZE', (0, 0), (-1, -1), fontsize),
        ('LEFTPADDING', (0, 0), (-1, -1), 2),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3)
    ]
    if header:
        cmds += [('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#e9eef4')), ('FONTNAME', (0, 0), (-1, 0), font)]
    t.setStyle(TableStyle(cmds))
    return t

def _add_docx_picture_safely(doc, img_src, width=None):
    """
    python-docx에 이미지를 안전하게 추가하는 함수.
    경로(str) 또는 io.BytesIO 스트림 형태를 모두 지원하며, 
    유효성을 검증하여 UnexpectedEndOfFileError를 예방합니다.
    """
    try:
        if isinstance(img_src, str):
            if os.path.exists(img_src) and os.path.getsize(img_src) > 0:
                doc.add_picture(img_src, width=width)
            else:
                print(f"[Warning] 유효하지 않은 파일 경로이거나 0바이트 파일입니다: {img_src}")
        elif isinstance(img_src, io.BytesIO):
            img_src.seek(0)
            doc.add_picture(img_src, width=width)
        elif isinstance(img_src, bytes):
            stream = io.BytesIO(img_src)
            stream.seek(0)
            doc.add_picture(stream, width=width)
    except Exception as e:
        print(f"[Error] DOCX 이미지 삽입 실패: {e}")

def build_pdf(summary_df, matrix_df, selected_thicknesses, params, proj_name, theory_paths, formula_paths, chart_path):
    font = _safe_pdf_font()
    path = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf').name
    doc = SimpleDocTemplate(path, pagesize=A4, rightMargin=12*mm, leftMargin=12*mm, topMargin=12*mm, bottomMargin=12*mm,
                            title='장비주행성 검토 구조계산서')
    styles = getSampleStyleSheet()
    body = ParagraphStyle('kr', parent=styles['BodyText'], fontName=font, fontSize=8.5, leading=12, spaceAfter=4)
    title = ParagraphStyle('title', parent=body, fontSize=16, leading=20, alignment=TA_CENTER, spaceAfter=8)
    sec = ParagraphStyle('sec', parent=body, fontSize=11, fontName=font, leading=14, spaceBefore=5, spaceAfter=5)
    story = []
    
    story.append(_pdf_table([[Paragraph('<b>PROJECT</b>', body), Paragraph(proj_name if proj_name.strip() else '&nbsp;', body)]], [32*mm, 142*mm], font=font, fontsize=9))
    story.append(Spacer(1, 5*mm))
    story.append(Paragraph('장비주행성 검토 구조계산서', title))
    story.append(Paragraph('1. 적용 이론 및 산정식', sec))
    
    labels = [
        ('가. 장비 접지압(P) 자동 산정식', 'P'),
        ('나. 원지반상 작용응력(σ) — 하중분산 응력법', 'sigma'),
        ('다. 허용지지력(qa) — Meyerhof and Hanna(1978)', 'qa')
    ]
    for lab, key in labels:
        img = RLImage(formula_paths[key], width=172*mm, height=13*mm, kind='bound')
        story.append(KeepTogether([Paragraph(lab, body), img, Spacer(1, 2*mm)]))
        
    story.append(Paragraph('<b>[설계 적용 매개변수]</b>', body))
    pheaders = ['γ1(kN/m³)', 'c2(kPa)', 'φ1(°)', 'φ2(°)', 'Ks', 'θ(°)', 'T(kN/m)', 'ε', 'Fs']
    pvals = [
        f"{params['gamma1']:.1f}", f"{params['c2']:.1f}", f"{params['phi1']:.1f}",
        f"{params['phi2']:.1f}", f"{params['Ks']:.2f}", f"{params['theta']:.1f}",
        f"{params['T']:.1f}", f"{params['impact']:.2f}", f"{params['Fs']:.1f}"
    ]
    story.append(_pdf_table([pheaders, pvals], [18.5*mm]*9, font=font, fontsize=6.3))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph('수식 주요 변수: W 장비 총중량, b 접지폭, L 접지길이, H 복토두께, P 접지압, σ 원지반 작용응력, qa 허용지지력, γ1 복토층 단위중량, c2 원지반 점착력, φ1·φ2 내부마찰각, Ks 펀칭전단계수, θ 분산/보강각도, T 토목섬유 허용인장력, Fs 안전율.', body))
    story.append(Spacer(1, 2*mm))
    story.append(RLImage(theory_paths['A'], width=178*mm, height=62*mm, kind='bound'))
    story.append(Spacer(1, 2*mm))
    story.append(RLImage(theory_paths['B'], width=178*mm, height=62*mm, kind='bound'))
    story.append(Spacer(1, 3*mm))

    story.append(PageBreak())
    story.append(Paragraph('2. 복토 두께별 작용응력', sec))
    chunk_size = 8
    for start in range(0, len(selected_thicknesses), chunk_size):
        hs = selected_thicknesses[start:start+chunk_size]
        hdr = ['장비명', '규격', 'P(kPa)'] + [f'σ({h:.1f}m)' for h in hs]
        data = [hdr]
        for _, r in matrix_df.iterrows():
            data.append([r['장비명'], r['규격'], f"{r['접지압P(kPa)']:.2f}"] + [f"{r.get(f'σ({h:.1f}m)','-'):.2f}" if isinstance(r.get(f'σ({h:.1f}m)'), (int, float)) else '-' for h in hs])
        widths = [26*mm, 18*mm, 17*mm] + [15*mm]*len(hs)
        story.append(_pdf_table(data, widths, font=font, fontsize=6.6))
        story.append(Spacer(1, 4*mm))
        if start + chunk_size < len(selected_thicknesses):
            story.append(PageBreak())

    story.append(PageBreak())
    story.append(Paragraph('3. 장비주행성 검토결과', sec))
    hdr = ['장비명', '규격', 'P\n(kPa)', 'b\n(m)', 'L\n(m)', 'H\n(m)', 'σ\n(kPa)', 'qa\n(kPa)', 'qa/σ', '판정']
    for start in range(0, len(summary_df), 15):
        part = summary_df.iloc[start:start+15]
        data = [hdr]
        for _, r in part.iterrows():
            data.append([r['장비명'], r['규격'], f"{r['접지압P(kPa)']:.2f}", f"{r['폭b(m)']:.2f}", f"{r['길이L(m)']:.2f}", f"{r['복토두께H(m)']:.1f}", f"{r['작용응력σ(kPa)']:.2f}", f"{r['허용지지력qa(kPa)']:.2f}", f"{r['안전율(qa/σ)']:.2f}", r['판정']])
        t = _pdf_table(data, [25*mm, 18*mm, 17*mm, 13*mm, 13*mm, 13*mm, 18*mm, 18*mm, 15*mm, 15*mm], font=font, fontsize=6.7)
        t.setStyle(TableStyle([('TEXTCOLOR', (-1, 1), (-1, -1), colors.black)]))
        story.append(t)
        if start + 15 < len(summary_df):
            story.append(PageBreak())
            
    story.append(PageBreak())
    story.append(Paragraph('4. 장비별 작용응력(σ) vs 허용지지력(qa) 비교 그래프', sec))
    story.append(RLImage(chart_path, width=178*mm, height=89*mm, kind='bound'))
    doc.build(story)
    return path

def build_docx(summary_df, matrix_df, selected_thicknesses, params, proj_name, theory_paths, formula_paths, chart_path):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Mm(210)
    sec.page_height = Mm(297)
    sec.top_margin = Mm(12)
    sec.bottom_margin = Mm(12)
    sec.left_margin = Mm(12)
    sec.right_margin = Mm(12)
    
    p = doc.add_table(rows=1, cols=2)
    p.alignment = WD_TABLE_ALIGNMENT.CENTER
    p.style = 'Table Grid'
    p.cell(0, 0).text = 'PROJECT'
    p.cell(0, 1).text = proj_name
    
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run('장비주행성 검토 구조계산서')
    r.bold = True
    r.font.size = Pt(16)
    
    doc.add_heading('1. 적용 이론 및 산정식', level=1)
    for lab, key in [('가. 장비 접지압(P) 자동 산정식', 'P'), ('나. 원지반상 작용응력(σ) — 하중분산 응력법', 'sigma'), ('다. 허용지지력(qa) — Meyerhof and Hanna(1978)', 'qa')]:
        doc.add_paragraph(lab).runs[0].bold = True
        _add_docx_picture_safely(doc, formula_paths[key], width=Mm(175))
        
    doc.add_paragraph('설계 적용 매개변수').runs[0].bold = True
    tb = doc.add_table(rows=2, cols=9)
    tb.style = 'Table Grid'
    headers = ['γ1', 'c2', 'φ1', 'φ2', 'Ks', 'θ', 'T', 'ε', 'Fs']
    vals = [params['gamma1'], params['c2'], params['phi1'], params['phi2'], params['Ks'], params['theta'], params['T'], params['impact'], params['Fs']]
    for j, h in enumerate(headers):
        tb.cell(0, j).text = str(h)
        tb.cell(1, j).text = f'{vals[j]:.2f}'
        
    doc.add_paragraph('수식 주요 변수: W 장비 총중량, b 접지폭, L 접지길이, H 복토두께, P 접지압, σ 원지반 작용응력, qa 허용지지력, γ1 복토층 단위중량, c2 원지반 점착력, φ1·φ2 내부마찰각, Ks 펀칭전단계수, θ 분산/보강각도, T 토목섬유 허용인장력, Fs 안전율.')
    
    # 안전한 이미지 삽입 함수를 사용하여 UnexpectedEndOfFileError 원인 차단
    _add_docx_picture_safely(doc, theory_paths['A'], width=Mm(175))
    _add_docx_picture_safely(doc, theory_paths['B'], width=Mm(175))
    doc.add_page_break()
    
    doc.add_heading('2. 복토 두께별 작용응력', level=1)
    for start in range(0, len(selected_thicknesses), 8):
        hs = selected_thicknesses[start:start+8]
        tb = doc.add_table(rows=1, cols=3+len(hs))
        tb.style = 'Table Grid'
        for j, h in enumerate(['장비명', '규격', 'P(kPa)'] + [f'σ({x:.1f}m)' for x in hs]):
            tb.cell(0, j).text = str(h)
        for _, r in matrix_df.iterrows():
            cells = tb.add_row().cells
            vals = [r['장비명'], r['규격'], f"{r['접지압P(kPa)']:.2f}"] + [f"{r.get(f'σ({h:.1f}m)','-'):.2f}" if isinstance(r.get(f'σ({h:.1f}m)'), (int, float)) else '-' for h in hs]
            for j, v in enumerate(vals):
                cells[j].text = str(v)
        if start + 8 < len(selected_thicknesses):
            doc.add_page_break()
            
    doc.add_page_break()
    doc.add_heading('3. 장비주행성 검토결과', level=1)
    for start in range(0, len(summary_df), 15):
        part = summary_df.iloc[start:start+15]
        tb = doc.add_table(rows=1, cols=10)
        tb.style = 'Table Grid'
        headers = ['장비명', '규격', 'P(kPa)', 'b(m)', 'L(m)', 'H(m)', 'σ(kPa)', 'qa(kPa)', 'qa/σ', '판정']
        for j, h in enumerate(headers):
            tb.cell(0, j).text = h
        for _, r in part.iterrows():
            vals = [r['장비명'], r['규격'], f"{r['접지압P(kPa)']:.2f}", f"{r['폭b(m)']:.2f}", f"{r['길이L(m)']:.2f}", f"{r['복토두께H(m)']:.1f}", f"{r['작용응력σ(kPa)']:.2f}", f"{r['허용지지력qa(kPa)']:.2f}", f"{r['안전율(qa/σ)']:.2f}", r['판정']]
            cells = tb.add_row().cells
            for j, v in enumerate(vals):
                cells[j].text = str(v)
        if start + 15 < len(summary_df):
            doc.add_page_break()
            
    doc.add_page_break()
    doc.add_heading('4. 장비별 작용응력(σ) vs 허용지지력(qa) 비교 그래프', level=1)
    _add_docx_picture_safely(doc, chart_path, width=Mm(175))
    
    path = tempfile.NamedTemporaryFile(delete=False, suffix='.docx').name
    doc.save(path)
    return path

def build_xlsx(summary_df, matrix_df, selected_thicknesses, params, proj_name, theory_paths, formula_paths, chart_path):
    wb = Workbook()
    ws = wb.active
    ws.title = '구조계산서'
    ws.sheet_view.showGridLines = False
    thin = Side(style='thin', color='000000')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill('solid', fgColor='E9EEF4')
    
    ws.merge_cells('A1:J1')
    ws['A1'] = '장비주행성 검토 구조계산서'
    ws['A1'].font = Font(bold=True, size=16)
    ws['A1'].alignment = Alignment(horizontal='center')
    
    ws['A3'] = 'PROJECT'
    ws['B3'] = proj_name
    ws['A3'].font = Font(bold=True)
    ws['A3'].border = ws['B3'].border = border
    
    ws['A5'] = '1. 적용 이론 및 산정식'
    ws['A5'].font = Font(bold=True, size=12)
    
    row = 7
    for lab, key in [('가. 장비 접지압(P) 자동 산정식', 'P'), ('나. 원지반상 작용응력(σ) — 하중분산 응력법', 'sigma'), ('다. 허용지지력(qa) — Meyerhof and Hanna(1978)', 'qa')]:
        ws.cell(row, 1, lab).font = Font(bold=True)
        im = XLImage(formula_paths[key])
        ratio = im.height / im.width
        im.width = 650
        im.height = int(650 * ratio)
        ws.add_image(im, f'A{row+1}')
        row += 6
        
    ws.cell(row, 1, '설계 적용 매개변수').font = Font(bold=True)
    row += 1
    headers = ['γ1', 'c2', 'φ1', 'φ2', 'Ks', 'θ', 'T', 'ε', 'Fs']
    vals = [params['gamma1'], params['c2'], params['phi1'], params['phi2'], params['Ks'], params['theta'], params['T'], params['impact'], params['Fs']]
    for j, h in enumerate(headers, 1):
        ws.cell(row, j, h)
        ws.cell(row, j).border = border
        ws.cell(row, j).fill = header_fill
        ws.cell(row+1, j, vals[j-1])
        ws.cell(row+1, j).border = border
        ws.cell(row+1, j).alignment = Alignment(horizontal='center')
        
    row += 4
    im = XLImage(theory_paths['A'])
    ratio = im.height / im.width
    im.width = 700
    im.height = int(700 * ratio)
    ws.add_image(im, f'A{row}')
    
    im2 = XLImage(theory_paths['B'])
    ratio2 = im2.height / im2.width
    im2.width = 700
    im2.height = int(700 * ratio2)
    ws.add_image(im2, f'A{row+30}')
    
    row += 30
    ws.cell(row, 1, '2. 복토 두께별 작용응력').font = Font(bold=True, size=12)
    row += 1
    for start in range(0, len(selected_thicknesses), 8):
        hs = selected_thicknesses[start:start+8]
        hdr = ['장비명', '규격', 'P(kPa)'] + [f'σ({h:.1f}m)' for h in hs]
        for j, h in enumerate(hdr, 1):
            ws.cell(row, j, h)
            ws.cell(row, j).border = border
            ws.cell(row, j).fill = header_fill
            ws.cell(row, j).alignment = Alignment(horizontal='center', wrap_text=True)
        for _, r in matrix_df.iterrows():
            row += 1
            vals = [r['장비명'], r['규격'], r['접지압P(kPa)']] + [r.get(f'σ({h:.1f}m)', '-') for h in hs]
            for j, v in enumerate(vals, 1):
                ws.cell(row, j, v)
                ws.cell(row, j).border = border
                ws.cell(row, j).alignment = Alignment(horizontal='center')
        row += 2
        
    row += 1
    ws.cell(row, 1, '3. 장비주행성 검토결과').font = Font(bold=True, size=12)
    row += 1
    hdr = ['장비명', '규격', 'P(kPa)', 'b(m)', 'L(m)', 'H(m)', 'σ(kPa)', 'qa(kPa)', 'qa/σ', '판정']
    for j, h in enumerate(hdr, 1):
        ws.cell(row, j, h)
        ws.cell(row, j).border = border
        ws.cell(row, j).fill = header_fill
        ws.cell(row, j).alignment = Alignment(horizontal='center')
    for _, r in summary_df.iterrows():
        row += 1
        vals = [r['장비명'], r['규격'], r['접지압P(kPa)'], r['폭b(m)'], r['길이L(m)'], r['복토두께H(m)'], r['작용응력σ(kPa)'], r['허용지지력qa(kPa)'], r['안전율(qa/σ)'], r['판정']]
        for j, v in enumerate(vals, 1):
            ws.cell(row, j, v)
            ws.cell(row, j).border = border
            ws.cell(row, j).alignment = Alignment(horizontal='center')
            
    row += 3
    ws.cell(row, 1, '4. 장비별 작용응력(σ) vs 허용지지력(qa) 비교 그래프').font = Font(bold=True, size=12)
    row += 1
    im = XLImage(chart_path)
    ratio = im.height / im.width
    im.width = 700
    im.height = int(700 * ratio)
    ws.add_image(im, f'A{row}')
    
    for col, w in {'A': 18, 'B': 14, 'C': 12, 'D': 10, 'E': 10, 'F': 10, 'G': 12, 'H': 12, 'I': 10, 'J': 10}.items():
        ws.column_dimensions[col].width = w
        
    ws.freeze_panes = 'A6'
    ws.print_area = f'A1:J{row+28}'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = .25
    ws.page_margins.right = .25
    ws.page_margins.top = .35
    ws.page_margins.bottom = .35
    ws.print_options.horizontalCentered = True
    
    path = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx').name
    wb.save(path)
    return path


if st.checkbox("🖨️ 인쇄용 구조계산서 양식 열기", value=False, key="open_print_calc_sheet"):
    if "last_summary" not in st.session_state:
        st.warning("위의 [▶ 검토 실행] 버튼을 눌러 계산 결과를 먼저 생성하세요.")
    else:
        summary_df = st.session_state.last_summary
        matrix_df = st.session_state.last_matrix
        st.markdown("### 장비주행성 검토 구조계산서")
        st.markdown(f"**PROJECT:** {proj_name}")
        st.markdown("#### 1. 적용 이론 및 산정식")
        st.latex(r"P=\frac{W}{2bL}\quad(덤프트럭: P=\frac{0.4W}{bL})")
        st.latex(r"\sigma=\frac{PbL(1+\varepsilon)}{(b+2H\tan\theta)(L+2H\tan\theta)}+\gamma_1H")
        st.latex(r"q_a=\frac{1}{F_s}[ (1+0.2b/L)c_2N_{c(2)}F_{cs(2)}F_{cd(2)}+\gamma_1H^2(1+b/L)K_s\tan\phi_1/b+2T\sin\theta/(b+H)]")
        c1, c2c = st.columns(2)
        with c1:
            st.image(f"data:image/png;base64,{THEORY_IMG_A}", caption='Meyerhof-Hanna 층상지반 파괴 메커니즘', use_container_width=True)
        with c2c:
            st.image(f"data:image/png;base64,{THEORY_IMG_B}", caption='하중분산 응력 모델 개념도', use_container_width=True)
        st.markdown("#### 2. 복토 두께별 작용응력")
        st.dataframe(matrix_df, use_container_width=True, hide_index=True, key='print_calc_matrix_df')
        st.markdown("#### 3. 장비주행성 검토결과")
        st.dataframe(summary_df, use_container_width=True, hide_index=True, key='print_calc_summary_df')
        st.markdown("#### 4. 장비별 작용응력(σ) vs 허용지지력(qa) 비교 그래프")
        if 'last_fig' in st.session_state:
            st.plotly_chart(st.session_state.last_fig, use_container_width=True, key="print_calc_chart")

        if st.button('📦 A4 구조계산서 파일 생성', key='generate_a4_calc_files'):
            with st.spinner('PDF·Word·Excel을 A4 인쇄용으로 생성 중입니다...'):
                # 이론 삽도 파일
                a_path = _img_from_b64(THEORY_IMG_A)
                b_path = _img_from_b64(THEORY_IMG_B)
                formula_bytes = _make_formula_images()
                formula_paths = {k: _img_from_b64(base64.b64encode(v).decode()) for k, v in formula_bytes.items()}
                chart_bytes = _make_chart_png(summary_df)
                chart_path = _img_from_b64(base64.b64encode(chart_bytes).decode())
                
                params = {
                    'gamma1': gamma1, 'c2': c2, 'phi1': phi1, 'phi2': phi2,
                    'Ks': ks_lookup(phi1), 'theta': theta_geo, 'T': T_allow,
                    'impact': impact, 'Fs': Fs
                }
                
                pdf_path = build_pdf(summary_df, matrix_df, selected_thicknesses, params, proj_name, {'A': a_path, 'B': b_path}, formula_paths, chart_path)
                docx_path = build_docx(summary_df, matrix_df, selected_thicknesses, params, proj_name, {'A': a_path, 'B': b_path}, formula_paths, chart_path)
                xlsx_path = build_xlsx(summary_df, matrix_df, selected_thicknesses, params, proj_name, {'A': a_path, 'B': b_path}, formula_paths, chart_path)
                
                with open(pdf_path, 'rb') as f:
                    pdf_data = f.read()
                with open(docx_path, 'rb') as f:
                    docx_data = f.read()
                with open(xlsx_path, 'rb') as f:
                    xlsx_data = f.read()
                st.session_state.generated_files = {'pdf': pdf_data, 'docx': docx_data, 'xlsx': xlsx_data}
                
        if 'generated_files' in st.session_state:
            st.success('A4 구조계산서 파일이 생성되었습니다.')
            st.download_button('📄 PDF 다운로드', st.session_state.generated_files['pdf'], '장비주행성_구조계산서_A4.pdf', 'application/pdf', key='download_a4_pdf')
            st.download_button('📝 Word 다운로드', st.session_state.generated_files['docx'], '장비주행성_구조계산서_A4.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', key='download_a4_docx')
            st.download_button('📊 Excel 다운로드', st.session_state.generated_files['xlsx'], '장비주행성_구조계산서_A4.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key='download_a4_xlsx')