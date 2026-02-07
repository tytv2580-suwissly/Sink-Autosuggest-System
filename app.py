import os
import sys
import base64
import subprocess
import platform
from dataclasses import dataclass
from typing import Dict, Tuple, List, Optional

import pandas as pd
import streamlit as st


# =========================
# 0) Path helper
# =========================
def resource_path(rel_path: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel_path)


# =========================
# 1) 기본 설정값
# =========================
DEFAULT_EXCEL_PATH = resource_path("Pre-AI_MATCH_202608_Updated.xlsx")
DEFAULT_DB_SHEET = "DB"
DEFAULT_INPUT_SHEET = "INPUT"
DEFAULT_PDF_DIR = resource_path("pdf")

W_FACILITY = 30
W_SHAPE = 20
W_SPACE = 30
W_EQUIP = 20


# =========================
# 2) 입력 데이터 모델
# =========================
@dataclass
class UserInput:
    facility: str
    shape: str
    W: float
    D: float
    H: float
    distributor: str
    water: str
    gas: str
    electric: str


# =========================
# 3) Rule 정의
# =========================
def evaluate_rules(u: UserInput) -> Tuple[bool, str, str, str]:
    fail_reasons: List[str] = []
    variable_notice: List[str] = []

    if u.H < 900:
        fail_reasons.append("F-01 공간H 미달")

    if (u.distributor == "C") and (u.shape == "ㄱ자") and (u.W < 3000):
        fail_reasons.append("F-02 분배기 C + ㄱ자 + W 부족")

    if (u.gas == "R") and (u.electric == "C") and (u.water == "C"):
        fail_reasons.append("F-03 가스/전기/수도 간섭 조합")

    if u.electric in ["L/R", "R/L"]:
        variable_notice.append("V-01 전기 위치 가변 → 복수안 유지")

    is_designable = (len(fail_reasons) == 0)
    primary_reason = fail_reasons[0] if not is_designable else ""
    all_reasons = "; ".join([r + ";" for r in fail_reasons]) if fail_reasons else ""
    var_notice = "; ".join([v + ";" for v in variable_notice]) if variable_notice else ""

    return is_designable, primary_reason, all_reasons, var_notice


# =========================
# 4) Score 계산
# =========================
def _safe_num(x) -> float:
    try:
        return float(str(x).strip())
    except Exception:
        return 0.0


def space_similarity(u_val: float, db_val: float) -> float:
    u_val = max(1.0, float(u_val))
    db_val = float(db_val)
    sim = 1.0 - abs(db_val - u_val) / u_val
    return max(0.0, min(1.0, sim))


def equip_match(user_val: str, db_val: str, field_name: str) -> float:
    u = str(user_val).strip()
    d = str(db_val).strip()

    if field_name == "electric":
        if u in ["L/R", "R/L"]:
            return 1.0 if d in ["L", "R", "L/R", "R/L"] else 0.0
        if d in ["L/R", "R/L"] and u in ["L", "R"]:
            return 1.0
        return 1.0 if u == d else 0.0

    return 1.0 if u == d else 0.0


def compute_score(u: UserInput, row: pd.Series) -> float:
    facility_db = str(row.get("시설구분", "")).strip()
    shape_db = str(row.get("형태", "")).strip()

    W_db = _safe_num(row.get("공간W (가로)", row.get("공간W", 0)))
    D_db = _safe_num(row.get("공간D (세로)", row.get("공간D", 0)))
    H_db = _safe_num(row.get("공간H (높이)", row.get("공간H", 0)))

    distributor_db = str(row.get("분배기 위치", "")).strip()
    water_db = str(row.get("수도 위치", "")).strip()
    gas_db = str(row.get("가스배관 위치", "")).strip()
    electric_db = str(row.get("전기 위치", "")).strip()

    s_facility = 1.0 if u.facility.strip() == facility_db else 0.0
    s_shape = 1.0 if u.shape.strip() == shape_db else 0.0

    s_space = (
        space_similarity(u.W, W_db) +
        space_similarity(u.D, D_db) +
        space_similarity(u.H, H_db)
    ) / 3.0

    s_equip = (
        equip_match(u.distributor, distributor_db, "distributor") +
        equip_match(u.water, water_db, "water") +
        equip_match(u.gas, gas_db, "gas") +
        equip_match(u.electric, electric_db, "electric")
    ) / 4.0

    score = (W_FACILITY * s_facility) + (W_SHAPE * s_shape) + (W_SPACE * s_space) + (W_EQUIP * s_equip)

    if (s_facility == 0.0) and (s_shape == 0.0):
        score *= 0.4

    return float(score)


# =========================
# 5) PDF 표시(인라인)
# =========================
def render_pdf_inline(pdf_path: str, height: int = 720):
    if not os.path.exists(pdf_path):
        st.error(f"PDF 파일이 존재하지 않습니다: {pdf_path}")
        return

    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    b64 = base64.b64encode(pdf_bytes).decode("utf-8")
    html = f"""
    <iframe src="data:application/pdf;base64,{b64}"
            width="100%" height="{height}" type="application/pdf"></iframe>
    """
    st.components.v1.html(html, height=height, scrolling=True)


# =========================
# 6) 엑셀 로드
# =========================
@st.cache_data(show_spinner=False)
def load_db(excel_path: str, db_sheet: str) -> pd.DataFrame:
    df = pd.read_excel(excel_path, sheet_name=db_sheet, engine="openpyxl")
    if "key" not in df.columns and "Key" in df.columns:
        df = df.rename(columns={"Key": "key"})
    return df


def try_load_input(excel_path: str, input_sheet: str) -> Optional[Dict[str, str]]:
    try:
        df = pd.read_excel(excel_path, sheet_name=input_sheet, engine="openpyxl", header=None)
        data: Dict[str, str] = {}
        for i in range(len(df)):
            k = str(df.iloc[i, 0]).strip()
            v = str(df.iloc[i, 1]).strip()
            if k and k != "nan":
                data[k] = v
        return data
    except Exception:
        return None


def uniq_sorted(df: pd.DataFrame, col: str, fallback: List[str]) -> List[str]:
    if col in df.columns:
        vals = sorted({str(x).strip() for x in df[col].dropna().tolist() if str(x).strip() != ""})
        return vals if vals else fallback
    return fallback


# =========================
# 7) Streamlit UI
# =========================
st.set_page_config(page_title="싱크대 도면 자동 추천(Pre-AI)", layout="wide")
st.title("싱크대 도면 자동 추천 시스템 (엑셀 로직 기반)")

with st.sidebar:
    st.header("환경 설정")
    excel_path = st.text_input("엑셀 경로", value=DEFAULT_EXCEL_PATH)
    db_sheet = st.text_input("DB 시트명", value=DEFAULT_DB_SHEET)
    pdf_dir = st.text_input("도면 PDF 폴더 경로", value=DEFAULT_PDF_DIR)
    st.caption("※ PDF 파일명은 'Key.pdf'와 동일해야 함")

# DB 로드
try:
    db = load_db(excel_path, db_sheet)
except Exception as e:
    st.error(f"DB 로드 실패: {e}")
    st.stop()

auto = try_load_input(excel_path, DEFAULT_INPUT_SHEET) or {}

facility_opts = uniq_sorted(db, "시설구분", ["돌봄센터"])
shape_opts = uniq_sorted(db, "형태", ["일자", "ㄱ자"])
dist_opts = uniq_sorted(db, "분배기 위치", ["L", "C", "R"])
water_opts = uniq_sorted(db, "수도 위치", ["L", "C", "R"])
gas_opts = uniq_sorted(db, "가스배관 위치", ["L", "C", "R"])
elec_opts = uniq_sorted(db, "전기 위치", ["L", "C", "R", "L/R"])

# session_state: TOP1 자동열기 1회 가드
if "opened_top1_key" not in st.session_state:
    st.session_state["opened_top1_key"] = None

colL, colR = st.columns([1, 1])

with colL:
    st.subheader("1) 입력")

    facility_default = auto.get("시설구분", facility_opts[0] if facility_opts else "")
    facility = st.selectbox("시설구분", facility_opts,
                            index=facility_opts.index(facility_default) if facility_default in facility_opts else 0)

    shape_default = auto.get("형태", shape_opts[0] if shape_opts else "")
    shape = st.selectbox("형태", shape_opts,
                         index=shape_opts.index(shape_default) if shape_default in shape_opts else 0)

    W = st.number_input("공간 W (mm)", min_value=0.0, value=float(auto.get("공간 W (가로, mm)", 3000) or 3000))
    D = st.number_input("공간 D (mm)", min_value=0.0, value=float(auto.get("공간 D (세로, mm)", 690) or 690))
    H = st.number_input("공간 H (mm)", min_value=0.0, value=float(auto.get("공간 H (높이, mm)", 800) or 800))

    distributor_default = auto.get("분배기 위치", dist_opts[0] if dist_opts else "C")
    distributor = st.selectbox("분배기 위치", dist_opts,
                               index=dist_opts.index(distributor_default) if distributor_default in dist_opts else 0)

    water_default = auto.get("수도 위치", water_opts[0] if water_opts else "C")
    water = st.selectbox("수도 위치", water_opts,
                         index=water_opts.index(water_default) if water_default in water_opts else 0)

    gas_default = auto.get("가스배관 위치", gas_opts[0] if gas_opts else "R")
    gas = st.selectbox("가스배관 위치", gas_opts,
                       index=gas_opts.index(gas_default) if gas_default in gas_opts else 0)

    elec_default = auto.get("전기 위치", elec_opts[0] if elec_opts else "C")
    electric = st.selectbox("전기 위치", elec_opts,
                            index=elec_opts.index(elec_default) if elec_default in elec_opts else 0)

    auto_open_top1 = st.checkbox("TOP1 도면 자동 열기(윈도우 기본 PDF 뷰어)", value=True)
    run_btn = st.button("추천 실행", type="primary")


with colR:
    st.subheader("2) 출력")

    if run_btn:
        u = UserInput(facility=facility, shape=shape, W=W, D=D, H=H,
                      distributor=distributor, water=water, gas=gas, electric=electric)

        is_ok, primary, all_reason, var_notice = evaluate_rules(u)

        scores: List[Tuple[str, float]] = []
        for _, row in db.iterrows():
            key = str(row.get("key", "")).strip()
            if not key:
                continue
            s = compute_score(u, row)
            scores.append((key, s))

        scores.sort(key=lambda x: x[1], reverse=True)
        top = scores[:3]

        # TOP1 자동 열기
        if auto_open_top1 and len(top) > 0:
            top1_key = top[0][0]
            top1_pdf_path = os.path.join(pdf_dir, f"{top1_key}.pdf")

            if st.session_state["opened_top1_key"] != top1_key:
                if os.path.exists(top1_pdf_path):
                    try:
                        if platform.system() == "Windows":
                            os.startfile(os.path.abspath(top1_pdf_path))  # Windows 전용
                            st.toast("TOP1 PDF를 기본 뷰어로 열었습니다.", icon="✅")
                        else:
                            st.info("배포 환경에서는 자동 팝업이 불가하여 인라인 미리보기로 표시합니다.")
                        st.session_state["opened_top1_key"] = top1_key
                    except Exception as e:
                        st.warning(f"TOP1 자동 열기 실패: {e}")
                else:
                    st.warning("TOP1 PDF가 폴더에 없습니다. Key와 파일명이 정확히 일치하는지 확인하세요.")

        # 출력 테이블
        out_rows = [
            ("설계 가능 여부", "가능" if is_ok else "불가"),
            ("불가 사유(우선)", "" if is_ok else primary),
            ("불가 사유(전체)", "" if is_ok else all_reason),
            ("가변 조건 안내", var_notice if var_notice else ""),
        ]
        st.table(pd.DataFrame(out_rows, columns=["항목", "값"]))

        # 추천 Key
        st.markdown("### 추천 TOP Key")
        for i, (key, s) in enumerate(top, start=1):
            st.write(f"TOP{i}  |  Score={s:.2f}  |  Key={key}")

        # PDF 보기
        st.markdown("### 도면 보기(PDF)")
        if len(top) == 0:
            st.warning("추천 결과가 없습니다. DB에 key 값이 있는지 확인하세요.")
        else:
            tabs = st.tabs([f"TOP{i}" for i in range(1, len(top) + 1)])
            for i, tab in enumerate(tabs):
                key = top[i][0]
                pdf_path = os.path.join(pdf_dir, f"{key}.pdf")
                with tab:
                    st.caption(f"파일: {pdf_path}")
                    if os.path.exists(pdf_path):
                        with open(pdf_path, "rb") as f:
                            st.download_button(label="PDF 다운로드", data=f.read(),
                                               file_name=f"{key}.pdf", mime="application/pdf")
                        if st.button("탐색기에서 PDF 위치 열기", key=f"open_{key}"):
                            if platform.system() == "Windows":
                                subprocess.Popen(f'explorer /select,"{os.path.abspath(pdf_path)}"')
                                st.toast("탐색기를 열었습니다.", icon="✅")
                            else:
                                st.info("배포 환경에서는 탐색기 열기가 지원되지 않습니다.")
                        render_pdf_inline(pdf_path, height=720)
                    else:
                        st.error("PDF를 찾지 못했습니다. 폴더 경로/파일명이 Key와 정확히 일치하는지 확인하세요.")
    else:
        st.info("좌측에서 입력 후 '추천 실행'을 누르세요.")
