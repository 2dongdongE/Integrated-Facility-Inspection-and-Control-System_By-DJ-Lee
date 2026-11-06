import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime, date, timedelta
import io
import altair as alt
import qrcode
from PIL import Image, ImageDraw, ImageFont
import hashlib
import os
import json

# ----------------------------------------------------
# 1. 페이지 레이아웃 및 스타일
# ----------------------------------------------------
st.set_page_config(
    page_title="전사 시설물 에너지 관리 현황",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    [data-testid="InputInstructions"],
    div[data-testid="InputInstructions"] span {
        display: none !important;
    }
    .system-banner {
        background-color: var(--secondary-background-color);
        border: 1px solid var(--border-color, rgba(128, 128, 128, 0.2));
        border-left: 6px solid #2563eb;
        border-radius: 8px;
        padding: 18px 24px;
        margin-bottom: 20px;
    }
    .system-banner h3 {
        margin: 0 0 6px 0;
        font-size: 20px;
        font-weight: 800;
        color: var(--text-color);
    }
    .system-banner p {
        margin: 0;
        font-size: 13px;
        opacity: 0.8;
        color: var(--text-color);
    }
    .admin-badge {
        background-color: rgba(239, 68, 68, 0.15);
        border-left: 4px solid #ef4444;
        padding: 10px 14px;
        border-radius: 4px;
        margin: 16px 0;
    }
    .inspector-badge {
        background-color: rgba(34, 197, 94, 0.15);
        border-left: 4px solid #22c55e;
        padding: 10px 14px;
        border-radius: 4px;
        margin: 16px 0;
    }
    .completed-box {
        background-color: rgba(34, 197, 94, 0.12);
        border: 1px solid rgba(34, 197, 94, 0.4);
        border-left: 6px solid #22c55e;
        border-radius: 8px;
        padding: 16px 20px;
        margin-bottom: 18px;
    }
</style>
""", unsafe_allow_html=True)

# --- 새로고침 후에도 구글 시트 결과 메시지 띄우기 ---
if "sheet_alert" in st.session_state:
    alert_type, alert_msg = st.session_state.pop("sheet_alert")
    if alert_type == "error":
        st.error(alert_msg)
    elif alert_type == "warning":
        st.warning(alert_msg)
    elif alert_type == "success":
        st.success(alert_msg)

# ----------------------------------------------------
# 2. 전역 설정 및 경로 정의
# ----------------------------------------------------
UTILITY_ORDER = ["상수도(메인)", "상수도(식당)", "지하수", "정수", "전기사용량", "LPG가스"]
UTILITY_COLORS = ['#2563eb', '#60a5fa', '#f59e0b', '#10b981', '#ef4444', '#9467bd']
WATER_UTILITIES = ["상수도(메인)", "상수도(식당)", "지하수", "정수"]
DEFAULT_HOST_BASE_URL = "https://administrative-expertise-saying-geological.trycloudflare.com"

db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "facility.db")

def hash_pw(password: str) -> str:
    return hashlib.sha256(password.strip().encode()).hexdigest()

def get_utility_unit(u_type):
    if "전기" in u_type:
        return "kWh"
    elif "LPG" in u_type:
        return "㎥"
    else:
        return "Ton"

# --- 구글 시트 연동 핵심 함수 (저장, 수정, 삭제, 전체 복원) ---
def get_gspread_client():
    try:
        import gspread
        if "gcp_json" in st.secrets:
            raw = st.secrets["gcp_json"]
            creds_dict = json.loads(raw) if isinstance(raw, str) else dict(raw)
            if "private_key" in creds_dict:
                creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
            return gspread.service_account_from_dict(creds_dict)
            
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            if "private_key" in creds_dict:
                creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
            return gspread.service_account_from_dict(creds_dict)

        base_dir = os.path.dirname(os.path.abspath(__file__))
        local_key = os.path.join(base_dir, "service_account.json")
        if os.path.exists(local_key):
            return gspread.service_account(filename=local_key)
    except Exception as e:
        st.session_state["sheet_alert"] = ("error", f"🚨 [구글 키 인증 파싱 에러]: {e}")
    return None

def append_to_sheet(sheet_name: str, row_values: list):
    try:
        gc = get_gspread_client()
        if not gc:
            return
        sh = gc.open("facility data")
        ws = sh.worksheet(sheet_name)
        clean_values = [str(v) if not isinstance(v, (int, float)) else v for v in row_values]
        ws.append_row(clean_values, value_input_option='USER_ENTERED')
    except Exception as e:
        st.session_state["sheet_alert"] = ("error", f"🚨 [구글 시트 저장 실패]: {e}")

def delete_row_from_sheet(sheet_name: str, target_id):
    try:
        gc = get_gspread_client()
        if not gc:
            return
        sh = gc.open("facility data")
        ws = sh.worksheet(sheet_name)
        col1_vals = ws.col_values(1)
        target_str = str(target_id)
        for idx, val in enumerate(col1_vals):
            if val == target_str:
                ws.delete_rows(idx + 1)
                break
    except Exception as e:
        st.session_state["sheet_alert"] = ("error", f"🚨 [구글 시트 행 삭제 실패]: {e}")

def update_cell_by_id(sheet_name: str, target_id, col_idx: int, new_value):
    try:
        gc = get_gspread_client()
        if not gc:
            return
        sh = gc.open("facility data")
        ws = sh.worksheet(sheet_name)
        col1_vals = ws.col_values(1)
        target_str = str(target_id)
        for idx, val in enumerate(col1_vals):
            if val == target_str:
                ws.update_cell(idx + 1, col_idx, new_value)
                break
    except Exception as e:
        st.session_state["sheet_alert"] = ("error", f"🚨 [구글 시트 업데이트 실패]: {e}")

def clear_sheet_records(sheet_name: str):
    try:
        gc = get_gspread_client()
        if not gc:
            return
        sh = gc.open("facility data")
        ws = sh.worksheet(sheet_name)
        all_vals = ws.get_all_values()
        if len(all_vals) > 1:
            ws.delete_rows(2, len(all_vals))
    except Exception as e:
        st.session_state["sheet_alert"] = ("error", f"🚨 [구글 시트 전체 삭제 실패]: {e}")

# --- 구글 시트 ➔ 시스템 DB 원클릭 전체 복원 함수 ---
def restore_data_from_sheets():
    gc = get_gspread_client()
    if not gc:
        return False, "구글 연동 키(Secrets)를 찾을 수 없어 복원할 수 없습니다."
    try:
        sh = gc.open("facility data")
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        # 1. 실제 소화기 자산
        try:
            ws = sh.worksheet("extinguishers")
            rows = ws.get_all_values()[1:]
            for r in rows:
                if len(r) >= 2 and r[0].strip():
                    cursor.execute("INSERT OR REPLACE INTO extinguishers (ext_id, location) VALUES (?, ?)", (r[0].strip(), r[1].strip()))
        except Exception: pass

        # 2. 소화기 점검 이력
        try:
            ws = sh.worksheet("inspection_history")
            rows = ws.get_all_values()[1:]
            for r in rows:
                if len(r) >= 7 and r[0].strip():
                    cursor.execute("""INSERT OR REPLACE INTO inspection_history 
                        (id, ext_id, inspector, status, notes, timestamp, resolved) 
                        VALUES (?, ?, ?, ?, ?, ?, ?)""", 
                        (int(r[0]), r[1], r[2], r[3], r[4], r[5], int(r[6]) if r[6].isdigit() else 0))
        except Exception: pass

        # 3. 유틸리티 검침 데이터
        try:
            ws = sh.worksheet("utility_usage")
            rows = ws.get_all_values()[1:]
            for r in rows:
                if len(r) >= 6 and r[0].strip():
                    cursor.execute("""INSERT OR REPLACE INTO utility_usage 
                        (id, record_date, utility_type, consumption, recorder, notes) 
                        VALUES (?, ?, ?, ?, ?, ?)""", 
                        (int(r[0]), r[1], r[2], float(r[3]) if r[3] else 0.0, r[4], r[5]))
        except Exception: pass

        # 4. 일일 설비 점검 일지
        try:
            ws = sh.worksheet("daily_exact_inspection")
            header = ws.row_values(1)
            rows = ws.get_all_values()[1:]
            if header and rows:
                cols = header
                placeholders = ", ".join(["?"] * len(cols))
                col_names = ", ".join(cols)
                for r in rows:
                    if r and r[0].strip():
                        vals = r + [''] * (len(cols) - len(r))
                        vals[0] = int(vals[0]) if vals[0].isdigit() else vals[0]
                        cursor.execute(f"INSERT OR REPLACE INTO daily_exact_inspection ({col_names}) VALUES ({placeholders})", vals[:len(cols)])
        except Exception: pass

        # 5. 밀폐공간 점검표
        try:
            ws = sh.worksheet("confined_space_inspection")
            rows = ws.get_all_values()[1:]
            for r in rows:
                if len(r) >= 6 and r[0].strip():
                    cursor.execute("""INSERT OR REPLACE INTO confined_space_inspection 
                        (id, record_date, inspector, q_data, notes, timestamp) 
                        VALUES (?, ?, ?, ?, ?, ?)""", 
                        (int(r[0]), r[1], r[2], r[3], r[4], r[5]))
        except Exception: pass

        # 6. 소방펌프 및 밸브 점검표
        try:
            ws = sh.worksheet("fire_pump_inspection")
            rows = ws.get_all_values()[1:]
            for r in rows:
                if len(r) >= 7 and r[0].strip():
                    cursor.execute("""INSERT OR REPLACE INTO fire_pump_inspection 
                        (id, record_date, category, inspector, q_data, notes, timestamp) 
                        VALUES (?, ?, ?, ?, ?, ?, ?)""", 
                        (int(r[0]), r[1], r[2], r[3], r[4], r[5], r[6]))
        except Exception: pass

        # 7. 관리자 계정 명부
        try:
            ws = sh.worksheet("admin_users")
            rows = ws.get_all_values()[1:]
            for r in rows:
                if len(r) >= 8 and r[0].strip():
                    cursor.execute("""INSERT OR REPLACE INTO admin_users 
                        (id, username, password_hash, name, dept, role, approved, created_at) 
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)""", 
                        (int(r[0]), r[1], r[2], r[3], r[4], r[5], int(r[6]) if r[6].isdigit() else 0, r[7]))
        except Exception: pass

        conn.commit()
        conn.close()
        return True, "구글 시트에 보관된 모든 점검 및 계정 데이터를 시스템으로 성공적으로 복원했습니다!"
    except Exception as e:
        return False, f"데이터 복원 실패: {e}"

# ----------------------------------------------------
# 3. 통합 데이터베이스 초기화 및 조작 함수
# ----------------------------------------------------
def init_db():
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
        
    cursor.execute('''CREATE TABLE IF NOT EXISTS inspection_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        ext_id TEXT, 
        inspector TEXT, 
        status TEXT, 
        notes TEXT, 
        timestamp TEXT, 
        resolved INTEGER DEFAULT 0
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS extinguishers (
        ext_id TEXT PRIMARY KEY, 
        location TEXT
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS utility_usage (
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        record_date TEXT, 
        utility_type TEXT, 
        consumption REAL, 
        recorder TEXT, 
        notes TEXT
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS daily_exact_inspection (
        id INTEGER PRIMARY KEY AUTOINCREMENT, record_date TEXT, inspector TEXT,
        al150_t1_1 TEXT, al150_t1_2 TEXT, al150_t1_3 TEXT, al150_t2_1 TEXT, al150_t2_2 TEXT, al150_t2_3 TEXT,
        oil_temp_1 TEXT, oil_temp_2 TEXT, oil_temp_3 TEXT, run_time_1 TEXT, run_time_2 TEXT, run_time_3 TEXT,
        load_time_1 TEXT, load_time_2 TEXT, load_time_3 TEXT, p1_1 TEXT, p1_2 TEXT, p1_3 TEXT,
        p2_1 TEXT, p2_2 TEXT, p2_3 TEXT, oil_p_1 TEXT, oil_p_2 TEXT, oil_p_3 TEXT,
        suc_t_1 TEXT, suc_t_2 TEXT, suc_t_3 TEXT, oil_lvl_1 TEXT, oil_lvl_2 TEXT, oil_lvl_3 TEXT,
        cool_p_1 TEXT, cool_p_2 TEXT, cool_p_3 TEXT, cool_in_1 TEXT, cool_in_2 TEXT, cool_in_3 TEXT,
        cool_out_1 TEXT, cool_out_2 TEXT, cool_out_3 TEXT, dryer_low_1 TEXT, dryer_low_2 TEXT, dryer_low_3 TEXT,
        dryer_high_1 TEXT, dryer_high_2 TEXT, dryer_high_3 TEXT, dryer_air_1 TEXT, dryer_air_2 TEXT, dryer_air_3 TEXT,
        dryer_drain_1 TEXT, dryer_drain_2 TEXT, dryer_drain_3 TEXT, tx_out_p TEXT, tx_oil_p TEXT, tx_suc_t TEXT, tx_oil_t TEXT,
        tx_run_t TEXT, tx_load_t TEXT, tx_vib TEXT, tx_gap TEXT, tx_cin_t TEXT, tx_cout_t TEXT, tx_head_p TEXT,
        tx_dryer_low TEXT, tx_dryer_high TEXT, tx_dryer_air TEXT, tx_dryer_drain TEXT,
        pneu_leak TEXT, pneu_trap TEXT, pneu_mois TEXT, pneu_vib TEXT, pneu_curr TEXT,
        tower_fan TEXT, tower_mach TEXT, tower_ext TEXT, tower_vib TEXT, tower_fill TEXT,
        w_u_lvl TEXT, w_i_lvl TEXT, w_p_lvl TEXT, w_u_stat TEXT, w_h_stat TEXT, w_m_stat TEXT,
        oil_sep TEXT, cool_cap TEXT, cool_t TEXT, hot_p TEXT, hot_t TEXT, hot_stat TEXT,
        temp_mach TEXT, temp_outside TEXT, p_comp_p TEXT, p_comp_n TEXT, p_comp_s TEXT,
        p_fac_p TEXT, p_fac_n TEXT, p_fac_s TEXT, p_off_p TEXT, p_off_n TEXT, p_off_s TEXT, etc_notes TEXT, timestamp TEXT
    )''')
    
    cursor.execute('''CREATE TABLE IF NOT EXISTS confined_space_inspection (
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        record_date TEXT, 
        inspector TEXT, 
        q_data TEXT, 
        notes TEXT, 
        timestamp TEXT
    )''')

    cursor.execute('''CREATE TABLE IF NOT EXISTS fire_pump_inspection (
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        record_date TEXT, 
        category TEXT, 
        inspector TEXT, 
        q_data TEXT, 
        notes TEXT, 
        timestamp TEXT
    )''')

    cursor.execute('''CREATE TABLE IF NOT EXISTS admin_users (
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        username TEXT UNIQUE, 
        password_hash TEXT, 
        name TEXT, 
        dept TEXT, 
        role TEXT DEFAULT 'admin', 
        approved INTEGER DEFAULT 0, 
        created_at TEXT
    )''')

    # 마스터 계정 초기 등록
    cursor.execute("SELECT COUNT(*) FROM admin_users WHERE username = 'dolee'")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO admin_users (username, password_hash, name, dept, role, approved, created_at) VALUES ('dolee', ?, '마스터 관리자', '', 'master', 1, ?)", 
                       (hash_pw("2051"), datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()

    # 실제 현장 소화기 181대 전체 정식 등록 (100% 온전하게 유지)
    cursor.execute("SELECT COUNT(*) FROM extinguishers")
    if cursor.fetchone()[0] < 181:
        cursor.execute("DELETE FROM extinguishers")
        initial_extinguishers = [
            ("ETC-1", "정문 경비실"), ("1-1-1", "중앙현관 1"), ("1-1-2", "중앙현관 2"), ("1-1-3", "전산 서버실"), ("1-1-4", "1층 사무실 1"),
            ("1-1-5", "1층 교육장 1"), ("1-1-6", "1층 교육장 2"), ("1-1-7", "1층사무실2"), ("1-2-1", "재무 문서고"), ("1-2-2", "2층 사무실 1"),
            ("1-2-3", "2층 사무실 2"), ("1-2-4", "2층 사무실 3"), ("1-2-5", "2층 사무실 4"), ("1-2-6", "2층 사무실 5"), ("1-2-7", "2층 복도 1"),
            ("1-2-8", "2층 복도 2"), ("1-2-9", "2층 사무실 1"), ("1-2-10", "2층사무실 2"), ("1-2-11", "2층사무실 3"), ("1-2-12", "2층사무실 4"),
            ("1-1-8", "29S측 탈의실"), ("1-1-9", "방화문 29S"), ("1-1-10", "방화문 29S"), ("1-1-11", "23S B1"), ("1-1-12", "23S B2"),
            ("1-1-13", "23S B2"), ("1-1-14", "식당 홀 1"), ("1-1-15", "식당 홀 2"), ("1-1-16", "조리실 1"), ("1-1-17", "조리실 2"),
            ("1-1-18", "29S 1"), ("1-1-19", "29S 2"), ("1-1-20", "29S 3"), ("1-1-21", "29S 4"), ("1-1-22", "ALT A"),
            ("1-1-23", "Alt B"), ("1-1-24", "방화문 Alt 1"), ("1-1-25", "방화문 Alt 2"), ("1-1-26", "IDM 1"), ("1-1-27", "IDM 2"),
            ("1-1-28", "IDM 3"), ("1-1-29", "방화문 IDM1"), ("1-1-30", "방화문 IDM2"), ("1-1-31", "탈의실 IDM"), ("1-1-32", "공무작업실"),
            ("1-1-33", "자재 자동문"), ("1-1-34", "자재 사무실"), ("1-1-35", "직장 사무실"), ("1-1-36", "품질 사무실1"), ("1-1-37", "품질 사무실2"),
            ("1-1-38", "완제품 탈의실"), ("1-1-39", "완제품 방화문"), ("1-1-40", "완제품 방화문"), ("1-1-41", "완제품 창고1"), ("1-1-42", "완제품 창고2"),
            ("1-1-43", "완제품 창고3"), ("1-1-44", "B Line 방화문1"), ("1-1-45", "B Line 방화문2"), ("1-1-46", "구DG 방화문1"), ("1-1-47", "구DG 방화문2"),
            ("1-1-48", "구DG 탈의실"), ("ETC-62", "PG측 파고라"), ("ETC-63", "후문 경비실"), ("ETC-64", "완제품 지게차 1"), ("ETC-65", "완제품 지게차 2"),
            ("ETC-66", "공박스 보관소 1"), ("ETC-67", "공박스 보관소 2"), ("ETC-68", "공박스 보관소 3"), ("ETC-69", "자재창고 앞 파고라"), ("ETC-70", "시편 측정대 1"),
            ("ETC-71", "시편 측정대 2"), ("ETC-72", "하역장 1"), ("ETC-73", "하역장 2"), ("ETC-74", "지게차 충전소1"), ("ETC-75", "지게차 충전소2"),
            ("ETC-76", "시설 창고"), ("ETC-77", "유류저장소 1"), ("ETC-78", "유류저장소 2"), ("ETC-79", "유류저장소 3"), ("ETC-80", "유류저장소 4"),
            ("ETC-81", "기계실 1"), ("ETC-82", "기계실 2"), ("ETC-83", "기계실 3"), ("ETC-84", "기계실 4"), ("ETC-85", "기계실 5"),
            ("ETC-86", "기계실 6"), ("ETC-87", "전기실 1"), ("ETC-88", "전기실 2"), ("ETC-89", "전기실 3"), ("ETC-90", "전기실 4"),
            ("ETC-91", "전기실 5"), ("ETC-92", "Engine Test Room 1"), ("ETC-93", "Engine Test Room 2"), ("ETC-94", "Engine Test Room 3"), ("ETC-95", "Engine Test Room 4"),
            ("ETC-96", "외국인 휴게실"), ("ETC-97", "CFS 창고 1"), ("ETC-98", "CFS 창고 2"), ("ETC-99", "CFS 창고 3"), ("ETC-100", "QA Lab 1"),
            ("ETC-101", "QA Lab 2"), ("ETC-102", "QA Lab 3"), ("ETC-103", "LPG 탱크 1"), ("ETC-104", "LPG 탱크 2"), ("ETC-105", "DR23S A 외부1"),
            ("ETC-106", "DR23S A 외부2"), ("10-1-1", "DR23S A 내부1"), ("10-1-2", "DR23S A 내부2"), ("10-1-3", "DR23S A 내부3"), ("10-1-4", "DR23S A 내부4"),
            ("10-1-5", "DR23S A 내부5"), ("10-1-6", "DR23S A 내부6"), ("ETC-113", "신축동 앞 파고라"), ("ETC-114", "신축동 측 집진기"), ("10-1-7", "신축동 로비"),
            ("10-1-8", "신축동 1F 1"), ("10-1-9", "신축동 1F 2"), ("10-1-10", "신축동 1F 3"), ("10-1-11", "신축동 1F 4"), ("10-1-12", "진동시험실"),
            ("10-1-13", "충격내구"), ("10-1-14", "내구시험실 1"), ("10-1-15", "내구시험실 2"), ("10-1-16", "내구시험실 3"), ("10-1-17", "무향실"),
            ("10-1-18", "부식시험실"), ("10-2-1", "신축동 2F 1"), ("10-2-2", "신축동 2F 2"), ("10-2-3", "시작품제작실 1"), ("10-2-4", "시작품제작실 2"),
            ("10-2-5", "시작품제작실 3"), ("12-1-1", "1F 로비"), ("12-1-2", "휴게실"), ("12-1-3", "측정실"), ("ETC-135", "위험물옥내저장소 1"),
            ("ETC-136", "위험물옥내저장소 2"), ("ETC-137", "위험물옥내저장소 3"), ("ETC-138", "위험물옥내저장소 4"), ("ETC-139", "위험물옥내저장소 5"), ("12-1-4", "Stator 1"),
            ("12-1-5", "Stator 2"), ("12-1-6", "Stator 3"), ("12-1-7", "Stator 4"), ("12-1-8", "Stator 5"), ("12-1-9", "Motor 1"),
            ("12-1-10", "Motor 2"), ("12-1-11", "Motor 3"), ("12-1-12", "창고 1"), ("12-1-13", "창고 2"), ("12-1-14", "창고 3"),
            ("12-1-15", "하역장 1"), ("12-1-16", "하역장 2"), ("12-1-17", "하역장 3"), ("12-1-18", "창고계단"), ("12-B1", "12동 B1, 펌프실"),
            ("12-2-8", "2F 계단"), ("12-3-1", "3F 창고 1"), ("12-3-2", "3F 창고 2"), ("12-3-3", "3F 창고 3"), ("12-3-4", "3F보관소 1"),
            ("12-3-5", "3F보관소 2"), ("12-3-6", "3F보관소 3"), ("12-3-7", "3F보관소 4"), ("12-3-8", "3F Invt 1"), ("12-3-9", "3F Invt 2"),
            ("12-3-10", "3F Invt 3"), ("12-3-11", "3F Invt 4"), ("12-3-12", "3F 전산실"), ("12-3-13", "3F 복도"), ("12-3-14", "Invt 입구"),
            ("12-3-15", "Invt 방풍실"), ("12-3-16", "분석실"), ("12-3-17", "회의실"), ("12-3-18", "휴게실"), ("12-2-1", "2F 남자 탈의실"),
            ("12-2-2", "2F 복도 1"), ("12-2-3", "2F 온수탱크실"), ("12-2-4", "2F 여자 탈의실"), ("12-2-5", "2F 복도 2"), ("12-2-6", "2F 제품연구팀"),
            ("12-2-7", "2F 회의실")
        ]
        for ext_id_val, location in initial_extinguishers:
            cursor.execute('INSERT OR IGNORE INTO extinguishers (ext_id, location) VALUES (?, ?)', (ext_id_val, location))
        conn.commit()

    # [핵심] 기존 DB에 남아있던 365일 치 가짜 검침 더미 데이터 영구 삭제
    cursor.execute("DELETE FROM utility_usage WHERE notes LIKE '%정기 검침%' OR notes LIKE '%이상 검침%'")
    conn.commit()

    conn.close()

def authenticate_user(username, password):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT username, name, role, approved FROM admin_users WHERE username = ? AND password_hash = ?", (username.strip(), hash_pw(password)))
    user = cur.fetchone()
    conn.close()
    return {"username": user[0], "name": user[1], "role": user[2], "approved": user[3]} if user else None

def register_admin_request(username, password, name):
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("INSERT INTO admin_users (username, password_hash, name, dept, role, approved, created_at) VALUES (?, ?, ?, '', 'admin', 0, ?)", 
                     (username.strip(), hash_pw(password), name.strip(), now_str))
        new_id = cur.lastrowid
        conn.commit()
        append_to_sheet("admin_users", [new_id, username.strip(), hash_pw(password), name.strip(), '', 'admin', 0, now_str])
        return True
    except:
        return False
    finally:
        conn.close()

def get_pending_admin_requests():
    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query("SELECT id, username AS 아이디, name AS 성명, created_at AS 신청일시 FROM admin_users WHERE approved = 0", conn)
    conn.close()
    return df

def approve_admin_request(user_id):
    conn = sqlite3.connect(db_path)
    conn.execute("UPDATE admin_users SET approved = 1 WHERE id = ?", (user_id,))
    conn.commit()
    conn.close()
    update_cell_by_id("admin_users", user_id, 7, 1)

def reject_or_delete_user(user_id):
    conn = sqlite3.connect(db_path)
    conn.execute("DELETE FROM admin_users WHERE id = ? AND role != 'master'", (user_id,))
    conn.commit()
    conn.close()
    delete_row_from_sheet("admin_users", user_id)

def get_all_approved_users():
    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query("SELECT id, username AS 아이디, name AS 성명, role AS 권한, created_at AS 등록일시 FROM admin_users WHERE approved = 1", conn)
    conn.close()
    return df

def save_inspection(ext_id_val, inspector, status, notes):
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute('INSERT INTO inspection_history (ext_id, inspector, status, notes, timestamp, resolved) VALUES (?, ?, ?, ?, ?, 0)', 
                 (ext_id_val, inspector, status, notes, now_str))
    new_id = cur.lastrowid
    conn.commit()
    conn.close()
    append_to_sheet("inspection_history", [new_id, ext_id_val, inspector, status, notes, now_str, 0])

def mark_as_resolved(history_id):
    conn = sqlite3.connect(db_path)
    conn.execute("UPDATE inspection_history SET resolved = 1 WHERE id = ?", (history_id,))
    conn.commit()
    conn.close()
    update_cell_by_id("inspection_history", history_id, 7, 1)

def delete_inspection_history(history_id):
    conn = sqlite3.connect(db_path)
    conn.execute("DELETE FROM inspection_history WHERE id = ?", (history_id,))
    conn.commit()
    conn.close()
    delete_row_from_sheet("inspection_history", history_id)

def delete_all_inspection_history():
    conn = sqlite3.connect(db_path)
    conn.execute("DELETE FROM inspection_history")
    conn.execute("DELETE FROM sqlite_sequence WHERE name='inspection_history'")
    conn.commit()
    conn.close()
    clear_sheet_records("inspection_history")

def add_extinguisher(ext_id_val, location):
    conn = sqlite3.connect(db_path)
    try:
        conn.execute("INSERT INTO extinguishers (ext_id, location) VALUES (?, ?)", (ext_id_val, location))
        conn.commit()
        append_to_sheet("extinguishers", [ext_id_val, location])
        return True
    except:
        return False
    finally:
        conn.close()

def save_exact_daily_log(d):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("PRAGMA table_info(daily_exact_inspection)")
    cols = [row[1] for row in cursor.fetchall() if row[1] != 'id']
    d_copy = dict(d)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    d_copy['timestamp'] = now_str
    c_in = [c for c in cols if c in d_copy]
    v_in = [d_copy[c] for c in c_in]
    cursor.execute(f"INSERT INTO daily_exact_inspection ({', '.join(c_in)}) VALUES ({', '.join(['?'] * len(v_in))})", v_in)
    new_id = cursor.lastrowid
    conn.commit()
    conn.close()
    row_values = [new_id] + [d_copy.get(col, '') for col in cols]
    append_to_sheet("daily_exact_inspection", row_values)

def get_exact_daily_logs():
    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query("SELECT * FROM daily_exact_inspection ORDER BY record_date DESC, id DESC", conn)
    conn.close()
    return df

def delete_exact_daily_log(log_id):
    conn = sqlite3.connect(db_path)
    conn.execute("DELETE FROM daily_exact_inspection WHERE id = ?", (log_id,))
    conn.commit()
    conn.close()
    delete_row_from_sheet("daily_exact_inspection", log_id)

def delete_all_exact_daily_logs():
    conn = sqlite3.connect(db_path)
    conn.execute("DELETE FROM daily_exact_inspection")
    conn.execute("DELETE FROM sqlite_sequence WHERE name='daily_exact_inspection'")
    conn.commit()
    conn.close()
    clear_sheet_records("daily_exact_inspection")

def get_all_extinguishers():
    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query("SELECT ext_id AS 소화기ID, location AS 설치위치 FROM extinguishers ORDER BY ext_id", conn)
    conn.close()
    return df

def get_history_with_id():
    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query("""
        SELECT h.id, h.timestamp AS 점검시간, h.ext_id AS 소화기ID, e.location AS 설치위치, 
               h.inspector AS 점검자, h.status AS 상태, h.notes AS 특이사항, 
               CASE WHEN h.resolved = 1 THEN '조치완료' 
                    WHEN h.status != '정상' THEN '조치필요' 
                    ELSE '정상' END AS 조치상태 
        FROM inspection_history h 
        LEFT JOIN extinguishers e ON h.ext_id = e.ext_id 
        ORDER BY h.id DESC
    """, conn)
    conn.close()
    return df

def get_unresolved_issues():
    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query("""
        SELECT h.id, h.timestamp AS 점검시간, h.ext_id AS 소화기ID, e.location AS 설치위치, 
               h.inspector AS 점검자, h.status AS 상태, h.notes AS 특이사항 
        FROM inspection_history h 
        LEFT JOIN extinguishers e ON h.ext_id = e.ext_id 
        WHERE h.status != '정상' AND (h.resolved = 0 OR h.resolved IS NULL) 
        ORDER BY h.id DESC
    """, conn)
    conn.close()
    return df

def get_utility_df():
    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query("SELECT id, record_date AS 기록일자, utility_type AS 유틸리티종류, consumption AS 지침값, recorder AS 기록자, notes AS 특이사항 FROM utility_usage ORDER BY record_date ASC, id ASC", conn)
    conn.close()
    return df

def save_utility_usage(record_date, utility_type, consumption, recorder, notes):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute('INSERT INTO utility_usage (record_date, utility_type, consumption, recorder, notes) VALUES (?, ?, ?, ?, ?)', (record_date, utility_type, consumption, recorder, notes))
    new_id = cur.lastrowid
    conn.commit()
    conn.close()
    append_to_sheet("utility_usage", [new_id, str(record_date), utility_type, consumption, recorder, notes])

def delete_utility_usage(util_id):
    conn = sqlite3.connect(db_path)
    conn.execute("DELETE FROM utility_usage WHERE id = ?", (util_id,))
    conn.commit()
    conn.close()
    delete_row_from_sheet("utility_usage", util_id)

def delete_all_utility_usage():
    conn = sqlite3.connect(db_path)
    conn.execute("DELETE FROM utility_usage")
    conn.execute("DELETE FROM sqlite_sequence WHERE name='utility_usage'")
    conn.commit()
    conn.close()
    clear_sheet_records("utility_usage")

def save_confined_space_log(record_date, inspector, q_data, notes):
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    q_str = json.dumps(q_data, ensure_ascii=False)
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("INSERT INTO confined_space_inspection (record_date, inspector, q_data, notes, timestamp) VALUES (?, ?, ?, ?, ?)",
                 (str(record_date), inspector.strip(), q_str, notes.strip(), now_str))
    new_id = cur.lastrowid
    conn.commit()
    conn.close()
    append_to_sheet("confined_space_inspection", [new_id, str(record_date), inspector.strip(), q_str, notes.strip(), now_str])

def get_confined_space_logs():
    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query("SELECT * FROM confined_space_inspection ORDER BY record_date DESC, id DESC", conn)
    conn.close()
    return df

def delete_confined_space_log(log_id):
    conn = sqlite3.connect(db_path)
    conn.execute("DELETE FROM confined_space_inspection WHERE id = ?", (log_id,))
    conn.commit()
    conn.close()
    delete_row_from_sheet("confined_space_inspection", log_id)

def delete_all_confined_space_logs():
    conn = sqlite3.connect(db_path)
    conn.execute("DELETE FROM confined_space_inspection")
    conn.execute("DELETE FROM sqlite_sequence WHERE name='confined_space_inspection'")
    conn.commit()
    conn.close()
    clear_sheet_records("confined_space_inspection")

def save_fire_pump_log(record_date, category, inspector, q_data, notes):
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    q_str = json.dumps(q_data, ensure_ascii=False)
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("INSERT INTO fire_pump_inspection (record_date, category, inspector, q_data, notes, timestamp) VALUES (?, ?, ?, ?, ?, ?)",
                 (str(record_date), category, inspector.strip(), q_str, notes.strip(), now_str))
    new_id = cur.lastrowid
    conn.commit()
    conn.close()
    append_to_sheet("fire_pump_inspection", [new_id, str(record_date), category, inspector.strip(), q_str, notes.strip(), now_str])

def get_fire_pump_logs():
    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query("SELECT * FROM fire_pump_inspection ORDER BY record_date DESC, id DESC", conn)
    conn.close()
    return df

def delete_fire_pump_log(log_id):
    conn = sqlite3.connect(db_path)
    conn.execute("DELETE FROM fire_pump_inspection WHERE id = ?", (log_id,))
    conn.commit()
    conn.close()
    delete_row_from_sheet("fire_pump_inspection", log_id)

def delete_all_fire_pump_logs():
    conn = sqlite3.connect(db_path)
    conn.execute("DELETE FROM fire_pump_inspection")
    conn.execute("DELETE FROM sqlite_sequence WHERE name='fire_pump_inspection'")
    conn.commit()
    conn.close()
    clear_sheet_records("fire_pump_inspection")

def get_recent_inspection_this_month(ext_id_val):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT timestamp, inspector, status, notes FROM inspection_history WHERE ext_id = ? AND timestamp LIKE ? ORDER BY id DESC LIMIT 1", (ext_id_val, f"{datetime.now().strftime('%Y-%m')}%"))
    res = cur.fetchone()
    conn.close()
    return res

def get_utility_record_today(utility_type, target_date_str):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT id, record_date, recorder, consumption, notes FROM utility_usage WHERE utility_type = ? AND record_date = ? ORDER BY id DESC LIMIT 1", (utility_type, target_date_str))
    res = cur.fetchone()
    conn.close()
    return res

def get_daily_inspection_today(target_date_str):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT id, record_date, inspector, timestamp FROM daily_exact_inspection WHERE record_date = ? ORDER BY id DESC LIMIT 1", (target_date_str,))
    res = cur.fetchone()
    conn.close()
    return res

def get_confined_space_today(target_date_str):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT id, record_date, inspector FROM confined_space_inspection WHERE record_date = ? ORDER BY id DESC LIMIT 1", (target_date_str,))
    res = cur.fetchone()
    conn.close()
    return res

def get_fire_pump_today(category, target_date_str):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT id, record_date, inspector FROM fire_pump_inspection WHERE category = ? AND record_date = ? ORDER BY id DESC LIMIT 1", (category, target_date_str))
    res = cur.fetchone()
    conn.close()
    return res

def generate_qr_image_with_label(main_text, sub_text, url):
    qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=10, border=3)
    qr.add_data(url)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="#0f172a", back_color="white").convert('RGB')
    new_img = Image.new('RGB', (qr_img.size[0], qr_img.size[1] + 70), 'white')
    new_img.paste(qr_img, (0, 0))
    draw = ImageDraw.Draw(new_img)
    try:
        f_b = ImageFont.truetype(r"C:\Windows\Fonts\malgunbd.ttf", 15)
        f_n = ImageFont.truetype(r"C:\Windows\Fonts\malgun.ttf", 13)
    except:
        f_b = f_n = ImageFont.load_default()
    draw.text((15, qr_img.size[1] + 8), main_text, fill="#0f172a", font=f_b)
    draw.text((15, qr_img.size[1] + 34), sub_text, fill="#64748b", font=f_n)
    buf = io.BytesIO()
    new_img.save(buf, format="PNG")
    return buf.getvalue()

def check_factory_anomalies(u_df):
    alerts = []
    if u_df.empty:
        return alerts
    for u_type in UTILITY_ORDER:
        sub = u_df[u_df['유틸리티종류'] == u_type].sort_values(['기록일자', 'id']).copy()
        if len(sub) < 2:
            continue
        sub['date_dt'] = pd.to_datetime(sub['기록일자'])
        sub['days_diff'] = (sub['date_dt'] - sub['date_dt'].shift(1)).dt.days.fillna(1).clip(lower=1)
        sub['daily_usage'] = ((sub['지침값'] - sub['지침값'].shift(1)).clip(lower=0) / sub['days_diff']).round(1)
        
        latest = sub.iloc[-1]
        valid_past = sub.iloc[:-1]['daily_usage'].dropna()
        if len(valid_past) < 1:
            continue
        avg_d = valid_past.mean()
        cur_d = latest['daily_usage']
        d_diff = int(latest['days_diff'])

        if u_type in WATER_UTILITIES:
            if avg_d > 0 and cur_d >= (avg_d * 1.25) and (cur_d - avg_d) >= 4.0:
                alerts.append({
                    'id': f"leak_{u_type}", 
                    'level': 'error', 
                    'title': f'🚨 [누수 의심 경보] {u_type}', 
                    'body': f"• 검침 일자: {latest['기록일자']}\n\n• 금일 사용량(1일 환산): {cur_d:,.0f} Ton (평소 평일 평균 {avg_d:,.0f} Ton)\n\n• 분석 결과: 평소 대비 {cur_d/avg_d:.1f}배 급증"
                })
                continue
            if len(sub) >= 4:
                r3 = sub.iloc[-3:]['daily_usage'].tolist()
                if r3[0] < r3[1] < r3[2] and (((r3[2] - r3[0]) / r3[0] * 100) >= 15.0) and ((r3[2] - r3[0]) >= 2.0):
                    alerts.append({
                        'id': f"micro_{u_type}", 
                        'level': 'warning', 
                        'title': f'⚠️ [미세 누수 주의] {u_type}', 
                        'body': f"• 최근 사용량 추이: {r3[0]:,.0f} ➔ {r3[1]:,.0f} ➔ {r3[2]:,.0f} Ton\n\n• 분석 결과: 3회 연속 점진적 증가 감지"
                    })
            if d_diff >= 3 and cur_d >= (avg_d * 0.4) and cur_d >= 4.0:
                alerts.append({
                    'id': f"weekend_{u_type}", 
                    'level': 'warning', 
                    'title': f'🔍 [주말 기저 누수 의심] {u_type}', 
                    'body': f"• 비가동 기간 총 소모량: {(cur_d * d_diff):,.0f} Ton (일평균 {cur_d:,.0f} Ton)\n\n• 분석 결과: 공장 휴무 기간 동안 평일의 40% 이상 물이 지속 소모됨"
                })
        else:
            if avg_d > 0 and cur_d >= (avg_d * 1.8) and (cur_d - avg_d) >= 10.0:
                unit = 'kWh' if '전기' in u_type else '㎥'
                alerts.append({
                    'id': f"over_{u_type}", 
                    'level': 'error', 
                    'title': f'🚨 [과다 소비 경고] {u_type}', 
                    'body': f"• 금일 사용량: {cur_d:,.1f} {unit} (평소 평균 {avg_d:,.1f} {unit})\n\n• 분석 결과: 평소 대비 {cur_d/avg_d:.1f}배 급증"
                })
    return alerts

init_db()

# ----------------------------------------------------
# 4. 세션 상태 관리 및 라우팅 파라미터 처리
# ----------------------------------------------------
if 'auth_user' not in st.session_state: st.session_state['auth_user'] = None
if 'custom_host_url' not in st.session_state: st.session_state['custom_host_url'] = DEFAULT_HOST_BASE_URL
if 'acknowledged_alerts' not in st.session_state: st.session_state['acknowledged_alerts'] = set()

is_admin = st.session_state['auth_user'] is not None
active_user = st.session_state['auth_user']
is_master = is_admin and active_user['role'] == 'master'
active_host_url = st.session_state['custom_host_url']

params = st.query_params
ext_id_param = params.get('ext_id', None)
utility_param = params.get('utility', None)
daily_scan_param = params.get('daily_scan', None)
cf_scan_param = params.get('cf_scan', None)
fp_scan_param = params.get('fp_scan', None)

menu_list = ["📊 통합 종합 대시보드", "📋 일일 통합 점검", "🧯 소화기 점검 관리", "⚡ 유틸리티 사용량 관리", "🧱 밀폐공간 점검", "🧯 소방 펌프 및 밸브 점검"]
default_menu_idx = 0
if daily_scan_param: default_menu_idx = 1
elif ext_id_param: default_menu_idx = 2
elif utility_param: default_menu_idx = 3
elif cf_scan_param: default_menu_idx = 4
elif fp_scan_param: default_menu_idx = 5

# ----------------------------------------------------
# 5. 사이드바 메뉴 및 관리자 도구
# ----------------------------------------------------
st.sidebar.markdown("### 🏢 시설안전 통합관제")
menu = st.sidebar.radio("📌 메뉴 선택", menu_list, index=default_menu_idx, key="main_menu_radio")
st.sidebar.markdown("---")

if not is_admin:
    st.sidebar.markdown("<div class='inspector-badge'><span style='color:#22c55e;font-size:13px;font-weight:800;'>👷 현장 점검자 모드</span><br><span style='font-size:11px;opacity:0.8;'>정기 점검 및 등록 전용</span></div>", unsafe_allow_html=True)
    with st.sidebar.expander("🔐 관리자 로그인 (원격PC)", expanded=False):
        with st.form("auth_login_form"):
            l_id = st.text_input("아이디")
            l_pw = st.text_input("비밀번호", type="password")
            if st.form_submit_button("로그인", use_container_width=True):
                u_info = authenticate_user(l_id, l_pw)
                if u_info and u_info['approved'] == 1:
                    st.session_state['auth_user'] = u_info
                    st.rerun()
                elif u_info: 
                    st.warning("승인 대기 중인 계정입니다.")
                else: 
                    st.error("아이디/비밀번호 오류")

    with st.sidebar.expander("📝 관리자 계정 신청", expanded=False):
        with st.form("auth_reg_form"):
            r_id = st.text_input("희망 아이디")
            r_pw = st.text_input("희망 비밀번호", type="password")
            r_nm = st.text_input("성명 (이름)")
            if st.form_submit_button("신청 제출", use_container_width=True):
                if r_id and r_pw and r_nm:
                    if register_admin_request(r_id, r_pw, r_nm): 
                        st.success("마스터 승인 대기 중!")
                    else: 
                        st.warning("이미 존재하는 아이디입니다.")
                else: 
                    st.error("모든 항목을 입력해주세요.")
else:
    st.sidebar.markdown(f"<div class='admin-badge'><span style='color:#ef4444;font-size:13px;font-weight:800;'>🔐 {'마스터' if is_master else '운영'} 관리자</span><br><span style='font-size:12px;font-weight:700;'>{active_user['name']}</span></div>", unsafe_allow_html=True)
    if st.sidebar.button("🚪 로그아웃", use_container_width=True, key="btn_logout_main"):
        st.session_state['auth_user'] = None
        st.rerun()

    st.sidebar.markdown("---")
    st.sidebar.subheader("⚙️ 관리자 도구")

    # [복원 기능] 구글 시트 원클릭 전체 데이터 복원 기능
    with st.sidebar.expander("🔄 구글 시트 데이터 복원", expanded=False):
        st.caption("서버 초기화 시 구글 시트에 누적된 모든 데이터를 앱 내부로 1초 만에 복구합니다.")
        if st.button("구글 시트 데이터 전체 복원 실행", use_container_width=True, key="btn_restore_all_from_sheets"):
            with st.spinner("구글 시트에서 전체 데이터를 복원하는 중..."):
                ok, msg = restore_data_from_sheets()
                if ok:
                    st.session_state["sheet_alert"] = ("success", f"✅ {msg}")
                else:
                    st.session_state["sheet_alert"] = ("error", f"🚨 {msg}")
                st.rerun()

    if is_master:
        with st.sidebar.expander("🌐 도메인(URL) 설정", expanded=False):
            i_url = st.text_input("접속 주소", value=active_host_url)
            if i_url.strip() and i_url.strip().rstrip('/') != active_host_url:
                st.session_state['custom_host_url'] = i_url.strip().rstrip('/')
                st.rerun()
        with st.sidebar.expander("👥 계정 권한 관리", expanded=False):
            pend = get_pending_admin_requests()
            for _, r in pend.iterrows():
                st.write(f"• {r['성명']} ({r['아이디']})")
                if st.button("승인", key=f"app_{r['id']}"):
                    approve_admin_request(r['id'])
                    st.rerun()
            appr = get_all_approved_users()
            nm = appr[appr['권한'] != 'master']
            if not nm.empty:
                d_id = st.selectbox("권한 회수", nm['id'].tolist(), format_func=lambda x: nm[nm['id']==x]['성명'].values[0], key="sel_del_user")
                if st.button("권한 삭제", type="primary", key="btn_del_user"):
                    reject_or_delete_user(d_id)
                    st.rerun()

    if menu == "📋 일일 통합 점검":
        with st.sidebar.expander("📥 현장 QR 발급", expanded=False):
            u = f"{active_host_url}?daily_scan=true"
            img = generate_qr_image_with_label("설비실 일일 통합", "현장 점검용 QR", u)
            st.image(img, use_container_width=True)
            st.download_button("QR 저장", img, "QR_Daily.png", "image/png", use_container_width=True)
        with st.sidebar.expander("🗑️ 기록 삭제 관리", expanded=False):
            ldf = get_exact_daily_logs()
            if not ldf.empty:
                s_id = st.selectbox("삭제할 일지", ldf['id'].tolist(), format_func=lambda x: f"[{ldf[ldf['id']==x]['record_date'].values[0]}] {ldf[ldf['id']==x]['inspector'].values[0]}", key="sel_del_daily")
                if st.button("개별 삭제", use_container_width=True, key="btn_del_daily_one"):
                    delete_exact_daily_log(s_id)
                    st.rerun()
                if st.checkbox("모든 기록 삭제 동의", key="chk_del_daily_all"):
                    if st.button("🚨 전체 삭제", use_container_width=True, key="btn_del_daily_all"):
                        delete_all_exact_daily_logs()
                        st.rerun()
            else:
                st.caption("기록 없음")

    elif menu == "🧯 소화기 점검 관리":
        ex = get_all_extinguishers()
        if not ex.empty:
            with st.sidebar.expander("📥 QR 라벨 발급", expanded=False):
                s_ext = st.selectbox("소화기 선택", ex['소화기ID'].tolist(), key="sel_qr_ext")
                loc = ex[ex['소화기ID']==s_ext]['설치위치'].values[0]
                img = generate_qr_image_with_label(f"ID: {s_ext}", f"위치: {loc}", f"{active_host_url}?ext_id={s_ext}")
                st.image(img, use_container_width=True)
                st.download_button("QR 저장", img, f"QR_{s_ext}.png", "image/png", use_container_width=True)
        with st.sidebar.expander("➕ 자산 등록", expanded=False):
            with st.form("ext_reg_form"):
                n_id = st.text_input("식별 ID")
                n_loc = st.text_input("설치 위치")
                if st.form_submit_button("등록", use_container_width=True):
                    if add_extinguisher(n_id, n_loc): 
                        st.rerun()
        with st.sidebar.expander("🗑️ 기록 삭제 관리", expanded=False):
            hdf = get_history_with_id()
            if not hdf.empty:
                s_hid = st.selectbox("삭제할 이력", hdf['id'].tolist(), format_func=lambda x: f"[{hdf[hdf['id']==x]['점검시간'].values[0]}] {hdf[hdf['id']==x]['소화기ID'].values[0]}", key="sel_del_ext")
                if st.button("개별 삭제", use_container_width=True, key="btn_del_ext_one"):
                    delete_inspection_history(s_hid)
                    st.rerun()
                if st.checkbox("모든 이력 삭제 동의", key="chk_del_ext_all"):
                    if st.button("🚨 전체 삭제", use_container_width=True, key="btn_del_ext_all"):
                        delete_all_inspection_history()
                        st.rerun()

    elif menu == "⚡ 유틸리티 사용량 관리":
        with st.sidebar.expander("📥 계량기 QR 발급", expanded=False):
            s_u = st.selectbox("계량기 지정", UTILITY_ORDER, key="sel_qr_util")
            img = generate_qr_image_with_label("유틸리티 검침", s_u, f"{active_host_url}?utility={s_u}")
            st.image(img, use_container_width=True)
            st.download_button("QR 저장", img, f"QR_{s_u}.png", "image/png", use_container_width=True)
        with st.sidebar.expander("🗑️ 검침 데이터 삭제", expanded=False):
            udf = get_utility_df()
            if not udf.empty:
                s_uid = st.selectbox("삭제할 로그", udf['id'].tolist(), format_func=lambda x: f"[{udf[udf['id']==x]['기록일자'].values[0]}] {udf[udf['id']==x]['유틸리티종류'].values[0]}", key="sel_del_util")
                if st.button("개별 삭제", use_container_width=True, key="btn_del_util_one"):
                    delete_utility_usage(s_uid)
                    st.rerun()
                if st.checkbox("모든 데이터 삭제 동의", key="chk_del_util_all"):
                    if st.button("🚨 전체 삭제", use_container_width=True, key="btn_del_util_all"):
                        delete_all_utility_usage()
                        st.rerun()

    elif menu == "🧱 밀폐공간 점검":
        with st.sidebar.expander("📥 현장 QR 발급", expanded=False):
            u = f"{active_host_url}?cf_scan=true"
            img = generate_qr_image_with_label("밀폐공간", "현장 점검용 QR", u)
            st.image(img, use_container_width=True)
            st.download_button("QR 저장", img, "QR_Confined.png", "image/png", use_container_width=True)
        with st.sidebar.expander("🗑️ 점검 기록 삭제", expanded=False):
            cs_df_all = get_confined_space_logs()
            if not cs_df_all.empty:
                s_cs_id = st.selectbox(
                    "삭제할 기록", 
                    cs_df_all['id'].tolist(), 
                    format_func=lambda x: f"[{cs_df_all[cs_df_all['id']==x]['record_date'].values[0]}] {cs_df_all[cs_df_all['id']==x]['inspector'].values[0]}", 
                    key="sel_del_cs"
                )
                if st.button("개별 삭제", use_container_width=True, key="btn_del_cs_one"):
                    delete_confined_space_log(s_cs_id)
                    st.rerun()
                if st.checkbox("모든 기록 삭제 동의", key="chk_del_cs_all"):
                    if st.button("🚨 전체 삭제", use_container_width=True, key="btn_del_cs_all"):
                        delete_all_confined_space_logs()
                        st.rerun()
            else:
                st.caption("기록 없음")

    elif menu == "🧯 소방 펌프 및 밸브 점검":
        with st.sidebar.expander("📥 현장 QR 발급", expanded=False):
            u = f"{active_host_url}?fp_scan=true"
            img = generate_qr_image_with_label("소방펌프실", "현장 점검용 QR", u)
            st.image(img, use_container_width=True)
            st.download_button("QR 저장", img, "QR_FirePump.png", "image/png", use_container_width=True)
        with st.sidebar.expander("🗑️ 점검 기록 삭제", expanded=False):
            fp_df_all = get_fire_pump_logs()
            if not fp_df_all.empty:
                s_fp_id = st.selectbox(
                    "삭제할 기록", 
                    fp_df_all['id'].tolist(), 
                    format_func=lambda x: f"[{fp_df_all[fp_df_all['id']==x]['record_date'].values[0]}] {fp_df_all[fp_df_all['id']==x]['category'].values[0]} ({fp_df_all[fp_df_all['id']==x]['inspector'].values[0]})", 
                    key="sel_del_fp"
                )
                if st.button("개별 삭제", use_container_width=True, key="btn_del_fp_one"):
                    delete_fire_pump_log(s_fp_id)
                    st.rerun()
                if st.checkbox("모든 기록 삭제 동의", key="chk_del_fp_all"):
                    if st.button("🚨 전체 삭제", use_container_width=True, key="btn_del_fp_all"):
                        delete_all_fire_pump_logs()
                        st.rerun()
            else:
                st.caption("기록 없음")

# ----------------------------------------------------
# 6. 메인 화면 본문
# ----------------------------------------------------
if menu == "📊 통합 종합 대시보드":
    st.markdown("""
        <div class='system-banner'>
            <h3>📊 전사 시설물 · 에너지 관리 현황</h3>
            <p>방재 설비 가동 상태 및 6대 에너지원의 소비 실적을 종합 관제합니다.</p>
        </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <style>
    .dash-kpi {
        background: var(--secondary-background-color);
        border: 1px solid var(--border-color, rgba(128,128,128,0.2));
        border-radius: 10px;
        padding: 14px 18px;
        height: 100%;
    }
    .dash-kpi .lbl { font-size: 12px; font-weight: 700; color: var(--text-color); opacity: 0.65; margin-bottom: 6px; }
    .dash-kpi .val { font-size: 26px; font-weight: 800; line-height: 1.15; }
    .dash-kpi .sub { font-size: 12px; margin-top: 6px; opacity: 0.85; color: var(--text-color); }
    .kpi-bar { height: 6px; background: rgba(128,128,128,0.18); border-radius: 3px; margin-top: 10px; overflow: hidden; }
    .kpi-bar-fill { height: 100%; border-radius: 3px; }
    .dash-section-title { font-size: 17px; font-weight: 800; color: var(--text-color); margin: 4px 0 2px 0; }
    .dash-section-desc { font-size: 12px; color: var(--text-color); opacity: 0.65; margin-bottom: 10px; }
    </style>
    """, unsafe_allow_html=True)

    def dash_kpi_card(label, value, sub="", accent="#2563eb", bar_pct=None):
        bar = f'<div class="kpi-bar"><div class="kpi-bar-fill" style="width:{min(max(bar_pct,0),100):.0f}%;background:{accent};"></div></div>' if bar_pct is not None else ""
        sub_html = f'<div class="sub">{sub}</div>' if sub else ""
        return (f'<div class="dash-kpi" style="border-top:3px solid {accent};">'
                f'<div class="lbl">{label}</div><div class="val" style="color:{accent};">{value}</div>{sub_html}{bar}</div>')

    ext_df = get_all_extinguishers()
    u_df = get_utility_df()
    history_df = get_history_with_id()
    unresolved_df = get_unresolved_issues()
    all_anomalies = check_factory_anomalies(u_df)
    
    anomalies = [a for a in all_anomalies if a['id'] not in st.session_state['acknowledged_alerts']]

    cur_m = datetime.now().month
    ins_m = len(history_df[pd.to_datetime(history_df['점검시간'], errors='coerce').dt.month == cur_m]['소화기ID'].dropna().unique()) if not history_df.empty else 0
    c_rate = (ins_m / len(ext_df) * 100) if not ext_df.empty else 0

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(dash_kpi_card("총 등록 소화기", f"{len(ext_df):,} 대", "전사 방재 설비 자산"), unsafe_allow_html=True)
    with k2:
        st.markdown(dash_kpi_card(f"{cur_m}월 소화기 점검률", f"{c_rate:.1f}%", f"{ins_m} / {len(ext_df)} 대 점검 완료", accent="#10b981", bar_pct=c_rate), unsafe_allow_html=True)
    with k3:
        st.markdown(dash_kpi_card("즉시 조치 필요", f"{len(unresolved_df):,} 건", "미조치 설비 결함" if len(unresolved_df) > 0 else "모든 설비 정상", accent="#ef4444" if len(unresolved_df) > 0 else "#10b981"), unsafe_allow_html=True)
    with k4:
        st.markdown(dash_kpi_card("누수/이상 경보", f"{len(anomalies):,} 건", "소비 패턴 이상 감지" if anomalies else "이상 없음", accent="#ef4444" if anomalies else "#10b981"), unsafe_allow_html=True)

    st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)

    if u_df.empty:
        st.info("등록된 검침 데이터가 없습니다. [유틸리티 사용량 관리] 메뉴에서 실제 계량기 검침값을 등록하시면 대시보드 통계가 활성화됩니다.")
    else:
        uf = u_df.copy()
        uf['parsed'] = pd.to_datetime(uf['기록일자'], errors='coerce')
        uf = uf.sort_values(['유틸리티종류', 'parsed', 'id'])
        uf['구간사용량'] = (uf['지침값'] - uf.groupby('유틸리티종류')['지침값'].shift(1)).clip(lower=0)
        uf['year'] = uf['parsed'].dt.year
        uf['month'] = uf['parsed'].dt.month

        st.markdown("<div class='dash-section-title'>기간 사용량 비교</div><div class='dash-section-desc'>선택한 두 날짜의 계량기 지침값 차이로 실제 사용량을 산정합니다.</div>", unsafe_allow_html=True)
        d1c, d2c, _ = st.columns([1, 1, 2])
        with d1c:
            date1 = st.date_input("기준일 1", value=date.today() - timedelta(days=30), key="dash_date1")
        with d2c:
            date2 = st.date_input("기준일 2", value=date.today(), key="dash_date2")

        lo, hi = (date1, date2) if date1 <= date2 else (date2, date1)
        cmp_rows = []
        for u_type in UTILITY_ORDER:
            sub = uf[uf['유틸리티종류'] == u_type]
            m1 = sub[sub['parsed'] <= pd.Timestamp(lo)]
            m2 = sub[sub['parsed'] <= pd.Timestamp(hi)]
            if m1.empty or m2.empty: continue
            v1 = float(m1.iloc[-1]['지침값'])
            v2 = float(m2.iloc[-1]['지침값'])
            cmp_rows.append({
                '유틸리티': u_type, '사용량': round(max(v2 - v1, 0.0), 1),
                '단위': get_utility_unit(u_type), '기준일1 지침': v1, '기준일2 지침': v2
            })
        cmp_df = pd.DataFrame(cmp_rows)

        if cmp_df.empty:
            st.warning("선택한 기간에 검침 기록이 없습니다.")
        else:
            cmp_df['라벨'] = cmp_df.apply(lambda r: f"{r['사용량']:,.1f} {r['단위']}", axis=1)

            water_df = cmp_df[cmp_df['유틸리티'].isin(WATER_UTILITIES)]
            lpg_df = cmp_df[cmp_df['유틸리티'] == 'LPG가스']
            elec_df = cmp_df[cmp_df['유틸리티'] == '전기사용량']

            def create_sub_bar_chart(df_sub, chart_title):
                base = alt.Chart(df_sub).mark_bar(cornerRadiusEnd=4).encode(
                    x=alt.X('사용량:Q', title=None),
                    y=alt.Y('유틸리티:N', sort=UTILITY_ORDER, title=None, axis=alt.Axis(labelFontSize=12)),
                    color=alt.Color('유틸리티:N', scale=alt.Scale(domain=UTILITY_ORDER, range=UTILITY_COLORS), legend=None),
                    tooltip=[alt.Tooltip('유틸리티:N', title='구분'), alt.Tooltip('사용량:Q', title='사용량'), alt.Tooltip('단위:N', title='단위')]
                )
                text = alt.Chart(df_sub).mark_text(align='left', dx=6, fontSize=12, fontWeight='bold', color='#3b82f6').encode(
                    x='사용량:Q', y=alt.Y('유틸리티:N', sort=UTILITY_ORDER), text='라벨:N'
                )
                return (base + text).properties(height=180, title=chart_title)

            c_w, c_l, c_e = st.columns(3)
            with c_w:
                if not water_df.empty:
                    st.altair_chart(create_sub_bar_chart(water_df, f"💧 급수 계통 사용량 ({lo} ~ {hi})"), use_container_width=True)
                else: 
                    st.caption("급수 검침 데이터 없음")
            with c_l:
                if not lpg_df.empty:
                    st.altair_chart(create_sub_bar_chart(lpg_df, f"🔥 LPG 가스 사용량 ({lo} ~ {hi})"), use_container_width=True)
                else: 
                    st.caption("LPG 검침 데이터 없음")
            with c_e:
                if not elec_df.empty:
                    st.altair_chart(create_sub_bar_chart(elec_df, f"⚡ 전기 사용량 ({lo} ~ {hi})"), use_container_width=True)
                else: 
                    st.caption("전기 검침 데이터 없음")

        st.markdown("---")

        st.markdown("<div class='dash-section-title'>월간 사용량 비교</div><div class='dash-section-desc'>당월 누계, 전월, 전년 동월 사용량을 전체 화면 너비로 비교합니다.</div>", unsafe_allow_html=True)
        today = date.today()
        prev_y, prev_m = (today.year, today.month - 1) if today.month > 1 else (today.year - 1, 12)

        def month_usage_sum(yy, mm):
            sel = uf[(uf['year'] == yy) & (uf['month'] == mm)]
            return sel.groupby('유틸리티종류', observed=False)['구간사용량'].sum()

        sums = {
            '당월': month_usage_sum(today.year, today.month),
            '전월': month_usage_sum(prev_y, prev_m),
            '전년동월': month_usage_sum(today.year - 1, today.month)
        }

        sel_u = st.selectbox("기준 유틸리티", UTILITY_ORDER, key="dash_month_util")
        unit = get_utility_unit(sel_u)
        v_cur = float(sums['당월'].get(sel_u, 0.0))
        v_prev = float(sums['전월'].get(sel_u, 0.0))
        v_ly = float(sums['전년동월'].get(sel_u, 0.0))
        d_prev = v_cur - v_prev
        d_pct = (d_prev / v_prev * 100) if v_prev > 0 else None
        cur_sub = "전월 데이터 없음" if d_pct is None else f"전월 대비 {d_pct:+.1f}% ({d_prev:+,.1f} {unit})"

        m1c, m2c, m3c = st.columns(3)
        with m1c:
            st.markdown(dash_kpi_card(f"당월 사용량 ({today.year}.{today.month:02d})", f"{v_cur:,.1f} {unit}", cur_sub, accent="#2563eb"), unsafe_allow_html=True)
        with m2c:
            st.markdown(dash_kpi_card(f"전월 사용량 ({prev_y}.{prev_m:02d})", f"{v_prev:,.1f} {unit}", "직전 1개월 실적", accent="#60a5fa"), unsafe_allow_html=True)
        with m3c:
            st.markdown(dash_kpi_card(f"전년 동월 사용량 ({today.year - 1}.{today.month:02d})", f"{v_ly:,.1f} {unit}", "동월 대비 기준 실적", accent="#94a3b8"), unsafe_allow_html=True)

        st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)

        long_rows = []
        for lbl, s in sums.items():
            for u_type in UTILITY_ORDER:
                long_rows.append({'유틸리티': u_type, '기간': lbl, '사용량': round(float(s.get(u_type, 0.0)), 1)})
        m_df = pd.DataFrame(long_rows)
        period_order = ['당월', '전월', '전년동월']

        grouped_monthly_chart = alt.Chart(m_df).mark_bar(cornerRadiusEnd=3).encode(
            x=alt.X('유틸리티:N', sort=UTILITY_ORDER, title=None, axis=alt.Axis(labelAngle=0)),
            xOffset=alt.X('기간:N', sort=period_order),
            y=alt.Y('사용량:Q', title='월간 사용량'),
            color=alt.Color('기간:N', scale=alt.Scale(domain=period_order, range=['#2563eb', '#60a5fa', '#94a3b8']), legend=alt.Legend(title=None, orient='top')),
            tooltip=[alt.Tooltip('유틸리티:N', title='유틸리티'), alt.Tooltip('기간:N', title='구분'), alt.Tooltip('사용량:Q', title='사용량', format=',.1f')]
        ).properties(height=340, title="전사 6대 유틸리티 월간 사용량 통합 비교 (당월 / 전월 / 전년동월)")
        st.altair_chart(grouped_monthly_chart, use_container_width=True)

    st.markdown("---")

    st.markdown("<div class='dash-section-title'>실시간 이상 모니터링</div><div class='dash-section-desc'>유틸리티 소비 이상 및 미조치 설비 결함을 상시 감시합니다.</div>", unsafe_allow_html=True)
    m_col1, m_col2 = st.columns(2)
    
    with m_col1:
        st.markdown("**유틸리티 이상 감지**")
        if anomalies:
            for a in anomalies:
                with st.container(border=True):
                    if a['level'] == 'error': 
                        st.error(f"{a['title']}\n\n{a['body']}")
                    else: 
                        st.warning(f"{a['title']}\n\n{a['body']}")
                    if st.button("✅ 확인 (조치완료)", key=f"ack_{a['id']}", use_container_width=True):
                        st.session_state['acknowledged_alerts'].add(a['id'])
                        st.rerun()
        else:
            with st.container(border=True):
                st.success("✨ 6대 에너지원 모두 정상 범위 내 가동 중입니다.")
                
    with m_col2:
        st.markdown("**미조치 설비 결함**")
        if not unresolved_df.empty:
            with st.container(border=True):
                st.dataframe(unresolved_df[['점검시간', '소화기ID', '설치위치', '상태', '점검자']], use_container_width=True, height=210)
                st.caption(f"총 {len(unresolved_df)}건의 결함이 조치 대기 중입니다.")
        else:
            with st.container(border=True):
                st.success("✨ 미조치 결함 없음. 모든 설비 정상 가동 중.")

elif menu == "📋 일일 통합 점검":
    st.markdown("<div class='system-banner'><h3>📋 일일 통합 설비 점검 일지</h3><p>설비실 점검 결과를 원본 서식 구조대로 기록하고 보존합니다.</p></div>", unsafe_allow_html=True)
    t1, t2 = st.tabs(["📝 일일 점검표 작성", "📂 점검 기록 조회 (리포트다운로드)"])

    with t1:
        tdy = str(date.today())
        rec = get_daily_inspection_today(tdy)
        allow = True
        
        if rec:
            st.markdown(f"<div class='completed-box'>✅ 오늘({tdy}) 일일 점검 완료 (작성자: {rec[2]})</div>", unsafe_allow_html=True)
            allow = st.checkbox("🔄 다시 작성(수정)하시겠습니까?", key="chk_redaily")
        elif daily_scan_param:
            st.success("📍 [QR 현장 인식 완료] 설비실 일일 점검")
            
        with st.form("exact_form"):
            c1, c2 = st.columns(2)
            with c1: rd = st.date_input("점검 일자", value=datetime.now())
            with c2: rin = st.text_input("점검자 성명", disabled=not allow)
            
            st.markdown("### 1️⃣ Compressor & Air Dryer")
            c_s1, c_s2, c_s3 = st.columns(3)
            with c_s1:
                st.markdown("#### # 1"); st.markdown("---")
                a1_1 = st.text_input("1단 토출온도 (200℃ 이하)", value="", disabled=not allow, key="a1_1")
                a2_1 = st.text_input("2단 토출온도 (200℃ 이하)", value="", disabled=not allow, key="a2_1")
                a3_1 = st.text_input("오일급유온도 (65℃ 이하)", value="", disabled=not allow, key="a3_1")
                a4_1 = st.text_input("Running 누적시간", value="", disabled=not allow, key="a4_1")
                a5_1 = st.text_input("Loading 누적시간", value="", disabled=not allow, key="a5_1")
                a6_1 = st.text_input("1단 토출압력 (2.0~2.5)", value="", disabled=not allow, key="a6_1")
                a7_1 = st.text_input("2단 토출압력 (7.2~8.5)", value="", disabled=not allow, key="a7_1")
                a8_1 = st.text_input("오일급유압력 (2.0 이상)", value="", disabled=not allow, key="a8_1")
                a9_1 = st.text_input("2단 흡입온도 (65℃ 이하)", value="", disabled=not allow, key="a9_1")
                a10_1 = st.text_input("오일급유유면 (중~상)", value="", disabled=not allow, key="a10_1")
                a11_1 = st.text_input("냉각수 입구압력 (3.5 이상)", value="", disabled=not allow, key="a11_1")
                a12_1 = st.text_input("냉각수 입구온도 (32℃이하)", value="", disabled=not allow, key="a12_1")
                a13_1 = st.text_input("냉각수 출구온도 (37℃이하)", value="", disabled=not allow, key="a13_1")
                st.markdown("##### ■ AIR DRYER (#1)")
                ad_l_1 = st.text_input("저압냉매 (3.5~6.0)", value="", disabled=not allow, key="adl_1")
                ad_h_1 = st.text_input("고압냉매 (10~18)", value="", disabled=not allow, key="adh_1")
                ad_a_1 = st.text_input("AIR배관 온도차", value="", disabled=not allow, key="ada_1")
                ad_d_1 = st.text_input("응축수 드레인", value="", disabled=not allow, key="add_1")
            with c_s2:
                st.markdown("#### # 2"); st.markdown("---")
                a1_2 = st.text_input("1단 토출온도 (200℃ 이하)", value="", key="a1_2", disabled=not allow)
                a2_2 = st.text_input("2단 토출온도 (200℃ 이하)", value="", key="a2_2", disabled=not allow)
                a3_2 = st.text_input("오일급유온도 (65℃ 이하)", value="", key="a3_2", disabled=not allow)
                a4_2 = st.text_input("Running 누적시간", value="", key="a4_2", disabled=not allow)
                a5_2 = st.text_input("Loading 누적시간", value="", key="a5_2", disabled=not allow)
                a6_2 = st.text_input("1단 토출압력 (2.0~2.5)", value="", key="a6_2", disabled=not allow)
                a7_2 = st.text_input("2단 토출압력 (7.2~8.5)", value="", key="a7_2", disabled=not allow)
                a8_2 = st.text_input("오일급유압력 (2.0 이상)", value="", key="a8_2", disabled=not allow)
                a9_2 = st.text_input("2단 흡입온도 (65℃ 이하)", value="", key="a9_2", disabled=not allow)
                a10_2 = st.text_input("오일급유유면 (중~상)", value="", key="a10_2", disabled=not allow)
                a11_2 = st.text_input("냉각수 입구압력 (3.5 이상)", value="", key="a11_2", disabled=not allow)
                a12_2 = st.text_input("냉각수 입구온도 (32℃이하)", value="", key="a12_2", disabled=not allow)
                a13_2 = st.text_input("냉각수 출구온도 (37℃이하)", value="", key="a13_2", disabled=not allow)
                st.markdown("##### ■ AIR DRYER (#2)")
                ad_l_2 = st.text_input("저압냉매 (3.5~6.0)", value="", key="adl_2", disabled=not allow)
                ad_h_2 = st.text_input("고압냉매 (10~18)", value="", key="adh_2", disabled=not allow)
                ad_a_2 = st.text_input("AIR배관 온도차", value="", key="ada_2", disabled=not allow)
                ad_d_2 = st.text_input("응축수 드레인", value="", key="add_2", disabled=not allow)
            with c_s3:
                st.markdown("#### # 6"); st.markdown("---")
                a1_3 = st.text_input("1단 토출온도 (200℃ 이하)", value="", key="a1_3", disabled=not allow)
                a2_3 = st.text_input("2단 토출온도 (200℃ 이하)", value="", key="a2_3", disabled=not allow)
                a3_3 = st.text_input("오일급유온도 (65℃ 이하)", value="", key="a3_3", disabled=not allow)
                a4_3 = st.text_input("Running 누적시간", value="", key="a4_3", disabled=not allow)
                a5_3 = st.text_input("Loading 누적시간", value="", key="a5_3", disabled=not allow)
                a6_3 = st.text_input("1단 토출압력 (2.0~2.5)", value="", key="a6_3", disabled=not allow)
                a7_3 = st.text_input("2단 토출압력 (7.2~8.5)", value="", key="a7_3", disabled=not allow)
                a8_3 = st.text_input("오일급유압력 (2.0 이상)", value="", key="a8_3", disabled=not allow)
                a9_3 = st.text_input("2단 흡입온도 (65℃ 이하)", value="", key="a9_3", disabled=not allow)
                a10_3 = st.text_input("오일급유유면 (중~상)", value="", key="a10_3", disabled=not allow)
                a11_3 = st.text_input("냉각수 입구압력 (3.5 이상)", value="", key="a11_3", disabled=not allow)
                a12_3 = st.text_input("냉각수 입구온도 (32℃이하)", value="", key="a12_3", disabled=not allow)
                a13_3 = st.text_input("냉각수 출구온도 (37℃이하)", value="", key="a13_3", disabled=not allow)
                st.markdown("##### ■ AIR DRYER (#6)")
                ad_l_3 = st.text_input("저압냉매 (3.5~6.0)", value="", key="adl_3", disabled=not allow)
                ad_h_3 = st.text_input("고압냉매 (10~18)", value="", key="adh_3", disabled=not allow)
                ad_a_3 = st.text_input("AIR배관 온도차", value="", key="ada_3", disabled=not allow)
                ad_d_3 = st.text_input("응축수 드레인", value="", key="add_3", disabled=not allow)

            st.markdown("---")
            t_c1, t_c2 = st.columns(2)
            with t_c1:
                st.markdown("##### ■ TX-280 TURBO COMP.")
                tx_out_p = st.text_input("토출압력 (0.637~0.735)", value="", disabled=not allow, key="tx_out_p")
                tx_oil_p = st.text_input("윤활유압력 (0.140~0.160)", value="", disabled=not allow, key="tx_oil_p")
                tx_suc_t = st.text_input("2단 흡입온도 (40℃ 이하)", value="", disabled=not allow, key="tx_suc_t")
                tx_oil_t = st.text_input("윤활유온도 (35~50℃)", value="", disabled=not allow, key="tx_oil_t")
                tx_run_t = st.text_input("Running 누적시간", value="", disabled=not allow, key="tx_run_t")
                tx_load_t = st.text_input("Loading 누적시간", value="", disabled=not allow, key="tx_load_t")
                tx_vib = st.text_input("축진동 (10um이하)", value="", disabled=not allow, key="tx_vib")
                tx_gap = st.text_input("GAP전압 (-2.5~+2.5V)", value="", disabled=not allow, key="tx_gap")
                tx_cin_t = st.text_input("냉각수 입구온도 (32℃이하)", value="", disabled=not allow, key="tx_cin_t")
                tx_cout_t = st.text_input("냉각수 출구온도 (37℃이하)", value="", disabled=not allow, key="tx_cout_t")
                tx_head_p = st.text_input("해더압력 (0.65 이상)", value="", disabled=not allow, key="tx_head_p")
            with t_c2:
                st.markdown("##### ■ AIR DRYER (TX)")
                tx_dryer_low = st.text_input("저압냉매 (TX)", value="", disabled=not allow, key="tx_dryer_low")
                tx_dryer_high = st.text_input("고압냉매 (TX)", value="", disabled=not allow, key="tx_dryer_high")
                tx_dryer_air = st.text_input("AIR배관 온도차 (TX)", value="", disabled=not allow, key="tx_dryer_air")
                tx_dryer_drain = st.text_input("응축수 드레인 (TX)", value="", disabled=not allow, key="tx_dryer_drain")

            st.markdown("---")
            p_col1, p_col2 = st.columns(2)
            with p_col1:
                st.markdown("##### ■ 공압설비")
                pneu_leak = st.text_input("누유/누수 여부", value="", disabled=not allow, key="pneu_leak")
                pneu_trap = st.text_input("Drain Trap 정상작동", value="", disabled=not allow, key="pneu_trap")
                pneu_mois = st.text_input("압축공기 내 수분", value="", disabled=not allow, key="pneu_mois")
                pneu_vib = st.text_input("진동이나 이상소음", value="", disabled=not allow, key="pneu_vib")
                pneu_curr = st.text_input("운전전류 정상", value="", disabled=not allow, key="pneu_curr")
            with p_col2:
                st.markdown("##### ■ 냉각탑")
                tower_fan = st.text_input("냉각탑 Fan 정상작동", value="", disabled=not allow, key="tower_fan")
                tower_mach = st.text_input("산기관 정상작동", value="", disabled=not allow, key="tower_mach")
                tower_ext = st.text_input("외관상태 양호", value="", disabled=not allow, key="tower_ext")
                tower_vib = st.text_input("진동소음 여부", value="", disabled=not allow, key="tower_vib")
                tower_fill = st.text_input("충진물 이물질", value="", disabled=not allow, key="tower_fill")

            st.markdown("---")
            w_col1, w_col2 = st.columns(2)
            with w_col1:
                st.markdown("##### ■ 급수설비")
                w_u_lvl = st.text_input("지하수조 수위 (55%이상)", value="", disabled=not allow, key="w_u_lvl")
                w_i_lvl = st.text_input("공업용수 수위 (35%이상)", value="", disabled=not allow, key="w_i_lvl")
                w_p_lvl = st.text_input("정수조 수위 (40%이상)", value="", disabled=not allow, key="w_p_lvl")
                w_u_stat = st.text_input("지하수조 외관 양호", value="", disabled=not allow, key="w_u_stat")
                w_h_stat = st.text_input("고가수조 외관 양호", value="", disabled=not allow, key="w_h_stat")
                w_m_stat = st.text_input("계량기 파손/누수", value="", disabled=not allow, key="w_m_stat")
            with w_col2:
                st.markdown("##### ■ 수조 및 온도")
                oil_sep = st.text_input("유수분리조 누유", value="", disabled=not allow, key="oil_sep")
                cool_cap = st.text_input("냉각수조 용량 정상", value="", disabled=not allow, key="cool_cap")
                cool_t = st.text_input("냉각수 온도 (37℃이하)", value="", disabled=not allow, key="cool_t")
                hot_p = st.text_input("온수Tank 압력 (5.0이하)", value="", disabled=not allow, key="hot_p")
                hot_t = st.text_input("온수Tank 온도 (40℃이상)", value="", disabled=not allow, key="hot_t")
                hot_stat = st.text_input("온수Tank 누수여부", value="", disabled=not allow, key="hot_stat")
                temp_mach = st.text_input("기계실 실내온도 (℃)", value="", disabled=not allow, key="temp_mach")
                temp_outside = st.text_input("외부 현재온도 (℃)", value="", disabled=not allow, key="temp_outside")

            st.markdown("---")
            pmp_c1, pmp_c2, pmp_c3 = st.columns(3)
            with pmp_c1:
                st.markdown("##### ■ COMP 냉각수 PUMP")
                p_comp_p = st.text_input("압력 (3.5 이상)", value="", key="pc_p", disabled=not allow)
                p_comp_n = st.text_input("소음진동 여부", value="", key="pc_n", disabled=not allow)
                p_comp_s = st.text_input("외관상태 양호", value="", key="pc_s", disabled=not allow)
            with pmp_c2:
                st.markdown("##### ■ 공장동 냉난방 PUMP")
                p_fac_p = st.text_input("압력 (3.5 이상)", value="", key="pf_p", disabled=not allow)
                p_fac_n = st.text_input("소음진동 여부", value="", key="pf_n", disabled=not allow)
                p_fac_s = st.text_input("외관상태 양호", value="", key="pf_s", disabled=not allow)
            with pmp_c3:
                st.markdown("##### ■ 사무동 냉난방 PUMP")
                p_off_p = st.text_input("압력 (3.5 이상)", value="", key="po_p", disabled=not allow)
                p_off_n = st.text_input("소음진동 여부", value="", key="po_n", disabled=not allow)
                p_off_s = st.text_input("외관상태 양호", value="", key="po_s", disabled=not allow)

            etc_notes = st.text_area("비고 및 특이사항", height=80, key="etc_notes_txt", disabled=not allow)
            if st.form_submit_button("💾 일일 점검표 저장", use_container_width=True, disabled=not allow):
                if not rin.strip(): 
                    st.error("성명을 입력하세요.")
                else:
                    save_exact_daily_log({
                        'record_date': str(rd), 'inspector': rin.strip(),
                        'al150_t1_1': a1_1, 'al150_t1_2': a1_2, 'al150_t1_3': a1_3,
                        'al150_t2_1': a2_1, 'al150_t2_2': a2_2, 'al150_t2_3': a2_3,
                        'oil_temp_1': a3_1, 'oil_temp_2': a3_2, 'oil_temp_3': a3_3,
                        'run_time_1': a4_1, 'run_time_2': a4_2, 'run_time_3': a4_3,
                        'load_time_1': a5_1, 'load_time_2': a5_2, 'load_time_3': a5_3,
                        'p1_1': a6_1, 'p1_2': a6_2, 'p1_3': a6_3,
                        'p2_1': a7_1, 'p2_2': a7_2, 'p2_3': a7_3,
                        'oil_p_1': a8_1, 'oil_p_2': a8_2, 'oil_p_3': a8_3,
                        'suc_t_1': a9_1, 'suc_t_2': a9_2, 'suc_t_3': a9_3,
                        'oil_lvl_1': a10_1, 'oil_lvl_2': a10_2, 'oil_lvl_3': a10_3,
                        'cool_p_1': a11_1, 'cool_p_2': a11_2, 'cool_p_3': a11_3,
                        'cool_in_1': a12_1, 'cool_in_2': a12_2, 'cool_in_3': a12_3,
                        'cool_out_1': a13_1, 'cool_out_2': a13_2, 'cool_out_3': a13_3,
                        'dryer_low_1': ad_l_1, 'dryer_low_2': ad_l_2, 'dryer_low_3': ad_l_3,
                        'dryer_high_1': ad_h_1, 'dryer_high_2': ad_h_2, 'dryer_high_3': ad_h_3,
                        'dryer_air_1': ad_a_1, 'dryer_air_2': ad_a_2, 'dryer_air_3': ad_a_3,
                        'dryer_drain_1': ad_d_1, 'dryer_drain_2': ad_d_2, 'dryer_drain_3': ad_d_3,
                        'tx_out_p': tx_out_p, 'tx_oil_p': tx_oil_p, 'tx_suc_t': tx_suc_t, 'tx_oil_t': tx_oil_t,
                        'tx_run_t': tx_run_t, 'tx_load_t': tx_load_t, 'tx_vib': tx_vib, 'tx_gap': tx_gap,
                        'tx_cin_t': tx_cin_t, 'tx_cout_t': tx_cout_t, 'tx_head_p': tx_head_p,
                        'tx_dryer_low': tx_dryer_low, 'tx_dryer_high': tx_dryer_high, 'tx_dryer_air': tx_dryer_air, 'tx_dryer_drain': tx_dryer_drain,
                        'pneu_leak': pneu_leak, 'pneu_trap': pneu_trap, 'pneu_mois': pneu_mois, 'pneu_vib': pneu_vib, 'pneu_curr': pneu_curr,
                        'tower_fan': tower_fan, 'tower_mach': tower_mach, 'tower_ext': tower_ext, 'tower_vib': tower_vib, 'tower_fill': tower_fill,
                        'w_u_lvl': w_u_lvl, 'w_i_lvl': w_i_lvl, 'w_p_lvl': w_p_lvl,
                        'w_u_stat': w_u_stat, 'w_h_stat': w_h_stat, 'w_m_stat': w_m_stat,
                        'oil_sep': oil_sep, 'cool_cap': cool_cap, 'cool_t': cool_t,
                        'hot_p': hot_p, 'hot_t': hot_t, 'hot_stat': hot_stat,
                        'temp_mach': temp_mach, 'temp_outside': temp_outside,
                        'p_comp_p': p_comp_p, 'p_comp_n': p_comp_n, 'p_comp_s': p_comp_s,
                        'p_fac_p': p_fac_p, 'p_fac_n': p_fac_n, 'p_fac_s': p_fac_s,
                        'p_off_p': p_off_p, 'p_off_n': p_off_n, 'p_off_s': p_off_s,
                        'etc_notes': etc_notes.strip()
                    })
                    st.rerun()

    with t2:
        st.subheader("📂 일일 점검 기록 리포트 조회 및 다운로드")
        ex_df = get_exact_daily_logs()
        if not ex_df.empty:
            sel_d_obj = st.date_input("📅 조회 및 다운로드할 점검 일자 선택", value=datetime.now())
            sel_d = str(sel_d_obj)
            
            filtered_df = ex_df[ex_df['record_date'] == sel_d]
            if not filtered_df.empty:
                t_row = filtered_df.iloc[0]
                rep_map = [
                    ["AL-150 SCREW COMP", "1단 토출온도 (200℃ 이하)", t_row.get('al150_t1_1',''), t_row.get('al150_t1_2',''), t_row.get('al150_t1_3',''), f"토출압력: {t_row.get('tx_out_p','')}"],
                    ["AL-150 SCREW COMP", "2단 토출온도 (200℃ 이하)", t_row.get('al150_t2_1',''), t_row.get('al150_t2_2',''), t_row.get('al150_t2_3',''), f"윤활유압력: {t_row.get('tx_oil_p','')}"],
                    ["AL-150 SCREW COMP", "오일급유온도 (65℃ 이하)", t_row.get('oil_temp_1',''), t_row.get('oil_temp_2',''), t_row.get('oil_temp_3',''), f"2단흡입온도: {t_row.get('tx_suc_t','')}"],
                    ["AL-150 SCREW COMP", "Running 누적시간", t_row.get('run_time_1',''), t_row.get('run_time_2',''), t_row.get('run_time_3',''), f"윤활유온도: {t_row.get('tx_oil_t','')}"],
                    ["AL-150 SCREW COMP", "Loading 누적시간", t_row.get('load_time_1',''), t_row.get('load_time_2',''), t_row.get('load_time_3',''), f"Run시간: {t_row.get('tx_run_t','')}"],
                    ["AL-150 SCREW COMP", "1단 토출압력 (2.0~2.5)", t_row.get('p1_1',''), t_row.get('p1_2',''), t_row.get('p1_3',''), f"Load시간: {t_row.get('tx_load_t','')}"],
                    ["AL-150 SCREW COMP", "2단 토출압력 (7.2~8.5)", t_row.get('p2_1',''), t_row.get('p2_2',''), t_row.get('p2_3',''), f"축진동: {t_row.get('tx_vib','')}"],
                    ["AL-150 SCREW COMP", "오일급유압력 (2.0 이상)", t_row.get('oil_p_1',''), t_row.get('oil_p_2',''), t_row.get('oil_p_3',''), f"GAP전압: {t_row.get('tx_gap','')}"],
                    ["AL-150 SCREW COMP", "2단 흡입온도 (65℃ 이하)", t_row.get('suc_t_1',''), t_row.get('suc_t_2',''), t_row.get('suc_t_3',''), f"냉각수 입구온도: {t_row.get('tx_cin_t','')}"],
                    ["AL-150 SCREW COMP", "오일급유유면 (중~상)", t_row.get('oil_lvl_1',''), t_row.get('oil_lvl_2',''), t_row.get('oil_lvl_3',''), f"냉각수 출구온도: {t_row.get('tx_cout_t','')}"],
                    ["AL-150 SCREW COMP", "냉각수 입구압력 (3.5 이상)", t_row.get('cool_p_1',''), t_row.get('cool_p_2',''), t_row.get('cool_p_3',''), f"해더압력: {t_row.get('tx_head_p','')}"],
                    ["AL-150 SCREW COMP", "냉각수 입구온도 (32℃이하)", t_row.get('cool_in_1',''), t_row.get('cool_in_2',''), t_row.get('cool_in_3',''), "-"],
                    ["AL-150 SCREW COMP", "냉각수 출구온도 (37℃이하)", t_row.get('cool_out_1',''), t_row.get('cool_out_2',''), t_row.get('cool_out_3',''), "-"],
                    ["AIR DRYER", "저압냉매 (3.5~6.0)", t_row.get('dryer_low_1',''), t_row.get('dryer_low_2',''), t_row.get('dryer_low_3',''), f"저압냉매: {t_row.get('tx_dryer_low','')}"],
                    ["AIR DRYER", "고압냉매 (10~18)", t_row.get('dryer_high_1',''), t_row.get('dryer_high_2',''), t_row.get('dryer_high_3',''), f"고압냉매: {t_row.get('tx_dryer_high','')}"],
                    ["AIR DRYER", "AIR배관 온도차", t_row.get('dryer_air_1',''), t_row.get('dryer_air_2',''), t_row.get('dryer_air_3',''), f"온도차: {t_row.get('tx_dryer_air','')}"],
                    ["AIR DRYER", "응축수 드레인", t_row.get('dryer_drain_1',''), t_row.get('dryer_drain_2',''), t_row.get('dryer_drain_3',''), f"드레인: {t_row.get('tx_dryer_drain','')}"],
                    ["공압설비", "누유 및 누수 여부", t_row.get('pneu_leak',''), "-", "-", "-"],
                    ["공압설비", "Drain Trap 정상작동", t_row.get('pneu_trap',''), "-", "-", "-"],
                    ["공압설비", "압축공기 내 수분 포함", t_row.get('pneu_mois',''), "-", "-", "-"],
                    ["공압설비", "진동이나 이상소음", t_row.get('pneu_vib',''), "-", "-", "-"],
                    ["공압설비", "운전전류 정상여부", t_row.get('pneu_curr',''), "-", "-", "-"],
                    ["냉각탑", "냉각탑 Fan 정상작동", t_row.get('tower_fan',''), "-", "-", "-"],
                    ["냉각탑", "산기관 정상작동", t_row.get('tower_mach',''), "-", "-", "-"],
                    ["냉각탑", "외관상태 양호", t_row.get('tower_ext',''), "-", "-", "-"],
                    ["냉각탑", "진동 및 이상소음", t_row.get('tower_vib',''), "-", "-", "-"],
                    ["냉각탑", "충진물루바 이물질", t_row.get('tower_fill',''), "-", "-", "-"],
                    ["급수설비", "지하수조 수위 (55%이상)", t_row.get('w_u_lvl',''), "-", "-", "-"],
                    ["급수설비", "공업용수 수위 (35%이상)", t_row.get('w_i_lvl',''), "-", "-", "-"],
                    ["급수설비", "정수조 수위 (40%이상)", t_row.get('w_p_lvl',''), "-", "-", "-"],
                    ["급수설비", "지하수조 외관 양호", t_row.get('w_u_stat',''), "-", "-", "-"],
                    ["급수설비", "고가수조 외관 양호", t_row.get('w_h_stat',''), "-", "-", "-"],
                    ["급수설비", "계량기 파손/누수", t_row.get('w_m_stat',''), "-", "-", "-"],
                    ["수조 및 온도", "유수분리조 누유", t_row.get('oil_sep',''), "-", "-", "-"],
                    ["수조 및 온도", "냉각수조 용량 정상", t_row.get('cool_cap',''), "-", "-", "-"],
                    ["수조 및 온도", "냉각수-2 온도 (37℃이하)", t_row.get('cool_t',''), "-", "-", "-"],
                    ["수조 및 온도", "온수Tank 압력 (5.0이하)", t_row.get('hot_p',''), "-", "-", "-"],
                    ["수조 및 온도", "온수Tank 온도 (40℃이상)", t_row.get('hot_t',''), "-", "-", "-"],
                    ["수조 및 온도", "파손 및 누수 여부", t_row.get('hot_stat',''), "-", "-", "-"],
                    ["수조 및 온도", "기계실 실내온도 (℃)", t_row.get('temp_mach',''), "-", "-", "-"],
                    ["수조 및 온도", "외부 현재온도 (℃)", t_row.get('temp_outside',''), "-", "-", "-"],
                    ["PUMP", "COMP 냉각수 PUMP 압력", t_row.get('p_comp_p',''), "-", "-", "-"],
                    ["PUMP", "COMP 냉각수 PUMP 소음진동", t_row.get('p_comp_n',''), "-", "-", "-"],
                    ["PUMP", "COMP 냉각수 PUMP 외관", t_row.get('p_comp_s',''), "-", "-", "-"],
                    ["PUMP", "공장동 냉난방 PUMP 압력", t_row.get('p_fac_p',''), "-", "-", "-"],
                    ["PUMP", "공장동 냉난방 PUMP 소음진동", t_row.get('p_fac_n',''), "-", "-", "-"],
                    ["PUMP", "공장동 냉난방 PUMP 외관", t_row.get('p_fac_s',''), "-", "-", "-"],
                    ["PUMP", "사무동 냉난방 PUMP 압력", t_row.get('p_off_p',''), "-", "-", "-"],
                    ["PUMP", "사무동 냉난방 PUMP 소음진동", t_row.get('p_off_n',''), "-", "-", "-"],
                    ["PUMP", "사무동 냉난방 PUMP 외관", t_row.get('p_off_s',''), "-", "-", "-"],
                    ["기타", "특이사항 및 비고", t_row.get('etc_notes',''), "-", "-", "-"]
                ]
                r_df = pd.DataFrame(rep_map, columns=["대분류", "점검 항목", "AL-150 #(1)", "AL-150 #(2)", "AL-150 #(6)", "TX-280 및 부대설비"])
                st.markdown(f"📌 [일자: {sel_d} / 점검자: {t_row.get('inspector','')}]")
                st.dataframe(r_df, use_container_width=True, height=450)
                
                st.download_button(
                    f"📥 [{sel_d}] 원본 서식 엑셀(CSV) 양식 다운로드",
                    r_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig'),
                    f"Daily_Inspection_{sel_d}.csv",
                    "text/csv",
                    use_container_width=True
                )
            else:
                st.warning(f"선택하신 날짜({sel_d})에는 점검 기록이 없습니다.")
            with st.expander("원본 가로형 데이터베이스 로그 전체 보기"):
                st.dataframe(ex_df, use_container_width=True)
        else: 
            st.info("기록 없음")

elif menu == "🧯 소화기 점검 관리":
    is_qr_scanned = bool(ext_id_param)
    ext_all = get_all_extinguishers()
    current_ext = ext_id_param if is_qr_scanned else (ext_all['소화기ID'].iloc[0] if not ext_all.empty else "")
    current_loc = ext_all[ext_all['소화기ID']==current_ext]['설치위치'].values[0] if not ext_all.empty and current_ext in ext_all['소화기ID'].values else ""

    st.markdown("<div class='system-banner'><h3>🧯 소화 설비 현장 점검 모듈</h3><p>기기를 점검하고 결함을 추적·조치합니다.</p></div>", unsafe_allow_html=True)
    t1, t2 = st.tabs(["📋 점검표 작성", "📈 현황 및 조치"])
    
    with t1:
        rec = get_recent_inspection_this_month(current_ext) if current_ext else None
        allow = True
        
        if rec:
            st.markdown(f"<div class='completed-box'>✅ 이번 달 정기 점검 완료 (점검자: {rec[1]})</div>", unsafe_allow_html=True)
            allow = st.checkbox("🔄 재점검(수정) 진행")
        elif is_qr_scanned:
            st.success(f"📍 [QR 인식 완료] {current_ext} ({current_loc})")
            
        with st.form("ext_form"):
            c1, c2 = st.columns(2)
            with c1:
                t_ext = st.text_input("ID", value=current_ext, disabled=True) if is_qr_scanned else st.selectbox("소화기 선택", ext_all['소화기ID'].tolist() if not ext_all.empty else [""])
                ins = st.text_input("점검자 성명", disabled=not allow)
            with c2:
                stat = st.selectbox("외관/압력", ["정상", "압력 부족", "파손/불량", "위치 이탈/미비치"], disabled=not allow)
                nts = st.text_area("특이사항", height=105, disabled=not allow)
            if st.form_submit_button("✅ 제출", use_container_width=True, disabled=not allow):
                if not ins.strip(): 
                    st.error("성명 입력")
                elif not (current_ext if is_qr_scanned else t_ext):
                    st.error("선택된 소화기가 없습니다.")
                else:
                    save_inspection(current_ext if is_qr_scanned else t_ext, ins.strip(), stat, nts.strip())
                    st.rerun()

    with t2:
        unr = get_unresolved_issues()
        if not unr.empty:
            st.subheader("🚨 결함 기기 (대기 목록)")
            for _, r in unr.iterrows():
                with st.expander(f"⚠️ [{r['소화기ID']}] {r['상태']}", expanded=True):
                    st.write(f"• 점검자: {r['점검자']} / 일시: {r['점검시간']} / 내용: {r['특이사항']}")
                    if is_admin and st.button("조치 완료", key=f"re_{r['id']}"): 
                        mark_as_resolved(r['id'])
                        st.rerun()
        st.subheader("📑 점검 이력")
        hdf = get_history_with_id()
        if not hdf.empty: 
            st.dataframe(hdf[['점검시간', '소화기ID', '설치위치', '점검자', '상태', '특이사항', '조치상태']], use_container_width=True)

elif menu == "⚡ 유틸리티 사용량 관리":
    is_u_qr = bool(utility_param)
    cur_u = utility_param if is_u_qr else UTILITY_ORDER[0]
    st.markdown("<div class='system-banner'><h3>⚡ 유틸리티 에너지 검침</h3><p>누수 감지 및 소비량 실시간 모니터링</p></div>", unsafe_allow_html=True)
    tdy = str(date.today())
    rec = get_utility_record_today(cur_u, tdy)
    allow = True
    
    if rec:
        st.markdown(f"<div class='completed-box'>✅ 오늘({tdy}) [{cur_u}] 검침 완료 (지침: {rec[3]:,.0f})</div>", unsafe_allow_html=True)
        allow = st.checkbox("🔄 다시 수정등록")
    elif is_u_qr:
        st.success(f"📍 [QR 인식 완료] {cur_u}")

    with st.form("u_form"):
        c1, c2 = st.columns(2)
        with c1:
            rd = st.date_input("검침 일자", value=datetime.now())
            u_t = st.selectbox("구분", UTILITY_ORDER, index=UTILITY_ORDER.index(cur_u) if cur_u in UTILITY_ORDER else 0, disabled=is_u_qr)
        with c2:
            con = st.number_input("지침값", min_value=0.0, step=1.0, value=0.0, disabled=not allow)
            rn = st.text_input("성명", disabled=not allow)
        nts = st.text_area("특이사항", height=75, disabled=not allow)
        if st.form_submit_button("💾 등록", use_container_width=True, disabled=not allow):
            if not rn.strip(): 
                st.error("성명 입력")
            else:
                save_utility_usage(str(rd), cur_u if is_u_qr else u_t, con, rn.strip(), nts)
                st.rerun()
    st.markdown("---")
    st.subheader("📂 이력 조회")
    udf = get_utility_df()
    if not udf.empty: 
        st.dataframe(udf, use_container_width=True)
    else: 
        st.info("데이터 없음")

elif menu == "🧱 밀폐공간 점검":
    st.markdown("<div class='system-banner'><h3>🧱 밀폐공간 안전 작업 허가 및 점검</h3><p>밀폐공간 출입 전 환기, 보호구, 측정기기 상태를 점검합니다.</p></div>", unsafe_allow_html=True)
    
    tab_cs1, tab_cs2 = st.tabs(["📝 점검표 작성", "📂 점검 기록 조회 (리포트다운로드)"])

    with tab_cs1:
        if cf_scan_param:
            st.success("📍 [QR 현장 인식 완료] 밀폐공간 통합 점검 모드입니다.")

        today_str = str(date.today())
        today_cs_record = get_confined_space_today(today_str)
        allow_cs = True

        if today_cs_record:
            st.markdown(f"<div class='completed-box'>✅ 오늘({today_str}) 밀폐공간 점검이 이미 완료되었습니다. (작성자: {today_cs_record[2]})</div>", unsafe_allow_html=True)
            allow_cs = st.checkbox("🔄 오늘 점검을 다시 작성(수정)하시겠습니까?", key="chk_re_cs")

        with st.form("cs_form"):
            st.markdown("##### 📌 기본 정보")
            c1, c2 = st.columns(2)
            with c1: cs_date = st.date_input("점검 일자", value=datetime.now())
            with c2: cs_ins = st.text_input("점검자 성명", disabled=not allow_cs)
            
            st.markdown("---")
            st.markdown("### 1️⃣ 출입통제 (시건장치 여부)")
            c_c1, c_c2 = st.columns(2)
            with c_c1:
                q1 = st.radio("1-1. 지하피트", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q2 = st.radio("1-2. 냉각수조", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q3 = st.radio("1-3. 고가수조", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q4 = st.radio("1-4. 지하수조", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q5 = st.radio("1-5. 정화조 1", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q6 = st.radio("1-6. 정화조 2", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q7 = st.radio("1-7. 정화조 3", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q8 = st.radio("1-8. 침전조", ["O", "X"], disabled=not allow_cs, horizontal=True)
            with c_c2:
                q9 = st.radio("1-9. 집수조", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q10 = st.radio("1-10. 유수분리조", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q11 = st.radio("1-11. 지하펌프실", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q12 = st.radio("1-12. 기계실 내부피트", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q13 = st.radio("1-13. 엔진룸 챔버", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q14 = st.radio("1-14. 정화조 4 (도로용 덮개 폐쇄 여부)", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q15 = st.radio("1-15. 정화조 5 (도로용 덮개 폐쇄 여부)", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q16 = st.radio("2. 각 위치별 위험표지판 부착 여부", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q17 = st.radio("3. 비상연락망 기재 여부", ["O", "X"], disabled=not allow_cs, horizontal=True)

            st.markdown("---")
            st.markdown("### 2️⃣ 환기장치 & 3️⃣ 측정장치 및 보호구")
            v_c1, v_c2 = st.columns(2)
            with v_c1:
                st.markdown("##### ■ 환기장치")
                q18 = st.radio("4. 정상작동 여부 (팬 작동/송풍관 누출)", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q19 = st.radio("5. 환기장치 가동상태 식별표시", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q20 = st.radio("6. 전선 피복 손상 여부", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q21 = st.radio("7. 팬 먼지 쌓임 여부", ["O", "X"], disabled=not allow_cs, horizontal=True)
                st.markdown("##### ■ 측정장치")
                q22 = st.radio("9. 측정기 경고장치 필증 (1년 이내)", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q23 = st.radio("10. 배터리 충분 여부", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q24 = st.radio("11. 산소 측정치 (18~23.5% 범위 내)", ["O", "X"], disabled=not allow_cs, horizontal=True)
            with v_c2:
                st.markdown("##### ■ 무전기 & 안전보호구")
                q25 = st.radio("8. 안전보호구 정해진 위치 보관", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q26 = st.radio("12. 무전기 배터리 충전 여부", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q27 = st.radio("13. 무전기 주파수 동일 설정", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q28 = st.radio("14. 송수신 가능 여부", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q29 = st.radio("15. 무전기 음량 충분", ["O", "X"], disabled=not allow_cs, horizontal=True)

            st.markdown("---")
            st.markdown("### 4️⃣ 들것, 삼각대, 공기호흡기")
            e_c1, e_c2 = st.columns(2)
            with e_c1:
                st.markdown("##### ■ 들것 및 삼각대")
                q30 = st.radio("16. 들것 파손/끈 끊어짐 여부", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q31 = st.radio("17. 삼각대 파손 부위 (설치 가능 여부)", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q32 = st.radio("18. 삼각대 윈치 정상 작동", ["O", "X"], disabled=not allow_cs, horizontal=True)
            with e_c2:
                st.markdown("##### ■ 공기호흡기")
                q33 = st.radio("19. 용기 내 산소용량 적정 (압력/용량)", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q34 = st.radio("20. 안면부 새는 곳 (보조 마스크)", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q35 = st.radio("21. 마스크 밸브 정상작동", ["O", "X"], disabled=not allow_cs, horizontal=True)
                q36 = st.radio("22. 연결호스 파손 여부", ["O", "X"], disabled=not allow_cs, horizontal=True)

            cs_notes = st.text_area("비고 및 특이사항", height=80, disabled=not allow_cs)

            if st.form_submit_button("💾 밀폐공간 점검표 저장", use_container_width=True, disabled=not allow_cs):
                if not cs_ins.strip():
                    st.error("점검자 성명을 입력해주세요.")
                else:
                    q_data = {
                        "1-1": q1, "1-2": q2, "1-3": q3, "1-4": q4, "1-5": q5, "1-6": q6, "1-7": q7, "1-8": q8,
                        "1-9": q9, "1-10": q10, "1-11": q11, "1-12": q12, "1-13": q13, "1-14": q14, "1-15": q15,
                        "2": q16, "3": q17, "4": q18, "5": q19, "6": q20, "7": q21, "9": q22, "10": q23, "11": q24,
                        "8": q25, "12": q26, "13": q27, "14": q28, "15": q29, "16": q30, "17": q31, "18": q32,
                        "19": q33, "20": q34, "21": q35, "22": q36
                    }
                    save_confined_space_log(cs_date, cs_ins, q_data, cs_notes)
                    st.rerun()

    with tab_cs2:
        st.subheader("📂 밀폐공간 점검 기록 리포트 조회 및 다운로드")
        cs_df = get_confined_space_logs()
        if not cs_df.empty:
            sel_d_obj = st.date_input("📅 조회 및 다운로드할 점검 일자 선택", value=datetime.now(), key="cs_date_input")
            sel_d = str(sel_d_obj)
            
            filtered_df = cs_df[cs_df['record_date'] == sel_d]
            if not filtered_df.empty:
                t_row = filtered_df.iloc[0]
                q_dict = json.loads(t_row['q_data'])
                
                rep_map = [
                    ["출입통제", "1-1. 지하피트", q_dict.get("1-1","")],
                    ["출입통제", "1-2. 냉각수조", q_dict.get("1-2","")],
                    ["출입통제", "1-3. 고가수조", q_dict.get("1-3","")],
                    ["출입통제", "1-4. 지하수조", q_dict.get("1-4","")],
                    ["출입통제", "1-5. 정화조 1", q_dict.get("1-5","")],
                    ["출입통제", "1-6. 정화조 2", q_dict.get("1-6","")],
                    ["출입통제", "1-7. 정화조 3", q_dict.get("1-7","")],
                    ["출입통제", "1-8. 침전조", q_dict.get("1-8","")],
                    ["출입통제", "1-9. 집수조", q_dict.get("1-9","")],
                    ["출입통제", "1-10. 유수분리조", q_dict.get("1-10","")],
                    ["출입통제", "1-11. 지하펌프실", q_dict.get("1-11","")],
                    ["출입통제", "1-12. 기계실 내부피트", q_dict.get("1-12","")],
                    ["출입통제", "1-13. 엔진룸 챔버", q_dict.get("1-13","")],
                    ["출입통제", "1-14. 정화조 4", q_dict.get("1-14","")],
                    ["출입통제", "1-15. 정화조 5", q_dict.get("1-15","")],
                    ["표지판연락망", "2. 위험표지판 부착", q_dict.get("2","")],
                    ["표지판연락망", "3. 비상연락망 기재", q_dict.get("3","")],
                    ["환기장치", "4. 정상작동 여부", q_dict.get("4","")],
                    ["환기장치", "5. 가동 식별표시", q_dict.get("5","")],
                    ["환기장치", "6. 전선 피복 상태", q_dict.get("6","")],
                    ["환기장치", "7. 팬 먼지 상태", q_dict.get("7","")],
                    ["안전보호구", "8. 정위치 보관 여부", q_dict.get("8","")],
                    ["측정장치", "9. 경고장치 필증", q_dict.get("9","")],
                    ["측정장치", "10. 배터리 충분 여부", q_dict.get("10","")],
                    ["측정장치", "11. 산소 측정치 범위", q_dict.get("11","")],
                    ["무전기", "12. 배터리 충전", q_dict.get("12","")],
                    ["무전기", "13. 주파수 일치", q_dict.get("13","")],
                    ["무전기", "14. 송수신 여부", q_dict.get("14","")],
                    ["무전기", "15. 음량 충분", q_dict.get("15","")],
                    ["들것", "16. 파손/끈 상태", q_dict.get("16","")],
                    ["삼각대", "17. 파손 부위 여부", q_dict.get("17","")],
                    ["삼각대", "18. 윈치 정상 작동", q_dict.get("18","")],
                    ["공기호흡기", "19. 산소용량 적정", q_dict.get("19","")],
                    ["공기호흡기", "20. 안면부 보조마스크", q_dict.get("20","")],
                    ["공기호흡기", "21. 밸브 정상작동", q_dict.get("21","")],
                    ["공기호흡기", "22. 연결호스 파손", q_dict.get("22","")],
                    ["기타", "비고", t_row.get("notes","")]
                ]
                r_df = pd.DataFrame(rep_map, columns=["구분", "점검항목", "평가(O/X)"])
                st.markdown(f"📌 [일자: {sel_d} / 점검자: {t_row.get('inspector','')}]")
                st.dataframe(r_df, use_container_width=True, height=450)
                st.download_button(
                    f"📥 [{sel_d}] 밀폐공간 점검표(CSV) 다운로드",
                    r_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig'),
                    f"ConfinedSpace_Inspection_{sel_d}.csv",
                    "text/csv",
                    use_container_width=True
                )
            else:
                st.warning(f"선택하신 날짜({sel_d})에는 점검 기록이 없습니다.")
            with st.expander("원본 데이터베이스 전체 로그 조회"):
                st.dataframe(cs_df, use_container_width=True)
        else:
            st.info("등록된 기록이 없습니다.")

elif menu == "🧯 소방 펌프 및 밸브 점검":
    st.markdown("<div class='system-banner'><h3>🧯 소방 펌프실 및 알람밸브 개폐 점검</h3><p>엔진펌프, 전기펌프, 스프링클러 밸브 등 핵심 소방설비를 점검합니다.</p></div>", unsafe_allow_html=True)
    
    tab_fp1, tab_fp2 = st.tabs(["📝 점검표 작성", "📂 점검 기록 조회 (리포트다운로드)"])

    with tab_fp1:
        if fp_scan_param:
            st.success("📍 [QR 현장 인식 완료] 소방펌프/밸브 점검 모드입니다.")

        pump_category = st.selectbox("📌 점검할 설비를 선택하세요", ["1. 소방 엔진펌프 점검", "2. 소방 전기펌프 점검", "3. 스프링클러 밸브점검"])
        
        today_str = str(date.today())
        today_fp_record = get_fire_pump_today(pump_category, today_str)
        allow_fp = True

        if today_fp_record:
            st.markdown(f"<div class='completed-box'>✅ 오늘({today_str}) 해당 설비({pump_category}) 점검이 이미 완료되었습니다. (작성자: {today_fp_record[2]})</div>", unsafe_allow_html=True)
            allow_fp = st.checkbox("🔄 오늘 점검을 다시 작성(수정)하시겠습니까?", key=f"chk_re_{pump_category}")

        with st.form(f"fp_form_{pump_category}"):
            c1, c2 = st.columns(2)
            with c1: fp_date = st.date_input("점검 일자", value=datetime.now())
            with c2: fp_ins = st.text_input("점검자 성명", disabled=not allow_fp)
            st.markdown("---")

            q_data = {}
            if pump_category == "1. 소방 엔진펌프 점검":
                st.markdown("### 🔥 소방 엔진펌프 점검")
                fc1, fc2 = st.columns(2)
                with fc1:
                    q_data['q1'] = st.radio("2. 제어반 표시등 점등 여부", ["Yes", "No"], disabled=not allow_fp, horizontal=True)
                    q_data['q2'] = st.radio("3. 연료탱크 경유 적정량", ["Full", "3/4 이상", "1/2 이상", "1/2 이하"], disabled=not allow_fp, horizontal=True)
                    q_data['q3'] = st.text_input("4. 배터리 #1 충전상태 (Unit: V)", disabled=not allow_fp)
                    q_data['q4'] = st.text_input("5. 배터리 #2 충전상태 (Unit: V)", disabled=not allow_fp)
                    q_data['q5'] = st.radio("6. 엔진오일 상태", ["Yes", "No"], disabled=not allow_fp, horizontal=True)
                    q_data['q6'] = st.radio("7. 배터리 충전기 작동여부", ["Yes", "No"], disabled=not allow_fp, horizontal=True)
                    q_data['q7'] = st.radio("8. 연료/유류 공급배관 파손여부", ["No", "Yes"], disabled=not allow_fp, horizontal=True)
                    q_data['q8'] = st.radio("9. 윤활유 유량 (정상확인)", ["Yes", "No"], disabled=not allow_fp, horizontal=True)
                    q_data['q9'] = st.radio("10. 풀리 벨트상태", ["Yes", "No"], disabled=not allow_fp, horizontal=True)
                    q_data['q10'] = st.radio("11. 배터리 'A' 작동 여부 (수동작동)", ["Yes", "No"], disabled=not allow_fp, horizontal=True)
                    q_data['q11'] = st.radio("12. 배터리 'B' 작동 여부 (수동작동)", ["Yes", "No"], disabled=not allow_fp, horizontal=True)
                with fc2:
                    q_data['q12'] = st.text_input("13. 기동압력 (Unit: MPa)", disabled=not allow_fp)
                    q_data['q13'] = st.text_input("14. 펌프 회전속도 (Unit: RPM)", disabled=not allow_fp)
                    q_data['q14'] = st.radio("15. 펌프 패킹 누수확인", ["Yes", "No"], disabled=not allow_fp, horizontal=True)
                    q_data['q15'] = st.text_input("16. 엔진펌프 기동 시 쿨링 압력 (Unit: kPa)", disabled=not allow_fp)
                    q_data['q16'] = st.text_input("17. 엔진펌프 가동 시 오일 압력 (Unit: kPa)", disabled=not allow_fp)
                    q_data['q17'] = st.text_input("18. 엔진펌프 가동 시 토출 압력 (Unit: Bar)", disabled=not allow_fp)
                    q_data['q18'] = st.radio("19. 냉각수 온도이상 알람 여부", ["No", "Yes"], disabled=not allow_fp, horizontal=True)
                    q_data['q19'] = st.radio("20. 수신반 작동알람 여부", ["Yes", "No"], disabled=not allow_fp, horizontal=True)
                    q_data['q20'] = st.radio("21. 기동 제어 스위치 자동설정 여부", ["Yes", "No"], disabled=not allow_fp, horizontal=True)
                    q_data['q21'] = st.radio("22. 점검 완료 후 2차측 밸브 개방여부", ["Yes", "No"], disabled=not allow_fp, horizontal=True)
                    q_data['q22'] = st.text_input("23. 점검 수행시간 (최소 10분 이상)", disabled=not allow_fp)

            elif pump_category == "2. 소방 전기펌프 점검":
                st.markdown("### ⚡ 소방 전기펌프 점검")
                ec1, ec2 = st.columns(2)
                with ec1:
                    q_data['q1'] = st.radio("2. 수원의 적정수량 유지 (수조 20% 이상)", ["Yes", "No"], disabled=not allow_fp, horizontal=True)
                    q_data['q2'] = st.text_input("3. 점검 수행시간 (최소 10분 이상)", disabled=not allow_fp)
                    q_data['q3'] = st.text_input("4. 기동 전 배관압력 (Unit: Bar)", disabled=not allow_fp)
                    q_data['q4'] = st.radio("5. 압력스위치 작동여부", ["Yes", "No"], disabled=not allow_fp, horizontal=True)
                    q_data['q5'] = st.radio("6. 주펌프 작동여부", ["Yes", "No"], disabled=not allow_fp, horizontal=True)
                    q_data['q6'] = st.radio("7. 충압펌프 작동여부", ["Yes", "No"], disabled=not allow_fp, horizontal=True)
                    q_data['q7'] = st.radio("8. 펌프 작동 후 가동 유지", ["Yes", "No"], disabled=not allow_fp, horizontal=True)
                    q_data['q8'] = st.text_input("9. 펌프 토출압력 (Unit: Bar)", disabled=not allow_fp)
                with ec2:
                    q_data['q9'] = st.radio("10. 릴리프밸브 작동여부", ["Yes", "No"], disabled=not allow_fp, horizontal=True)
                    q_data['q10'] = st.radio("11. 모터 과열 여부", ["No", "Yes"], disabled=not allow_fp, horizontal=True)
                    q_data['q11'] = st.radio("12. 펌프실 누수 여부", ["No", "Yes"], disabled=not allow_fp, horizontal=True)
                    q_data['q12'] = st.radio("13. 배관 누수여부", ["No", "Yes"], disabled=not allow_fp, horizontal=True)
                    q_data['q13'] = st.radio("14. 펌프실 내부 정리정돈 상태", ["Yes", "No"], disabled=not allow_fp, horizontal=True)
                    q_data['q14'] = st.radio("15. 펌프실 배수상태 확인", ["Yes", "No"], disabled=not allow_fp, horizontal=True)
                    q_data['q15'] = st.radio("16. 수신반 가동 표시 알람 여부", ["Yes", "No"], disabled=not allow_fp, horizontal=True)
                    q_data['q16'] = st.radio("17. 흡입/토출 밸브 개방 가능여부", ["Yes", "No"], disabled=not allow_fp, horizontal=True)

            elif pump_category == "3. 스프링클러 밸브점검":
                st.markdown("### 🚿 스프링클러 밸브점검 (Open/Shut)")
                vc1, vc2 = st.columns(2)
                with vc1:
                    q_data['q1'] = st.radio("1. P3-V1 주펌프 (Main pump)", ["Open", "Shut"], disabled=not allow_fp, horizontal=True)
                    q_data['q2'] = st.radio("2. P3-V2 충압펌프 (Jocky pump)", ["Open", "Shut"], disabled=not allow_fp, horizontal=True)
                    q_data['q3'] = st.radio("3. P3-V3 엔진펌프 (Engine pump)", ["Open", "Shut"], disabled=not allow_fp, horizontal=True)
                    q_data['q4'] = st.radio("4. P3-V4 1F 스프링클러 (Alarm valve)", ["Open", "Shut"], disabled=not allow_fp, horizontal=True)
                    q_data['q5'] = st.radio("5. P3-V5 1F 스프링클러 (Pre action VV)", ["Open", "Shut"], disabled=not allow_fp, horizontal=True)
                with vc2:
                    q_data['q6'] = st.radio("6. P3-V6 1F 드렌처 설비 (Pre action VV)", ["Open", "Shut"], disabled=not allow_fp, horizontal=True)
                    q_data['q7'] = st.radio("7. P3-V7 2F 스프링클러 (Alarm valve)", ["Open", "Shut"], disabled=not allow_fp, horizontal=True)
                    q_data['q8'] = st.radio("8. P3-V8 3F 스프링클러 (Alarm valve)", ["Open", "Shut"], disabled=not allow_fp, horizontal=True)
                    q_data['q9'] = st.radio("9. P3-V9 3F 스프링클러 (Alarm valve)", ["Open", "Shut"], disabled=not allow_fp, horizontal=True)

            fp_notes = st.text_area("기타 특이사항 (점검간 이상증상 및 조치 필요 사항)", height=80, disabled=not allow_fp)

            if st.form_submit_button(f"💾 {pump_category} 저장", use_container_width=True, disabled=not allow_fp):
                if not fp_ins.strip():
                    st.error("점검자 성명을 입력해주세요.")
                else:
                    save_fire_pump_log(fp_date, pump_category, fp_ins, q_data, fp_notes)
                    st.rerun()

    with tab_fp2:
        st.subheader("📂 소방펌프/밸브 점검 기록 리포트 조회 및 다운로드")
        fp_df = get_fire_pump_logs()
        if not fp_df.empty:
            sc_c1, sc_c2 = st.columns(2)
            with sc_c1:
                sel_cat = st.selectbox("출력할 설비 선택", ["1. 소방 엔진펌프 점검", "2. 소방 전기펌프 점검", "3. 스프링클러 밸브점검"])
            with sc_c2:
                filtered_cat_df = fp_df[fp_df['category'] == sel_cat]
                if not filtered_cat_df.empty:
                    sel_d_obj = st.date_input("📅 조회 및 다운로드할 점검 일자 선택", value=datetime.now(), key="fp_date_input")
                    sel_d = str(sel_d_obj)
                else:
                    st.warning(f"선택하신 '{sel_cat}'의 기록이 없습니다.")
                    sel_d = None

            if sel_d and not filtered_cat_df.empty:
                filtered_df = filtered_cat_df[filtered_cat_df['record_date'] == sel_d]
                if not filtered_df.empty:
                    t_row = filtered_df.iloc[0]
                    q_dict = json.loads(t_row['q_data'])
                    
                    if sel_cat == "1. 소방 엔진펌프 점검":
                        rep_map = [
                            ["제어반 표시등 점등 여부", q_dict.get('q1','')], ["연료탱크 경유 적정량", q_dict.get('q2','')],
                            ["배터리 #1 충전상태 (V)", q_dict.get('q3','')], ["배터리 #2 충전상태 (V)", q_dict.get('q4','')],
                            ["엔진오일 상태", q_dict.get('q5','')], ["배터리 충전기 작동여부", q_dict.get('q6','')],
                            ["연료/유류 배관 파손여부", q_dict.get('q7','')], ["윤활유 유량", q_dict.get('q8','')],
                            ["쿨링 벨트상태", q_dict.get('q9','')], ["배터리 'A' 작동 여부", q_dict.get('q10','')],
                            ["배터리 'B' 작동 여부", q_dict.get('q11','')], ["기동압력 (MPa)", q_dict.get('q12','')],
                            ["펌프 회전속도 (RPM)", q_dict.get('q13','')], ["펌프 패킹 누수확인", q_dict.get('q14','')],
                            ["기동 시 쿨링 압력 (kPa)", q_dict.get('q15','')], ["기동 시 오일 압력 (kPa)", q_dict.get('q16','')],
                            ["가동 시 토출 압력 (Bar)", q_dict.get('q17','')], ["냉각수 온도이상 알람", q_dict.get('q18','')],
                            ["수신반 작동알람 여부", q_dict.get('q19','')], ["기동 제어 스위치 자동설정", q_dict.get('q20','')],
                            ["2차측 밸브 개방여부", q_dict.get('q21','')], ["점검 수행시간", q_dict.get('q22','')],
                            ["기타 특이사항", t_row.get('notes','')]
                        ]
                    elif sel_cat == "2. 소방 전기펌프 점검":
                        rep_map = [
                            ["수원의 적정수량 유지", q_dict.get('q1','')], ["점검 수행시간", q_dict.get('q2','')],
                            ["기동 전 체절압력 (Bar)", q_dict.get('q3','')], ["압력스위치 작동여부", q_dict.get('q4','')],
                            ["주펌프 작동여부", q_dict.get('q5','')], ["충압펌프 작동여부", q_dict.get('q6','')],
                            ["펌프 작동 후 가동 유지", q_dict.get('q7','')], ["펌프 토출압력 (Bar)", q_dict.get('q8','')],
                            ["릴리프밸브 작동여부", q_dict.get('q9','')], ["모터 과열 여부", q_dict.get('q10','')],
                            ["펌프실 누수 여부", q_dict.get('q11','')], ["배관 누수여부", q_dict.get('q12','')],
                            ["펌프실 내부 정리정돈", q_dict.get('q13','')], ["펌프실 배수상태", q_dict.get('q14','')],
                            ["수신반 가동 표시 알람", q_dict.get('q15','')], ["흡입/토출 밸브 개방 여부", q_dict.get('q16','')],
                            ["기타 특이사항", t_row.get('notes','')]
                        ]
                    else:
                        rep_map = [
                            ["P3-V1 주펌프", q_dict.get('q1','')], ["P3-V2 충압펌프", q_dict.get('q2','')],
                            ["P3-V3 엔진펌프", q_dict.get('q3','')], ["P3-V4 1F 스프링클러(Alarm)", q_dict.get('q4','')],
                            ["P3-V5 1F 스프링클러(Pre action)", q_dict.get('q5','')], ["P3-V6 1F 드렌처 설비", q_dict.get('q6','')],
                            ["P3-V7 2F 스프링클러(Alarm)", q_dict.get('q7','')], ["P3-V8 3F 스프링클러(Alarm)", q_dict.get('q8','')],
                            ["P3-V9 3F 스프링클러(Alarm)", q_dict.get('q9','')], ["기타 특이사항", t_row.get('notes','')]
                        ]

                    r_df = pd.DataFrame(rep_map, columns=["점검 항목", "점검 결과"])
                    st.markdown(f"📌 [{sel_cat}] 일자: {sel_d} / 점검자: {t_row.get('inspector','')}]")
                    st.dataframe(r_df, use_container_width=True, height=450)
                    st.download_button(
                        f"📥 [{sel_d}] {sel_cat} 점검표 다운로드",
                        r_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig'),
                        f"Fire_Inspection_{sel_d}.csv",
                        "text/csv",
                        use_container_width=True
                    )
                else:
                    st.warning(f"선택하신 날짜({sel_d})에는 '{sel_cat}' 점검 기록이 없습니다.")
            with st.expander("원본 가로형 데이터베이스 로그 전체 보기"):
                st.dataframe(fp_df, use_container_width=True)
        else:
            st.info("등록된 기록이 없습니다.")
