"""
재고실사 웹 대시보드 (Streamlit)

여러 담당자가 구역별로 실사 수량을 입력하고, 전체 진행률과 장부 대비
차이 결과를 실시간으로 확인할 수 있는 웹앱.

저장소: 로컬 CSV (실사기록.csv) — Streamlit Cloud에 배포하면 서버 하나에
모든 사용자가 접속하는 구조라 이 파일 하나를 공유 저장소처럼 쓸 수 있음.
(주의: Streamlit Cloud 무료 플랜은 앱 재시작 시 파일이 초기화될 수 있음.
 장기 운영 시엔 Google Sheets(gspread) 연동으로 교체 권장 — 하단 주석 참고)

실행:
    pip install streamlit pandas openpyxl plotly
    streamlit run app.py
"""

import os
from datetime import datetime

import pandas as pd
import streamlit as st
import plotly.express as px

LEDGER_PATH = "장부재고.xlsx"
RECORD_PATH = "실사기록.csv"
ZONES = ["지상 창고", "지하 창고", "매장 1", "매장 2", "매장 3"]

st.set_page_config(page_title="재고실사 대시보드", layout="wide")


# ---------- 데이터 로드/저장 ----------

@st.cache_data(ttl=5)
def load_ledger():
    return pd.read_excel(LEDGER_PATH)  # 품목코드, 품목명, 장부수량


def load_records():
    if not os.path.exists(RECORD_PATH):
        return pd.DataFrame(columns=["구역", "담당자", "품목코드", "품목명", "실사수량", "입력시각"])
    return pd.read_csv(RECORD_PATH)


def append_record(zone, staff, item_code, item_name, qty):
    df = load_records()
    new_row = {
        "구역": zone,
        "담당자": staff,
        "품목코드": item_code,
        "품목명": item_name,
        "실사수량": qty,
        "입력시각": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_csv(RECORD_PATH, index=False)


def delete_record(row_index):
    """실사기록.csv에서 특정 행(원본 인덱스 기준)을 삭제한다."""
    df = load_records()
    df = df.drop(index=row_index).reset_index(drop=True)
    df.to_csv(RECORD_PATH, index=False)


def update_record(row_index, zone, staff, item_code, item_name, qty):
    """실사기록.csv에서 특정 행의 값을 수정한다."""
    df = load_records()
    df.loc[row_index, ["구역", "담당자", "품목코드", "품목명", "실사수량"]] = [
        zone, staff, item_code, item_name, qty
    ]
    df.loc[row_index, "입력시각"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    df.to_csv(RECORD_PATH, index=False)


def classify(row):
    if pd.isna(row["장부수량"]):
        return "신규(장부누락)"
    if pd.isna(row["실사수량"]):
        return "실사누락"
    if row["장부수량"] == row["실사수량"]:
        return "일치"
    return "수량차이"


def reconcile(ledger, records):
    # 같은 품목이 여러 구역/담당자에 의해 나눠서 입력될 수 있으므로,
    # 품목코드 기준으로 모든 입력 수량을 합산한다 (예: A구역 후드티 3개 + B구역 후드티 5개 = 8개)
    counted = records.groupby("품목코드", as_index=False)["실사수량"].sum()

    merged = pd.merge(
        ledger[["품목코드", "품목명", "장부수량"]],
        counted,
        on="품목코드",
        how="outer",
    )
    merged["품목명"] = merged["품목명"].fillna(
        merged["품목코드"].map(records.drop_duplicates("품목코드").set_index("품목코드")["품목명"])
    )
    merged["차이수량"] = merged["실사수량"] - merged["장부수량"]
    merged["상태"] = merged.apply(classify, axis=1)
    return merged.sort_values(by=["상태", "품목코드"]).reset_index(drop=True)


def highlight_status(row):
    color_map = {
        "일치": "background-color: #E2EFDA",
        "수량차이": "background-color: #FFF2CC",
        "실사누락": "background-color: #F8CBAD",
        "신규(장부누락)": "background-color: #D9E1F2",
    }
    color = color_map.get(row["상태"], "")
    return [color] * len(row)


# ---------- 화면 구성 ----------

st.title("📦 재고실사 대시보드")
page = st.sidebar.radio("메뉴", ["실사 입력", "진행률", "대사 결과"])

ledger = load_ledger()
records = load_records()

if page == "실사 입력":
    st.subheader("실사 수량 입력")
    col1, col2 = st.columns(2)
    with col1:
        zone = st.selectbox("구역 선택", ZONES)
        staff = st.text_input("담당자 이름")
        item_code = st.selectbox("품목코드", ledger["품목코드"].tolist())
        item_name = ledger.loc[ledger["품목코드"] == item_code, "품목명"].values[0]
        st.write(f"품목명: **{item_name}**")
    with col2:
        qty = st.number_input("실사 수량", min_value=0, step=1)

    if st.button("입력 저장", type="primary"):
        if not staff.strip():
            st.warning("담당자 이름을 입력해주세요.")
        else:
            append_record(zone, staff, item_code, item_name, qty)
            st.success(f"{zone} / {staff} / {item_name} / {qty}개 저장 완료")
            st.cache_data.clear()
            st.rerun()

    st.divider()
    st.caption("최근 입력 내역 (잘못 입력한 건은 아래에서 수정/삭제하세요)")

    recent = records.sort_values("입력시각", ascending=False).head(10)
    if recent.empty:
        st.info("아직 입력된 내역이 없습니다.")
    else:
        st.dataframe(recent, use_container_width=True)

        # 원본 records의 인덱스를 그대로 유지해야 정확한 행을 수정/삭제할 수 있음
        options = {
            f"[{idx}] {row['구역']} / {row['담당자']} / {row['품목명']} / {row['실사수량']}개 ({row['입력시각']})": idx
            for idx, row in recent.iterrows()
        }
        selected_label = st.selectbox("수정 또는 삭제할 항목 선택", list(options.keys()))
        selected_idx = options[selected_label]
        selected_row = records.loc[selected_idx]

        st.write("**선택한 항목 수정**")
        ecol1, ecol2 = st.columns(2)
        with ecol1:
            new_zone = st.selectbox("구역", ZONES, index=ZONES.index(selected_row["구역"]) if selected_row["구역"] in ZONES else 0, key="edit_zone")
            new_staff = st.text_input("담당자", value=selected_row["담당자"], key="edit_staff")
        with ecol2:
            new_qty = st.number_input("실사 수량", min_value=0, step=1, value=int(selected_row["실사수량"]), key="edit_qty")

        bcol1, bcol2 = st.columns(2)
        with bcol1:
            if st.button("이 항목 수정 저장", type="primary"):
                update_record(selected_idx, new_zone, new_staff, selected_row["품목코드"], selected_row["품목명"], new_qty)
                st.success("수정 완료")
                st.cache_data.clear()
                st.rerun()
        with bcol2:
            if st.button("이 항목 삭제", type="secondary"):
                delete_record(selected_idx)
                st.success("삭제 완료")
                st.cache_data.clear()
                st.rerun()

elif page == "진행률":
    st.subheader("실사 진행률")

    total_qty = ledger["장부수량"].sum()
    counted_qty = records["실사수량"].sum() if not records.empty else 0
    qty_progress = min(counted_qty / total_qty, 1.0) if total_qty else 0

    total_items = ledger["품목코드"].nunique()
    touched_items = records["품목코드"].nunique() if not records.empty else 0

    col1, col2 = st.columns(2)
    with col1:
        st.metric("수량 기준 진행률", f"{int(counted_qty):,} / {int(total_qty):,}개", f"{qty_progress*100:.1f}%")
        st.progress(qty_progress)
        st.caption("장부수량 대비 지금까지 실제로 센 수량의 비율입니다. 한 품목을 세는 도중이어도 정확히 반영돼요.")
    with col2:
        item_progress = touched_items / total_items if total_items else 0
        st.metric("품목 커버리지", f"{touched_items} / {total_items} 품목", f"{item_progress*100:.1f}%")
        st.progress(item_progress)
        st.caption("한 건이라도 입력이 시작된 품목의 비율입니다. (수량이 다 안 세어졌어도 카운트됨)")

    st.divider()
    st.caption("구역별 입력 수량")
    if not records.empty:
        zone_summary = records.groupby("구역")["실사수량"].sum().reset_index()
        zone_summary.columns = ["구역", "입력한 수량"]
        st.dataframe(zone_summary, use_container_width=True)

        st.divider()
        st.caption("담당자별 입력 수량")
        staff_summary = records.groupby("담당자")["실사수량"].sum().reset_index()
        staff_summary.columns = ["담당자", "입력한 수량"]
        st.dataframe(staff_summary, use_container_width=True)

        st.divider()
        st.caption("품목별 진행률 (장부수량 대비 실사 완료 수량)")
        item_counted = records.groupby("품목코드", as_index=False)["실사수량"].sum()
        item_progress_df = pd.merge(ledger[["품목코드", "품목명", "장부수량"]], item_counted, on="품목코드", how="left")
        item_progress_df["실사수량"] = item_progress_df["실사수량"].fillna(0)
        item_progress_df["진행률(%)"] = (item_progress_df["실사수량"] / item_progress_df["장부수량"] * 100).round(1)
        item_progress_df["진행률(캡100%)"] = item_progress_df["진행률(%)"].clip(upper=100)

        chart_df = item_progress_df.sort_values("진행률(%)")
        fig = px.bar(
            chart_df,
            x="진행률(캡100%)",
            y="품목명",
            orientation="h",
            text=chart_df["진행률(%)"].astype(str) + "%",
            range_x=[0, 100],
            color="진행률(캡100%)",
            color_continuous_scale=["#F8CBAD", "#FFF2CC", "#E2EFDA"],
        )
        fig.update_layout(
            xaxis_title="진행률 (%)", yaxis_title="",
            coloraxis_showscale=False, height=80 + 40 * len(chart_df),
        )
        st.plotly_chart(fig, use_container_width=True)

        with st.expander("표로 보기"):
            st.dataframe(item_progress_df.drop(columns=["진행률(캡100%)"]), use_container_width=True)
    else:
        st.info("아직 입력된 실사 기록이 없습니다.")

elif page == "대사 결과":
    st.subheader("장부 vs 실사 대사 결과")
    if records.empty:
        st.info("실사 입력이 아직 없어 대사할 데이터가 없습니다.")
    else:
        result = reconcile(ledger, records)
        diff_count = len(result[result["상태"] != "일치"])
        st.metric("차이 발견 항목", f"{diff_count}건")

        styled = result.style.apply(highlight_status, axis=1).format(
            {"장부수량": "{:.0f}", "실사수량": "{:.0f}", "차이수량": "{:.0f}"}, na_rep="-"
        )
        st.dataframe(styled, use_container_width=True)

        csv = result.to_csv(index=False).encode("utf-8-sig")
        st.download_button("대사결과 CSV 다운로드", csv, "대사결과.csv", "text/csv")

# ---------------------------------------------------------------
# [확장 옵션] Google Sheets 연동으로 교체하려면:
#   import gspread
#   from google.oauth2.service_account import Credentials
#   gc = gspread.service_account(filename="credentials.json")
#   sheet = gc.open("재고실사기록").sheet1
#   -> append_record(): sheet.append_row([...])
#   -> load_records(): pd.DataFrame(sheet.get_all_records())
# 로 load_records/append_record 두 함수만 교체하면 나머지 로직은 그대로 재사용 가능
# ---------------------------------------------------------------
