#!/usr/bin/env python3
"""
매출 데이터 자동 업데이트 스크립트
사용법: python3 update_data.py 파일1.xlsx 파일2.xlsx 파일3.xlsx 파일4.xlsx
  - 주문별_국내, 주문별_해외, 상품별_국내, 상품별_해외 순서 무관 (자동 판별)
"""

import sys, re, json, os
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from datetime import datetime

# ── 설정 ────────────────────────────────────────────────────────────────────
USD_TO_KRW   = 1460  # fallback (2026-07-15 이후 기본값)

def fx_rate(date_str):
    """구간별 환율: 2025년=1450 고정, 2026-01-01~2026-07-14=1500 고정,
    2026-07-15 이후=1460 (2026-07-28 대표님 확정, 실제 환율 1500원 붕괴 반영)"""
    ds = str(date_str)
    if ds[:4] == "2025":
        return 1450
    if ds >= "2026-07-15":
        return 1460
    return 1500
BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
MIN_DATE     = "2025-01-01"
CANCEL_STATUS = {"취소 완료","반품 완료","반품 요청","입금 대기","교환 완료"}
CC_NAME = {
    "US":"United States","CA":"Canada","AU":"Australia","GB":"United Kingdom",
    "DE":"Germany","FR":"France","JP":"Japan","SG":"Singapore","HK":"Hong Kong",
    "TW":"Taiwan","NL":"Netherlands","SE":"Sweden","NO":"Norway","DK":"Denmark",
    "NZ":"New Zealand","CH":"Switzerland","AT":"Austria","BE":"Belgium","IT":"Italy",
    "ES":"Spain","MX":"Mexico","TH":"Thailand","MY":"Malaysia","PH":"Philippines",
    "ID":"Indonesia","VN":"Vietnam","AE":"United Arab Emirates","SA":"Saudi Arabia",
    "TR":"Turkey","PL":"Poland","IE":"Ireland","KW":"Kuwait","QA":"Qatar",
    "GR":"Greece","PT":"Portugal","FI":"Finland","CZ":"Czech Republic","HU":"Hungary",
    "EE":"Estonia","RO":"Romania","SK":"Slovakia","LU":"Luxembourg","CN":"China",
    "CL":"Chile","IL":"Israel","PR":"Puerto Rico","GU":"Guam",
}

# 26HS 핫썸머 시즌 상품명 목록 (신규 상품 자동 시즌 분류용, _26HS 상세가이드 기준)
SEASON_26HS_NAMES = {
    "Rope Kiko Doule Botton Shorts Khaki","Rope Kiko Doule Botton Shorts Mint",
    "Ring Buoy Bermuda Pants","Strawberry Mlik Punch Mini Dress","Turquoise Wave Lace Mini Dress",
    "Coconut Stripe Sleeveless","Strawberry Mlik Punch Sleeveless","Fruit Mlik Punch Sleeveless",
    "Holo Dot Botton Half Shirts","Holo Stripe Botton Half Shirts","Sailing Stripe Lace-up Half Shirts",
    "Sailing Dot Lace-up Half Shirts","Strawberry Milk Vacation Cardigan","Fruit Milk Vacation Cardigan",
    "Jelly Lemon Dimsum Bag","Jelly Mint Dimsum Bag","Jelly Peach Dimsum Bag",
    "Jelly Leomon Dumpling Bag","Jelly Mint Dumpling Bag","Jelly Peach Dumpling Bag",
    "Rattan Beige Dimsum Bag","Rattan Pistachio Dimsum Bag","Rattan Black Dimsum Bag",
    "Mini Glaze Lemon Dimsum Bag","Mini Glaze Pistachio Dimsum Bag","Mini Glaze Night Dimsum Bag",
    "Sunset Beach Dumpling Bag","Ocean Blue Dumpling Bag",
    "Mini Jelly Lemon Dimsum Bag","Mini Jelly Mint Dimsum Bag","Mini Jelly Peach Dimsum Bag",
}
SEASON_26HS_NAMES_STRIP = {n.strip() for n in SEASON_26HS_NAMES}

# 26 FW_1 시즌 상품명 목록 (2026-08-19 대표님 지정, 프리오더 신상 — Ships on Sep 10)
SEASON_26FW1_NAMES = {
    "Cocoa Check Pillow Bag","Sky Berry Check Pillow Bag",
    "Hazy Leopard Charcoal Pillow Bag","Hazy Leopard Charcoal Dimsum Bag",
    "Mini Hazy Leopard Charcoal Dimsum Bag",
}
SEASON_26FW1_NAMES_STRIP = {n.strip() for n in SEASON_26FW1_NAMES}

def season_for_new_product(nm, season_map):
    """기존 시즌맵에 없는 신규 상품의 시즌을 결정. 26HS/26FW_1 가이드에 있으면 해당 시즌, 아니면 기타."""
    se = season_map.get(nm)
    if se: return se
    nms = str(nm).strip()
    if nms in SEASON_26HS_NAMES_STRIP:
        return "26 HS 핫썸머"
    if nms in SEASON_26FW1_NAMES_STRIP:
        return "26 FW_1"
    return "기타"

errors   = []
warnings = []

def log(msg):  print(f"  {msg}")
def ok(msg):   print(f"  ✅ {msg}")
def warn(msg): print(f"  ⚠️  {msg}"); warnings.append(msg)
def err(msg):  print(f"  ❌ {msg}"); errors.append(msg)

# ── 파일 자동 판별 ───────────────────────────────────────────────────────────
def classify_files(paths):
    result = {}
    for p in paths:
        df = pd.read_excel(p, header=0, nrows=3)
        cols = list(df.columns)
        col_str = " ".join(str(c) for c in cols)
        is_product = "상품 이름" in col_str or "상품명" in col_str
        # 해외 판별: 결제금액이 소수(USD) 또는 주소에 영문 국가코드
        sample = df.iloc[0] if len(df) > 0 else pd.Series()
        pay_col = [c for c in cols if "결제 금액" in str(c)]
        is_overseas = False
        if pay_col:
            val = sample.get(pay_col[0], 0)
            try:
                fval = float(val)
                is_overseas = (fval < 5000)  # USD 기준 (KRW는 최소 수천)
            except: pass
        if not is_overseas:
            addr_col = [c for c in cols if "배송지 주소" in str(c)]
            if addr_col:
                addr = str(sample.get(addr_col[0], ""))
                is_overseas = bool(re.search(r"\b[A-Z]{2}\s*$", addr.strip()))
        key = ("product" if is_product else "order") + "_" + ("ovs" if is_overseas else "dom")
        if key in result:
            warn(f"동일 종류 파일 중복 감지: {key} — {os.path.basename(p)}, {os.path.basename(result[key])}")
        result[key] = p
        log(f"{os.path.basename(p)} → [{key}]")
    return result

def extract_cc(addr):
    if pd.isna(addr): return "US"
    m = re.search(r"\b([A-Z]{2})\s*$", str(addr).strip())
    return m.group(1) if m else "US"

# ── 엑셀 읽기 + 전처리 ──────────────────────────────────────────────────────
def load_and_clean(path):
    df = pd.read_excel(path, header=0)
    df["주문 일자"] = pd.to_datetime(df["주문 일자"], errors="coerce")
    df["날짜"]     = df["주문 일자"].dt.date.astype(str)
    df = df[~df["주문 상태"].isin(CANCEL_STATUS)].copy()
    return df

# ── 01_일별매출현황.csv ──────────────────────────────────────────────────────
def update_daily(dom_ord, ovs_ord, max_date_in_csv):
    print("\n[1/6] 01_일별매출현황.csv")
    path = os.path.join(BASE_DIR, "01_일별매출현황.csv")
    df_cur = pd.read_csv(path)

    dom_g = dom_ord.drop_duplicates("주문번호").groupby("날짜").agg(
        국내매출_KRW=("결제 금액","sum"), 국내주문건수=("주문번호","count")).reset_index()
    ovs_g = ovs_ord.drop_duplicates("주문번호").groupby("날짜").agg(
        해외매출_USD=("결제 금액","sum"), 해외주문건수=("주문번호","count")).reset_index()

    all_dates = sorted(set(dom_g["날짜"].tolist() + ovs_g["날짜"].tolist()))
    # 이미 있는 날짜 skip
    existing = set(df_cur["날짜"].astype(str).tolist())
    new_dates = [d for d in all_dates if d > max_date_in_csv]
    if not new_dates:
        warn("추가할 새 날짜 없음 (이미 반영됐거나 날짜 범위 확인 필요)")
        return df_cur

    rows = []
    for d in new_dates:
        dr = dom_g[dom_g["날짜"]==d]
        or_ = ovs_g[ovs_g["날짜"]==d]
        dom_rev = float(dr["국내매출_KRW"].values[0]) if len(dr) else 0.0
        dom_cnt = int(dr["국내주문건수"].values[0])   if len(dr) else 0
        ovs_usd = float(or_["해외매출_USD"].values[0]) if len(or_) else 0.0
        ovs_cnt = int(or_["해외주문건수"].values[0])   if len(or_) else 0
        ovs_krw = round(ovs_usd * fx_rate(d), 0)
        rows.append({"날짜":d,"월":d[:7],
            "통합매출_KRW":dom_rev+ovs_krw,"통합주문건수":float(dom_cnt+ovs_cnt),
            "국내매출_KRW":dom_rev,"국내주문건수":dom_cnt,
            "해외매출_USD":round(ovs_usd,2),"해외매출_KRW환산":ovs_krw,"해외주문건수":float(ovs_cnt)})
    df_new = pd.DataFrame(rows)
    df_out = pd.concat([df_cur, df_new], ignore_index=True)
    df_out.to_csv(path, index=False)
    ok(f"{len(new_dates)}일 추가 → 총 {len(df_out)}행 (최신: {df_out['날짜'].max()})")
    return df_out

# ── 02_상품별판매현황.csv ────────────────────────────────────────────────────
def update_products(dom_prod, ovs_prod, dom_ord, ovs_ord):
    print("\n[2/6] 02_상품별판매현황.csv")
    path = os.path.join(BASE_DIR, "02_상품별판매현황.csv")
    df_cur = pd.read_csv(path)
    season_map = df_cur.set_index("상품명")["시즌"].to_dict()

    dom_agg = dom_prod.groupby("상품 이름").agg(
        국내판매수량=("수량","sum"), 국내매출_KRW=("상품별 결제 금액","sum"),
        국내주문건수=("주문번호","nunique")).reset_index().rename(columns={"상품 이름":"상품명"})
    ovs_prod["상품별_KRW"] = ovs_prod["상품별 결제 금액"].fillna(0) * ovs_prod["날짜"].apply(fx_rate)
    ovs_agg = ovs_prod.groupby("상품 이름").agg(
        해외판매수량=("수량","sum"), 해외매출_KRW환산=("상품별_KRW","sum"),
        해외주문건수=("주문번호","nunique")).reset_index().rename(columns={"상품 이름":"상품명"})

    delta = pd.merge(dom_agg, ovs_agg, on="상품명", how="outer").fillna(0)
    delta["통합판매수량"] = delta["국내판매수량"] + delta["해외판매수량"]
    delta["통합매출_KRW"] = delta["국내매출_KRW"] + delta["해외매출_KRW환산"]
    delta["총주문건수"]   = delta["국내주문건수"] + delta["해외주문건수"]
    delta["채널구분"] = delta.apply(lambda r:
        "국내전용" if r["해외판매수량"]==0 else ("해외전용" if r["국내판매수량"]==0 else "국내+해외"), axis=1)

    idx = df_cur.set_index("상품명")
    new_cnt = 0
    for _, row in delta.iterrows():
        nm = row["상품명"]
        if nm in idx.index:
            for col, add in [
                ("통합판매수량",row["통합판매수량"]),("국내판매수량",row["국내판매수량"]),
                ("해외판매수량",row["해외판매수량"]),("통합매출_KRW",row["통합매출_KRW"]),
                ("국내매출_KRW",row["국내매출_KRW"]),("해외매출_KRW환산",row["해외매출_KRW환산"]),
                ("총주문건수",row["총주문건수"]),("국내주문건수",row["국내주문건수"]),
                ("해외주문건수",row["해외주문건수"]),("통합주문건수",row["총주문건수"]),
            ]:
                idx.at[nm, col] = idx.at[nm, col] + add
            tq = idx.at[nm,"통합판매수량"]
            if tq > 0:
                idx.at[nm,"평균단가_KRW"] = int(idx.at[nm,"통합매출_KRW"] / tq)
        else:
            new_cnt += 1
            se = season_for_new_product(nm, season_map)
            tq = row["통합판매수량"]
            avg = int(row["통합매출_KRW"]/tq) if tq > 0 else 0
            all_d = pd.concat([
                dom_prod[dom_prod["상품 이름"]==nm]["주문 일자"],
                ovs_prod[ovs_prod["상품 이름"]==nm]["주문 일자"]
            ]).dropna()
            first = str(all_d.min().date()) if len(all_d) else "2026-01-01"
            nr = pd.Series({"상품명":nm,"채널구분":row["채널구분"],
                "통합판매수량":tq,"국내판매수량":row["국내판매수량"],"해외판매수량":row["해외판매수량"],
                "통합매출_KRW":row["통합매출_KRW"],"국내매출_KRW":row["국내매출_KRW"],
                "해외매출_KRW환산":row["해외매출_KRW환산"],"평균단가_KRW":avg,
                "총주문건수":row["총주문건수"],"국내주문건수":row["국내주문건수"],
                "해외주문건수":row["해외주문건수"],"최초판매일":first,"시즌":se,
                "통합주문건수":row["총주문건수"]})
            idx = pd.concat([idx, nr.to_frame().T.set_index("상품명")])

    df_out = idx.reset_index()
    df_out.to_csv(path, index=False)
    ok(f"갱신 완료 → 총 {len(df_out)}개 상품 (신규 {new_cnt}개)")
    return df_out

# ── 03_주문원본.csv ──────────────────────────────────────────────────────────
def update_raw(dom_ord, ovs_ord, max_date_in_csv):
    print("\n[3/6] 03_주문원본.csv")
    path = os.path.join(BASE_DIR, "03_주문원본.csv")
    df_cur = pd.read_csv(path)

    rows = []
    for _, r in dom_ord.drop_duplicates("주문번호").iterrows():
        if r["날짜"] > max_date_in_csv:
            rows.append({"날짜":r["날짜"],"월":r["날짜"][:7],"채널":"국내자사몰",
                "결제금액_KRW":float(r["결제 금액"]),"결제금액_USD":0.0,
                "상품수":int(r["주문 품목 개수"]) if pd.notna(r.get("주문 품목 개수")) else 1})
    for _, r in ovs_ord.drop_duplicates("주문번호").iterrows():
        if r["날짜"] > max_date_in_csv:
            usd = float(r["결제 금액"])
            rows.append({"날짜":r["날짜"],"월":r["날짜"][:7],"채널":"해외자사몰",
                "결제금액_KRW":round(usd*fx_rate(r["날짜"]),0),"결제금액_USD":usd,
                "상품수":int(r["주문 품목 개수"]) if pd.notna(r.get("주문 품목 개수")) else 1})

    if not rows:
        warn("03: 추가할 새 주문 없음")
        return df_cur
    df_out = pd.concat([df_cur, pd.DataFrame(rows)], ignore_index=True)
    df_out.to_csv(path, index=False)
    ok(f"{len(rows)}건 추가 → 총 {len(df_out)}행")
    return df_out

# ── 해외배송국가_DB.xlsx ─────────────────────────────────────────────────────
def update_overseas_db(ovs_ord, ovs_prod, max_date_in_csv):
    print("\n[4/6] 해외배송국가_DB.xlsx")
    path = os.path.join(BASE_DIR, "해외배송국가_DB.xlsx")
    df_cur = pd.read_excel(path, sheet_name="해외주문 전체")

    ovs_ord = ovs_ord.copy()
    ovs_ord["cc"] = ovs_ord["배송지 주소"].apply(extract_cc)
    order_products = ovs_prod.groupby("주문번호").apply(
        lambda g: ", ".join(f"{r['상품 이름']}(x{int(r['수량'])})" for _, r in g.iterrows())
    ).to_dict()

    rows = []
    for _, r in ovs_ord.drop_duplicates("주문번호").iterrows():
        if r["날짜"] <= max_date_in_csv:
            continue
        cc = r["cc"]
        usd = float(r["결제 금액"])
        rows.append({"날짜":r["날짜"],"월":r["날짜"][:7],"주문번호":r["주문번호"],
            "국가코드":cc,"국가명":CC_NAME.get(cc,"기타"),
            "결제금액(USD)":usd,"결제금액(KRW)":round(usd*fx_rate(r["날짜"]),0),
            "상품수":int(r.get("주문 품목 개수",1)) if pd.notna(r.get("주문 품목 개수")) else 1,
            "상품명":order_products.get(r["주문번호"],"")})

    if not rows:
        warn("해외DB: 추가할 새 주문 없음")
        return

    df_combined = pd.concat([df_cur, pd.DataFrame(rows)], ignore_index=True)
    df_combined["날짜"] = df_combined["날짜"].astype(str)
    by_cc = df_combined.groupby(["국가코드","국가명"]).agg(
        총주문건수=("주문번호","count"), 총매출_USD=("결제금액(USD)","sum"),
        총매출_KRW=("결제금액(KRW)","sum"), 평균주문금액_USD=("결제금액(USD)","mean")
    ).reset_index().sort_values("총매출_USD", ascending=False)
    cross = df_combined.groupby(["월","국가코드"])["결제금액(USD)"].sum().unstack(fill_value=0).round(2)

    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df_combined.to_excel(w, sheet_name="해외주문 전체", index=False)
        by_cc.round(2).to_excel(w, sheet_name="국가별 요약", index=False)
        cross.to_excel(w, sheet_name="월별×국가 크로스탭")
    ok(f"{len(rows)}건 추가 → 총 {len(df_combined)}행")

# ── HTML 갱신 ────────────────────────────────────────────────────────────────
def update_html(dom_ord, ovs_ord, dom_prod, ovs_prod, df_products, max_date):
    print("\n[5/6] HTML 파일 갱신")

    season_map = df_products.set_index("상품명")["시즌"].to_dict()

    dom_ok = dom_ord.drop_duplicates("주문번호")
    ovs_ok = ovs_ord.drop_duplicates("주문번호")
    ovs_ok = ovs_ok.copy()
    ovs_ok["cc"] = ovs_ok["배송지 주소"].apply(extract_cc)
    ovs_prod_cc = ovs_prod.copy()
    ovs_prod_cc["cc"] = ovs_prod_cc["배송지 주소"].apply(extract_cc)

    # ── sales_dashboard: DOMESTIC / OVERSEAS ──
    _update_sales_dashboard(dom_ok, ovs_ok)

    # ── insights: OVS_DAILY / OVS_RAW_ORDERS ──
    _update_insights(ovs_prod_cc, ovs_ok)

    # ── product_report: PRODUCTS / DAILY / MONTHS ──
    _update_product_report(dom_prod, ovs_prod, df_products, season_map, max_date)

    # ── product_analytics: DATA.rows ──
    _update_product_analytics(dom_prod, ovs_prod, season_map)

    # ── 4개 파일 공통: date input HTML attributes ──
    _update_date_inputs(max_date)

    ok("HTML 4개 파일 갱신 완료")

def _update_sales_dashboard(dom_ok, ovs_ok):
    path = os.path.join(BASE_DIR, "sales_dashboard.html")
    with open(path) as f: c = f.read()

    dom_g = dom_ok.groupby("날짜").agg(orders=("주문번호","count"), revenue=("결제 금액","sum")).reset_index()
    ovs_g = ovs_ok.groupby("날짜").agg(orders=("주문번호","count"), revenue_usd=("결제 금액","sum")).reset_index()
    ovs_g["revenue_krw"] = (ovs_g["revenue_usd"]*ovs_g["날짜"].apply(fx_rate)).round(0).astype(int)

    # 이미 있는 날짜 skip
    existing = set(re.findall(r"date: '(\d{4}-\d{2}-\d{2})'", c))

    dom_new = "".join(
        f"\n  {{ date: '{r.날짜}', orders: {r.orders}, revenue: {int(r.revenue)} }},"
        for r in dom_g.itertuples() if r.날짜 not in existing)
    ovs_new = "".join(
        f"\n  {{ date: '{r.날짜}', orders: {r.orders}, revenue_usd: {round(r.revenue_usd,2)}, revenue_krw: {int(r.revenue_krw)} }},"
        for r in ovs_g.itertuples() if r.날짜 not in existing)

    if dom_new:
        c = re.sub(r"(const DOMESTIC\s*=\s*\[)(.*?)(\];)",
                   lambda m: m.group(1)+m.group(2)+dom_new+m.group(3), c, flags=re.DOTALL)
    if ovs_new:
        c = re.sub(r"(const OVERSEAS\s*=\s*\[)(.*?)(\];)",
                   lambda m: m.group(1)+m.group(2)+ovs_new+m.group(3), c, flags=re.DOTALL)

    # 업데이트 날짜 텍스트
    c = re.sub(r"최근 업데이트 \d{4}-\d{2}-\d{2}", f"최근 업데이트 {max_date_global}", c)

    # YoY 카드는 product 페이지의 renderYoYCards()가 DOMESTIC/OVERSEAS 배열로 클라이언트에서 직접 계산함 (정적 갱신 불필요)

    with open(path,"w") as f: f.write(c)

def _update_insights(ovs_prod_cc, ovs_ok):
    path = os.path.join(BASE_DIR, "insights.html")
    with open(path) as f: c = f.read()

    existing_d = set(re.findall(r'"d":"(\d{4}-\d{2}-\d{2})"', c[:c.find("const OVS_RAW_ORDERS")]))

    # OVS_DAILY {d, cc, n, q}
    new_daily = [
        f'{{"d":"{r["날짜"]}","cc":"{r["cc"]}","n":"{r["상품 이름"]}","q":{int(r["수량"])}}}'
        for _, r in ovs_prod_cc.iterrows() if r["날짜"] not in existing_d
    ]
    if new_daily:
        c = re.sub(r"(const OVS_DAILY\s*=\s*\[)(.*?)(\];)",
                   lambda m: m.group(1)+m.group(2)+","+",".join(new_daily)+m.group(3), c, flags=re.DOTALL)

    existing_r = set(re.findall(r'"d":"(\d{4}-\d{2}-\d{2})"',
                                c[c.find("const OVS_RAW_ORDERS"):c.find("const OVS_RAW_ORDERS")+500000]))
    new_raw = []
    for _, r in ovs_ok.iterrows():
        if r["날짜"] in existing_r: continue
        u = float(r["결제 금액"])
        u_str = f"{u:.2f}" if u != int(u) else f"{u:.1f}"
        q = int(r.get("주문 품목 개수",1)) if pd.notna(r.get("주문 품목 개수")) else 1
        new_raw.append(f'{{"d":"{r["날짜"]}","cc":"{r["cc"]}","u":{u_str},"q":{q}}}')
    if new_raw:
        c = re.sub(r"(const OVS_RAW_ORDERS\s*=\s*\[)(.*?)(\];)",
                   lambda m: m.group(1)+m.group(2)+","+",".join(new_raw)+m.group(3), c, flags=re.DOTALL)

    # I_MAX fallback 갱신
    c = re.sub(r"(I_MAX=\(\(\)=>\{.*?return ds\.length\?ds\[ds\.length-1\]:\')[\d-]+(\';\}\)\(\))",
               rf"\g<1>{max_date_global}\g<2>", c)

    # hsub 초기 문구 (JS가 로드 시 덮어쓰지만 순간적으로 stale 값이 보이는 것 방지)
    try:
        d01 = pd.read_csv(os.path.join(BASE_DIR, "01_일별매출현황.csv"))
        dom_cnt = int(d01["국내주문건수"].sum()); ovs_cnt = int(d01["해외주문건수"].sum())
        c = re.sub(r'(id="hsub">전체 · 국내 )[\d,]+(건 · 해외 )[\d,]+(건)',
                   rf"\g<1>{dom_cnt:,}\g<2>{ovs_cnt:,}\g<3>", c)
    except Exception as e:
        warn(f"insights hsub 갱신 실패: {e}")

    with open(path,"w") as f: f.write(c)

def _update_product_report(dom_prod, ovs_prod, df_products, season_map, max_date):
    path = os.path.join(BASE_DIR, "product_report.html")
    with open(path) as f: c = f.read()

    # PRODUCTS 배열 전체 재생성
    rows_json = ",".join(json.dumps({
        "상품명":str(r["상품명"]),"채널구분":str(r["채널구분"]),
        "통합판매수량":str(r["통합판매수량"]),"국내판매수량":str(r["국내판매수량"]),
        "해외판매수량":str(r["해외판매수량"]),"통합매출_KRW":str(r["통합매출_KRW"]),
        "국내매출_KRW":str(r["국내매출_KRW"]),"해외매출_KRW환산":str(r["해외매출_KRW환산"]),
        "평균단가_KRW":str(int(float(r["평균단가_KRW"]))) if pd.notna(r["평균단가_KRW"]) else "0",
        "총주문건수":str(r["총주문건수"]),"국내주문건수":str(r["국내주문건수"]),
        "해외주문건수":str(r["해외주문건수"]),"최초판매일":str(r["최초판매일"]),
        "시즌":str(r["시즌"]),"통합주문건수":str(r["통합주문건수"]),
    }, ensure_ascii=False) for _, r in df_products.iterrows())
    c = re.sub(r"(const PRODUCTS\s*=\s*\[)(.*?)(\];)",
               lambda m: m.group(1)+rows_json+m.group(3), c, flags=re.DOTALL)

    # DAILY 배열 — 상품별 일자 데이터 추가
    dom_prod["날짜"] = pd.to_datetime(dom_prod["주문 일자"], errors="coerce").dt.date.astype(str)
    ovs_prod["날짜"] = pd.to_datetime(ovs_prod["주문 일자"], errors="coerce").dt.date.astype(str)
    ovs_prod["상품별_KRW"] = ovs_prod["상품별 결제 금액"].fillna(0) * ovs_prod["날짜"].apply(fx_rate)

    existing_daily_dates = {}  # product -> set of dates already in DAILY
    m_daily = re.search(r"const DAILY\s*=\s*(\{.*?\});", c, re.DOTALL)
    if m_daily:
        try:
            daily_obj = json.loads(m_daily.group(1))
            for pname, pdata in daily_obj.items():
                existing_daily_dates[pname] = set(pdata.get("d",[]))
        except: pass

    all_products = set(dom_prod["상품 이름"].tolist() + ovs_prod["상품 이름"].tolist())
    daily_updates = {}
    for pname in all_products:
        d_rows = dom_prod[dom_prod["상품 이름"]==pname]
        o_rows = ovs_prod[ovs_prod["상품 이름"]==pname]
        all_dates = sorted(set(d_rows["날짜"].tolist()+o_rows["날짜"].tolist()))
        existing = existing_daily_dates.get(pname, set())
        for d in all_dates:
            if d in existing: continue
            dr = d_rows[d_rows["날짜"]==d]
            orr = o_rows[o_rows["날짜"]==d]
            dq = int(dr["수량"].sum()) if len(dr) else 0
            oq = int(orr["수량"].sum()) if len(orr) else 0
            drv = int(dr["상품별 결제 금액"].sum()) if len(dr) else 0
            orv = int(orr["상품별_KRW"].sum()) if len(orr) else 0
            do_ = int(dr["주문번호"].nunique()) if len(dr) else 0
            oo_ = int(orr["주문번호"].nunique()) if len(orr) else 0
            if pname not in daily_updates:
                daily_updates[pname] = {"d":[],"dq":[],"oq":[],"dr":[],"or":[],"do":[],"oo":[]}
            daily_updates[pname]["d"].append(d)
            daily_updates[pname]["dq"].append(dq); daily_updates[pname]["oq"].append(oq)
            daily_updates[pname]["dr"].append(drv); daily_updates[pname]["or"].append(orv)
            daily_updates[pname]["do"].append(do_); daily_updates[pname]["oo"].append(oo_)

    # DAILY 배열에 새 데이터 추가
    # 2026-07-28 버그 수정: 아래 re.sub는 DAILY에 해당 상품 키가 "이미 존재할 때"만
    # 매치되어 값을 채워넣는다. 신규 출시 상품(그 주에 처음 판매된 상품)은 키 자체가
    # 없어서 매치가 안 되고 re.sub가 조용히 아무 것도 안 한 채 넘어가 — 그 상품은
    # DAILY에서 영원히 누락된다(26개 "26 HS 핫썸머" 상품에서 실제 발생, product_report.html
    # 드릴다운에서 통째로 빠짐). 매치 실패 시 새 키를 직접 추가하도록 수정.
    new_keys_added = []
    for pname, upd in daily_updates.items():
        escaped = re.escape(pname)
        pat = rf'("{escaped}"\s*:\s*\{{)(.*?)(\}}(?=,|\s*\}}))'
        matched = [False]
        def inject_daily(m, upd=upd):
            matched[0] = True
            inner = m.group(2)
            for key in ["d","dq","oq","dr","or","do","oo"]:
                new_vals = json.dumps(upd[key], ensure_ascii=False)[1:-1]  # [x,y] → x,y
                inner = re.sub(rf'("{key}"\s*:\s*\[)(.*?)(\])',
                               lambda mm, nv=new_vals: mm.group(1)+mm.group(2)+","+nv+mm.group(3),
                               inner, flags=re.DOTALL)
            return m.group(1)+inner+m.group(3)
        c = re.sub(pat, inject_daily, c, flags=re.DOTALL)
        if not matched[0]:
            # 신규 상품: DAILY 객체 맨 앞에 새 키로 직접 삽입
            entry_json = json.dumps({k: upd[k] for k in ["d","dq","oq","dr","or","do","oo"]}, ensure_ascii=False)
            m_daily_start = re.search(r'const DAILY\s*=\s*\{', c)
            insert_at = m_daily_start.end()
            c = c[:insert_at] + json.dumps(pname, ensure_ascii=False) + ":" + entry_json + "," + c[insert_at:]
            new_keys_added.append(pname)
    if new_keys_added:
        ok(f"DAILY에 신규 상품 키 {len(new_keys_added)}개 추가: {', '.join(new_keys_added[:5])}{' 등' if len(new_keys_added)>5 else ''}")

    # MONTHS 배열 — 새 월 추가
    cur_month = max_date[:7]
    m_months = re.search(r'(const MONTHS\s*=\s*\[)([^\]]*?)(\])', c)
    if m_months:
        months_str = m_months.group(2)
        months_list = [x.strip().strip('"') for x in months_str.split(",") if x.strip()]
        if cur_month not in months_list:
            months_list.append(cur_month)
            months_list.sort()
            new_arr = ", ".join(f'"{x}"' for x in months_list)
            c = c.replace(m_months.group(), m_months.group(1)+new_arr+m_months.group(3))
            log(f"  MONTHS에 {cur_month} 추가")

    # 정적 텍스트 날짜
    c = re.sub(r"2025-01-01 ~ \d{4}-\d{2}-\d{2}", f"2025-01-01 ~ {max_date}", c)
    # 하단 푸터 (2026-07-28 이전엔 누락되어 SUM.period_end와 함께 stale 되던 버그 수정)
    c = re.sub(r"📅 업데이트 \d{4}-\d{2}-\d{2}", f"📅 업데이트 {max_date}", c)

    # MONTHLY 재계산 (DAILY로부터) — 월별 상세 정확도 보장
    c = _recompute_monthly_from_daily(c)

    # SUM 오브젝트 재계산 (01/02 CSV 기준 — 2026-07-28 이전엔 이 스텝이 아예 없어서 SUM.period_end가 오래 stale된 채 방치됐음)
    c = _recompute_sum(c, df_products, max_date)

    with open(path,"w") as f: f.write(c)

def _recompute_sum(c, df_products, max_date):
    d01 = pd.read_csv(os.path.join(BASE_DIR, "01_일별매출현황.csv"))
    os_usd = float(d01["해외매출_USD"].sum())
    os_krw = float(d01["해외매출_KRW환산"].sum())
    sum_obj = {
        "period_start": MIN_DATE, "period_end": max_date,
        "total_revenue_krw": float(d01["통합매출_KRW"].sum()),
        "total_orders": int(d01["통합주문건수"].sum()),
        "total_qty": int(df_products["통합판매수량"].sum()),
        "dom_revenue_krw": float(d01["국내매출_KRW"].sum()),
        "dom_orders": int(d01["국내주문건수"].sum()),
        "dom_qty": int(df_products["국내판매수량"].sum()),
        "os_revenue_usd": round(os_usd, 2), "os_revenue_krw": os_krw,
        "os_orders": int(d01["해외주문건수"].sum()),
        "os_qty": int(df_products["해외판매수량"].sum()),
        "num_products": int(len(df_products)),
        "usd_to_krw": round(os_krw / os_usd) if os_usd else USD_TO_KRW,
    }
    m = re.search(r"const SUM\s*=\s*\{.*?\};", c, re.DOTALL)
    if not m:
        warn("SUM 오브젝트를 찾지 못해 재계산 건너뜀")
        return c
    new_sum = "const SUM = " + json.dumps(sum_obj, ensure_ascii=False, separators=(',', ':')) + ";"
    return c[:m.start()] + new_sum + c[m.end():]

def _recompute_monthly_from_daily(c):
    """DAILY 배열을 월별로 집계해 MONTHLY 오브젝트 전체를 재생성"""
    idx_d = c.find('const DAILY = ')
    idx_m = c.find('const MONTHLY = ', idx_d)
    if idx_d < 0 or idx_m < 0:
        warn("MONTHLY 재계산 실패: DAILY/MONTHLY 위치 못 찾음")
        return c
    # DAILY 파싱 (bracket-depth로 정확한 끝 위치 탐색)
    arr_start = idx_d + len('const DAILY = ')
    depth = in_str = esc = 0
    end_d = None
    for i in range(arr_start, len(c)):
        ch = c[i]
        if in_str:
            esc = (ch == '\\' and not esc)
            if ch == '"' and not esc: in_str = False
            continue
        if ch == '"': in_str = True
        elif ch in '{[': depth += 1
        elif ch in '}]':
            depth -= 1
            if depth == 0: end_d = i+1; break
    if end_d is None:
        warn("MONTHLY 재계산 실패: DAILY 끝 못 찾음")
        return c
    try:
        daily = json.loads(c[arr_start:end_d])
    except Exception as e:
        warn(f"MONTHLY 재계산 실패: DAILY JSON 파싱 오류 — {e}")
        return c

    monthly = {}
    for pname, d in daily.items():
        by_m = {}
        for i, date in enumerate(d.get('d', [])):
            m = date[:7]
            if m not in by_m:
                by_m[m] = [0, 0, 0, 0]
            by_m[m][0] += d['dq'][i] if i < len(d.get('dq',[])) else 0
            by_m[m][1] += d['oq'][i] if i < len(d.get('oq',[])) else 0
            by_m[m][2] += d['dr'][i] if i < len(d.get('dr',[])) else 0
            by_m[m][3] += d['or'][i] if i < len(d.get('or',[])) else 0
        ms = sorted(by_m)
        monthly[pname] = {
            'months': ms,
            'dom_qty': [by_m[m][0] for m in ms],
            'os_qty':  [by_m[m][1] for m in ms],
            'dom_rev': [by_m[m][2] for m in ms],
            'os_rev':  [by_m[m][3] for m in ms],
        }

    # MONTHLY 오브젝트 교체
    m_json = json.dumps(monthly, ensure_ascii=False, separators=(',', ':'))
    # 끝 위치 탐색
    arr_start_m = idx_m + len('const MONTHLY = ')
    depth = in_str = esc = 0
    end_m = None
    for i in range(arr_start_m, len(c)):
        ch = c[i]
        if in_str:
            esc = (ch == '\\' and not esc)
            if ch == '"' and not esc: in_str = False
            continue
        if ch == '"': in_str = True
        elif ch in '{[': depth += 1
        elif ch in '}]':
            depth -= 1
            if depth == 0: end_m = i+1; break
    if end_m is None:
        warn("MONTHLY 재계산 실패: MONTHLY 끝 못 찾음")
        return c

    c = c[:arr_start_m] + m_json + c[end_m:]
    log(f"  MONTHLY 재계산 완료 ({len(monthly)}개 상품)")
    return c

def _update_product_analytics(dom_prod, ovs_prod, season_map):
    path = os.path.join(BASE_DIR, "product_analytics.html")
    with open(path) as f: c = f.read()

    dom_prod = dom_prod.copy()
    ovs_prod = ovs_prod.copy()
    dom_prod["날짜"] = pd.to_datetime(dom_prod["주문 일자"], errors="coerce").dt.date.astype(str)
    ovs_prod["날짜"] = pd.to_datetime(ovs_prod["주문 일자"], errors="coerce").dt.date.astype(str)
    ovs_prod["상품별_KRW"] = ovs_prod["상품별 결제 금액"].fillna(0) * ovs_prod["날짜"].apply(fx_rate)

    # 이미 있는 날짜
    existing_dates = set(re.findall(r'"d":"(\d{4}-\d{2}-\d{2})"', c))

    pa_rows = []
    dom_pa = dom_prod.groupby(["날짜","상품 이름"]).agg(
        dq=("수량","sum"), dr=("상품별 결제 금액","sum"), do=("주문번호","nunique")).reset_index()
    for _, r in dom_pa.iterrows():
        if r["날짜"] in existing_dates: continue
        nm = r["상품 이름"]; se = season_for_new_product(nm, season_map)
        pa_rows.append({"d":r["날짜"],"s":"dom","n":str(nm),"se":se,
            "dq":int(r["dq"]),"oq":0,"dr":int(r["dr"]),"or":0,"do":int(r["do"]),"oo":0})

    ovs_pa = ovs_prod.groupby(["날짜","상품 이름"]).agg(
        oq=("수량","sum"), or_krw=("상품별_KRW","sum"), oo=("주문번호","nunique")).reset_index()
    for _, r in ovs_pa.iterrows():
        if r["날짜"] in existing_dates: continue
        nm = r["상품 이름"]; se = season_for_new_product(nm, season_map)
        found = [x for x in pa_rows if x["d"]==r["날짜"] and x["n"]==nm]
        if found:
            found[0]["oq"]=int(r["oq"]); found[0]["or"]=int(round(r["or_krw"])); found[0]["oo"]=int(r["oo"])
        else:
            pa_rows.append({"d":r["날짜"],"s":"ovs","n":str(nm),"se":se,
                "dq":0,"oq":int(r["oq"]),"dr":0,"or":int(round(r["or_krw"])),"do":0,"oo":int(r["oo"])})

    if pa_rows:
        pa_rows.sort(key=lambda x: x["d"])
        new_str = ","+",".join(json.dumps(r, ensure_ascii=False) for r in pa_rows)
        # bracket-matching: re.sub의 비탐욕(.*?) 매칭이 JSON 내부의 첫 ']}'에서
        # 잘못 멈추는 버그를 방지하기 위해 직접 괄호 깊이를 추적해 rows 배열의
        # 진짜 끝 위치를 찾는다.
        marker = 'const DATA = {"rows":['
        start = c.index(marker)
        arr_start = start + len(marker) - 1  # '[' 위치
        depth = 0
        in_str = False
        esc = False
        end = None
        for i in range(arr_start, len(c)):
            ch = c[i]
            if in_str:
                if esc:
                    esc = False
                elif ch == '\\':
                    esc = True
                elif ch == '"':
                    in_str = False
                continue
            if ch == '"':
                in_str = True
            elif ch == '[':
                depth += 1
            elif ch == ']':
                depth -= 1
                if depth == 0:
                    end = i
                    break
        if end is None:
            err("product_analytics.html: DATA.rows 배열의 끝을 찾지 못함")
        else:
            c = c[:end] + new_str + c[end:]

    # date_max 필드 갱신
    c = re.sub(r'("date_max"\s*:\s*")\d{4}-\d{2}-\d{2}(")', rf"\g<1>{max_date_global}\g<2>", c)

    # 정적 텍스트
    c = re.sub(r"2025-01-01 ~ \d{4}-\d{2}-\d{2}", f"2025-01-01 ~ {max_date_global}", c)

    with open(path,"w") as f: f.write(c)

def _update_date_inputs(max_date):
    """4개 HTML의 date input value/max 속성을 Python으로 직접 교체 (JS 의존 없음)"""
    targets = {
        "sales_dashboard.html": [("startDate","2025-01-01",max_date),("endDate",max_date,max_date)],
        "insights.html":        [("dtFrom","2025-01-01",max_date),   ("dtTo",max_date,max_date)],
        "product_report.html":  [("startDate","2025-01-01",max_date),("endDate",max_date,max_date)],
        "product_analytics.html":[("startDate","2025-01-01",max_date),("endDate",max_date,max_date)],
    }
    for fname, inputs in targets.items():
        path = os.path.join(BASE_DIR, fname)
        with open(path) as f: c = f.read()
        for iid, val, mx in inputs:
            # value 교체 또는 추가
            c = re.sub(rf'(<input\b[^>]*\bid="{iid}"[^>]*?)\s+value="[^"]*"', rf'\1', c)
            c = re.sub(rf'(<input\b[^>]*\bid="{iid}"[^>]*?)\s+max="[^"]*"', rf'\1', c)
            c = re.sub(rf'(<input\b[^>]*\bid="{iid}")', rf'\1 value="{val}" max="{mx}"', c)
        with open(path,"w") as f: f.write(c)

# ── 검증 ─────────────────────────────────────────────────────────────────────
def validate(max_date):
    print("\n[6/6] 검증")

    # 1. CSV 날짜 최신 확인
    df_d = pd.read_csv(os.path.join(BASE_DIR,"01_일별매출현황.csv"))
    assert df_d["날짜"].max() == max_date, f"01 최신일 불일치: {df_d['날짜'].max()} ≠ {max_date}"
    ok(f"01 최신일: {df_d['날짜'].max()}")

    # 2. date input 속성 검증
    targets = {
        "sales_dashboard.html":  [("endDate",max_date,max_date)],
        "insights.html":         [("dtTo",max_date,max_date)],
        "product_report.html":   [("endDate",max_date,max_date)],
        "product_analytics.html":[("endDate",max_date,max_date)],
    }
    for fname, inputs in targets.items():
        with open(os.path.join(BASE_DIR,fname)) as f: c = f.read()
        for iid, val, mx in inputs:
            m = re.search(rf'<input\b[^>]*\bid="{iid}"[^>]*>', c)
            if not m: err(f"{fname} #{iid} input 못 찾음"); continue
            tag = m.group()
            v_m = re.search(r'value="([^"]+)"', tag)
            x_m = re.search(r'max="([^"]+)"', tag)
            v = v_m.group(1) if v_m else "없음"
            x = x_m.group(1) if x_m else "없음"
            if v == val and x == mx:
                ok(f"{fname} #{iid}: value={v}, max={x}")
            else:
                err(f"{fname} #{iid}: value={v}(기대:{val}), max={x}(기대:{mx})")

    # 3. 이중 콤마 검사
    for fname in ["sales_dashboard.html","insights.html","product_report.html","product_analytics.html"]:
        with open(os.path.join(BASE_DIR,fname)) as f: c = f.read()
        doubles = re.findall(r'\},,|\],,', c)
        if doubles:
            err(f"{fname}: 이중 콤마 {len(doubles)}건 발견")
        else:
            ok(f"{fname}: 이중 콤마 없음")

    # 4. 구버전 날짜 잔존 검사 (데이터 레코드 제외)
    prev_date = (pd.to_datetime(max_date) - pd.Timedelta(days=1)).strftime("%Y-%m-%d")
    for fname in ["sales_dashboard.html","insights.html","product_report.html","product_analytics.html"]:
        with open(os.path.join(BASE_DIR,fname)) as f: c = f.read()
        # 정적 영역(value=, max=, 부제 텍스트)에만 체크
        static_parts = re.findall(r'(?:value|max)="[^"]*"', c)
        stale = [x for x in static_parts if prev_date in x and max_date not in x]
        if stale:
            warn(f"{fname}: 이전 날짜 잔존 가능성 — {stale[:2]}")

    # 5. 05/31 T00:00:00 timezone 버그
    for fname in ["sales_dashboard.html","insights.html","product_report.html","product_analytics.html"]:
        with open(os.path.join(BASE_DIR,fname)) as f: c = f.read()
        if "T00:00:00" in c:
            err(f"{fname}: T00:00:00 timezone 버그 발견 → T12:00:00으로 교체 필요")
        else:
            ok(f"{fname}: timezone 안전")

# ── 메인 ─────────────────────────────────────────────────────────────────────
def main():
    global max_date_global

    if len(sys.argv) < 5:
        print("사용법: python3 update_data.py 파일1.xlsx 파일2.xlsx 파일3.xlsx 파일4.xlsx")
        print("  (주문별_국내, 주문별_해외, 상품별_국내, 상품별_해외 순서 무관)")
        sys.exit(1)

    paths = sys.argv[1:5]
    print("=" * 60)
    print("  매출 데이터 자동 업데이트")
    print(f"  실행 시각: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("=" * 60)

    # 파일 분류
    print("\n[0/6] 파일 자동 판별")
    files = classify_files(paths)

    required = ["order_dom","order_ovs","product_dom","product_ovs"]
    for k in required:
        if k not in files:
            err(f"필수 파일 없음: {k}"); sys.exit(1)

    dom_ord  = load_and_clean(files["order_dom"])
    ovs_ord  = load_and_clean(files["order_ovs"])
    dom_prod = load_and_clean(files["product_dom"])
    ovs_prod = load_and_clean(files["product_ovs"])

    # 업로드 파일 날짜 범위
    all_dates = pd.concat([dom_ord["날짜"], ovs_ord["날짜"]]).dropna()
    max_date_global = all_dates.max()
    min_date_upload = all_dates.min()
    log(f"업로드 기간: {min_date_upload} ~ {max_date_global}")

    # 기존 최신일
    df_daily_cur = pd.read_csv(os.path.join(BASE_DIR,"01_일별매출현황.csv"))
    prev_max = df_daily_cur["날짜"].max()
    log(f"기존 최신일: {prev_max}")

    if max_date_global <= prev_max:
        warn(f"업로드 데이터({max_date_global})가 기존 데이터({prev_max})보다 새롭지 않습니다.")

    max_date = max_date_global  # 문자열

    # 순차 실행
    update_daily(dom_ord, ovs_ord, prev_max)
    df_products = update_products(dom_prod, ovs_prod, dom_ord, ovs_ord)
    update_raw(dom_ord, ovs_ord, prev_max)
    update_overseas_db(ovs_ord, ovs_prod, prev_max)
    update_html(dom_ord, ovs_ord, dom_prod, ovs_prod, df_products, max_date)
    validate(max_date)

    print("\n" + "=" * 60)
    if errors:
        print(f"  완료 (오류 {len(errors)}건, 경고 {len(warnings)}건)")
        for e in errors: print(f"    ❌ {e}")
    elif warnings:
        print(f"  완료 ✅ (경고 {len(warnings)}건)")
        for w in warnings: print(f"    ⚠️  {w}")
    else:
        print(f"  완료 ✅ 오류 없음")
    print(f"  데이터 기준일: {max_date}")
    print("  다음 단계: git add . && git commit -m '데이터 업데이트: {max_date}' && git push")
    print("=" * 60)

if __name__ == "__main__":
    main()
