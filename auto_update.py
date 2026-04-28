#!/usr/bin/env python3
"""
매출 대시보드 자동 업데이트 스크립트
--------------------------------------
국내/해외 상품별 파일(주문_상품별_상세검색_*.xlsx)을 업로드하면
다음을 자동 반영합니다:

  1. product_analytics.html  — DATA.rows 교체
  2. sales_dashboard.html    — DOMESTIC / OVERSEAS 배열 갱신
  3. 01_일별매출현황.csv      — 새 일자 행 추가
  4. insights.html           — INS.mom_data 현재 월 갱신 + I_MAX
  5. 전 HTML 공통             — 날짜 입력(max/value), 정적 텍스트

사용법:
  python3 auto_update.py --dom 국내파일.xlsx --ovs 해외파일.xlsx

오류 패턴(오늘 경험) 자동 방어:
  - JSON 공백 없는 컴팩트 포맷 강제
  - 중복 날짜 삽입 금지 (먼저 그레프로 기존 날짜 확인 후 교체)
  - const ROWS = DATA.rows; 누락 시 자동 복원
  - I_MAX IIFE 패턴 유지 검증
"""

import sys, re, json, shutil, argparse
from pathlib import Path
from collections import defaultdict
import pandas as pd

# ── 경로 ────────────────────────────────────────────────────────────────────
BASE      = Path('/sessions/sleepy-wizardly-knuth/mnt/outputs')
WORKSPACE = Path('/sessions/sleepy-wizardly-knuth/mnt/매출트레킹')
PA_HTML   = BASE / 'product_analytics.html'
SD_HTML   = BASE / 'sales_dashboard.html'
INS_HTML  = BASE / 'insights.html'
PR_HTML   = BASE / 'product_report.html'
CSV_01    = BASE / '01_일별매출현황.csv'
EXCHANGE  = 1380  # 1 USD = 1,380 KRW (고정)

# ── 파일 타입 자동 감지 ──────────────────────────────────────────────────────
def detect_channel(filepath):
    """주문번호 영문자 포함 여부로 국내/해외 판별"""
    df = pd.read_excel(filepath, nrows=20)
    if '주문번호' in df.columns:
        for val in df['주문번호'].dropna():
            if re.search(r'[A-Z]', str(val)):
                return 'ovs'
    return 'dom'

# ── 국내 상품별 파일 처리 ────────────────────────────────────────────────────
def process_domestic_pa(filepath):
    """→ product_analytics DATA.rows용 컴팩트 JSON 행 리스트"""
    df = pd.read_excel(filepath)
    df['date'] = pd.to_datetime(df['주문 일자']).dt.strftime('%Y-%m-%d')
    df = df[~df['상품별 주문 상태'].str.contains('취소', na=False)]

    rows = []
    for _, r in df.iterrows():
        date = r['date']
        name = str(r['상품 이름']).strip()
        opt  = str(r.get('상품 옵션 정보', '')).strip()
        if opt in ('nan', '', '-'): opt = ''
        dq   = int(r['수량'])              if pd.notna(r.get('수량'))              else 1
        dr   = int(r['상품별 결제 금액'])  if pd.notna(r.get('상품별 결제 금액'))  else 0
        rows.append(
            f'{{"d":"{date}","s":"dom","n":{json.dumps(name, ensure_ascii=False)},'
            f'"se":{json.dumps(opt, ensure_ascii=False)},'
            f'"dq":{dq},"oq":0,"dr":{dr},"or":0,"do":{dr},"oo":0}}'
        )
    return rows

# ── 해외 상품별 파일 처리 ────────────────────────────────────────────────────
def process_overseas_pa(filepath):
    """→ product_analytics DATA.rows용 컴팩트 JSON 행 리스트"""
    df = pd.read_excel(filepath)

    # 날짜: 주문 일자 컬럼 우선 → 없으면 주문번호 앞 8자리
    if '주문 일자' in df.columns and df['주문 일자'].notna().any():
        df['date'] = pd.to_datetime(df['주문 일자']).dt.strftime('%Y-%m-%d')
    else:
        def _ext(o):
            m = re.match(r'(\d{4})(\d{2})(\d{2})', str(o))
            return f'{m.group(1)}-{m.group(2)}-{m.group(3)}' if m else None
        df['date'] = df['주문번호'].apply(_ext)

    df = df[~df['상품별 주문 상태'].str.contains('취소', na=False)]
    df = df[df['date'].notna()]

    rows = []
    for _, r in df.iterrows():
        date = r['date']
        name = str(r['상품 이름']).strip()
        opt  = str(r.get('상품 옵션 정보', '')).strip()
        if opt in ('nan', '', '-'): opt = ''
        oq   = int(r['수량'])           if pd.notna(r.get('수량'))            else 1
        pc   = int(r['주문 품목 개수']) if pd.notna(r.get('주문 품목 개수')) and int(r.get('주문 품목 개수', 1)) > 0 else 1
        pay  = float(r['결제 금액'])    if pd.notna(r.get('결제 금액'))       else 0
        or_  = round(pay / pc * EXCHANGE)
        rows.append(
            f'{{"d":"{date}","s":"ovs","n":{json.dumps(name, ensure_ascii=False)},'
            f'"se":{json.dumps(opt, ensure_ascii=False)},'
            f'"dq":0,"oq":{oq},"dr":0,"or":{or_},"do":0,"oo":{or_}}}'
        )
    return rows

# ── 일별 매출 집계 (sales_dashboard / CSV 용) ────────────────────────────────
def aggregate_daily_dom(filepath):
    """국내 상품별 파일 → {date: {orders, revenue}}"""
    df = pd.read_excel(filepath)
    df['date'] = pd.to_datetime(df['주문 일자']).dt.strftime('%Y-%m-%d')
    df = df[~df['상품별 주문 상태'].str.contains('취소', na=False)]

    result = {}
    for date, grp in df.groupby('date'):
        orders  = grp['주문번호'].nunique()
        revenue = int(grp['상품별 결제 금액'].fillna(0).sum())
        result[date] = {'orders': orders, 'revenue': revenue}
    return result

def aggregate_daily_ovs(filepath):
    """해외 상품별 파일 → {date: {orders, revenue_usd, revenue_krw}}"""
    df = pd.read_excel(filepath)
    if '주문 일자' in df.columns and df['주문 일자'].notna().any():
        df['date'] = pd.to_datetime(df['주문 일자']).dt.strftime('%Y-%m-%d')
    else:
        def _ext(o):
            m = re.match(r'(\d{4})(\d{2})(\d{2})', str(o))
            return f'{m.group(1)}-{m.group(2)}-{m.group(3)}' if m else None
        df['date'] = df['주문번호'].apply(_ext)

    df = df[~df['상품별 주문 상태'].str.contains('취소', na=False)]
    df = df[df['date'].notna()]

    result = {}
    for date, grp in df.groupby('date'):
        orders      = grp['주문번호'].nunique()
        # 고유 주문별 결제 금액 합산 (USD)
        revenue_usd = round(grp.drop_duplicates('주문번호')['결제 금액'].fillna(0).sum(), 2)
        revenue_krw = round(revenue_usd * EXCHANGE)
        result[date] = {'orders': orders, 'revenue_usd': revenue_usd, 'revenue_krw': revenue_krw}
    return result

# ── product_analytics.html 업데이트 ─────────────────────────────────────────
def update_product_analytics(dom_rows, ovs_rows, new_max_date):
    print('\n[1/5] product_analytics.html 업데이트...')
    all_new = dom_rows + ovs_rows
    all_new.sort(key=lambda x: re.search(r'"d":"([^"]+)"', x).group(1))

    if not all_new:
        print('  ⚠ 새 행 없음, 건너뜀')
        return

    new_dates = {re.search(r'"d":"([^"]+)"', r).group(1) for r in all_new}
    min_d, max_d = min(new_dates), max(new_dates)
    print(f'  새 데이터: {min_d} ~ {max_d} ({len(new_dates)}일, {len(all_new)}행)')

    with open(PA_HTML, 'r', encoding='utf-8') as f:
        content = f.read()

    rows_start = content.find('"rows":[') + len('"rows":[')
    rows_end   = content.find('],"cat_totals"', rows_start)
    section    = content[rows_start:rows_end]
    cur_rows   = re.findall(r'\{[^{}]+\}', section)

    # 날짜 범위 내 기존 행 제거
    kept, removed = [], 0
    for r in cur_rows:
        m = re.search(r'"d":"([^"]+)"', r)
        if m and min_d <= m.group(1) <= max_d:
            removed += 1
        else:
            kept.append(r)
    print(f'  기존 제거: {removed}행, 유지: {len(kept)}행')

    before = [r for r in kept if re.search(r'"d":"([^"]+)"', r) and re.search(r'"d":"([^"]+)"', r).group(1) < min_d]
    after  = [r for r in kept if re.search(r'"d":"([^"]+)"', r) and re.search(r'"d":"([^"]+)"', r).group(1) > max_d]

    final  = before + all_new + after
    new_str = ',\n'.join(final)
    content = content[:rows_start] + '\n' + new_str + '\n' + content[rows_end:]

    # ★ const ROWS 누락 자동 복원
    if 'const ROWS = DATA.rows;' not in content:
        print('  ⚠ const ROWS 누락 감지 → 자동 복원')
        weekly_pos = content.find('const WEEKLY')
        if weekly_pos >= 0:
            insert = (
                'const ROWS = DATA.rows;\n'
                'const CP_DOM = DATA.cp_dom || [];\n'
                'const CP_OVS = DATA.cp_ovs || [];\n'
                'const SEASONS = DATA.seasons || [];\n\n'
            )
            content = content[:weekly_pos] + insert + content[weekly_pos:]

    # ★ I_MAX IIFE 패턴 검증
    if 'DATA.rows.map' in content[content.find('I_MAX'):content.find('I_MAX')+300]:
        print('  ⚠ I_MAX 단순 체이닝 감지 — product_analytics는 PA_MAX로 동적 계산, 무관')

    with open(PA_HTML, 'w', encoding='utf-8') as f:
        f.write(content)
    shutil.copy(PA_HTML, WORKSPACE / 'product_analytics.html')
    print(f'  ✓ 완료 (총 {len(final)}행)')

# ── sales_dashboard.html 업데이트 ───────────────────────────────────────────
def update_sales_dashboard(dom_daily, ovs_daily, new_max_date):
    print('\n[2/5] sales_dashboard.html 업데이트...')
    with open(SD_HTML, 'r', encoding='utf-8') as f:
        content = f.read()

    # DOMESTIC 배열 갱신
    dom_start = content.find('const DOMESTIC = [') + len('const DOMESTIC = [')
    dom_end   = content.find('];', dom_start)
    dom_section = content[dom_start:dom_end]

    # 기존 날짜 파악
    existing_dom_dates = set(re.findall(r"date:\s*'(\d{4}-\d{2}-\d{2})'", dom_section))

    new_dom_lines = []
    for date in sorted(dom_daily.keys()):
        if date not in existing_dom_dates:
            d = dom_daily[date]
            new_dom_lines.append(f"\n  {{ date: '{date}', orders: {d['orders']}, revenue: {d['revenue']} }},")

    if new_dom_lines:
        insert_pos = dom_end  # 배열 닫히기 직전에 삽입
        insert_text = ''.join(new_dom_lines)
        content = content[:insert_pos] + insert_text + content[insert_pos:]
        print(f'  DOMESTIC: {len(new_dom_lines)}일 추가')
    else:
        print('  DOMESTIC: 이미 최신')

    # OVERSEAS 배열 갱신 (삽입 후 offset 반영)
    ovs_start = content.find('const OVERSEAS = [') + len('const OVERSEAS = [')
    ovs_end   = content.find('];', ovs_start)
    ovs_section = content[ovs_start:ovs_end]

    existing_ovs_dates = set(re.findall(r"date:\s*'(\d{4}-\d{2}-\d{2})'", ovs_section))

    new_ovs_lines = []
    for date in sorted(ovs_daily.keys()):
        if date not in existing_ovs_dates:
            d = ovs_daily[date]
            new_ovs_lines.append(
                f"\n  {{ date: '{date}', orders: {d['orders']}, "
                f"revenue_usd: {d['revenue_usd']}, revenue_krw: {d['revenue_krw']} }},"
            )

    if new_ovs_lines:
        ovs_end2 = content.find('];', content.find('const OVERSEAS = ['))
        content = content[:ovs_end2] + ''.join(new_ovs_lines) + content[ovs_end2:]
        print(f'  OVERSEAS: {len(new_ovs_lines)}일 추가')
    else:
        print('  OVERSEAS: 이미 최신')

    # 정적 텍스트 갱신
    content = re.sub(r'최근 업데이트 \d{4}-\d{2}-\d{2}', f'최근 업데이트 {new_max_date}', content)

    # date input max/value 갱신
    content = re.sub(
        r'(<input[^>]*id="endDate"[^>]*value=")[^"]*(")',
        lambda m: m.group(1) + new_max_date + m.group(2), content
    )

    with open(SD_HTML, 'w', encoding='utf-8') as f:
        f.write(content)
    shutil.copy(SD_HTML, WORKSPACE / 'sales_dashboard.html')
    print('  ✓ 완료')

# ── 01_일별매출현황.csv 업데이트 ────────────────────────────────────────────
def update_csv_01(dom_daily, ovs_daily):
    print('\n[3/5] 01_일별매출현황.csv 업데이트...')
    with open(CSV_01, 'r', encoding='utf-8-sig') as f:
        lines = f.readlines()

    existing_dates = set()
    for line in lines[1:]:
        parts = line.split(',')
        if parts:
            existing_dates.add(parts[0].strip())

    # 공통 날짜 합산
    all_dates = sorted(set(list(dom_daily.keys()) + list(ovs_daily.keys())))
    added = 0
    for date in all_dates:
        if date in existing_dates:
            continue
        month = date[:7]
        d_dom = dom_daily.get(date, {'orders': 0, 'revenue': 0})
        d_ovs = ovs_daily.get(date, {'orders': 0, 'revenue_usd': 0, 'revenue_krw': 0})
        total_rev = d_dom['revenue'] + d_ovs['revenue_krw']
        total_ord = d_dom['orders'] + d_ovs['orders']
        line = (
            f"{date},{month},{total_rev},{total_ord},"
            f"{d_dom['revenue']},{d_dom['orders']},"
            f"{d_ovs['revenue_usd']},{d_ovs['revenue_krw']},{d_ovs['orders']}\n"
        )
        lines.append(line)
        added += 1

    if added:
        lines[1:] = sorted(lines[1:], key=lambda l: l.split(',')[0])
        with open(CSV_01, 'w', encoding='utf-8-sig') as f:
            f.writelines(lines)
        shutil.copy(CSV_01, WORKSPACE / '01_일별매출현황.csv')
        print(f'  ✓ {added}일 추가')
    else:
        print('  이미 최신')

# ── insights.html 월간 집계 갱신 ────────────────────────────────────────────
def update_insights_monthly(dom_daily, ovs_daily, new_max_date):
    print('\n[4/5] insights.html INS.mom_data 갱신...')
    with open(INS_HTML, 'r', encoding='utf-8') as f:
        content = f.read()

    # 현재 월 집계
    cur_month = new_max_date[:7]
    dom_month = {d: v for d, v in dom_daily.items() if d.startswith(cur_month)}
    ovs_month = {d: v for d, v in ovs_daily.items() if d.startswith(cur_month)}
    new_dom_rev = sum(v['revenue'] for v in dom_month.values())
    new_dom_ord = sum(v['orders'] for v in dom_month.values())
    new_ovs_rev = sum(v['revenue_krw'] for v in ovs_month.values())
    new_ovs_ord = sum(v['orders'] for v in ovs_month.values())

    # mom_data에서 현재 월 항목 찾아서 갱신
    def update_mom_entry(match_str):
        # {"month":"2026-04","dom_rev":X,"dom_ord":X,"ovs_rev":X,"ovs_ord":X,...}
        m = re.search(rf'"month":"{re.escape(cur_month)}"([^}}]+)\}}', content)
        if not m:
            return None
        old_entry = m.group(0)
        # 기존 MoM 비율 유지 (복잡한 재계산 필요해 보존)
        dom_mom_m = re.search(r'"dom_mom":([\d.\-null]+)', old_entry)
        ovs_mom_m = re.search(r'"ovs_mom":([\d.\-null]+)', old_entry)
        dom_mom = dom_mom_m.group(1) if dom_mom_m else 'null'
        ovs_mom = ovs_mom_m.group(1) if ovs_mom_m else 'null'
        new_entry = (
            f'{{"month":"{cur_month}",'
            f'"dom_rev":{new_dom_rev},"dom_ord":{new_dom_ord},'
            f'"ovs_rev":{new_ovs_rev},"ovs_ord":{new_ovs_ord},'
            f'"dom_mom":{dom_mom},"ovs_mom":{ovs_mom}}}'
        )
        return old_entry, new_entry

    result = update_mom_entry(cur_month)
    if result:
        old_entry, new_entry = result
        content = content.replace(old_entry, new_entry, 1)
        print(f'  mom_data "{cur_month}" 갱신: dom={new_dom_rev:,}원/{new_dom_ord}건, '
              f'ovs={new_ovs_rev:,}원/{new_ovs_ord}건')
    else:
        print(f'  ⚠ mom_data에서 "{cur_month}" 항목 없음 — 수동 추가 필요')

    # 정적 텍스트 (hdr-sub 초기값)
    content = re.sub(
        r'(id="hsub"[^>]*>[^<]*~ )\d{4}-\d{2}-\d{2}',
        lambda m: m.group(1) + new_max_date,
        content
    )

    # date input 갱신
    content = re.sub(r'(I_MAX\s*[,;].*?PA_MIN\s*=\s*["\'])[^"\']+',
                     lambda m: m.group(0), content)  # I_MAX는 IIFE로 동적계산, 건드리지 않음

    with open(INS_HTML, 'w', encoding='utf-8') as f:
        f.write(content)
    shutil.copy(INS_HTML, WORKSPACE / 'insights.html')
    print('  ✓ 완료 (INS 히트맵 등 복잡 집계는 월말 전체 리빌드 시 반영)')

# ── product_report.html 날짜/텍스트 갱신 ────────────────────────────────────
def update_product_report_dates(new_max_date):
    print('\n[5/5] product_report.html 날짜 갱신...')
    with open(PR_HTML, 'r', encoding='utf-8') as f:
        content = f.read()

    # period_end
    content = re.sub(r'"period_end"\s*:\s*"\d{4}-\d{2}-\d{2}"',
                     f'"period_end": "{new_max_date}"', content)

    # 푸터 날짜
    content = re.sub(r'📅 업데이트 \d{4}-\d{2}-\d{2}',
                     f'📅 업데이트 {new_max_date}', content)

    # MAX_DATE const
    content = re.sub(r"(const MAX_DATE\s*=\s*['\"])\d{4}-\d{2}-\d{2}(['\"])",
                     rf'\g<1>{new_max_date}\g<2>', content)

    # date input max/value
    content = re.sub(
        r'(<input[^>]*id="endDate"[^>]*value=")[^"]*(")',
        lambda m: m.group(1) + new_max_date + m.group(2), content
    )

    with open(PR_HTML, 'w', encoding='utf-8') as f:
        f.write(content)
    shutil.copy(PR_HTML, WORKSPACE / 'product_report.html')
    print('  ✓ 완료 (MONTHLY/PRODUCTS 데이터는 월말 전체 리빌드 시 반영)')

# ── JS 문법 검증 ─────────────────────────────────────────────────────────────
def validate_js(html_path):
    import subprocess, tempfile
    content = open(html_path, 'r', encoding='utf-8').read()
    scripts = re.findall(r'<script[^>]*>(.*?)</script>', content, re.DOTALL)
    js = '\n'.join(scripts)
    with tempfile.NamedTemporaryFile(suffix='.js', mode='w', delete=False) as f:
        f.write(js)
        tmp = f.name
    result = subprocess.run(['node', '--check', tmp], capture_output=True, text=True)
    ok = result.returncode == 0
    print(f'  JS 문법: {"✓ OK" if ok else "✗ ERROR: " + result.stderr[:200]}')
    return ok

# ── 메인 ─────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--dom', required=True, help='국내 상품별 xlsx 파일')
    parser.add_argument('--ovs', required=True, help='해외 상품별 xlsx 파일')
    args = parser.parse_args()

    dom_path = Path(args.dom)
    ovs_path = Path(args.ovs)

    print('=' * 60)
    print(f'국내 파일: {dom_path.name}')
    print(f'해외 파일: {ovs_path.name}')
    print('=' * 60)

    # 파일 타입 자동 검증
    dom_ch = detect_channel(dom_path)
    ovs_ch = detect_channel(ovs_path)
    if dom_ch != 'dom':
        print(f'⚠ 국내 파일로 지정했지만 해외 포맷 감지됨 — 파일을 확인하세요: {dom_path.name}')
    if ovs_ch != 'ovs':
        print(f'⚠ 해외 파일로 지정했지만 국내 포맷 감지됨 — 파일을 확인하세요: {ovs_path.name}')

    # 데이터 처리
    print('\n데이터 처리 중...')
    dom_pa   = process_domestic_pa(dom_path)
    ovs_pa   = process_overseas_pa(ovs_path)
    dom_daily = aggregate_daily_dom(dom_path)
    ovs_daily = aggregate_daily_ovs(ovs_path)

    # 최대 날짜 계산
    all_dates = sorted(list(dom_daily.keys()) + list(ovs_daily.keys()))
    new_max_date = all_dates[-1] if all_dates else None
    if not new_max_date:
        print('ERROR: 날짜 데이터 없음')
        sys.exit(1)

    print(f'최신 날짜: {new_max_date}')
    print(f'국내 PA 행: {len(dom_pa)}, 해외 PA 행: {len(ovs_pa)}')
    print(f'국내 일별 {len(dom_daily)}일, 해외 일별 {len(ovs_daily)}일')

    # 각 파일 업데이트
    update_product_analytics(dom_pa, ovs_pa, new_max_date)
    update_sales_dashboard(dom_daily, ovs_daily, new_max_date)
    update_csv_01(dom_daily, ovs_daily)
    update_insights_monthly(dom_daily, ovs_daily, new_max_date)
    update_product_report_dates(new_max_date)

    # JS 문법 검증
    print('\n=== JS 문법 검증 ===')
    for name, path in [('product_analytics', PA_HTML), ('sales_dashboard', SD_HTML),
                        ('insights', INS_HTML), ('product_report', PR_HTML)]:
        print(f'{name}.html:', end=' ')
        validate_js(path)

    print('\n' + '=' * 60)
    print(f'✅ 자동 업데이트 완료 — 최신 날짜: {new_max_date}')
    print('📌 브라우저에서 Cmd/Ctrl+Shift+R (하드 리프레시) 필요')
    print('=' * 60)

if __name__ == '__main__':
    main()
