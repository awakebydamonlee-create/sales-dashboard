#!/usr/bin/env python3
# 오늘의 대표 대시보드 생성기
# 입력: 매출트래킹 폴더의 01_일별매출현황.csv, 02_상품별판매현황.csv, daily_report_summary.json (출고표 요약)
# 출력: 상태페이지.html (자체 포함형, 매일 재생성)
import csv, json, os, datetime

BASE = os.path.dirname(os.path.abspath(__file__))
def p(f): return os.path.join(BASE, f)

def num(x):
    try: return float(str(x).replace(",", ""))
    except: return 0.0

# ---- 일별 매출 ----
daily = []
with open(p("01_일별매출현황.csv"), encoding="utf-8") as f:
    for r in csv.DictReader(f):
        d = r.get("날짜", "")
        if not d or len(d) != 10: continue
        daily.append({
            "date": d,
            "rev": num(r.get("통합매출_KRW")),
            "ord": num(r.get("통합주문건수")),
            "dom": num(r.get("국내매출_KRW")),
            "ovs": num(r.get("해외매출_KRW환산")),
        })
daily.sort(key=lambda x: x["date"])

last = daily[-1]; prev = daily[-2] if len(daily) > 1 else None
last7 = daily[-7:]; prev7 = daily[-14:-7]
def s(arr, k): return sum(a[k] for a in arr)
w_rev = s(last7, "rev"); pw_rev = s(prev7, "rev") if prev7 else 0
dom7 = s(last7, "dom"); ovs7 = s(last7, "ovs"); tot7 = dom7 + ovs7 or 1
trend = daily[-21:]

# ---- 베스트셀러 ----
prods = []
with open(p("02_상품별판매현황.csv"), encoding="utf-8") as f:
    for r in csv.DictReader(f):
        name = r.get("상품명", "")
        qty = num(r.get("통합판매수량"))
        if not name or qty <= 0: continue
        ovsq = num(r.get("해외판매수량"))
        prods.append({
            "name": name, "qty": qty, "rev": num(r.get("통합매출_KRW")),
            "ovs_pct": round(ovsq / qty * 100) if qty else 0,
            "season": r.get("시즌", ""),
        })
prods.sort(key=lambda x: x["qty"], reverse=True)
best = prods[:10]

# ---- 출고표 요약 ----
inv = {}
try:
    with open(p("daily_report_summary.json"), encoding="utf-8") as f:
        inv = json.load(f)
except Exception:
    inv = {}

data = {
    "generated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
    "last": last, "prev": prev,
    "w_rev": w_rev, "pw_rev": pw_rev,
    "dom7": dom7, "ovs7": ovs7,
    "trend": trend, "best": best, "inv": inv,
}

HTML = """<!DOCTYPE html>
<html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>오늘의 대표 대시보드</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.5.0/dist/chart.umd.js" integrity="sha384-iU8HYtnGQ8Cy4zl7gbNMOhsDTTKX02BTXptVP/vqAWIaTfM7isw76iyZCsjL2eVi" crossorigin="anonymous"></script>
<style>
:root{color-scheme:light}*{box-sizing:border-box}
body{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Apple SD Gothic Neo","Malgun Gothic",sans-serif;background:#f7f7f8;color:#1a1a1a}
.wrap{max-width:900px;margin:0 auto;padding:20px 18px 48px}
h1{font-size:20px;margin:0 0 2px}.sub{color:#888;font-size:12.5px;margin-bottom:16px}
.grid{display:grid;gap:12px}.cards{grid-template-columns:repeat(4,1fr)}
@media(max-width:640px){.cards{grid-template-columns:repeat(2,1fr)}}
.card{background:#fff;border:1px solid #ececef;border-radius:12px;padding:14px}
.card .label{font-size:11.5px;color:#888;margin-bottom:6px}.card .val{font-size:19px;font-weight:700}
.card .delta{font-size:11.5px;margin-top:3px}.up{color:#1a9d4e}.down{color:#d23f3f}
.sec{background:#fff;border:1px solid #ececef;border-radius:12px;padding:16px;margin-top:14px}
.sec h2{font-size:14px;margin:0 0 12px}.pill{display:inline-block;background:#eef3fb;color:#2b6cb0;border-radius:6px;padding:1px 7px;font-size:11px;margin-left:6px}
table{width:100%;border-collapse:collapse;font-size:13px}th,td{text-align:left;padding:7px 6px;border-bottom:1px solid #f0f0f2}
th{color:#999;font-weight:600;font-size:11.5px}td.num,th.num{text-align:right;font-variant-numeric:tabular-nums}.rank{color:#bbb;width:20px}
.alert{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}
@media(max-width:640px){.alert{grid-template-columns:repeat(2,1fr)}}
.abox{border-radius:10px;padding:12px;border:1px solid}
.a-red{background:#fdf0f0;border-color:#f5d0d0}.a-amber{background:#fdf7ec;border-color:#f0e2c0}.a-blue{background:#eef3fb;border-color:#d3e0f2}.a-gray{background:#f4f4f6;border-color:#e6e6ea}
.abox .n{font-size:22px;font-weight:700}.abox .t{font-size:11.5px;color:#666;margin-top:2px}
.taglist{margin:10px 0 0;font-size:12.5px;color:#555}.taglist b{color:#222}
ul{margin:6px 0 0;padding-left:18px;font-size:12.5px}li{margin:3px 0}
</style></head><body><div class="wrap">
<h1>오늘의 대표 대시보드</h1>
<div class="sub" id="sub"></div>
<div class="grid cards" id="cards"></div>
<div class="sec"><h2>출고·재고 알림 <span class="pill">출고표 기준</span></h2>
<div class="alert" id="alerts"></div><div class="taglist" id="invtags"></div></div>
<div class="sec"><h2>일별 매출 추이 <span class="pill" id="trange"></span></h2><canvas id="chart" height="120"></canvas></div>
<div class="sec"><h2>베스트셀러 Top 10 <span class="pill">누적 판매수량</span></h2>
<table><thead><tr><th class="rank">#</th><th>상품명</th><th class="num">수량</th><th class="num">매출</th><th class="num">해외%</th></tr></thead><tbody id="best"></tbody></table></div>
</div>
<script>
const D = __DATA__;
const KRW=n=>"₩"+Math.round(n).toLocaleString("ko-KR");
const KM=n=>n>=1e8?(n/1e8).toFixed(2)+"억":n>=1e4?Math.round(n/1e4).toLocaleString()+"만":Math.round(n).toLocaleString();
function delta(c,p){if(!p)return"";const d=(c-p)/p*100;return '<span class="'+(d>=0?"up":"down")+'">'+(d>=0?"▲":"▼")+Math.abs(d).toFixed(1)+"%</span>";}
document.getElementById("sub").textContent="폴더 자료 + 출고표 연동 · 마지막 갱신 "+D.generated+" · 최신일 "+D.last.date;
const dr=D.prev?(D.last.rev-D.prev.rev)/D.prev.rev*100:0, dor=D.prev?(D.last.ord-D.prev.ord)/D.prev.ord*100:0;
const wr=D.pw_rev?(D.w_rev-D.pw_rev)/D.pw_rev*100:0;
const tot7=D.dom7+D.ovs7||1;
document.getElementById("cards").innerHTML=
 card("최신일 매출",KRW(D.last.rev),(D.prev?(dr>=0?"▲":"▼")+Math.abs(dr).toFixed(1)+"% 전일":""),dr>=0)
+card("최신일 주문",Math.round(D.last.ord)+"건",(D.prev?(dor>=0?"▲":"▼")+Math.abs(dor).toFixed(1)+"% 전일":""),dor>=0)
+card("최근 7일 매출",KM(D.w_rev),(D.pw_rev?(wr>=0?"▲":"▼")+Math.abs(wr).toFixed(1)+"% WoW":""),wr>=0)
+card("국내 : 해외",Math.round(D.dom7/tot7*100)+" : "+Math.round(D.ovs7/tot7*100),"최근 7일 매출비중",true);
function card(l,v,d,up){return '<div class="card"><div class="label">'+l+'</div><div class="val">'+v+'</div><div class="delta '+(up?"up":"down")+'">'+d+'</div></div>';}
const iv=D.inv||{};
document.getElementById("alerts").innerHTML=
 abox("a-red",iv.urgent_count||0,"긴급 품절/임박")
+abox("a-amber",iv.refill_count||0,"재입고 필요")
+abox("a-blue",iv.incoming_planned_count||0,"입고 예정")
+abox("a-gray",iv.zero_stock_count||0,"재고 0");
function abox(c,n,t){return '<div class="abox '+c+'"><div class="n">'+n+'</div><div class="t">'+t+'</div></div>';}
let tags="";
if(iv.urgent_top&&iv.urgent_top.length)tags+='<div>긴급: <b>'+iv.urgent_top.join(" · ")+'</b></div>';
if(iv.refill_top&&iv.refill_top.length){tags+='<div style="margin-top:6px">재입고 임박: '+iv.refill_top.map(r=>'<b>'+r.product+'</b> ('+r.days_left+'일)').join(" · ")+'</div>';}
document.getElementById("invtags").innerHTML=tags;
const t=D.trend;
document.getElementById("trange").textContent=t[0].date+" ~ "+t[t.length-1].date;
new Chart(document.getElementById("chart"),{data:{labels:t.map(x=>x.date.slice(5)),datasets:[
 {type:"bar",label:"매출",data:t.map(x=>x.rev),backgroundColor:"#cdddf5",yAxisID:"y",borderRadius:3},
 {type:"line",label:"주문수",data:t.map(x=>x.ord),borderColor:"#2b6cb0",backgroundColor:"#2b6cb0",yAxisID:"y1",tension:.3,pointRadius:0}]},
 options:{responsive:true,plugins:{legend:{labels:{boxWidth:12,font:{size:11}}}},
 scales:{y:{ticks:{callback:v=>KM(v),font:{size:10}}},y1:{position:"right",grid:{drawOnChartArea:false},ticks:{font:{size:10}}},x:{ticks:{font:{size:9},maxRotation:0,autoSkip:true,maxTicksLimit:11}}}}});
document.getElementById("best").innerHTML=D.best.map((b,i)=>
 '<tr><td class="rank">'+(i+1)+'</td><td>'+b.name+'</td><td class="num">'+Math.round(b.qty).toLocaleString()+'</td><td class="num">'+KM(b.rev)+'</td><td class="num">'+b.ovs_pct+'%</td></tr>').join("");
</script></body></html>"""

out = HTML.replace("__DATA__", json.dumps(data, ensure_ascii=False))
with open(p("상태페이지.html"), "w", encoding="utf-8") as f:
    f.write(out)
print("OK", p("상태페이지.html"), "| last:", last["date"], "| best:", len(best), "| urgent:", inv.get("urgent_count"))
