#!/usr/bin/env python3
"""KR Lost Customer Tracker

คำสั่ง
  build  : อ่าน booking export (.xls/.xlsx) -> สร้างไฟล์ HTML สำหรับส่งให้เซล
  merge  : รวมไฟล์ HTML ที่เซลส่งกลับ -> Excel สรุป feedback + HTML ฉบับรวม

ตัวอย่าง
  python3 kr_tracker.py build KR_May_to_Oct.xls
  python3 kr_tracker.py build KR_May_to_Oct.xls -o dist/KR_Tracker.html
  python3 kr_tracker.py merge returned/*.html
  python3 kr_tracker.py merge returned/*.html --base dist/KR_Lost_Customer_Tracker.html -o merged

ต้องติดตั้ง: pip install pandas xlrd openpyxl
"""
from __future__ import annotations

import argparse
import glob
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
TEMPLATE = ROOT / "src" / "template.html"
MONTH_TH = {1: "ม.ค.", 2: "ก.พ.", 3: "มี.ค.", 4: "เม.ย.", 5: "พ.ค.", 6: "มิ.ย.",
            7: "ก.ค.", 8: "ส.ค.", 9: "ก.ย.", 10: "ต.ค.", 11: "พ.ย.", 12: "ธ.ค."}
RESULT_TH = {
    "back": "จะกลับมาจอง / มีงานเดือนหน้า",
    "shift": "เลื่อนชิปเมนต์ / ไปอยู่เรือเดือนถัดไป",
    "price": "ราคาสู้คู่แข่งไม่ได้",
    "competitor": "ย้ายไปใช้สายเรืออื่น",
    "space": "ปัญหาพื้นที่ / ตารางเรือ / อุปกรณ์",
    "noorder": "ออเดอร์ลด / ไม่มีสินค้า",
    "spot": "เป็นงาน spot ไม่ประจำ",
    "closed": "เลิกส่งเส้นทางนี้ / ปิดกิจการ",
    "other": "อื่นๆ (ดูหมายเหตุ)",
}
TYPE_TH = {"lost": "หายไปเดือนนี้", "drop": "ลดลง ≥50%", "gone": "หายไปก่อนหน้า", "new": "ลูกค้าใหม่"}
FB_RE = re.compile(r'(<script type="application/json" id="fbdata">)([\s\S]*?)(</script>)')


# ---------------------------------------------------------------- build
def load_bookings(path: Path) -> pd.DataFrame:
    d = pd.read_excel(path)
    need = ["ETD", "TEU", "SHPCD", "Shipper", "SALESNM", "POD"]
    miss = [c for c in need if c not in d.columns]
    if miss:
        sys.exit(f"ไม่พบคอลัมน์ {miss} ในไฟล์ {path.name}")
    d["ETD"] = pd.to_datetime(d["ETD"], errors="coerce")
    d = d.dropna(subset=["ETD"]).copy()
    d["TEU"] = pd.to_numeric(d["TEU"], errors="coerce").fillna(0)
    d["ym"] = d["ETD"].dt.to_period("M")
    org = d["Org Shipper"] if "Org Shipper" in d.columns else d["Shipper"]
    d["Org"] = (org.fillna(d["Shipper"]).astype(str).str.strip().str.upper()
                .str.replace(r"\s+", " ", regex=True))
    com = d["Commodity"] if "Commodity" in d.columns else pd.Series("", index=d.index)
    item = d["Main Item"] if "Main Item" in d.columns else pd.Series("", index=d.index)
    d["Commodity"] = com.fillna(item).fillna("").astype(str).str.strip().str.upper()
    return d


def summarise(d: pd.DataFrame, periods: list, key: str, namecol: str, other: str) -> list[dict]:
    out = []
    for k, g in d.groupby(key):
        by_m = [g[g.ym == p] for p in periods]
        last = g.sort_values("ETD").iloc[-1]
        sales = g.groupby("SALESNM")["TEU"].sum().sort_values(ascending=False)
        top_o = g.groupby(other)["TEU"].sum().sort_values(ascending=False).head(4)
        pods = g.groupby("POD")["TEU"].sum().sort_values(ascending=False).head(3)
        com = g.groupby("Commodity")["TEU"].sum().sort_values(ascending=False).head(2)
        out.append(dict(
            k=str(k), n=str(g[namecol].mode().iloc[0]).strip(),
            s=str(last.SALESNM), ss=[str(x) for x in sales.index[:3]],
            t=[float(x.TEU.sum()) for x in by_m], b=[int(len(x)) for x in by_m],
            last=last.ETD.strftime("%Y-%m-%d"),
            lastByMonth=[x.ETD.max().strftime("%d/%m") if len(x) else "" for x in by_m],
            o=[[str(a), round(float(v), 1)] for a, v in top_o.items()],
            pod=[str(a) for a in pods.index], com=[str(a)[:40] for a in com.index]))
    return out


def build_data(d: pd.DataFrame) -> dict:
    periods = list(pd.period_range(d.ym.min(), d.ym.max(), freq="M"))
    asof = d["Update Date"].max() if "Update Date" in d.columns else d.ETD.max()
    return dict(
        months=[str(p) for p in periods],
        monthNames=[MONTH_TH[p.month] for p in periods],
        asof=str(asof)[:10], etdMax=d.ETD.max().strftime("%Y-%m-%d"),
        total=[float(d[d.ym == p].TEU.sum()) for p in periods],
        totalBk=[int((d.ym == p).sum()) for p in periods],
        cust=summarise(d, periods, "SHPCD", "Shipper", "Org"),
        org=summarise(d, periods, "Org", "Org", "Shipper"))


def render_html(data: dict, template: Path = TEMPLATE) -> str:
    t = template.read_text(encoding="utf-8")
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    # month labels follow the data instead of the fixed May–Oct list
    t = re.sub(r"const MNAME = \[[^\]]*\];", "const MNAME = DATA.monthNames;", t, count=1)
    t = t.replace("let state = {m:5,", "let state = {m:DATA.months.length-1,", 1)
    t = t.replace("if(m===5) note+=", "if(m===DATA.months.length-1) note+=", 1)
    t = t.replace("(ETD ล่าสุดในไฟล์ ${DATA.etdMax.slice(8)}/10)",
                  "(ETD ล่าสุดในไฟล์ ${DATA.etdMax.slice(8)}/${DATA.etdMax.slice(5,7)})", 1)
    body = t.replace("__DATA__", payload)
    return ('<!doctype html>\n<html lang="th"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width,initial-scale=1"></head><body>\n'
            + body + "\n</body></html>")


def print_summary(data: dict) -> None:
    m = len(data["months"]) - 1
    names = data["monthNames"]
    print("\nTEU รายเดือน: " + " | ".join(f"{n} {int(v)}" for n, v in zip(names, data["total"])))
    if m < 1:
        return
    rows = data["cust"]
    lost = sorted([c for c in rows if c["t"][m] == 0 and c["t"][m - 1] > 0], key=lambda c: -c["t"][m - 1])
    drop = [c for c in rows if c["t"][m] > 0 and c["t"][m - 1] > 0 and c["t"][m] <= c["t"][m - 1] * .5]
    gone = [c for c in rows if c["t"][m] == 0 and c["t"][m - 1] == 0 and any(c["t"][:m - 1])]
    print(f"{names[m]} เทียบ {names[m-1]}: หายไป {len(lost)} ราย ({sum(c['t'][m-1] for c in lost):.0f} TEU), "
          f"ลดลง ≥50% {len(drop)} ราย, หายไปก่อนหน้า {len(gone)} ราย")
    for c in lost[:10]:
        print(f"  - {c['n'][:45]:45s} {c['s'][:20]:20s} {c['t'][m-1]:>5.0f} TEU")


def cmd_build(a) -> None:
    src = Path(a.xls)
    d = load_bookings(src)
    data = build_data(d)
    out = Path(a.output) if a.output else ROOT / "dist" / "KR_Lost_Customer_Tracker.html"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(render_html(data), encoding="utf-8")
    if a.json:
        Path(a.json).write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    print(f"อ่าน {len(d)} bookings จาก {src.name} -> {out}")
    print_summary(data)


# ---------------------------------------------------------------- merge
def read_feedback(path: Path) -> dict:
    m = FB_RE.search(path.read_text(encoding="utf-8", errors="replace"))
    if not m:
        raise ValueError("ไม่พบข้อมูล feedback ในไฟล์")
    return json.loads(m.group(2)).get("items", {})


def merge_items(files: list[Path]) -> tuple[dict, list[str]]:
    merged, bad = {}, []
    for f in files:
        try:
            items = read_feedback(f)
        except Exception as e:  # noqa: BLE001
            bad.append(f"{f.name} ({e})")
            continue
        for k, r in items.items():
            if not isinstance(r, dict):
                continue
            if k not in merged or (r.get("at") or 0) > (merged[k].get("at") or 0):
                merged[k] = {**r, "_file": f.name}
    return merged, bad


def cmd_merge(a) -> None:
    files = sorted({Path(p) for pat in a.files for p in glob.glob(pat)})
    if not files:
        sys.exit("ไม่พบไฟล์ที่ระบุ")
    merged, bad = merge_items(files)
    out_dir = Path(a.output)
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")

    rows = []
    for r in merged.values():
        at = r.get("at") or 0
        rows.append({
            "เดือน": r.get("month", ""),
            "มุมมอง": "Shipper" if r.get("view") == "cust" else "Org Shipper",
            "รหัส": r.get("key", "") if r.get("view") == "cust" else "",
            "ลูกค้า": r.get("name", ""),
            "Sales": r.get("sales", ""),
            "ประเภท": TYPE_TH.get(r.get("type", ""), r.get("type", "")),
            "ผลการเช็ค": RESULT_TH.get(r.get("result", ""), ""),
            "หมายเหตุ": r.get("note", ""),
            "ผู้เช็ค": r.get("by", ""),
            "เวลาบันทึก": datetime.fromtimestamp(at / 1000).strftime("%Y-%m-%d %H:%M") if at else "",
            "จากไฟล์": r.get("_file", ""),
        })
    df = pd.DataFrame(rows)
    xlsx = out_dir / f"KR_feedback_{stamp}.xlsx"
    with pd.ExcelWriter(xlsx, engine="openpyxl") as w:
        if df.empty:
            pd.DataFrame({"หมายเหตุ": ["ยังไม่มีคำตอบในไฟล์ที่รวม"]}).to_excel(w, sheet_name="Feedback", index=False)
        else:
            df = df.sort_values(["เดือน", "Sales", "ลูกค้า"])
            df.to_excel(w, sheet_name="Feedback", index=False)
            done = df[df["ผลการเช็ค"] != ""]
            pd.crosstab(done["Sales"], done["ผลการเช็ค"], margins=True, margins_name="รวม") \
                .to_excel(w, sheet_name="สรุปตาม Sales")
            done.groupby("ผลการเช็ค").size().sort_values(ascending=False) \
                .rename("จำนวนราย").to_excel(w, sheet_name="สรุปเหตุผล")
        for ws in w.book.worksheets:
            for col in ws.columns:
                width = max(len(str(c.value or "")) for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 60)
            ws.freeze_panes = "A2"
    print(f"รวม {len(files) - len(bad)} ไฟล์ · {len(merged)} คำตอบ -> {xlsx}")

    base = Path(a.base) if a.base else ROOT / "dist" / "KR_Lost_Customer_Tracker.html"
    if base.exists():
        payload = json.dumps({"rev": f"merged-{stamp}", "items": merged}, ensure_ascii=False) \
            .replace("<", "\\u003c")
        html = FB_RE.sub(lambda m: m.group(1) + payload + m.group(3), base.read_text(encoding="utf-8"), count=1)
        out_html = out_dir / f"KR_Lost_Customer_Tracker_merged_{stamp}.html"
        out_html.write_text(html, encoding="utf-8")
        print(f"HTML ฉบับรวม -> {out_html}")
    else:
        print(f"(ข้ามการสร้าง HTML ฉบับรวม: ไม่พบไฟล์ฐาน {base})")
    for b in bad:
        print(f"  อ่านไม่ได้: {b}")


def main() -> None:
    p = argparse.ArgumentParser(description="KR Lost Customer Tracker")
    sub = p.add_subparsers(dest="cmd", required=True)
    b = sub.add_parser("build", help="สร้างไฟล์ HTML จาก booking export")
    b.add_argument("xls", help="ไฟล์ booking export (.xls / .xlsx)")
    b.add_argument("-o", "--output", help="ไฟล์ HTML ปลายทาง (ค่าเริ่มต้น dist/KR_Lost_Customer_Tracker.html)")
    b.add_argument("--json", help="บันทึกข้อมูลที่คำนวณเป็น JSON ด้วย")
    b.set_defaults(func=cmd_build)
    m = sub.add_parser("merge", help="รวมไฟล์ HTML ที่เซลส่งกลับ")
    m.add_argument("files", nargs="+", help="ไฟล์ HTML ที่ได้รับคืน (ใช้ *.html ได้)")
    m.add_argument("--base", help="ไฟล์ HTML ฐานสำหรับสร้างฉบับรวม")
    m.add_argument("-o", "--output", default="merged", help="โฟลเดอร์ผลลัพธ์ (ค่าเริ่มต้น merged/)")
    m.set_defaults(func=cmd_merge)
    a = p.parse_args()
    a.func(a)


if __name__ == "__main__":
    main()
