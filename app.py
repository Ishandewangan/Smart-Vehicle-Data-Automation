import re, io, os, zipfile, datetime as dt
from copy import copy
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image, ImageOps
import pytesseract

st.set_page_config("Smart Vehicle Data Automation", "🚗", layout="wide")
S = st.session_state
for k, v in dict(wb=None, ws=None, fname=None, recs=[], hist=[], unmatched=[], dups=0, imgs=0).items():
    S.setdefault(k, v)
os.makedirs("images", exist_ok=True)
HDRS = ["Date", "Shift", "Meter Reading", "Extracted Text", "Image"]

norm = lambda s: re.sub(r"\D", "", str(s or ""))

def nums(t):
    return [norm(m) for m in re.findall(r"\d[\d\s-]*\d", t)]

PFX = re.compile(r"^\[?\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4},?\s+\d{1,2}:\d{2}(:\d{2})?\s*([ap]\.?m\.?)?\]?\s*-?\s*[^:]*:", re.I)

def captions(text):
    """WhatsApp export: line with image filename, next line = caption (vehicle number)."""
    m, last = {}, None
    for ln in text.splitlines():
        if PFX.match(ln):
            rest = PFX.sub("", ln).strip()
            f = re.match(r"(?:<attached:\s*)?(.+?\.(?:jpe?g|png))", rest, re.I)
            last = f.group(1).strip() if f else None
        elif last:
            v = next((x for x in nums(ln) if len(x) in (4, 6)), None)
            if v:
                m[last] = v
            last = None
    return m

def prepare_image(data):
    im = ImageOps.autocontrast(Image.open(io.BytesIO(data)).convert("L"))
    return im.resize((im.width * 2, im.height * 2))

def read_meter(data):
    """OCR digits from the odometer photo; pick the longest 4-7 digit number."""
    im = prepare_image(data)
    text = pytesseract.image_to_string(im, config="--psm 6 -c tessedit_char_whitelist=0123456789")
    candidates = [x for x in re.findall(r"\d+", text) if 4 <= len(x) <= 7]
    return max(candidates, key=len) if candidates else ""

def read_image_text(data):
    """Extract general text from an image using Tesseract OCR."""
    im = prepare_image(data)
    return pytesseract.image_to_string(im, config="--psm 6").strip()

def parse(text, src, date, shift, img=None):
    """Return records: vehicle = first 4/6 digit number, meter = next other number."""
    out = []
    lines = text.splitlines() if img is None else [text.replace("\n", "  ")]
    for ln in lines:
        ln = PFX.sub("", ln)
        n = nums(ln)
        veh = next((x for x in n if len(x) in (4, 6)), None)
        if not veh:
            continue
        rest = [x for x in n if x != veh and len(x) >= 2]
        meter = max(rest, key=len) if rest else ""
        if img is None and not meter:
            continue
        out.append(dict(ok=False, vehicle=veh, meter=meter, text="", date=str(date), shift=shift, source=src, img=img))
    return out

def match(v, cell):
    c = norm(cell)
    return bool(c) and (c == v or (len(v) == 4 and c.endswith(v)))

def load_excel(f):
    if f.name.lower().endswith(".xls"):
        b = io.BytesIO()
        pd.read_excel(f).to_excel(b, index=False)
        b.seek(0)
        wb = load_workbook(b)
    else:
        wb = load_workbook(f)
    S.wb, S.ws, S.fname = wb, wb.active, f.name
    S.orig = pd.DataFrame(wb.active.values)
    ws = wb.active
    heads = [str(c.value or "") for c in ws[1]]
    S.vcol = next((i + 1 for i, h in enumerate(heads) if "vehicle" in h.lower() or "number" in h.lower()), 1)
    S.cols = {}
    for h in HDRS:
        i = next((i + 1 for i, x in enumerate(heads) if x.strip().lower() == h.lower()), None)
        if not i:
            i = ws.max_column + 1
            ws.cell(1, i, h)
            heads.append(h)
        S.cols[h] = i

def remove_images_from_row(ws, row):
    """Remove existing embedded images anchored to a worksheet row."""
    kept = []
    for drawing in ws._images:
        anchor_row = getattr(getattr(drawing, "anchor", None), "_from", None)
        if anchor_row is None or anchor_row.row + 1 != row:
            kept.append(drawing)
    ws._images = kept

def apply(r, on_dup):
    ws, H = S.ws, S.cols
    rows = [x for x in range(2, ws.max_row + 1) if match(r["vehicle"], ws.cell(x, S.vcol).value)]
    if not rows:
        return "unmatched"
    same = [x for x in rows if str(ws.cell(x, H["Date"]).value) == r["date"] and str(ws.cell(x, H["Shift"]).value) == r["shift"]]
    if same:
        if on_dup == "Skip":
            return "duplicate-skipped"
        x, st_ = same[0], "duplicate-updated"
    else:
        empty = [x for x in rows if not ws.cell(x, H["Date"]).value]
        if empty:
            x = empty[0]
        else:
            x = ws.max_row + 1
            for c in range(1, ws.max_column + 1):
                s, d = ws.cell(rows[0], c), ws.cell(x, c)
                d.value = s.value if c == S.vcol else None
                d._style = copy(s._style)
        st_ = "updated"

    ws.cell(x, H["Date"], r["date"])
    ws.cell(x, H["Shift"], r["shift"])
    ws.cell(x, H["Meter Reading"], r["meter"])
    ws.cell(x, H["Extracted Text"], r.get("text", ""))

    path = ""
    if r.get("img") is not None:
        remove_images_from_row(ws, x)
        path = f"images/{r['vehicle']}_{r['date']}_{r['shift']}.png"
        pil_im = Image.open(io.BytesIO(r["img"])).convert("RGB")
        pil_im.save(path)

        xl_im = XLImage(io.BytesIO(r["img"]))
        xl_im.width = 160
        xl_im.height = 90
        ws.add_image(xl_im, f"{ws.cell(1, H['Image']).column_letter}{x}")
        ws.row_dimensions[x].height = 75

    ws.cell(x, H["Image"], path)
    return st_

def wb_bytes():
    b = io.BytesIO()
    S.wb.save(b)
    return b.getvalue()

# ---------------- UI ----------------
st.title("🚗 Smart Vehicle Data Automation")
tabs = st.tabs(["📊 Overview", "⬆️ Upload & Process", "✅ Verification", "📑 Excel", "🕘 History"])

with tabs[1]:
    c1, c2 = st.columns(2)
    date = c1.date_input("Date", dt.date.today())
    shift = c2.selectbox("Shift", ["Morning", "Evening", "Night"])
    f = st.file_uploader("1) Vehicle list Excel", ["xlsx", "xls"])
    if f and S.fname != f.name:
        load_excel(f)
    if S.ws is not None:
        hd = [str(c.value) for c in S.ws[1]]
        S.vcol = st.selectbox("Vehicle number column", range(1, len(hd) + 1), S.vcol - 1, format_func=lambda i: hd[i - 1]) if len(hd) else 1
        st.success(f"Loaded {S.fname}")

    wa = st.file_uploader("2) WhatsApp data (exported .zip, or .txt + images together)", ["zip", "txt", "jpg", "jpeg", "png"], accept_multiple_files=True)
    if st.button("▶ Process", type="primary"):
        new, bar = [], st.progress(0.0)
        files = []
        for u in wa:
            if u.name.lower().endswith(".zip"):
                with zipfile.ZipFile(u) as z:
                    files += [(os.path.basename(n), z.read(n)) for n in z.namelist() if not n.endswith("/")]
            else:
                files.append((u.name, u.read()))

        cap = {}
        for name, d in files:
            if name.lower().endswith(".txt"):
                body = d.decode(errors="ignore")
                cap.update(captions(body))
                new += parse(body, name, date, shift)

        ims = [(n, d) for n, d in files if n.lower().endswith((".jpg", ".jpeg", ".png"))]
        for i, (name, data) in enumerate(ims):
            try:
                meter = read_meter(data)
                extracted_text = read_image_text(data)
            except Exception as e:
                st.error(f"OCR failed for {name}: {e}")
                meter, extracted_text = "", ""
            new.append(dict(ok=False, vehicle=cap.get(name, ""), meter=meter, text=extracted_text,
                            date=str(date), shift=shift, source=name, img=data))
            S.imgs += 1
            bar.progress((i + 1) / len(ims) if ims else 1.0)
        S.recs += new
        S.hist.append(dict(time=dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), event="processed", detail=f"{len(new)} records"))
        st.info(f"{len(new)} records extracted → go to Verification tab (nothing is saved until you confirm).")

with tabs[2]:
    if not S.recs:
        st.write("Nothing to verify.")
    else:
        df = pd.DataFrame([{k: r[k] for k in ("ok", "vehicle", "meter", "text", "date", "shift", "source")} for r in S.recs])
        ed = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="ed")
        image_rows = [i for i, r in enumerate(S.recs) if r["img"] is not None]
        pick = st.selectbox("View image for row", image_rows,
                            format_func=lambda i: f"{i}: {S.recs[i]['source']}") if image_rows else None
        if pick is not None:
            st.image(S.recs[pick]["img"], width=400)
        on_dup = st.radio("If duplicate (vehicle+date+shift)", ["Skip", "Update"], horizontal=True)
        a, b = st.columns(2)
        if a.button("💾 Save accepted to Excel"):
            if S.ws is None:
                st.error("Upload Excel first")
            else:
                for i, row in ed.iterrows():
                    if not row["ok"]:
                        continue
                    r = {**row.to_dict(), "vehicle": norm(row["vehicle"]),
                         "img": S.recs[i]["img"] if i < len(S.recs) else None}
                    res = apply(r, on_dup)
                    if res == "unmatched":
                        S.unmatched.append(r["vehicle"])
                    if res.startswith("duplicate"):
                        S.dups += 1
                        st.warning(f"Duplicate: {r['vehicle']} {r['date']} {r['shift']} → {res}")
                    S.hist.append(dict(time=dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), event=res,
                                       detail=f"{r['vehicle']} {r['meter']} {r['date']} {r['shift']}"))
                S.recs = [r for i, r in enumerate(S.recs) if not (i in ed.index and ed.loc[i, "ok"])]
                st.success("Saved. See Excel tab.")
        if b.button("🗑 Reject all"):
            S.recs = []
            st.rerun()

with tabs[3]:
    if S.ws is None:
        st.write("Upload an Excel file first.")
    else:
        st.subheader("Original")
        st.dataframe(S.orig, use_container_width=True)
        st.subheader("Updated")
        st.dataframe(pd.DataFrame(S.ws.values), use_container_width=True)
        download_name = "updated_" + (S.fname.replace(".xls", ".xlsx") if S.fname.lower().endswith(".xls") else S.fname)
        st.download_button("⬇ Updated Excel", wb_bytes(), download_name)
        st.download_button("⬇ Unmatched report", pd.DataFrame({"unmatched_vehicle": S.unmatched}).to_csv(index=False), "unmatched.csv")
        st.download_button("⬇ Processing report", pd.DataFrame(S.hist).to_csv(index=False), "report.csv")

with tabs[4]:
    st.dataframe(pd.DataFrame(S.hist), use_container_width=True)

with tabs[0]:
    h = [x["event"] for x in S.hist]
    total = (S.ws.max_row - 1) if S.ws is not None else 0
    c = st.columns(5)
    c[0].metric("Total vehicles", total)
    c[1].metric("Matched", h.count("updated") + h.count("duplicate-updated"))
    c[2].metric("Unmatched", len(S.unmatched))
    c[3].metric("Duplicates", S.dups)
    c[4].metric("Images processed", S.imgs)
    if S.unmatched:
        st.warning("Unmatched: " + ", ".join(S.unmatched))
