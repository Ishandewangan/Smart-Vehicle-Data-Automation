import re, io, zipfile, hashlib, datetime as dt
from copy import copy
import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
from openpyxl.styles import Font, Alignment
from PIL import Image, ImageOps

HDR = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{2,4}), (\d{1,2}:\d{2})[\s ]*([ap]m)[\s ]*-[\s ]*([^:]+): ?(.*)$", re.I)
PLATE = re.compile(r"[A-Z]{0,2}\d{0,2}[A-Z]{0,2}\d{4}")
FILE = re.compile(r"([\w\-. ]+\.(?:jpe?g|png|webp)) \(file attached\)", re.I)


def parse_chat(text):
    """Return one dict per photo: date, time, sender, caption text, vehicle guess, file name."""
    msgs = []
    for ln in text.splitlines():
        m = HDR.match(ln.strip("‎‏  "))
        if m:
            y = int(m[3])
            msgs.append(dict(date=dt.date(y + 2000 if y < 100 else y, int(m[2]), int(m[1])),
                             time=f"{m[4]} {m[5].upper()}", sender=m[6], lines=[m[7]]))
        elif msgs:
            msgs[-1]["lines"].append(ln)
    clean = lambda ls: " ".join(l.strip() for l in ls if l.strip() and "<Media omitted>" not in l and "(file attached)" not in l)
    out = []
    for i, m in enumerate(msgs):
        f = FILE.search(m["lines"][0])
        if not f:
            continue
        cap = clean(m["lines"][1:])
        p = msgs[i - 1] if i else None
        if not cap and p and p["sender"] == m["sender"] and p["date"] == m["date"] and not FILE.search(p["lines"][0]):
            cap = clean(p["lines"])  # caption sent just before the photo
        v = PLATE.search(re.sub(r"[^A-Z0-9]", "", cap.upper()))
        out.append(dict(date=m["date"], time=m["time"], sender=m["sender"], caption=cap,
                        vehicle=v[0] if v else "", file=f[1].strip()))
    return out


alnum = lambda x: re.sub(r"[^A-Z0-9]", "", str(x or "").upper())
EXTRA = ["Date", "Time", "Text / Caption", "Meter Reading", "Image"]


def place(ws, data, col, row):
    xi = XLImage(io.BytesIO(data))  # photo fills its cell exactly
    xi.anchor = TwoCellAnchor("twoCell", AnchorMarker(col - 1, 0, row - 1, 0), AnchorMarker(col, 0, row, 0))
    ws.add_image(xi)
    ws.row_dimensions[row].height = 80


def cp(a, b):
    b.font, b.fill, b.border = copy(a.font), copy(a.fill), copy(a.border)
    b.alignment, b.number_format = copy(a.alignment), a.number_format


def read_vehicles(up):
    if up.name.lower().endswith(".xls"):
        buf = io.BytesIO(); pd.read_excel(up, header=None).to_excel(buf, header=False, index=False); buf.seek(0); wb = load_workbook(buf)
    else:
        wb = load_workbook(up)
    ws = wb.active
    for r in range(1, min(ws.max_row, 10) + 1):
        for c in range(1, ws.max_column + 1):
            if re.search(r"vehi|number|veh", str(ws.cell(r, c).value or ""), re.I):
                return ws, r, c
    return ws, 1, 1


def fill_row(ws, row, c0, r):
    ws.cell(row, c0, r["date"]).number_format = "DD/MM/YYYY"
    ws.cell(row, c0 + 1, r["time"]); ws.cell(row, c0 + 2, r["caption"])
    for k in range(c0, c0 + 5):
        ws.cell(row, k).alignment = Alignment(wrap_text=True, vertical="center")
    place(ws, r["bytes"], c0 + 4, row)


def build_excel(rows, vehicles=None):
    """vehicles=None -> one flat sheet. Else: original sheet with matched photos beside each vehicle + 'Unmatched' tab."""
    wb = Workbook(); out = wb.active
    plates = {}
    if vehicles is not None:
        ws0, h0, v0 = vehicles
        plates = {alnum(ws0.cell(r, v0).value): str(ws0.cell(r, v0).value) for r in range(h0 + 1, ws0.max_row + 1) if ws0.cell(r, v0).value}
    for r in rows:
        g = alnum(r["vehicle"]); c = [p for k, p in plates.items() if len(g) >= 4 and k.endswith(g)]
        r["plate"] = c[0] if len(c) == 1 else None
    seen, keep, dups = set(), [], []
    for r in sorted(rows, key=lambda x: (x["date"], x["time"])):  # same photo, or same vehicle on same day -> keep the first only
        ks = {("h", r["hash"])} | ({("v", r["plate"] or alnum(r["vehicle"]), r["date"])} if (r["plate"] or alnum(r["vehicle"])) else set())
        if ks & seen:
            dups.append(r)
        else:
            keep.append(r); seen |= ks
    rows = keep
    if vehicles is None:
        out.title = "Meter Readings"; out.append(["Vehicle No."] + EXTRA + ["Sender", "File name"])
        unm, nc, hrow, vcol, ws = rows, 1, 1, 1, None
        matched = {}
    else:
        ws, hrow, vcol = vehicles
        out.title = ws.title; nc = ws.max_column
        matched, unm = {}, []
        for r in rows:
            g = alnum(r["vehicle"])
            if r["plate"]:
                matched.setdefault(r["plate"], []).append(r)
            else:
                r["possible"] = ", ".join(p for k, p in plates.items() if len(g) >= 4 and k.endswith(g))
                unm.append(r)
        for r in range(1, hrow + 1):
            for c in range(1, nc + 1):
                cp(ws.cell(r, c), out.cell(r, c, ws.cell(r, c).value))
        for i, h in enumerate(EXTRA):
            cp(ws.cell(hrow, vcol), out.cell(hrow, nc + 1 + i, h))
        for c in range(1, nc + 1):  # wide enough that rows never auto-grow (keeps photos aligned)
            L = out.cell(1, c).column_letter; out.column_dimensions[L].width = max(ws.column_dimensions[L].width or 8, 18)
        cur = hrow + 1
        for r in range(hrow + 1, ws.max_row + 1):
            v = ws.cell(r, vcol).value
            if v is None:
                continue
            for rec in sorted(matched.get(str(v), []), key=lambda x: (x["date"], x["time"])) or [None]:
                for c in range(1, nc + 1):
                    cp(ws.cell(r, c), out.cell(cur, c, ws.cell(r, c).value))
                if rec:
                    fill_row(out, cur, nc + 1, rec)
                cur += 1
    base = nc if vehicles is not None else 1
    if vehicles is None:
        for i, r in enumerate(unm, 2):
            out.cell(i, 1, r["vehicle"]); fill_row(out, i, 2, r)
            out.cell(i, 7, r["sender"]); out.cell(i, 8, r["file"])
        widths = (14, 12, 10, 34, 14, 24, 18, 28)
    else:
        widths = None
        for k, w in zip(range(nc + 1, nc + 6), (12, 10, 30, 14, 24)):
            out.column_dimensions[out.cell(1, k).column_letter].width = w
    if widths:
        for k, w in enumerate(widths, 1):
            out.column_dimensions[out.cell(1, k).column_letter].width = w
    for c in out[hrow]:
        c.font = Font(bold=True)
    if vehicles is not None and unm:
        u = wb.create_sheet("Unmatched")
        u.append(["Vehicle (from text)", "Possible vehicles"] + EXTRA + ["Sender", "File name"])
        for c in u[1]:
            c.font = Font(bold=True)
        for i, r in enumerate(unm, 2):
            u.cell(i, 1, r["vehicle"]); u.cell(i, 2, r["possible"]); fill_row(u, i, 3, r)
            u.cell(i, 8, r["sender"]); u.cell(i, 9, r["file"])
        for col, w in zip("ABCDEFGHI", (18, 24, 12, 10, 30, 14, 24, 18, 28)):
            u.column_dimensions[col].width = w
    if dups:
        d = wb.create_sheet("Duplicates")
        d.append(["Vehicle (from text)", "Matched vehicle"] + EXTRA + ["Sender / group", "File name"])
        for c in d[1]:
            c.font = Font(bold=True)
        for i, r in enumerate(dups, 2):
            d.cell(i, 1, r["vehicle"]); d.cell(i, 2, r["plate"] or ""); fill_row(d, i, 3, r)
            d.cell(i, 8, r["sender"]); d.cell(i, 9, r["file"])
        for col, w in zip("ABCDEFGHI", (18, 18, 12, 10, 30, 14, 24, 24, 28)):
            d.column_dimensions[col].width = w
    b = io.BytesIO(); wb.save(b)
    return b.getvalue(), len(unm), sum(len(v) for v in matched.values()), len(dups)


st.set_page_config("WhatsApp Photos → Excel", "🚗", layout="wide")
st.title("🚗 WhatsApp Photos → Excel")
st.caption("Upload the WhatsApp export .zip (and optionally your vehicle list). Matched photos + text go beside their vehicle; the rest go to an 'Unmatched' tab. "
           "Type the meter reading from the photo in the 'Meter Reading' column.")
c1, c2 = st.columns(2)
xl = c1.file_uploader("1) Vehicle list Excel (optional)", ["xlsx", "xls"])
wa = c2.file_uploader("2) WhatsApp export .zip files (you can select several groups)", ["zip"], accept_multiple_files=True)

if wa and st.button("▶ Create Excel", type="primary"):
    rows, skipped, seen = [], [], set()
    for up in wa:
        z = zipfile.ZipFile(up)
        group = re.sub(r"^WhatsApp Chat with |\.zip$|\s*\(\d+\)$", "", up.name)
        names = {n.split("/")[-1]: n for n in z.namelist()}
        chat = next((z.read(n).decode("utf-8-sig", "replace") for n in z.namelist() if n.lower().endswith(".txt")), "")
        for it in parse_chat(chat):
            key = (group, it["file"], it["date"], it["time"])
            if key in seen:
                continue
            seen.add(key)
            try:
                im = ImageOps.exif_transpose(Image.open(io.BytesIO(z.read(names[it["file"]])))).convert("RGB")
                im.thumbnail((900, 900))
                bb = io.BytesIO(); im.save(bb, "JPEG", quality=80)
                rows.append({**it, "sender": f"{it['sender']} [{group}]", "bytes": bb.getvalue(),
                             "hash": hashlib.md5(z.read(names[it["file"]])).hexdigest()})
            except Exception:
                skipped.append(it["file"])
    data, nun, nm, nd = build_excel(rows, read_vehicles(xl) if xl else None)
    st.session_state.out = (data, len(rows), skipped, [r["bytes"] for r in rows], nun, nm, bool(xl), nd)

if st.session_state.get("out"):
    data, n, skipped, imgs, nun, nm, had_xl, nd = st.session_state.out
    st.success(f"{n - nd} photos saved in the Excel." + (f" {nm} matched to your vehicle list, {nun} placed in the 'Unmatched' tab." if had_xl else ""))
    if nd:
        st.info(f"{nd} duplicate(s) removed (same photo or same vehicle on the same date in another group). They are listed in the 'Duplicates' tab.")
    if skipped:
        st.warning("Empty/unreadable images skipped: " + ", ".join(skipped))
    st.download_button("⬇ Download Excel", data, "whatsapp_meter_photos.xlsx", type="primary")
    st.image(imgs[:12], width=140)
