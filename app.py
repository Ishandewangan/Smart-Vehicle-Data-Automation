import re, io, os, zipfile, datetime as dt
from copy import copy
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image, ImageOps, ImageFilter
import pytesseract

st.set_page_config("Smart Vehicle Data Automation", "🚗", layout="wide")
S = st.session_state
for k, v in dict(wb=None, ws=None, fname=None, recs=[], hist=[], unmatched=[], dups=0, imgs=0).items():
    S.setdefault(k, v)
os.makedirs("images", exist_ok=True)
HDRS = ["Date", "Shift", "Meter Reading", "Extracted Text", "Image"]
IMAGE_EXTS = (".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff", ".gif")

norm = lambda s: re.sub(r"\D", "", str(s or ""))

def nums(text):
    return [norm(m) for m in re.findall(r"\d[\d\s-]*\d", str(text or ""))]

PFX = re.compile(r"^\[?\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4},?\s+\d{1,2}:\d{2}(:\d{2})?\s*([ap]\.?m\.?)?\]?\s*-?\s*[^:]*:", re.I)

def captions(text):
    """Find image captions in a WhatsApp text export, keyed by image filename."""
    found, last = {}, None
    for line in text.splitlines():
        if PFX.match(line):
            rest = PFX.sub("", line).strip()
            file_match = re.search(r"([^/\\<>]+\.(?:jpe?g|png|webp|bmp|tiff?|gif))", rest, re.I)
            last = file_match.group(1).strip() if file_match else None
        elif last:
            vehicle = next((x for x in nums(line) if len(x) in (4, 6)), None)
            if vehicle:
                found[last] = vehicle
            last = None
    return found

def normalize_image(data):
    """Validate image bytes and return an EXIF-corrected PNG."""
    with Image.open(io.BytesIO(data)) as source:
        source.load()
        image = ImageOps.exif_transpose(source).convert("RGB")
        output = io.BytesIO()
        image.save(output, format="PNG")
        return output.getvalue()

def ocr_variants(data):
    """Prepare several useful image variants to improve OCR on phone photos."""
    with Image.open(io.BytesIO(normalize_image(data))) as source:
        gray = ImageOps.grayscale(source)
    gray = ImageOps.autocontrast(gray).filter(ImageFilter.SHARPEN)
    if max(gray.size) < 1800:
        scale = min(3, max(2, 1800 // max(gray.size)))
        gray = gray.resize((gray.width * scale, gray.height * scale), Image.Resampling.LANCZOS)
    threshold = gray.point(lambda px: 255 if px > 150 else 0)
    return [gray, threshold]

def read_image_text(data):
    """Run OCR with more than one layout mode and choose the clearest result."""
    candidates = []
    for image in ocr_variants(data):
        for psm in (6, 11, 12):
            text = pytesseract.image_to_string(image, config=f"--oem 3 --psm {psm}").strip()
            if text:
                candidates.append(text)
    if not candidates:
        return ""
    return max(candidates, key=lambda text: (len(re.sub(r"\s", "", text)), sum(ch.isalnum() for ch in text)))

def read_meter(data):
    """Try multiple preprocessing/layout options and return a plausible odometer value."""
    candidates = []
    for image in ocr_variants(data):
        for psm in (6, 7, 11, 13):
            text = pytesseract.image_to_string(
                image, config=f"--oem 3 --psm {psm} -c tessedit_char_whitelist=0123456789"
            )
            candidates.extend(x for x in re.findall(r"\d+", text) if 4 <= len(x) <= 7)
    return max(candidates, key=len) if candidates else ""

def parse(text, src, date, shift):
    """Parse vehicle/meter pairs from text exports; keep unmatched lines for review."""
    output = []
    for line in text.splitlines():
        line = PFX.sub("", line)
        values = nums(line)
        if not values:
            continue
        vehicle = next((x for x in values if len(x) in (4, 6)), "")
        other = [x for x in values if x != vehicle and len(x) >= 2]
        meter = max(other, key=len) if other else ""
        output.append(dict(ok=False, vehicle=vehicle, meter=meter, text=line.strip(),
                           date=str(date), shift=shift, source=src, img=None))
    return output

def match(vehicle, cell):
    cell_value = norm(cell)
    return bool(vehicle and cell_value) and (
        cell_value == vehicle or (len(vehicle) == 4 and cell_value.endswith(vehicle))
    )

def load_excel(upload):
    if upload.name.lower().endswith(".xls"):
        buffer = io.BytesIO()
        pd.read_excel(upload).to_excel(buffer, index=False)
        buffer.seek(0)
        workbook = load_workbook(buffer)
    else:
        workbook = load_workbook(upload)
    S.wb, S.ws, S.fname = workbook, workbook.active, upload.name
    S.orig = pd.DataFrame(workbook.active.values)
    worksheet = workbook.active
    headers = [str(cell.value or "") for cell in worksheet[1]]
    S.vcol = next((i + 1 for i, header in enumerate(headers)
                   if "vehicle" in header.lower() or "number" in header.lower()), 1)
    S.cols = {}
    for header in HDRS:
        col = next((i + 1 for i, value in enumerate(headers)
                    if value.strip().lower() == header.lower()), None)
        if col is None:
            col = worksheet.max_column + 1
            worksheet.cell(1, col, header)
            headers.append(header)
        S.cols[header] = col

def remove_images_from_row(worksheet, row):
    kept = []
    for drawing in worksheet._images:
        anchor = getattr(getattr(drawing, "anchor", None), "_from", None)
        if anchor is None or anchor.row + 1 != row:
            kept.append(drawing)
    worksheet._images = kept

def append_excel_record(record):
    """Append a record when it cannot be matched to an existing vehicle row."""
    worksheet, cols = S.ws, S.cols
    row = worksheet.max_row + 1
    if row > 2 and worksheet.max_column:
        for col in range(1, worksheet.max_column + 1):
            worksheet.cell(row, col)._style = copy(worksheet.cell(row - 1, col)._style)
    if record.get("vehicle"):
        worksheet.cell(row, S.vcol, record["vehicle"])
    worksheet.cell(row, cols["Date"], record["date"])
    worksheet.cell(row, cols["Shift"], record["shift"])
    worksheet.cell(row, cols["Meter Reading"], record.get("meter", ""))
    worksheet.cell(row, cols["Extracted Text"], record.get("text", ""))
    add_excel_image(record, row)
    return "appended"

def add_excel_image(record, row):
    image_bytes = record.get("img")
    if image_bytes is None:
        return
    try:
        image_bytes = normalize_image(image_bytes)
    except Exception:
        st.warning(f"Could not embed unreadable image {record.get('source', '')}.")
        return
    filename = f"images/{record.get('vehicle') or 'unmatched'}_{record['date']}_{record['shift']}_{row}.png"
    with Image.open(io.BytesIO(image_bytes)) as image:
        image.save(filename, format="PNG")
    embedded = XLImage(io.BytesIO(image_bytes))
    embedded.width, embedded.height = 160, 90
    S.ws.add_image(embedded, f"{S.ws.cell(1, S.cols['Image']).column_letter}{row}")
    S.ws.cell(row, S.cols["Image"], filename)
    S.ws.row_dimensions[row].height = 75

def apply(record, on_dup):
    worksheet, cols = S.ws, S.cols
    vehicle = norm(record.get("vehicle", ""))
    record["vehicle"] = vehicle
    # Do not discard an image/text record just because OCR did not identify a vehicle,
    # or because that vehicle is not present in the uploaded vehicle list.
    rows = [row for row in range(2, worksheet.max_row + 1)
            if match(vehicle, worksheet.cell(row, S.vcol).value)] if vehicle else []
    if not rows:
        result = append_excel_record(record)
        return "image-appended" if record.get("img") is not None and not vehicle else result

    same = [row for row in rows
            if str(worksheet.cell(row, cols["Date"]).value) == str(record["date"])
            and str(worksheet.cell(row, cols["Shift"]).value) == str(record["shift"])]
    if same:
        if on_dup == "Skip":
            return "duplicate-skipped"
        row, status = same[0], "duplicate-updated"
    else:
        empty = [row for row in rows if not worksheet.cell(row, cols["Date"]).value]
        if empty:
            row, status = empty[0], "updated"
        else:
            row, status = worksheet.max_row + 1, "updated"
            for col in range(1, worksheet.max_column + 1):
                source, target = worksheet.cell(rows[0], col), worksheet.cell(row, col)
                target.value = source.value if col == S.vcol else None
                target._style = copy(source._style)
    worksheet.cell(row, cols["Date"], record["date"])
    worksheet.cell(row, cols["Shift"], record["shift"])
    worksheet.cell(row, cols["Meter Reading"], record.get("meter", ""))
    worksheet.cell(row, cols["Extracted Text"], record.get("text", ""))
    add_excel_image(record, row)
    return status

def workbook_bytes():
    buffer = io.BytesIO()
    S.wb.save(buffer)
    return buffer.getvalue()

# ---------------- UI ----------------
st.title("🚗 Smart Vehicle Data Automation")
tabs = st.tabs(["📊 Overview", "⬆️ Upload & Process", "✅ Verification", "📑 Excel", "🕘 History"])

with tabs[1]:
    col1, col2 = st.columns(2)
    selected_date = col1.date_input("Date", dt.date.today())
    selected_shift = col2.selectbox("Shift", ["Morning", "Evening", "Night"])
    excel_upload = st.file_uploader("1) Vehicle list Excel", ["xlsx", "xls"])
    if excel_upload and S.fname != excel_upload.name:
        load_excel(excel_upload)
    if S.ws is not None:
        headers = [str(cell.value) for cell in S.ws[1]]
        S.vcol = st.selectbox("Vehicle number column", range(1, len(headers) + 1), S.vcol - 1,
                              format_func=lambda i: headers[i - 1]) if headers else 1
        st.success(f"Loaded {S.fname}")

    uploads = st.file_uploader("2) WhatsApp export (.zip/.txt) or image files", [
        "zip", "txt", "jpg", "jpeg", "png", "webp", "bmp", "tif", "tiff", "gif"
    ], accept_multiple_files=True)
    if st.button("▶ Process", type="primary"):
        new_records, files = [], []
        progress = st.progress(0.0)
        for upload in uploads:
            if upload.name.lower().endswith(".zip"):
                try:
                    with zipfile.ZipFile(upload) as archive:
                        files.extend((os.path.basename(name), archive.read(name))
                                     for name in archive.namelist() if not name.endswith("/"))
                except zipfile.BadZipFile:
                    st.warning(f"Could not read {upload.name}: the ZIP file is damaged.")
            else:
                files.append((upload.name, upload.read()))

        caption_map = {}
        for name, data in files:
            if name.lower().endswith(".txt"):
                body = data.decode("utf-8-sig", errors="replace")
                caption_map.update(captions(body))
                new_records.extend(parse(body, name, selected_date, selected_shift))

        images = [(name, data) for name, data in files if name.lower().endswith(IMAGE_EXTS)]
        if not images:
            progress.progress(1.0)
        for index, (name, data) in enumerate(images):
            try:
                normalized = normalize_image(data)
            except Exception as error:
                st.warning(f"Skipped unreadable image {name}: {error}. Other photos will still be processed.")
                progress.progress((index + 1) / len(images))
                continue
            meter, extracted_text, errors = "", "", []
            try:
                meter = read_meter(normalized)
            except Exception as error:
                errors.append(f"meter OCR: {error}")
            try:
                extracted_text = read_image_text(normalized)
            except Exception as error:
                errors.append(f"text OCR: {error}")
            vehicle = caption_map.get(name, "")
            if not vehicle:
                vehicle = next((value for value in nums(extracted_text) if len(value) in (4, 6)), "")
            new_records.append(dict(ok=False, vehicle=vehicle, meter=meter, text=extracted_text,
                                    date=str(selected_date), shift=selected_shift, source=name, img=normalized))
            S.imgs += 1
            if errors:
                st.warning(f"OCR issue for {name}: {'; '.join(errors)}. The image is kept and can still be added to Excel.")
            elif not meter and not extracted_text:
                st.info(f"No text recognized in {name}; the photo is kept for adding to Excel.")
            progress.progress((index + 1) / len(images))
        S.recs.extend(new_records)
        S.hist.append(dict(time=dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                           event="processed", detail=f"{len(new_records)} records"))
        st.info(f"{len(new_records)} records ready for review. Select records and save them to Excel.")

with tabs[2]:
    if not S.recs:
        st.write("Nothing to verify.")
    else:
        select_all_col, deselect_col = st.columns(2)
        if select_all_col.button("☑ Select all records"):
            for record in S.recs:
                record["ok"] = True
            S.pop("ed", None)
            st.rerun()
        if deselect_col.button("☐ Deselect all records"):
            for record in S.recs:
                record["ok"] = False
            S.pop("ed", None)
            st.rerun()

        table = pd.DataFrame([{key: record.get(key, "") for key in
                               ("ok", "vehicle", "meter", "text", "date", "shift", "source")}
                              for record in S.recs])
        edited = st.data_editor(table, num_rows="dynamic", use_container_width=True, key="ed",
                                column_config={"ok": st.column_config.CheckboxColumn("Add to Excel")})
        image_indexes = [i for i, record in enumerate(S.recs) if record.get("img") is not None]
        selected_image = st.selectbox("Preview image", image_indexes,
                                      format_func=lambda i: f"{i}: {S.recs[i]['source']}") if image_indexes else None
        if selected_image is not None:
            st.image(S.recs[selected_image]["img"], width=400)
        duplicate_mode = st.radio("If duplicate vehicle + date + shift", ["Skip", "Update"], horizontal=True)
        save_col, reject_col = st.columns(2)
        if save_col.button("💾 Save selected records to Excel"):
            if S.ws is None:
                st.error("Upload Excel first.")
            else:
                saved = 0
                for index, row in edited.iterrows():
                    if not bool(row["ok"]):
                        continue
                    original = S.recs[index] if index < len(S.recs) else {}
                    record = {**row.to_dict(), "vehicle": norm(row.get("vehicle", "")),
                              "img": original.get("img"), "source": original.get("source", "")}
                    result = apply(record, duplicate_mode)
                    if result == "duplicate-skipped":
                        S.dups += 1
                        st.warning(f"Duplicate skipped: {record['vehicle']} {record['date']} {record['shift']}")
                        continue
                    S.hist.append(dict(time=dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), event=result,
                                       detail=f"{record['vehicle']} {record.get('meter', '')} {record['date']} {record['shift']}"))
                    saved += 1
                S.recs = [record for index, record in enumerate(S.recs)
                          if not (index in edited.index and bool(edited.loc[index, "ok"]))]
                st.success(f"Saved {saved} selected record(s). Check the Excel tab and download the updated workbook.")
                st.rerun()
        if reject_col.button("🗑 Reject all records"):
            S.recs = []
            S.pop("ed", None)
            st.rerun()

with tabs[3]:
    if S.ws is None:
        st.write("Upload an Excel file first.")
    else:
        st.subheader("Original")
        st.dataframe(S.orig, use_container_width=True)
        st.subheader("Updated")
        st.dataframe(pd.DataFrame(S.ws.values), use_container_width=True)
        output_name = "updated_" + (S.fname[:-4] + ".xlsx" if S.fname.lower().endswith(".xls") else S.fname)
        st.download_button("⬇ Download updated Excel", workbook_bytes(), output_name)
        st.download_button("⬇ Unmatched report", pd.DataFrame({"unmatched_vehicle": S.unmatched}).to_csv(index=False), "unmatched.csv")
        st.download_button("⬇ Processing report", pd.DataFrame(S.hist).to_csv(index=False), "report.csv")

with tabs[4]:
    st.dataframe(pd.DataFrame(S.hist), use_container_width=True)

with tabs[0]:
    events = [entry["event"] for entry in S.hist]
    total = S.ws.max_row - 1 if S.ws is not None else 0
    metrics = st.columns(5)
    metrics[0].metric("Vehicle rows", total)
    metrics[1].metric("Updated", events.count("updated") + events.count("duplicate-updated"))
    metrics[2].metric("Images processed", S.imgs)
    metrics[3].metric("Records awaiting review", len(S.recs))
    metrics[4].metric("Duplicates skipped", S.dups)
