import streamlit as st
import pdfplumber
import pandas as pd
import re
import os

# ---------- helpers ----------
def clean_sheet_name(name: str) -> str:
    """Excel sheet names must be <=31 chars and cannot contain : \\ / ? * [ ]"""
    name = re.sub(r'[:\\/?*\[\]]', '', (name or '').strip())
    return name[:31] or "Table"

def horiz_overlap_ratio(b1, b2):
    """Compute horizontal overlap ratio between two bounding boxes"""
    overlap = max(0, min(b1[2], b2[2]) - max(b1[0], b2[0]))
    w1, w2 = (b1[2] - b1[0]), (b2[2] - b2[0])
    base = min(w1, w2) if min(w1, w2) > 0 else 1
    return overlap / base

def normalize_rows(rows, ncols, title=None, columns=None):
    """Fix rows to equal number of columns, drop fully empty rows, skip repeats."""
    fixed = []
    for r in rows:
        r = [(c if c is not None else "") for c in r]

        if len(r) < ncols:
            r = r + [""] * (ncols - len(r))
        else:
            r = r[:ncols]

        if not any(str(x).strip() for x in r):
            continue
        if title and " ".join(r).strip() == " ".join(title).strip():
            continue
        if columns and [c.strip() for c in r] == [c.strip() for c in columns]:
            continue

        fixed.append(r)
    return fixed

# ---------- main extractor ----------
def pdf_to_excel_pair_header_with_body(
    pdf_path: str,
    excel_path: str,
    keep_width_frac: float = 1.0,
    keep_height_frac: float = 1.0,
    overlap_threshold: float = 0.5
):
    sheets = {}

    with pdfplumber.open(pdf_path) as pdf:
        for pageno, page in enumerate(pdf.pages, start=1):
            if keep_width_frac < 1.0 or keep_height_frac < 1.0:
                W, H = page.width, page.height
                page = page.crop((0, 0, W * keep_width_frac, H * keep_height_frac))

            tables = page.find_tables()
            items = []
            for t in tables:
                data = t.extract()
                if data and any(any(cell for cell in row) for row in data):
                    items.append({"bbox": t.bbox, "data": data})

            items.sort(key=lambda it: it["bbox"][1])
            used = set()
            i = 0
            while i < len(items):
                if i in used:
                    i += 1
                    continue

                head = items[i]
                header_rows = head["data"]

                if len(header_rows) >= 2:
                    title = " ".join([c for c in header_rows[0] if c])
                    cols_raw = header_rows[1]

                    body_idx = None
                    for j in range(i + 1, len(items)):
                        if j in used:
                            continue
                        b = items[j]
                        below = b["bbox"][1] >= head["bbox"][3] - 2
                        overlap = horiz_overlap_ratio(head["bbox"], b["bbox"])
                        if below and overlap >= overlap_threshold:
                            body_idx = j
                            break
                    if body_idx is None:
                        i += 1
                        continue

                    cols = [c if c is not None else "" for c in cols_raw]
                    ncols = len(cols)
                    body_rows = []

                    j = body_idx
                    while j < len(items):
                        if j in used:
                            j += 1
                            continue
                        cand = items[j]
                        below = cand["bbox"][1] >= head["bbox"][3] - 2
                        overlap = horiz_overlap_ratio(head["bbox"], cand["bbox"])
                        is_new_header = (
                            len(cand["data"]) >= 2
                            and (cand["bbox"][1] - head["bbox"][3]) > 5
                            and overlap < overlap_threshold
                        )
                        if below and overlap >= overlap_threshold and not is_new_header:
                            body_rows.extend(cand["data"])
                            used.add(j)
                            j += 1
                        else:
                            break

                    body_rows = normalize_rows(body_rows, ncols, title=title.split(), columns=cols)
                    if not body_rows:
                        i += 1
                        continue

                    df = pd.DataFrame(body_rows, columns=cols)

                    sheet_key = (clean_sheet_name(title), tuple(cols))
                    if sheet_key in sheets:
                        sheets[sheet_key] = pd.concat([sheets[sheet_key], df], ignore_index=True)
                    else:
                        sheets[sheet_key] = df

                    used.add(i)
                    i += 1
                else:
                    i += 1

    if not sheets:
        return None

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        for (title, _cols), df in sheets.items():
            sheet_name = clean_sheet_name(title)
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    return excel_path


# ---------- Streamlit UI ----------
st.set_page_config(
    page_title="PDF to Excel Converter",
    page_icon="pdf.png",
    layout="centered"
)

st.title("📑 PDF to Excel Converter")
st.write(
    "Easily extract structured tables from your PDF documents "
    "and export them into clean, well-formatted Excel files. "
    "Simply upload your PDF and download the converted result in seconds."
)

uploaded_pdf = st.file_uploader("📂 Upload your PDF file", type=["pdf"])

if uploaded_pdf is not None:
    with open("temp.pdf", "wb") as f:
        f.write(uploaded_pdf.getbuffer())

    st.success("✅ PDF uploaded successfully!")

    if st.button("Convert to Excel"):
        # derive Excel file name from PDF
        pdf_filename = uploaded_pdf.name
        base_name = os.path.splitext(pdf_filename)[0]
        excel_file = f"{base_name}.xlsx"

        result = pdf_to_excel_pair_header_with_body("temp.pdf", excel_file)

        if result:
            with open(result, "rb") as f:
                st.download_button("⬇️ Download Excel", f, file_name=excel_file)
            st.success(f"✅ Conversion complete! Saved as {excel_file}")
        else:
            st.warning("⚠️ No tables found in the PDF.")
