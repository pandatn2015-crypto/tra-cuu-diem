
import streamlit as st
import pandas as pd
from pathlib import Path
import re
import io

st.set_page_config(
    page_title="Tra cứu điểm học sinh",
    page_icon="📘",
    layout="wide"
)

st.title("📘 TRA CỨU ĐIỂM HỌC SINH")
st.caption(
    "Tra cứu điểm bằng Mã định danh Bộ GD&ĐT."
)

# Đọc các file Excel nằm cùng cấp với home.py
ROOT = Path(__file__).resolve().parent


def norm(value):
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().casefold()


def clean(value):
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def read_workbook(path: Path):
    records = []
    xls = pd.ExcelFile(path, engine="openpyxl")
    class_name = path.stem

    for sheet_name in xls.sheet_names:
        raw = pd.read_excel(
            path,
            sheet_name=sheet_name,
            header=None,
            dtype=object,
            engine="openpyxl"
        )

        header_row = None
        id_col = None

        for r in range(min(len(raw), 20)):
            for c in range(raw.shape[1]):
                cell = norm(raw.iat[r, c])
                if "mã định danh" in cell and (
                    "gd" in cell or "đt" in cell
                ):
                    header_row, id_col = r, c
                    break
            if header_row is not None:
                break

        if header_row is None:
            continue

        top_headers = raw.iloc[header_row]
        sub_headers = (
            raw.iloc[header_row + 1]
            if header_row + 1 < len(raw)
            else pd.Series(dtype=object)
        )

        headers = []
        current_group = ""

        for c in range(raw.shape[1]):
            top = clean(top_headers.iloc[c])
            sub = clean(sub_headers.iloc[c])

            top = re.sub(r"\s+", " ", top.replace("\n", " ")).strip()
            sub = re.sub(r"\s+", " ", sub.replace("\n", " ")).strip()

            if c >= 5 and top:
                current_group = top
            elif c >= 5 and not top and current_group:
                top = current_group

            if top and sub and norm(top) != norm(sub):
                label = f"{top} - {sub}"
            else:
                label = top or sub or f"Cột {c + 1}"

            headers.append(re.sub(r"\s+", " ", label).strip())

        name_col = next(
            (
                c for c in range(raw.shape[1])
                if "họ và tên" in norm(raw.iat[header_row, c])
            ),
            2
        )

        dob_col = next(
            (
                c for c in range(raw.shape[1])
                if "ngày sinh" in norm(raw.iat[header_row, c])
            ),
            3
        )

        gender_col = next(
            (
                c for c in range(raw.shape[1])
                if "giới tính" in norm(raw.iat[header_row, c])
            ),
            4
        )

        for r in range(header_row + 2, len(raw)):
            student_id = clean(raw.iat[r, id_col])

            if not student_id or not re.search(r"\d", student_id):
                continue

            row = raw.iloc[r]
            scores = {}

            for c in range(raw.shape[1]):
                if c <= max(id_col, name_col, dob_col, gender_col):
                    continue

                value = row.iloc[c]

                if value is None or pd.isna(value) or clean(value) == "":
                    continue

                label = headers[c]
                if label and not label.startswith("Cột "):
                    scores[label] = clean(value)

            records.append({
                "Mã định danh": student_id,
                "Họ và tên": clean(row.iloc[name_col]),
                "Ngày sinh": clean(row.iloc[dob_col]),
                "Giới tính": clean(row.iloc[gender_col]),
                "Lớp": class_name,
                "Môn học": sheet_name,
                "Điểm": scores,
            })

    return records


with st.sidebar:
    st.header("⚙️ Dữ liệu")
    st.caption("Đọc file Excel cùng cấp với home.py.")
    refresh = st.button(
        "🔄 Đọc lại dữ liệu",
        use_container_width=True
    )

if refresh:
    st.cache_data.clear()
    st.rerun()


@st.cache_data(show_spinner=False)
def load_all_data(file_paths, file_mtimes):
    all_records = []
    errors = []

    for file_path in file_paths:
        try:
            all_records.extend(read_workbook(Path(file_path)))
        except Exception as e:
            errors.append(f"{Path(file_path).name}: {e}")

    return all_records, errors


# Chỉ đọc Excel cùng cấp, bỏ qua file tạm và file đầu ra
files = sorted(
    p for p in ROOT.glob("*.xlsx")
    if not p.name.startswith("~$")
    and not p.name.startswith("ket_qua_")
)

if not files:
    st.warning(
        "Chưa tìm thấy file Excel. "
        "Hãy tải các file lớp như 6A1.xlsx, 6A2.xlsx "
        "lên cùng cấp với home.py trên GitHub."
    )
    st.stop()

file_paths = tuple(str(p) for p in files)
file_mtimes = tuple(p.stat().st_mtime for p in files)

with st.spinner("Đang đọc dữ liệu điểm..."):
    records, errors = load_all_data(file_paths, file_mtimes)

if errors:
    with st.expander("⚠️ File gặp lỗi khi đọc"):
        for err in errors:
            st.warning(err)

if not records:
    st.error(
        "Không tìm thấy dữ liệu học sinh. "
        "Kiểm tra cột 'Mã định danh Bộ GD&ĐT' trong các sheet."
    )
    st.stop()

st.success(
    f"Đã đọc {len(files)} file lớp và tìm thấy "
    f"{len(set(norm(x['Mã định danh']) for x in records))} mã định danh."
)

query = st.text_input(
    "🔎 Nhập Mã định danh Bộ GD&ĐT",
    placeholder="Nhập mã định danh học sinh..."
).strip()

if query:
    matches = [
        r for r in records
        if norm(r["Mã định danh"]) == norm(query)
    ]

    if not matches:
        st.error("Không tìm thấy học sinh có mã định danh này.")
    else:
        first = matches[0]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Họ và tên", first["Họ và tên"] or "—")
        c2.metric(
            "Lớp",
            ", ".join(sorted(set(r["Lớp"] for r in matches)))
        )
        c3.metric("Ngày sinh", first["Ngày sinh"] or "—")
        c4.metric("Giới tính", first["Giới tính"] or "—")

        st.subheader("📚 Điểm theo từng môn")

        for rec in matches:
            with st.expander(
                f"{rec['Môn học']} — lớp {rec['Lớp']}",
                expanded=True
            ):
                if rec["Điểm"]:
                    df = pd.DataFrame([
                        {"Cột điểm / nội dung": k, "Kết quả": v}
                        for k, v in rec["Điểm"].items()
                    ])
                    st.dataframe(
                        df,
                        hide_index=True,
                        use_container_width=True
                    )
                else:
                    st.info("Sheet môn học này chưa có điểm.")

        export_rows = []

        for rec in matches:
            row = {
                "Mã định danh": rec["Mã định danh"],
                "Họ và tên": rec["Họ và tên"],
                "Lớp": rec["Lớp"],
                "Môn học": rec["Môn học"],
            }
            row.update(rec["Điểm"])
            export_rows.append(row)

        out = io.BytesIO()

        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            pd.DataFrame(export_rows).to_excel(
                writer,
                index=False,
                sheet_name="KetQua"
            )

        st.download_button(
            "⬇️ Xuất kết quả học sinh ra Excel",
            data=out.getvalue(),
            file_name=f"ket_qua_{clean(query)}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with st.expander("📁 Các file đang được đọc"):
    st.write([p.name for p in files])
