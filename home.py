import streamlit as st
import pandas as pd
from pathlib import Path
import re
import io
import html

# =====================================================
# CẤU HÌNH TRANG
# =====================================================

st.set_page_config(
    page_title="Tra cứu điểm - THCS Lý Tự Trọng",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded"
)

ROOT = Path(__file__).resolve().parent

# =====================================================
# CSS GIAO DIỆN
# =====================================================

st.markdown("""
<style>

/* Nền trang */
.stApp {
    background-color: #f4f7fb;
}

/* Giới hạn nội dung */
.block-container {
    padding-top: 2rem;
    padding-bottom: 3rem;
    max-width: 1250px;
}

/* Tiêu đề chính */
.main-title {
    color: #174b85;
    font-size: 28px;
    font-weight: 800;
    text-align: center;
    margin-bottom: 5px;
}

.sub-title {
    color: #64748b;
    text-align: center;
    font-size: 15px;
    margin-bottom: 25px;
}

/* Thanh tiêu đề */
.header-box {
    background: linear-gradient(120deg, #174b85, #287ac2);
    padding: 25px;
    border-radius: 15px;
    color: white;
    text-align: center;
    margin-bottom: 25px;
    box-shadow: 0 5px 15px rgba(23,75,133,0.15);
}

.header-box h1 {
    color: white;
    font-size: 27px;
    font-weight: 800;
    margin: 0;
}

.header-box p {
    color: #e5efff;
    margin-top: 8px;
    margin-bottom: 0;
    font-size: 15px;
}

/* Thông tin học sinh */
.student-card {
    background: white;
    padding: 22px 25px;
    border-radius: 14px;
    border-left: 6px solid #287ac2;
    box-shadow: 0 3px 12px rgba(0,0,0,0.05);
    margin: 15px 0 20px 0;
}

.student-label {
    color: #64748b;
    font-size: 13px;
    font-weight: 600;
    letter-spacing: 1px;
    margin-bottom: 5px;
}

.student-name {
    color: #174b85;
    font-size: 27px;
    font-weight: 800;
    line-height: 1.5;
    overflow-wrap: anywhere;
}

/* Thẻ thông tin */
.info-card {
    background: white;
    padding: 16px;
    border-radius: 12px;
    text-align: center;
    border: 1px solid #e2e8f0;
    box-shadow: 0 2px 8px rgba(0,0,0,0.03);
    min-height: 95px;
}

.info-label {
    color: #64748b;
    font-size: 13px;
    margin-bottom: 8px;
}

.info-value {
    color: #174b85;
    font-size: 19px;
    font-weight: 750;
    overflow-wrap: anywhere;
}

/* Khu vực tra cứu */
.search-box {
    background: white;
    padding: 20px;
    border-radius: 14px;
    border: 1px solid #e2e8f0;
    margin-bottom: 20px;
}

/* Tiêu đề phần */
.section-title {
    color: #174b85;
    font-size: 20px;
    font-weight: 750;
    margin-top: 25px;
    margin-bottom: 12px;
}

/* Nút */
.stButton > button {
    border-radius: 9px;
    font-weight: 650;
    min-height: 42px;
}

.stDownloadButton > button {
    background-color: #16834a;
    color: white;
    border-radius: 9px;
    font-weight: 700;
    min-height: 45px;
    border: none;
}

.stDownloadButton > button:hover {
    background-color: #11683b;
    color: white;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background-color: #ffffff;
    border-right: 1px solid #e2e8f0;
}

/* Bảng điểm */
div[data-testid="stDataFrame"] {
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    overflow: hidden;
}

/* Footer */
.footer {
    text-align: center;
    color: #94a3b8;
    font-size: 13px;
    padding-top: 25px;
    margin-top: 40px;
    border-top: 1px solid #e2e8f0;
}

@media (max-width: 768px) {
    .header-box h1 {
        font-size: 21px;
    }

    .student-name {
        font-size: 22px;
    }

    .block-container {
        padding-left: 1rem;
        padding-right: 1rem;
    }
}
/* Lưới hiển thị điểm tự co giãn */
.score-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 10px;
    width: 100%;
    margin-top: 10px;
}

.score-item {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 12px;
    min-width: 0;
    box-shadow: 0 2px 5px rgba(0,0,0,0.03);
}

.score-label {
    color: #64748b;
    font-size: 13px;
    font-weight: 600;
    overflow-wrap: anywhere;
    margin-bottom: 8px;
}

.score-value {
    color: #174b85;
    font-size: 20px;
    font-weight: 800;
    overflow-wrap: anywhere;
}

/* Điện thoại */
@media (max-width: 480px) {
    .score-grid {
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 7px;
    }

    .score-item {
        padding: 10px;
    }

    .score-label {
        font-size: 12px;
    }

    .score-value {
        font-size: 18px;
    }
}
</style>
""", unsafe_allow_html=True)


# =====================================================
# HÀM XỬ LÝ DỮ LIỆU
# =====================================================

def norm(value):
    if value is None or pd.isna(value):
        return ""

    return re.sub(
        r"\s+", " ", str(value)
    ).strip().casefold()


def clean(value):
    if value is None or pd.isna(value):
        return ""

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value).strip()


# =====================================================
# ĐỌC FILE EXCEL
# =====================================================

def read_workbook(path: Path):

    records = []

    xls = pd.ExcelFile(
        path,
        engine="openpyxl"
    )

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

        # Tìm dòng tiêu đề
        for r in range(min(len(raw), 20)):

            for c in range(raw.shape[1]):

                cell = norm(raw.iat[r, c])

                if "mã định danh" in cell and (
                    "gd" in cell or "đt" in cell
                ):
                    header_row = r
                    id_col = c
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

            top = re.sub(
                r"\s+", " ", top.replace("\n", " ")
            ).strip()

            sub = re.sub(
                r"\s+", " ", sub.replace("\n", " ")
            ).strip()

            if c >= 5 and top:
                current_group = top

            elif c >= 5 and not top and current_group:
                top = current_group

            if top and sub and norm(top) != norm(sub):
                label = f"{top} - {sub}"
            else:
                label = top or sub or f"Cột {c + 1}"

            headers.append(
                re.sub(r"\s+", " ", label).strip()
            )

        # Xác định cột thông tin
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

        # Đọc từng học sinh
        for r in range(header_row + 2, len(raw)):

            student_id = clean(raw.iat[r, id_col])

            if not student_id or not re.search(r"\d", student_id):
                continue

            row = raw.iloc[r]
            scores = {}

            for c in range(raw.shape[1]):

                if c <= max(
                    id_col,
                    name_col,
                    dob_col,
                    gender_col
                ):
                    continue

                value = row.iloc[c]

                if value is None or pd.isna(value):
                    continue

                if clean(value) == "":
                    continue

                label = headers[c]

                if label and not label.startswith("Cột "):
                    scores[label] = clean(value)

            records.append({

                "Mã định danh": student_id,

                "Họ và tên": clean(
                    row.iloc[name_col]
                ),

                "Ngày sinh": clean(
                    row.iloc[dob_col]
                ),

                "Giới tính": clean(
                    row.iloc[gender_col]
                ),

                "Lớp": class_name,

                "Môn học": sheet_name,

                "Điểm": scores

            })

    return records


# =====================================================
# SIDEBAR
# =====================================================

logo1_path = ROOT / "logoltt.jpg"
logo2_path = ROOT / "K.png"

with st.sidebar:

    col1, col2 = st.columns(2)

    with col1:
        if logo1_path.exists():
            st.image(
                str(logo1_path),
                use_container_width=True
            )

    with col2:
        if logo2_path.exists():
            st.image(
                str(logo2_path),
                use_container_width=True
            )

    st.markdown(
        """
        <div style="text-align:center; padding:10px 0;">
            <div style="font-size:19px; font-weight:800; color:#174b85;">
                THCS LÝ TỰ TRỌNG
            </div>
            <div style="font-size:14px; color:#64748b;">
                PHƯỜNG LONG HOA
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.divider()

    st.markdown("### ⚙️ QUẢN LÝ DỮ LIỆU")

    st.caption(
        "Hệ thống tự động đọc các file Excel "
        "được đặt cùng thư mục với home.py."
    )

    refresh = st.button(
        "🔄 ĐỌC LẠI DỮ LIỆU",
        use_container_width=True
    )

    if refresh:
        st.cache_data.clear()
        st.rerun()

    st.divider()

    st.markdown("### 📂 CẤU TRÚC DỮ LIỆU")

    st.caption(
        """
        Các file Excel được đặt cùng cấp với home.py.

        Ví dụ:

        - 6A1.xlsx
        - 6A2.xlsx
        - 7A1.xlsx
        - 8A1.xlsx
        - 9A1.xlsx
        """
    )

    st.divider()

    st.caption(
        "© THCS Lý Tự Trọng - Phường Long Hoa"
    )


# =====================================================
# TẢI TOÀN BỘ DỮ LIỆU
# =====================================================

@st.cache_data(show_spinner=False)
def load_all_data(file_paths, file_mtimes):

    all_records = []
    errors = []

    for file_path in file_paths:

        try:
            all_records.extend(
                read_workbook(Path(file_path))
            )

        except Exception as e:

            errors.append(
                f"{Path(file_path).name}: {e}"
            )

    return all_records, errors


files = sorted(
    p for p in ROOT.glob("*.xlsx")
    if not p.name.startswith("~$")
    and not p.name.startswith("ket_qua_")
)

if not files:

    st.warning(
        "Chưa tìm thấy file Excel. "
        "Hãy tải các file lớp như 6A1.xlsx, 6A2.xlsx "
        "lên cùng cấp với home.py."
    )

    st.stop()


file_paths = tuple(str(p) for p in files)

file_mtimes = tuple(
    p.stat().st_mtime for p in files
)

with st.spinner("⏳ Đang tải dữ liệu điểm học sinh..."):

    records, errors = load_all_data(
        file_paths,
        file_mtimes
    )


if errors:

    with st.expander("⚠️ Các file gặp lỗi khi đọc"):

        for err in errors:
            st.warning(err)


if not records:

    st.error(
        "Không tìm thấy dữ liệu học sinh. "
        "Vui lòng kiểm tra cột Mã định danh Bộ GD&ĐT."
    )

    st.stop()


total_students = len(
    set(norm(x["Mã định danh"]) for x in records)
)


# =====================================================
# HEADER TRANG CHỦ
# =====================================================

st.markdown(
    """
    <div class="header-box">
        <h1>📘 TRA CỨU ĐIỂM HỌC SINH</h1>
        <h1> TRƯỜNG THCS LÝ TỰ TRỌNG - PHƯỜNG LONG HOA</h1>
    </div>
    """,
    unsafe_allow_html=True
)


# =====================================================
# THỐNG KÊ DỮ LIỆU
# =====================================================

st.markdown(
    '<div class="section-title">📊 THÔNG TIN HỆ THỐNG</div>',
    unsafe_allow_html=True
)

col1, col2, col3 = st.columns(3)

with col1:

    st.markdown(
        f"""
        <div class="info-card">
            <div class="info-label">📁 SỐ FILE LỚP</div>
            <div class="info-value">{len(files)}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with col2:

    st.markdown(
        f"""
        <div class="info-card">
            <div class="info-label">👨‍🎓 SỐ HỌC SINH</div>
            <div class="info-value">{total_students:,}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

with col3:

    st.markdown(
        f"""
        <div class="info-card">
            <div class="info-label">📋 SỐ BẢN GHI</div>
            <div class="info-value">{len(records):,}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


# =====================================================
# KHU VỰC TRA CỨU
# =====================================================

st.markdown(
    '<div class="section-title">🔎 TRA CỨU THÔNG TIN KẾT QUẢ HỌC TẬP CỦA HỌC SINH</div>',
    unsafe_allow_html=True
)

# st.markdown(
#     '<div class="search-box">',
#     unsafe_allow_html=True
# )

query = st.text_input(
    "Nhập Mã định danh Bộ GD&ĐT",
    placeholder="Ví dụ: 012345678901",
    help="Nhập chính xác mã định danh của học sinh."
).strip()

st.markdown(
    """
    <div style="font-size:13px; color:#64748b; margin-top:5px;">
        💡 Nhập mã định danh Bộ GD và ĐT để xem thông tin và kết quả học tập.
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("</div>", unsafe_allow_html=True)


# =====================================================
# HIỂN THỊ KẾT QUẢ
# =====================================================

if query:

    matches = [
        r for r in records
        if norm(r["Mã định danh"]) == norm(query)
    ]

    if not matches:

        st.error(
            "❌ Không tìm thấy học sinh có mã định danh này. "
            "Vui lòng kiểm tra lại thông tin."
        )

    else:

        first = matches[0]

        # ---------------------------------------------
        # THÔNG TIN CÁ NHÂN
        # ---------------------------------------------

        st.markdown(
            '<div class="section-title">👨‍🎓 THÔNG TIN HỌC SINH</div>',
            unsafe_allow_html=True
        )

        safe_name = html.escape(
            first["Họ và tên"] or "—"
        )

        st.markdown(
            f"""
            <div class="student-card">
                <div class="student-label">
                    HỌ VÀ TÊN HỌC SINH
                </div>
                <div class="student-name">
                    {safe_name}
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        # Các lớp
        classes = ", ".join(
            sorted(set(r["Lớp"] for r in matches))
        )

        c1, c2, c3 = st.columns(3)

        with c1:

            st.markdown(
                f"""
                <div class="info-card">
                    <div class="info-label">🏫 LỚP</div>
                    <div class="info-value">
                        {html.escape(classes or "—")}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

        with c2:

            st.markdown(
                f"""
                <div class="info-card">
                    <div class="info-label">🎂 NGÀY SINH</div>
                    <div class="info-value">
                        {html.escape(first["Ngày sinh"] or "—")}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

        with c3:

            st.markdown(
                f"""
                <div class="info-card">
                    <div class="info-label">👤 GIỚI TÍNH</div>
                    <div class="info-value">
                        {html.escape(first["Giới tính"] or "—")}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

        
                # ---------------------------------------------
        # ĐIỂM THEO TỪNG MÔN
        # ---------------------------------------------

        st.markdown(
            '<div class="section-title">📚 KẾT QUẢ HỌC TẬP</div>',
            unsafe_allow_html=True
        )

        st.caption(
            f"Tìm thấy {len(matches)} bảng điểm tương ứng với học sinh."
        )

        for i, rec in enumerate(matches):

            subject = rec["Môn học"]
            class_name = rec["Lớp"]

            with st.expander(
                f"📖 {subject}  |  Lớp {class_name}",
                expanded=True
            ):

                if rec["Điểm"]:

                    score_items = []

                    for k, v in rec["Điểm"].items():

                        ten_cot = html.escape(str(k))
                        gia_tri = html.escape(str(v))

                        score_items.append(
                            f'<div class="score-item">'
                            f'<div class="score-label">{ten_cot}</div>'
                            f'<div class="score-value">{gia_tri}</div>'
                            f'</div>'
                        )

                    score_html = (
                        '<div class="score-grid">'
                        + "".join(score_items)
                        + '</div>'
                    )

                    st.markdown(
                        score_html,
                        unsafe_allow_html=True
                    )

                else:

                    st.info("Chưa có dữ liệu điểm cho môn học này.")

        # ---------------------------------------------
        # XUẤT EXCEL
        # ---------------------------------------------

        st.markdown(
            '<div class="section-title">📥 TẢI KẾT QUẢ HỌC TẬP</div>',
            unsafe_allow_html=True
        )

        export_rows = []

        for rec in matches:

            row = {

                "Mã định danh": rec["Mã định danh"],

                "Họ và tên": rec["Họ và tên"],

                "Lớp": rec["Lớp"],

                "Môn học": rec["Môn học"]

            }

            row.update(rec["Điểm"])

            export_rows.append(row)

        out = io.BytesIO()

        with pd.ExcelWriter(
            out,
            engine="openpyxl"
        ) as writer:

            pd.DataFrame(export_rows).to_excel(
                writer,
                index=False,
                sheet_name="KetQua"
            )

        st.download_button(
            label="⬇️ TẢI KẾT QUẢ HỌC TẬP EXCEL",
            data=out.getvalue(),
            file_name=f"ket_qua_{clean(query)}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        st.success(
            "Tra cứu thành công! "
            "Bạn có thể tải kết quả học tập về máy tính."
        )


# =====================================================
# THÔNG TIN FILE DỮ LIỆU
# =====================================================

with st.expander("📂 DANH SÁCH FILE EXCEL ĐANG SỬ DỤNG"):

    for p in files:

        st.write(f"📄 {p.name}")


# =====================================================
# FOOTER
# =====================================================

st.markdown(
    """
    <div class="footer">
        <b>TRƯỜNG THCS LÝ TỰ TRỌNG - PHƯỜNG LONG HOA</b>
        <br>
        Hệ thống tra cứu kết quả học tập học sinh miễn phí, không thương mại
        <br>
        © 2026 - Bản quyền thuộc TK và nhà trường
    </div>
    """,
    unsafe_allow_html=True
)
