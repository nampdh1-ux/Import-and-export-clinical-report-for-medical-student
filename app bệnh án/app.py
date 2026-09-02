import streamlit as st
from fpdf import FPDF
from datetime import datetime
import os
import tempfile
import json

# --- CẤU HÌNH TRANG ĐẦU TIÊN ---
st.set_page_config(page_title="Công cụ làm Bệnh án Lâm sàng", layout="wide")

# --- CSS TÙY BIẾN DẢI ĐỀ MỤC THEO MÀU XANH AMBOSS KẾT HỢP VẠCH KÉP HMU ---
st.markdown("""
<style>
    /* Khung expander tổng thể */
    div[data-testid="stExpander"] {
        border: 1px solid #d4eaf0;
        border-radius: 6px;
        margin-bottom: 12px;
        background-color: #ffffff;
    }
    
    /* Thanh tiêu đề expander: Màu nền xanh AMBOSS (#EBF7F9) + Viền kép HMU */
    div[data-testid="stExpander"] > details > summary {
        background-color: #ebf7f9 !important;
        /* Vạch kép HMU: 4px hồng đậm (#c2185b) và 4px xanh navy (#0d47a1) */
        box-shadow: inset 4px 0 0 #c2185b, inset 8px 0 0 #0d47a1 !important;
        border-left: none !important;
        border-radius: 5px 5px 0 0;
        padding: 10px 14px 10px 18px !important;
        font-weight: 700 !important;
        color: #06445c !important;
        font-size: 1.05rem !important;
    }
    
    div[data-testid="stExpander"] > details > summary:hover {
        background-color: #ddf2f5 !important;
        color: #032b3b !important;
    }

    /* Khối đề mục ở sidebar theo tone AMBOSS */
    .sidebar-header-amboss {
        background-color: #ebf7f9;
        box-shadow: inset 4px 0 0 #c2185b, inset 8px 0 0 #0d47a1;
        padding: 8px 12px 8px 16px;
        border-radius: 4px;
        font-size: 1.05rem;
        font-weight: 700;
        color: #06445c;
        margin-bottom: 8px;
    }

    /* Tiểu mục con bên trong các expander: Tone xanh pastel nhạt dịu mắt */
    .sub-section-header {
        background-color: #f2fafb;
        box-shadow: inset 3px 0 0 #c2185b, inset 6px 0 0 #0d47a1;
        padding: 6px 12px 6px 14px;
        border-radius: 3px;
        margin-top: 10px;
        margin-bottom: 8px;
        font-size: 0.95rem;
        font-weight: 600;
        color: #0c4d63;
    }

    /* Chữ đỏ đậm y khoa cho chẩn đoán xác định */
    .highlight-dx {
        color: #b40000;
        font-weight: bold;
        font-size: 1.02rem;
    }
</style>
""", unsafe_allow_html=True)

# --- DANH MỤC CÁC TRƯỜNG DỮ LIỆU ---
default_fields = [
    "ho_ten", "tuoi", "gioi_tinh", "dan_tok", "nghe_nghiep", "khoa_phong", "dia_chi", "ngay_vao_vien", "sinh_vien",
    "ly_do_vao_vien", "benh_su", "ts_noi_khoa", "ts_ngoai_khoa", "ts_loi_song", "ts_gia_dinh",
    "kham_vao_vien", "kham_toan_than", "kham_tuan_hoan", "kham_ho_hap", "kham_tieu_hoa",
    "kham_than_kinh", "kham_tiet_nieu", "kham_co_xuong_khop", "kham_co_quan_khac",
    "tom_tat", "chan_doan_so_bo", "chan_doan_phan_biet", "bien_luan",
    "cls_dx_xac_dinh", "cls_dx_dieu_tri", "cls_dx_khac",
    "chan_doan_xac_dinh", "bien_luan_xac_dinh", "dt_muc_tieu", "dt_cu_the", "dt_theo_doi",
    "tien_luong", "tu_van"
]

if "so_hang_cls" not in st.session_state:
    st.session_state["so_hang_cls"] = 3

for field in default_fields:
    if field not in st.session_state:
        if field == "tuoi":
            st.session_state[field] = 45
        elif field == "gioi_tinh":
            st.session_state[field] = "Nam"
        elif field == "dan_tok":
            st.session_state[field] = "Kinh"
        elif field == "ngay_vao_vien":
            st.session_state[field] = datetime.now().strftime("%d/%m/%Y %H:%M")
        else:
            st.session_state[field] = ""

for i in range(st.session_state["so_hang_cls"]):
    if f"cls_kq_{i}" not in st.session_state:
        st.session_state[f"cls_kq_{i}"] = ""
    if f"cls_pg_{i}" not in st.session_state:
        st.session_state[f"cls_pg_{i}"] = ""

# --- HÀM HỖ TRỢ ĐỊNH DẠNG BULLET POINTS TỰ ĐỘNG ---
def format_bullet_points(text):
    if not text or not str(text).strip():
        return "Chưa ghi nhận thông tin."
    lines = str(text).strip().split("\n")
    formatted_lines = []
    for line in lines:
        cleaned = line.strip()
        if cleaned:
            if not cleaned.startswith("-") and not cleaned.startswith("*"):
                formatted_lines.append(f"- {cleaned}")
            else:
                formatted_lines.append(cleaned)
    return "\n".join(formatted_lines)

# --- CLASS TẠO PDF BỆNH ÁN ---
class BenhAnPDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.add_font("Roboto", "", "Roboto-Regular.ttf")
        self.add_font("Roboto-Bold", "", "Roboto-Bold.ttf")

    def header(self):
        if self.page_no() == 1:
            self.set_font("Roboto-Bold", "", 15)
            self.cell(0, 8, "BỆNH ÁN LÂM SÀNG", align="C", new_x="LMARGIN", new_y="NEXT")
            self.set_font("Roboto", "", 8)
            self.cell(0, 4, f"Thời gian lập hồ sơ: {datetime.now().strftime('%d/%m/%Y %H:%M')}", align="C", new_x="LMARGIN", new_y="NEXT")
            self.ln(3)

    def footer(self):
        self.set_y(-12)
        self.set_font("Roboto", "", 8)
        self.cell(0, 10, f"Trang {self.page_no()}/{{nb}}", align="C")

    def add_section_header(self, title):
        self.set_font("Roboto-Bold", "", 11)
        self.set_fill_color(225, 235, 245)
        self.cell(0, 7, title, fill=True, new_x="LMARGIN", new_y="NEXT")
        self.ln(1)

    def add_subsection_header(self, title):
        self.set_font("Roboto-Bold", "", 10)
        self.cell(0, 6, title, new_x="LMARGIN", new_y="NEXT")

    def add_body_text(self, text):
        self.set_font("Roboto", "", 9.5)
        self.set_text_color(0, 0, 0)
        self.multi_cell(0, 5, str(text).strip() if str(text).strip() else "Chưa ghi nhận thông tin.")
        self.ln(2)

    def add_highlight_text(self, text):
        self.set_font("Roboto-Bold", "", 10.5)
        self.set_text_color(180, 0, 0)
        self.multi_cell(0, 5.5, str(text).strip() if str(text).strip() else "Chưa ghi nhận thông tin.")
        self.set_text_color(0, 0, 0)
        self.ln(2)

    def render_tom_tat_pdf(self, text):
        if not text or not str(text).strip():
            self.add_body_text("Chưa ghi nhận thông tin.")
            return

        lines = [line.strip() for line in str(text).strip().split("\n") if line.strip()]
        if not lines:
            self.add_body_text("Chưa ghi nhận thông tin.")
            return

        self.set_font("Roboto", "", 9.5)
        self.set_text_color(0, 0, 0)
        cau_dan = lines[0]
        self.multi_cell(0, 5, cau_dan)
        self.ln(1)

        orig_l_margin = self.l_margin
        indent = 8.0
        self.set_left_margin(orig_l_margin + indent)

        for line in lines[1:]:
            bullet_line = line
            if not (bullet_line.startswith("-") or bullet_line.startswith("*")):
                bullet_line = f"- {bullet_line}"
            self.multi_cell(0, 5, bullet_line)
            self.ln(1)

        self.set_left_margin(orig_l_margin)
        self.ln(2)

    def render_table_cls(self, cls_rows):
        col_w = 95
        line_h = 5.0
        
        self.set_font("Roboto-Bold", "", 9.5)
        self.set_fill_color(230, 235, 245)
        start_x = self.get_x()
        start_y = self.get_y()
        
        if start_y > 260:
            self.add_page()
            start_y = self.get_y()
            
        self.cell(col_w, 7, "KẾT QUẢ CẬN LÂM SÀNG", border=1, align="C", fill=True)
        self.cell(col_w, 7, "PHIÊN GIẢI / BIỆN GIẢI", border=1, align="C", fill=True, new_x="LMARGIN", new_y="NEXT")
        
        temp_files = []
        try:
            self.set_font("Roboto", "", 9)
            for kq, pg, img in cls_rows:
                txt_kq = format_bullet_points(kq) if kq else ""
                txt_pg = format_bullet_points(pg) if pg else "-"
                
                temp_img_path = None
                img_h = 0
                if img:
                    suffix = os.path.splitext(img.name)[1]
                    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as t_img:
                        t_img.write(img.getbuffer())
                        temp_img_path = t_img.name
                        temp_files.append(temp_img_path)
                    img_h = 45
                    
                nb_lines_left = len(self.multi_cell(col_w - 4, line_h, txt_kq, dry_run=True, output="LINES"))
                nb_lines_right = len(self.multi_cell(col_w - 4, line_h, txt_pg, dry_run=True, output="LINES"))
                
                h_left = nb_lines_left * line_h + 4 + (img_h + 4 if img else 0)
                h_right = nb_lines_right * line_h + 4
                row_h = max(h_left, h_right, 8)
                
                if self.get_y() + row_h > 275:
                    self.add_page()
                    self.set_font("Roboto-Bold", "", 9.5)
                    self.set_fill_color(230, 235, 245)
                    self.cell(col_w, 7, "KẾT QUẢ CẬN LÂM SÀNG", border=1, align="C", fill=True)
                    self.cell(col_w, 7, "PHIÊN GIẢI / BIỆN GIẢI", border=1, align="C", fill=True, new_x="LMARGIN", new_y="NEXT")
                    self.set_font("Roboto", "", 9)

                curr_x = self.get_x()
                curr_y = self.get_y()
                
                self.rect(curr_x, curr_y, col_w, row_h)
                self.rect(curr_x + col_w, curr_y, col_w, row_h)
                
                self.set_xy(curr_x + 2, curr_y + 2)
                if txt_kq:
                    self.multi_cell(col_w - 4, line_h, txt_kq)
                
                if temp_img_path:
                    y_img = self.get_y() + 1
                    img_w = min(col_w - 8, 70)
                    x_img = curr_x + (col_w - img_w) / 2
                    try:
                        self.image(temp_img_path, x=x_img, y=y_img, w=img_w, h=img_h)
                    except:
                        pass
                
                self.set_xy(curr_x + col_w + 2, curr_y + 2)
                self.multi_cell(col_w - 4, line_h, txt_pg)
                self.set_xy(curr_x, curr_y + row_h)
                
            self.ln(3)
        finally:
            for p in temp_files:
                if os.path.exists(p):
                    try:
                        os.remove(p)
                    except:
                        pass

def export_pdf(data):
    pdf = BenhAnPDF()
    pdf.alias_nb_pages()
    pdf.add_page()
    
    # I. HÀNH CHÍNH
    pdf.add_section_header("I. PHẦN HÀNH CHÍNH")
    hc_text = (
        f"- Họ và tên: {str(data['ho_ten']).upper()}    |    Tuổi: {data['tuoi']}    |    Giới tính: {data['gioi_tinh']}\n"
        f"- Dân tộc: {data['dan_tok']}    |    Nghề nghiệp: {data['nghe_nghiep']}\n"
        f"- Khoa phòng: {data['khoa_phong']}\n"
        f"- Địa chỉ: {data['dia_chi']}\n"
        f"- Ngày giờ vào viện: {data['ngay_vao_vien']}\n"
        f"- Bác sĩ hoặc Sinh viên thực hiện: {data['sinh_vien']}"
    )
    pdf.add_body_text(hc_text)

    # II. LÝ DO VÀO VIỆN
    pdf.add_section_header("II. LÝ DO VÀO VIỆN")
    pdf.add_body_text(data['ly_do_vao_vien'])

    # III. BỆNH SỬ
    pdf.add_section_header("III. BỆNH SỬ")
    pdf.add_body_text(data['benh_su'])

    # IV. TIỀN SỬ
    pdf.add_section_header("IV. TIỀN SỬ")
    pdf.add_subsection_header("1. Tiền sử nội khoa:")
    pdf.add_body_text(format_bullet_points(data['ts_noi_khoa']))
    pdf.add_subsection_header("2. Tiền sử ngoại khoa và dị ứng:")
    pdf.add_body_text(format_bullet_points(data['ts_ngoai_khoa']))
    pdf.add_subsection_header("3. Lối sống và thói quen:")
    pdf.add_body_text(format_bullet_points(data['ts_loi_song']))
    pdf.add_subsection_header("4. Tiền sử gia đình:")
    pdf.add_body_text(format_bullet_points(data['ts_gia_dinh']))

    # V. THĂM KHÁM LÂM SÀNG
    pdf.add_section_header("V. THĂM KHÁM LÂM SÀNG")
    pdf.add_subsection_header("1. Thăm khám lúc vào viện:")
    pdf.add_body_text(format_bullet_points(data['kham_vao_vien']))
    
    pdf.add_subsection_header("2. Thăm khám hiện tại:")
    pdf.add_body_text("a. Toàn thân:")
    pdf.add_body_text(format_bullet_points(data['kham_toan_than']))
    
    pdf.add_body_text("b. Các cơ quan:")
    pdf.add_body_text(f"Tuần hoàn:\n{format_bullet_points(data['kham_tuan_hoan'])}")
    pdf.add_body_text(f"Hô hấp:\n{format_bullet_points(data['kham_ho_hap'])}")
    pdf.add_body_text(f"Tiêu hóa:\n{format_bullet_points(data['kham_tieu_hoa'])}")
    pdf.add_body_text(f"Thần kinh:\n{format_bullet_points(data['kham_than_kinh'])}")
    pdf.add_body_text(f"Thận - Tiết niệu:\n{format_bullet_points(data['kham_tiet_nieu'])}")
    pdf.add_body_text(f"Cơ xương khớp:\n{format_bullet_points(data['kham_co_xuong_khop'])}")
    pdf.add_body_text(f"Các cơ quan khác:\n{format_bullet_points(data['kham_co_quan_khac'])}")

    # VI. TÓM TẮT BỆNH ÁN
    pdf.add_section_header("VI. TÓM TẮT BỆNH ÁN")
    pdf.render_tom_tat_pdf(data['tom_tat'])

    # VII. CHẨN ĐOÁN SƠ BỘ
    pdf.add_section_header("VII. CHẨN ĐOÁN SƠ BỘ")
    pdf.add_body_text(data['chan_doan_so_bo'])

    # VIII. CHẨN ĐOÁN PHÂN BIỆT
    pdf.add_section_header("VIII. CHẨN ĐOÁN PHÂN BIỆT")
    pdf.add_body_text(data['chan_doan_phan_biet'])

    # IX. BIỆN LUẬN CHẨN ĐOÁN SƠ BỘ
    pdf.add_section_header("IX. BIỆN LUẬN CHẨN ĐOÁN SƠ BỘ")
    pdf.add_body_text(data['bien_luan'])

    # X. ĐỀ XUẤT CẬN LÂM SÀNG
    pdf.add_section_header("X. ĐỀ XUẤT CẬN LÂM SÀNG")
    pdf.add_subsection_header("1. Phục vụ chẩn đoán xác định:")
    pdf.add_body_text(format_bullet_points(data['cls_dx_xac_dinh']))
    pdf.add_subsection_header("2. Phục vụ điều trị:")
    pdf.add_body_text(format_bullet_points(data['cls_dx_dieu_tri']))
    pdf.add_subsection_header("3. Cận lâm sàng khác:")
    pdf.add_body_text(format_bullet_points(data['cls_dx_khac']))

    # XI. CẬN LÂM SÀNG ĐÃ CÓ
    pdf.add_section_header("XI. CẬN LÂM SÀNG ĐÃ CÓ")
    cls_rows = []
    so_hang = data.get("so_hang_cls", 3)
    for i in range(so_hang):
        kq = data.get(f"cls_kq_{i}", "").strip()
        pg = data.get(f"cls_pg_{i}", "").strip()
        img = data.get(f"cls_img_{i}", None)
        if kq or pg or img:
            cls_rows.append((kq, pg, img))

    if not cls_rows:
        pdf.add_body_text("Chưa ghi nhận kết quả cận lâm sàng.")
    else:
        pdf.render_table_cls(cls_rows)

    # XII. CHẨN ĐOÁN XÁC ĐỊNH
    pdf.add_section_header("XII. CHẨN ĐOÁN XÁC ĐỊNH")
    pdf.add_highlight_text(format_bullet_points(data['chan_doan_xac_dinh']))

    # XIII. BIỆN LUẬN CHẨN ĐOÁN XÁC ĐỊNH
    pdf.add_section_header("XIII. BIỆN LUẬN CHẨN ĐOÁN XÁC ĐỊNH")
    pdf.add_body_text(format_bullet_points(data['bien_luan_xac_dinh']))

    # XIV. ĐIỀU TRỊ
    pdf.add_section_header("XIV. ĐIỀU TRỊ")
    pdf.add_subsection_header("1. Mục tiêu điều trị:")
    pdf.add_body_text(format_bullet_points(data['dt_muc_tieu']))
    pdf.add_subsection_header("2. Điều trị cụ thể:")
    pdf.add_body_text(format_bullet_points(data['dt_cu_the']))
    pdf.add_subsection_header("3. Theo dõi sau điều trị:")
    pdf.add_body_text(format_bullet_points(data['dt_theo_doi']))

    # XV. TIÊN LƯỢNG (CHỈ HIỆN KHI CÓ NHẬP NỘI DUNG)
    noi_dung_tien_luong = str(data.get("tien_luong", "")).strip()
    if noi_dung_tien_luong:
        pdf.add_section_header("XV. TIÊN LƯỢNG")
        pdf.add_body_text(format_bullet_points(noi_dung_tien_luong))

    # XVI. TƯ VẤN (CHỈ HIỆN KHI CÓ NHẬP NỘI DUNG)
    noi_dung_tu_van = str(data.get("tu_van", "")).strip()
    if noi_dung_tu_van:
        # Nếu có mục Tiên lượng thì đề mục là XVI, nếu không có thì đẩy lên thành XV
        ten_de_muc_tu_van = "XVI. TƯ VẤN" if noi_dung_tien_luong else "XV. TƯ VẤN"
        pdf.add_section_header(ten_de_muc_tu_van)
        pdf.add_body_text(format_bullet_points(noi_dung_tu_van))

    return bytes(pdf.output())

# --- SIDEBAR: XỬ LÝ LƯU & NẠP BẢN NHÁP ---
with st.sidebar:
    st.markdown("<div class='sidebar-header-amboss'>Quản lý bản nháp</div>", unsafe_allow_html=True)
    st.caption("Lưu hoặc khôi phục dữ liệu bệnh án từ tập tin JSON.")
    
    current_data = {k: st.session_state.get(k, "") for k in default_fields}
    current_data["so_hang_cls"] = st.session_state.get("so_hang_cls", 3)
    for i in range(current_data["so_hang_cls"]):
        current_data[f"cls_kq_{i}"] = st.session_state.get(f"cls_kq_{i}", "")
        current_data[f"cls_pg_{i}"] = st.session_state.get(f"cls_pg_{i}", "")
        
    json_string = json.dumps(current_data, ensure_ascii=False, indent=2)
    ten_benh_nhan = str(st.session_state.get("ho_ten", "chua_dat_ten")).strip().replace(" ", "_")
    if not ten_benh_nhan:
        ten_benh_nhan = "chua_dat_ten"
        
    st.download_button(
        label="📥 Lưu bản nháp về máy (.json)",
        data=json_string,
        file_name=f"Ban_nhap_{ten_benh_nhan}_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
        mime="application/json",
        use_container_width=True
    )
    
    st.markdown("---")
    st.markdown("**Khôi phục dữ liệu từ bản nháp:**")
    file_nhap = st.file_uploader("Chọn tập tin .json đã lưu:", type=["json"], key="uploader_nhap_json")
    
    if file_nhap is not None:
        if st.button("🔄 Nhấn vào đây để nạp dữ liệu", type="primary", use_container_width=True):
            try:
                loaded_data = json.load(file_nhap)
                if "so_hang_cls" in loaded_data:
                    st.session_state["so_hang_cls"] = int(loaded_data["so_hang_cls"])
                for k in default_fields:
                    if k in loaded_data:
                        st.session_state[k] = loaded_data[k]
                for i in range(st.session_state["so_hang_cls"]):
                    if f"cls_kq_{i}" in loaded_data:
                        st.session_state[f"cls_kq_{i}"] = loaded_data[f"cls_kq_{i}"]
                    if f"cls_pg_{i}" in loaded_data:
                        st.session_state[f"cls_pg_{i}"] = loaded_data[f"cls_pg_{i}"]
                        
                st.success("Đã nạp bản nháp thành công!")
                st.rerun()
            except Exception as e:
                st.error(f"Không thể đọc file: {e}")

# --- GIAO DIỆN CHÍNH ---
st.title("Công cụ Nhập và Xuất Bệnh Án Lâm Sàng")
st.caption("Cấu trúc bệnh án phân tích chuyên sâu phục vụ học tập, giao ban và thực hành lâm sàng.")

tab1, tab2 = st.tabs(["Nhập liệu hồ sơ", "Xem trước và Xuất tập tin"])

with tab1:
    # 1. HÀNH CHÍNH (EXPANDER)
    with st.expander("I. PHẦN HÀNH CHÍNH", expanded=True):
        c_hc1, c_hc2, c_hc3 = st.columns(3)
        with c_hc1:
            st.text_input("Họ và tên người bệnh", key="ho_ten", placeholder="Nguyễn Văn A")
            st.text_input("Dân tộc", key="dan_tok", placeholder="Kinh, Tày, Nùng...")
        with c_hc2:
            st.number_input("Tuổi", min_value=0, max_value=120, key="tuoi")
            st.text_input("Nghề nghiệp", key="nghe_nghiep", placeholder="Kỹ sư, Hưu trí, Nông dân...")
        with c_hc3:
            st.selectbox("Giới tính", ["Nam", "Nữ", "Khác"], key="gioi_tinh")
            st.text_input("Khoa / Phòng điều trị", key="khoa_phong", placeholder="Khoa Cấp cứu, Khoa Nội Tim mạch...")
        
        c_hc4, c_hc5, c_hc6 = st.columns(3)
        with c_hc4:
            st.text_input("Địa chỉ", key="dia_chi", placeholder="Quận Đống Đa, TP. Hà Nội")
        with c_hc5:
            st.text_input("Bác sĩ hoặc Sinh viên phụ trách", key="sinh_vien", placeholder="Bác sĩ nội trú, Sinh viên Y...")
        with c_hc6:
            st.text_input("Ngày giờ vào viện", key="ngay_vao_vien")

    # 2. LÝ DO VÀO VIỆN & BỆNH SỬ (EXPANDER)
    with st.expander("II VÀ III. LÝ DO VÀO VIỆN VÀ BỆNH SỬ", expanded=True):
        st.text_area("Lý do vào viện:", key="ly_do_vao_vien", placeholder="Ví dụ: Đau thắt ngực trái giờ thứ 2 lan vai trái...", height=65)
        st.text_area("Bệnh sử:", key="benh_su", placeholder="Mô tả hoàn cảnh khởi phát, triệu chứng cơ năng điển hình...", height=130)

    # 3. TIỀN SỬ (EXPANDER)
    with st.expander("IV. TIỀN SỬ", expanded=True):
        st.caption("Lưu ý: Mỗi lần nhấn xuống dòng, hệ thống sẽ tự động chuyển thành gạch đầu dòng trong tập tin xuất ra.")
        c_ts1, c_ts2 = st.columns(2)
        with c_ts1:
            st.markdown("<div class='sub-section-header'>1. Tiền sử nội khoa</div>", unsafe_allow_html=True)
            st.text_area("Nội dung tiền sử nội khoa:", key="ts_noi_khoa", height=90, label_visibility="collapsed")
            
            st.markdown("<div class='sub-section-header'>2. Tiền sử ngoại khoa và dị ứng</div>", unsafe_allow_html=True)
            st.text_area("Nội dung tiền sử ngoại khoa và dị ứng:", key="ts_ngoai_khoa", height=90, label_visibility="collapsed")
        with c_ts2:
            st.markdown("<div class='sub-section-header'>3. Lối sống và thói quen</div>", unsafe_allow_html=True)
            st.text_area("Nội dung lối sống và thói quen:", key="ts_loi_song", height=90, label_visibility="collapsed")
            
            st.markdown("<div class='sub-section-header'>4. Tiền sử gia đình</div>", unsafe_allow_html=True)
            st.text_area("Nội dung tiền sử gia đình:", key="ts_gia_dinh", height=90, label_visibility="collapsed")

    # 4. THĂM KHÁM LÂM SÀNG (EXPANDER)
    with st.expander("V. THĂM KHÁM LÂM SÀNG", expanded=True):
        st.caption("Lưu ý: Tất cả các ô thăm khám khi nhấn xuống dòng sẽ tự động tạo gạch đầu dòng trong tập tin xuất ra.")
        st.markdown("<div class='sub-section-header'>1. Thăm khám lúc vào viện</div>", unsafe_allow_html=True)
        st.text_area("Nội dung khám lúc vào viện:", key="kham_vao_vien", height=80, label_visibility="collapsed")
        
        st.markdown("<div class='sub-section-header'>2. Thăm khám hiện tại - Toàn thân</div>", unsafe_allow_html=True)
        st.text_area("Nội dung khám toàn thân:", key="kham_toan_than", height=90, label_visibility="collapsed")
        
        st.markdown("<div class='sub-section-header'>3. Thăm khám hiện tại - Các cơ quan</div>", unsafe_allow_html=True)
        c_cq1, c_cq2 = st.columns(2)
        with c_cq1:
            st.text_area("Tuần hoàn:", key="kham_tuan_hoan", height=85)
            st.text_area("Hô hấp:", key="kham_ho_hap", height=85)
            st.text_area("Tiêu hóa:", key="kham_tieu_hoa", height=85)
            st.text_area("Thần kinh:", key="kham_than_kinh", height=85)
        with c_cq2:
            st.text_area("Thận - Tiết niệu:", key="kham_tiet_nieu", height=85)
            st.text_area("Cơ xương khớp:", key="kham_co_xuong_khop", height=85)
            st.text_area("Các cơ quan khác:", key="kham_co_quan_khac", height=85)

    # 5. TỔNG HỢP & BIỆN LUẬN CHẨN ĐOÁN SƠ BỘ (EXPANDER)
    with st.expander("VI ĐẾN IX. TÓM TẮT VÀ BIỆN LUẬN CHẨN ĐOÁN SƠ BỘ", expanded=True):
        st.caption("Lưu ý: Dòng đầu tiên là câu dẫn. Từ lần xuống dòng tiếp theo sẽ tự động thụt lề và thêm gạch đầu dòng.")
        st.text_area("VI. Tóm tắt bệnh án:", key="tom_tat", height=110)
        
        c_cd1, c_cd2 = st.columns(2)
        with c_cd1:
            st.text_area("VII. Chẩn đoán sơ bộ:", key="chan_doan_so_bo", height=85)
        with c_cd2:
            st.text_area("VIII. Chẩn đoán phân biệt:", key="chan_doan_phan_biet", height=85)
            
        st.text_area("IX. Biện luận chẩn đoán sơ bộ:", key="bien_luan", height=110)

    # 6. CẬN LÂM SÀNG (EXPANDER)
    with st.expander("X VÀ XI. CẬN LÂM SÀNG", expanded=True):
        st.markdown("<div class='sub-section-header'>X. Đề xuất cận lâm sàng</div>", unsafe_allow_html=True)
        c_cls1, c_cls2, c_cls3 = st.columns(3)
        with c_cls1:
            st.text_area("1. Phục vụ chẩn đoán xác định:", key="cls_dx_xac_dinh", height=100)
        with c_cls2:
            st.text_area("2. Phục vụ điều trị:", key="cls_dx_dieu_tri", height=100)
        with c_cls3:
            st.text_area("3. Cận lâm sàng khác:", key="cls_dx_khac", height=100)
            
        st.markdown(f"<div class='sub-section-header'>XI. Cận lâm sàng đã có (Hiện có {st.session_state['so_hang_cls']} hàng)</div>", unsafe_allow_html=True)
        st.caption("Mỗi hàng: Cột trái nhập kết quả hoặc kèm ảnh; Cột phải nhập biện giải tương ứng. Khi xuất PDF sẽ tự động xếp thành bảng 2 cột đối chiếu.")
        
        uploaded_imgs = {}
        for i in range(st.session_state["so_hang_cls"]):
            st.markdown(f"**Hàng {i + 1}:**")
            col_left, col_right = st.columns([1, 1])
            with col_left:
                st.text_area(
                    f"Kết quả cận lâm sàng {i + 1}:",
                    key=f"cls_kq_{i}",
                    placeholder=f"Ví dụ: Điện tâm đồ 12 chuyển đạo, Sinh hóa máu, X-quang...",
                    height=75
                )
                img = st.file_uploader(
                    f"Đính kèm ảnh cho hàng {i + 1} (tùy chọn):",
                    type=["png", "jpg", "jpeg"],
                    key=f"uploader_cls_img_{i}"
                )
                if img:
                    uploaded_imgs[f"cls_img_{i}"] = img
                    st.image(img, width=180, caption=f"Ảnh hàng {i + 1}")
            with col_right:
                st.text_area(
                    f"Biện giải cận lâm sàng {i + 1}:",
                    key=f"cls_pg_{i}",
                    placeholder=f"Ý nghĩa bệnh lý, đánh giá tăng/giảm hoặc tổn thương tương ứng...",
                    height=130
                )
            st.divider()

        col_btn_them, col_btn_bot, _ = st.columns([2, 2, 6])
        with col_btn_them:
            if st.button("➕ Thêm kết quả cận lâm sàng"):
                st.session_state["so_hang_cls"] += 1
                st.rerun()
        with col_btn_bot:
            if st.session_state["so_hang_cls"] > 1:
                if st.button("➖ Bớt hàng cuối"):
                    st.session_state["so_hang_cls"] -= 1
                    st.rerun()

    # 7. CHẨN ĐOÁN XÁC ĐỊNH (EXPANDER)
    with st.expander("XII VÀ XIII. CHẨN ĐOÁN XÁC ĐỊNH VÀ BIỆN LUẬN", expanded=True):
        st.caption("Lưu ý: Chẩn đoán xác định sẽ được in chữ màu ĐỎ ĐẬM trong PDF.")
        st.text_area("XII. Chẩn đoán xác định:", key="chan_doan_xac_dinh", height=90)
        st.text_area("XIII. Biện luận chẩn đoán xác định:", key="bien_luan_xac_dinh", height=110)

    # 8. ĐIỀU TRỊ (EXPANDER)
    with st.expander("XIV. ĐIỀU TRỊ", expanded=True):
        c_dt1, c_dt2, c_dt3 = st.columns(3)
        with c_dt1:
            st.text_area("1. Mục tiêu điều trị:", key="dt_muc_tieu", height=110)
        with c_dt2:
            st.text_area("2. Điều trị cụ thể:", key="dt_cu_the", height=110)
        with c_dt3:
            st.text_area("3. Theo dõi sau điều trị:", key="dt_theo_doi", height=110)

    # 9. TIÊN LƯỢNG & TƯ VẤN (EXPANDER)
    with st.expander("XV VÀ XVI. TIÊN LƯỢNG VÀ TƯ VẤN", expanded=True):
        c_pl, c_tv = st.columns(2)
        with c_pl:
            st.text_area("XV. Tiên lượng:", key="tien_luong", height=100, placeholder="Nếu để trống, mục này sẽ không xuất hiện trong file PDF...")
        with c_tv:
            st.text_area("XVI. Tư vấn:", key="tu_van", height=100, placeholder="Nếu để trống, mục này sẽ không xuất hiện trong file PDF...")

# Gom dữ liệu từ session_state sang dict để xuất file PDF
data_benh_an = {k: st.session_state.get(k, "") for k in default_fields}
data_benh_an["so_hang_cls"] = st.session_state.get("so_hang_cls", 3)
for i in range(data_benh_an["so_hang_cls"]):
    data_benh_an[f"cls_kq_{i}"] = st.session_state.get(f"cls_kq_{i}", "")
    data_benh_an[f"cls_pg_{i}"] = st.session_state.get(f"cls_pg_{i}", "")
data_benh_an.update(uploaded_imgs)

with tab2:
    st.markdown("<div class='sidebar-header-amboss'>XEM TRƯỚC THÔNG TIN TỔNG QUAN</div>", unsafe_allow_html=True)
    ho_ten_val = str(st.session_state.get("ho_ten", "")).strip()
    if ho_ten_val:
        st.info(f"Bệnh nhân: {ho_ten_val.upper()} | {st.session_state.get('tuoi')} tuổi | Giới tính: {st.session_state.get('gioi_tinh')} | Dân tộc: {st.session_state.get('dan_tok')}")
        st.write(f"Khoa phòng: {st.session_state.get('khoa_phong', 'Chưa điền')} | Lý do vào viện: {st.session_state.get('ly_do_vao_vien')}")
        if st.session_state.get("chan_doan_so_bo"):
            st.write(f"Chẩn đoán sơ bộ: {st.session_state.get('chan_doan_so_bo')}")
        if st.session_state.get("chan_doan_xac_dinh"):
            st.markdown(f"<div class='highlight-dx'>Chẩn đoán xác định: {st.session_state.get('chan_doan_xac_dinh')}</div>", unsafe_allow_html=True)
        
        so_hang = st.session_state.get("so_hang_cls", 3)
        dem_cls = sum(1 for i in range(so_hang) if str(st.session_state.get(f"cls_kq_{i}", "")).strip() or uploaded_imgs.get(f"cls_img_{i}"))
        st.write(f"Số lượng cận lâm sàng đã nhập vào bảng: {dem_cls}/{so_hang} hàng")
    else:
        st.warning("Vui lòng điền thông tin bên tab Nhập liệu hồ sơ.")

    st.markdown("---")
    if st.button("Tạo và tải tập tin PDF bệnh án", type="primary", use_container_width=True):
        if not ho_ten_val:
            st.error("Vui lòng điền tối thiểu Họ và tên người bệnh trước khi xuất tập tin!")
        elif not os.path.exists("Roboto-Regular.ttf") or not os.path.exists("Roboto-Bold.ttf"):
            st.error("Chưa tìm thấy tập tin font 'Roboto-Regular.ttf' và 'Roboto-Bold.ttf' trong cùng thư mục với app.py. Hãy kiểm tra lại!")
        else:
            with st.spinner("Đang kết xuất văn bản PDF..."):
                pdf_bytes = export_pdf(data_benh_an)
                ten_file = f"Benh_an_{ho_ten_val.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.pdf"
                
                st.success("Tạo văn bản bệnh án thành công!")
                st.download_button(
                    label="Nhấn vào đây để tải tập tin PDF về máy",
                    data=pdf_bytes,
                    file_name=ten_file,
                    mime="application/pdf",
                    use_container_width=True
                )