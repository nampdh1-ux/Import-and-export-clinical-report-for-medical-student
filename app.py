import google.generativeai as genai
import streamlit as st
from fpdf import FPDF
from datetime import datetime
import os
import tempfile
import json
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# --- CẤU HÌNH TRANG ĐẦU TIÊN ---
st.set_page_config(page_title="Công cụ làm Bệnh án Lâm sàng", layout="wide")

# --- CSS TÙY BIẾN DẢI ĐỀ MỤC THEO MÀU XANH AMBOSS KẾT HỢP VẠCH KÉP HMU ---
st.markdown("""
<style>
    div[data-testid="stExpander"] {
        border: 1px solid #d4eaf0;
        border-radius: 6px;
        margin-bottom: 12px;
        background-color: #ffffff;
    }
    div[data-testid="stExpander"] > details > summary {
        background-color: #ebf7f9 !important;
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
        start_y = self.get_y()
        
        if start_y > 260:
            self.add_page()
            
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

# --- HÀM XUẤT FILE PDF ---
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

    # XV. TIÊN LƯỢNG
    noi_dung_tien_luong = str(data.get("tien_luong", "")).strip()
    if noi_dung_tien_luong:
        pdf.add_section_header("XV. TIÊN LƯỢNG")
        pdf.add_body_text(format_bullet_points(noi_dung_tien_luong))

    # XVI. TƯ VẤN
    noi_dung_tu_van = str(data.get("tu_van", "")).strip()
    if noi_dung_tu_van:
        ten_de_muc_tu_van = "XVI. TƯ VẤN" if noi_dung_tien_luong else "XV. TƯ VẤN"
        pdf.add_section_header(ten_de_muc_tu_van)
        pdf.add_body_text(format_bullet_points(noi_dung_tu_van))

    return bytes(pdf.output())

# --- HÀM XUẤT FILE POWERPOINT (PPTX) ---
# --- HÀM XUẤT FILE POWERPOINT (PPTX) ĐÃ ĐƯỢC LÀM NỔI BẬT ĐỀ MỤC ---
def export_pptx(data):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    COLOR_PRIMARY = RGBColor(13, 71, 161)    # Xanh navy chuẩn HMU
    COLOR_ACCENT = RGBColor(194, 24, 91)     # Hồng sẫm HMU
    COLOR_TEXT = RGBColor(30, 41, 59)        # Màu chữ nội dung
    COLOR_RED = RGBColor(180, 0, 0)          # Màu đỏ đậm chẩn đoán xác định

    def add_slide_with_header(title_text):
        slide = prs.slides.add_slide(blank_layout)
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.9))
        tf = header_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title_text
        p.font.name = "Calibri"
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = COLOR_PRIMARY

        # Vạch trang trí kép nhận diện HMU
        line_rose = slide.shapes.add_shape(1, Inches(0.8), Inches(1.3), Inches(0.6), Inches(0.06))
        line_rose.fill.solid()
        line_rose.fill.fore_color.rgb = COLOR_ACCENT
        line_rose.line.fill.background()

        line_blue = slide.shapes.add_shape(1, Inches(1.45), Inches(1.3), Inches(11.083), Inches(0.06))
        line_blue.fill.solid()
        line_blue.fill.fore_color.rgb = COLOR_PRIMARY
        line_blue.line.fill.background()

        return slide

    def add_content_with_overflow(title_text, items_list, is_red=False):
        """
        Nhận vào danh sách tuple: (text, is_header)
        is_header=True -> Tự động IN ĐẬM, GẠCH CHÂN và đổi sang màu xanh thương hiệu.
        """
        if not items_list:
            return
            
        MAX_LINES_PER_SLIDE = 7
        total_chunks = [items_list[i:i + MAX_LINES_PER_SLIDE] for i in range(0, len(items_list), MAX_LINES_PER_SLIDE)]

        for idx, chunk in enumerate(total_chunks):
            current_title = title_text if idx == 0 else f"{title_text} (tiếp theo)"
            slide = add_slide_with_header(current_title)
            
            box = slide.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.733), Inches(5.2))
            tf = box.text_frame
            tf.word_wrap = True
            
            for line_idx, item in enumerate(chunk):
                text, is_sub_header = item
                p = tf.paragraphs[0] if line_idx == 0 else tf.add_paragraph()
                p.text = text
                p.font.name = "Calibri"
                
                if is_sub_header:
                    # NỔI BẬT: IN ĐẬM + GẠCH CHÂN
                    p.font.size = Pt(19)
                    p.font.bold = True
                    p.font.underline = True
                    p.font.color.rgb = COLOR_PRIMARY
                    p.space_after = Pt(6)
                else:
                    p.font.size = Pt(16.5)
                    p.font.bold = is_red
                    p.font.underline = False
                    p.font.color.rgb = COLOR_RED if is_red else COLOR_TEXT
                    p.space_after = Pt(10)

    # --- SLIDE 1: TRANG TIÊU ĐỀ BỆNH ÁN ---
    title_slide = prs.slides.add_slide(blank_layout)
    t_box = title_slide.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.333), Inches(3.5))
    tf_t = t_box.text_frame
    tf_t.word_wrap = True
    p1 = tf_t.paragraphs[0]
    p1.text = "BỆNH ÁN LÂM SÀNG"
    p1.font.size = Pt(38)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_PRIMARY
    p1.alignment = PP_ALIGN.CENTER
    p1.space_after = Pt(16)

    p2 = tf_t.add_paragraph()
    p2.text = f"Bệnh nhân: {str(data['ho_ten']).upper()} | {data['tuoi']} tuổi | Giới tính: {data['gioi_tinh']}"
    p2.font.size = Pt(20)
    p2.font.color.rgb = COLOR_TEXT
    p2.alignment = PP_ALIGN.CENTER
    p2.space_after = Pt(8)

    p3 = tf_t.add_paragraph()
    p3.text = f"Khoa phòng: {data['khoa_phong']} | Người thực hiện: {data['sinh_vien']}"
    p3.font.size = Pt(16)
    p3.font.color.rgb = RGBColor(100, 116, 139)
    p3.alignment = PP_ALIGN.CENTER

    # --- SLIDE: I. HÀNH CHÍNH ---
    hc_items = [
        (f"Họ và tên: {str(data['ho_ten']).upper()}", False),
        (f"Tuổi: {data['tuoi']}   |   Giới tính: {data['gioi_tinh']}   |   Dân tộc: {data['dan_tok']}", False),
        (f"Nghề nghiệp: {data['nghe_nghiep']}", False),
        (f"Khoa / Phòng: {data['khoa_phong']}", False),
        (f"Địa chỉ: {data['dia_chi']}", False),
        (f"Ngày giờ vào viện: {data['ngay_vao_vien']}", False),
        (f"Bác sĩ hoặc Sinh viên thực hiện: {data['sinh_vien']}", False)
    ]
    add_content_with_overflow("I. PHẦN HÀNH CHÍNH", hc_items)

    # --- SLIDE: II & III. LÝ DO VÀO VIỆN VÀ BỆNH SỬ ---
    bs_items = [("1. Lý do vào viện:", True)]
    bs_items.append((f"- {data['ly_do_vao_vien']}", False))
    
    bs_text_lines = [l.strip() for l in str(data['benh_su']).split("\n") if l.strip()]
    if bs_text_lines:
        bs_items.append(("2. Bệnh sử:", True))
        for l in bs_text_lines:
            bs_items.append((f"- {l}" if not l.startswith("-") else l, False))
    add_content_with_overflow("II VÀ III. LÝ DO VÀO VIỆN VÀ BỆNH SỬ", bs_items)

    # --- SLIDE: IV. TIỀN SỬ ---
    ts_items = []
    for label, key in [("1. Tiền sử nội khoa:", "ts_noi_khoa"),
                       ("2. Tiền sử ngoại khoa & dị ứng:", "ts_ngoai_khoa"),
                       ("3. Lối sống & thói quen:", "ts_loi_song"),
                       ("4. Tiền sử gia đình:", "ts_gia_dinh")]:
        lines = [l.strip() for l in str(data.get(key, "")).split("\n") if l.strip()]
        if lines:
            ts_items.append((label, True))
            for l in lines:
                ts_items.append((f"- {l}" if not l.startswith("-") else l, False))
    if ts_items:
        add_content_with_overflow("IV. TIỀN SỬ", ts_items)

    # --- SLIDE: V. THĂM KHÁM LÂM SÀNG ---
    tk_items = []
    if str(data.get("kham_vao_vien", "")).strip():
        tk_items.append(("1. Khám lúc vào viện:", True))
        for l in str(data['kham_vao_vien']).split("\n"):
            if l.strip():
                tk_items.append((f"- {l.strip()}", False))
                
    if str(data.get("kham_toan_than", "")).strip():
        tk_items.append(("2. Thăm khám hiện tại - Toàn thân:", True))
        for l in str(data['kham_toan_than']).split("\n"):
            if l.strip():
                tk_items.append((f"- {l.strip()}", False))
    
    cq_list = [("Tuần hoàn", "kham_tuan_hoan"), ("Hô hấp", "kham_ho_hap"), ("Tiêu hóa", "kham_tieu_hoa"),
               ("Thần kinh", "kham_than_kinh"), ("Thận - Tiết niệu", "kham_tiet_nieu"), ("Cơ xương khớp", "kham_co_xuong_khop")]
    cq_has_data = any(str(data.get(k, "")).strip() for _, k in cq_list)
    if cq_has_data:
        tk_items.append(("3. Thăm khám hiện tại - Các cơ quan:", True))
        for name, key in cq_list:
            val = str(data.get(key, "")).strip()
            if val:
                tk_items.append((f"- {name}: {val}", False))
    if tk_items:
        add_content_with_overflow("V. THĂM KHÁM LÂM SÀNG", tk_items)

    # --- SLIDE: VI. TÓM TẮT BỆNH ÁN ---
    tt_items = []
    lines_tt = [l.strip() for l in str(data.get("tom_tat", "")).split("\n") if l.strip()]
    if lines_tt:
        tt_items.append(("Tóm tắt diễn biến ca bệnh:", True))
        for l in lines_tt:
            tt_items.append((f"- {l}" if not l.startswith("-") else l, False))
        add_content_with_overflow("VI. TÓM TẮT BỆNH ÁN", tt_items)

    # --- SLIDE: VII & VIII. CHẨN ĐOÁN SƠ BỘ VÀ PHÂN BIỆT ---
    cd_items = []
    if str(data.get("chan_doan_so_bo", "")).strip():
        cd_items.append(("Chẩn đoán sơ bộ:", True))
        for l in str(data['chan_doan_so_bo']).split("\n"):
            if l.strip():
                cd_items.append((f"- {l.strip()}", False))
                
    if str(data.get("chan_doan_phan_biet", "")).strip():
        cd_items.append(("Chẩn đoán phân biệt:", True))
        for l in str(data['chan_doan_phan_biet']).split("\n"):
            if l.strip():
                cd_items.append((f"- {l.strip()}", False))
    if cd_items:
        add_content_with_overflow("VII VÀ VIII. CHẨN ĐOÁN SƠ BỘ VÀ PHÂN BIỆT", cd_items)

    # --- SLIDE: IX. BIỆN LUẬN CHẨN ĐOÁN SƠ BỘ ---
    bl_items = []
    lines_bl = [l.strip() for l in str(data.get("bien_luan", "")).split("\n") if l.strip()]
    if lines_bl:
        bl_items.append(("Biện luận lâm sàng:", True))
        for l in lines_bl:
            bl_items.append((f"- {l}" if not l.startswith("-") else l, False))
        add_content_with_overflow("IX. BIỆN LUẬN CHẨN ĐOÁN SƠ BỘ", bl_items)

    # --- SLIDE: X. ĐỀ XUẤT CẬN LÂM SÀNG ---
    cls_items = []
    for label, key in [("1. Phục vụ chẩn đoán xác định:", "cls_dx_xac_dinh"),
                       ("2. Phục vụ điều trị:", "cls_dx_dieu_tri"),
                       ("3. Cận lâm sàng khác:", "cls_dx_khac")]:
        lines = [l.strip() for l in str(data.get(key, "")).split("\n") if l.strip()]
        if lines:
            cls_items.append((label, True))
            for l in lines:
                cls_items.append((f"- {l}" if not l.startswith("-") else l, False))
    if cls_items:
        add_content_with_overflow("X. ĐỀ XUẤT CẬN LÂM SÀNG", cls_items)

    # --- SLIDE: XI. CẬN LÂM SÀNG ĐÃ CÓ (BẢNG & HÌNH ẢNH) ---
    cls_rows = []
    so_hang = data.get("so_hang_cls", 3)
    for i in range(so_hang):
        kq = data.get(f"cls_kq_{i}", "").strip()
        pg = data.get(f"cls_pg_{i}", "").strip()
        img = data.get(f"cls_img_{i}", None)
        if kq or pg or img:
            cls_rows.append((kq, pg, img))

    if cls_rows:
        rows_text_only = [item for item in cls_rows if not item[2]]
        rows_with_img = [item for item in cls_rows if item[2]]

        # Bảng các cận lâm sàng dạng chữ
        if rows_text_only:
            MAX_ROWS_PER_SLIDE = 3
            table_chunks = [rows_text_only[i:i + MAX_ROWS_PER_SLIDE] for i in range(0, len(rows_text_only), MAX_ROWS_PER_SLIDE)]
            for c_idx, chunk in enumerate(table_chunks):
                t_title = "XI. CẬN LÂM SÀNG ĐÃ CÓ" if c_idx == 0 else "XI. CẬN LÂM SÀNG ĐÃ CÓ (tiếp theo)"
                slide = add_slide_with_header(t_title)
                
                rows_cnt = len(chunk) + 1
                table_shape = slide.shapes.add_table(rows_cnt, 2, Inches(0.8), Inches(1.6), Inches(11.733), Inches(1.0 + len(chunk) * 1.3))
                table = table_shape.table
                table.columns[0].width = Inches(5.866)
                table.columns[1].width = Inches(5.866)

                table.cell(0, 0).text = "KẾT QUẢ CẬN LÂM SÀNG"
                table.cell(0, 1).text = "PHIÊN GIẢI / BIỆN GIẢI"
                for col_i in range(2):
                    cell_p = table.cell(0, col_i).text_frame.paragraphs[0]
                    cell_p.font.bold = True
                    cell_p.font.size = Pt(14)
                    cell_p.font.color.rgb = COLOR_PRIMARY

                for r_i, (kq, pg, _) in enumerate(chunk):
                    table.cell(r_i + 1, 0).text = kq if kq else "-"
                    table.cell(r_i + 1, 1).text = pg if pg else "-"
                    for col_i in range(2):
                        cell_p = table.cell(r_i + 1, col_i).text_frame.paragraphs[0]
                        cell_p.font.size = Pt(13)
                        cell_p.font.color.rgb = COLOR_TEXT

        # Các cận lâm sàng có hình ảnh
        temp_img_files = []
        try:
            for kq, pg, img in rows_with_img:
                slide = add_slide_with_header("XI. CẬN LÂM SÀNG ĐÃ CÓ (HÌNH ẢNH)")
                suffix = os.path.splitext(img.name)[1]
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as t_img:
                    t_img.write(img.getbuffer())
                    temp_path = t_img.name
                    temp_img_files.append(temp_path)

                try:
                    slide.shapes.add_picture(temp_path, Inches(0.8), Inches(1.6), width=Inches(5.6))
                except Exception:
                    pass

                right_box = slide.shapes.add_textbox(Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.0))
                tf_r = right_box.text_frame
                tf_r.word_wrap = True

                p_kq_title = tf_r.paragraphs[0]
                p_kq_title.text = "Kết quả ghi nhận:"
                p_kq_title.font.bold = True
                p_kq_title.font.underline = True
                p_kq_title.font.size = Pt(16)
                p_kq_title.font.color.rgb = COLOR_PRIMARY
                p_kq_title.space_after = Pt(4)

                p_kq = tf_r.add_paragraph()
                p_kq.text = kq if kq else "Hình ảnh xét nghiệm đính kèm"
                p_kq.font.size = Pt(14.5)
                p_kq.font.color.rgb = COLOR_TEXT
                p_kq.space_after = Pt(16)

                p_pg_title = tf_r.add_paragraph()
                p_pg_title.text = "Biện giải / Phiên giải:"
                p_pg_title.font.bold = True
                p_pg_title.font.underline = True
                p_pg_title.font.size = Pt(16)
                p_pg_title.font.color.rgb = COLOR_PRIMARY
                p_pg_title.space_after = Pt(4)

                p_pg = tf_r.add_paragraph()
                p_pg.text = pg if pg else "-"
                p_pg.font.size = Pt(14.5)
                p_pg.font.color.rgb = COLOR_TEXT
        finally:
            for p in temp_img_files:
                if os.path.exists(p):
                    try:
                        os.remove(p)
                    except:
                        pass

    # --- SLIDE: XII. CHẨN ĐOÁN XÁC ĐỊNH (IN ĐỎ ĐẬM) ---
    cdxd_items = []
    lines_cdxd = [l.strip() for l in str(data.get("chan_doan_xac_dinh", "")).split("\n") if l.strip()]
    if lines_cdxd:
        cdxd_items.append(("Chẩn đoán xác định:", True))
        for l in lines_cdxd:
            cdxd_items.append((f"- {l}" if not l.startswith("-") else l, False))
        add_content_with_overflow("XII. CHẨN ĐOÁN XÁC ĐỊNH", cdxd_items, is_red=True)

    # --- SLIDE: XIII. BIỆN LUẬN CHẨN ĐOÁN XÁC ĐỊNH ---
    blxd_items = []
    lines_blxd = [l.strip() for l in str(data.get("bien_luan_xac_dinh", "")).split("\n") if l.strip()]
    if lines_blxd:
        blxd_items.append(("Biện luận chẩn đoán xác định:", True))
        for l in lines_blxd:
            blxd_items.append((f"- {l}" if not l.startswith("-") else l, False))
        add_content_with_overflow("XIII. BIỆN LUẬN CHẨN ĐOÁN XÁC ĐỊNH", blxd_items)

    # --- SLIDE: XIV. ĐIỀU TRỊ ---
    dt_items = []
    for label, key in [("1. Mục tiêu điều trị:", "dt_muc_tieu"),
                       ("2. Điều trị cụ thể:", "dt_cu_the"),
                       ("3. Theo dõi sau điều trị:", "dt_theo_doi")]:
        lines = [l.strip() for l in str(data.get(key, "")).split("\n") if l.strip()]
        if lines:
            dt_items.append((label, True))
            for l in lines:
                dt_items.append((f"- {l}" if not l.startswith("-") else l, False))
    if dt_items:
        add_content_with_overflow("XIV. ĐIỀU TRỊ", dt_items)

    # --- SLIDE: XV. TIÊN LƯỢNG (CHỈ TẠO KHI CÓ DỮ LIỆU) ---
    lines_tl = [l.strip() for l in str(data.get("tien_luong", "")).split("\n") if l.strip()]
    if lines_tl:
        tl_items = [("Đánh giá tiên lượng bệnh nhân:", True)]
        for l in lines_tl:
            tl_items.append((f"- {l}" if not l.startswith("-") else l, False))
        add_content_with_overflow("XV. TIÊN LƯỢNG", tl_items)

    # --- SLIDE: XVI. TƯ VẤN (CHỈ TẠO KHI CÓ DỮ LIỆU) ---
    lines_tv = [l.strip() for l in str(data.get("tu_van", "")).split("\n") if l.strip()]
    if lines_tv:
        t_label = "XVI. TƯ VẤN" if lines_tl else "XV. TƯ VẤN"
        tv_items = [("Hướng dẫn và tư vấn cho người bệnh:", True)]
        for l in lines_tv:
            tv_items.append((f"- {l}" if not l.startswith("-") else l, False))
        add_content_with_overflow(t_label, tv_items)

    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io.getvalue()
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    COLOR_PRIMARY = RGBColor(13, 71, 161)
    COLOR_ACCENT = RGBColor(194, 24, 91)
    COLOR_TEXT = RGBColor(30, 41, 59)
    COLOR_RED = RGBColor(180, 0, 0)

    def add_slide_with_header(title_text):
        slide = prs.slides.add_slide(blank_layout)
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.9))
        tf = header_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title_text
        p.font.name = "Calibri"
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = COLOR_PRIMARY

        line_rose = slide.shapes.add_shape(1, Inches(0.8), Inches(1.3), Inches(0.6), Inches(0.06))
        line_rose.fill.solid()
        line_rose.fill.fore_color.rgb = COLOR_ACCENT
        line_rose.line.fill.background()

        line_blue = slide.shapes.add_shape(1, Inches(1.45), Inches(1.3), Inches(11.083), Inches(0.06))
        line_blue.fill.solid()
        line_blue.fill.fore_color.rgb = COLOR_PRIMARY
        line_blue.line.fill.background()

        return slide

    def add_content_with_overflow(title_text, lines_list, is_red=False):
        if not lines_list:
            return
        MAX_LINES_PER_SLIDE = 7
        total_chunks = [lines_list[i:i + MAX_LINES_PER_SLIDE] for i in range(0, len(lines_list), MAX_LINES_PER_SLIDE)]

        for idx, chunk in enumerate(total_chunks):
            current_title = title_text if idx == 0 else f"{title_text} (tiếp theo)"
            slide = add_slide_with_header(current_title)
            
            box = slide.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.733), Inches(5.2))
            tf = box.text_frame
            tf.word_wrap = True
            
            for line_idx, line in enumerate(chunk):
                p = tf.paragraphs[0] if line_idx == 0 else tf.add_paragraph()
                p.text = line
                p.font.name = "Calibri"
                p.font.size = Pt(17)
                p.font.color.rgb = COLOR_RED if is_red else COLOR_TEXT
                p.font.bold = is_red
                p.space_after = Pt(12)

    # Slide 1: Bìa
    title_slide = prs.slides.add_slide(blank_layout)
    t_box = title_slide.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.333), Inches(3.5))
    tf_t = t_box.text_frame
    tf_t.word_wrap = True
    p1 = tf_t.paragraphs[0]
    p1.text = "BỆNH ÁN LÂM SÀNG"
    p1.font.size = Pt(38)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_PRIMARY
    p1.alignment = PP_ALIGN.CENTER
    p1.space_after = Pt(16)

    p2 = tf_t.add_paragraph()
    p2.text = f"Bệnh nhân: {str(data['ho_ten']).upper()} | {data['tuoi']} tuổi | Giới tính: {data['gioi_tinh']}"
    p2.font.size = Pt(20)
    p2.font.color.rgb = COLOR_TEXT
    p2.alignment = PP_ALIGN.CENTER
    p2.space_after = Pt(8)

    p3 = tf_t.add_paragraph()
    p3.text = f"Khoa phòng: {data['khoa_phong']} | Người thực hiện: {data['sinh_vien']}"
    p3.font.size = Pt(16)
    p3.font.color.rgb = RGBColor(100, 116, 139)
    p3.alignment = PP_ALIGN.CENTER

    # Slide: I. Hành chính
    hc_lines = [
        f"Họ và tên: {str(data['ho_ten']).upper()}",
        f"Tuổi: {data['tuoi']}   |   Giới tính: {data['gioi_tinh']}   |   Dân tộc: {data['dan_tok']}",
        f"Nghề nghiệp: {data['nghe_nghiep']}",
        f"Khoa / Phòng: {data['khoa_phong']}",
        f"Địa chỉ: {data['dia_chi']}",
        f"Ngày giờ vào viện: {data['ngay_vao_vien']}",
        f"Bác sĩ hoặc Sinh viên thực hiện: {data['sinh_vien']}"
    ]
    add_content_with_overflow("I. PHẦN HÀNH CHÍNH", hc_lines)

    # Slide: II & III. Lý do vào viện & Bệnh sử
    bs_lines = [f"Lý do vào viện:\n- {data['ly_do_vao_vien']}"]
    bs_text_lines = [l.strip() for l in str(data['benh_su']).split("\n") if l.strip()]
    if bs_text_lines:
        bs_lines.append("Bệnh sử:")
        bs_lines.extend([f"- {l}" if not l.startswith("-") else l for l in bs_text_lines])
    add_content_with_overflow("II VÀ III. LÝ DO VÀO VIỆN VÀ BỆNH SỬ", bs_lines)

    # Slide: IV. Tiền sử
    ts_lines = []
    for label, key in [("1. Tiền sử nội khoa", "ts_noi_khoa"),
                       ("2. Tiền sử ngoại khoa & dị ứng", "ts_ngoai_khoa"),
                       ("3. Lối sống & thói quen", "ts_loi_song"),
                       ("4. Tiền sử gia đình", "ts_gia_dinh")]:
        items = [l.strip() for l in str(data.get(key, "")).split("\n") if l.strip()]
        if items:
            ts_lines.append(f"[{label}]")
            ts_lines.extend([f"- {i}" if not i.startswith("-") else i for i in items])
    if ts_lines:
        add_content_with_overflow("IV. TIỀN SỬ", ts_lines)

    # Slide: V. Thăm khám
    tk_lines = []
    if str(data.get("kham_vao_vien", "")).strip():
        tk_lines.append("[1. Khám lúc vào viện]")
        tk_lines.extend([f"- {l.strip()}" for l in str(data['kham_vao_vien']).split("\n") if l.strip()])
    if str(data.get("kham_toan_than", "")).strip():
        tk_lines.append("[2. Toàn thân]")
        tk_lines.extend([f"- {l.strip()}" for l in str(data['kham_toan_than']).split("\n") if l.strip()])
    
    co_quan = [("Tuần hoàn", "kham_tuan_hoan"), ("Hô hấp", "kham_ho_hap"), ("Tiêu hóa", "kham_tieu_hoa"),
               ("Thần kinh", "kham_than_kinh"), ("Thận - Tiết niệu", "kham_tiet_nieu"), ("Cơ xương khớp", "kham_co_xuong_khop")]
    for cq_name, cq_key in co_quan:
        val = str(data.get(cq_key, "")).strip()
        if val:
            tk_lines.append(f"[{cq_name}]: {val}")
    add_content_with_overflow("V. THĂM KHÁM LÂM SÀNG", tk_lines)

    # Slide: VI. Tóm tắt bệnh án
    tt_lines = [l.strip() for l in str(data.get("tom_tat", "")).split("\n") if l.strip()]
    if tt_lines:
        add_content_with_overflow("VI. TÓM TẮT BỆNH ÁN", tt_lines)

    # Slide: VII & VIII. Chẩn đoán sơ bộ & phân biệt
    cd_lines = []
    if str(data.get("chan_doan_so_bo", "")).strip():
        cd_lines.append("[Chẩn đoán sơ bộ]:")
        cd_lines.extend([f"- {l.strip()}" for l in str(data['chan_doan_so_bo']).split("\n") if l.strip()])
    if str(data.get("chan_doan_phan_biet", "")).strip():
        cd_lines.append("[Chẩn đoán phân biệt]:")
        cd_lines.extend([f"- {l.strip()}" for l in str(data['chan_doan_phan_biet']).split("\n") if l.strip()])
    add_content_with_overflow("VII VÀ VIII. CHẨN ĐOÁN SƠ BỘ VÀ PHÂN BIỆT", cd_lines)

    # Slide: IX. Biện luận sơ bộ
    bl_sb_lines = [l.strip() for l in str(data.get("bien_luan", "")).split("\n") if l.strip()]
    if bl_sb_lines:
        add_content_with_overflow("IX. BIỆN LUẬN CHẨN ĐOÁN SƠ BỘ", bl_sb_lines)

    # Slide: X. Đề xuất cận lâm sàng
    cls_dx_lines = []
    for label, key in [("1. Phục vụ chẩn đoán xác định", "cls_dx_xac_dinh"),
                       ("2. Phục vụ điều trị", "cls_dx_dieu_tri"),
                       ("3. Cận lâm sàng khác", "cls_dx_khac")]:
        items = [l.strip() for l in str(data.get(key, "")).split("\n") if l.strip()]
        if items:
            cls_dx_lines.append(f"[{label}]")
            cls_dx_lines.extend([f"- {i}" if not i.startswith("-") else i for i in items])
    if cls_dx_lines:
        add_content_with_overflow("X. ĐỀ XUẤT CẬN LÂM SÀNG", cls_dx_lines)

    # Slide: XI. Cận lâm sàng đã có (Bảng)
    # --- SLIDE: XI. CẬN LÂM SÀNG ĐÃ CÓ (BẢNG & ẢNH ĐÍNH KÈM) ---
    cls_rows = []
    so_hang = data.get("so_hang_cls", 3)
    for i in range(so_hang):
        kq = data.get(f"cls_kq_{i}", "").strip()
        pg = data.get(f"cls_pg_{i}", "").strip()
        img = data.get(f"cls_img_{i}", None)
        if kq or pg or img:
            cls_rows.append((kq, pg, img))

    if cls_rows:
        rows_text_only = []
        rows_with_img = []
        for item in cls_rows:
            if item[2]:  # Có ảnh đính kèm
                rows_with_img.append(item)
            else:
                rows_text_only.append(item)

        # 1. Hiển thị các cận lâm sàng dạng chữ (Dạng bảng 2 cột)
        if rows_text_only:
            MAX_ROWS_PER_SLIDE = 3
            table_chunks = [rows_text_only[i:i + MAX_ROWS_PER_SLIDE] for i in range(0, len(rows_text_only), MAX_ROWS_PER_SLIDE)]
            for c_idx, chunk in enumerate(table_chunks):
                t_title = "XI. CẬN LÂM SÀNG ĐÃ CÓ" if c_idx == 0 else "XI. CẬN LÂM SÀNG ĐÃ CÓ (tiếp theo)"
                slide = add_slide_with_header(t_title)
                
                rows_cnt = len(chunk) + 1
                table_shape = slide.shapes.add_table(rows_cnt, 2, Inches(0.8), Inches(1.6), Inches(11.733), Inches(1.0 + len(chunk) * 1.3))
                table = table_shape.table
                table.columns[0].width = Inches(5.866)
                table.columns[1].width = Inches(5.866)

                table.cell(0, 0).text = "KẾT QUẢ CẬN LÂM SÀNG"
                table.cell(0, 1).text = "PHIÊN GIẢI / BIỆN GIẢI"
                for col_i in range(2):
                    cell_p = table.cell(0, col_i).text_frame.paragraphs[0]
                    cell_p.font.bold = True
                    cell_p.font.size = Pt(14)
                    cell_p.font.color.rgb = COLOR_PRIMARY

                for r_i, (kq, pg, _) in enumerate(chunk):
                    table.cell(r_i + 1, 0).text = kq if kq else "-"
                    table.cell(r_i + 1, 1).text = pg if pg else "-"
                    for col_i in range(2):
                        cell_p = table.cell(r_i + 1, col_i).text_frame.paragraphs[0]
                        cell_p.font.size = Pt(13)
                        cell_p.font.color.rgb = COLOR_TEXT

        # 2. Hiển thị các cận lâm sàng có kèm ảnh (Mỗi ảnh 1 slide riêng, chia đôi màn hình)
        temp_img_files = []
        try:
            for kq, pg, img in rows_with_img:
                slide = add_slide_with_header("XI. CẬN LÂM SÀNG ĐÃ CÓ (HÌNH ẢNH)")
                
                # Lưu tạm ảnh để nạp vào slide
                suffix = os.path.splitext(img.name)[1]
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as t_img:
                    t_img.write(img.getbuffer())
                    temp_path = t_img.name
                    temp_img_files.append(temp_path)

                # Nửa bên trái: Ảnh (Kích thước tự động tối đa rộng 5.8 inch, cao 5.0 inch)
                try:
                    slide.shapes.add_picture(temp_path, Inches(0.8), Inches(1.6), width=Inches(5.6))
                except Exception:
                    pass

                # Nửa bên phải: Kết quả & Biện giải
                right_box = slide.shapes.add_textbox(Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.0))
                tf_r = right_box.text_frame
                tf_r.word_wrap = True

                p_kq_title = tf_r.paragraphs[0]
                p_kq_title.text = "Kết quả ghi nhận:"
                p_kq_title.font.bold = True
                p_kq_title.font.size = Pt(15)
                p_kq_title.font.color.rgb = COLOR_PRIMARY
                p_kq_title.space_after = Pt(4)

                p_kq = tf_r.add_paragraph()
                p_kq.text = kq if kq else "Hình ảnh đính kèm"
                p_kq.font.size = Pt(14)
                p_kq.font.color.rgb = COLOR_TEXT
                p_kq.space_after = Pt(16)

                p_pg_title = tf_r.add_paragraph()
                p_pg_title.text = "Biện giải / Phiên giải:"
                p_pg_title.font.bold = True
                p_pg_title.font.size = Pt(15)
                p_pg_title.font.color.rgb = COLOR_PRIMARY
                p_pg_title.space_after = Pt(4)

                p_pg = tf_r.add_paragraph()
                p_pg.text = pg if pg else "-"
                p_pg.font.size = Pt(14)
                p_pg.font.color.rgb = COLOR_TEXT
        finally:
            for p in temp_img_files:
                if os.path.exists(p):
                    try:
                        os.remove(p)
                    except:
                        pass

    # Slide: XII. Chẩn đoán xác định (Đỏ đậm)
    cdxd_lines = [l.strip() for l in str(data.get("chan_doan_xac_dinh", "")).split("\n") if l.strip()]
    if cdxd_lines:
        add_content_with_overflow("XII. CHẨN ĐOÁN XÁC ĐỊNH", cdxd_lines, is_red=True)

    # Slide: XIII. Biện luận chẩn đoán xác định
    blxd_lines = [l.strip() for l in str(data.get("bien_luan_xac_dinh", "")).split("\n") if l.strip()]
    if blxd_lines:
        add_content_with_overflow("XIII. BIỆN LUẬN CHẨN ĐOÁN XÁC ĐỊNH", blxd_lines)

    # Slide: XIV. Điều trị
    dt_lines = []
    for label, key in [("1. Mục tiêu điều trị", "dt_muc_tieu"),
                       ("2. Điều trị cụ thể", "dt_cu_the"),
                       ("3. Theo dõi sau điều trị", "dt_theo_doi")]:
        items = [l.strip() for l in str(data.get(key, "")).split("\n") if l.strip()]
        if items:
            dt_lines.append(f"[{label}]")
            dt_lines.extend([f"- {i}" if not i.startswith("-") else i for i in items])
    if dt_lines:
        add_content_with_overflow("XIV. ĐIỀU TRỊ", dt_lines)

    # Slide: XV. Tiên lượng
    tl_lines = [l.strip() for l in str(data.get("tien_luong", "")).split("\n") if l.strip()]
    if tl_lines:
        add_content_with_overflow("XV. TIÊN LƯỢNG", tl_lines)

    # Slide: XVI. Tư vấn
    tv_lines = [l.strip() for l in str(data.get("tu_van", "")).split("\n") if l.strip()]
    if tv_lines:
        t_label = "XVI. TƯ VẤN" if tl_lines else "XV. TƯ VẤN"
        add_content_with_overflow(t_label, tv_lines)

    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io.getvalue()

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
        
        # Nút bấm tích hợp AI cho phần Đề xuất Cận lâm sàng
        if st.button("🪄 Tự động đề xuất Cận lâm sàng bằng AI", type="primary"):
            if "GEMINI_API_KEY" not in st.secrets:
                st.error("⚠️ Hệ thống chưa được cài đặt API Key bí mật. Vui lòng kiểm tra lại cấu hình Secrets!")
            else:
                with st.spinner("Bác sĩ AI đang phân tích lập luận lâm sàng để chỉ định cận lâm sàng tối ưu..."):
                    try:
                        # Gom dữ kiện lâm sàng quan trọng để định hướng chỉ định
                        context_cls = f"- Tuổi: {st.session_state.get('tuoi')}, Giới tính: {st.session_state.get('gioi_tinh')}\n"
                        context_cls += f"- Tiền sử: {st.session_state.get('ts_noi_khoa')} | {st.session_state.get('ts_ngoai_khoa')}\n"
                        context_cls += f"- Lý do vào viện & Bệnh sử: {st.session_state.get('ly_do_vao_vien')} - {st.session_state.get('benh_su')}\n"
                        context_cls += f"- Thăm khám lâm sàng: {st.session_state.get('kham_toan_than')} | Tuần hoàn: {st.session_state.get('kham_tuan_hoan')} | Hô hấp: {st.session_state.get('kham_ho_hap')}\n"
                        context_cls += f"- Chẩn đoán sơ bộ: {st.session_state.get('chan_doan_so_bo')}\n"
                        context_cls += f"- Chẩn đoán phân biệt: {st.session_state.get('chan_doan_phan_biet')}"

                        # Nạp Key và gọi Model
                        genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
                        model = genai.GenerativeModel('gemini-3.6-flash')

                        # Prompt chuẩn bác sĩ lâm sàng: phân tích logic trước khi chỉ định, chú thích tiếng Anh
                        prompt_cls = f"""
                        Bạn là một bác sĩ lâm sàng thực thụ và giàu kinh nghiệm. Hãy nhìn vào toàn thể ca bệnh dưới đây, phân tích logic trước khi đi vào chi tiết để đưa ra các chỉ định CẬN LÂM SÀNG hợp lý, có tính ứng dụng cao, tránh lạm dụng xét nghiệm nhưng không được bỏ sót tổn thương.
                        Bắt buộc: Khi xuất hiện tên xét nghiệm, kỹ thuật chẩn đoán hình ảnh hoặc thuật ngữ y khoa mới, phải ghi kèm thuật ngữ tiếng Anh trong ngoặc đơn bên cạnh.

                        Dữ kiện ca bệnh:
                        {context_cls}

                        YÊU CẦU ĐẦU RA:
                        Trình bày gạch đầu dòng rõ ràng theo đúng 3 nhóm nhãn sau (không viết thêm lời dẫn thừa):
                        [CLS_XAC_DINH]
                        - Các cận lâm sàng phục vụ chẩn đoán xác định nguyên nhân hoặc phân biệt với các chẩn đoán phân biệt.
                        [CLS_DIEU_TRI]
                        - Các cận lâm sàng đánh giá toàn trạng, chức năng gan thận, bilan tiền phẫu hoặc tiên lượng phục vụ phác đồ điều trị.
                        [CLS_KHAC]
                        - Các cận lâm sàng thường quy, theo dõi tiến triển hoặc tầm soát thêm nếu có yếu tố nguy cơ.
                        """

                        response_cls = model.generate_content(prompt_cls)
                        res_cls_text = response_cls.text

                        # Bóc tách kết quả đưa vào session_state
                        if "[CLS_XAC_DINH]" in res_cls_text and "[CLS_DIEU_TRI]" in res_cls_text:
                            p1 = res_cls_text.split("[CLS_DIEU_TRI]")
                            part_xd = p1[0].replace("[CLS_XAC_DINH]", "").strip()
                            
                            if "[CLS_KHAC]" in p1[1]:
                                p2 = p1[1].split("[CLS_KHAC]")
                                part_dt = p2[0].strip()
                                part_khac = p2[1].strip()
                            else:
                                part_dt = p1[1].strip()
                                part_khac = ""

                            st.session_state["cls_dx_xac_dinh"] = part_xd
                            st.session_state["cls_dx_dieu_tri"] = part_dt
                            st.session_state["cls_dx_khac"] = part_khac
                            st.success("✨ Đã gợi ý danh mục cận lâm sàng thành công! Bạn có thể chỉnh sửa trực tiếp bên dưới.")
                            st.rerun()
                        else:
                            st.error("AI trả về không đúng cấu trúc định dạng, vui lòng thử lại.")
                    except Exception as e:
                        st.error(f"Lỗi khi kết nối với AI: {e}")

        st.caption("Các cận lâm sàng định hướng sẽ tự động điền vào 3 ô bên dưới, bạn có thể tự do chỉnh sửa theo ý muốn.")
        c_cls1, c_cls2, c_cls3 = st.columns(3)
        with c_cls1:
            st.text_area("1. Phục vụ chẩn đoán xác định:", key="cls_dx_xac_dinh", height=130)
        with c_cls2:
            st.text_area("2. Phục vụ điều trị:", key="cls_dx_dieu_tri", height=130)
        with c_cls3:
            st.text_area("3. Cận lâm sàng khác:", key="cls_dx_khac", height=130)
            
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
    with st.expander("XIV. HƯỚNG DẪN VÀ KẾ HOẠCH ĐIỀU TRỊ", expanded=True):
        if st.button("🪄 Tự động xây dựng Kế hoạch Điều trị bằng AI", key="btn_ai_dt", type="primary"):
            if "GEMINI_API_KEY" not in st.secrets:
                st.error("⚠️ Hệ thống chưa được cài đặt API Key bí mật. Vui lòng kiểm tra lại cấu hình Secrets!")
            else:
                with st.spinner("AI đang phân tích phác đồ điều trị và cá thể hóa đơn thuốc..."):
                    try:
                        # Gom dữ kiện kết quả CLS đã có để điều trị an toàn (chức năng gan/thận...)
                        cls_da_co_str = ""
                        so_h = st.session_state.get("so_hang_cls", 3)
                        for i in range(so_h):
                            kq_i = st.session_state.get(f"cls_kq_{i}", "").strip()
                            pg_i = st.session_state.get(f"cls_pg_{i}", "").strip()
                            if kq_i:
                                cls_da_co_str += f"+ {kq_i} -> {pg_i}\n"

                        context_dt = f"- Bệnh nhân: {st.session_state.get('tuoi')} tuổi, Giới tính: {st.session_state.get('gioi_tinh')}\n"
                        context_dt += f"- Tiền sử / Dị ứng: {st.session_state.get('ts_noi_khoa')} | Dị ứng & Ngoại khoa: {st.session_state.get('ts_ngoai_khoa')}\n"
                        context_dt += f"- Chẩn đoán xác định: {st.session_state.get('chan_doan_xac_dinh')}\n"
                        context_dt += f"- Kết quả CLS quan trọng đã có:\n{cls_da_co_str}"

                        genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
                        model = genai.GenerativeModel('gemini-3.1-flash-lite')

                        prompt_dt = f"""
                        Bạn là một bác sĩ điều trị lâm sàng chính. Hãy xây dựng một kế hoạch điều trị toàn diện, cá thể hóa cho ca bệnh dưới đây theo đúng các phác đồ y khoa chuẩn mực (Guidelines).
                        Đặc biệt lưu ý chống chỉ định hoặc chỉnh liều nếu có bệnh nền hoặc bất thường cận lâm sàng.

                        Thông tin ca bệnh:
                        {context_dt}

                        YÊU CẦU ĐẦU RA (đúng 3 tag sau, trình bày rõ ràng từng gạch đầu dòng):
                        [MUC_TIEU]
                        (Nêu mục tiêu ngắn hạn và dài hạn: kiểm soát triệu chứng, ngăn ngừa biến chứng, cải thiện chất lượng sống...)
                        [DIEU_TRI_CU_THE]
                        - Không dùng thuốc (chế độ nghỉ ngơi, dinh dưỡng, lý liệu...)
                        - Dùng thuốc (hoặc điều trị ngoại khoa nếu cần): Ghi rõ tên hoạt chất (kèm tên thương mại phổ biến nếu có), liều lượng, số lần/ngày, đường dùng, thời điểm uống/tiêm.
                        
                        [THEO_DOI]
                        (Các chỉ số sinh tồn, triệu chứng cơ năng, xét nghiệm cần làm lại và lịch đánh giá lại đáp ứng điều trị)
                        """

                        resp = model.generate_content(prompt_dt)
                        txt = resp.text

                        if "[MUC_TIEU]" in txt and "[DIEU_TRI_CU_THE]" in txt and "[THEO_DOI]" in txt:
                            p1 = txt.split("[DIEU_TRI_CU_THE]")
                            dt_mt = p1[0].replace("[MUC_TIEU]", "").strip()
                            p2 = p1[1].split("[THEO_DOI]")
                            dt_ct = p2[0].strip()
                            dt_td = p2[1].strip()

                            st.session_state["dt_muc_tieu"] = dt_mt
                            st.session_state["dt_cu_the"] = dt_ct
                            st.session_state["dt_theo_doi"] = dt_td
                            st.success("✨ Đã lên phác đồ điều trị thành công!")
                            st.rerun()
                        else:
                            st.error("AI phản hồi sai cấu trúc, vui lòng thử lại.")
                    except Exception as e:
                        st.error(f"Lỗi khi kết nối với AI: {e}")

        st.caption("AI sẽ cá thể hóa phác đồ điều trị dựa trên Chẩn đoán xác định, Bệnh nền và Kết quả xét nghiệm.")
        c_mt, c_ct, c_td = st.columns(3)
        with c_mt:
            st.text_area("1. Mục tiêu điều trị:", key="dt_muc_tieu", height=220)
        with c_ct:
            st.text_area("2. Điều trị cụ thể:", key="dt_cu_the", height=220)
        with c_td:
            st.text_area("3. Theo dõi:", key="dt_theo_doi", height=220)

    # 9. TIÊN LƯỢNG & TƯ VẤN (EXPANDER)
    with st.expander("XV VÀ XVI. TIÊN LƯỢNG VÀ TƯ VẤN", expanded=True):
        # Nút bấm tích hợp AI
        if st.button("🪄 Tự động suy luận Tiên lượng & Tư vấn bằng AI", type="primary"):
            # Lấy API Key bí mật từ máy chủ Streamlit
            if "GEMINI_API_KEY" not in st.secrets:
                st.error("⚠️ Hệ thống chưa được cài đặt API Key bí mật. Vui lòng kiểm tra lại cấu hình Secrets!")
            else:
                with st.spinner("AI đang phân tích logic lâm sàng của ca bệnh..."):
                    try:
                        # Gom nhặt các dữ kiện quan trọng nhất để mớm cho AI
                        context = f"- Tuổi: {st.session_state.get('tuoi')}, Giới tính: {st.session_state.get('gioi_tinh')}\n"
                        context += f"- Tiền sử: {st.session_state.get('ts_noi_khoa')} {st.session_state.get('ts_ngoai_khoa')}\n"
                        context += f"- Bệnh sử: {st.session_state.get('benh_su')}\n"
                        context += f"- Chẩn đoán xác định: {st.session_state.get('chan_doan_xac_dinh')}\n"
                        context += f"- Hướng điều trị cụ thể: {st.session_state.get('dt_cu_the')}"

                        # Tự động nạp Key từ máy chủ (st.secrets)
                        genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
                        
                        # Chỉ định đích danh phiên bản Google yêu cầu trong thông báo lỗi
                        model = genai.GenerativeModel('gemini-3.1-flash-lite')

                        # Thiết kế Prompt ép AI suy luận như Bác sĩ thực thụ
                        prompt = f"""
                        Bạn là một bác sĩ lâm sàng thực thụ và giàu kinh nghiệm. Hãy nhìn vào toàn thể ca bệnh dưới đây, phân tích logic trước khi đi vào chi tiết để đưa ra nội dung cho 2 mục: TIÊN LƯỢNG và TƯ VẤN.
                        Đảm bảo ưu tiên tính chính xác về mặt khoa học. Nội dung áp dụng sát với thực tế lâm sàng. Trình bày chỉ cần xuống dòng, không cần gạch đầu dòng.

                        Thông tin ca bệnh tóm tắt:
                        {context}

                        YÊU CẦU ĐẦU RA:
                        Trình bày đúng định dạng sau (không viết thêm văn vẻ thừa):
                        [TIEN_LUONG]
                        (Viết tiên lượng gần và tiên lượng xa. Dựa vào chẩn đoán, thể trạng và bệnh nền).
                        [TU_VAN]
                        (Viết tư vấn cụ thể về chế độ dinh dưỡng, sinh hoạt, phục hồi chức năng, và đặc biệt là các dấu hiệu cảnh báo nguy hiểm cần tái khám ngay. Giải thích dễ hiểu để bệnh nhân có thể áp dụng).
                        """

                        response = model.generate_content(prompt)
                        res_text = response.text

                        # Bóc tách kết quả từ AI
                        if "[TIEN_LUONG]" in res_text and "[TU_VAN]" in res_text:
                            parts = res_text.split("[TU_VAN]")
                            tl_part = parts[0].replace("[TIEN_LUONG]", "").strip()
                            tv_part = parts[1].strip()

                            # Cập nhật trực tiếp vào ô text_area
                            st.session_state["tien_luong"] = tl_part
                            st.session_state["tu_van"] = tv_part
                            st.success("✨ Đã tạo gợi ý thành công! Bạn có thể xem lại và tùy chỉnh nội dung bên dưới.")
                            st.rerun() # Tải lại giao diện để hiện chữ
                        else:
                            st.error("AI trả về sai định dạng, hãy thử bấm lại.")
                    except Exception as e:
                        st.error(f"Lỗi khi kết nối với AI: {e}")

        st.caption("Bạn có thể để AI phân tích và gợi ý, sau đó chỉnh sửa lại nội dung trong ô dưới đây sao cho phù hợp nhất với bệnh nhân.")
        c_pl, c_tv = st.columns(2)
        with c_pl:
            st.text_area("XV. Tiên lượng:", key="tien_luong", height=250, placeholder="Nếu để trống, mục này sẽ không xuất hiện trong file...")
        with c_tv:
            st.text_area("XVI. Tư vấn:", key="tu_van", height=250, placeholder="Nếu để trống, mục này sẽ không xuất hiện trong file...")

# Gom dữ liệu từ session_state sang dict để chuẩn bị xuất file
data_benh_an = {k: st.session_state.get(k, "") for k in default_fields}
data_benh_an["so_hang_cls"] = st.session_state.get("so_hang_cls", 3)
for i in range(data_benh_an["so_hang_cls"]):
    data_benh_an[f"cls_kq_{i}"] = st.session_state.get(f"cls_kq_{i}", "")
    data_benh_an[f"cls_pg_{i}"] = st.session_state.get(f"cls_pg_{i}", "")
data_benh_an.update(uploaded_imgs)


# --- TAB 2: XEM TRƯỚC VÀ XUẤT TẬP TIN ---
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
    col_dl_pdf, col_dl_pptx = st.columns(2)
    
    with col_dl_pdf:
        if st.button("Tạo tập tin PDF bệnh án", type="primary", use_container_width=True):
            if not ho_ten_val:
                st.error("Vui lòng điền tối thiểu Họ và tên người bệnh trước khi xuất tập tin!")
            elif not os.path.exists("Roboto-Regular.ttf") or not os.path.exists("Roboto-Bold.ttf"):
                st.error("Chưa tìm thấy tập tin font 'Roboto-Regular.ttf' và 'Roboto-Bold.ttf' trong cùng thư mục với app.py!")
            else:
                with st.spinner("Đang kết xuất văn bản PDF..."):
                    pdf_bytes = export_pdf(data_benh_an)
                    ten_file = f"Benh_an_{ho_ten_val.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.pdf"
                    st.success("Tạo PDF thành công!")
                    st.download_button(
                        label="Nhấn vào đây để tải PDF về máy",
                        data=pdf_bytes,
                        file_name=ten_file,
                        mime="application/pdf",
                        use_container_width=True
                    )

    with col_dl_pptx:
        if st.button("Tạo tập tin PowerPoint (PPTX)", type="secondary", use_container_width=True):
            if not ho_ten_val:
                st.error("Vui lòng điền tối thiểu Họ và tên người bệnh trước khi xuất tập tin!")
            else:
                with st.spinner("Đang kết xuất bản trình chiếu PowerPoint..."):
                    pptx_bytes = export_pptx(data_benh_an)
                    ten_file_pptx = f"Trinh_chieu_Benh_an_{ho_ten_val.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.pptx"
                    st.success("Tạo PowerPoint thành công!")
                    st.download_button(
                        label="Nhấn vào đây để tải file PowerPoint (.pptx) về máy",
                        data=pptx_bytes,
                        file_name=ten_file_pptx,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True
                    )
