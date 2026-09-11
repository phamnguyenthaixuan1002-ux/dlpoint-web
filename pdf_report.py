from fpdf import FPDF
from datetime import datetime
import os
import json
from collections import Counter

from database import * 
from config import resource_path, GVCN_NHAN_XET_MAU, DIEM_KHOI_DAU, xep_loai_hanh_kiem, get_school_week_number, HOLIDAY_SETTINGS_KEY

# --- Hằng số cho PDF ---
FONT_DIR = resource_path("fonts")
FONT_REGULAR_PATH = os.path.join(FONT_DIR, 'times.ttf')
FONT_BOLD_PATH = os.path.join(FONT_DIR, 'timesbd.ttf')
FONT_ITALIC_PATH = os.path.join(FONT_DIR, 'timesi.ttf')


class PDF(FPDF):
    def __init__(self, month, year):
        super().__init__()
        self.month = month
        self.year = year
        self.full_title = f'PHIẾU BÁO CÁO RÈN LUYỆN THÁNG {self.month}/{self.year}'

    def header(self):
        self.set_font('Times', 'B', 16)
        self.cell(0, 10, self.full_title, 0, 1, 'C')
        self.ln(2) 

    def footer(self):
        self.set_y(-15)
        self.set_font('Times', '', 8)
        self.cell(0, 10, f'Trang {self.page_no()}', 0, 0, 'C')
        self.set_x(-40)
        self.cell(0, 10, f'Ngày in: {datetime.now().strftime("%d/%m/%Y")}', 0, 0, 'R')

# <<< HÀM ĐÃ ĐƯỢC NÂNG CẤP TOÀN DIỆN >>>
def generate_dynamic_comment(classification, events):
    if not events:
        return "Trong tháng vừa qua, chưa có nhiều dữ liệu rèn luyện được ghi nhận. Thầy đề nghị con tích cực tham gia các hoạt động trong tháng tới."

    khen_thuong_counter = Counter(event[1] for event in events if event[2] == 'Khen thưởng')
    vi_pham_counter = Counter(event[1] for event in events if event[2] == 'Vi phạm')

    top_khen_thuong = khen_thuong_counter.most_common(1)
    top_vi_pham = vi_pham_counter.most_common(1)

    comment_parts = []

    # 1. Ghi nhận Ưu điểm
    if top_khen_thuong:
        hanh_vi_tot, so_lan = top_khen_thuong[0]
        if so_lan > 1:
            comment_parts.append(f"Ưu điểm nổi bật: Tích cực {hanh_vi_tot.lower()} (ghi nhận {so_lan} lần).")
        else:
            comment_parts.append(f"Ưu điểm: Có ý thức {hanh_vi_tot.lower()}.")
    else:
        if classification in ["Xuất sắc", "Tốt"]:
             comment_parts.append("Ưu điểm: Chấp hành tốt nội quy, nề nếp ổn định.")

    # 2. Phân tích Hạn chế (nếu có)
    if top_vi_pham:
        hanh_vi_xau, so_lan = top_vi_pham[0]
        if so_lan > 1:
            comment_parts.append(f"Hạn chế cần khắc phục: Còn vi phạm lỗi {hanh_vi_xau.lower()} (lặp lại {so_lan} lần).")
        else:
            comment_parts.append(f"Hạn chế: Cần chú ý hơn để không vi phạm lỗi {hanh_vi_xau.lower()}.")
            
    # 3. Đề xuất Hướng khắc phục / Phát huy
    if classification == "Xuất sắc":
        comment_parts.append("Hướng phát huy: Tiếp tục duy trì và lan tỏa tinh thần tích cực này đến các bạn trong tổ.")
    elif classification == "Tốt":
         comment_parts.append("Hướng phát huy: Duy trì những mặt đã làm tốt và nỗ lực hơn để đạt kết quả xuất sắc trong tháng tới.")
    elif classification == "Khá" and top_vi_pham:
        comment_parts.append(f"Hướng khắc phục: Trong tháng tới, con cần đặt mục tiêu cụ thể là giảm số lần vi phạm lỗi trên xuống dưới {max(1, top_vi_pham[0][1]//2)} lần.")
    else: # Đạt và Cần cố gắng
        comment_parts.append("Hướng khắc phục: Con cần chấn chỉnh nghiêm túc các vi phạm về nề nếp và chủ động hơn trong học tập. Thầy sẽ quan tâm, hỗ trợ thêm.")

    return " ".join(comment_parts)


def generate_pdf_report(student_id, year, month, teacher_name, filepath):
    student_data_tuple = lay_thong_tin_day_du_hoc_sinh_db(student_id)
    if not student_data_tuple: return False, "Không tìm thấy thông tin học sinh."

    db_fields_with_id = ["id"] + STUDENT_FIELDS_DB
    student_data = dict(zip(db_fields_with_id, student_data_tuple))
    events_in_month = lay_su_kien_trong_thang_cua_hoc_sinh_db(student_id, year, month)

    pdf = PDF(month=month, year=year)
    
    try:
        pdf.add_font('Times', '', FONT_REGULAR_PATH, uni=True)
        pdf.add_font('Times', 'B', FONT_BOLD_PATH, uni=True)
        pdf.add_font('Times', 'I', FONT_ITALIC_PATH, uni=True)
    except RuntimeError as e:
        return False, f"Lỗi font: {e}\n\nHãy chắc chắn các file .ttf của Times New Roman có trong thư mục 'fonts'."
    
    pdf.add_page()
    
    pdf.set_font('Times', 'B', 12)
    pdf.cell(0, 10, 'I. THÔNG TIN HỌC SINH', 0, 1)
    pdf.set_font('Times', 'B', 11)
    pdf.cell(40, 7, '   - Họ và tên:'); pdf.set_font('Times', '', 11); pdf.cell(90, 7, student_data.get('ten', ''))
    pdf.set_font('Times', 'B', 11); pdf.cell(15, 7, 'Lớp:'); pdf.set_font('Times', '', 11); pdf.cell(0, 7, student_data.get('lop', ''))
    pdf.ln()
    pdf.set_font('Times', 'B', 11)
    pdf.cell(40, 7, '   - Ngày sinh:'); pdf.set_font('Times', '', 11); pdf.cell(90, 7, student_data.get('ngay_sinh', ''))
    pdf.set_font('Times', 'B', 11); pdf.cell(15, 7, 'Tổ:'); pdf.set_font('Times', '', 11); pdf.cell(0, 7, student_data.get('to_nhiem_vu', ''))
    pdf.ln(4) 

    pdf.set_font('Times', 'B', 12)
    pdf.cell(0, 10, 'II. TỔNG KẾT RÈN LUYỆN THEO TUẦN', 0, 1)
    weekly_summary = []
    if events_in_month:
        start_date_str = load_setting('school_year_start_date', '2025-09-08'); holidays_json = load_setting(HOLIDAY_SETTINGS_KEY, '[]')
        school_start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date(); holidays_raw = json.loads(holidays_json)
        holidays_in_year = [(datetime.strptime(start, '%Y-%m-%d').date(), datetime.strptime(end, '%Y-%m-%d').date()) for start, end in holidays_raw]
        events_by_week = {}
        for event in events_in_month:
            week_num = get_school_week_number(event[4].date(), school_start_date, holidays_in_year)
            if week_num not in events_by_week: events_by_week[week_num] = []
            events_by_week[week_num].append(event[3])
        for week_num in sorted(events_by_week.keys()):
            week_score = DIEM_KHOI_DAU + sum(events_by_week[week_num])
            weekly_summary.append({'week': week_num, 'score': week_score, 'rank': xep_loai_hanh_kiem(week_score)})
    if not weekly_summary:
        pdf.set_font('Times', '', 11); pdf.cell(0, 7, '   - Không có dữ liệu rèn luyện theo tuần.', 0, 1)
    else:
        pdf.set_font('Times', 'B', 10); pdf.set_fill_color(230, 230, 230)
        pdf.cell(60, 8, 'Tuần số', 1, 0, 'C', 1); pdf.cell(65, 8, 'Tổng điểm', 1, 0, 'C', 1); pdf.cell(65, 8, 'Xếp loại', 1, 1, 'C', 1)
        pdf.set_font('Times', '', 10)
        for summary in weekly_summary:
            pdf.cell(60, 7, f"Tuần {summary['week']}", 1, 0, 'C'); pdf.cell(65, 7, str(summary['score']), 1, 0, 'C'); pdf.cell(65, 7, summary['rank'], 1, 1, 'C')
    pdf.ln(4)

    pdf.set_font('Times', 'B', 12)
    pdf.cell(0, 10, 'III. CHI TIẾT CÁC GHI NHẬN TRONG THÁNG', 0, 1)
    current_score = DIEM_KHOI_DAU
    if not events_in_month:
        pdf.set_font('Times', '', 11); pdf.cell(0, 10, '   - Không có ghi nhận nào trong tháng.', 0, 1)
    else:
        pdf.set_font('Times', 'B', 10); pdf.set_fill_color(230, 230, 230)
        pdf.cell(30, 8, 'Ngày', 1, 0, 'C', 1); pdf.cell(110, 8, 'Nội dung', 1, 0, 'C', 1)
        pdf.cell(25, 8, 'Loại', 1, 0, 'C', 1); pdf.cell(25, 8, 'Điểm', 1, 1, 'C', 1)
        pdf.set_font('Times', '', 10)
        for event in events_in_month:
            current_score += event[3]
            content_text = f' {event[1]}'; line_height = 6
            num_lines = len(pdf.multi_cell(110, line_height, content_text, split_only=True))
            row_height = num_lines * line_height
            start_x, start_y = pdf.get_x(), pdf.get_y()
            pdf.rect(start_x, start_y, 30, row_height); pdf.multi_cell(30, line_height, event[4].strftime('%d/%m/%Y'), align='C')
            pdf.set_xy(start_x + 30, start_y)
            pdf.rect(start_x + 30, start_y, 110, row_height); pdf.multi_cell(110, line_height, content_text, align='L')
            pdf.set_xy(start_x + 140, start_y)
            pdf.rect(start_x + 140, start_y, 25, row_height); pdf.multi_cell(25, line_height, event[2], align='C')
            pdf.set_xy(start_x + 165, start_y)
            pdf.rect(start_x + 165, start_y, 25, row_height); pdf.multi_cell(25, line_height, f"{'+' if event[3] > 0 else ''}{event[3]}", align='C')
            pdf.set_y(start_y + row_height)
    pdf.ln(4) 

    pdf.set_font('Times', 'B', 12)
    pdf.cell(0, 10, 'IV. TỔNG KẾT VÀ NHẬN XÉT', 0, 1)
    final_score = max(0, current_score)
    final_classification = xep_loai_hanh_kiem(final_score)
    pdf.set_font('Times', 'B', 11); pdf.cell(50, 7, '   - Điểm rèn luyện tháng:'); pdf.set_font('Times', '', 11); pdf.cell(0, 7, f'{final_score} điểm', 0, 1)
    pdf.set_font('Times', 'B', 11); pdf.cell(50, 7, '   - Xếp loại tháng:'); pdf.set_font('Times', 'B', 12); pdf.cell(0, 7, f'{final_classification}', 0, 1)
    pdf.ln(3)
    pdf.set_font('Times', 'B', 11); pdf.cell(50, 7, '   - Nhận xét của GVCN:')
    
    dynamic_comment = generate_dynamic_comment(final_classification, events_in_month)
    pdf.set_font('Times', '', 11); pdf.multi_cell(0, 6, dynamic_comment)
    pdf.ln(8)

    SIGNATURE_BLOCK_HEIGHT = 45
    if pdf.get_y() + SIGNATURE_BLOCK_HEIGHT > (pdf.h - pdf.b_margin): pdf.add_page()
    pdf.set_font('Times', 'I', 11)
    today_date = "Tân Lộc, ngày {:>2} tháng {:>2} năm {}".format(datetime.now().day, datetime.now().month, datetime.now().year)
    pdf.cell(0, 8, today_date, 0, 1, 'R'); pdf.ln(1)
    y_before_signatures = pdf.get_y(); cell_height = 6
    pdf.set_y(y_before_signatures); pdf.set_x(20); pdf.set_font('Times', 'B', 11); pdf.multi_cell(60, cell_height, 'PHỤ HUYNH HỌC SINH', 0, 'C')
    pdf.set_y(y_before_signatures + cell_height); pdf.set_x(20); pdf.set_font('Times', '', 11); pdf.multi_cell(60, cell_height, '(Ký và ghi rõ họ tên)', 0, 'C')
    pdf.set_y(y_before_signatures + cell_height * 4); pdf.set_x(20); pdf.multi_cell(60, cell_height, '...........................................', 0, 'C')
    pdf.set_y(y_before_signatures); pdf.set_x(75); pdf.set_font('Times', 'B', 11); pdf.multi_cell(60, cell_height, 'HỌC SINH', 0, 'C')
    pdf.set_y(y_before_signatures + cell_height); pdf.set_x(75); pdf.set_font('Times', '', 11); pdf.multi_cell(60, cell_height, '(Ký và ghi rõ họ tên)', 0, 'C')
    pdf.set_y(y_before_signatures + cell_height * 4); pdf.set_x(75); pdf.set_font('Times', 'B', 11); pdf.multi_cell(60, cell_height, student_data.get('ten', ''), 0, 'C')
    pdf.set_y(y_before_signatures); pdf.set_x(130); pdf.set_font('Times', 'B', 11); pdf.multi_cell(60, cell_height, 'GIÁO VIÊN CHỦ NHIỆM', 0, 'C')
    pdf.set_y(y_before_signatures + cell_height); pdf.set_x(130); pdf.set_font('Times', '', 11); pdf.multi_cell(60, cell_height, '(Ký và ghi rõ họ tên)', 0, 'C')
    pdf.set_y(y_before_signatures + cell_height * 4); pdf.set_x(130); pdf.set_font('Times', 'B', 11); pdf.multi_cell(60, cell_height, teacher_name, 0, 'C')
    
    try:
        pdf.output(filepath)
        return True, "Tạo báo cáo PDF thành công."
    except Exception as e:
        return False, f"Lỗi khi lưu file PDF: {e}"