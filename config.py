# File: config.py

import os
import sys
from datetime import date, datetime, timedelta

# --- CÁC HẰNG SỐ MỚI THÊM ---
SPEAKING_EVENT_DESCRIPTION = "Phát biểu xây dựng bài"
SPEAKING_TOP_N = 5 # Số lượng học sinh tích cực nhất cần hiển thị
DISORDERLY_EVENT_DESCRIPTION = "Mất trật tự, nói chuyện riêng" # <<< THÊM MỚI

# --- CÁC HẰNG SỐ DÙNG CHUNG ---
DIEM_KHOI_DAU = 100
MUC_DO_XU_LY = {
    0: "Không Xử Lý",
    1: "Mức 1: Nhắc nhở",
    2: "Mức 2: GVCN làm việc riêng",
    3: "Mức 3: Viết bản tự kiểm điểm (có PH ký)",
    4: "Mức 4: Mời PH đến trường làm việc"
}
# CẬP NHẬT THANG ĐIỂM XẾP LOẠI MỚI
THRESHOLD_XUATSAC_MIN = 120
THRESHOLD_TOT_MIN = 100
THRESHOLD_KHA_MIN = 85
THRESHOLD_DAT_MIN = 70

GVCN_NHAN_XET_MAU = {
    "Xuất sắc": [
        "Con là một tấm gương xuất sắc về nề nếp và học tập. Rất đáng biểu dương.",
        "Điểm rèn luyện của con cho thấy sự nỗ lực vượt trội, luôn là ngôi sao sáng của lớp.",
    ],
    "Tốt": [
        "Con có nhiều tiến bộ, tích cực tham gia các hoạt động của lớp.",
        "Ngoan, lễ phép, hoàn thành tốt các nhiệm vụ được giao.",
        "Luôn gương mẫu, là tấm gương tốt cho các bạn noi theo.",
        "Chấp hành tốt nội quy, có ý thức giúp đỡ bạn bè trong học tập."
    ],
    "Khá": [
        "Con có cố gắng nhưng đôi lúc còn chưa tập trung trong giờ học.",
        "Cần tích cực phát biểu xây dựng bài hơn nữa.",
        "Nhìn chung đã hoàn thành các yêu cầu, cần nỗ lực để đạt kết quả tốt hơn.",
        "Chấp hành nội quy của lớp, tuy nhiên cần chủ động hơn trong các hoạt động tập thể."
    ],
    "Đạt": [
        "Con cần cố gắng nhiều hơn trong học tập và rèn luyện.",
        "Còn vi phạm nội quy, cần nghiêm túc chấn chỉnh.",
        "Chưa thực sự tập trung vào việc học, cần có sự quan tâm hơn từ gia đình.",
        "Chưa có ý thức phấn đấu, hay làm ảnh hưởng đến các bạn trong lớp."
    ],
    "Cần cố gắng": [
        "Con thường xuyên vi phạm nội quy của trường, lớp. Cần có biện pháp giáo dục kịp thời.",
        "Kết quả rèn luyện chưa tốt, cần sự phối hợp chặt chẽ giữa gia đình và nhà trường.",
        "Cần nghiêm khắc kiểm điểm và thay đổi thái độ học tập và rèn luyện.",
        "Chưa có ý thức phấn đấu, hay làm ảnh hưởng đến các bạn trong lớp."
    ]
}

# --- CÁC HẰNG SỐ CẤU HÌNH CƠ SỞ DỮ LIỆU ---
QUICK_BUTTON_SETTINGS_KEY = "quick_buttons"
HOLIDAY_SETTINGS_KEY = "holidays" 

# --- CÁC HÀM TIỆN ÍCH DÙNG CHUNG ---
def resource_path(relative_path):
    try:
        # PyInstaller tạo một thư mục tạm thời và lưu đường dẫn trong _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # Lấy đường dẫn của file script đang chạy làm gốc
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)

def xep_loai_hanh_kiem(diem):
    if diem >= THRESHOLD_XUATSAC_MIN:
        return "Xuất sắc"
    if diem >= THRESHOLD_TOT_MIN:
        return "Tốt"
    if diem >= THRESHOLD_KHA_MIN:
        return "Khá"
    if diem >= THRESHOLD_DAT_MIN:
        return "Đạt"
    return "Cần cố gắng"

def is_date_in_holiday(check_date, holidays_in_year):
    """Kiểm tra một ngày có nằm trong kỳ nghỉ không."""
    for start_holiday, end_holiday in holidays_in_year:
        if start_holiday <= check_date <= end_holiday:
            return True
    return False

def get_school_week_number(current_date, start_date_of_school_year, holidays_in_year):
    """
    Tính số tuần học chính xác, có bỏ qua các kỳ nghỉ.
    Args:
        current_date (date): Ngày hiện tại cần tính tuần học.
        start_date_of_school_year (date): Ngày bắt đầu năm học.
        holidays_in_year (list of tuples): Danh sách các kỳ nghỉ [(start_date, end_date), ...].
    Returns:
        int: Số thứ tự của tuần học.
    """
    if current_date < start_date_of_school_year:
        return 0

    # Chuyển đổi datetime sang date để so sánh
    if isinstance(current_date, datetime):
        current_date = current_date.date()
    if isinstance(start_date_of_school_year, datetime):
        start_date_of_school_year = start_date_of_school_year.date()
    
    current_iter_date = start_date_of_school_year
    school_week_count = 0
    
    # Lặp qua từng tuần, tối đa 52 tuần học để tránh vòng lặp vô hạn
    while current_iter_date <= current_date and school_week_count < 53:
        # Giả định một tuần học từ Thứ Hai đến Thứ Bảy
        week_end = current_iter_date + timedelta(days=5) # Chủ Nhật
        
        # Kiểm tra xem có bất kỳ ngày nào trong tuần này là ngày học không
        is_learning_week = False
        temp_date = current_iter_date
        while temp_date <= week_end:
            if not is_date_in_holiday(temp_date, holidays_in_year):
                is_learning_week = True
                break
            temp_date += timedelta(days=1)
            
        if is_learning_week:
            school_week_count += 1
        
        # Di chuyển đến tuần tiếp theo
        current_iter_date += timedelta(weeks=1)

    return school_week_count