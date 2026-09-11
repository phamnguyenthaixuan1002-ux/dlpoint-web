# File: app.py
from PIL import Image, ImageDraw, ImageFont # Dùng để vẽ Bằng khen
import io
import streamlit as st
import pandas as pd
import json
from datetime import date, datetime, timedelta
import tempfile
import os
import calendar
import io # <<< THÊM THƯ VIỆN NÀY (Để xử lý file Excel mẫu tải về)

from database import (
    init_db, verify_user, lay_danh_sach_hoc_sinh_db,
    lay_tat_ca_danh_muc_su_kien_db, lay_chi_tiet_danh_muc_su_kien_db,
    lay_lich_su_ky_luat_cua_hoc_sinh_db, them_su_kien_va_ky_luat_db,
    lay_su_kien_cho_thong_ke_df, lay_su_kien_trong_khoang_ngay_db,
    lay_thong_tin_day_du_hoc_sinh_db, get_week_dates_by_number,
    STUDENT_FIELDS_DB, lay_du_lieu_bieu_do_ca_nhan, luu_muc_tieu_phan_hoi_db,
    get_all_users_db, them_hoac_cap_nhat_user_db, 
    them_hoac_cap_nhat_danh_muc_db, load_setting, save_setting,
    get_distinct_classes_and_groups_db,
    them_hoc_sinh_db, xoa_nhieu_hoc_sinh_db,
    them_su_kien_ren_luyen_db # <<< ĐÃ BỔ SUNG HÀM BỊ THIẾU Ở ĐÂY
)
from config import DIEM_KHOI_DAU, xep_loai_hanh_kiem
from excel_export import generate_weekly_summary_excel, generate_monthly_summary_excel
from config import DIEM_KHOI_DAU, xep_loai_hanh_kiem, get_school_week_number, HOLIDAY_SETTINGS_KEY # <<< THÊM 2 HÀM/BIẾN CUỐI


# --- CẤU HÌNH TRANG WEB ---
# Lệnh này phải đặt ở đầu tiên
st.set_page_config(page_title="DLPOINT Web", page_icon="🎓", layout="wide")

# Khởi tạo kết nối CSDL (dùng chung file database.py cũ của thầy)
init_db()

# --- QUẢN LÝ TRẠNG THÁI ĐĂNG NHẬP (SESSION STATE) ---
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.user_info = None

# --- HÀM 1: GIAO DIỆN ĐĂNG NHẬP ---
def show_login_page():
    st.title("🎓 Hệ thống Quản lý DLPOINT 2.0")
    st.markdown("---")
    
    # Tạo 3 cột để căn giữa form đăng nhập
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.subheader("Đăng nhập")
        username = st.text_input("Tên đăng nhập")
        password = st.text_input("Mật khẩu", type="password")
        
        if st.button("Đăng nhập", width="stretch", type="primary"):
            user = verify_user(username, password)
            if user:
                # Đăng nhập thành công, lưu thông tin vào session
                st.session_state.logged_in = True
                st.session_state.user_info = user
                st.rerun() # Tải lại trang web
            else:
                st.error("Tên đăng nhập hoặc mật khẩu không đúng!")

# --- HÀM 2: GIAO DIỆN CHÍNH (SAU KHI ĐĂNG NHẬP) ---
# --- HÀM 3: QUẢN LÝ LỚP HỌC ---
# --- HÀM 3: QUẢN LÝ LỚP HỌC & HỒ SƠ 360 ---
def show_class_management():
    st.header("👨‍🎓 Quản lý Lớp học & Hồ sơ 360°")
    st.markdown("---")
    
    user = st.session_state.user_info
    raw_data = lay_danh_sach_hoc_sinh_db(user['role'], user['class'], user['group'])
    
    if not raw_data:
        st.info("Chưa có dữ liệu học sinh nào trong phạm vi quản lý của thầy/cô.")
        return

    # 1. THU GỌN BẢNG DANH SÁCH LỚP VÀO MỘT KHUNG CÓ THỂ ĐÓNG/MỞ (EXPANDER)
    with st.expander("📋 Click để xem danh sách tổng quát của lớp", expanded=False):
        df = pd.DataFrame(raw_data, columns=["ID", "Họ và Tên", "Lớp", "Tổ", "SĐT Học sinh", "SĐT Zalo", "Ảnh thẻ"])
        df.insert(0, 'STT', range(1, 1 + len(df)))
        st.dataframe(df[['STT', 'Họ và Tên', 'Lớp', 'Tổ', 'SĐT Học sinh', "SĐT Zalo"]], width="stretch", hide_index=True)

    # 2. KHU VỰC HỒ SƠ 360 ĐỘ
    st.subheader("🔍 Tra cứu Hồ sơ 360°")
    
    # Tạo danh sách chọn học sinh
    student_dict = {f"{hs[1]} (Lớp {hs[2]})": hs[0] for hs in raw_data}
    selected_student_name = st.selectbox("Chọn học sinh để xem hồ sơ chi tiết:", options=["-- Hãy chọn một học sinh --"] + list(student_dict.keys()))

    if selected_student_name != "-- Hãy chọn một học sinh --":
        hs_id = student_dict[selected_student_name]
        
        # Lấy thông tin chi tiết từ CSDL
        student_info_tuple = lay_thong_tin_day_du_hoc_sinh_db(student_id=hs_id)
        if not student_info_tuple:
            st.error("Lỗi: Không tìm thấy thông tin chi tiết.")
            return
            
        # Ghép tên cột với dữ liệu
        info = dict(zip(["id"] + STUDENT_FIELDS_DB, student_info_tuple))

        # --- TẠO 3 TAB HIỂN THỊ ---
        tab_tong_quan, tab_lich_su, tab_muc_tieu = st.tabs([
            "📝 Tổng quan & Biểu đồ", 
            "📜 Lịch sử Rèn luyện & Kỷ luật", 
            "🎯 Mục tiêu & Phản hồi GVCN"
        ])

        # === TAB 1: TỔNG QUAN ===
        with tab_tong_quan:
            col_info, col_chart = st.columns([1, 1.5], gap="large")
            
            with col_info:
                st.markdown(f"### {info.get('ten', 'Chưa cập nhật')}")
                st.write(f"**Lớp:** {info.get('lop', '')} | **Tổ:** {info.get('to_nhiem_vu', '')}")
                st.write(f"**Ngày sinh:** {info.get('ngay_sinh', '')}")
                st.write(f"**SĐT Học sinh:** {info.get('sdt_hoc_sinh', 'Không có')}")
                st.write(f"**Phụ huynh:** {info.get('ten_cha', '')} / {info.get('ten_me', '')}")
                st.write(f"**SĐT Phụ huynh:** {info.get('sdt_cha', '')} / {info.get('sdt_me', '')}")
                
            with col_chart:
                st.write("**📈 Xu hướng rèn luyện (4 tuần gần nhất)**")
                chart_data = lay_du_lieu_bieu_do_ca_nhan(hs_id, num_weeks=4)
                if chart_data:
                    # Chuyển đổi dữ liệu sang định dạng Streamlit hiểu để vẽ
                    weeks = [f"Tuần {int(d[0])}" for d in chart_data]
                    scores = [DIEM_KHOI_DAU + d[1] for d in chart_data]
                    
                    df_chart = pd.DataFrame({"Tổng điểm": scores}, index=weeks)
                    # Vẽ biểu đồ Line cực nhanh của Streamlit
                    st.line_chart(df_chart, color="#0078D7")
                else:
                    st.info("Học sinh chưa có dữ liệu rèn luyện trong 4 tuần qua để vẽ biểu đồ.")

        # === TAB 2: LỊCH SỬ RÈN LUYỆN ===
        with tab_lich_su:
            st.write("**Lịch sử Kỷ luật (Theo TT19)**")
            history_kl = lay_lich_su_ky_luat_cua_hoc_sinh_db(hs_id)
            if history_kl:
                df_kl = pd.DataFrame(history_kl, columns=["ID", "Hình thức", "Ngày áp dụng", "Ghi chú"])
                st.dataframe(df_kl[["Ngày áp dụng", "Hình thức", "Ghi chú"]], width="stretch", hide_index=True)
            else:
                st.success("Học sinh chưa bị áp dụng hình thức kỷ luật nào.")
            
            st.write("**Chi tiết các Sự kiện Rèn luyện (+/- điểm)**")
            # Tái sử dụng hàm lay_su_kien_trong_khoang_ngay_db với mốc thời gian rộng
            all_events = lay_su_kien_trong_khoang_ngay_db('2020-01-01', '2100-01-01', user['role'], user['class'], user['group'])
            hs_events = [e for e in all_events if e[0] == hs_id]
            
            if hs_events:
                # Cấu trúc tuple: (hs_id, ten, lop, to, mo_ta, loai_su_kien, diem, ngay_tao)
                df_events = pd.DataFrame(hs_events, columns=["ID", "Tên", "Lớp", "Tổ", "Nội dung", "Loại", "Điểm", "Ngày ghi nhận"])
                st.dataframe(df_events[["Ngày ghi nhận", "Nội dung", "Loại", "Điểm"]].sort_values("Ngày ghi nhận", ascending=False), width="stretch", hide_index=True)
            else:
                st.info("Chưa có sự kiện rèn luyện nào được ghi nhận.")

        # === TAB 3: MỤC TIÊU & PHẢN HỒI ===
        with tab_muc_tieu:
            st.write("Giáo viên Chủ nhiệm có thể ghi chú mục tiêu phấn đấu và nhận xét cá nhân tại đây:")
            
            # Khởi tạo giá trị trong session_state để quản lý Form
            muc_tieu_cu = info.get('muc_tieu_thang') or ""
            phan_hoi_cu = info.get('phan_hoi_gvcn') or ""
            
            new_muc_tieu = st.text_area("🎯 Mục tiêu tháng tới:", value=muc_tieu_cu, height=100)
            new_phan_hoi = st.text_area("💬 Nhận xét & Phản hồi của GVCN:", value=phan_hoi_cu, height=150)
            
            if st.button("💾 Lưu Mục tiêu & Phản hồi", type="primary"):
                if luu_muc_tieu_phan_hoi_db(hs_id, new_muc_tieu, new_phan_hoi):
                    st.toast("Đã lưu mục tiêu và phản hồi thành công!", icon="✅")
                else:
                    st.error("Có lỗi xảy ra khi lưu vào CSDL.")
# --- HÀM 8: GHI NHẬN NHANH (TỐI ƯU CHO BCS & GVCN) ---
def show_quick_record_page():
    st.header("📝 Ghi nhận Điểm Cộng / Trừ Nhanh")
    st.markdown("---")

    user = st.session_state.user_info
    
    # 1. Chuẩn bị dữ liệu Học sinh và Sự kiện
    raw_students = lay_danh_sach_hoc_sinh_db(user['role'], user['class'], user['group'])
    if not raw_students:
        st.warning("Không có học sinh nào trong phạm vi quản lý.")
        return

    student_dict = {f"{hs[1]} (Tổ {hs[3] or '?'})": hs[0] for hs in raw_students}
    
    all_events = lay_tat_ca_danh_muc_su_kien_db()
    # Tạo dictionary cho sự kiện để hiển thị kèm điểm cho trực quan: "Tên sự kiện (+2đ)"
    event_dict = {}
    for ev in all_events:
        # ev: (id, ten, loai, diem, muc_do)
        ten, loai, diem = ev[1], ev[2], ev[3]
        diem_str = f"+{diem}đ" if diem > 0 else f"{diem}đ"
        label = f"[{loai}] {ten} ({diem_str})"
        event_dict[label] = (ten, loai, diem) # Lưu lại giá trị gốc để ghi vào DB

    # Lựa chọn ngày chung cho cả 2 tab
    event_date = st.date_input("📅 Chọn Ngày ghi nhận:", format="DD/MM/YYYY")
    ngay_tao_str = event_date.strftime('%Y-%m-%d %H:%M:%S')
    st.markdown("---")

    # 2. Tạo 2 Tab cho 2 kịch bản nhập liệu
    tab1, tab2 = st.tabs(["👥 Nhiều Học sinh -> Cùng 1 Lỗi/Khen", "👤 1 Học sinh -> Cùng lúc Nhiều Lỗi/Khen"])

    # ==========================================
    # KỊCH BẢN 1: 1 SỰ KIỆN ÁP DỤNG CHO NHIỀU HS
    # ==========================================
    with tab1:
        st.info("💡 Dùng khi: Nguyên một nhóm mất trật tự, hoặc cả tổ đều mang đủ sách vở...")
        col1_1, col1_2 = st.columns([1, 1])
        
        with col1_1:
            selected_event_1 = st.selectbox("1. Chọn Sự kiện:", options=["-- Chọn --"] + list(event_dict.keys()), key="ev_tab1")
        with col1_2:
            selected_students_1 = st.multiselect("2. Chọn các Học sinh:", options=list(student_dict.keys()), key="hs_tab1")
            
        if st.button("🚀 Ghi nhận cho danh sách trên", type="primary", key="btn_tab1"):
            if selected_event_1 != "-- Chọn --" and selected_students_1:
                ten_sk, loai_sk, diem_sk = event_dict[selected_event_1]
                count = 0
                for hs_name in selected_students_1:
                    hs_id = student_dict[hs_name]
                    if them_su_kien_ren_luyen_db(hs_id, ten_sk, loai_sk, diem_sk, ngay_tao_str):
                        count += 1
                st.success(f"✅ Đã ghi nhận **{ten_sk}** cho **{count}** học sinh thành công!")
            else:
                st.error("Vui lòng chọn 1 sự kiện và ít nhất 1 học sinh.")

    # ==========================================
    # KỊCH BẢN 2: NHIỀU SỰ KIỆN CHO 1 HS
    # ==========================================
    with tab2:
        st.info("💡 Dùng khi: Một em vừa đi trễ, vừa không thuộc bài, nhưng lại có phát biểu gỡ điểm...")
        col2_1, col2_2 = st.columns([1, 1])
        
        with col2_1:
            selected_student_2 = st.selectbox("1. Chọn Học sinh:", options=["-- Chọn --"] + list(student_dict.keys()), key="hs_tab2")
        with col2_2:
            selected_events_2 = st.multiselect("2. Chọn các Sự kiện:", options=list(event_dict.keys()), key="ev_tab2")

        if st.button("🚀 Ghi nhận các sự kiện trên", type="primary", key="btn_tab2"):
            if selected_student_2 != "-- Chọn --" and selected_events_2:
                hs_id = student_dict[selected_student_2]
                count = 0
                for ev_label in selected_events_2:
                    ten_sk, loai_sk, diem_sk = event_dict[ev_label]
                    if them_su_kien_ren_luyen_db(hs_id, ten_sk, loai_sk, diem_sk, ngay_tao_str):
                        count += 1
                
                # Tách tên học sinh cho đẹp
                hs_ten_ngan = selected_student_2.split("(")[0].strip()
                st.success(f"✅ Đã ghi nhận **{count}** sự kiện cho học sinh **{hs_ten_ngan}** thành công!")
            else:
                st.error("Vui lòng chọn 1 học sinh và ít nhất 1 sự kiện.")
# --- HÀM 4: GHI NHẬN KỶ LUẬT (TT19) ---
def show_discipline_page():
    st.header("⚖️ Ghi nhận Kỷ luật (Theo Thông tư 19)")
    st.markdown("---")

    user = st.session_state.user_info
    
    # 1. Lấy danh sách Học sinh & Lỗi vi phạm từ CSDL
    raw_students = lay_danh_sach_hoc_sinh_db(user['role'], user['class'], user['group'])
    if not raw_students:
        st.warning("Không có học sinh nào trong phạm vi quản lý.")
        return

    # Tạo một "Từ điển" để web hiển thị Tên đẹp mắt, nhưng ngầm hiểu ID ở bên trong
    student_dict = {f"{hs[1]} (Lớp {hs[2]})": hs[0] for hs in raw_students}
    
    all_events = lay_tat_ca_danh_muc_su_kien_db()
    # Lấy các sự kiện là Vi phạm và có Mức độ (1, 2 hoặc 3)
    violation_events = sorted([e[1] for e in all_events if e[2] == "Vi phạm" and e[4] is not None])

    # 2. Tạo giao diện 2 cột
    col1, col2 = st.columns([1, 1], gap="large")

    # --- CỘT TRÁI: NHẬP LIỆU ---
    with col1:
        st.subheader("1. Chọn đối tượng & Vi phạm")
        selected_student_name = st.selectbox("Chọn học sinh:", options=["-- Chọn học sinh --"] + list(student_dict.keys()))
        event_date = st.date_input("Ngày vi phạm:", format="DD/MM/YYYY")
        selected_violation = st.selectbox("Chọn vi phạm:", options=["-- Chọn lỗi vi phạm --"] + violation_events)

    # --- CỘT PHẢI: PHÂN TÍCH & ĐỀ XUẤT ---
    with col2:
        st.subheader("2. Phân tích & Đề xuất (Tự động)")
        
        # Chỉ phân tích khi GV đã chọn đủ HS và Lỗi
        if selected_student_name != "-- Chọn học sinh --" and selected_violation != "-- Chọn lỗi vi phạm --":
            student_id = student_dict[selected_student_name]
            event_details = lay_chi_tiet_danh_muc_su_kien_db(selected_violation)

            if event_details and event_details[4]:
                muc_do = event_details[4]
                st.info(f"**Mức độ vi phạm:** Mức {muc_do}")

                # Lấy lịch sử cũ
                history = lay_lich_su_ky_luat_cua_hoc_sinh_db(student_id)

                # Logic đề xuất (Giống hệt app Desktop)
                de_xuat = ""
                if muc_do == 1:
                    if any(h[1] == "Nhắc nhở" for h in history): de_xuat = "Phê bình"
                    else: de_xuat = "Nhắc nhở"
                elif muc_do == 2:
                    if any(h[1] == "Phê bình" for h in history): de_xuat = "Viết bản tự kiểm điểm"
                    else: de_xuat = "Phê bình"
                elif muc_do == 3:
                    de_xuat = "Viết bản tự kiểm điểm"

                st.success(f"**Đề xuất xử lý:** {de_xuat}")

                # Nút xem nhanh lịch sử
                with st.expander("Xem nhanh lịch sử kỷ luật của học sinh này"):
                    if history:
                        for h in history[:3]: # Hiện 3 lần gần nhất
                            st.write(f"- {h[1]} (Ngày: {h[2].strftime('%d/%m/%Y')})")
                    else:
                        st.write("Chưa có lịch sử kỷ luật.")

                # Form xác nhận cuối cùng
                st.markdown("---")
                hinh_thuc_ap_dung = st.selectbox(
                    "Xác nhận Hình thức áp dụng:", 
                    ["Nhắc nhở", "Phê bình", "Viết bản tự kiểm điểm"], 
                    index=["Nhắc nhở", "Phê bình", "Viết bản tự kiểm điểm"].index(de_xuat) if de_xuat else 0
                )
                ghi_chu = st.text_area("Ghi chú của GV:")

                # Nút lưu dữ liệu
                if st.button("💾 Lưu và Áp dụng", type="primary", width="stretch"):
                    _, _, loai_sk, diem_sk, _ = event_details
                    ngay_tao_str = event_date.strftime('%Y-%m-%d %H:%M:%S')

                    success, msg = them_su_kien_va_ky_luat_db(
                        student_id, selected_violation, loai_sk, diem_sk, ngay_tao_str, hinh_thuc_ap_dung, ghi_chu
                    )
                    
                    if success:
                        st.toast(f"Đã lưu kỷ luật thành công!", icon="✅") # Hiện thông báo pop-up đẹp ở góc màn hình
                    else:
                        st.error(f"Lỗi: {msg}")
        else:
            st.write("👈 Vui lòng chọn học sinh và lỗi vi phạm ở cột bên trái để hệ thống phân tích.")
# --- HÀM 5: THỐNG KÊ & BÁO CÁO ---
def show_statistics_page():
    st.header("📊 Thống kê & Báo cáo Rèn luyện")
    st.markdown("---")

    # 1. Bộ lọc thời gian (Nằm ngang cho đẹp)
    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        # Mặc định lấy lùi lại 30 ngày từ hôm nay
        start_date = st.date_input("Từ ngày:", date.today() - timedelta(days=30), format="DD/MM/YYYY")
    with col2:
        end_date = st.date_input("Đến ngày:", date.today(), format="DD/MM/YYYY")
    with col3:
        st.write("") # Tạo khoảng trống cho nút bấm ngang hàng
        st.write("")
        btn_thong_ke = st.button("Lấy dữ liệu Thống kê", type="primary", width="stretch")

    st.markdown("---")

    # 2. Xử lý và Vẽ biểu đồ khi bấm nút
    if btn_thong_ke:
        if start_date > end_date:
            st.error("Ngày bắt đầu không được lớn hơn ngày kết thúc!")
            return
            
        with st.spinner('Đang tổng hợp dữ liệu...'):
            # Gọi hàm CSDL trả về Pandas DataFrame (Giống hệt app cũ)
            df = lay_su_kien_cho_thong_ke_df(start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d'))
            
            if df.empty:
                st.warning("Không có dữ liệu rèn luyện nào trong khoảng thời gian này.")
                return

            # Lọc ra chỉ các sự kiện "Vi phạm"
            df_vipham = df[df['loai_su_kien'] == 'Vi phạm']

            if df_vipham.empty:
                st.success("Tuyệt vời! Không có vi phạm nào trong khoảng thời gian này.")
                return

            # --- VẼ BIỂU ĐỒ (Sức mạnh của Streamlit) ---
            
            # Khung chứa 2 biểu đồ nằm cạnh nhau
            chart_col1, chart_col2 = st.columns(2)
            
            with chart_col1:
                st.subheader("1. Tổng Vi phạm theo Lớp")
                # Đếm số vi phạm theo lớp
                vp_theo_lop = df_vipham['lop'].value_counts()
                # Vẽ biểu đồ cột trực tiếp từ Pandas Series chỉ với 1 dòng lệnh
                st.bar_chart(vp_theo_lop, color="#ff4b4b") 

            with chart_col2:
                st.subheader("2. Top Lỗi Vi phạm phổ biến")
                # Đếm số lượng theo tên mô tả lỗi (Top 10)
                top_loi = df_vipham['mo_ta'].value_counts().head(10)
                st.bar_chart(top_loi, color="#ffa600")

            # Bảng chi tiết ở dưới cùng
            st.subheader("3. Chi tiết các dữ liệu vi phạm")
            # Hiển thị bảng có thể lọc, tìm kiếm
            st.dataframe(df_vipham[['lop', 'to_nhiem_vu', 'mo_ta', 'diem_ap_dung', 'ngay_tao']], width="stretch")
# --- HÀM 6: TỔNG KẾT & XUẤT EXCEL ---
def show_summary_page():
    st.header("📑 Tổng kết & Xuất Báo cáo Excel")
    st.markdown("---")
    
    user = st.session_state.user_info
    
    # Tạo 2 tab trên giao diện Web
    tab_tuan, tab_thang = st.tabs(["📅 Tổng kết Tuần", "🗓️ Tổng kết Tháng"])
    
    # ==========================================
    # TAB 1: TỔNG KẾT TUẦN
    # ==========================================
    with tab_tuan:
        col_input1, col_input2 = st.columns([1, 3])
        with col_input1:
            week_num = st.number_input("Chọn tuần số:", min_value=1, max_value=52, value=1, step=1)
            btn_tuan = st.button("Xem Tổng Kết Tuần", type="primary", width="stretch")
            
        # Dùng session_state để giữ bảng dữ liệu không bị mất khi ấn nút Tải File
        if btn_tuan or st.session_state.get('show_tuan', False):
            st.session_state.show_tuan = True 
            
            start_date, end_date = get_week_dates_by_number(week_num)
            if not start_date:
                st.error("Không thể xác định ngày của tuần này. Hãy kiểm tra Cài đặt năm học.")
            else:
                end_date_inc = end_date + timedelta(days=1)
                st.info(f"**Tuần {week_num}:** Từ {start_date.strftime('%d/%m/%Y')} đến {end_date.strftime('%d/%m/%Y')}")
                
                # --- Tính toán dữ liệu ---
                with st.spinner("Đang tính toán điểm..."):
                    all_events = lay_su_kien_trong_khoang_ngay_db(start_date.strftime('%Y-%m-%d'), end_date_inc.strftime('%Y-%m-%d'), user['role'], user['class'], user['group'])
                    all_students = lay_thong_tin_day_du_hoc_sinh_db(user_role=user['role'], assigned_class=user['class'], assigned_group=user['group'])
                    
                    events_by_student = {}
                    for ev in all_events:
                        hs_id = ev[0]
                        if hs_id not in events_by_student: events_by_student[hs_id] = []
                        events_by_student[hs_id].append(ev)
                        
                    student_list = []
                    for hs in all_students:
                        hs_id, ten_hs, lop_hs, to_hs = hs[0], hs[1], hs[2], hs[3] or ""
                        hs_events = events_by_student.get(hs_id, [])
                        score = DIEM_KHOI_DAU + sum(e[6] for e in hs_events) # Cột 6 là điểm
                        final_score = max(0, score)
                        student_list.append({
                            'id': hs_id, 'ten': ten_hs, 'lop': lop_hs, 'to': to_hs, 
                            'diem': final_score, 'xeploai': xep_loai_hanh_kiem(final_score)
                        })
                    student_list.sort(key=lambda x: (-x['diem'], x['to']))
                
                # Hiển thị bảng
                df_tuan = pd.DataFrame(student_list)
                df_tuan.insert(0, 'STT', range(1, len(df_tuan) + 1))
                st.dataframe(df_tuan[['STT', 'ten', 'lop', 'to', 'diem', 'xeploai']], width="stretch", hide_index=True)
                
                # --- XUẤT EXCEL TUẦN ---
                st.markdown("---")
                
                # Tạo file Excel ngầm trong hệ thống
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    filepath = tmp.name
                    
                # Chạy file code xuất Excel cũ của thầy    
                success, msg = generate_weekly_summary_excel(
                    student_list, week_num, user['full_name'], user['class'] or "Toàn trường", user['group'] or "", 
                    start_date, end_date, filepath, [], [], [] 
                )
                
                if success:
                    # Nếu thành công, đọc file đó lên web và tạo nút Tải Xuống
                    with open(filepath, "rb") as f:
                        excel_data = f.read()
                    
                    st.download_button(
                        label="📥 Tải xuống File Excel Tuần",
                        data=excel_data,
                        file_name=f"TongKet_Tuan_{week_num}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        width="stretch"
                    )
                else:
                    st.error(f"Lỗi tạo Excel: {msg}")
                os.unlink(filepath) # Dọn rác file tạm
                
    # ==========================================
    # TAB 2: TỔNG KẾT THÁNG
    # ==========================================
    with tab_thang:
        col_m1, col_m2, col_m3 = st.columns([1, 1, 2])
        with col_m1:
            month = st.selectbox("Chọn tháng:", range(1, 13), index=datetime.now().month - 1)
        with col_m2:
            year = st.number_input("Chọn năm:", value=datetime.now().year)
        with col_m3:
            st.write("")
            st.write("")
            btn_thang = st.button("Xem Báo cáo Tháng", type="primary", width="stretch")
            
        if btn_thang or st.session_state.get('show_thang', False):
            st.session_state.show_thang = True
            
            start_date_month = datetime(year, month, 1)
            _, num_days = calendar.monthrange(year, month)
            end_date_month = datetime(year, month, num_days)
            
            st.info(f"**Tháng {month}/{year}:** Từ {start_date_month.strftime('%d/%m/%Y')} đến {end_date_month.strftime('%d/%m/%Y')}")
            
            with st.spinner("Đang tính toán điểm tháng..."):
                all_events_m = lay_su_kien_trong_khoang_ngay_db(start_date_month.strftime('%Y-%m-%d'), end_date_month.strftime('%Y-%m-%d'), user['role'], user['class'], user['group'])
                all_students_m = lay_thong_tin_day_du_hoc_sinh_db(user_role=user['role'], assigned_class=user['class'], assigned_group=user['group'])
                
                events_by_student_m = {}
                for ev in all_events_m:
                    hs_id = ev[0]
                    if hs_id not in events_by_student_m: events_by_student_m[hs_id] = []
                    events_by_student_m[hs_id].append(ev)
                    
                student_list_m = []
                for hs in all_students_m:
                    hs_id, ten_hs, lop_hs, to_hs = hs[0], hs[1], hs[2], hs[3] or ""
                    hs_events = events_by_student_m.get(hs_id, [])
                    score = DIEM_KHOI_DAU + sum(e[6] for e in hs_events)
                    student_list_m.append({
                        'id': hs_id, 'ten': ten_hs, 'lop': lop_hs, 'to': to_hs, 
                        'diem': score, 'xeploai': xep_loai_hanh_kiem(score)
                    })
                student_list_m.sort(key=lambda x: (-x['diem'], x['to']))
            
            df_thang = pd.DataFrame(student_list_m)
            df_thang.insert(0, 'STT', range(1, len(df_thang) + 1))
            st.dataframe(df_thang[['STT', 'ten', 'lop', 'to', 'diem', 'xeploai']], width="stretch", hide_index=True)
            
            # --- XUẤT EXCEL THÁNG ---
            st.markdown("---")
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_m:
                filepath_m = tmp_m.name
                
            success_m, msg_m = generate_monthly_summary_excel(student_list_m, month, year, user['full_name'], user['class'] or "Toàn trường", filepath_m)
            
            if success_m:
                with open(filepath_m, "rb") as f_m:
                    excel_data_m = f_m.read()
                
                st.download_button(
                    label="📥 Tải xuống File Excel Tháng",
                    data=excel_data_m,
                    file_name=f"TongKet_Thang_{month}_{year}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    width="stretch"
                )
            else:
                st.error(f"Lỗi tạo Excel: {msg_m}")
            os.unlink(filepath_m)
# --- HÀM 7: QUẢN TRỊ HỆ THỐNG (ADMIN) ---
def show_admin_page():
    st.header("⚙️ Quản trị Hệ thống")
    st.markdown("---")

    # Kiểm tra quyền: Chỉ Admin mới được xem trang này
    user = st.session_state.user_info
    if user['role'] != 'admin':
        st.error("🔒 Chức năng này bảo mật, chỉ dành cho tài khoản Admin.")
        return

  # TẠO 4 TAB (Đã thêm Tab thứ 4)
    tab_users, tab_events, tab_settings, tab_class = st.tabs([
        "👥 Quản lý Người dùng", 
        "📋 Danh mục Sự kiện", 
        "🗓️ Cài đặt Năm học",
        "🏫 Quản lý Lớp & HS (Excel)" # <<< TAB MỚI
    ])
    # ==========================================
    # TAB 1: QUẢN LÝ NGƯỜI DÙNG
    # ==========================================
    with tab_users:
        col_list, col_form = st.columns([1.5, 1], gap="large")
        
        with col_list:
            st.subheader("Danh sách Tài khoản")
            users = get_all_users_db()
            if users:
                df_users = pd.DataFrame(users, columns=["ID", "Tên đăng nhập", "Họ tên", "Vai trò", "Lớp", "Tổ"])
                st.dataframe(df_users, width="stretch", hide_index=True)
                
        with col_form:
            st.subheader("Thêm/Sửa Tài khoản")
            st.info("Nhập Tên đăng nhập đã có để SỬA, hoặc Tên mới để THÊM.")
            with st.form("form_user"):
                username = st.text_input("Tên đăng nhập *")
                fullname = st.text_input("Họ và tên *")
                password = st.text_input("Mật khẩu (Nhập để đổi/tạo mới)", type="password")
                role = st.selectbox("Vai trò", ["admin", "gvcn", "bcs"])
                aclass = st.text_input("Lớp quản lý (Để trống nếu là admin)")
                agroup = st.text_input("Tổ quản lý (Gõ số, VD: 1)")

                if st.form_submit_button("💾 Lưu Tài khoản", type="primary"):
                    if username and fullname and role:
                        if them_hoac_cap_nhat_user_db(username, fullname, password, role, aclass or None, agroup or None):
                            st.success(f"Đã lưu tài khoản **{username}** thành công!")
                            st.rerun() # Tải lại trang để cập nhật bảng
                    else:
                        st.error("Vui lòng nhập đủ các trường bắt buộc (*).")

    # ==========================================
    # TAB 2: QUẢN LÝ DANH MỤC SỰ KIỆN (TT19)
    # ==========================================
    with tab_events:
        col_elist, col_eform = st.columns([1.5, 1], gap="large")
        
        with col_elist:
            st.subheader("Danh sách Lỗi & Khen thưởng")
            events = lay_tat_ca_danh_muc_su_kien_db()
            if events:
                df_events = pd.DataFrame(events, columns=["ID", "Tên Sự kiện", "Loại", "Điểm", "Mức độ Vi phạm"])
                st.dataframe(df_events, width="stretch", hide_index=True)

        with col_eform:
            st.subheader("Thêm/Sửa Sự kiện")
            with st.form("form_event"):
                e_name = st.text_input("Tên sự kiện *")
                e_type = st.selectbox("Loại", ["Vi phạm", "Khen thưởng"])
                e_points = st.number_input("Điểm (+ hoặc -)", value=0)
                e_level = st.selectbox("Mức độ vi phạm theo TT19 (Chỉ dành cho lỗi)", ["Không xét KL", "1", "2", "3"])

                if st.form_submit_button("💾 Lưu Sự kiện", type="primary"):
                    if e_name:
                        # Logic: Khen thưởng thì không có mức độ kỷ luật
                        muc_do = int(e_level) if e_level != "Không xét KL" and e_type == "Vi phạm" else None
                        if them_hoac_cap_nhat_danh_muc_db(e_name, e_type, e_points, muc_do):
                            st.success(f"Đã lưu sự kiện **{e_name}**!")
                            st.rerun()
                    else:
                        st.error("Vui lòng nhập tên sự kiện.")

    # ==========================================
    # TAB 3: CÀI ĐẶT CHUNG
    # ==========================================
    with tab_settings:
        st.subheader("Cài đặt Ngày bắt đầu Năm học")
        st.write("Ngày này được dùng để tính số thứ tự Tuần học trong các Báo cáo.")
        
        current_start = load_setting("school_year_start_date", "2025-09-08")
        start_date = st.date_input("Ngày bắt đầu năm học (Ngày đầu tiên của Tuần 1):", datetime.strptime(current_start, "%Y-%m-%d").date())
        
        if st.button("💾 Lưu Cài đặt", type="primary"):
            save_setting("school_year_start_date", start_date.strftime("%Y-%m-%d"))
            st.toast("Đã lưu ngày bắt đầu năm học thành công!", icon="✅")
     # TAB 4: QUẢN LÝ LỚP & HỌC SINH (NHẬP/XÓA HÀNG LOẠT)
    # ==========================================
    with tab_class:
        st.info("Khu vực này giúp thầy khởi tạo lớp mới nhanh chóng bằng Excel, hoặc dọn dẹp dữ liệu học sinh lớp cũ/chuyển trường.")
        
        col_import, col_delete = st.columns([1, 1], gap="large")

        # ---------------------------------------
        # CỘT TRÁI: TẠO LỚP BẰNG EXCEL
        # ---------------------------------------
        with col_import:
            st.subheader("📥 1. Tạo Lớp (Nhập từ Excel)")
            
            # Tạo file Excel mẫu ngay trên bộ nhớ để tải về
            df_template = pd.DataFrame(columns=["Họ tên học sinh", "Lớp", "Tổ", "Ngày tháng năm sinh", "Địa chỉ", "SĐT Học sinh", "SĐT Zalo", "Ghi chú"])
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_template.to_excel(writer, index=False)
            excel_data = output.getvalue()

            st.download_button(
                label="⬇️ Tải File Excel Mẫu",
                data=excel_data,
                file_name="Mau_Nhap_Hoc_Sinh.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.markdown("---")
            uploaded_file = st.file_uploader("Tải lên file danh sách học sinh (đã điền):", type=['xlsx', 'xls'])
            
            if uploaded_file is not None:
                if st.button("🚀 Xử lý & Nhập vào hệ thống", type="primary"):
                    try:
                        df_import = pd.read_excel(uploaded_file)
                        headers = df_import.columns.tolist()
                        
                        if "Họ tên học sinh" not in headers or "Lớp" not in headers:
                            st.error("File Excel không hợp lệ. Bắt buộc phải có cột 'Họ tên học sinh' và 'Lớp'.")
                        else:
                            # Map cột Excel sang cột DB
                            db_map = {
                                "Họ tên học sinh": "ten", "Lớp": "lop", "Tổ": "to_nhiem_vu", 
                                "Ngày tháng năm sinh": "ngay_sinh", "Địa chỉ": "dia_chi",
                                "SĐT Học sinh": "sdt_hoc_sinh", "SĐT Zalo": "sdt_zalo", "Ghi chú": "ghi_chu"
                            }
                            
                            success_count, fail_count = 0, 0
                            with st.spinner("Đang lưu dữ liệu vào CSDL..."):
                                for _, row in df_import.iterrows():
                                    if pd.isna(row.get("Họ tên học sinh")) or pd.isna(row.get("Lớp")):
                                        continue # Bỏ qua dòng trống
                                    
                                    data_db = {}
                                    for excel_col, db_col in db_map.items():
                                        val = row.get(excel_col)
                                        data_db[db_col] = str(val).strip() if pd.notna(val) else ""
                                        
                                    if them_hoc_sinh_db(data_db): success_count += 1
                                    else: fail_count += 1
                                    
                            st.success(f"Đã nhập thành công {success_count} học sinh. Bỏ qua/Lỗi: {fail_count} dòng.")
                    except Exception as e:
                        st.error(f"Có lỗi khi đọc file: {e}")

        # ---------------------------------------
        # CỘT PHẢI: XÓA HỌC SINH / XÓA LỚP
        # ---------------------------------------
        with col_delete:
            st.subheader("🗑️ 2. Dọn dẹp Dữ liệu (Xóa)")
            
            classes, _ = get_distinct_classes_and_groups_db()
            if not classes:
                st.write("Hiện chưa có lớp nào trong hệ thống.")
            else:
                delete_mode = st.radio("Chọn phương thức xóa:", ["Xóa từng học sinh được chọn", "Xóa toàn bộ một lớp"])
                
                if delete_mode == "Xóa toàn bộ một lớp":
                    class_to_delete = st.selectbox("Chọn lớp muốn XÓA HOÀN TOÀN:", classes)
                    st.warning(f"⚠️ CẢNH BÁO: Hành động này sẽ xóa vĩnh viễn toàn bộ học sinh lớp {class_to_delete}, bao gồm cả điểm số, lịch sử kỷ luật của các em.")
                    
                    # Xác nhận 2 bước để chống bấm nhầm
                    confirm_delete = st.checkbox("Tôi hiểu rủi ro và xác nhận muốn xóa lớp này.")
                    if confirm_delete:
                        if st.button(f"🚨 TIẾN HÀNH XÓA LỚP {class_to_delete}", type="primary", use_container_width=True):
                            # Lấy ID của tất cả HS lớp đó (chạy bằng quyền admin)
                            hs_lop_do = lay_danh_sach_hoc_sinh_db('admin', class_to_delete, None)
                            ids_to_delete = [hs[0] for hs in hs_lop_do]
                            
                            if ids_to_delete:
                                deleted_count = xoa_nhieu_hoc_sinh_db(ids_to_delete)
                                st.success(f"Đã xóa thành công {deleted_count} học sinh thuộc lớp {class_to_delete} khỏi hệ thống.")
                                st.rerun() # Tải lại trang

                elif delete_mode == "Xóa từng học sinh được chọn":
                    class_to_filter = st.selectbox("Lọc danh sách theo lớp:", ["Tất cả"] + classes)
                    
                    # Lấy danh sách HS để hiển thị
                    c_filter = None if class_to_filter == "Tất cả" else class_to_filter
                    hs_list = lay_danh_sach_hoc_sinh_db('admin', c_filter, None)
                    
                    if hs_list:
                        # Biến danh sách thành DataFrame và thêm cột Checkbox "Chọn Xóa"
                        df_hs = pd.DataFrame(hs_list, columns=["ID", "Họ Tên", "Lớp", "Tổ", "SĐT", "Zalo", "Ảnh"])
                        df_hs = df_hs[["ID", "Họ Tên", "Lớp", "Tổ"]] # Chỉ lấy cột cần thiết
                        df_hs.insert(0, "Chọn xóa", False) # Cột checkbox mặc định là False
                        
                        st.write("Đánh dấu tick [x] vào các học sinh muốn xóa:")
                        # Hiển thị bảng cho phép chỉnh sửa (tick checkbox)
                        edited_df = st.data_editor(
                            df_hs, 
                            hide_index=True, 
                            column_config={"Chọn xóa": st.column_config.CheckboxColumn(required=True)},
                            disabled=["ID", "Họ Tên", "Lớp", "Tổ"], # Khóa các cột thông tin, chỉ cho tick
                            use_container_width=True
                        )
                        
                        # Lọc ra các ID đã được tick True
                        selected_rows = edited_df[edited_df["Chọn xóa"] == True]
                        ids_to_delete = selected_rows["ID"].tolist()
                        
                        if len(ids_to_delete) > 0:
                            st.warning(f"Bạn đang chọn xóa {len(ids_to_delete)} học sinh.")
                            if st.button("🗑️ Xác nhận xóa các HS đã chọn", type="primary"):
                                deleted_count = xoa_nhieu_hoc_sinh_db(ids_to_delete)
                                st.success(f"Đã xóa thành công {deleted_count} học sinh.")
                                st.rerun()
# --- HÀM 2: GIAO DIỆN CHÍNH (SAU KHI ĐĂNG NHẬP) ---
# --- HÀM 2: GIAO DIỆN CHÍNH (SAU KHI ĐĂNG NHẬP) ---
def show_main_dashboard():
    user = st.session_state.user_info
    role, aclass, agroup = user['role'], user['class'], user['group']
    
    # --- THANH BÊN (SIDEBAR) VÀ PHÂN QUYỀN MENU ---
    st.sidebar.title(f"Chào, {user['full_name']} 👋")
    st.sidebar.caption(f"Vai trò: {role.upper()} | Lớp: {aclass or 'Toàn trường'}")
    st.sidebar.markdown("---")
    
    # 1. Menu cơ bản ai cũng thấy (BCS, GVCN, Admin)
    menu_options = [
        "🏠 Bảng điều khiển", 
        "👨‍🎓 Quản lý Lớp học", 
        "📝 Ghi nhận Nhanh"
    ]
    
    # 2. Menu chỉ dành cho GVCN và Admin (BCS không thấy)
    if role in ['gvcn', 'admin']:
        menu_options.extend([
            "⚖️ Ghi nhận Kỷ luật (TT19)", 
            "📊 Thống kê & Báo cáo", 
            "📑 Tổng kết & Xuất Excel"
        ])
        
    # 3. Menu chỉ dành cho Admin
    if role == 'admin':
        menu_options.append("⚙️ Quản trị Hệ thống")

    choice = st.sidebar.radio("Chọn chức năng:", menu_options)
    
    st.sidebar.markdown("---")
    if st.sidebar.button("Đăng xuất", width="stretch"):
        st.session_state.logged_in = False
        st.session_state.user_info = None
        st.rerun()

    # --- KHU VỰC NỘI DUNG CHÍNH ---
    if choice == "🏠 Bảng điều khiển":
        st.header("📊 Bảng điều khiển Tổng quan")
        st.markdown("---")

        # ==========================================
        # 1. TÍNH TOÁN DỮ LIỆU THỰC TẾ
        # ==========================================
        hs_list = lay_danh_sach_hoc_sinh_db(role, aclass, agroup)
        total_hs = len(hs_list) if hs_list else 0

        today = date.today()
        try:
            start_date_str = load_setting('school_year_start_date', '2025-09-08')
            holidays_json = load_setting(HOLIDAY_SETTINGS_KEY, '[]')
            school_start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()
            holidays_raw = json.loads(holidays_json)
            holidays_in_year = [(datetime.strptime(s, '%Y-%m-%d').date(), datetime.strptime(e, '%Y-%m-%d').date()) for s, e in holidays_raw]
            current_week = get_school_week_number(today, school_start_date, holidays_in_year)
            if current_week <= 0: current_week = 1
        except:
            current_week = 1
            
        start_w, end_w = get_week_dates_by_number(current_week)
        if not start_w: # Dự phòng nếu lỗi
            start_w = today - timedelta(days=today.weekday())
            end_w = start_w + timedelta(days=6)

        events_week = lay_su_kien_trong_khoang_ngay_db(start_w.strftime('%Y-%m-%d'), (end_w + timedelta(days=1)).strftime('%Y-%m-%d'), role, aclass, agroup)
        
        khen_thuong_count = sum(1 for ev in events_week if ev[5] == "Khen thưởng")
        vi_pham_count = sum(1 for ev in events_week if ev[5] == "Vi phạm")

        # ==========================================
        # 2. HIỂN THỊ GIAO DIỆN (UI)
        # ==========================================
        st.markdown(f"##### 📅 Thông số Tuần {current_week} (Từ {start_w.strftime('%d/%m')} đến {end_w.strftime('%d/%m')})")
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Tổng Học sinh", total_hs, "Đang quản lý")
        col2.metric("Sự kiện (Tuần này)", len(events_week), f"{khen_thuong_count} Tốt / {vi_pham_count} Lỗi", delta_color="off")
        col3.metric("Khen thưởng", khen_thuong_count, "Tích cực", delta_color="normal")
        col4.metric("Vi phạm", vi_pham_count, "Cần lưu ý", delta_color="inverse")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        col_chart, col_feed = st.columns([1.5, 1], gap="large")
        
        with col_chart:
            st.subheader("📈 Xu hướng 30 ngày qua")
            start_30d = today - timedelta(days=30)
            events_30d = lay_su_kien_trong_khoang_ngay_db(start_30d.strftime('%Y-%m-%d'), (today + timedelta(days=1)).strftime('%Y-%m-%d'), role, aclass, agroup)
            
            if events_30d:
                df_30d = pd.DataFrame(events_30d, columns=["hs_id", "ten", "lop", "to", "mo_ta", "loai", "diem", "ngay"])
                df_30d['Ngày'] = pd.to_datetime(df_30d['ngay']).dt.strftime('%d/%m')
                daily_stats = df_30d.groupby(['Ngày', 'loai']).size().unstack(fill_value=0)
                st.bar_chart(daily_stats, color=["#00D274", "#FF4B4B"] if len(daily_stats.columns) == 2 else None)
            else:
                st.info("Chưa có dữ liệu rèn luyện để vẽ biểu đồ 30 ngày qua.")

        with col_feed:
            st.subheader("🔔 Hoạt động mới nhất")
            if events_week:
                recent_events = events_week[:5]
                for ev in recent_events:
                    is_khen = ev[5] == "Khen thưởng"
                    icon = "🟢" if is_khen else "🔴"
                    border_color = "#00D274" if is_khen else "#ff4b4b"
                    bg_color = "#f4fbfa" if is_khen else "#fff4f4"
                    
                    st.markdown(f"""
                    <div style='background-color: {bg_color}; padding: 12px; border-radius: 8px; margin-bottom: 10px; border-left: 5px solid {border_color}; box-shadow: 1px 1px 4px rgba(0,0,0,0.05);'>
                        <strong style='font-size: 1.1em;'>{icon} {ev[1]}</strong> <span style='color: gray; font-size: 0.9em;'>({ev[2]})</span><br>
                        <span style='color: #333;'>{ev[4]}</span><br>
                        <i style='color: gray; font-size: 0.8em;'>Lúc: {ev[7].strftime('%H:%M - %d/%m')}</i>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.write("Tuần này lớp/trường khá im ắng, chưa có hoạt động nào được ghi nhận.")

    # ==========================================
    # CÁC ĐOẠN MÃ ĐIỀU HƯỚNG TỚI CÁC TRANG KHÁC
    # ==========================================
    elif choice == "👨‍🎓 Quản lý Lớp học":
        show_class_management()
        
    elif choice == "📝 Ghi nhận Nhanh":
        show_quick_record_page()
        
    elif choice == "⚖️ Ghi nhận Kỷ luật (TT19)":
        show_discipline_page()
        
    elif choice == "📊 Thống kê & Báo cáo":
        show_statistics_page()
        
    elif choice == "📑 Tổng kết & Xuất Excel":
        show_summary_page()
        
    elif choice == "⚙️ Quản trị Hệ thống":
        show_admin_page()

# --- ĐIỀU HƯỚNG ---
if not st.session_state.logged_in:
    show_login_page()
else:
    show_main_dashboard()
