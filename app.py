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
import base64 # <<< THÊM THƯ VIỆN NÀY ĐỂ XỬ LÝ ẢNH TRÊN WEB
import zipfile  # <<< THÊM DÒNG NÀY
import google.generativeai as genai
import streamlit.components.v1 as components # Thư viện để nhúng mã PWA
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
    them_su_kien_ren_luyen_db,
    cap_nhat_anh_the_db # <<< THÊM HÀM NÀY
    xoa_nhieu_danh_muc_su_kien_db # <<< THÊM HÀM NÀY
)
from config import DIEM_KHOI_DAU, xep_loai_hanh_kiem
from excel_export import generate_weekly_summary_excel, generate_monthly_summary_excel
from config import DIEM_KHOI_DAU, xep_loai_hanh_kiem, get_school_week_number, HOLIDAY_SETTINGS_KEY # <<< THÊM 2 HÀM/BIẾN CUỐI
# Thêm hàm tạo PDF từ file cũ
from pdf_report import generate_pdf_report  # <<< THÊM DÒNG NÀY
from streamlit_option_menu import option_menu # Thư viện tạo Menu bo góc
# BỘ NHỚ ĐỆM (CACHING) GIÚP APP CHẠY SIÊU NHANH TRÊN MOBILE
# =========================================================
@st.cache_data(ttl=600) # Tự động làm mới sau 10 phút
def get_cached_students(role, aclass, agroup):
    return lay_danh_sach_hoc_sinh_db(role, aclass, agroup)

@st.cache_data(ttl=600)
def get_cached_events():
    return lay_tat_ca_danh_muc_su_kien_db()
# =========================================================
# --- CẤU HÌNH TRANG WEB ---
# Lệnh này phải đặt ở đầu tiên
st.set_page_config(page_title="DLPOINT Web", page_icon="🎓", layout="wide", initial_sidebar_state="collapsed")
# =========================================================
# THIẾT LẬP APP TRÀN VIỀN TRÊN ĐIỆN THOẠI (PWA)
# =========================================================
def setup_pwa():
    pwa_code = """
    <script>
        // Lấy thẻ head của trang web mẹ
        const parentHead = window.parent.document.getElementsByTagName('head')[0];
        
        // Tránh chèn nhiều lần khi tải lại trang
        if (!window.parent.document.getElementById('pwa-manifest')) {
            
            // 1. Khai báo cho iPhone/iPad (iOS)
            const meta1 = window.parent.document.createElement('meta');
            meta1.name = "apple-mobile-web-app-capable";
            meta1.content = "yes";
            parentHead.appendChild(meta1);

            const meta2 = window.parent.document.createElement('meta');
            meta2.name = "apple-mobile-web-app-status-bar-style";
            meta2.content = "default";
            parentHead.appendChild(meta2);

            const meta3 = window.parent.document.createElement('meta');
            meta3.name = "apple-touch-icon";
            meta3.content = "https://cdn-icons-png.flaticon.com/512/149/149071.png"; // Có thể thay link logo trường thầy vào đây
            parentHead.appendChild(meta3);

            // 2. Tạo file Cấu hình ứng dụng (Manifest) cho Android/Chrome
            const manifestJSON = {
                "name": "DLPOINT App",
                "short_name": "DLPOINT",
                "description": "Phần mềm Quản lý Rèn luyện",
                "start_url": ".",
                "display": "standalone", // Lệnh này giúp app mở tràn viền, mất thanh URL
                "background_color": "#ffffff",
                "theme_color": "#0056b3",
                "icons": [{
                    "src": "https://cdn-icons-png.flaticon.com/512/149/149071.png",
                    "sizes": "512x512",
                    "type": "image/png"
                }]
            };
            
            // Chèn file cấu hình vào trang web
            const blob = new Blob([JSON.stringify(manifestJSON)], {type: 'application/json'});
            const manifestURL = URL.createObjectURL(blob);
            const linkManifest = window.parent.document.createElement('link');
            linkManifest.id = 'pwa-manifest';
            linkManifest.rel = 'manifest';
            linkManifest.href = manifestURL;
            parentHead.appendChild(linkManifest);
        }
    </script>
    """
    components.html(pwa_code, height=0, width=0)
# Gọi hàm chạy ngầm ngay khi mở web
setup_pwa()
# =========================================================
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
# --- HÀM 3: QUẢN LÝ LỚP HỌC & HỒ SƠ 360 ---
# --- HÀM 3: QUẢN LÝ LỚP HỌC & HỒ SƠ 360 (GIAO DIỆN TƯƠNG TÁC MỚI) ---
def show_class_management():
    st.header("👨‍🎓 Quản lý Lớp học & Hồ sơ 360°")
    st.markdown("---")
    
    user = st.session_state.user_info
    raw_data = lay_danh_sach_hoc_sinh_db(user['role'], user['class'], user['group'])
    
    if not raw_data:
        st.info("Chưa có dữ liệu học sinh nào trong phạm vi quản lý của thầy/cô.")
        return

    # Chuẩn bị dữ liệu bảng
    df = pd.DataFrame(raw_data, columns=["ID", "Họ và Tên", "Lớp", "Tổ", "SĐT Học sinh", "SĐT Zalo", "Ảnh thẻ"])
    df.insert(0, 'STT', range(1, 1 + len(df)))

    # Chia layout: Danh sách (trái) - Chi tiết (phải). Trên điện thoại nó sẽ tự xếp dọc.
    col_list, col_detail = st.columns([1.2, 2], gap="large")

    # ==========================================
    # KHU VỰC 1: DANH SÁCH HỌC SINH (LUÔN HIỂN THỊ)
    # ==========================================
    with col_list:
        st.write("👇 **Bấm vào 1 dòng để xem hồ sơ:**")
        
        # Tạo bảng tương tác (Bấm vào dòng nào, Streamlit tự nhận diện dòng đó)
        selection = st.dataframe(
            df[['ID', 'STT', 'Họ và Tên', 'Tổ']], 
            hide_index=True,
            use_container_width=True,
            height=500, # Giới hạn chiều cao để có thanh cuộn, không làm trang web quá dài
            column_config={
                "ID": None, # Ẩn cột ID đi cho đẹp, nhưng vẫn giữ để truy vấn
                "STT": st.column_config.NumberColumn(width="small"),
                "Tổ": st.column_config.TextColumn(width="small")
            },
            selection_mode="single-row", # Chỉ cho phép chọn 1 dòng
            on_select="rerun" # Tự động chạy lại app khi bấm chọn
        )

    # ==========================================
    # KHU VỰC 2: HỒ SƠ 360 ĐỘ CỦA HS ĐƯỢC CHỌN
    # ==========================================
    with col_detail:
        # Kiểm tra xem thầy có đang bấm chọn dòng nào trong bảng không
        selected_rows = selection.selection.rows
        
        if len(selected_rows) == 0:
            # Giao diện chờ khi chưa chọn ai
            st.info("👈 Vui lòng bấm chọn một học sinh từ danh sách bên trái để xem Hồ sơ 360°.")
            st.image("https://cdn-icons-png.flaticon.com/512/3135/3135715.png", width=150) # Ảnh minh họa vui vẻ
        else:
            # Lấy ID của học sinh đang được bấm chọn
            selected_index = selected_rows[0]
            hs_id = int(df.iloc[selected_index]['ID'])
            
            # --- CODE TẢI HỒ SƠ (GIỮ NGUYÊN NHƯ CŨ) ---
            student_info_tuple = lay_thong_tin_day_du_hoc_sinh_db(student_id=hs_id)
            if not student_info_tuple: return
                
            info = dict(zip(["id"] + STUDENT_FIELDS_DB, student_info_tuple))

            tab_tong_quan, tab_lich_su, tab_muc_tieu = st.tabs(["📝 Tổng quan & Biểu đồ", "📜 Lịch sử Rèn luyện", "🎯 Mục tiêu & Phản hồi"])

            # === TAB 1: TỔNG QUAN ===
            with tab_tong_quan:
                col_img, col_info = st.columns([1, 2], gap="medium")
                
                with col_img:
                    import os
                    if not os.path.exists('student_photos'): os.makedirs('student_photos')
                    anh_path = info.get('anh_the_path')
                    full_img_path = os.path.join('student_photos', anh_path) if anh_path else ""
                    
                    if anh_path and os.path.exists(full_img_path):
                        st.image(full_img_path, width=130)
                    else:
                        st.image("https://cdn-icons-png.flaticon.com/512/149/149071.png", width=130)
                    
                    uploaded_file = st.file_uploader("Đổi ảnh", type=['jpg', 'jpeg', 'png'], key=f"up_{hs_id}", label_visibility="collapsed")
                    if uploaded_file is not None:
                        if st.button("Lưu ảnh", type="primary", key=f"btn_{hs_id}", width="stretch"):
                            ext = uploaded_file.name.split('.')[-1]
                            new_filename = f"hs_{hs_id}.{ext}"
                            with open(os.path.join('student_photos', new_filename), "wb") as f:
                                f.write(uploaded_file.getbuffer())
                            if cap_nhat_anh_the_db(hs_id, new_filename):
                                st.toast("Cập nhật ảnh thành công!", icon="✅")
                                st.rerun()
                
                with col_info:
                    st.markdown(f"#### {info.get('ten', '')}")
                    st.write(f"**Lớp:** {info.get('lop', '')} | **Tổ:** {info.get('to_nhiem_vu', '')}")
                    st.write(f"**Ngày sinh:** {info.get('ngay_sinh', '')}")
                    st.write(f"**Phụ huynh:** {info.get('ten_cha', '')} / {info.get('ten_me', '')}")
                    st.write(f"**SĐT LH:** {info.get('sdt_cha', '')} / {info.get('sdt_me', '')}")
                    
                st.markdown("---")
                st.write("**📈 Xu hướng rèn luyện (4 tuần)**")
                chart_data = lay_du_lieu_bieu_do_ca_nhan(hs_id, num_weeks=4)
                if chart_data:
                    weeks = [f"Tuần {int(d[0])}" for d in chart_data]
                    scores = [DIEM_KHOI_DAU + d[1] for d in chart_data]
                    st.line_chart(pd.DataFrame({"Tổng điểm": scores}, index=weeks), color="#0078D7")
                else:
                    st.info("Chưa có dữ liệu rèn luyện.")

            # === TAB LỊCH SỬ ===
            with tab_lich_su:
                st.write("**Lịch sử Kỷ luật (Theo TT19)**")
                history_kl = lay_lich_su_ky_luat_cua_hoc_sinh_db(hs_id)
                if history_kl:
                    df_kl = pd.DataFrame(history_kl, columns=["ID", "Hình thức", "Ngày", "Ghi chú"])
                    df_kl['Ngày'] = pd.to_datetime(df_kl['Ngày']).dt.strftime('%d/%m/%Y')
                    st.dataframe(df_kl[["Ngày", "Hình thức", "Ghi chú"]], width="stretch", hide_index=True)
                else: st.success("Chưa bị kỷ luật.")
                
                st.write("**Chi tiết Sự kiện (+/- điểm)**")
                all_events = lay_su_kien_trong_khoang_ngay_db('2020-01-01', '2100-01-01', user['role'], user['class'], user['group'])
                hs_events = [e for e in all_events if e[0] == hs_id]
                if hs_events:
                    df_events = pd.DataFrame(hs_events, columns=["ID", "Tên", "Lớp", "Tổ", "Nội dung", "Loại", "Điểm", "Ngày"])
                    df_events['Ngày'] = pd.to_datetime(df_events['Ngày']).dt.strftime('%d/%m/%Y')
                    st.dataframe(df_events[["Ngày", "Nội dung", "Loại", "Điểm"]].sort_values("Ngày", ascending=False), width="stretch", hide_index=True)
                else: st.info("Chưa có sự kiện nào.")

            # === TAB MỤC TIÊU ===
            with tab_muc_tieu:
                muc_tieu_cu = info.get('muc_tieu_thang') or ""
                phan_hoi_cu = info.get('phan_hoi_gvcn') or ""
                
                if st.button("✨ Trợ lý AI viết nhận xét", type="secondary", width="stretch", key=f"ai_{hs_id}"):
                    if "GEMINI_API_KEY" not in st.secrets: st.error("Chưa cấu hình API Key!")
                    else:
                        with st.spinner("🤖 AI đang soạn lời phê..."):
                            try:
                                from google import genai
                                import calendar
                                client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
                                now = datetime.now()
                                start_m = datetime(now.year, now.month, 1).date()
                                _, last_day = calendar.monthrange(now.year, now.month)
                                end_m = datetime(now.year, now.month, last_day).date()
                                sk_thang = [e for e in hs_events if start_m <= e[7].date() <= end_m]
                                diem_thang = DIEM_KHOI_DAU + sum(e[6] for e in sk_thang)
                                kt_str = ", ".join([e[4] for e in sk_thang if e[5] == "Khen thưởng"]) or "Chưa có nổi bật"
                                vp_str = ", ".join([e[4] for e in sk_thang if e[5] == "Vi phạm"]) or "Không vi phạm"
                                
                                prompt = f"Đóng vai là GVCN. Hãy viết 1 đoạn nhận xét cuối tháng (3-4 câu) cho học sinh {info.get('ten')}. Điểm rèn luyện: {diem_thang}/100. Ưu điểm: {kt_str}. Khuyết điểm: {vp_str}. Yêu cầu: Giọng chuẩn mực sư phạm, xưng 'thầy/cô' gọi 'em', động viên khen ngợi, nhắc nhở khéo léo. Trả về đúng đoạn văn."
                                response = client.models.generate_content(model='gemini-3.6-flash', contents=prompt)
                                st.session_state[f"ai_phan_hoi_{hs_id}"] = response.text
                                st.toast("AI đã viết xong!", icon="🎉")
                            except Exception as e: st.error(f"Lỗi AI: {e}")

                hien_thi_phan_hoi = st.session_state.get(f"ai_phan_hoi_{hs_id}", phan_hoi_cu)
                new_muc_tieu = st.text_area("🎯 Mục tiêu tháng tới:", value=muc_tieu_cu, height=80)
                new_phan_hoi = st.text_area("💬 Nhận xét & Phản hồi:", value=hien_thi_phan_hoi, height=120)
                
                if st.button("💾 Lưu Nhận xét", type="primary", width="stretch", key=f"save_{hs_id}"):
                    if luu_muc_tieu_phan_hoi_db(hs_id, new_muc_tieu, new_phan_hoi):
                        st.toast("Đã lưu thành công!", icon="✅")
                        if f"ai_phan_hoi_{hs_id}" in st.session_state: del st.session_state[f"ai_phan_hoi_{hs_id}"]
# --- HÀM 8: GHI NHẬN NHANH (TỐI ƯU MOBILE BẰNG GIAO DIỆN PILLS & CACHING) ---
def show_quick_record_page():
    st.header("📝 Ghi nhận Điểm Cộng / Trừ Nhanh")
    st.markdown("---")

    user = st.session_state.user_info
    
    # 1. Dùng bộ nhớ đệm (Cache) để tải học sinh và sự kiện SIÊU NHANH
    raw_students = get_cached_students(user['role'], user['class'], user['group'])
    if not raw_students:
        st.warning("Không có học sinh nào trong phạm vi quản lý.")
        return

    student_dict = {f"{hs[1]} (Tổ {hs[3] or '?'})": hs[0] for hs in raw_students}
    
    all_events = get_cached_events()
    event_dict = {}
    for ev in all_events:
        ten, loai, diem = ev[1], ev[2], ev[3]
        diem_str = f"+{diem}đ" if diem > 0 else f"{diem}đ"
        # Thêm Icon 🟢/🔴 cho rực rỡ và dễ phân biệt
        icon = "🟢" if loai == "Khen thưởng" else "🔴"
        label = f"{icon} {ten} ({diem_str})"
        event_dict[label] = (ten, loai, diem)

    # Đã thêm key mới cho ngày tháng
    event_date = st.date_input("📅 Chọn Ngày ghi nhận:", format="DD/MM/YYYY", key="date_quick_v2")
    ngay_tao_str = event_date.strftime('%Y-%m-%d %H:%M:%S')
    st.markdown("---")

    # 2. Tạo 2 Tab 
    tab1, tab2 = st.tabs(["👥 Cả nhóm cùng chung 1 Lỗi/Khen", "👤 1 Học sinh mắc nhiều Lỗi/Khen"])

    # ==========================================
    # KỊCH BẢN 1: 1 SỰ KIỆN ÁP DỤNG CHO NHIỀU HS
    # ==========================================
    with tab1:
        st.info("💡 Chạm chọn 1 Sự kiện, sau đó chọn danh sách học sinh bên dưới.")
        
        # SỬ DỤNG GIAO DIỆN PILLS (Đã thêm key mới)
        selected_event_1 = st.pills("1. Chạm chọn 1 Sự kiện:", options=list(event_dict.keys()), selection_mode="single", key="pill_tab1_v2")
        selected_students_1 = st.multiselect("2. Chọn các Học sinh:", options=list(student_dict.keys()), key="ms_tab1_v2")
            
        # Đã đổi key thành btn_tab1_v2
        if st.button("🚀 Ghi nhận cho danh sách trên", type="primary", width="stretch", key="btn_tab1_v2"):
            if selected_event_1 and selected_students_1: 
                ten_sk, loai_sk, diem_sk = event_dict[selected_event_1]
                count = 0
                for hs_name in selected_students_1:
                    hs_id = student_dict[hs_name]
                    if them_su_kien_ren_luyen_db(hs_id, ten_sk, loai_sk, diem_sk, ngay_tao_str):
                        count += 1
                st.success(f"✅ Đã ghi nhận **{ten_sk}** cho **{count}** học sinh thành công!")
            else:
                st.error("👆 Vui lòng chạm chọn 1 sự kiện ở trên và chọn ít nhất 1 học sinh.")

    # ==========================================
    # KỊCH BẢN 2: NHIỀU SỰ KIỆN CHO 1 HS
    # ==========================================
    with tab2:
        st.info("💡 Chọn 1 Học sinh, sau đó chạm để chọn nhiều Sự kiện cùng lúc.")
        
        # Đã đổi key thành sb_tab2_v2
        selected_student_2 = st.selectbox("1. Chọn Học sinh:", options=["-- Chọn --"] + list(student_dict.keys()), key="sb_tab2_v2")
        
        # SỬ DỤNG GIAO DIỆN PILLS CHỌN NHIỀU (Đã thêm key mới)
        selected_events_2 = st.pills("2. Chạm chọn các Sự kiện (được chọn nhiều):", options=list(event_dict.keys()), selection_mode="multi", key="pill_tab2_v2")

        # Đã đổi key thành btn_tab2_v2
        if st.button("🚀 Ghi nhận các sự kiện trên", type="primary", width="stretch", key="btn_tab2_v2"):
            if selected_student_2 != "-- Chọn --" and selected_events_2:
                hs_id = student_dict[selected_student_2]
                count = 0
                for ev_label in selected_events_2:
                    ten_sk, loai_sk, diem_sk = event_dict[ev_label]
                    if them_su_kien_ren_luyen_db(hs_id, ten_sk, loai_sk, diem_sk, ngay_tao_str):
                        count += 1
                
                hs_ten_ngan = selected_student_2.split("(")[0].strip()
                st.success(f"✅ Đã ghi nhận **{count}** sự kiện cho học sinh **{hs_ten_ngan}** thành công!")
            else:
                st.error("👆 Vui lòng chọn 1 học sinh và chạm chọn ít nhất 1 sự kiện.")
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
# --- HÀM 6: TỔNG KẾT & XUẤT BÁO CÁO (NÂNG CẤP FULL GIAO DIỆN) ---
def show_summary_page():
    st.header("📑 Báo cáo Tổng kết Toàn diện")
    st.markdown("---")
    
    user = st.session_state.user_info
    
    tab_tuan, tab_thang, tab_pdf = st.tabs(["📅 Tổng kết Tuần", "🗓️ Tổng kết Tháng", "🖨️ Xuất Phiếu PDF"])
    
    # ==========================================
    # TAB 1: TỔNG KẾT TUẦN
    # ==========================================
    with tab_tuan:
        col_input1, col_input2, col_input3 = st.columns([1, 1, 2])
        with col_input1:
            week_num = st.number_input("Chọn tuần số:", min_value=1, max_value=52, value=1, step=1)
        with col_input2:
            st.write("")
            st.write("")
            btn_tuan = st.button("📊 Truy xuất Báo cáo", type="primary", width="stretch")
            
        if btn_tuan or st.session_state.get('show_tuan', False):
            st.session_state.show_tuan = True 
            
            start_date, end_date = get_week_dates_by_number(week_num)
            if not start_date:
                st.error("Không thể xác định ngày của tuần này. Hãy kiểm tra Cài đặt năm học.")
            else:
                end_date_inc = end_date + timedelta(days=1)
                st.success(f"**ĐANG XEM BÁO CÁO TUẦN {week_num}:** Từ {start_date.strftime('%d/%m/%Y')} đến {end_date.strftime('%d/%m/%Y')}")
                
                with st.spinner("Đang tổng hợp dữ liệu toàn diện..."):
                    all_events = lay_su_kien_trong_khoang_ngay_db(start_date.strftime('%Y-%m-%d'), end_date_inc.strftime('%Y-%m-%d'), user['role'], user['class'], user['group'])
                    all_students = lay_thong_tin_day_du_hoc_sinh_db(user_role=user['role'], assigned_class=user['class'], assigned_group=user['group'])
                    
                    events_by_student = {}
                    for ev in all_events:
                        events_by_student.setdefault(ev[0], []).append(ev)
                        
                    student_list = []
                    for hs in all_students:
                        hs_id, ten_hs, lop_hs, to_hs = hs[0], hs[1], hs[2], hs[3] or ""
                        hs_events = events_by_student.get(hs_id, [])
                        score = DIEM_KHOI_DAU + sum(e[6] for e in hs_events)
                        final_score = max(0, score)
                        student_list.append({
                            'id': hs_id, 'ten': ten_hs, 'lop': lop_hs, 'to': to_hs, 
                            'diem': final_score, 'xeploai': xep_loai_hanh_kiem(final_score)
                        })
                    student_list.sort(key=lambda x: (-x['diem'], x['to']))
                    df_tuan = pd.DataFrame(student_list)
                    if not df_tuan.empty: df_tuan.insert(0, 'STT', range(1, len(df_tuan) + 1))

                # --- 1. HIỂN THỊ BẢNG XẾP HẠNG TỔ ---
                st.markdown("#### 🏆 Xếp hạng Thi đua Tổ")
                if not df_tuan.empty and 'to' in df_tuan.columns:
                    # Tính điểm trung bình theo tổ
                    df_team = df_tuan[df_tuan['to'] != ""].groupby('to')['diem'].mean().reset_index()
                    df_team.rename(columns={'to': 'Tên Tổ', 'diem': 'Điểm Trung Bình'}, inplace=True)
                    df_team.sort_values(by='Điểm Trung Bình', ascending=False, inplace=True)
                    df_team.insert(0, 'Hạng', range(1, len(df_team) + 1))
                    
                    # Vẽ bảng ngang siêu đẹp
                    cols_team = st.columns(len(df_team) if len(df_team) > 0 else 1)
                    for i, (_, row) in enumerate(df_team.iterrows()):
                        with cols_team[i]:
                            bg_color = "#e6f9ec" if i == 0 else "#f8f9fa"
                            icon = "🥇" if i == 0 else "🥈" if i == 1 else "🥉" if i == 2 else "🏅"
                            st.markdown(f"<div style='background-color:{bg_color}; padding:15px; border-radius:10px; text-align:center; border: 1px solid #ddd;'>"
                                        f"<h4>{icon} {row['Tên Tổ']}</h4>"
                                        f"<h2 style='color:#0078D7; margin:0;'>{row['Điểm Trung Bình']:.1f}</h2>"
                                        f"</div>", unsafe_allow_html=True)
                
                st.write("") # Cắt dòng
                
                # Chia 2 cột cho Vi phạm và Danh sách tổng
                col_vipham, col_tong = st.columns([1, 1.5], gap="large")

                # --- 2. CHI TIẾT LỖI VI PHẠM (BÊN TRÁI) ---
                with col_vipham:
                    st.markdown("#### ⚠️ Tổng hợp Vi phạm")
                    vipham_list = [{"Họ Tên": ev[1], "Tổ": ev[3] or "", "Lỗi Vi Phạm": ev[4]} for ev in all_events if ev[5] == "Vi phạm"]
                    
                    if vipham_list:
                        df_vp = pd.DataFrame(vipham_list)
                        # Đếm số lần vi phạm của mỗi HS cho từng lỗi
                        df_vp_grouped = df_vp.groupby(['Họ Tên', 'Tổ', 'Lỗi Vi Phạm']).size().reset_index(name='Số lần')
                        df_vp_grouped = df_vp_grouped.sort_values(by=['Tổ', 'Họ Tên'])
                        st.dataframe(df_vp_grouped, width="stretch", hide_index=True)
                    else:
                        st.success("Tuyệt vời! Tuần này không có học sinh nào vi phạm nội quy.")

                # --- 3. BẢNG TỔNG KẾT ĐIỂM (BÊN PHẢI) ---
                with col_tong:
                    st.markdown("#### 📋 Bảng Điểm & Xếp loại Lớp")
                    if not df_tuan.empty:
                        # Dùng data_editor để hiển thị đẹp hơn
                        st.dataframe(
                            df_tuan[['STT', 'ten', 'to', 'diem', 'xeploai']].rename(columns={'ten': 'Họ và Tên', 'to': 'Tổ', 'diem': 'Tổng điểm', 'xeploai': 'Xếp loại'}),
                            width="stretch", 
                            hide_index=True,
                            height=400
                        )

                # --- 4. TẢI FILE EXCEL & ZALO ---
                st.markdown("---")
                col_btn1, col_btn2 = st.columns(2)
                
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    filepath = tmp.name
                    
                success, msg = generate_weekly_summary_excel(
                    student_list, week_num, user['full_name'], user['class'] or "Toàn trường", user['group'] or "", 
                    start_date, end_date, filepath, [], [], [] 
                )
                
                with col_btn1:
                    if success:
                        with open(filepath, "rb") as f: excel_data = f.read()
                        st.download_button("📥 TẢI FILE EXCEL BÁO CÁO TUẦN (BẢN ĐẦY ĐỦ)", data=excel_data, file_name=f"TongKet_Tuan_{week_num}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", width="stretch")
                    os.unlink(filepath)
                    
                with col_btn2:
                    st.info("👆 File Excel đính kèm chứa 5 sheet chi tiết giống hệt bản Desktop (Xếp hạng, Vi phạm, Khen thưởng...).")

                # (Giữ nguyên Trợ lý Zalo ở cuối)
                with st.expander("💬 Mở Trợ lý tạo tin nhắn Zalo gửi Phụ huynh"):
                    top_students = [hs['ten'] for hs in student_list[:5] if hs['diem'] > DIEM_KHOI_DAU]
                    if not top_students: top_students = [hs['ten'] for hs in student_list[:3]] 
                    vipham_dict = {}
                    for ev in all_events:
                        if ev[5] == 'Vi phạm': vipham_dict.setdefault(ev[1], []).append(ev[4])
                    vipham_text_list = [f"{ten} ({', '.join(list(set(loi)))})" for ten, loi in vipham_dict.items()]

                    msg_zalo = f"🌟 KÍNH GỬI QUÝ PHỤ HUYNH LỚP {user['class'] or '...'} - TỔNG KẾT TUẦN {week_num} 🌟\n\n"
                    if top_students: msg_zalo += f"🏆 Khen ngợi các em tích cực: {', '.join(top_students)}.\n\n"
                    if vipham_text_list: msg_zalo += f"⚠️ Nhắc nhở các em còn vi phạm:\n- {';\n- '.join(vipham_text_list)}.\n\nKính mong quý phụ huynh nhắc nhở thêm các em.\n\n"
                    else: msg_zalo += f"✨ Tuyệt vời! Tuần này lớp chúng ta thực hiện nề nếp rất tốt, không có bạn nào vi phạm nội quy.\n\n"
                    msg_zalo += "📌 Chi tiết quý vị xem trong file Excel đính kèm. Trân trọng!"
                    st.text_area("Copy đoạn văn bản này dán vào nhóm Zalo:", value=msg_zalo, height=200)

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
            btn_thang = st.button("📊 Xem Báo cáo Tháng", type="primary", width="stretch")
            
        if btn_thang or st.session_state.get('show_thang', False):
            st.session_state.show_thang = True
            start_date_month = datetime(year, month, 1)
            _, num_days = calendar.monthrange(year, month)
            end_date_month = datetime(year, month, num_days)
            
            st.success(f"**BÁO CÁO THÁNG {month}/{year}:** Từ {start_date_month.strftime('%d/%m/%Y')} đến {end_date_month.strftime('%d/%m/%Y')}")
            
            with st.spinner("Đang tính toán điểm tháng..."):
                all_events_m = lay_su_kien_trong_khoang_ngay_db(start_date_month.strftime('%Y-%m-%d'), end_date_month.strftime('%Y-%m-%d'), user['role'], user['class'], user['group'])
                all_students_m = lay_thong_tin_day_du_hoc_sinh_db(user_role=user['role'], assigned_class=user['class'], assigned_group=user['group'])
                
                events_by_student_m = {}
                for ev in all_events_m: events_by_student_m.setdefault(ev[0], []).append(ev)
                    
                student_list_m = []
                for hs in all_students_m:
                    hs_events = events_by_student_m.get(hs[0], [])
                    score = DIEM_KHOI_DAU + sum(e[6] for e in hs_events)
                    student_list_m.append({'id': hs[0], 'ten': hs[1], 'lop': hs[2], 'to': hs[3] or "", 'diem': score, 'xeploai': xep_loai_hanh_kiem(score)})
                student_list_m.sort(key=lambda x: (-x['diem'], x['to']))
            
            df_thang = pd.DataFrame(student_list_m)
            if not df_thang.empty: df_thang.insert(0, 'STT', range(1, len(df_thang) + 1))
            
            # --- HIỂN THỊ BẢNG XẾP HẠNG TỔ THÁNG ---
            st.markdown("#### 🏆 Xếp hạng Thi đua Tổ (Tháng)")
            if not df_thang.empty and 'to' in df_thang.columns:
                df_team_m = df_thang[df_thang['to'] != ""].groupby('to')['diem'].mean().reset_index()
                df_team_m.rename(columns={'to': 'Tên Tổ', 'diem': 'Điểm Trung Bình'}, inplace=True)
                df_team_m.sort_values(by='Điểm Trung Bình', ascending=False, inplace=True)
                
                cols_team_m = st.columns(len(df_team_m) if len(df_team_m) > 0 else 1)
                for i, (_, row) in enumerate(df_team_m.iterrows()):
                    with cols_team_m[i]:
                        bg_color = "#e6f9ec" if i == 0 else "#f8f9fa"
                        st.markdown(f"<div style='background-color:{bg_color}; padding:15px; border-radius:10px; text-align:center; border: 1px solid #ddd;'>"
                                    f"<h4>Tổ {row['Tên Tổ']}</h4>"
                                    f"<h2 style='color:#0078D7; margin:0;'>{row['Điểm Trung Bình']:.1f}</h2>"
                                    f"</div>", unsafe_allow_html=True)
            
            st.markdown("#### 📋 Bảng Xếp loại Tháng của Lớp")
            st.dataframe(df_thang[['STT', 'ten', 'to', 'diem', 'xeploai']].rename(columns={'ten': 'Họ Tên', 'to': 'Tổ', 'diem': 'Điểm', 'xeploai': 'Xếp loại'}), width="stretch", hide_index=True)
            
            st.markdown("---")
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_m: filepath_m = tmp_m.name
            success_m, msg_m = generate_monthly_summary_excel(student_list_m, month, year, user['full_name'], user['class'] or "Toàn trường", filepath_m)
            
            if success_m:
                with open(filepath_m, "rb") as f_m: excel_data_m = f_m.read()
                st.download_button("📥 TẢI FILE EXCEL BÁO CÁO THÁNG", data=excel_data_m, file_name=f"TongKet_Thang_{month}_{year}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", width="stretch")
            os.unlink(filepath_m)

    
    # ==========================================
    # ==========================================
    # TAB 3: XUẤT PHIẾU LIÊN LẠC PDF (ĐÃ SỬA LỖI TRÙNG KEY)
    # ==========================================
    with tab_pdf:
        st.subheader("Tạo và Tải xuống Phiếu Liên Lạc (PDF)")
        
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            # Đổi key thành pdf_month_v3
            pdf_month = st.selectbox("Chọn tháng xuất:", range(1, 13), index=datetime.now().month - 1, key="pdf_month_v3")
        with col_p2:
            # Đổi key thành pdf_year_v3
            pdf_year = st.number_input("Chọn năm xuất:", value=datetime.now().year, key="pdf_year_v3")
            
        st.markdown("---")
        # Đổi key thành export_scope_v3
        export_scope = st.radio("Phạm vi xuất:", ["Toàn bộ lớp", "Chọn học sinh cụ thể"], key="export_scope_v3")
        
        # Lấy danh sách HS để chọn
        all_students = lay_thong_tin_day_du_hoc_sinh_db(user_role=user['role'], assigned_class=user['class'], assigned_group=user['group'])
        student_dict = {f"{hs[1]} (Tổ {hs[3] or '?'})": hs[0] for hs in all_students}
        
        selected_student_ids = []
        if export_scope == "Chọn học sinh cụ thể":
            # Đổi key thành ms_export_hs_v3
            selected_names = st.multiselect("Chọn học sinh muốn xuất:", options=list(student_dict.keys()), key="ms_export_hs_v3")
            selected_student_ids = [student_dict[name] for name in selected_names]
        else:
            selected_student_ids = [hs[0] for hs in all_students]
            st.info(f"Sẽ xuất phiếu cho toàn bộ {len(selected_student_ids)} học sinh.")
            
        # Đổi key thành btn_export_pdf_v3
        if st.button("🚀 Xử lý & Tạo Phiếu Liên Lạc", type="primary", width="stretch", key="btn_export_pdf_v3"):
            if not selected_student_ids:
                st.warning("Vui lòng chọn ít nhất một học sinh.")
            else:
                with st.spinner(f"Đang tự động tạo {len(selected_student_ids)} phiếu liên lạc PDF... Thầy chờ chút nhé!"):
                    # Tạo thư mục tạm trên server để lưu PDF
                    with tempfile.TemporaryDirectory() as tmpdirname:
                        pdf_files = []
                        success_count, fail_count = 0, 0
                        
                        # Duyệt qua từng HS để tạo PDF
                        for hs in all_students:
                            if hs[0] in selected_student_ids:
                                # Làm sạch tên file để tránh lỗi tiếng Việt có dấu
                                safe_name = "".join(x for x in hs[1] if x.isalnum() or x in " _-").replace(" ", "_")
                                filename = f"PhieuLienLac_{safe_name}_T{pdf_month}-{pdf_year}.pdf"
                                filepath = os.path.join(tmpdirname, filename)
                                
                                # Gọi hàm tạo PDF từ code cũ
                                success, msg = generate_pdf_report(hs[0], pdf_year, pdf_month, user['full_name'], filepath)
                                if success:
                                    pdf_files.append((filename, filepath))
                                    success_count += 1
                                else:
                                    fail_count += 1
                                    
                        # Xử lý sau khi tạo xong
                        if success_count > 0:
                            st.success(f"✅ Đã tạo thành công {success_count} phiếu PDF. Lỗi: {fail_count}.")
                            
                            # Nếu chỉ xuất 1 người, cho tải thẳng file PDF
                            if len(pdf_files) == 1:
                                with open(pdf_files[0][1], "rb") as f:
                                    pdf_data = f.read()
                                st.download_button(
                                    label=f"📥 Tải xuống Phiếu: {pdf_files[0][0]}",
                                    data=pdf_data,
                                    file_name=pdf_files[0][0],
                                    mime="application/pdf",
                                    type="primary",
                                    width="stretch",
                                    key="btn_download_single_v3"
                                )
                            # Nếu xuất nhiều người, Nén Zip lại cho tiện
                            else:
                                zip_buffer = io.BytesIO()
                                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                                    for fname, fpath in pdf_files:
                                        zip_file.write(fpath, arcname=fname)
                                
                                st.download_button(
                                    label=f"📦 Tải xuống File ZIP chứa {success_count} Phiếu Liên Lạc",
                                    data=zip_buffer.getvalue(),
                                    file_name=f"PhieuLienLac_T{pdf_month}_{pdf_year}.zip",
                                    mime="application/zip",
                                    type="primary",
                                    width="stretch",
                                    key="btn_download_zip_v3"
                                )
                        else:
                            st.error("Không thể tạo phiếu liên lạc nào. Vui lòng kiểm tra lại dữ liệu và Font chữ.")
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
    # TAB 2: QUẢN LÝ DANH MỤC SỰ KIỆN (TT19) - ĐÃ NÂNG CẤP EXCEL & XÓA NHIỀU
    # ==========================================
    with tab_events:
        st.info("Quản lý danh mục các lỗi vi phạm và thành tích khen thưởng để áp dụng toàn trường.")

        # --- 1. KHU VỰC NHẬP BẰNG EXCEL ---
        with st.expander("📥 Bấm vào đây để TẠO/CẬP NHẬT HÀNG LOẠT bằng file Excel", expanded=False):
            col_ex1, col_ex2 = st.columns([1, 1])
            
            with col_ex1:
                st.write("1. Tải file mẫu về máy và điền danh sách sự kiện.")
                # Tạo file Excel mẫu trên RAM
                df_mau_sk = pd.DataFrame({
                    "Tên sự kiện": ["Đi học trễ", "Không đồng phục", "Nhặt được của rơi"],
                    "Loại": ["Vi phạm", "Vi phạm", "Khen thưởng"],
                    "Điểm": [-2, -5, 10],
                    "Mức độ vi phạm": [1, 2, None]
                })
                output_sk = io.BytesIO()
                with pd.ExcelWriter(output_sk, engine='openpyxl') as writer:
                    df_mau_sk.to_excel(writer, index=False)
                
                st.download_button(
                    label="⬇️ Tải File Excel Mẫu",
                    data=output_sk.getvalue(),
                    file_name="Mau_Nhap_Su_Kien.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_dl_sk_template"
                )
                st.caption("Lưu ý: Cột 'Loại' phải gõ đúng chữ 'Vi phạm' hoặc 'Khen thưởng'.")

            with col_ex2:
                st.write("2. Tải file đã điền lên hệ thống.")
                uploaded_sk = st.file_uploader("Chọn file Excel sự kiện:", type=['xlsx', 'xls'], key="up_sk_excel")
                
                if uploaded_sk is not None:
                    if st.button("🚀 Xử lý & Nhập dữ liệu", type="primary", key="btn_import_sk"):
                        try:
                            df_imp_sk = pd.read_excel(uploaded_sk)
                            success_sk, fail_sk = 0, 0
                            
                            with st.spinner("Đang lưu vào hệ thống..."):
                                for _, row in df_imp_sk.iterrows():
                                    ten = str(row.get("Tên sự kiện", "")).strip()
                                    loai = str(row.get("Loại", "")).strip()
                                    diem_val = row.get("Điểm")
                                    md_val = row.get("Mức độ vi phạm")
                                    
                                    if not ten or ten == "nan" or loai not in ["Vi phạm", "Khen thưởng"]:
                                        fail_sk += 1; continue
                                        
                                    try: diem = int(diem_val)
                                    except: fail_sk += 1; continue
                                    
                                    # Xử lý mức độ
                                    muc_do = None
                                    if loai == "Vi phạm" and pd.notna(md_val):
                                        try: muc_do = int(md_val)
                                        except: pass
                                        
                                    # Hàm này có tính năng UPSERT: Đã có tên thì cập nhật, chưa có thì thêm mới
                                    if them_hoac_cap_nhat_danh_muc_db(ten, loai, diem, muc_do):
                                        success_sk += 1
                                    else: fail_sk += 1
                                    
                            st.success(f"✅ Nhập hoàn tất: Thành công {success_sk}. Lỗi/Bỏ qua {fail_sk}.")
                            st.rerun() # Tải lại trang để hiện list mới
                        except Exception as e:
                            st.error(f"Lỗi file: {e}")

        st.markdown("---")

        # --- 2. KHU VỰC DANH SÁCH (CÓ CHECKBOX XÓA) & FORM SỬA TAY ---
        col_elist, col_eform = st.columns([1.5, 1], gap="large")
        
        with col_elist:
            st.subheader("📋 Danh sách Sự kiện (Xóa hàng loạt)")
            events = lay_tat_ca_danh_muc_su_kien_db()
            if events:
                df_events = pd.DataFrame(events, columns=["ID", "Tên Sự kiện", "Loại", "Điểm", "Mức độ"])
                # Thêm cột Checkbox để tick chọn
                df_events.insert(0, "Chọn xóa", False)
                
                st.write("Đánh dấu tick [x] vào các sự kiện muốn xóa:")
                edited_ev_df = st.data_editor(
                    df_events,
                    hide_index=True,
                    column_config={
                        "Chọn xóa": st.column_config.CheckboxColumn(required=True),
                        "ID": None # Ẩn ID cho đẹp
                    },
                    disabled=["Tên Sự kiện", "Loại", "Điểm", "Mức độ"], # Không cho sửa chữ, chỉ cho tick ô vuông
                    use_container_width=True,
                    key="editor_events_v2"
                )
                
                # Lọc ra ID của các dòng được tick True
                selected_ev_rows = edited_ev_df[edited_ev_df["Chọn xóa"] == True]
                ids_to_del_ev = selected_ev_rows["ID"].tolist()
                
                if len(ids_to_del_ev) > 0:
                    st.warning(f"Bạn đang chọn xóa {len(ids_to_del_ev)} sự kiện.")
                    if st.button("🗑️ Xác nhận Xóa", type="primary", key="btn_del_events_v2"):
                        del_ev_count = xoa_nhieu_danh_muc_su_kien_db(ids_to_del_ev)
                        st.success(f"Đã xóa thành công {del_ev_count} sự kiện.")
                        st.rerun()

        with col_eform:
            st.subheader("Thêm/Sửa 1 Sự kiện")
            st.info("💡 Mẹo: Nhập tên sự kiện ĐÃ CÓ để Cập nhật lại điểm/mức độ.")
            with st.form("form_event_v2"):
                e_name = st.text_input("Tên sự kiện *")
                e_type = st.selectbox("Loại", ["Vi phạm", "Khen thưởng"])
                e_points = st.number_input("Điểm (+ hoặc -)", value=0)
                e_level = st.selectbox("Mức độ vi phạm theo TT19", ["Không xét KL", "1", "2", "3"])

                if st.form_submit_button("💾 Lưu Sự kiện", type="primary", use_container_width=True):
                    if e_name:
                        muc_do = int(e_level) if e_level != "Không xét KL" and e_type == "Vi phạm" else None
                        if them_hoac_cap_nhat_danh_muc_db(e_name, e_type, e_points, muc_do):
                            st.success(f"Đã lưu sự kiện **{e_name}**!")
                            st.rerun()
                    else:
                        st.error("Vui lòng nhập tên sự kiện.")

    # ==========================================
    # TAB 3: CÀI ĐẶT CHUNG
    # ==========================================
    # ==========================================
    # TAB 3: CÀI ĐẶT NĂM HỌC & KỲ NGHỈ LỄ
    # ==========================================
    with tab_settings:
        st.subheader("1. Cài đặt Ngày bắt đầu Năm học")
        st.write("Ngày này được dùng làm mốc (Tuần 1) để tính số thứ tự Tuần học trong các Báo cáo.")
        
        current_start = load_setting("school_year_start_date", "2025-09-08")
        start_date = st.date_input("Ngày bắt đầu năm học:", datetime.strptime(current_start, "%Y-%m-%d").date())
        
        if st.button("💾 Lưu Ngày bắt đầu", type="primary", width="stretch", key="btn_save_start"):
            save_setting("school_year_start_date", start_date.strftime("%Y-%m-%d"))
            st.toast("Đã lưu ngày bắt đầu năm học thành công!", icon="✅")

        st.markdown("---")
        st.subheader("2. Quản lý Kỳ nghỉ (Lễ, Tết)")
        st.write("Các khoảng thời gian này sẽ KHÔNG được tính vào số tuần học thực tế.")
        st.info("💡 **Hướng dẫn:** Bấm vào lịch trong bảng để chọn ngày. Bấm dấu **[+]** ở góc dưới bảng để thêm kỳ nghỉ mới. Đánh dấu ô vuông bên trái dòng và bấm phím **Delete** (hoặc biểu tượng thùng rác) để xóa.")
        
        # Đọc dữ liệu kỳ nghỉ từ CSDL
        import json
        holidays_json = load_setting(HOLIDAY_SETTINGS_KEY, '[]')
        try:
            holidays_list = json.loads(holidays_json)
            # Chuyển chuỗi thành định dạng ngày tháng để hiển thị lịch trên Web
            for i in range(len(holidays_list)):
                holidays_list[i][0] = datetime.strptime(holidays_list[i][0], "%Y-%m-%d").date()
                holidays_list[i][1] = datetime.strptime(holidays_list[i][1], "%Y-%m-%d").date()
            df_holidays = pd.DataFrame(holidays_list, columns=["Từ ngày", "Đến ngày"])
        except:
            df_holidays = pd.DataFrame(columns=["Từ ngày", "Đến ngày"])

        # Bảng dữ liệu tương tác (Cho phép Thêm/Sửa/Xóa trực tiếp)
        edited_holidays = st.data_editor(
            df_holidays,
            column_config={
                "Từ ngày": st.column_config.DateColumn("Từ ngày", required=True, format="DD/MM/YYYY"),
                "Đến ngày": st.column_config.DateColumn("Đến ngày", required=True, format="DD/MM/YYYY")
            },
            num_rows="dynamic", # Cho phép thêm/xóa dòng
            use_container_width=True,
            key="holiday_editor"
        )
        
        if st.button("💾 Lưu Danh sách Kỳ nghỉ", type="primary", width="stretch", key="btn_save_holidays"):
            valid_holidays = []
            has_error = False
            
            # Kiểm tra và định dạng lại dữ liệu trước khi lưu
            for index, row in edited_holidays.iterrows():
                start_d = row["Từ ngày"]
                end_d = row["Đến ngày"]
                
                if pd.notna(start_d) and pd.notna(end_d):
                    if start_d > end_d:
                        st.error(f"Lỗi ở dòng {index + 1}: Ngày bắt đầu ({start_d.strftime('%d/%m/%Y')}) không thể lớn hơn ngày kết thúc ({end_d.strftime('%d/%m/%Y')}).")
                        has_error = True
                        break
                    # Lưu lại định dạng chuỗi YYYY-MM-DD vào CSDL
                    valid_holidays.append([start_d.strftime('%Y-%m-%d'), end_d.strftime('%Y-%m-%d')])
            
            if not has_error:
                save_setting(HOLIDAY_SETTINGS_KEY, json.dumps(valid_holidays))
                st.success("✅ Đã lưu danh sách kỳ nghỉ thành công! Hệ thống sẽ tự động tính lại số Tuần trên toàn ứng dụng.")
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
# --- HÀM 9: ĐIỂM DANH HÀNG NGÀY ---
def show_attendance_page():
    st.header("📅 Điểm danh Hàng ngày")
    st.markdown("---")

    user = st.session_state.user_info
    role, aclass, agroup = user['role'], user['class'], user['group']

    # 1. Chọn ngày
    col_d1, col_d2 = st.columns([1, 3])
    with col_d1:
        attendance_date = st.date_input("Ngày điểm danh:", date.today(), format="DD/MM/YYYY")

    # 2. Lấy danh sách học sinh
    raw_students = lay_danh_sach_hoc_sinh_db(role, aclass, agroup)
    if not raw_students:
        st.warning("Không có học sinh nào trong phạm vi quản lý.")
        return

    # Chuẩn bị dữ liệu cho Bảng tương tác (Mặc định tất cả là "Có mặt")
    student_data = []
    for i, hs in enumerate(raw_students, 1):
        student_data.append({
            "ID": hs[0],
            "STT": i,
            "Họ và Tên": hs[1],
            "Tổ": hs[3] or "",
            "Trạng thái": "Có mặt",  # Mặc định
            "Ghi chú (Tùy chọn)": ""
        })

    df_attendance = pd.DataFrame(student_data)

    st.info("💡 **Hướng dẫn:** Nhấn đúp chuột vào cột **'Trạng thái'** để đổi sang Vắng/Trễ. Hệ thống sẽ tự động trừ điểm rèn luyện tương ứng khi bấm Lưu.")

    # 3. Hiển thị Bảng Điểm danh tương tác (Data Editor của Streamlit)
    edited_df = st.data_editor(
        df_attendance,
        hide_index=True,
        column_config={
            "ID": None, # Ẩn cột ID đi cho đẹp
            "STT": st.column_config.NumberColumn(disabled=True),
            "Họ và Tên": st.column_config.TextColumn(disabled=True),
            "Tổ": st.column_config.TextColumn(disabled=True),
            "Trạng thái": st.column_config.SelectboxColumn(
                options=["Có mặt", "Vắng có phép", "Vắng không phép", "Đi trễ"],
                required=True
            ),
            "Ghi chú (Tùy chọn)": st.column_config.TextColumn()
        },
        use_container_width=True
    )

    st.markdown("---")
    
    # 4. Xử lý khi bấm nút Lưu
    if st.button("💾 Lưu Điểm danh", type="primary", width="stretch"):
        # Lấy danh mục sự kiện hiện có trong CSDL để map điểm
        all_events = lay_tat_ca_danh_muc_su_kien_db()
        event_map = {e[1]: (e[2], e[3]) for e in all_events} # {Tên: (Loại, Điểm)}

        # Mức điểm dự phòng nếu thầy chưa tạo các sự kiện này trong mục Quản trị
        fallback_map = {
            "Vắng không phép": ("Vi phạm", -5),
            "Vắng có phép": ("Vi phạm", 0),  # Không trừ điểm
            "Đi trễ": ("Vi phạm", -2)
        }

        ngay_tao_str = attendance_date.strftime('%Y-%m-%d %H:%M:%S')
        success_count = 0
        
        with st.spinner("Đang lưu dữ liệu điểm danh..."):
            # Quét qua bảng dữ liệu vừa chỉnh sửa
            for index, row in edited_df.iterrows():
                status = row["Trạng thái"]
                
                # Chỉ ghi nhận vào CSDL nếu HS đó Vắng hoặc Trễ
                if status != "Có mặt":
                    hs_id = row["ID"]
                    ghi_chu = row["Ghi chú (Tùy chọn)"]
                    # Nối ghi chú vào tên sự kiện (nếu có)
                    mo_ta = f"{status}" + (f" ({ghi_chu})" if ghi_chu else "")
                    
                    # Tìm xem CSDL đã cài đặt điểm cho lỗi này chưa, chưa thì dùng dự phòng
                    if status in event_map:
                        loai_sk, diem_sk = event_map[status]
                    else:
                        loai_sk, diem_sk = fallback_map.get(status, ("Vi phạm", 0))
                        
                    # Lưu vào CSDL
                    if them_su_kien_ren_luyen_db(hs_id, mo_ta, loai_sk, diem_sk, ngay_tao_str):
                        success_count += 1

        if success_count > 0:
            st.success(f"✅ Đã lưu điểm danh! Có **{success_count}** học sinh (Vắng/Trễ) được ghi nhận và tự động cập nhật điểm rèn luyện.")
        else:
            st.success("✅ Đã lưu điểm danh! Hôm nay 100% học sinh có mặt đầy đủ.")
# --- HÀM 2: GIAO DIỆN CHÍNH (SAU KHI ĐĂNG NHẬP) ---
# --- HÀM 2: GIAO DIỆN CHÍNH (SAU KHI ĐĂNG NHẬP) ---
def show_main_dashboard():
    user = st.session_state.user_info
    role, aclass, agroup = user['role'], user['class'], user['group']
    
    # --- THANH BÊN (SIDEBAR) HIỆN ĐẠI ---
    with st.sidebar:
        st.title(f"Chào, {user['full_name']} 👋")
        st.caption(f"Vai trò: {role.upper()} | Lớp: {aclass or 'Toàn trường'}")
        st.markdown("---")
        
        # 1. Thiết lập danh sách Menu và Icon (Sử dụng Bootstrap Icons)
        options = ["Bảng điều khiển", "Bảng Vàng Thi Đua", "Quản lý Lớp học", "Điểm danh hàng ngày", "Ghi nhận Nhanh"]
        icons = ["house", "trophy", "people", "calendar2-check", "lightning-charge"]
        
        if role in ['gvcn', 'admin']:
            options.extend(["Ghi nhận Kỷ luật (TT19)", "Thống kê & Báo cáo", "Tổng kết & Xuất Excel"])
            icons.extend(["shield-exclamation", "bar-chart-steps", "file-earmark-spreadsheet"])
            
        if role == 'admin':
            options.append("Quản trị Hệ thống")
            icons.append("gear")

        # 2. Vẽ Menu với Style hiện đại (Theme Ocean Blue)
        choice = option_menu(
            menu_title=None,  # Ẩn tiêu đề chữ Menu
            options=options,
            icons=icons,
            menu_icon="cast",
            default_index=0,
            styles={
                "container": {"padding": "0!important", "background-color": "transparent"},
                "icon": {"font-size": "18px"}, 
                "nav-link": {
                    "font-size": "15px", "text-align": "left", "margin":"4px 0", 
                    "border-radius": "8px", "--hover-color": "#e6f2ff"
                },
                "nav-link-selected": {"background-color": "#0056b3", "color": "white"},
            }
        )
        
        st.markdown("---")
        if st.button("Đăng xuất", type="secondary", width="stretch"):
            st.session_state.logged_in = False
            st.session_state.user_info = None
            st.rerun()
# ====================================================================
    # <<< DÁN KHỐI CODE TỰ ĐỘNG ĐÓNG MENU (BẢN TỐI THƯỢNG) VÀO ĐÂY >>>
    # ====================================================================
    import streamlit.components.v1 as components
    
    if 'last_menu_choice' not in st.session_state:
        st.session_state.last_menu_choice = choice

    if choice != st.session_state.last_menu_choice:
        st.session_state.last_menu_choice = choice
        
        js_close_sidebar = """
        <script>
            // Đợi 0.5 giây để đảm bảo trình duyệt điện thoại đã vẽ xong giao diện
            setTimeout(function() {
                var parentDoc = window.parent.document;
                var parentWin = window.parent;
                
                // Chuẩn màn hình điện thoại/tablet của Streamlit là dưới 991px
                if (parentWin.innerWidth <= 991) {
                    
                    // Phương pháp 1: Ra lệnh bấm phím ESCAPE mạnh mẽ hơn
                    var escEvent = new KeyboardEvent('keydown', {
                        key: 'Escape', code: 'Escape', keyCode: 27, which: 27, bubbles: true
                    });
                    parentDoc.dispatchEvent(escEvent);
                    
                    // Phương pháp 2: Tìm và bấm thẳng vào nút "X" của thanh Menu
                    var sidebar = parentDoc.querySelector('[data-testid="stSidebar"]');
                    if (sidebar) {
                        var buttons = sidebar.querySelectorAll('button');
                        if (buttons.length > 0) {
                            buttons[0].click(); // Nút đầu tiên trên thanh menu luôn là nút X
                        }
                    }
                }
            }, 500);
        </script>
        """
        components.html(js_close_sidebar, height=0, width=0)
    # ====================================================================
    # --- KHU VỰC NỘI DUNG CHÍNH (ĐIỀU HƯỚNG THEO LỰA CHỌN) ---
    if choice == "Bảng điều khiển":
        st.header("📊 Bảng điều khiển Tổng quan")
        st.markdown("---")

        # ==========================================
        # TÍNH TOÁN DỮ LIỆU THỰC TẾ
        # ==========================================
        hs_list = lay_danh_sach_hoc_sinh_db(role, aclass, agroup)
        total_hs = len(hs_list) if hs_list else 0

        today = date.today()
        try:
            start_date_str = load_setting('school_year_start_date', '2025-09-08')
            h_json = json.loads(load_setting(HOLIDAY_SETTINGS_KEY, '[]'))
            h_in_year = [(datetime.strptime(s, '%Y-%m-%d').date(), datetime.strptime(e, '%Y-%m-%d').date()) for s, e in h_json]
            current_week = get_school_week_number(today, datetime.strptime(start_date_str, '%Y-%m-%d').date(), h_in_year)
            if current_week <= 0: current_week = 1
        except:
            current_week = 1
            
        start_w, end_w = get_week_dates_by_number(current_week)
        if not start_w: 
            start_w, end_w = today - timedelta(days=today.weekday()), today + timedelta(days=6-today.weekday())

        events_week = lay_su_kien_trong_khoang_ngay_db(start_w.strftime('%Y-%m-%d'), (end_w + timedelta(days=1)).strftime('%Y-%m-%d'), role, aclass, agroup)
        
        khen_thuong_count = sum(1 for ev in events_week if ev[5] == "Khen thưởng")
        vi_pham_count = sum(1 for ev in events_week if ev[5] == "Vi phạm")

        # ==========================================
        # HIỂN THỊ GIAO DIỆN (UI)
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
                for ev in events_week[:5]:
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
                st.write("Tuần này khá im ắng, chưa có hoạt động nào được ghi nhận.")

    # ==========================================
    # CÁC ĐOẠN MÃ ĐIỀU HƯỚNG
    # ==========================================
    elif choice == "Bảng Vàng Thi Đua":
        show_leaderboard_page()
    elif choice == "Quản lý Lớp học":
        show_class_management()
    elif choice == "Điểm danh hàng ngày":
        show_attendance_page()
    elif choice == "Ghi nhận Nhanh":
        show_quick_record_page()
    elif choice == "Ghi nhận Kỷ luật (TT19)":
        show_discipline_page()
    elif choice == "Thống kê & Báo cáo":
        show_statistics_page()
    elif choice == "Tổng kết & Xuất Excel":
        show_summary_page()
    elif choice == "Quản trị Hệ thống":
        show_admin_page()

# --- HÀM HỖ TRỢ: TẠO ẢNH BẰNG KHEN ---
def create_certificate_image(student_name, class_name, achievement_text):
    from PIL import Image, ImageDraw, ImageFont
    import io
    
    # Tạo một bức ảnh nền màu vàng nhạt sang trọng
    img = Image.new('RGB', (800, 600), color='#FDF5E6')
    draw = ImageDraw.Draw(img)

    # Vẽ khung viền vàng (Đậm ở ngoài, mảnh ở trong)
    draw.rectangle([20, 20, 780, 580], outline='#DAA520', width=12)
    draw.rectangle([35, 35, 765, 565], outline='#DAA520', width=3)

    # Tải Font chữ (Dự phòng lỗi nếu máy chủ mất font)
    try:
        font_title = ImageFont.truetype("fonts/timesbd.ttf", 55)
        font_subtitle = ImageFont.truetype("fonts/times.ttf", 30)
        font_name = ImageFont.truetype("fonts/timesbd.ttf", 65)
    except:
        font_title = font_subtitle = font_name = ImageFont.load_default()

    # Viết chữ lên Bằng khen
    draw.text((400, 100), "BẢNG VÀNG VINH DANH", fill="#B22222", font=font_title, anchor="mt")
    draw.text((400, 200), "Tuyên dương học sinh:", fill="#333333", font=font_subtitle, anchor="mt")
    draw.text((400, 280), str(student_name).upper(), fill="#000080", font=font_name, anchor="mt")
    draw.text((400, 380), f"Học sinh lớp: {class_name}", fill="#333333", font=font_subtitle, anchor="mt")
    draw.text((400, 450), str(achievement_text), fill="#D2691E", font=font_subtitle, anchor="mt")

    # Xuất ra định dạng Byte để Web tải về
    buf = io.BytesIO()
    img.save(buf, format='JPEG', quality=90)
    return buf.getvalue()


# --- HÀM 8: BẢNG VÀNG THI ĐUA ĐƯỢC NÂNG CẤP CÓ ẢNH ---
def show_leaderboard_page():
    import streamlit as st
    import pandas as pd
    import json
    import os
    import base64 # Khai báo trực tiếp trong hàm để chống lỗi
    from datetime import date, datetime, timedelta
    
    st.header("🏆 Bảng Vàng Thi Đua & Vinh Danh")
    st.markdown("---")

    user = st.session_state.user_info
    role, aclass, agroup = user['role'], user['class'], user['group']

    today = date.today()
    try:
        start_date_str = load_setting('school_year_start_date', '2025-09-08')
        h_json = json.loads(load_setting(HOLIDAY_SETTINGS_KEY, '[]'))
        h_in_year = [(datetime.strptime(s, '%Y-%m-%d').date(), datetime.strptime(e, '%Y-%m-%d').date()) for s, e in h_json]
        current_week = get_school_week_number(today, datetime.strptime(start_date_str, '%Y-%m-%d').date(), h_in_year)
        if current_week <= 0: current_week = 1
    except: current_week = 1
        
    start_w, end_w = get_week_dates_by_number(current_week)
    if not start_w: 
        start_w = today - timedelta(days=today.weekday())
        end_w = today + timedelta(days=6-today.weekday())

    all_events = lay_su_kien_trong_khoang_ngay_db(start_w.strftime('%Y-%m-%d'), (end_w + timedelta(days=1)).strftime('%Y-%m-%d'), role, aclass, agroup)
    all_students = lay_thong_tin_day_du_hoc_sinh_db(user_role=role, assigned_class=aclass, assigned_group=agroup)
    
    events_by_student = {}
    for ev in all_events:
        events_by_student.setdefault(ev[0], []).append(ev)
        
    student_list = []
    for hs_tuple in all_students:
        try:
            hs = dict(zip(["id"] + STUDENT_FIELDS_DB, hs_tuple))
            score = DIEM_KHOI_DAU + sum(e[6] for e in events_by_student.get(hs['id'], []))
            student_list.append({
                'id': hs['id'], 'ten': hs['ten'], 'lop': hs['lop'], 'to': hs.get('to_nhiem_vu', '') or "Không rõ", 
                'diem': score, 'anh_the': hs.get('anh_the_path', '')
            })
        except Exception as e:
            continue
    
    student_list.sort(key=lambda x: x['diem'], reverse=True)

    if not student_list:
        st.info("Chưa có dữ liệu học sinh để xếp hạng.")
        return

    st.subheader(f"🌟 TOP 3 HỌC SINH XUẤT SẮC NHẤT TUẦN {current_week}")
    
    top3 = student_list[:3]
    col2, col1, col3 = st.columns([1, 1.2, 1], gap="medium")
    
    # Hàm con mã hóa ảnh an toàn chống lỗi
    def get_image_base64(path):
        if path and isinstance(path, str) and os.path.exists(os.path.join('student_photos', path)):
            try:
                with open(os.path.join('student_photos', path), "rb") as img_file:
                    encoded = base64.b64encode(img_file.read()).decode()
                    return f"data:image/jpeg;base64,{encoded}"
            except: pass
        return "https://cdn-icons-png.flaticon.com/512/149/149071.png" # Ảnh mặc định

    def draw_podium(col, hs, rank, icon, color):
        img_b64 = get_image_base64(hs['anh_the'])
        with col:
            st.markdown(f"""
            <div style='background-color: {color}; padding: 20px; border-radius: 15px; text-align: center; box-shadow: 2px 2px 10px rgba(0,0,0,0.1); height: 100%;'>
                <h1 style='margin:0;'>{icon}</h1>
                <img src="{img_b64}" style="width: 140px; height: 140px; object-fit: cover; border-radius: 50%; border: 4px solid #DAA520; margin: 10px 0; box-shadow: 0px 4px 8px rgba(0,0,0,0.2);">
                <h3 style='margin:5px 0 5px 0; color: #333;'>{hs['ten']}</h3>
                <p style='margin:0; font-size:18px; font-weight:bold; color: #d63031;'>{hs['diem']} điểm</p>
                <p style='margin:0; color: #636e72;'>Lớp: {hs['lop']} | Tổ: {hs['to']}</p>
            </div>
            """, unsafe_allow_html=True)
            
            try:
                cert_img = create_certificate_image(hs['ten'], hs['lop'], f"Đạt Top {rank} Xuất sắc Tuần {current_week}")
                st.download_button("📥 Tải Bằng Khen", data=cert_img, file_name=f"BangKhen_Top{rank}_{hs['ten']}.jpg", mime="image/jpeg", width="stretch", key=f"btn_cert_{hs['id']}")
            except Exception as e:
                st.error("Lỗi tạo bằng khen")

    if len(top3) > 0: 
        draw_podium(col1, top3[0], 1, "🥇", "#FFF9E6") 
        st.balloons() 
    if len(top3) > 1: draw_podium(col2, top3[1], 2, "🥈", "#F2F2F2")
    if len(top3) > 2: draw_podium(col3, top3[2], 3, "🥉", "#FFF0E6")

    st.markdown("<br><hr>", unsafe_allow_html=True)

    col_t_left, col_t_right = st.columns([1, 1])
    with col_t_left:
        st.subheader("👥 Bảng Xếp Hạng Tổ (Tuần này)")
        df_hs = pd.DataFrame(student_list)
        if not df_hs.empty:
            team_ranking = df_hs.groupby('to')['diem'].mean().reset_index()
            team_ranking.rename(columns={'to': 'Tên Tổ', 'diem': 'Điểm Trung Bình'}, inplace=True)
            team_ranking.sort_values(by='Điểm Trung Bình', ascending=False, inplace=True)
            team_ranking.insert(0, 'Hạng', range(1, len(team_ranking) + 1))
            st.dataframe(team_ranking, hide_index=True, width="stretch")

    with col_t_right:
        st.subheader("📜 Top 4-10 (Cố gắng vươn lên)")
        if len(student_list) > 3:
            df_rest = pd.DataFrame(student_list[3:10])
            df_rest.insert(0, 'Hạng', range(4, 4 + len(df_rest)))
            df_rest.rename(columns={'ten': 'Họ Tên', 'lop': 'Lớp', 'diem': 'Điểm'}, inplace=True)
            st.dataframe(df_rest[['Hạng', 'Họ Tên', 'Lớp', 'Điểm']], hide_index=True, width="stretch")
# --- ĐIỀU HƯỚNG ---
if not st.session_state.logged_in:
    show_login_page()
else:
    show_main_dashboard()
