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
    cap_nhat_anh_the_db,             # <<< Đã có dấu phẩy ở đây
    xoa_nhieu_danh_muc_su_kien_db, # Của lần trước
    cap_nhat_xu_thuong_db, 
    lay_so_du_xu_db, #<<< THÊM 2 HÀM NÀY
    xoa_su_kien_ren_luyen_db,  # <<< THÊM HÀM NÀY VÀO ĐÂY# 
    luu_nhat_ky_phan_hoi_thang_db, lay_lich_su_phan_hoi_db # <<< THÊM 2 HÀM NÀY
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
# =========================================================
# TỐI ƯU HÓA KHÔNG GIAN GIAO DIỆN (ÉP SÁT LỀ TỐI ĐA)
# =========================================================
st.markdown("""
    <style>
        /* 1. Ép sát trên cùng cho phần Nội dung chính bên phải */
        .block-container {
            padding-top: 2rem !important; 
            padding-bottom: 1rem !important;
        }
        
        /* 2. ÉP SÁT LÊN TRÊN CÙNG CHO THANH SIDEBAR BÊN TRÁI (QUAN TRỌNG NHẤT) */
        [data-testid="stSidebar"] > div:first-child {
            padding-top: 1.5rem !important;
        }
        [data-testid="stSidebarUserContent"] {
            padding-top: 0rem !important;
        }

        /* 3. Thu nhỏ các đường gạch ngang (---) */
        hr {
            margin-top: 0.5rem !important;
            margin-bottom: 0.5rem !important;
        }
        
        /* 4. Ép các tiêu đề h1, h2, h3 không bị sinh khoảng trắng thừa */
        h1, h2, h3 {
            padding-top: 0rem !important;
            margin-top: 0rem !important;
        }
    </style>
""", unsafe_allow_html=True)
# Gọi hàm chạy ngầm ngay khi mở web
setup_pwa()
# =========================================================
# Khởi tạo kết nối CSDL (dùng chung file database.py cũ của thầy)
init_db()

# --- QUẢN LÝ TRẠNG THÁI ĐĂNG NHẬP (SESSION STATE) ---
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.user_info = None

# --- HÀM 1: GIAO DIỆN ĐĂNG NHẬP (ĐÃ KHÓA FORM CHỐNG LỖI) ---
def show_login_page():
    st.title("🎓 Hệ thống Quản lý DLPOINT 2.0")
    st.markdown("---")
    
    # Tạo 3 cột để căn giữa form đăng nhập
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.subheader("Đăng nhập")
        
        # <<< SỬ DỤNG st.form ĐỂ KHÓA CHẶT QUÁ TRÌNH NHẬP LIỆU >>>
        with st.form("form_dang_nhap"):
            username = st.text_input("Tên đăng nhập")
            password = st.text_input("Mật khẩu", type="password")
            
            # Nút Đăng nhập giờ đây là Nút Submit của Form
            submitted = st.form_submit_button("Đăng nhập", type="primary", use_container_width=True)
            
            if submitted:
                # Chỉ kiểm tra CSDL khi người dùng thực sự bấm nút
                if not username or not password:
                    st.warning("Vui lòng nhập đầy đủ tài khoản và mật khẩu.")
                else:
                    user = verify_user(username, password)
                    if user:
                        # Đăng nhập thành công, lưu thông tin vào session
                        st.session_state.logged_in = True
                        st.session_state.user_info = user
                        st.rerun() # Tải lại trang web để chuyển vào Dashboard
                    else:
                        st.error("❌ Tên đăng nhập hoặc mật khẩu không đúng!")

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

            # === TAB LỊCH SỬ (ĐÃ NÂNG CẤP TÍNH NĂNG XÓA) ===
            with tab_lich_su:
                st.write("**1. Lịch sử Kỷ luật (Theo TT19)**")
                history_kl = lay_lich_su_ky_luat_cua_hoc_sinh_db(hs_id)
                if history_kl:
                    df_kl = pd.DataFrame(history_kl, columns=["ID", "Hình thức", "Ngày", "Ghi chú"])
                    df_kl['Ngày'] = pd.to_datetime(df_kl['Ngày']).dt.strftime('%d/%m/%Y')
                    st.dataframe(df_kl[["Ngày", "Hình thức", "Ghi chú"]], width="stretch", hide_index=True)
                else: st.success("Học sinh chưa bị áp dụng hình thức kỷ luật nào.")
                
                st.markdown("---")
                st.write("**2. Chi tiết Sự kiện (+/- điểm)**")
                
                all_events = lay_su_kien_trong_khoang_ngay_db('2020-01-01', '2100-01-01', user['role'], user['class'], user['group'])
                hs_events = [e for e in all_events if e[0] == hs_id]
                
                if hs_events:
                    df_events = pd.DataFrame(hs_events, columns=["ID", "Tên", "Lớp", "Tổ", "Nội dung", "Loại", "Điểm", "Ngày"])
                    df_events['Ngày'] = pd.to_datetime(df_events['Ngày']).dt.strftime('%d/%m/%Y %H:%M')
                    
                    # Thêm cột Checkbox để chọn xóa
                    df_events.insert(0, "Chọn xóa", False)
                    
                    st.info("💡 Nếu ghi nhận nhầm, thầy/cô hãy đánh dấu tick [x] vào sự kiện ở bảng dưới và bấm nút Xóa.")
                    
                    # Tạo bảng tương tác Data Editor
                    edited_ev_df = st.data_editor(
                        df_events[["Chọn xóa", "Ngày", "Nội dung", "Loại", "Điểm", "ID"]].sort_values("Ngày", ascending=False),
                        hide_index=True,
                        column_config={
                            "Chọn xóa": st.column_config.CheckboxColumn("Xóa", required=True),
                            "ID": None # Ẩn cột ID
                        },
                        disabled=["Ngày", "Nội dung", "Loại", "Điểm", "ID"], # Khóa các cột khác, chỉ cho phép tick ô Xóa
                        use_container_width=True,
                        key=f"editor_del_ev_{hs_id}"
                    )
                    
                    # Lọc ra các ID sự kiện được tick True
                    selected_ev_rows = edited_ev_df[edited_ev_df["Chọn xóa"] == True]
                    ids_to_del = selected_ev_rows["ID"].tolist()
                    
                    if len(ids_to_del) > 0:
                        if st.button(f"🗑️ Xác nhận xóa {len(ids_to_del)} sự kiện bị nhầm", type="primary", key=f"btn_del_ev_{hs_id}"):
                            count_del = 0
                            with st.spinner("Đang xóa dữ liệu..."):
                                for ev_id in ids_to_del:
                                    if xoa_su_kien_ren_luyen_db(ev_id):
                                        count_del += 1
                                        
                            if count_del > 0:
                                st.toast(f"Đã gỡ bỏ thành công {count_del} sự kiện!", icon="✅")
                                
                                # <<< THÊM 2 DÒNG NÀY ĐỂ XÓA BỘ NHỚ TẠM CỦA BẢNG TRƯỚC KHI TẢI LẠI >>>
                                if f"editor_del_ev_{hs_id}" in st.session_state:
                                    del st.session_state[f"editor_del_ev_{hs_id}"]
                                
                                st.rerun() # Tải lại trang
                else: 
                    st.info("Chưa có sự kiện nào.")

            # === TAB 3: MỤC TIÊU & PHẢN HỒI (LƯU THEO LỊCH SỬ THÁNG) ===
            with tab_muc_tieu:
                st.write("📌 **Viết nhận xét & Đặt mục tiêu cho tháng:**")
                
                # 1. Chọn tháng để viết nhận xét (Mặc định là tháng hiện tại)
                now = datetime.now()
                thang_hien_tai = f"{now.month:02d}/{now.year}"
                
                # Lấy lịch sử cũ từ CSDL để hiển thị
                lich_su_ph = lay_lich_su_phan_hoi_db(hs_id)
                dict_ph = {row[0]: {"muc_tieu": row[1], "phan_hoi": row[2]} for row in lich_su_ph}
                
                # Danh sách các tháng để chọn (Tháng hiện tại + Các tháng đã có lịch sử)
                danh_sach_thang = list(set([thang_hien_tai] + list(dict_ph.keys())))
                danh_sach_thang.sort(reverse=True) # Sắp xếp mới nhất lên đầu
                
                selected_month = st.selectbox("Chọn kỳ đánh giá:", danh_sach_thang, key=f"month_sel_{hs_id}")
                
                # Tải dữ liệu của tháng được chọn lên Form
                muc_tieu_cu = dict_ph.get(selected_month, {}).get("muc_tieu", "")
                phan_hoi_cu = dict_ph.get(selected_month, {}).get("phan_hoi", "")
                
                # --- NÚT GỌI TRỢ LÝ AI (Chỉ phân tích dữ liệu của tháng được chọn) ---
                if st.button(f"✨ Nhờ AI viết nhận xét cho {selected_month}", type="secondary", width="stretch", key=f"ai_btn_{hs_id}"):
                    if "GEMINI_API_KEY" not in st.secrets:
                        st.error("Chưa cấu hình API Key của Google Gemini!")
                    else:
                        with st.spinner("🤖 AI đang đọc hồ sơ và soạn lời phê..."):
                            try:
                                from google import genai
                                import calendar
                                client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
                                
                                # Lọc dữ liệu đúng cái tháng đang được chọn
                                m, y = map(int, selected_month.split('/'))
                                start_m = datetime(y, m, 1).date()
                                _, last_day = calendar.monthrange(y, m)
                                end_m = datetime(y, m, last_day).date()
                                
                                sk_thang = [e for e in hs_events if start_m <= e[7].date() <= end_m]
                                
                                diem_thang = DIEM_KHOI_DAU + sum(e[6] for e in sk_thang)
                                kt_str = ", ".join([e[4] for e in sk_thang if e[5] == "Khen thưởng"]) or "Chưa có nổi bật"
                                vp_str = ", ".join([e[4] for e in sk_thang if e[5] == "Vi phạm"]) or "Không vi phạm"
                                
                                prompt = f"Đóng vai là GVCN. Hãy viết 1 đoạn nhận xét cuối tháng (3-4 câu) cho học sinh {info.get('ten')}. Điểm rèn luyện tháng này: {diem_thang}/100. Ưu điểm: {kt_str}. Khuyết điểm: {vp_str}. Yêu cầu: Giọng chuẩn mực sư phạm, xưng 'thầy/cô' gọi 'em', động viên khen ngợi, nhắc nhở khéo léo. Trả về đúng đoạn văn."
                                response = client.models.generate_content(model='gemini-3.6-flash', contents=prompt)
                                
                                st.session_state[f"ai_phan_hoi_{hs_id}_{selected_month}"] = response.text
                                st.toast("AI đã viết xong!", icon="🎉")
                            except Exception as e: st.error(f"Lỗi AI: {e}")

                hien_thi_phan_hoi = st.session_state.get(f"ai_phan_hoi_{hs_id}_{selected_month}", phan_hoi_cu)
                
                # 2. Ô Nhập liệu
                new_muc_tieu = st.text_area("🎯 Mục tiêu do Học sinh/GVCN đề ra:", value=muc_tieu_cu, height=80, key=f"mt_{hs_id}")
                new_phan_hoi = st.text_area("💬 Nhận xét & Phản hồi của GVCN:", value=hien_thi_phan_hoi, height=120, key=f"ph_{hs_id}")
                
                if st.button("💾 Lưu Nhận xét Tháng này", type="primary", width="stretch", key=f"save_ph_{hs_id}"):
                    if luu_nhat_ky_phan_hoi_thang_db(hs_id, selected_month, new_muc_tieu, new_phan_hoi):
                        st.success(f"✅ Đã lưu hồ sơ thành công cho tháng {selected_month}!")
                        if f"ai_phan_hoi_{hs_id}_{selected_month}" in st.session_state: 
                            del st.session_state[f"ai_phan_hoi_{hs_id}_{selected_month}"]
                        st.rerun()

                # 3. Khu vực hiển thị Lịch sử các tháng cũ (Dạng Dòng thời gian - Timeline)
                st.markdown("---")
                st.write("🕰️ **Lịch sử Phản hồi các tháng trước:**")
                if lich_su_ph:
                    for row in lich_su_ph:
                        thang, mt, ph, ngay_cap_nhat = row[0], row[1], row[2], row[3]
                        with st.expander(f"Tháng {thang} (Cập nhật: {ngay_cap_nhat.strftime('%d/%m/%Y')})", expanded=False):
                            st.write(f"**🎯 Mục tiêu:** {mt if mt else 'Không có'}")
                            st.write(f"**💬 Nhận xét:** {ph if ph else 'Không có'}")
                else:
                    st.info("Chưa có lịch sử nhận xét nào được lưu.")
# --- HÀM 8: GHI NHẬN NHANH (GIAO DIỆN HỢP NHẤT - KHÔNG GIẬT TRANG) ---
def show_quick_record_page():
    st.header("📝 Ghi nhận Điểm Cộng / Trừ Nhanh")
    st.markdown("---")

    user = st.session_state.user_info
    role, aclass, agroup = user['role'], user['class'], user['group']
    
    if role == 'bcs':
        st.info(f"🔒 **Chế độ Tổ trưởng:** Ghi nhận cho các thành viên **Tổ {agroup} - Lớp {aclass}**.")
    
    raw_students = get_cached_students(role, aclass, agroup)
    if not raw_students:
        st.warning("Không có học sinh nào trong phạm vi quản lý của bạn.")
        return

    # Chuẩn bị dữ liệu Sự kiện
    all_events = get_cached_events()
    event_list_for_table = [] 
    for ev in all_events:
        ten, loai, diem = ev[1], ev[2], ev[3]
        diem_str = f"+{diem}" if diem > 0 else f"{diem}"
        icon = "🟢" if loai == "Khen thưởng" else "🔴"
        event_list_for_table.append({"Số lần": 0, "Loại": icon, "Tên Sự Kiện": ten, "Điểm gốc": diem, "Loại gốc": loai})

    # Cấu hình Ngày
    event_date = st.date_input("📅 Chọn Ngày ghi nhận:", format="DD/MM/YYYY", key="date_quick_unified")
    ngay_tao_str = event_date.strftime('%Y-%m-%d %H:%M:%S')
    st.markdown("---")

    # MẸO: Dùng session_state để tạo "chìa khóa" reset bảng sau khi lưu thành công
    if 'reset_table_key' not in st.session_state:
        st.session_state.reset_table_key = 0

    st.info("💡 **Cách dùng siêu nhanh:** (1) Tick chọn một hoặc nhiều Học sinh bên trái. (2) Tăng 'Số lần' ở một hoặc nhiều Sự kiện bên phải. (3) Bấm Ghi nhận.")

    # Chia 2 cột tỷ lệ 1:1.2
    col_hs, col_ev = st.columns([1, 1.2], gap="large")

    # ==========================================
    # CỘT TRÁI: CHỌN HỌC SINH (BẢNG CHECKBOX)
    # ==========================================
    with col_hs:
        st.write("🧑‍🎓 **1. Đánh dấu [x] chọn Học sinh:**")
        df_hs = pd.DataFrame(raw_students, columns=["ID", "Họ Tên", "Lớp", "Tổ", "SĐT", "Zalo", "Ảnh"])
        df_hs.insert(0, "Chọn", False)
        df_hs.insert(1, "STT", range(1, len(df_hs) + 1))
        
        # Gắn key động để tự reset khi cần
        key_hs = f"editor_hs_unified_{st.session_state.reset_table_key}"
        
        edited_hs_df = st.data_editor(
            df_hs[["Chọn", "STT", "Họ Tên", "Tổ", "ID"]],
            hide_index=True,
            column_config={
                "Chọn": st.column_config.CheckboxColumn("Tích", required=True), 
                "ID": None, 
                "STT": st.column_config.NumberColumn(disabled=True, width="small"), 
                "Họ Tên": st.column_config.TextColumn(disabled=True), 
                "Tổ": st.column_config.TextColumn(disabled=True, width="small")
            },
            disabled=["STT", "Họ Tên", "Tổ", "ID"],
            use_container_width=True, height=500, key=key_hs
        )
        
        # Lọc ra danh sách ID HS được chọn
        selected_rows_hs = edited_hs_df[edited_hs_df["Chọn"] == True]
        ids_to_apply = selected_rows_hs["ID"].tolist()
        names_to_apply = selected_rows_hs["Họ Tên"].tolist()

    # ==========================================
    # CỘT PHẢI: CHỌN SỰ KIỆN & SỐ LẦN
    # ==========================================
    with col_ev:
        st.write("📋 **2. Bấm (+) điền số lần Sự kiện:**")
        df_ev = pd.DataFrame(event_list_for_table)
        
        key_ev = f"editor_ev_unified_{st.session_state.reset_table_key}"
        
        edited_ev_df = st.data_editor(
            df_ev,
            hide_index=True,
            column_config={
                "Số lần": st.column_config.NumberColumn("Số lần", min_value=0, max_value=50, step=1, format="%d"),
                "Loại": st.column_config.TextColumn(disabled=True, width="small"),
                "Tên Sự Kiện": st.column_config.TextColumn(disabled=True),
                "Điểm gốc": None, # Ẩn cột dữ liệu gốc đi cho gọn
                "Loại gốc": None  # Ẩn cột dữ liệu gốc đi cho gọn
            },
            disabled=["Loại", "Tên Sự Kiện", "Điểm gốc", "Loại gốc"], 
            use_container_width=True, height=500, key=key_ev
        )
        
        # Lọc ra các sự kiện có số lần > 0
        selected_rows_ev = edited_ev_df[edited_ev_df["Số lần"] > 0]

    # ==========================================
    # KHU VỰC XÁC NHẬN VÀ LƯU DỮ LIỆU
    # ==========================================
    st.markdown("---")
    
    # Hiển thị tóm tắt trước khi lưu
    sl_hs = len(ids_to_apply)
    sl_sk = selected_rows_ev["Số lần"].sum()
    
    if sl_hs > 0 and sl_sk > 0:
        st.info(f"👉 Sắp áp dụng **{sl_sk}** lượt sự kiện cho **{sl_hs}** học sinh đã chọn.")
    
    # Nút bấm trung tâm
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        if st.button("🚀 XÁC NHẬN GHI NHẬN", type="primary", use_container_width=True):
            if sl_hs == 0 or sl_sk == 0:
                st.error("👆 Thầy/Cô cần tick chọn ít nhất 1 Học sinh (Bảng trái) VÀ tăng số lần ở ít nhất 1 Sự kiện (Bảng phải).")
            else:
                count_luu = 0
                with st.spinner("Đang xử lý dữ liệu..."):
                    # Vòng lặp kép: Quét từng HS, áp dụng từng sự kiện
                    for hs_id in ids_to_apply:
                        for _, row_ev in selected_rows_ev.iterrows():
                            ev_name = row_ev["Tên Sự Kiện"]
                            so_lan = int(row_ev["Số lần"])
                            loai_sk = row_ev["Loại gốc"]
                            diem_goc = int(row_ev["Điểm gốc"])
                            
                            # Tính toán tổng điểm và chuỗi mô tả
                            diem_final = diem_goc * so_lan
                            mo_ta_final = f"{ev_name} ({so_lan} lần)" if so_lan > 1 else ev_name
                            
                            # Gọi hàm CSDL
                            if them_su_kien_ren_luyen_db(hs_id, mo_ta_final, loai_sk, diem_final, ngay_tao_str): 
                                count_luu += 1
                
                if count_luu > 0:
                    st.success(f"✅ HOÀN TẤT! Đã lưu {count_luu} bản ghi dữ liệu thành công.")
                    st.balloons()
                    # MẸO: Tăng biến key lên 1 để Streamlit tự động "làm sạch" 2 bảng (bỏ tick và reset số về 0)
                    st.session_state.reset_table_key += 1
                    st.rerun() # Tải lại giao diện ngay lập tức
                else:
                    st.error("Có lỗi xảy ra khi lưu CSDL.")
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
# --- HÀM 5: PHÂN TÍCH & CẢNH BÁO RÈN LUYỆN (TÍCH HỢP BÁO ĐỘNG & BẢN ĐỒ HÀNH VI) ---
def show_statistics_page():
    st.header("📊 Phân tích & Cảnh báo Rèn luyện")
    st.markdown("---")

    user = st.session_state.user_info
    role, aclass, agroup = user['role'], user['class'], user['group']

    # ==========================================
    # MODULE 1: HỆ THỐNG BÁO ĐỘNG ĐỎ (EARLY WARNING)
    # ==========================================
    st.subheader("🚨 Hệ thống Cảnh báo Sớm (30 ngày qua)")
    st.write("Hệ thống tự động rà soát dữ liệu 30 ngày gần nhất để phát hiện các học sinh cần GVCN can thiệp khẩn cấp.")
    
    today = date.today()
    start_30d = today - timedelta(days=30)
    
    with st.spinner("Đang quét dữ liệu rèn luyện để tìm rủi ro..."):
        events_30d = lay_su_kien_trong_khoang_ngay_db(start_30d.strftime('%Y-%m-%d'), (today + timedelta(days=1)).strftime('%Y-%m-%d'), role, aclass, agroup)
        hs_list = get_cached_students(role, aclass, agroup)
        all_categories = get_cached_events()
        
        muc_do_map = {cat[1]: cat[4] for cat in all_categories if cat[4] is not None}
        warnings = []
        
        if hs_list and events_30d:
            ev_by_hs = {}
            for ev in events_30d:
                ev_by_hs.setdefault(ev[0], []).append(ev)
                
            for hs in hs_list:
                hs_id, hs_ten, hs_lop = hs[0], hs[1], hs[2]
                my_events = ev_by_hs.get(hs_id, [])
                
                if not my_events: continue
                    
                # 1. Cảnh báo Điểm rèn luyện thấp (Dưới 70)
                score = DIEM_KHOI_DAU + sum(e[6] for e in my_events)
                if score < 70:
                    warnings.append({'hs': hs_ten, 'lop': hs_lop, 'icon': '📉', 'color': '#e74c3c', 'title': 'Điểm rèn luyện rớt xuống mức Báo động', 'detail': f'Điểm hiện tại: {score}/100'})
                    
                # 2. Cảnh báo Chuyên cần (Vắng Không Phép >= 2 hoặc Trễ >= 3)
                vang_kp = sum(1 for e in my_events if "Vắng không phép" in e[4])
                di_tre = sum(1 for e in my_events if "Đi trễ" in e[4])
                if vang_kp >= 2 or di_tre >= 3:
                    msg = []
                    if vang_kp > 0: msg.append(f"Vắng K.Phép: {vang_kp} lần")
                    if di_tre > 0: msg.append(f"Đi trễ: {di_tre} lần")
                    warnings.append({'hs': hs_ten, 'lop': hs_lop, 'icon': '⚠️', 'color': '#f39c12', 'title': 'Nguy cơ chểnh mảng / Bỏ học', 'detail': " | ".join(msg)})
                    
                # 3. Cảnh báo Vi phạm Nghiêm trọng (Mức 2, Mức 3 theo TT19)
                loi_nghiem_trong = set()
                for e in my_events:
                    mo_ta = e[4]
                    for cat_name, m_do in muc_do_map.items():
                        if cat_name in mo_ta and m_do >= 2:
                            loi_nghiem_trong.add(f"{cat_name} (Mức {m_do})")
                            
                if loi_nghiem_trong:
                    warnings.append({'hs': hs_ten, 'lop': hs_lop, 'icon': '🚫', 'color': '#8e44ad', 'title': 'Hành vi vi phạm nghiêm trọng', 'detail': ", ".join(list(loi_nghiem_trong))})

        if not warnings:
            st.success("✨ Tình hình ổn định. Không phát hiện học sinh nào có rủi ro cao trong 30 ngày qua.")
        else:
            st.error(f"Phát hiện **{len(warnings)}** rủi ro cần GVCN lưu tâm và can thiệp sớm:")
            cols = st.columns(3)
            for i, w in enumerate(warnings):
                with cols[i % 3]:
                    st.markdown(f"""
                    <div style='background-color: #ffffff; border-left: 5px solid {w['color']}; padding: 15px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.05); margin-bottom: 15px;'>
                        <h4 style='margin-top: 0; color: {w['color']}; font-size: 1.1em;'>{w['icon']} {w['hs']}</h4>
                        <p style='margin: 0; font-size: 0.9em; color: gray;'>Lớp: {w['lop']}</p>
                        <p style='margin: 8px 0 0 0; font-weight: bold; color: #333;'>{w['title']}</p>
                        <p style='margin: 0; font-size: 0.9em; color: #d63031;'>{w['detail']}</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
    st.markdown("<br><hr>", unsafe_allow_html=True)

    # ==========================================
    # MODULE 2: BẢN ĐỒ HÀNH VI LỚP HỌC (BEHAVIOR HEATMAP)
    # ==========================================
    st.subheader("🗺️ Bản đồ Hành vi & Dữ liệu Tổng hợp")
    
    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        start_date = st.date_input("Từ ngày:", date.today() - timedelta(days=30), format="DD/MM/YYYY")
    with col2:
        end_date = st.date_input("Đến ngày:", date.today(), format="DD/MM/YYYY")
    with col3:
        st.write("") 
        st.write("")
        btn_thong_ke = st.button("📊 Chạy Phân tích Dữ liệu", type="primary", width="stretch")

    if btn_thong_ke or st.session_state.get('show_stats', False):
        st.session_state.show_stats = True
        if start_date > end_date:
            st.error("Ngày bắt đầu không được lớn hơn ngày kết thúc!")
            return
            
        with st.spinner('Đang dùng AI nội bộ phân tích dữ liệu...'):
            df = lay_su_kien_cho_thong_ke_df(start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d'))
            
            if df.empty:
                st.warning("Không có dữ liệu rèn luyện nào trong khoảng thời gian này.")
                return

            # Thêm thông tin Tên học sinh vào df để phân tích
            raw_events = lay_su_kien_trong_khoang_ngay_db(start_date.strftime('%Y-%m-%d'), (end_date + timedelta(days=1)).strftime('%Y-%m-%d'), role, aclass, agroup)
            df_full = pd.DataFrame(raw_events, columns=["ID", "Họ Tên", "Lớp", "Tổ", "Lỗi", "Loại", "Điểm", "Ngày"])
            
            df_vipham = df_full[df_full['Loại'] == 'Vi phạm'].copy()

            if df_vipham.empty:
                st.success("Tuyệt vời! Không có vi phạm nào trong khoảng thời gian này.")
                return

            # --- TẠO CÁC BIỂU ĐỒ PHÂN TÍCH SÂU ---
            st.markdown("---")
            row1_col1, row1_col2 = st.columns(2, gap="large")
            
            with row1_col1:
                st.write("🔥 **Top 5 Học sinh vi phạm nhiều nhất**")
                # Đếm số lỗi theo học sinh
                top_hs = df_vipham['Họ Tên'].value_counts().head(5)
                # Vẽ biểu đồ ngang cho dễ đọc tên
                st.bar_chart(top_hs, color="#e74c3c")
                
            with row1_col2:
                st.write("📅 **Chu kỳ vi phạm theo Thứ trong tuần**")
                # Trích xuất "Thứ" từ Ngày
                df_vipham['Ngày'] = pd.to_datetime(df_vipham['Ngày'])
                df_vipham['Thứ_Index'] = df_vipham['Ngày'].dt.dayofweek
                thu_map = {0: 'Thứ 2', 1: 'Thứ 3', 2: 'Thứ 4', 3: 'Thứ 5', 4: 'Thứ 6', 5: 'Thứ 7', 6: 'Chủ Nhật'}
                df_vipham['Thứ'] = df_vipham['Thứ_Index'].map(thu_map)
                
                # Sắp xếp đúng thứ tự từ Thứ 2 đến Chủ Nhật
                days_order = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6', 'Thứ 7', 'Chủ Nhật']
                vp_theo_thu = df_vipham['Thứ'].value_counts().reindex(days_order).fillna(0)
                st.bar_chart(vp_theo_thu, color="#f39c12")
                
            st.write("<br>", unsafe_allow_html=True)
            row2_col1, row2_col2 = st.columns(2, gap="large")

            with row2_col1:
                st.write("📌 **Top Lỗi Vi phạm phổ biến nhất**")
                top_loi = df_vipham['Lỗi'].value_counts().head(8)
                st.bar_chart(top_loi, color="#8e44ad")

            with row2_col2:
                st.write("🏫 **Phân bổ Vi phạm theo Tổ/Lớp**")
                if role == 'admin':
                    vp_theo_vung = df_vipham['Lớp'].value_counts()
                else:
                    vp_theo_vung = df_vipham['Tổ'].fillna("Không rõ").value_counts()
                st.bar_chart(vp_theo_vung, color="#3498db")

            st.markdown("---")
            st.write("📝 **Chi tiết Dữ liệu Vi phạm (Bảng thô)**")
            # ==========================================
            # MODULE 3: TRỢ LÝ AI PHÂN TÍCH LỚP HỌC
            # ==========================================
            st.markdown("---")
            st.subheader("🤖 Trợ lý AI Phân tích Chiến lược Lớp học")
            st.info("AI sẽ tổng hợp toàn bộ số liệu biểu đồ ở trên để viết ra một Báo cáo phân tích chuyên sâu, giúp GVCN nắm bắt gốc rễ vấn đề và có hướng giải quyết hiệu quả.")

            if st.button("✨ Nhờ AI Phân tích Tình hình Lớp học", type="secondary", width="stretch", key="btn_ai_class"):
                if "GEMINI_API_KEY" not in st.secrets:
                    st.error("Chưa cấu hình API Key của Google Gemini trong Streamlit Secrets!")
                else:
                    with st.spinner("🤖 AI đang phân tích hàng trăm dòng dữ liệu và suy nghĩ chiến lược..."):
                        try:
                            from google import genai
                            client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

                            # Lọc số liệu chuẩn bị đưa cho AI đọc
                            so_vi_pham = len(df_vipham)
                            df_khen = df_full[df_full['Loại'] == 'Khen thưởng']
                            so_khen_thuong = len(df_khen)

                            top_loi_str = ", ".join(df_vipham['Lỗi'].value_counts().head(3).index.tolist()) if so_vi_pham > 0 else "Không có lỗi nào"
                            top_hs_vp_str = ", ".join(df_vipham['Họ Tên'].value_counts().head(3).index.tolist()) if so_vi_pham > 0 else "Không có học sinh vi phạm"
                            top_hs_khen_str = ", ".join(df_khen['Họ Tên'].value_counts().head(3).index.tolist()) if so_khen_thuong > 0 else "Chưa có"

                            # Ra lệnh cho AI (Prompt Engineering)
                            prompt = f"""
                            Bạn là một chuyên gia quản lý giáo dục và tư vấn tâm lý học đường cấp cao.
                            Dưới đây là số liệu thống kê rèn luyện của lớp học trong khoảng thời gian vừa qua:
                            - Tổng số lượt khen thưởng: {so_khen_thuong} lượt (Top học sinh xuất sắc: {top_hs_khen_str})
                            - Tổng số lượt vi phạm: {so_vi_pham} lượt.
                            - Các lỗi vi phạm phổ biến nhất của lớp: {top_loi_str}.
                            - Những học sinh cá biệt vi phạm nhiều nhất cần chú ý: {top_hs_vp_str}.

                            Dựa vào các số liệu thực tế trên, hãy viết một báo cáo tư vấn ngắn gọn (khoảng 150-200 chữ) cho Giáo viên chủ nhiệm bao gồm 3 phần:
                            1. Đánh giá tổng quan về nề nếp của lớp.
                            2. Phân tích nguyên nhân sâu xa của các lỗi phổ biến (dựa trên tâm lý học sinh).
                            3. Đề xuất 2-3 biện pháp cụ thể, thiết thực và mang tính giáo dục tích cực (Không dùng đòn roi/chửi mắng) để GVCN chấn chỉnh lớp trong tháng tới.
                            
                            Văn phong: Chuyên nghiệp, tôn trọng, mang tính tư vấn, xưng là "Trợ lý AI DLPOINT".
                            """

                            # Gọi AI sinh kết quả
                            response = client.models.generate_content(
                                model='gemini-3.6-flash',
                                contents=prompt
                            )

                            st.success("✅ Trợ lý AI đã phân tích xong!")
                            
                            # Hiển thị kết quả trong một khung xanh dương tuyệt đẹp
                            st.markdown(f"""
                            <div style='background-color: #f0f7ff; padding: 20px; border-radius: 10px; border-left: 5px solid #0078D7; font-size: 1.05em; line-height: 1.6; box-shadow: 2px 2px 8px rgba(0,0,0,0.05);'>
                                {response.text}
                            </div>
                            <br>
                            """, unsafe_allow_html=True)

                        except Exception as e:
                            st.error(f"Lỗi khi gọi AI: {e}")
            df_vipham['Ngày'] = df_vipham['Ngày'].dt.strftime('%d/%m/%Y %H:%M')
            st.dataframe(df_vipham[['Ngày', 'Họ Tên', 'Lớp', 'Tổ', 'Lỗi', 'Điểm']], width="stretch", hide_index=True)
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
                
                # ==========================================
                # LỌC DỮ LIỆU ĐỂ BƠM VÀO FILE EXCEL CHI TIẾT
                # ==========================================
                student_id_to_name = {hs[0]: hs[1] for hs in all_students}
                
                # 1. Lọc điểm cộng
                positive_points = {}
                for ev in all_events:
                    if ev[5] == "Khen thưởng" and ev[6] > 0:
                        positive_points[ev[0]] = positive_points.get(ev[0], 0) + ev[6]
                pos_list = sorted([(student_id_to_name.get(sid, ""), pts) for sid, pts in positive_points.items()], key=lambda x: x[1], reverse=True)
                
                # 2. Lọc danh sách chưa có điểm cộng
                zero_list = sorted([name for sid, name in student_id_to_name.items() if sid not in positive_points])
                
                # 3. Tổng hợp vi phạm
                vio_summary = {}
                for ev in all_events:
                    if ev[5] == "Vi phạm":
                        key = (ev[0], ev[4]) # Gom nhóm theo ID và Tên lỗi
                        if key not in vio_summary:
                            vio_summary[key] = {'ten': ev[1], 'lop': ev[2], 'to': ev[3] or '', 'mo_ta': ev[4], 'count': 0}
                        vio_summary[key]['count'] += 1
                vio_list = sorted(vio_summary.values(), key=lambda x: (x['ten'], x['mo_ta']))

                # ==========================================
                # TẠO VÀ XUẤT FILE EXCEL
                # ==========================================
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    filepath = tmp.name
                    
                # Gọi hàm xuất Excel và truyền đầy đủ 3 danh sách vừa lọc vào
                success, msg = generate_weekly_summary_excel(
                    student_list, week_num, user['full_name'], user['class'] or "Toàn trường", user['group'] or "", 
                    start_date, end_date, filepath, pos_list, zero_list, vio_list 
                )
                
                with col_btn1:
                    if success:
                        with open(filepath, "rb") as f: excel_data = f.read()
                        st.download_button("📥 TẢI FILE EXCEL BÁO CÁO TUẦN (BẢN ĐẦY ĐỦ)", data=excel_data, file_name=f"TongKet_Tuan_{week_num}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", width="stretch")
                    os.unlink(filepath)
                    
                with col_btn2:
                    st.info("👆 File Excel đính kèm chứa 5 sheet chi tiết giống hệt bản Desktop (Xếp hạng, Vi phạm, Khen thưởng...).")

                # ==========================================
                # TÍNH NĂNG MỚI: TRỢ LÝ SOẠN TIN NHẮN / BÁO CÁO
                # ==========================================
                # 1. TRỢ LÝ CHO GIÁO VIÊN CHỦ NHIỆM / ADMIN
                if user['role'] in ['gvcn', 'admin']:
                    with st.expander("💬 Mở Trợ lý tạo tin nhắn Zalo gửi Phụ huynh"):
                        st.info("Hệ thống tự động quét dữ liệu và soạn sẵn tin nhắn. Thầy/Cô có thể copy dán vào nhóm Zalo lớp.")
                        
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
                        
                        st.text_area("Văn bản tin nhắn (Có thể chỉnh sửa):", value=msg_zalo, height=200, key="txt_zalo_gv")

                # 2. TRỢ LÝ BÁO CÁO DÀNH CHO BAN CÁN SỰ (TỔ TRƯỞNG)
                elif user['role'] == 'bcs':
                    with st.expander("💬 Mở Trợ lý soạn Báo cáo Tổ (Dùng để đọc trong giờ Sinh hoạt)"):
                        st.info("Hệ thống đã tự động tổng hợp số liệu. Em có thể cầm điện thoại, đọc nguyên văn bản báo cáo này trước lớp vào giờ sinh hoạt cuối tuần.")
                        
                        top_students_bcs = [hs['ten'] for hs in student_list[:3] if hs['diem'] > DIEM_KHOI_DAU]
                        vipham_dict_bcs = {}
                        for ev in all_events:
                            if ev[5] == 'Vi phạm': vipham_dict_bcs.setdefault(ev[1], []).append(ev[4])

                        group_name = user['group'] or "..."
                        class_name = user['class'] or "..."
                        
                        msg_bcs = f"Kính thưa Thầy/Cô giáo chủ nhiệm cùng toàn thể các bạn trong lớp.\n"
                        msg_bcs += f"Sau đây, em xin thay mặt Tổ {group_name} báo cáo tình hình thi đua của tổ trong Tuần {week_num} vừa qua như sau:\n\n"
                        
                        msg_bcs += "📌 1. ƯU ĐIỂM / ĐIỂM SÁNG:\n"
                        if top_students_bcs: 
                            msg_bcs += f"- Nhìn chung các bạn có ý thức tốt. Đặc biệt biểu dương các bạn rất tích cực, đạt điểm thi đua cao là: {', '.join(top_students_bcs)}.\n\n"
                        else:
                            msg_bcs += "- Đa số các bạn trong tổ có ý thức chấp hành nội quy, đi học và trực nhật đầy đủ.\n\n"
                        
                        msg_bcs += "📌 2. HẠN CHẾ / TỒN TẠI:\n"
                        if vipham_dict_bcs: 
                            msg_bcs += "- Tuy nhiên, tổ vẫn còn một số tồn tại làm ảnh hưởng đến điểm thi đua chung. Cụ thể:\n"
                            for ten, cac_loi in vipham_dict_bcs.items():
                                loi_rut_gon = ", ".join(list(set(cac_loi)))
                                msg_bcs += f"   + Bạn {ten} (Mắc lỗi: {loi_rut_gon}).\n"
                            msg_bcs += "\n"
                        else: 
                            msg_bcs += "- Rất đáng khen là tuần này tổ chúng ta thực hiện xuất sắc, không có bạn nào vi phạm nội quy!\n\n"
                            
                        msg_bcs += "📌 3. ĐỀ XUẤT:\n"
                        if top_students_bcs:
                            msg_bcs += "- Kính đề nghị Thầy/Cô tuyên dương các bạn có thành tích tốt để làm gương.\n"
                        if vipham_dict_bcs:
                            msg_bcs += "- Yêu cầu các bạn còn vi phạm nghiêm túc rút kinh nghiệm và sửa đổi. Nếu tuần sau tái phạm, tổ sẽ đề nghị Thầy/Cô xử lý kỷ luật.\n"
                        else:
                            msg_bcs += "- Mong các bạn trong tổ tiếp tục phát huy tinh thần kỷ luật tốt như tuần này.\n"

                        msg_bcs += f"\nDạ, phần báo cáo của Tổ {group_name} đến đây là hết. Em xin cảm ơn Thầy/Cô và các bạn đã lắng nghe!"
                        
                        st.text_area("Văn bản báo cáo (Có thể chỉnh sửa thêm nếu cần):", value=msg_bcs, height=350, key="txt_report_bcs")

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
                            
                            # Làm sạch tên cột (cắt dấu cách thừa) để chống lỗi
                            df_imp_sk.columns = df_imp_sk.columns.str.strip()
                            
                            if "Tên sự kiện" not in df_imp_sk.columns or "Loại" not in df_imp_sk.columns:
                                st.error("❌ File Excel không đúng chuẩn. Vui lòng tải lại File Mẫu mới nhất ở cột bên trái!")
                            else:
                                with st.spinner("Đang lưu vào hệ thống..."):
                                    for _, row in df_imp_sk.iterrows():
                                        # Lấy và làm sạch dữ liệu an toàn
                                        ten = str(row.get("Tên sự kiện", "")).strip()
                                        loai_goc = str(row.get("Loại", "")).strip().lower()
                                        diem_val = row.get("Điểm")
                                        md_val = row.get("Mức độ vi phạm", None)
                                        
                                        if not ten or ten == "nan":
                                            continue
                                            
                                        # Hệ thống tự động sửa lỗi chính tả cho cột Loại
                                        if "vi" in loai_goc or "phạm" in loai_goc: 
                                            loai = "Vi phạm"
                                        elif "khen" in loai_goc or "thưởng" in loai_goc: 
                                            loai = "Khen thưởng"
                                        else: 
                                            fail_sk += 1; continue
                                            
                                        try: diem = int(float(diem_val)) # Xử lý an toàn nếu Excel tự định dạng
                                        except: fail_sk += 1; continue
                                        
                                        # Xử lý mức độ
                                        muc_do = None
                                        if loai == "Vi phạm" and pd.notna(md_val) and str(md_val).strip() != "nan":
                                            try: muc_do = int(float(md_val))
                                            except: pass
                                            
                                        # Ghi vào CSDL (Có thì cập nhật, Chưa thì thêm mới)
                                        if them_hoac_cap_nhat_danh_muc_db(ten, loai, diem, muc_do):
                                            success_sk += 1
                                        else: fail_sk += 1
                                        
                                if success_sk > 0:
                                    st.success(f"✅ Nhập hoàn tất: Thành công {success_sk} sự kiện. Bỏ qua dòng trống/lỗi: {fail_sk}.")
                                    st.rerun() # Tải lại trang để hiện list mới
                                else:
                                    st.error(f"❌ Không có sự kiện nào được nhập. Bị lỗi {fail_sk} dòng. Vui lòng kiểm tra lại file Excel (Điểm phải là số).")
                        except Exception as e:
                            st.error(f"Lỗi đọc file: {e}")

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
    # ==========================================
    # TAB 4: QUẢN LÝ LỚP & HỌC SINH (NHẬP/XÓA/CẬP NHẬT ẢNH HÀNG LOẠT)
    # ==========================================
    with tab_class:
        st.info("Khu vực này giúp thầy khởi tạo lớp mới nhanh chóng, dọn dẹp dữ liệu, hoặc gắn ảnh thẻ cho cả lớp chỉ bằng 1 file Excel.")
        
        # Chia làm 3 cột
        col_import, col_photo, col_delete = st.columns([1, 1, 1], gap="medium")

        # ---------------------------------------
        # CỘT 1: TẠO LỚP (GIỮ NGUYÊN NHƯ CŨ)
        # ---------------------------------------
        with col_import:
            st.subheader("📥 1. Tạo Lớp Mới")
            df_template = pd.DataFrame(columns=["Họ tên học sinh", "Lớp", "Tổ", "Ngày tháng năm sinh", "Địa chỉ", "SĐT Học sinh", "SĐT Zalo", "Ghi chú"])
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer: df_template.to_excel(writer, index=False)
            
            st.download_button("⬇️ Tải File Mẫu (Tạo lớp)", data=output.getvalue(), file_name="Mau_Tao_Lop.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            
            uploaded_file = st.file_uploader("Tải lên danh sách học sinh mới:", type=['xlsx', 'xls'])
            if uploaded_file is not None:
                if st.button("🚀 Nhập học sinh mới", type="primary", use_container_width=True):
                    try:
                        df_import = pd.read_excel(uploaded_file)
                        db_map = {"Họ tên học sinh": "ten", "Lớp": "lop", "Tổ": "to_nhiem_vu", "Ngày tháng năm sinh": "ngay_sinh", "Địa chỉ": "dia_chi", "SĐT Học sinh": "sdt_hoc_sinh", "SĐT Zalo": "sdt_zalo", "Ghi chú": "ghi_chu"}
                        success_count, fail_count = 0, 0
                        with st.spinner("Đang lưu dữ liệu vào CSDL..."):
                            for _, row in df_import.iterrows():
                                if pd.isna(row.get("Họ tên học sinh")) or pd.isna(row.get("Lớp")): continue
                                data_db = {db_col: str(row.get(ex_col)).strip() if pd.notna(row.get(ex_col)) else "" for ex_col, db_col in db_map.items()}
                                if them_hoc_sinh_db(data_db): success_count += 1
                                else: fail_count += 1
                        st.success(f"Đã nhập thành công {success_count} học sinh. Lỗi: {fail_count} dòng.")
                    except Exception as e: st.error(f"Lỗi: {e}")

        # ---------------------------------------
        # CỘT 2: CẬP NHẬT ẢNH THẺ HÀNG LOẠT (TÍNH NĂNG MỚI)
        # ---------------------------------------
        with col_photo:
            st.subheader("🖼️ 2. Gắn Ảnh Thẻ")
            st.write("Cập nhật tên file ảnh cho các học sinh ĐÃ CÓ trong hệ thống.")
            
            # Tạo file mẫu cập nhật ảnh
            df_photo_template = pd.DataFrame(columns=["Họ và Tên", "Lớp", "Tên file ảnh"])
            out_photo = io.BytesIO()
            with pd.ExcelWriter(out_photo, engine='openpyxl') as writer: df_photo_template.to_excel(writer, index=False)
            
            st.download_button("⬇️ Tải File Mẫu (Gắn ảnh)", data=out_photo.getvalue(), file_name="Mau_Cap_Nhat_Anh.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            
            uploaded_photo_file = st.file_uploader("Tải lên danh sách gắn ảnh:", type=['xlsx', 'xls'], key="up_photo")
            if uploaded_photo_file is not None:
                if st.button("🖼️ Chạy Cập nhật Ảnh", type="primary", use_container_width=True):
                    try:
                        df_p = pd.read_excel(uploaded_photo_file)
                        if "Họ và Tên" not in df_p.columns or "Lớp" not in df_p.columns or "Tên file ảnh" not in df_p.columns:
                            st.error("File không hợp lệ. Phải có 3 cột: Họ và Tên, Lớp, Tên file ảnh.")
                        else:
                            # Lấy danh sách toàn bộ học sinh để dò tìm ID
                            all_hs = lay_danh_sach_hoc_sinh_db('admin', None, None)
                            # Tạo bộ từ điển: Key là (Tên, Lớp) -> Value là ID
                            hs_dict = {(str(hs[1]).strip().lower(), str(hs[2]).strip().lower()): hs[0] for hs in all_hs}
                            
                            s_count, f_count = 0, 0
                            with st.spinner("Đang dò tìm học sinh và gắn ảnh..."):
                                for _, row in df_p.iterrows():
                                    ten = str(row.get("Họ và Tên", "")).strip().lower()
                                    lop = str(row.get("Lớp", "")).strip().lower()
                                    ten_file_anh = str(row.get("Tên file ảnh", "")).strip()
                                    
                                    if not ten or not lop or not ten_file_anh or ten_file_anh == "nan": continue
                                    
                                    # Dò tìm ID học sinh dựa vào Tên và Lớp
                                    hs_id = hs_dict.get((ten, lop))
                                    if hs_id:
                                        # Gọi hàm CSDL cập nhật ảnh
                                        if cap_nhat_anh_the_db(hs_id, ten_file_anh): s_count += 1
                                        else: f_count += 1
                                    else: f_count += 1 # Không tìm thấy học sinh này
                                        
                            st.success(f"✅ Đã gắn ảnh thành công cho {s_count} học sinh. Bỏ qua: {f_count} học sinh (Do sai tên/lớp hoặc trống).")
                    except Exception as e: st.error(f"Lỗi: {e}")

        # ---------------------------------------
        # CỘT 3: DỌN DẸP / XÓA (GIỮ NGUYÊN NHƯ CŨ)
        # ---------------------------------------
        with col_delete:
            st.subheader("🗑️ 3. Dọn dẹp (Xóa)")
            classes, _ = get_distinct_classes_and_groups_db()
            if not classes: st.write("Chưa có lớp nào.")
            else:
                delete_mode = st.radio("Chọn phương thức xóa:", ["Xóa từng học sinh", "Xóa nguyên lớp"])
                if delete_mode == "Xóa nguyên lớp":
                    class_to_delete = st.selectbox("Chọn lớp muốn XÓA:", classes)
                    if st.checkbox("Xác nhận muốn xóa vĩnh viễn lớp này."):
                        if st.button(f"🚨 XÓA LỚP {class_to_delete}", type="primary", use_container_width=True):
                            hs_lop_do = lay_danh_sach_hoc_sinh_db('admin', class_to_delete, None)
                            ids_to_delete = [hs[0] for hs in hs_lop_do]
                            if ids_to_delete:
                                xoa_nhieu_hoc_sinh_db(ids_to_delete)
                                st.success("Đã xóa thành công."); st.rerun()

                elif delete_mode == "Xóa từng học sinh":
                    class_to_filter = st.selectbox("Lọc theo lớp:", ["Tất cả"] + classes)
                    c_filter = None if class_to_filter == "Tất cả" else class_to_filter
                    hs_list = lay_danh_sach_hoc_sinh_db('admin', c_filter, None)
                    if hs_list:
                        df_hs = pd.DataFrame(hs_list, columns=["ID", "Họ Tên", "Lớp", "Tổ", "SĐT", "Zalo", "Ảnh"])[["ID", "Họ Tên", "Lớp"]]
                        df_hs.insert(0, "Chọn xóa", False)
                        edited_df = st.data_editor(df_hs, hide_index=True, column_config={"Chọn xóa": st.column_config.CheckboxColumn(required=True)}, disabled=["ID", "Họ Tên", "Lớp"], use_container_width=True)
                        ids_to_delete = edited_df[edited_df["Chọn xóa"] == True]["ID"].tolist()
                        if len(ids_to_delete) > 0:
                            if st.button("🗑️ Xác nhận xóa", type="primary"):
                                xoa_nhieu_hoc_sinh_db(ids_to_delete)
                                st.success("Đã xóa."); st.rerun()
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
def show_main_dashboard():
    user = st.session_state.user_info
    role, aclass, agroup = user['role'], user['class'], user['group']
    
    # --- THANH BÊN (SIDEBAR) HIỆN ĐẠI & ÉP KHÔNG GIAN ---
    with st.sidebar:
        role_display = "TỔ TRƯỞNG" if role == 'bcs' else role.upper()
        group_display = f" | Tổ {agroup}" if agroup else ""
        
        # Bỏ st.title, Dùng HTML để ép lời chào và vai trò dính sát vào nhau
        st.markdown(f"""
            <div style='margin-top: -10px; margin-bottom: 10px;'>
                <h4 style='margin-bottom: 0px;'>Chào, {user['full_name']} 👋</h4>
                <p style='color: gray; font-size: 13.5px; margin-top: 2px; margin-bottom: 5px;'>
                    Vai trò: {role_display} | Lớp: {aclass or 'Toàn trường'}{group_display}
                </p>
            </div>
        """, unsafe_allow_html=True)
        
        options = ["Bảng điều khiển", "Bảng Vàng Thi Đua", "Quầy Đổi Thưởng", "Quản lý Lớp học", "Sơ đồ Chỗ ngồi", "Điểm danh hàng ngày", "Ghi nhận Nhanh"]
        icons = ["house", "trophy", "gift", "people", "grid-3x3", "calendar2-check", "lightning-charge"]
        
        if role in ['gvcn', 'admin', 'bcs']: 
            options.append("Tổng kết & Xuất Báo cáo")
            icons.append("file-earmark-spreadsheet")

        if role in ['gvcn', 'admin']:
            options.extend(["Ghi nhận Kỷ luật (TT19)", "Thống kê & Báo cáo"])
            icons.extend(["shield-exclamation", "bar-chart-steps"])
            
        if role == 'admin':
            options.append("Quản trị Hệ thống")
            icons.append("gear")

        # Tối ưu Menu: Chữ nhỏ lại một xíu, bỏ khoảng cách giữa các nút
        choice = option_menu(
            menu_title=None, 
            options=options,
            icons=icons,
            menu_icon="cast",
            default_index=0,
            styles={
                "container": {"padding": "0!important", "background-color": "transparent"},
                "icon": {"font-size": "15px"}, 
                "nav-link": {
                    "font-size": "13.5px",     # Thu nhỏ chữ
                    "text-align": "left", 
                    "margin": "0px 0px",       # Xóa hoàn toàn khoảng cách giữa 2 nút
                    "padding": "8px 10px",     
                    "border-radius": "6px", 
                    "--hover-color": "#e6f2ff"
                },
                "nav-link-selected": {"background-color": "#0056b3", "color": "white", "font-weight": "bold"},
            }
        )
        
        st.markdown("---")
        if st.button("Đăng xuất", type="secondary", width="stretch"):
            st.session_state.logged_in = False
            st.session_state.user_info = None
            st.rerun()
    
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
     # <<< THÊM 2 DÒNG NÀY VÀO ĐÂY >>>
    elif choice == "Quầy Đổi Thưởng":
        show_reward_store_page()
    elif choice == "Quản lý Lớp học":
        show_class_management()
        # <<< THÊM 2 DÒNG NÀY VÀO >>>
    elif choice == "Sơ đồ Chỗ ngồi":
        show_seating_chart_page()
    elif choice == "Điểm danh hàng ngày":
        show_attendance_page()
    elif choice == "Ghi nhận Nhanh":
        show_quick_record_page()
    elif choice == "Ghi nhận Kỷ luật (TT19)":
        show_discipline_page()
    elif choice == "Thống kê & Báo cáo":
        show_statistics_page()
    elif choice == "Tổng kết & Xuất Báo cáo":
        show_summary_page()
    elif choice == "Quản trị Hệ thống":
        show_admin_page()

# --- HÀM HỖ TRỢ: TẠO ẢNH BẰNG KHEN (TỐI ƯU CĂN CHỈNH & DẢI LỤA ĐỎ) ---
def create_certificate_image(student_name, class_name, achievement_text, photo_filename=None):
    from PIL import Image, ImageDraw, ImageFont, ImageOps
    import io
    import os
    
    # 1. Nền Trắng ngà sang trọng (1000x750)
    img = Image.new('RGB', (1000, 750), color='#FFFFF0')
    draw = ImageDraw.Draw(img)

    # 2. Khung viền Hoàng gia
    draw.rectangle([25, 25, 975, 725], outline='#001F3F', width=15)
    draw.rectangle([45, 45, 955, 705], outline='#DAA520', width=4)
    draw.rectangle([55, 55, 945, 695], outline='#DAA520', width=1)

    # 3. Nạp Font chữ
    try:
        font_title = ImageFont.truetype("fonts/timesbd.ttf", 65) # To hơn
        font_subtitle = ImageFont.truetype("fonts/times.ttf", 35)
        font_name = ImageFont.truetype("fonts/timesbd.ttf", 85)  # To hơn
        font_achievement = ImageFont.truetype("fonts/timesbd.ttf", 45)
    except:
        font_title = font_subtitle = font_name = font_achievement = ImageFont.load_default()

    # 4. Tiêu đề
    draw.text((500, 80), "BẢNG VÀNG VINH DANH", fill="#B22222", font=font_title, anchor="mt")

    # 5. XỬ LÝ GHÉP ẢNH HỌC SINH
    avatar_size = 200 
    avatar_x = 400    
    avatar_y = 170    # Đẩy ảnh lên cao một chút cho thoáng

    has_valid_photo = False
    if photo_filename and isinstance(photo_filename, str) and os.path.exists(os.path.join('student_photos', photo_filename)):
        try:
            avatar = Image.open(os.path.join('student_photos', photo_filename)).convert("RGBA")
            avatar = avatar.resize((avatar_size, avatar_size), Image.Resampling.LANCZOS)
            
            mask = Image.new('L', (avatar_size, avatar_size), 0)
            mask_draw = ImageDraw.Draw(mask)
            mask_draw.ellipse((0, 0, avatar_size, avatar_size), fill=255)
            
            output = ImageOps.fit(avatar, mask.size, centering=(0.5, 0.5))
            output.putalpha(mask)
            img.paste(output, (avatar_x, avatar_y), output)
            has_valid_photo = True
        except: pass

    if has_valid_photo:
        draw.ellipse([avatar_x - 6, avatar_y - 6, avatar_x + avatar_size + 6, avatar_y + avatar_size + 6], outline='#DAA520', width=8)
    else:
        # Nếu chưa có ảnh, vẽ một huy chương vàng tượng trưng
        draw.ellipse([avatar_x, avatar_y, avatar_x + avatar_size, avatar_y + avatar_size], fill='#FFF8DC', outline='#DAA520', width=5)
        draw.text((500, avatar_y + 100), "🏆", fill="#DAA520", font=font_title, anchor="mm")

    # 6. Thông tin học sinh
    draw.text((500, 400), "Tuyên dương học sinh:", fill="#555555", font=font_subtitle, anchor="mt")
    draw.text((500, 450), str(student_name).upper(), fill="#001F3F", font=font_name, anchor="mt")
    draw.text((500, 550), f"Học sinh lớp: {class_name}", fill="#333333", font=font_subtitle, anchor="mt")
    
    # 7. Dải lụa đỏ cho Thành tích (Tạo điểm nhấn mạnh mẽ)
    # Vẽ một hình chữ nhật bo góc giả làm dải lụa
    box_w = 700
    box_h = 70
    box_x = 500 - (box_w/2)
    box_y = 620
    draw.rounded_rectangle([box_x, box_y, box_x + box_w, box_y + box_h], radius=15, fill="#B22222")
    
    # Chữ thành tích màu vàng gold nằm trên dải lụa đỏ
    draw.text((500, 630), str(achievement_text), fill="#FFD700", font=font_achievement, anchor="mt")

    # Xuất file
    buf = io.BytesIO()
    img.save(buf, format='JPEG', quality=95)
    return buf.getvalue()
# --- HÀM 8: BẢNG VÀNG THI ĐUA ĐƯỢC NÂNG CẤP (CÓ CHỌN TUẦN) ---
def show_leaderboard_page():
    import streamlit as st
    import pandas as pd
    import json
    import os
    import base64
    from datetime import date, datetime, timedelta
    
    st.header("🏆 Bảng Vàng Thi Đua & Vinh Danh")
    st.markdown("---")

    user = st.session_state.user_info
    role, aclass, agroup = user['role'], user['class'], user['group']

    # --- 1. TÍNH TUẦN HIỆN TẠI LÀM MẶC ĐỊNH ---
    today = date.today()
    try:
        start_date_str = load_setting('school_year_start_date', '2025-09-08')
        h_json = json.loads(load_setting(HOLIDAY_SETTINGS_KEY, '[]'))
        h_in_year = [(datetime.strptime(s, '%Y-%m-%d').date(), datetime.strptime(e, '%Y-%m-%d').date()) for s, e in h_json]
        current_week_default = get_school_week_number(today, datetime.strptime(start_date_str, '%Y-%m-%d').date(), h_in_year)
        if current_week_default <= 0: current_week_default = 1
    except: current_week_default = 1
        
    # --- 2. BỘ LỌC CHỌN TUẦN ---
    col_w1, col_w2 = st.columns([1, 3])
    with col_w1:
        # Cho phép người dùng chọn tuần, mặc định là tuần hiện tại
        selected_week = st.number_input("Tùy chọn Tuần xem Bảng Vàng:", min_value=1, max_value=52, value=current_week_default, step=1)

    # --- 3. LẤY NGÀY THEO TUẦN ĐƯỢC CHỌN ---
    start_w, end_w = get_week_dates_by_number(selected_week)
    if not start_w: 
        st.error(f"Lỗi: Không thể tính toán ngày tháng cho Tuần {selected_week}.")
        return

    st.info(f"✨ Đang hiển thị thành tích **Tuần {selected_week}** (Từ {start_w.strftime('%d/%m')} đến {end_w.strftime('%d/%m')})")
    st.markdown("---")

    # --- 4. TRUY XUẤT DỮ LIỆU ---
    all_events = lay_su_kien_trong_khoang_ngay_db(start_w.strftime('%Y-%m-%d'), (end_w + timedelta(days=1)).strftime('%Y-%m-%d'), role, aclass, agroup)
    all_students = lay_thong_tin_day_du_hoc_sinh_db(user_role=role, assigned_class=aclass, assigned_group=agroup)
    
    events_by_student = {}
    for ev in all_events:
        events_by_student.setdefault(ev[0], []).append(ev)
        
    # ==========================================
    # THUẬT TOÁN TÍNH ĐIỂM VÀ TIÊU CHÍ PHỤ
    # ==========================================
    student_list = []
    for hs_tuple in all_students:
        try:
            hs = dict(zip(["id"] + STUDENT_FIELDS_DB, hs_tuple))
            hs_events = events_by_student.get(hs['id'], [])
            
            # 1. Tính tổng điểm
            score = DIEM_KHOI_DAU + sum(e[6] for e in hs_events)
            
            # 2. Tính các tiêu chí phụ (Tie-breakers)
            diem_cong = sum(e[6] for e in hs_events if e[6] > 0) # Tổng điểm cộng
            so_lan_vp = sum(1 for e in hs_events if e[5] == "Vi phạm") # Số lần mắc lỗi
            ten_chinh = hs['ten'].split()[-1].lower() if hs['ten'] else "" # Lấy tên để xếp A-Z
            
            student_list.append({
                'id': hs['id'], 'ten': hs['ten'], 'lop': hs['lop'], 'to': hs.get('to_nhiem_vu', '') or "Không rõ", 
                'diem': score, 'anh_the': hs.get('anh_the_path', ''),
                'diem_cong': diem_cong, 'so_lan_vp': so_lan_vp, 'ten_chinh': ten_chinh
            })
        except Exception as e:
            continue
    
    # 3. Sắp xếp thông minh đa tiêu chí:
    # Dùng dấu (-) cho các tiêu chí muốn xếp giảm dần (Cao xếp trước)
    # Dùng dấu (+) cho các tiêu chí muốn xếp tăng dần (Thấp/Ít xếp trước)
    student_list.sort(key=lambda x: (
        -x['diem'],         # Ưu tiên 1: Điểm tổng cao nhất
        -x['diem_cong'],    # Ưu tiên 2: Điểm cộng nhiều nhất
        x['so_lan_vp'],     # Ưu tiên 3: Số lần vi phạm ÍT nhất
        x['ten_chinh']      # Ưu tiên 4: Tên theo vần A-Z
    ))

    if not student_list:
        st.warning(f"Tuần {selected_week} chưa có dữ liệu rèn luyện nào để xếp hạng.")
        return

    # === KHU VỰC 1: BỤC VINH QUANG TOP 3 ===
    st.subheader(f"🌟 TOP 3 HỌC SINH XUẤT SẮC NHẤT TUẦN {selected_week}")
    
    top3 = student_list[:3]
    col2, col1, col3 = st.columns([1, 1.2, 1], gap="medium")
    
    def get_image_base64(path):
        if path and isinstance(path, str) and os.path.exists(os.path.join('student_photos', path)):
            try:
                with open(os.path.join('student_photos', path), "rb") as img_file:
                    encoded = base64.b64encode(img_file.read()).decode()
                    return f"data:image/jpeg;base64,{encoded}"
            except: pass
        return "https://cdn-icons-png.flaticon.com/512/149/149071.png" 

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
                # Cập nhật chữ trong Bằng khen theo tuần được chọn
                cert_img = create_certificate_image(hs['ten'], hs['lop'], f"Đạt Top {rank} Xuất sắc Tuần {selected_week}", hs['anh_the'])
                st.download_button("📥 Tải Bằng Khen", data=cert_img, file_name=f"BangKhen_T{selected_week}_Top{rank}_{hs['ten']}.jpg", mime="image/jpeg", width="stretch", key=f"btn_cert_{hs['id']}_{selected_week}")
            except Exception as e:
                st.error("Lỗi tạo bằng khen")

    if len(top3) > 0: 
        draw_podium(col1, top3[0], 1, "🥇", "#FFF9E6") 
        # Chỉ bắn bóng bay nếu đang xem tuần hiện tại
        if selected_week == current_week_default:
            st.balloons() 
    if len(top3) > 1: draw_podium(col2, top3[1], 2, "🥈", "#F2F2F2")
    if len(top3) > 2: draw_podium(col3, top3[2], 3, "🥉", "#FFF0E6")

    st.markdown("<br><hr>", unsafe_allow_html=True)

    # === KHU VỰC 2: BẢNG XẾP HẠNG TỔ VÀ TOP 4-10 ===
    col_t_left, col_t_right = st.columns([1, 1.2], gap="large")
    
    with col_t_left:
        st.subheader("👥 Xếp Hạng Tổ")
        df_hs = pd.DataFrame(student_list)
        if not df_hs.empty:
            team_ranking = df_hs.groupby('to')['diem'].mean().reset_index()
            team_ranking.rename(columns={'to': 'Tên Tổ', 'diem': 'Điểm Trung Bình'}, inplace=True)
            team_ranking.sort_values(by='Điểm Trung Bình', ascending=False, inplace=True)
            
            html_teams = "<div style='display: flex; flex-direction: column; gap: 12px; margin-top: 10px;'>"
            for i, (_, row) in enumerate(team_ranking.iterrows()):
                rank = i + 1
                ten_to = row['Tên Tổ']
                diem_tb = row['Điểm Trung Bình']
                
                if rank == 1:
                    bg, color, icon, border = "linear-gradient(135deg, #FFD700 0%, #FDB931 100%)", "#8B6508", "🏆", "border: 2px solid #DAA520; box-shadow: 0 4px 15px rgba(218,165,32,0.4);"
                elif rank == 2:
                    bg, color, icon, border = "linear-gradient(135deg, #E0E0E0 0%, #BDBDBD 100%)", "#424242", "🥈", "box-shadow: 0 4px 10px rgba(0,0,0,0.1);"
                elif rank == 3:
                    bg, color, icon, border = "linear-gradient(135deg, #F4A460 0%, #CD853F 100%)", "#5C3A21", "🥉", "box-shadow: 0 4px 10px rgba(0,0,0,0.1);"
                else:
                    bg, color, icon, border = "#ffffff", "#333333", "🏅", "border-left: 5px solid #bdc3c7; border-top: 1px solid #eee; border-bottom: 1px solid #eee; border-right: 1px solid #eee;"

                html_teams += f"<div style='background: {bg}; padding: 15px 20px; border-radius: 12px; {border} display: flex; justify-content: space-between; align-items: center; transition: transform 0.2s;'><div style='font-size: 18px; font-weight: bold; color: {color};'><span style='font-size: 24px; vertical-align: middle;'>{icon}</span> Hạng {rank}: Tổ {ten_to}</div><div style='font-size: 22px; font-weight: 900; color: {color};'>{diem_tb:.1f} <span style='font-size: 14px; font-weight: normal;'>điểm</span></div></div>"
                
            html_teams += "</div>"
            st.markdown(html_teams, unsafe_allow_html=True)

    with col_t_right:
        st.subheader("📜 Bảng Danh Dự (Top 4 - Top 10)")
        if len(student_list) > 3:
            top_rest = student_list[3:10]
            html_list = ""
            for i, hs in enumerate(top_rest):
                rank = i + 4
                bg_color = "#f8f9fa" if i % 2 == 0 else "#ffffff"
                html_list += f"<div style='background-color: {bg_color}; padding: 10px 15px; border-radius: 8px; margin-bottom: 8px; border-left: 4px solid #3498db; display: flex; justify-content: space-between; align-items: center;'><div style='font-size: 16px;'><strong style='color: #2980b9;'>Hạng {rank} 🏅</strong> &nbsp; | &nbsp; <b>{hs['ten']}</b> <span style='color: gray; font-size: 14px;'>(Tổ {hs['to']})</span></div><div style='font-size: 18px; font-weight: bold; color: #d63031;'>{hs['diem']}đ</div></div>"
            st.markdown(html_list, unsafe_allow_html=True)

    # === KHU VỰC 3: NHẮC NHỞ HỌC SINH CHƯA CÓ ĐIỂM CỘNG ===
    st.markdown("<br>", unsafe_allow_html=True)
    khen_thuong_ids = set([ev[0] for ev in all_events if ev[5] == "Khen thưởng" and ev[6] > 0])
    hs_chua_co_diem = [hs['ten'] for hs in student_list if hs['id'] not in khen_thuong_ids]
    
    if len(hs_chua_co_diem) > 0:
        with st.expander(f"⚠️ Xem danh sách {len(hs_chua_co_diem)} học sinh CHƯA có điểm cộng trong tuần {selected_week}:", expanded=False):
            st.info("💡 Các em có tên dưới đây tuần này chưa hăng hái phát biểu hoặc làm việc tốt. Cần cố gắng hơn!")
            cols_no_bonus = st.columns(3)
            for i, ten_hs in enumerate(hs_chua_co_diem):
                cols_no_bonus[i % 3].markdown(f"🔹 {ten_hs}")
    else:
        st.success(f"✨ Thật tuyệt vời! 100% học sinh đều có điểm cộng tích cực trong tuần {selected_week}!")
# --- HÀM 10: SIÊU THỊ ĐỔI THƯỞNG (GAMIFICATION ĐA DẠNG) ---
def show_reward_store_page():
    st.header("🎁 Siêu thị Đổi Thưởng (Reward Store)")
    st.markdown("---")

    user = st.session_state.user_info
    role, aclass, agroup = user['role'], user['class'], user['group']

    # ==========================================
    # DANH MỤC HÀNG HÓA SIÊU THỊ (Chia làm 3 Hạng)
    # ==========================================
    SUPERMARKET = {
        "🟢 Khu Vực Phổ Thông (Dễ đạt được)": [
            {"name": "Được quyền chọn bài hát giờ ra chơi", "cost": 15, "icon": "🎵", "color": "#e8f5e9"},
            {"name": "Gia hạn nộp Bài tập về nhà thêm 1 ngày", "cost": 25, "icon": "⏳", "color": "#e8f5e9"},
            {"name": "Thẻ miễn truy bài miệng (1 lần)", "cost": 35, "icon": "🛡️", "color": "#e8f5e9"},
        ],
        "🔵 Khu Vực Đặc Quyền (Tích lũy trung bình)": [
            {"name": "Miễn trực nhật vệ sinh (1 buổi)", "cost": 60, "icon": "🧹", "color": "#e3f2fd"},
            {"name": "Được quyền tự chọn chỗ ngồi (1 tuần)", "cost": 80, "icon": "🪑", "color": "#e3f2fd"},
            {"name": "Một phần quà vặt nhỏ từ Thầy/Cô", "cost": 100, "icon": "🍬", "color": "#e3f2fd"},
        ],
        "👑 Khu Vực VIP (Mục tiêu dài hạn)": [
            {"name": "Thư khen ngợi đặc biệt gửi về Phụ huynh", "cost": 150, "icon": "💌", "color": "#fff8e1"},
            {"name": "Voucher 1 ly Trà sữa từ Thầy/Cô", "cost": 250, "icon": "🧋", "color": "#f3e5f5"},
            {"name": "Thẻ Kim Bài: Trải nghiệm làm Lớp Trưởng 1 ngày", "cost": 400, "icon": "👑", "color": "#ffebee"},
        ]
    }

    tab_doiqua, tab_phatxu = st.tabs(["🛒 Siêu thị Mua sắm", "💰 Phát Xu Thưởng (Cho GVCN)"])

    # ==========================================
    # TAB 1: ĐỔI QUÀ (HỌC SINH / BAN CÁN SỰ THAO TÁC)
    # ==========================================
    with tab_doiqua:
        raw_students = get_cached_students(role, aclass, agroup)
        if not raw_students: st.warning("Không có dữ liệu học sinh."); return
        
        student_dict = {f"{hs[1]} (Tổ {hs[3] or '?'})": hs[0] for hs in raw_students}
        
        col_chon_hs, col_so_du = st.columns([2, 1])
        with col_chon_hs:
            st.subheader("1. Chọn Học sinh muốn mua đồ")
            selected_student = st.selectbox("Tên khách hàng:", ["-- Chọn --"] + list(student_dict.keys()), label_visibility="collapsed")

        if selected_student != "-- Chọn --":
            hs_id = student_dict[selected_student]
            so_du_hien_tai = lay_so_du_xu_db(hs_id)
            hs_ten_ngan = selected_student.split("(")[0].strip()

            # Hiển thị số dư
            with col_so_du:
                st.markdown(f"""
                    <div style='text-align: center; padding: 10px; background-color: #fff; border-radius: 10px; border: 2px solid #f1c40f; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-top: -10px;'>
                        <p style='margin: 0; color: #555; font-size: 14px;'>Ví của {hs_ten_ngan}</p>
                        <h2 style='margin: 0; color: #f39c12;'>🪙 {so_du_hien_tai} Xu</h2>
                    </div>
                """, unsafe_allow_html=True)
            
            st.markdown("---")
            st.subheader("2. Gian hàng Quà tặng")
            
            # Quét từng khu vực trong siêu thị để vẽ giao diện
            for category_name, items in SUPERMARKET.items():
                st.markdown(f"#### {category_name}")
                
                # Chia làm 3 cột để xếp đồ cho đẹp
                cols = st.columns(3)
                for i, item in enumerate(items):
                    with cols[i % 3]:
                        # Vẽ Thẻ sản phẩm
                        st.markdown(f"""
                            <div style='background-color: {item['color']}; padding: 15px; border-radius: 12px; text-align: center; border: 1px solid #ddd; height: 180px; display: flex; flex-direction: column; justify-content: center; box-shadow: 0 2px 4px rgba(0,0,0,0.05); margin-bottom: 10px;'>
                                <h1 style='margin:0; font-size: 40px;'>{item['icon']}</h1>
                                <p style='margin:10px 0; color: #333; font-weight: bold; font-size: 15px; line-height: 1.2;'>{item['name']}</p>
                                <h4 style='color: #d35400; margin:0;'>🪙 {item['cost']}</h4>
                            </div>
                        """, unsafe_allow_html=True)
                        
                        # Nút Mua hàng
                        if so_du_hien_tai >= item['cost']:
                            # Dùng key động tạo từ tên món quà để chống lỗi trùng lặp
                            if st.button(f"🛒 Mua ngay", key=f"btn_buy_{hs_id}_{item['name']}", type="primary", use_container_width=True):
                                # 1. Trừ xu
                                cap_nhat_xu_thuong_db(hs_id, -item['cost'])
                                # 2. Ghi vào lịch sử rèn luyện
                                ngay_doi = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                                them_su_kien_ren_luyen_db(hs_id, f"🛍️ Đã mua: {item['name']} (-{item['cost']} Xu)", "Khen thưởng", 0, ngay_doi)
                                
                                st.success(f"🎉 Giao dịch thành công! {hs_ten_ngan} đã mua: {item['name']}!")
                                st.balloons()
                                st.rerun() # Tải lại trang để update ví
                        else:
                            st.button(f"🔒 Cần thêm {item['cost'] - so_du_hien_tai} Xu", disabled=True, key=f"btn_lock_{hs_id}_{item['name']}", use_container_width=True)
                
                st.write("<br>", unsafe_allow_html=True) # Khoảng cách giữa các khu vực

    # ==========================================
    # TAB 2: PHÁT XU THƯỞNG (GVCN THAO TÁC CUỐI TUẦN)
    # ==========================================
    with tab_phatxu:
        if role not in ['gvcn', 'admin']:
            st.error("Chức năng này chỉ dành cho Giáo viên chủ nhiệm.")
            return
            
        st.info("Hàng tuần, sau khi xem Tổng kết, GVCN vào đây để thưởng Xu tự động cho các em đạt loại Tốt và Xuất sắc.")
        
        col_px1, col_px2 = st.columns([1, 2])
        with col_px1:
            week_num = st.number_input("Chọn tuần để quét:", min_value=1, max_value=52, value=1, step=1)
            
        st.markdown("Quy tắc thưởng: **Xuất sắc (≥ 115đ) = +20 Xu** | **Tốt (100 - 114đ) = +10 Xu**")
        
        if st.button("🔍 Quét kết quả & Phát Xu hàng loạt", type="primary", width="stretch"):
            start_w, end_w = get_week_dates_by_number(week_num)
            if not start_w: st.error("Lỗi ngày tháng."); return
            
            with st.spinner("Đang tính điểm xếp loại rèn luyện tuần..."):
                all_events = lay_su_kien_trong_khoang_ngay_db(start_w.strftime('%Y-%m-%d'), (end_w + timedelta(days=1)).strftime('%Y-%m-%d'), role, aclass, agroup)
                
                events_by_student = {}
                for ev in all_events: events_by_student.setdefault(ev[0], []).append(ev)
                
                phat_xu_count = 0
                for hs in raw_students:
                    hs_id = hs[0]
                    score = DIEM_KHOI_DAU + sum(e[6] for e in events_by_student.get(hs_id, []))
                    xep_loai = xep_loai_hanh_kiem(max(0, score)) 
                    
                    xu_thuong = 0
                    if xep_loai == "Xuất sắc": xu_thuong = 20
                    elif xep_loai == "Tốt": xu_thuong = 10
                    
                    if xu_thuong > 0:
                        cap_nhat_xu_thuong_db(hs_id, xu_thuong)
                        ngay_tao = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        them_su_kien_ren_luyen_db(hs_id, f"🪙 Thưởng Xu Tuần {week_num} (Xếp loại Rèn luyện {xep_loai})", "Khen thưởng", 0, ngay_tao)
                        phat_xu_count += 1
                        
            st.success(f"✅ Đã phát Xu thưởng thành công cho **{phat_xu_count}** học sinh đạt loại Tốt/Xuất sắc trong Tuần {week_num}!")
# --- HÀM 11: SƠ ĐỒ LỚP HỌC (SMART SEATING ALGORITHM) ---
def show_seating_chart_page():
    # --- THÊM NÚT LÀM MỚI BỘ NHỚ ĐỆM ---
    col_hd1, col_hd2 = st.columns([3, 1])
    with col_hd1:
        st.header("🪑 Sơ đồ Chỗ ngồi & Thuật toán Tối ưu")
    with col_hd2:
        st.write("")
        if st.button("🔄 Làm mới dữ liệu", use_container_width=True, help="Bấm vào đây nếu thầy vừa thêm/xóa học sinh mà sơ đồ chưa cập nhật"):
            st.cache_data.clear() # Xóa toàn bộ bộ nhớ tạm
            st.rerun()

    st.markdown("---")

    user = st.session_state.user_info
    role, aclass, agroup = user['role'], user['class'], user['group']

    if role not in ['gvcn', 'admin']:
        st.error("Chức năng này chỉ dành cho Giáo viên chủ nhiệm.")
        return

    # 1. Lấy dữ liệu
    raw_students = get_cached_students(role, aclass, agroup)
    
    # <<< ĐIỂM SỬA LỖI: BỎ LỆNH return Ở ĐÂY ĐỂ VẪN VẼ ĐƯỢC SƠ ĐỒ TRỐNG >>>
    if not raw_students:
        st.warning(f"Lớp {aclass or 'của thầy/cô'} hiện chưa có dữ liệu học sinh. Sơ đồ sẽ hiển thị các ghế trống.")

    today = date.today()
    start_30d = today - timedelta(days=30)
    events_30d = lay_su_kien_trong_khoang_ngay_db(start_30d.strftime('%Y-%m-%d'), (today + timedelta(days=1)).strftime('%Y-%m-%d'), role, aclass, agroup)
    
    events_by_student = {}
    for ev in events_30d: events_by_student.setdefault(ev[0], []).append(ev)

    hs_data = []
    to_set = set()
    if raw_students:
        for hs in raw_students:
            hs_id, ten, lop, to, _, _, anh_the = hs
            to_name = str(to).strip() if to else "Khác"
            to_set.add(to_name)
            
            my_events = events_by_student.get(hs_id, [])
            diem = DIEM_KHOI_DAU + sum(e[6] for e in my_events)
            
            loi_mat_trat_tu = sum(1 for e in my_events if "Mất trật tự" in e[4] or "Nói chuyện" in e[4])
            ten_ngan = ten.split()[-1] + " " + ten.split()[0] if len(ten.split()) > 1 else ten
            
            hs_data.append({
                "id": hs_id, "ten_goc": ten, "ten_ngan": ten_ngan, "to": to_name,
                "diem": diem, "anh_the": anh_the, "mat_trat_tu": loi_mat_trat_tu
            })

    default_day = len(to_set) if len(to_set) > 0 else 4
    col_set1, col_set2, col_set3 = st.columns([1, 1, 2])
    with col_set1:
        so_day = st.number_input("Số Dãy (Tương ứng với Tổ):", min_value=1, max_value=6, value=default_day)
    with col_set2:
        so_ban = st.number_input("Số Bàn mỗi dãy (Ngồi 2 em/bàn):", min_value=1, max_value=10, value=5)

    st.markdown("---")

    # --- NÚT GỌI THUẬT TOÁN SẮP XẾP ---
    col_ai1, col_ai2 = st.columns([1, 2.5])
    with col_ai1:
        btn_ai_sort = st.button("🚀 Tự động Xếp Chỗ (Smart Logic)", type="primary", width="stretch")
    with col_ai2:
        st.info("💡 **Quy tắc:** 1 Dãy = 1 Tổ | Xếp bạn điểm thấp lên bàn đầu | Tách 2 bạn hay nói chuyện khỏi 1 bàn.")

    setting_key = f"seating_plan_v2_{aclass}"
    
    if btn_ai_sort:
        if not raw_students:
            st.error("Chưa có học sinh nào để xếp chỗ! Thầy vui lòng kiểm tra lại danh sách lớp.")
        else:
            with st.spinner("Đang chạy thuật toán sư phạm sắp xếp bàn ghế..."):
                new_seating_plan = {}
                to_groups = {}
                for hs in hs_data:
                    to_groups.setdefault(hs['to'], []).append(hs)
                
                sorted_to_names = sorted(to_groups.keys())
                
                for d_idx, to_name in enumerate(sorted_to_names):
                    if d_idx >= so_day: break
                    day_num = d_idx + 1
                    students = to_groups[to_name]
                    students.sort(key=lambda x: x['diem'])
                    
                    talkers = [s for s in students if s['mat_trat_tu'] > 0]
                    quiet = [s for s in students if s['mat_trat_tu'] == 0]
                    
                    for ban_idx in range(1, so_ban + 1):
                        hs_left, hs_right = None, None
                        if talkers: hs_left = talkers.pop(0)
                        elif quiet: hs_left = quiet.pop(0)
                        if quiet: hs_right = quiet.pop(0)
                        elif talkers: hs_right = talkers.pop(0)
                        
                        if hs_left: new_seating_plan[str(hs_left['id'])] = f"D{day_num}-B{ban_idx}-L"
                        if hs_right: new_seating_plan[str(hs_right['id'])] = f"D{day_num}-B{ban_idx}-R"
                        
                save_setting(setting_key, json.dumps(new_seating_plan))
                st.toast("Đã xếp chỗ hoàn tất!", icon="✅")
                st.rerun()
# ====================================================================
    # KHU VỰC NÚT IN SƠ ĐỒ LỚP (BẢN KHẮC PHỤC BẢO MẬT STREAMLIT)
    # ====================================================================
    # 1. CSS ẩn các menu khi in ra giấy
    st.markdown("""
        <style>
            @media print {
                [data-testid="stSidebar"], 
                [data-testid="stHeader"],
                .stApp > header,
                .stButton, 
                div:has(> button), 
                iframe { /* Ẩn cái nút in đi khi ra giấy */
                    display: none !important; 
                }
                .block-container {
                    max-width: 100% !important;
                    padding: 0 !important;
                    margin: 0 !important;
                }
                * {
                    -webkit-print-color-adjust: exact !important;
                    print-color-adjust: exact !important;
                    color-adjust: exact !important;
                }
            }
        </style>
    """, unsafe_allow_html=True)
    
    # 2. Nút in chứa mã JavaScript chạy trong môi trường an toàn (Iframe)
    import streamlit.components.v1 as components
    components.html("""
        <div style="display: flex; justify-content: flex-end; margin-bottom: 5px;">
            <button onclick="window.parent.print()" style="background-color: #0078D7; color: white; border: none; padding: 10px 20px; border-radius: 8px; font-weight: bold; font-size: 15px; cursor: pointer; box-shadow: 0 4px 6px rgba(0,0,0,0.1); font-family: Arial, sans-serif;">
                🖨️ In Sơ Đồ Lớp
            </button>
        </div>
    """, height=50)
    # ====================================================================

 
    # --- HIỂN THỊ SƠ ĐỒ LỚP ---
    st.markdown("<br><h3 style='text-align: center; color: #2c3e50; letter-spacing: 2px;'>BẢNG ĐEN / BỤC GIẢNG</h3>", unsafe_allow_html=True)
    st.markdown("<div style='height: 6px; background: linear-gradient(90deg, #bdc3c7 0%, #2c3e50 50%, #bdc3c7 100%); border-radius: 3px; margin-bottom: 40px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'></div>", unsafe_allow_html=True)

    try: current_plan = json.loads(load_setting(setting_key, '{}'))
    except: current_plan = {}

    import base64
    import os
    def get_img_b64(path):
        if path and os.path.exists(os.path.join('student_photos', path)):
            with open(os.path.join('student_photos', path), "rb") as f: return f"data:image/jpeg;base64,{base64.b64encode(f.read()).decode()}"
        return "https://cdn-icons-png.flaticon.com/512/149/149071.png"

    def create_seat_html(hs_id):
        if not hs_id:
            return "<div style='width: 48%; background:#f8f9fa; border: 2px dashed #dfe6e9; border-radius:10px; text-align:center; color:#b2bec3; height:130px; display:flex; align-items:center; justify-content:center; box-sizing: border-box;'>Trống</div>"
        hs_info = next((item for item in hs_data if item["id"] == hs_id), None)
        if not hs_info: return ""
        img_b64 = get_img_b64(hs_info['anh_the'])
        bg_color = "#f0fff4" if hs_info['diem'] >= 115 else "#ffffff" if hs_info['diem'] >= 80 else "#fff5f5"
        accent_color = "#27ae60" if hs_info['diem'] >= 115 else "#bdc3c7" if hs_info['diem'] >= 80 else "#e74c3c"
        icon = "🗣️" if hs_info['mat_trat_tu'] > 0 else ""
        return f"<div style='width: 48%; background:{bg_color}; border:1px solid #eee; border-bottom: 4px solid {accent_color}; border-radius:10px; padding:8px 4px; text-align:center; height:130px; display:flex; flex-direction:column; justify-content:center; align-items:center; box-shadow: 0 2px 5px rgba(0,0,0,0.04); box-sizing: border-box;'><img src='{img_b64}' style='width:52px; height:52px; object-fit:cover; border-radius:50%; border: 2px solid white; box-shadow: 0 1px 3px rgba(0,0,0,0.1); margin-bottom: 6px;'><div style='font-size:11.5px; font-weight:bold; color:#2d3436; width:100%; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; line-height:1.2;' title='{hs_info['ten_goc']}'>{hs_info['ten_ngan']}</div><div style='font-size:10.5px; color:#636e72; font-weight:600; background:#f1f2f6; padding:2px 6px; border-radius:8px; margin-top:4px;'>{hs_info['diem']}đ {icon}</div></div>"

    to_names_sorted = sorted(list(set([hs['to'] for hs in hs_data])))
    
    cols = st.columns(so_day, gap="medium")
    for c in range(1, so_day + 1):
        with cols[c-1]:
            to_hien_thi = to_names_sorted[c-1] if c-1 < len(to_names_sorted) else f"Dãy {c}"
            col_html = f"<div style='background-color: #f4f6f9; border-radius: 16px; padding: 12px; border: 1px solid #e2e8f0; height: 100%; box-sizing: border-box;'><div style='text-align:center; background: linear-gradient(135deg, #0056b3 0%, #0078D7 100%); color:white; padding:10px; border-radius:8px; margin-bottom:15px; font-size:15px; font-weight:bold; box-shadow: 0 4px 10px rgba(0,86,179,0.3);'>Tổ {to_hien_thi}</div>"
            
            for r in range(1, so_ban + 1):
                seat_l_id, seat_r_id = None, None
                for hid_str, scode in current_plan.items():
                    if scode == f"D{c}-B{r}-L": seat_l_id = int(hid_str)
                    if scode == f"D{c}-B{r}-R": seat_r_id = int(hid_str)
                col_html += f"<div style='background-color: #ffffff; padding: 10px; border-radius: 12px; margin-bottom: 15px; box-shadow: 0 2px 8px rgba(0,0,0,0.04); border: 1px solid #edf2f7; box-sizing: border-box;'><div style='text-align:center; font-size:11px; color:#b2bec3; font-weight:800; margin-bottom:8px; text-transform: uppercase; letter-spacing: 1px;'>BÀN {r}</div><div style='display:flex; justify-content:space-between; align-items:center; width:100%;'>{create_seat_html(seat_l_id)}{create_seat_html(seat_r_id)}</div></div>"
            
            col_html += "</div>"
            st.markdown(col_html, unsafe_allow_html=True)
# --- ĐIỀU HƯỚNG ---
if not st.session_state.logged_in:
    show_login_page()
else:
    show_main_dashboard()

