# File: database.py
import streamlit as st
import psycopg2
from psycopg2 import pool
import hashlib
import pandas as pd
from contextlib import contextmanager
from datetime import datetime, timedelta
import json 
from config import HOLIDAY_SETTINGS_KEY
from sqlalchemy import create_engine

# --- BẢO MẬT CHUỖI KẾT NỐI CSDL ---
try:
    # Khi chạy trên Web, hệ thống sẽ tự tìm mật khẩu được giấu kín (Secrets)
    DB_CONNECTION_STRING = st.secrets["DB_CONNECTION_STRING"]
except (FileNotFoundError, KeyError):
    # Nếu chạy ở máy tính ở nhà (chưa có secrets), nó sẽ dùng chuỗi tạm này
    DB_CONNECTION_STRING = "postgresql://neondb_owner:npg_qulkd1b2QcFv@ep-snowy-voice-a1yf4u9r-pooler.ap-southeast-1.aws.neon.tech/neondb?sslmode=require&channel_binding=require"
# --- CÁC HẰNG SỐ ---
STUDENT_FIELDS_DB = [
    "ten", "lop", "to_nhiem_vu", "ngay_sinh", "dia_chi",
    "sdt_hoc_sinh", "sdt_zalo", "anh_the_path",
    "ten_cha", "nghe_nghiep_cha", "nam_sinh_cha", "sdt_cha",
    "ten_me", "nghe_nghiep_me", "nam_sinh_me", "sdt_me", "ghi_chu"
]

# --- CONNECTION POOL SETUP ---
try:
    connection_pool = psycopg2.pool.SimpleConnectionPool(1, 10, dsn=DB_CONNECTION_STRING)
    print("Khởi tạo Connection Pool thành công!")
except psycopg2.OperationalError as e:
    print("Lỗi Khởi tạo CSDL", f"Không thể tạo Connection Pool.\nKiểm tra chuỗi kết nối và mạng Internet.\n\nLỗi: {e}")
    connection_pool = None

@contextmanager
def get_db_connection():
    if not connection_pool:
        yield None
        return
    
    conn = None
    try:
        conn = connection_pool.getconn()
        yield conn
    finally:
        if conn:
            connection_pool.putconn(conn)

def init_db():
    try:
        with get_db_connection() as conn:
            if conn:
                print("Lấy kết nối từ Pool thành công!")
            else:
                print("Lấy kết nối từ Pool thất bại.")
    except ConnectionError as e:
        print(f"Lỗi khi kiểm tra kết nối: {e}")

def save_setting(key, value):
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    sql = "INSERT INTO settings (key, value) VALUES (%s, %s) ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value"
                    cursor.execute(sql, (key, value))
                conn.commit()
            else:
                print("Không có kết nối CSDL, không thể lưu cài đặt.")
                print("Lỗi CSDL", "Không có kết nối CSDL, không thể lưu cài đặt.")
    except (psycopg2.Error, ConnectionError) as e:
        print("Lỗi CSDL", f"Không thể lưu cài đặt: {e}")

def load_setting(key, default_value=None):
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    cursor.execute("SELECT value FROM settings WHERE key = %s", (key,))
                    result = cursor.fetchone()
                    return result[0] if result else default_value
            else:
                print("Không có kết nối CSDL, không thể tải cài đặt.")
                return default_value
    except (psycopg2.Error, ConnectionError) as e:
        print("Lỗi CSDL", f"Không thể tải cài đặt: {e}")
        return default_value

def get_week_dates_by_number(week_number):
    start_date_str = load_setting('school_year_start_date', '2025-09-08')
    holiday_periods_json = load_setting(HOLIDAY_SETTINGS_KEY, '[]')
    
    try:
        school_start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
        holiday_periods = json.loads(holiday_periods_json)
        # Chuyển đổi các chuỗi ngày tháng trong holiday_periods thành đối tượng datetime
        holiday_periods_dt = []
        for period in holiday_periods:
            start_dt = datetime.strptime(period[0], '%Y-%m-%d')
            end_dt = datetime.strptime(period[1], '%Y-%m-%d')
            holiday_periods_dt.append((start_dt, end_dt))
    except (ValueError, TypeError, json.JSONDecodeError):
        return None, None
        
    current_date, learning_week_count = school_start_date, 0
    
    for _ in range(100): # Lặp tối đa 100 tuần để tránh vòng lặp vô hạn
        is_holiday_week = False
        for start_holiday, end_holiday in holiday_periods_dt:
            if start_holiday <= current_date < end_holiday + timedelta(days=1):
                is_holiday_week = True
                break
        
        if is_holiday_week:
            current_date += timedelta(weeks=1)
            continue
        
        learning_week_count += 1
        if learning_week_count == week_number:
            return current_date, current_date + timedelta(days=5)
        
        current_date += timedelta(weeks=1)
        
    return None, None


def verify_user(username, password):
    password_hash = hashlib.sha256(password.encode()).hexdigest()
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    cursor.execute("SELECT role, assigned_class, assigned_group, full_name FROM users WHERE username = %s AND password_hash = %s", (username, password_hash))
                    user_data = cursor.fetchone()
                    if user_data:
                        return {
                            "username": username, 
                            "role": user_data[0], 
                            "class": user_data[1], 
                            "group": user_data[2],
                            "full_name": user_data[3] or username
                        }
            else:
                return None
    except (psycopg2.Error, ConnectionError) as e:
        print("Lỗi CSDL", f"Lỗi xác thực người dùng: {e}")
    return None

def get_all_users_db():
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    cursor.execute("SELECT id, username, full_name, role, assigned_class, assigned_group FROM users ORDER BY role, username")
                    return cursor.fetchall()
            else:
                return []
    except (psycopg2.Error, ConnectionError) as e:
        print("Lỗi CSDL", f"Lỗi lấy danh sách người dùng: {e}")
        return []

def add_user_db(username, full_name, password, role, assigned_class, assigned_group):
    if not password:
        return False, "Mật khẩu không được để trống."
    password_hash = hashlib.sha256(password.encode()).hexdigest()
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    cursor.execute("INSERT INTO users (username, full_name, password_hash, role, assigned_class, assigned_group) VALUES (%s, %s, %s, %s, %s, %s)",
                                   (username, full_name, password_hash, role, assigned_class, assigned_group))
                conn.commit()
                return True, "Thêm người dùng thành công."
            else:
                return False, "Không có kết nối CSDL."
    except psycopg2.IntegrityError:
        return False, f"Tên đăng nhập '{username}' đã tồn tại."
    except (psycopg2.Error, ConnectionError) as e:
        return False, f"Lỗi CSDL: {e}"

def update_user_db(user_id, username, full_name, role, assigned_class, assigned_group, new_password):
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    if new_password:
                        password_hash = hashlib.sha256(new_password.encode()).hexdigest()
                        cursor.execute("UPDATE users SET username=%s, full_name=%s, role=%s, assigned_class=%s, assigned_group=%s, password_hash=%s WHERE id=%s",
                                       (username, full_name, role, assigned_class, assigned_group, password_hash, user_id))
                    else:
                        cursor.execute("UPDATE users SET username=%s, full_name=%s, role=%s, assigned_class=%s, assigned_group=%s WHERE id=%s",
                                       (username, full_name, role, assigned_class, assigned_group, user_id))
                conn.commit()
                return True, "Cập nhật thành công."
            else:
                return False, "Không có kết nối CSDL."
    except psycopg2.IntegrityError:
        return False, f"Tên đăng nhập '{username}' có thể đã tồn tại."
    except (psycopg2.Error, ConnectionError) as e:
        return False, f"Lỗi CSDL: {e}"

def delete_user_db(user_id):
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    cursor.execute("DELETE FROM users WHERE id=%s", (user_id,))
                conn.commit()
                return True
            else:
                return False
    except (psycopg2.Error, ConnectionError) as e:
        print("Lỗi CSDL", f"Lỗi khi xóa người dùng: {e}")
        return False

def get_distinct_classes_and_groups_db():
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    cursor.execute("SELECT DISTINCT lop FROM hoc_sinh WHERE lop IS NOT NULL ORDER BY lop")
                    classes = [row[0] for row in cursor.fetchall()]
                    cursor.execute("SELECT DISTINCT to_nhiem_vu FROM hoc_sinh WHERE to_nhiem_vu IS NOT NULL ORDER BY to_nhiem_vu")
                    groups = [row[0] for row in cursor.fetchall()]
                    return classes, groups
            else:
                return [], []
    except (psycopg2.Error, ConnectionError) as e:
        print("Lỗi CSDL", f"Lỗi lấy danh sách lớp/tổ: {e}")
        return [], []

def them_danh_muc_su_kien_db(ten, loai, diem, muc_do):
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    # <<< THAY ĐỔI: Thêm cột muc_do_vi_pham >>>
                    sql = """INSERT INTO danh_muc_su_kien 
                             (ten_su_kien, loai_mac_dinh, diem_mac_dinh, muc_do_vi_pham) 
                             VALUES (%s, %s, %s, %s)"""
                    cursor.execute(sql, (ten, loai, diem, muc_do))
                conn.commit()
                return True
            else:
                print("Lỗi", "Không có kết nối CSDL.")
                return False
    except psycopg2.IntegrityError:
        print("Lỗi", f"Tên sự kiện '{ten}' đã tồn tại.")
        return False
    except (psycopg2.Error, ConnectionError) as e:
        print("Lỗi CSDL", f"{e}")
        return False

def lay_tat_ca_danh_muc_su_kien_db():
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    # <<< THAY ĐỔI: Thêm cột muc_do_vi_pham >>>
                    sql = """SELECT id, ten_su_kien, loai_mac_dinh, diem_mac_dinh, muc_do_vi_pham 
                             FROM danh_muc_su_kien 
                             ORDER BY loai_mac_dinh DESC, muc_do_vi_pham, ten_su_kien"""
                    cursor.execute(sql)
                    return cursor.fetchall()
            else:
                return []
    except (psycopg2.Error, ConnectionError) as e:
        print("Lỗi CSDL", f"{e}")
        return []

def sua_danh_muc_su_kien_db(id_dm, ten, loai, diem, muc_do):
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    # <<< THAY ĐỔI: Thêm cột muc_do_vi_pham >>>
                    sql = """UPDATE danh_muc_su_kien 
                             SET ten_su_kien = %s, loai_mac_dinh = %s, diem_mac_dinh = %s, muc_do_vi_pham = %s 
                             WHERE id = %s"""
                    cursor.execute(sql, (ten, loai, diem, muc_do, id_dm))
                conn.commit()
                return True
            else:
                print("Lỗi", "Không có kết nối CSDL.")
                return False
    except psycopg2.IntegrityError:
        print("Lỗi", f"Tên '{ten}' có thể đã tồn tại.")
        return False
    except (psycopg2.Error, ConnectionError) as e:
        print("Lỗi CSDL", f"{e}")
        return False

def xoa_danh_muc_su_kien_db(id_dm):
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    cursor.execute("DELETE FROM danh_muc_su_kien WHERE id = %s", (id_dm,))
                conn.commit()
                return True
            else:
                print("Lỗi", "Không có kết nối CSDL.")
                return False
    except (psycopg2.Error, ConnectionError) as e:
        print("Lỗi CSDL", f"{e}")
        return False

def them_hoc_sinh_db(student_data_dict):
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    cols = ', '.join(student_data_dict.keys())
                    placeholders = ', '.join(['%s'] * len(student_data_dict))
                    sql = f"INSERT INTO hoc_sinh ({cols}) VALUES ({placeholders}) RETURNING id"
                    cursor.execute(sql, list(student_data_dict.values()))
                    new_id = cursor.fetchone()[0]
                conn.commit()
                return new_id
            else:
                print("Không có kết nối CSDL, không thể thêm học sinh.")
                return None
    except (psycopg2.Error, ConnectionError) as e:
        print(f"Lỗi CSDL thêm HS: {e}")
        return None

def sua_hoc_sinh_db(student_id, student_data_dict):
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    set_clause = ', '.join([f"{key} = %s" for key in student_data_dict.keys()])
                    sql = f"UPDATE hoc_sinh SET {set_clause} WHERE id = %s"
                    vals = list(student_data_dict.values()) + [student_id]
                    cursor.execute(sql, vals)
                conn.commit()
                return True
            else:
                print("Lỗi", "Không có kết nối CSDL.")
                return False
    except (psycopg2.Error, ConnectionError) as e:
        print("Lỗi CSDL sửa HS", f"{e}")
        return False

def lay_danh_sach_hoc_sinh_db(user_role, assigned_class=None, assigned_group=None):
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    sql = "SELECT id, ten, lop, to_nhiem_vu, sdt_hoc_sinh, sdt_zalo, anh_the_path FROM hoc_sinh"
                    params = []
                    where_clauses = []
                    if user_role in ('gvcn', 'bcs'):
                        where_clauses.append("lop = %s")
                        params.append(assigned_class)
                        if assigned_group:
                            where_clauses.append("to_nhiem_vu = %s")
                            params.append(assigned_group)
                    
                    if where_clauses:
                        sql += " WHERE " + " AND ".join(where_clauses)

                    sql += " ORDER BY lop, to_nhiem_vu, ten"
                    cursor.execute(sql, params)
                    return cursor.fetchall()
            else:
                return []
    except (psycopg2.Error, ConnectionError) as e:
        print("Lỗi CSDL", f"Lỗi lấy danh sách học sinh: {e}")
        return []

def lay_thong_tin_day_du_hoc_sinh_db(student_id=None, user_role='admin', assigned_class=None, assigned_group=None):
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    db_fields_str = ', '.join(STUDENT_FIELDS_DB)
                    if student_id:
                        sql = f"SELECT id, {db_fields_str} FROM hoc_sinh WHERE id = %s"
                        cursor.execute(sql, (student_id,))
                        return cursor.fetchone()
                    else:
                        sql = f"SELECT id, {db_fields_str} FROM hoc_sinh"
                        params = []
                        where_clauses = []
                        if user_role in ('gvcn', 'bcs'):
                            where_clauses.append("lop = %s")
                            params.append(assigned_class)
                            if assigned_group:
                                where_clauses.append("to_nhiem_vu = %s")
                                params.append(assigned_group)
                        if where_clauses:
                            sql += " WHERE " + " AND ".join(where_clauses)
                        sql += " ORDER BY lop, to_nhiem_vu, ten"
                        cursor.execute(sql, params)
                        return cursor.fetchall()
            else:
                return None if student_id else []
    except (psycopg2.Error, ConnectionError) as e:
        print("Lỗi CSDL", f"Lỗi lấy thông tin học sinh: {e}")
        return None if student_id else []
    
def tim_hoc_sinh_db(ten, aclass, agroup=None, role='admin'):
    """
    Tìm ID của học sinh dựa trên tên và phạm vi quản lý (lớp, tổ).
    Trả về ID hoặc None.
    """
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    sql = "SELECT id FROM hoc_sinh WHERE ten = %s AND lop = %s"
                    params = [ten, aclass]
                    
                    if role == 'bcs' and agroup:
                        sql += " AND to_nhiem_vu = %s"
                        params.append(agroup)
                        
                    cursor.execute(sql, tuple(params))
                    result = cursor.fetchone()
                    return result[0] if result else None
            else:
                return None
    except (psycopg2.Error, ConnectionError) as e:
        print(f"Lỗi tìm học sinh: {e}")
        return None

def lay_thong_tin_su_kien_mau_db(ten_su_kien):
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    cursor.execute("SELECT loai_mac_dinh, diem_mac_dinh FROM danh_muc_su_kien WHERE ten_su_kien = %s", (ten_su_kien,))
                    result = cursor.fetchone()
                    return result if result else None
            else:
                return None
    except (psycopg2.Error, ConnectionError) as e:
        print(f"Lỗi lấy thông tin sự kiện mẫu: {e}")
        return None

def xoa_nhieu_hoc_sinh_db(student_ids):
    if not student_ids:
        return 0
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    conn.autocommit = False 
                    sql = "DELETE FROM hoc_sinh WHERE id = ANY(%s)"
                    cursor.execute(sql, (student_ids,))
                    deleted_count = cursor.rowcount
                conn.commit()
                return deleted_count
            else:
                return 0
    except (psycopg2.Error, ConnectionError) as e:
        print("Lỗi CSDL Xóa Nhiều HS", f"Đã xảy ra lỗi:\n{e}")
        return 0

def them_su_kien_ren_luyen_db(hs_id, mo_ta, loai, diem, ngay_tao_str):
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    cursor.execute("INSERT INTO su_kien_ren_luyen (hoc_sinh_id, mo_ta, loai_su_kien, diem_ap_dung, ngay_tao) VALUES (%s, %s, %s, %s, %s)",
                                   (hs_id, mo_ta, loai, diem, ngay_tao_str))
                conn.commit()
                return True
            else:
                print("Lỗi", "Không có kết nối CSDL.")
                return False
    except (psycopg2.Error, ConnectionError) as e:
        print("Lỗi CSDL", f"{e}")
        return False

def lay_su_kien_ren_luyen_cua_hoc_sinh_db(hs_id):
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    cursor.execute("SELECT id, mo_ta, loai_su_kien, diem_ap_dung, ngay_tao FROM su_kien_ren_luyen WHERE hoc_sinh_id = %s ORDER BY ngay_tao DESC, id DESC", (hs_id,))
                    return cursor.fetchall()
            else:
                return []
    except (psycopg2.Error, ConnectionError) as e:
        print("Lỗi CSDL", f"{e}")
        return []

def lay_su_kien_ren_luyen_nhieu_hoc_sinh_db(student_ids):
    if not student_ids:
        return []
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    sql = """
                        SELECT
                            hs.ten, hs.lop, hs.to_nhiem_vu,
                            skrl.ngay_tao, skrl.mo_ta, skrl.loai_su_kien, skrl.diem_ap_dung
                        FROM su_kien_ren_luyen AS skrl
                        JOIN hoc_sinh AS hs ON skrl.hoc_sinh_id = hs.id
                        WHERE skrl.hoc_sinh_id = ANY(%s)
                        ORDER BY hs.ten, skrl.ngay_tao DESC
                    """
                    cursor.execute(sql, (student_ids,))
                    return cursor.fetchall()
            else:
                return []
    except (psycopg2.Error, ConnectionError) as e:
        print("Lỗi CSDL", f"Không thể lấy lịch sử rèn luyện:\n{e}")
        return []

def xoa_su_kien_ren_luyen_db(ev_id):
    """Đổi tên từ xoa_su_kien_db cho thống nhất với bảng"""
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    cursor.execute("DELETE FROM su_kien_ren_luyen WHERE id = %s", (ev_id,))
                conn.commit()
                return True
            else:
                return False
    except (psycopg2.Error, ConnectionError) as e:
        print("Lỗi CSDL", f"{e}")
        return False

def lay_su_kien_trong_khoang_ngay_db(start_date_str, end_date_str, user_role='admin', assigned_class=None, assigned_group=None):
    end_date_inclusive_str = end_date_str + " 23:59:59"
    sql = """SELECT skrl.hoc_sinh_id, hs.ten, hs.lop, hs.to_nhiem_vu,
                      skrl.mo_ta, skrl.loai_su_kien, skrl.diem_ap_dung, skrl.ngay_tao
               FROM su_kien_ren_luyen skrl JOIN hoc_sinh hs ON skrl.hoc_sinh_id = hs.id"""
    params = []
    where_clauses = ["skrl.ngay_tao BETWEEN %s AND %s"]
    params.extend([start_date_str, end_date_inclusive_str])

    if user_role in ('gvcn', 'bcs'):
        where_clauses.append("hs.lop = %s")
        params.append(assigned_class)
        if assigned_group:
            where_clauses.append("hs.to_nhiem_vu = %s")
            params.append(assigned_group)

    sql += " WHERE " + " AND ".join(where_clauses)
    sql += " ORDER BY hs.lop, hs.to_nhiem_vu, hs.ten, skrl.ngay_tao DESC"
    
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    cursor.execute(sql, params)
                    return cursor.fetchall()
            else:
                return []
    except (psycopg2.Error, ConnectionError) as e:
        print("Lỗi CSDL", f"Lỗi lấy sự kiện theo ngày: {e}")
        return []

# TÌM VÀ THAY THẾ HÀM NÀY TRONG FILE database.py

def lay_su_kien_cho_thong_ke_df(start_date_str, end_date_str):
    end_date_inclusive_str = end_date_str + " 23:59:59"
    # <<< THAY ĐỔI: Thêm cột hs.to_nhiem_vu vào câu lệnh SELECT >>>
    sql = """
        SELECT hs.lop, hs.to_nhiem_vu, skrl.mo_ta, skrl.loai_su_kien, skrl.diem_ap_dung, skrl.ngay_tao
        FROM su_kien_ren_luyen skrl
        JOIN hoc_sinh hs ON skrl.hoc_sinh_id = hs.id
        WHERE skrl.ngay_tao BETWEEN %(start_date)s AND %(end_date)s
    """
    try:
        engine = create_engine(DB_CONNECTION_STRING)
        df = pd.read_sql_query(
            sql, 
            engine, 
            params={'start_date': start_date_str, 'end_date': end_date_inclusive_str}
        )
        return df
    except Exception as e:
        print("Lỗi CSDL", f"Không thể lấy dữ liệu thống kê: {e}")
        return pd.DataFrame()

def lay_su_kien_trong_thang_cua_hoc_sinh_db(hs_id, year, month):
    """Lấy tất cả sự kiện của một học sinh trong một tháng cụ thể."""
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    sql = """
                        SELECT id, mo_ta, loai_su_kien, diem_ap_dung, ngay_tao 
                        FROM su_kien_ren_luyen 
                        WHERE hoc_sinh_id = %s 
                            AND EXTRACT(YEAR FROM ngay_tao) = %s 
                            AND EXTRACT(MONTH FROM ngay_tao) = %s
                        ORDER BY ngay_tao ASC
                    """
                    cursor.execute(sql, (hs_id, year, month))
                    return cursor.fetchall()
            else:
                return []
    except (psycopg2.Error, ConnectionError) as e:
        print("Lỗi CSDL", f"Không thể lấy dữ liệu rèn luyện tháng:\n{e}")
        return []

def them_hinh_thuc_xu_ly_db(student_id, muc_do, hoc_ky, ghi_chu=''):
    """Thêm mới hoặc cập nhật hình thức xử lý cho học sinh theo HỌC KỲ."""
    sql = """
        INSERT INTO hinh_thuc_xu_ly (hoc_sinh_id, muc_do, hoc_ky, ghi_chu)
        VALUES (%s, %s, %s, %s)
        ON CONFLICT (hoc_sinh_id, hoc_ky)
        DO UPDATE SET muc_do = EXCLUDED.muc_do, ghi_chu = EXCLUDED.ghi_chu;
    """
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    cursor.execute(sql, (student_id, muc_do, hoc_ky, ghi_chu))
                conn.commit()
                return True
    except (psycopg2.Error, ConnectionError) as e:
        print(f"Lỗi CSDL khi thêm/cập nhật hình thức xử lý: {e}")
        return False

def lay_hinh_thuc_xu_ly_db(hoc_ky):
    """Lấy danh sách các hình thức xử lý đã được áp dụng trong một HỌC KỲ."""
    sql = "SELECT hoc_sinh_id, muc_do FROM hinh_thuc_xu_ly WHERE hoc_ky = %s"
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    cursor.execute(sql, (hoc_ky,))
                    return dict(cursor.fetchall())
        return {}
    except (psycopg2.Error, ConnectionError) as e:
        print(f"Lỗi CSDL khi lấy hình thức xử lý: {e}")
        return {}
# DÁN HÀM MỚI NÀY VÀO CUỐI FILE database.py

def them_hoac_cap_nhat_danh_muc_db(ten, loai, diem, muc_do):
    """
    Thêm một danh mục sự kiện mới. 
    Nếu tên sự kiện đã tồn tại, cập nhật lại thông tin của nó.
    """
    sql = """
        INSERT INTO danh_muc_su_kien (ten_su_kien, loai_mac_dinh, diem_mac_dinh, muc_do_vi_pham)
        VALUES (%s, %s, %s, %s)
        ON CONFLICT (ten_su_kien) 
        DO UPDATE SET 
            loai_mac_dinh = EXCLUDED.loai_mac_dinh,
            diem_mac_dinh = EXCLUDED.diem_mac_dinh,
            muc_do_vi_pham = EXCLUDED.muc_do_vi_pham;
    """
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    cursor.execute(sql, (ten, loai, diem, muc_do))
                conn.commit()
                return True
    except (psycopg2.Error, ConnectionError) as e:
        print(f"Lỗi CSDL khi thêm/cập nhật danh mục: {e}")
        return False
# DÁN HÀM MỚI NÀY VÀO CUỐI FILE database.py

def them_hoac_cap_nhat_user_db(username, full_name, password, role, assigned_class, assigned_group):
    """
    Thêm người dùng mới. Nếu username đã tồn tại, cập nhật lại thông tin và mật khẩu của họ.
    """
    password_hash = hashlib.sha256(password.encode()).hexdigest()
    sql = """
        INSERT INTO users (username, full_name, password_hash, role, assigned_class, assigned_group)
        VALUES (%s, %s, %s, %s, %s, %s)
        ON CONFLICT (username)
        DO UPDATE SET
            full_name = EXCLUDED.full_name,
            password_hash = EXCLUDED.password_hash,
            role = EXCLUDED.role,
            assigned_class = EXCLUDED.assigned_class,
            assigned_group = EXCLUDED.assigned_group;
    """
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    cursor.execute(sql, (username, full_name, password_hash, role, assigned_class, assigned_group))
                conn.commit()
                return True
    except (psycopg2.Error, ConnectionError) as e:
        print(f"Lỗi CSDL khi thêm/cập nhật người dùng: {e}")
        return False
# DÁN CÁC HÀM MỚI NÀY VÀO CUỐI FILE database.py

def lay_du_lieu_bieu_do_ca_nhan(hs_id, num_weeks=4):
    """Lấy dữ liệu điểm rèn luyện của 1 HS trong N tuần gần nhất."""
    sql = """
        SELECT 
            EXTRACT(WEEK FROM ngay_tao) as week_num, 
            SUM(diem_ap_dung) as total_points
        FROM su_kien_ren_luyen
        WHERE hoc_sinh_id = %s 
              AND ngay_tao >= NOW() - INTERVAL '%s weeks'
        GROUP BY week_num
        ORDER BY week_num;
    """
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    cursor.execute(sql, (hs_id, num_weeks))
                    return cursor.fetchall()
        return []
    except Exception as e:
        print(f"Lỗi lấy dữ liệu biểu đồ cá nhân: {e}")
        return []

def luu_muc_tieu_phan_hoi_db(hs_id, muc_tieu, phan_hoi):
    """Lưu hoặc cập nhật mục tiêu và phản hồi cho học sinh."""
    sql = """
        INSERT INTO hoc_sinh (id, muc_tieu_thang, phan_hoi_gvcn)
        VALUES (%s, %s, %s)
        ON CONFLICT (id) DO UPDATE SET
            muc_tieu_thang = EXCLUDED.muc_tieu_thang,
            phan_hoi_gvcn = EXCLUDED.phan_hoi_gvcn;
    """
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    cursor.execute(sql, (hs_id, muc_tieu, phan_hoi))
                conn.commit()
                return True
    except Exception as e:
        print(f"Lỗi lưu mục tiêu/phản hồi: {e}")
        return False

def lay_muc_tieu_phan_hoi_db(hs_id):
    """Lấy mục tiêu và phản hồi đã lưu của học sinh."""
    sql = "SELECT muc_tieu_thang, phan_hoi_gvcn FROM hoc_sinh WHERE id = %s"
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    cursor.execute(sql, (hs_id,))
                    return cursor.fetchone()
        return (None, None)
    except Exception as e:
        print(f"Lỗi lấy mục tiêu/phản hồi: {e}")
        return (None, None)
# File: database.py
# DÁN CÁC HÀM MỚI NÀY VÀO CUỐI FILE

def lay_chi_tiet_danh_muc_su_kien_db(ten_su_kien):
    """Lấy đầy đủ thông tin của một sự kiện trong danh mục, bao gồm cả mức độ vi phạm."""
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    sql = "SELECT id, ten_su_kien, loai_mac_dinh, diem_mac_dinh, muc_do_vi_pham FROM danh_muc_su_kien WHERE ten_su_kien = %s"
                    cursor.execute(sql, (ten_su_kien,))
                    return cursor.fetchone()
            return None
    except (psycopg2.Error, ConnectionError) as e:
        print(f"Lỗi CSDL khi lấy chi tiết danh mục sự kiện: {e}")
        return None

def lay_lich_su_ky_luat_cua_hoc_sinh_db(hoc_sinh_id):
    """Lấy toàn bộ lịch sử các biện pháp kỷ luật đã áp dụng cho một học sinh."""
    try:
        with get_db_connection() as conn:
            if conn:
                with conn.cursor() as cursor:
                    # Sắp xếp theo ngày gần nhất lên đầu
                    sql = "SELECT id, hinh_thuc, ngay_ap_dung, ghi_chu FROM lich_su_ky_luat WHERE hoc_sinh_id = %s ORDER BY ngay_ap_dung DESC"
                    cursor.execute(sql, (hoc_sinh_id,))
                    return cursor.fetchall()
            return []
    except (psycopg2.Error, ConnectionError) as e:
        print(f"Lỗi CSDL khi lấy lịch sử kỷ luật: {e}")
        return []

def them_su_kien_va_ky_luat_db(hs_id, mo_ta, loai_sk, diem_sk, ngay_tao_str, hinh_thuc_kl, ghi_chu_kl):
    """
    Thực hiện ghi nhận vi phạm và hình thức kỷ luật trong cùng một giao dịch (transaction).
    Đảm bảo cả hai đều thành công hoặc cả hai đều thất bại.
    """
    conn = None
    try:
        if not connection_pool:
            return False, "Không có connection pool."
        
        conn = connection_pool.getconn()
        conn.autocommit = False # Bắt đầu transaction
        
        with conn.cursor() as cursor:
            # Bước 1: Thêm sự kiện vi phạm vào bảng su_kien_ren_luyen
            sql_them_su_kien = """
                INSERT INTO su_kien_ren_luyen (hoc_sinh_id, mo_ta, loai_su_kien, diem_ap_dung, ngay_tao) 
                VALUES (%s, %s, %s, %s, %s) RETURNING id
            """
            cursor.execute(sql_them_su_kien, (hs_id, mo_ta, loai_sk, diem_sk, ngay_tao_str))
            
            # Lấy ID của sự kiện vừa được tạo
            new_event_id = cursor.fetchone()[0]
            if not new_event_id:
                raise Exception("Không thể tạo bản ghi sự kiện rèn luyện.")

            # Bước 2: Thêm hình thức kỷ luật vào bảng lich_su_ky_luat
            sql_them_ky_luat = """
                INSERT INTO lich_su_ky_luat (hoc_sinh_id, hinh_thuc, ngay_ap_dung, ghi_chu, su_kien_lien_quan_id)
                VALUES (%s, %s, %s, %s, %s)
            """
            cursor.execute(sql_them_ky_luat, (hs_id, hinh_thuc_kl, ngay_tao_str, ghi_chu_kl, new_event_id))

        # Nếu cả hai bước thành công, commit transaction
        conn.commit()
        return True, "Ghi nhận thành công."

    except (psycopg2.Error, Exception) as e:
        if conn:
            conn.rollback() # Nếu có lỗi, hủy bỏ mọi thay đổi
        return False, f"Lỗi CSDL: {e}"
    finally:
        if conn:
            connection_pool.putconn(conn)