import streamlit as st
import docx
import re

# 1. CẤU HÌNH GIAO DIỆN TRANG WEB
st.set_page_config(page_title="Phần Mềm Kiểm Tra Hợp Đồng", page_icon="🔍", layout="wide")

# 2. CÁC HÀM XỬ LÝ DỮ LIỆU
def normalize_text(text):
    """Hàm làm sạch chữ: Xóa mọi khoảng trắng thừa, dấu tab, xuống dòng để so sánh chính xác tuyệt đối"""
    if not text: return ""
    return re.sub(r'\s+', '', str(text)).lower()

def extract_input_data(input_text):
    """Hàm bóc tách dữ liệu từ Input gốc của Zalo/Email"""
    data = {
        "cong_ty": "", "mst": "", "dia_chi": "", "dai_dien": "", "chuc_vu": "",
        "tong_tien": "", "bang_chu": "", "dia_chi_giao": "",
        "san_pham": []
    }
    
    # Tách từng dòng và xóa khoảng trắng 2 đầu
    lines = [line.strip() for line in input_text.split('\n') if line.strip()]
    
    i = 0
    while i < len(lines):
        line = lines[i]
        line_lower = line.lower()
        
        if line_lower.startswith("tên công ty"):
            data["cong_ty"] = line.split(":", 1)[1].strip()
        elif line_lower.startswith("mã số thuế"):
            data["mst"] = line.split(":", 1)[1].strip()
        elif line_lower.startswith("địa chỉ :") or line_lower.startswith("địa chỉ:"):
            data["dia_chi"] = line.split(":", 1)[1].strip()
        elif line_lower.startswith("người đại diện"):
            data["dai_dien"] = line.split(":", 1)[1].strip()
        elif line_lower.startswith("chức vụ"):
            data["chuc_vu"] = line.split(":", 1)[1].strip()
        elif line_lower.startswith("tổng :") or line_lower.startswith("tổng:"):
            data["tong_tien"] = line.split(":", 1)[1].strip()
        elif line_lower.startswith("địa chỉ nhận hàng"):
            data["dia_chi_giao"] = line.split(":", 1)[1].strip()
            
        # Thuật toán bắt sản phẩm (Bắt dòng có dạng: Giá x Số lượng = Thành tiền)
        elif i > 0 and "x" in line and "=" in line and any(c.isdigit() for c in line):
            ten_sp = lines[i-1] # Tên sản phẩm luôn nằm ở dòng ngay trước đó
            match = re.search(r'([\d\.]+)\s*x\s*(\d+)\s*=\s*([\d\.]+)', line)
            if match:
                data["san_pham"].append({
                    "ten": ten_sp,
                    "don_gia": match.group(1),
                    "so_luong": match.group(2),
                    "thanh_tien": match.group(3),
                    "chuoi_gia_goc": line
                })
        # Bắt dòng chữ số tiền (thường có chữ đồng hoặc ./. ở cuối)
        elif "đồng./." in line_lower or ("đồng" in line_lower and len(line) > 15):
             data["bang_chu"] = line.strip()
        i += 1
        
    return data

# 3. GIAO DIỆN CHÍNH VÀ XỬ LÝ ĐỐI CHIẾU
def main():
    st.title("Phần Mềm Kiểm Tra Hợp Đồng (Phiên Bản Tự Động Đối Chiếu Bên A) 🔍")
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1. Nhập thông tin gốc")
        input_text = st.text_area("Paste dữ liệu Input của bạn vào đây:", height=400)
        
    with col2:
        st.subheader("2. Tải lên file Word Hợp đồng")
        uploaded_file = st.file_uploader("Chọn file .docx", type=["docx"])
        
    if st.button("🚀 BẮT ĐẦU ĐỐI CHIẾU", use_container_width=True):
        if not input_text or not uploaded_file:
            st.warning("Vui lòng nhập đủ Thông tin gốc và Tải lên file Word!")
            return
            
        with st.spinner("Đang phân tích và đối chiếu..."):
            # 1. Trích xuất Input
            input_data = extract_input_data(input_text)
            
            # 2. Đọc file Word (Đoạn này đã fix lỗi khoảng trắng ẩn)
            doc = docx.Document(uploaded_file)
            
            # Đọc đoạn văn thường
            full_word_text = "\n".join([p.text for p in doc.paragraphs])
            
            # Đọc bảng (Lọc sạch dấu xuống dòng và khoảng trắng)
            table_rows_clean = []
            for table in doc.tables:
                for row in table.rows:
                    row_data = []
                    for cell in row.cells:
                        # FIX LỖI: Thay thế khoảng trắng ẩn, dấu xuống dòng \n bằng 1 dấu cách duy nhất
                        clean_cell = re.sub(r'\s+', ' ', cell.text).strip()
                        row_data.append(clean_cell)
                    table_rows_clean.append(" | ".join(row_data))
                    
            full_table_text = "\n".join(table_rows_clean)
            
            # Gom chung để đối chiếu các trường text chung
            full_doc_combined = full_word_text + "\n" + full_table_text
            full_doc_normalized = normalize_text(full_doc_combined)

            # 3. Hiển thị Kết quả
            st.markdown("---")
            st.subheader("📊 KẾT QUẢ ĐỐI CHIẾU")
            
            def show_result(label, is_match, original_text):
                if is_match:
                    st.success(f"✅ **{label}:** Khớp ({original_text})")
                else:
                    st.error(f"❌ **{label}:** LỖI SAI LỆCH HOẶC KHÔNG TÌM THẤY! (Gốc: {original_text})")

            # A. ĐỐI CHIẾU THÔNG TIN CƠ BẢN
            st.markdown("### I. Thông tin cơ bản")
            show_result("Tên công ty", normalize_text(input_data['cong_ty']) in full_doc_normalized, input_data['cong_ty'])
            show_result("Mã số thuế", normalize_text(input_data['mst']) in full_doc_normalized, input_data['mst'])
            show_result("Địa chỉ", normalize_text(input_data['dia_chi']) in full_doc_normalized, input_data['dia_chi'])
            show_result("Người đại diện", normalize_text(input_data['dai_dien']) in full_doc_normalized, input_data['dai_dien'])
            show_result("Chức vụ", normalize_text(input_data['chuc_vu']) in full_doc_normalized, input_data['chuc_vu'])
            show_result("Tổng tiền", normalize_text(input_data['tong_tien']) in full_doc_normalized, input_data['tong_tien'])
            show_result("Số tiền bằng chữ", normalize_text(input_data['bang_chu']) in full_doc_normalized, input_data['bang_chu'])
            show_result("Địa chỉ nhận hàng", normalize_text(input_data['dia_chi_giao']) in full_doc_normalized, input_data['dia_chi_giao'])
            
            # B. ĐỐI CHIẾU SẢN PHẨM
            st.markdown("### II. Chi tiết Hàng hóa / Sản phẩm")
            if not input_data['san_pham']:
                st.info("Không tìm thấy danh sách sản phẩm trong Input gốc.")
                
            for idx, sp in enumerate(input_data['san_pham']):
                # Tên sản phẩm
                is_name_match = normalize_text(sp['ten']) in full_doc_normalized
                show_result(f"Tên Sản Phẩm {idx+1}", is_name_match, sp['ten'])
                
                # Giá và Số lượng (Tìm xem có dòng nào trong bảng khớp giá và số lượng không)
                is_qty_price_match = False
                for row_text in table_rows_clean:
                    row_norm = normalize_text(row_text)
                    # Bắt buộc trên cùng 1 dòng của bảng phải có cả Số Lượng và Đơn Giá
                    if normalize_text(sp['so_luong']) in row_norm and normalize_text(sp['don_gia']) in row_norm:
                        # Kết hợp check cả thành tiền để chắc chắn 100% không bắt nhầm sang dòng sản phẩm khác
                        if normalize_text(sp['thanh_tien']) in row_norm or normalize_text(sp['ten'][:15]) in row_norm:
                            is_qty_price_match = True
                            break
                
                if is_qty_price_match:
                    st.success(f"✅ **Giá & Số lượng SP {idx+1}:** Khớp toàn bộ thông số ({sp['chuoi_gia_goc']})")
                else:
                    st.error(f"❌ **Giá & Số lượng SP {idx+1}:** SAI SỐ LƯỢNG HOẶC ĐƠN GIÁ! Không tìm thấy dữ liệu khớp trong bảng. (Gốc: {sp['chuoi_gia_goc']})")

if __name__ == "__main__":
    main()
