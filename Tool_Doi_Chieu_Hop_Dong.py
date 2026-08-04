import streamlit as st
import docx
import re
import unicodedata

# 1. CẤU HÌNH GIAO DIỆN TRANG WEB
st.set_page_config(page_title="Phần Mềm Kiểm Tra Hợp Đồng", page_icon="🔍", layout="wide")

# 2. CÁC HÀM XỬ LÝ DỮ LIỆU
def normalize_text(text):
    """Hàm làm sạch chữ: Xóa khoảng trắng và đồng bộ mã tiếng Việt"""
    if not text: return ""
    text_nfc = unicodedata.normalize('NFC', str(text))
    return re.sub(r'\s+', '', text_nfc).lower()

def extract_input_data(input_text):
    """Hàm bóc tách dữ liệu từ Input gốc của Zalo/Email"""
    data = {
        "cong_ty": "", "mst": "", "dia_chi": "", "dai_dien": "", "chuc_vu": "",
        "tong_tien": "", "bang_chu": "", "dia_chi_giao": "",
        "san_pham": []
    }
    
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
            
        elif i > 0 and "x" in line and "=" in line and any(c.isdigit() for c in line):
            ten_sp = lines[i-1]
            match = re.search(r'([\d\.]+)\s*x\s*(\d+)\s*=\s*([\d\.]+)', line)
            if match:
                data["san_pham"].append({
                    "ten": ten_sp,
                    "don_gia": match.group(1),
                    "so_luong": match.group(2),
                    "thanh_tien": match.group(3),
                    "chuoi_gia_goc": line
                })
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
            input_data = extract_input_data(input_text)
            doc = docx.Document(uploaded_file)
            
            # 1. Lấy toàn bộ text ở các đoạn văn thường
            full_word_text = "\n".join([p.text for p in doc.paragraphs])
            
            danh_sach_sp_word = []
            all_table_text = [] # KHO LƯU TOÀN BỘ CHỮ TRONG BẢNG
            
            for table in doc.tables:
                for row in table.rows:
                    # Gom chữ trong bảng lại
                    row_text = [cell.text for cell in row.cells]
                    all_table_text.append(" ".join(row_text))
                    
                    cells = row.cells
                    # Bảng hợp đồng chuẩn có 6 cột: STT(0) | Tên(1) | Mô tả(2) | Số lượng(3) | Đơn giá(4) | Thành tiền(5)
                    if len(cells) >= 5: 
                        ten_sp_word = normalize_text(cells[1].text)
                        
                        # Chỉ lấy những dòng có chữ dài (Tên SP) để loại trừ dòng tiêu đề hoặc dòng tính tổng
                        if ten_sp_word and len(ten_sp_word) > 10: 
                            danh_sach_sp_word.append({
                                "ten": ten_sp_word,
                                "so_luong": normalize_text(cells[3].text), # Ép cứng lấy cột Số lượng
                                "don_gia": normalize_text(cells[4].text)   # Ép cứng lấy cột Đơn giá
                            })
            
            # SỬA LỖI Ở ĐÂY: Gộp text của cả đoạn văn VÀ text trong bảng lại với nhau để không sót Tổng tiền
            full_table_text = "\n".join(all_table_text)
            full_doc_combined = full_word_text + "\n" + full_table_text
            full_doc_normalized = normalize_text(full_doc_combined)

            st.markdown("---")
            st.subheader("📊 KẾT QUẢ ĐỐI CHIẾU")
            
            def show_result(label, is_match, original_text):
                if is_match:
                    st.success(f"✅ **{label}:** Khớp ({original_text})")
                else:
                    st.error(f"❌ **{label}:** LỖI SAI LỆCH HOẶC KHÔNG TÌM THẤY! (Gốc: {original_text})")

            st.markdown("### I. Thông tin cơ bản")
            show_result("Tên công ty", normalize_text(input_data['cong_ty']) in full_doc_normalized, input_data['cong_ty'])
            show_result("Mã số thuế", normalize_text(input_data['mst']) in full_doc_normalized, input_data['mst'])
            show_result("Địa chỉ", normalize_text(input_data['dia_chi']) in full_doc_normalized, input_data['dia_chi'])
            show_result("Người đại diện", normalize_text(input_data['dai_dien']) in full_doc_normalized, input_data['dai_dien'])
            show_result("Chức vụ", normalize_text(input_data['chuc_vu']) in full_doc_normalized, input_data['chuc_vu'])
            show_result("Tổng tiền", normalize_text(input_data['tong_tien']) in full_doc_normalized, input_data['tong_tien'])
            show_result("Số tiền bằng chữ", normalize_text(input_data['bang_chu']) in full_doc_normalized, input_data['bang_chu'])
            show_result("Địa chỉ nhận hàng", normalize_text(input_data['dia_chi_giao']) in full_doc_normalized, input_data['dia_chi_giao'])
            
            st.markdown("### II. Chi tiết Hàng hóa / Sản phẩm")
            if not input_data['san_pham']:
                st.info("Không tìm thấy danh sách sản phẩm trong Input gốc.")
                
            for idx, sp in enumerate(input_data['san_pham']):
                # 1. So sánh Tên sản phẩm
                ten_goc_norm = normalize_text(sp['ten'])
                is_name_match = False
                is_qty_price_match = False
                qty_goc_norm = normalize_text(sp['so_luong'])
                price_goc_norm = normalize_text(sp['don_gia'])
                
                # 2. Tìm sản phẩm tương ứng trong danh sách các dòng của file Word
                for sp_word in danh_sach_sp_word:
                    # Dùng 30 ký tự đầu để định vị xem đang check đúng dòng sản phẩm đó chưa
                    if ten_goc_norm[:30] in sp_word['ten']:
                        is_name_match = True
                        if qty_goc_norm == sp_word['so_luong'] and price_goc_norm in sp_word['don_gia']:
                            is_qty_price_match = True
                        break 
                
                show_result(f"Tên Sản Phẩm {idx+1}", is_name_match, sp['ten'])
                
                if is_qty_price_match:
                    st.success(f"✅ **Giá & Số lượng SP {idx+1}:** Khớp toàn bộ thông số ({sp['chuoi_gia_goc']})")
                else:
                    st.error(f"❌ **Giá & Số lượng SP {idx+1}:** SAI SỐ LƯỢNG HOẶC ĐƠN GIÁ! Vui lòng kiểm tra lại cột Số Lượng và Đơn Giá. (Gốc: {sp['chuoi_gia_goc']})")

if __name__ == "__main__":
    main()
