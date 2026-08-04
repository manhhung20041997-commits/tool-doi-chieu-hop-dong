import streamlit as st
import docx
import re
import unicodedata
import io # Thêm thư viện io để xử lý tải file

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

# 3. HÀM TỰ ĐỘNG SỬA FILE WORD (AUTO-FIX)
def auto_fix_document(doc, input_data):
    """Hàm tự động tìm và thay thế nội dung sai trong Word bằng nội dung gốc"""
    
    # A. Sửa lỗi trong BẢNG (Sản phẩm, Tổng tiền, Bằng chữ)
    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            
            # Sửa thông tin Sản Phẩm
            if len(cells) >= 5: 
                ten_sp_word = normalize_text(cells[1].text)
                if ten_sp_word and len(ten_sp_word) > 10: 
                    for sp in input_data['san_pham']:
                        # Nếu đúng dòng sản phẩm đó -> Ghi đè SL và Đơn giá
                        if normalize_text(sp['ten'])[:30] in ten_sp_word:
                            cells[3].text = str(sp['so_luong'])
                            cells[4].text = str(sp['don_gia'])
                            if len(cells) >= 6:
                                cells[5].text = str(sp['thanh_tien'])
            
            # Sửa Tổng tiền và Bằng chữ (nếu nằm trong bảng)
            for cell in cells:
                text_lower = normalize_text(cell.text)
                if "tổng(đãbaogồmvat)" in text_lower:
                    cell.text = f"Tổng (đã bao gồm VAT): {input_data['tong_tien']}"
                elif "bằngchữ:" in text_lower:
                    cell.text = f"Bằng chữ: {input_data['bang_chu']}"

    # B. Sửa lỗi Thông tin cơ bản (Đoạn văn)
    # Dùng biến đếm để chỉ sửa thông tin của Bên A (xuất hiện lần đầu), tránh sửa nhầm của Bên B
    count_mst = 0
    count_dc = 0
    count_dd = 0
    
    for p in doc.paragraphs:
        text_lower = p.text.lower()
        
        if "bên mua (bên a):" in text_lower:
            p.text = f"Bên Mua (Bên A): {input_data['cong_ty'].upper()}"
        elif "mã số thuế" in text_lower:
            count_mst += 1
            if count_mst == 1: p.text = f"Mã số thuế\t: {input_data['mst']}"
        elif "địa chỉ" in text_lower and ("phường" in text_lower or "khu phố" in text_lower or count_dc == 0):
            count_dc += 1
            if count_dc == 1: p.text = f"Địa chỉ\t: {input_data['dia_chi']}"
        elif "đại diện" in text_lower and "chức vụ" in text_lower:
            count_dd += 1
            if count_dd == 1: p.text = f"Đại diện\t: {input_data['dai_dien']}\tChức vụ: {input_data['chuc_vu']}"
        elif "địa điểm giao hàng" in text_lower:
            p.text = f"4.3.\tĐịa điểm giao hàng: {input_data['dia_chi_giao']}"
            
        # Dự phòng trường hợp Tổng tiền/Bằng chữ nằm ngoài bảng
        elif "tổng (đã bao gồm vat):" in text_lower:
            p.text = f"Tổng (đã bao gồm VAT): {input_data['tong_tien']}"
        elif "bằng chữ:" in text_lower and "đồng" in text_lower:
            p.text = f"Bằng chữ: {input_data['bang_chu']}"

    return doc

# 4. GIAO DIỆN CHÍNH
def main():
    st.title("Phần Mềm Kiểm Tra & Tự Động Sửa Hợp Đồng 🤖")
    st.markdown("*Tool sẽ quét lỗi. Nếu phát hiện sai sót, tool sẽ tự động tạo ra một file Word hoàn chỉnh mới cho bạn tải về!*")
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1. Nhập thông tin gốc")
        input_text = st.text_area("Paste dữ liệu Input của bạn vào đây:", height=400)
        
    with col2:
        st.subheader("2. Tải lên file Word Hợp đồng")
        uploaded_file = st.file_uploader("Chọn file .docx", type=["docx"])
        
    if st.button("🚀 BẮT ĐẦU ĐỐI CHIẾU & FIX LỖI", use_container_width=True):
        if not input_text or not uploaded_file:
            st.warning("Vui lòng nhập đủ Thông tin gốc và Tải lên file Word!")
            return
            
        with st.spinner("Đang phân tích, đối chiếu và sửa lỗi..."):
            input_data = extract_input_data(input_text)
            
            # --- ĐỌC VÀ KIỂM TRA LỖI (GIỮ NGUYÊN LOGIC CŨ) ---
            doc = docx.Document(uploaded_file)
            full_word_text = "\n".join([p.text for p in doc.paragraphs])
            
            danh_sach_sp_word = []
            all_table_text = [] 
            
            for table in doc.tables:
                for row in table.rows:
                    row_text = [cell.text for cell in row.cells]
                    all_table_text.append(" ".join(row_text))
                    
                    cells = row.cells
                    if len(cells) >= 5: 
                        ten_sp_word = normalize_text(cells[1].text)
                        if ten_sp_word and len(ten_sp_word) > 10: 
                            danh_sach_sp_word.append({
                                "ten": ten_sp_word,
                                "so_luong": normalize_text(cells[3].text),
                                "don_gia": normalize_text(cells[4].text)
                            })
            
            full_table_text = "\n".join(all_table_text)
            full_doc_combined = full_word_text + "\n" + full_table_text
            full_doc_normalized = normalize_text(full_doc_combined)

            st.markdown("---")
            st.subheader("📊 KẾT QUẢ ĐỐI CHIẾU")
            
            has_error = False # Biến theo dõi xem có lỗi hay không
            
            def show_result(label, is_match, original_text):
                nonlocal has_error
                if is_match:
                    st.success(f"✅ **{label}:** Khớp ({original_text})")
                else:
                    has_error = True
                    st.error(f"❌ **{label}:** LỖI! (Gốc: {original_text}) -> Sẽ được Auto-Fix")

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
                ten_goc_norm = normalize_text(sp['ten'])
                is_name_match = False
                is_qty_price_match = False
                qty_goc_norm = normalize_text(sp['so_luong'])
                price_goc_norm = normalize_text(sp['don_gia'])
                
                for sp_word in danh_sach_sp_word:
                    if ten_goc_norm[:30] in sp_word['ten']:
                        is_name_match = True
                        if qty_goc_norm == sp_word['so_luong'] and price_goc_norm in sp_word['don_gia']:
                            is_qty_price_match = True
                        break 
                
                show_result(f"Tên Sản Phẩm {idx+1}", is_name_match, sp['ten'])
                
                if is_qty_price_match:
                    st.success(f"✅ **Giá & Số lượng SP {idx+1}:** Khớp toàn bộ thông số ({sp['chuoi_gia_goc']})")
                else:
                    has_error = True
                    st.error(f"❌ **Giá & Số lượng SP {idx+1}:** SAI! -> Sẽ được Auto-Fix. (Gốc: {sp['chuoi_gia_goc']})")

            # --- NẾU CÓ LỖI, GỌI HÀM AUTO FIX VÀ XUẤT FILE MỚI ---
            st.markdown("---")
            if has_error:
                st.info("🤖 Đã phát hiện lỗi sai! Tool đang tiến hành tự động thay thế bằng dữ liệu gốc...")
                
                # Gọi hàm fix file word
                doc_fixed = auto_fix_document(doc, input_data)
                
                # Lưu file vào bộ nhớ tạm (io.BytesIO) thay vì lưu vào ổ cứng
                bio = io.BytesIO()
                doc_fixed.save(bio)
                
                # Hiển thị nút tải xuống
                st.success("🎉 Đã sửa xong! Bạn hãy tải file Word hoàn chỉnh ở nút bên dưới:")
                st.download_button(
                    label="📥 TẢI XUỐNG FILE HỢP ĐỒNG ĐÃ SỬA",
                    data=bio.getvalue(),
                    file_name="Hop_Dong_Da_Fix.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.success("🎉 Chúc mừng! Hợp đồng này chính xác 100%, không cần phải sửa gì thêm!")

if __name__ == "__main__":
    main()
