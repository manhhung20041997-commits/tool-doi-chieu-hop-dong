import streamlit as st
import docx
import re
import unicodedata
import io 

# 1. CẤU HÌNH GIAO DIỆN TRANG WEB
st.set_page_config(page_title="Phần Mềm Kiểm Tra Hợp Đồng", page_icon="🔍", layout="wide")

# 2. CÁC HÀM XỬ LÝ DỮ LIỆU
def normalize_text(text):
    if not text: return ""
    text_nfc = unicodedata.normalize('NFC', str(text))
    return re.sub(r'\s+', '', text_nfc).lower()

def extract_input_data(input_text):
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

# 3. HÀM TỰ ĐỘNG SỬA FILE WORD
def auto_fix_document(doc, input_data):
    
    def replace_text_keep_format(paragraph, new_text):
        """Hàm thay chữ nhưng giữ nguyên format của cục chữ đầu tiên"""
        if not paragraph.runs:
            paragraph.text = str(new_text)
            return
        paragraph.runs[0].text = str(new_text)
        for i in range(len(paragraph.runs) - 1, 0, -1):
            paragraph._element.remove(paragraph.runs[i]._element)

    def replace_mixed_format(paragraph, segments):
        """Hàm thay chữ linh hoạt: cho phép chỉ định đoạn nào in đậm, đoạn nào in thường"""
        if not paragraph.runs:
            for text, is_bold in segments:
                run = paragraph.add_run(text)
                run.bold = is_bold
            return
            
        # Lưu lại font gốc
        base_font_name = paragraph.runs[0].font.name
        base_font_size = paragraph.runs[0].font.size
        base_italic = paragraph.runs[0].font.italic
        
        # Xóa các run cũ
        paragraph.clear()
        
        # Tạo run mới với định dạng chuẩn
        for text, is_bold in segments:
            run = paragraph.add_run(text)
            run.bold = is_bold # <-- Phép màu nằm ở đây (Tự động in đậm theo yêu cầu)
            if base_italic is not None:
                run.font.italic = base_italic
            if base_font_name:
                run.font.name = base_font_name
            if base_font_size:
                run.font.size = base_font_size

    # A. Sửa lỗi trong BẢNG
    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            if len(cells) >= 5: 
                ten_sp_word = normalize_text(cells[1].text)
                if ten_sp_word and len(ten_sp_word) > 10: 
                    for sp in input_data['san_pham']:
                        if normalize_text(sp['ten'])[:30] in ten_sp_word:
                            if cells[3].paragraphs: replace_text_keep_format(cells[3].paragraphs[0], sp['so_luong'])
                            if cells[4].paragraphs: replace_text_keep_format(cells[4].paragraphs[0], sp['don_gia'])
                            if len(cells) >= 6 and cells[5].paragraphs:
                                replace_text_keep_format(cells[5].paragraphs[0], sp['thanh_tien'])
            
            for cell in cells:
                for p in cell.paragraphs:
                    text_lower = normalize_text(p.text)
                    if "tổng(đãbaogồmvat)" in text_lower:
                        replace_mixed_format(p, [("Tổng (đã bao gồm VAT): ", False), (input_data['tong_tien'], True)])
                    elif "bằngchữ:" in text_lower:
                        replace_mixed_format(p, [("Bằng chữ: ", False), (input_data['bang_chu'], True)])

    # B. Sửa lỗi Thông tin cơ bản (Đoạn văn)
    count_mst = 0
    count_dc = 0
    count_dd = 0
    passed_4_2 = False 
    
    for p in doc.paragraphs:
        text_lower = p.text.lower().strip()
        
        if "bên mua (bên a):" in text_lower:
            # Sửa Tên Công Ty (Tự động in đậm Tên công ty)
            replace_mixed_format(p, [
                ("Bên Mua (Bên A): ", False),
                (input_data['cong_ty'].upper(), True)
            ])
        elif "mã số thuế" in text_lower:
            count_mst += 1
            if count_mst == 1: replace_text_keep_format(p, f"Mã số thuế\t: {input_data['mst']}")
        elif "địa chỉ" in text_lower and ("phường" in text_lower or "khu phố" in text_lower or count_dc == 0) and not text_lower.startswith("4.3") and "nhận hàng" not in text_lower:
            count_dc += 1
            if count_dc == 1: replace_text_keep_format(p, f"Địa chỉ\t: {input_data['dia_chi']}")
        elif "đại diện" in text_lower and "chức vụ" in text_lower:
            count_dd += 1
            if count_dd == 1: 
                # Sửa Đại diện và Chức vụ (Tự động in đậm tên người và chức vụ)
                replace_mixed_format(p, [
                    ("Đại diện\t: ", False),
                    (input_data['dai_dien'], True),
                    ("\tChức vụ: ", False),
                    (input_data['chuc_vu'], True)
                ])
                
        elif text_lower.startswith("4.2"):
            passed_4_2 = True 
            
        elif text_lower.startswith("4.3") and "địa điểm giao hàng" in text_lower:
            if passed_4_2: 
                replace_text_keep_format(p, f"4.3.\tĐịa điểm giao hàng: {input_data['dia_chi_giao']}")
                
        elif "tổng (đã bao gồm vat):" in text_lower:
            replace_mixed_format(p, [("Tổng (đã bao gồm VAT): ", False), (input_data['tong_tien'], True)])
        elif "bằng chữ:" in text_lower and "đồng" in text_lower:
            replace_mixed_format(p, [("Bằng chữ: ", False), (input_data['bang_chu'], True)])

    return doc

# 4. GIAO DIỆN CHÍNH
def main():
    st.title("Phần Mềm Kiểm Tra & Tự Động Sửa Hợp Đồng 🤖")
    st.markdown("*Tool sẽ quét lỗi. Nếu phát hiện sai sót, tool sẽ tự động thay thế bằng dữ liệu gốc (Đảm bảo 100% không lệch Font & In đậm)!*")
    
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
            has_error = False 
            
            def show_result(label, is_match, original_text):
                nonlocal has_error
                if is_match:
                    st.success(f"✅ **{label}:** Khớp ({original_text})")
                else:
                    has_error = True
                    st.error(f"❌ **{label}:** LỖI! (Gốc: {original_text}) -> Đã Auto-Fix")

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
                    st.error(f"❌ **Giá & Số lượng SP {idx+1}:** SAI! -> Đã Auto-Fix. (Gốc: {sp['chuoi_gia_goc']})")

            st.markdown("---")
            if has_error:
                st.info("🤖 Đã phát hiện lỗi sai! Tool đang tiến hành tự động thay thế bằng dữ liệu gốc (Giữ nguyên định dạng Word)...")
                
                doc_fixed = auto_fix_document(doc, input_data)
                bio = io.BytesIO()
                doc_fixed.save(bio)
                
                st.success("🎉 Đã sửa xong! Bạn hãy tải file Word hoàn chỉnh ở nút bên dưới:")
                st.download_button(
                    label="📥 TẢI XUỐNG FILE HỢP ĐỒNG ĐÃ SỬA CHUẨN FONT & IN ĐẬM",
                    data=bio.getvalue(),
                    file_name="Hop_Dong_Da_Fix.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.success("🎉 Chúc mừng! Hợp đồng này chính xác 100%, không cần phải sửa gì thêm!")

if __name__ == "__main__":
    main()
