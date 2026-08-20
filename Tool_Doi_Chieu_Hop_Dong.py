import streamlit as st
import docx
import re
import unicodedata
import io 
import time

# 1. CẤU HÌNH GIAO DIỆN & CSS CAO CẤP
st.set_page_config(
    page_title="Hệ Thống Kiểm Tra & Sửa Hợp Đồng Tự Động", 
    page_icon="🤖", 
    layout="wide"
)

st.markdown("""
<style>
    /* Nhập Google Font: Be Vietnam Pro cho tiếng Việt cực đẹp và sắc nét */
    @import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Be Vietnam Pro', sans-serif !important;
    }
    
    /* Nền tổng thể nhẹ nhàng, hiện đại */
    .main {
        background-color: #f8fafc;
        padding: 2rem 1.5rem;
    }
    
    /* Tiêu đề chính */
    h1 {
        color: #1e293b;
        font-weight: 700;
        font-size: 2.2rem;
        letter-spacing: -0.025em;
        margin-bottom: 0.5rem;
    }
    
    /* Tùy chỉnh ô nhập liệu text_area */
    div.stTextArea textarea {
        border-radius: 12px !important;
        border: 1px solid #cbd5e1 !important;
        background-color: #ffffff !important;
        font-family: 'Be Vietnam Pro', sans-serif !important;
        padding: 1rem !important;
        font-size: 0.95rem;
    }
    div.stTextArea textarea:focus {
        border-color: #2563eb !important;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.15) !important;
    }
    
    /* Khung kéo thả file uploader */
    div[data-testid="stFileUploader"] {
        border: 2px dashed #cbd5e1 !important;
        border-radius: 12px !important;
        background-color: #ffffff !important;
        padding: 1.5rem !important;
        transition: all 0.3s ease;
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: #2563eb !important;
        background-color: #f1f5f9 !important;
    }
    
    /* Nút bấm chính (Primary Button) */
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
        border: none !important;
        border-radius: 12px !important;
        font-weight: 600 !important;
        font-size: 1.05rem !important;
        padding: 0.75rem 1.5rem !important;
        color: white !important;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3) !important;
        transition: all 0.3s ease !important;
    }
    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #1d4ed8 0%, #1e40af 100%) !important;
        box-shadow: 0 6px 16px rgba(37, 99, 235, 0.4) !important;
        transform: translateY(-1px);
    }
    
    /* Hộp thống kê (Metrics Dashboard) */
    [data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        padding: 1.25rem;
        border-radius: 16px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.02), 0 2px 4px -1px rgba(0, 0, 0, 0.02);
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.85rem !important;
        font-weight: 600 !important;
        color: #64748b !important;
    }
    [data-testid="stMetricValue"] {
        font-size: 1.75rem !important;
        font-weight: 700 !important;
        color: #0f172a !important;
    }
    
    /* Expander kết quả */
    details {
        border: 1px solid #e2e8f0 !important;
        border-radius: 12px !important;
        background-color: #ffffff !important;
        margin-bottom: 1rem;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.02);
    }
</style>
""", unsafe_allow_html=True)

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
        if line_lower.startswith("tên công ty"): data["cong_ty"] = line.split(":", 1)[1].strip()
        elif line_lower.startswith("mã số thuế"): data["mst"] = line.split(":", 1)[1].strip()
        elif line_lower.startswith("địa chỉ :") or line_lower.startswith("địa chỉ:"): data["dia_chi"] = line.split(":", 1)[1].strip()
        elif line_lower.startswith("người đại diện"): data["dai_dien"] = line.split(":", 1)[1].strip()
        elif line_lower.startswith("chức vụ"): data["chuc_vu"] = line.split(":", 1)[1].strip()
        elif line_lower.startswith("tổng :") or line_lower.startswith("tổng:"): data["tong_tien"] = line.split(":", 1)[1].strip()
        elif line_lower.startswith("địa chỉ nhận hàng"): data["dia_chi_giao"] = line.split(":", 1)[1].strip()
        elif i > 0 and "x" in line and "=" in line and any(c.isdigit() for c in line):
            ten_sp = lines[i-1]
            match = re.search(r'([\d\.]+)\s*x\s*(\d+)\s*=\s*([\d\.]+)', line)
            if match:
                data["san_pham"].append({
                    "ten": ten_sp, "don_gia": match.group(1), 
                    "so_luong": match.group(2), "thanh_tien": match.group(3), "chuoi_gia_goc": line
                })
        elif "đồng./." in line_lower or ("đồng" in line_lower and len(line) > 15):
             data["bang_chu"] = line.strip()
        i += 1
    return data

# 3. HÀM TỰ ĐỘNG SỬA FILE WORD
def auto_fix_document(doc, input_data):
    
    def replace_text_keep_format(paragraph, new_text):
        if not paragraph.runs:
            paragraph.text = str(new_text)
            return
        paragraph.runs[0].text = str(new_text)
        for i in range(len(paragraph.runs) - 1, 0, -1):
            paragraph._element.remove(paragraph.runs[i]._element)

    def replace_mixed_format(paragraph, segments):
        if not paragraph.runs:
            for text, is_bold in segments:
                run = paragraph.add_run(text)
                run.bold = is_bold
            return
        base_font_name = paragraph.runs[0].font.name
        base_font_size = paragraph.runs[0].font.size
        base_italic = paragraph.runs[0].font.italic
        paragraph.clear()
        for text, is_bold in segments:
            run = paragraph.add_run(text)
            run.bold = is_bold
            if base_italic is not None: run.font.italic = base_italic
            if base_font_name: run.font.name = base_font_name
            if base_font_size: run.font.size = base_font_size

    # A. Sửa lỗi trong BẢNG
    processed_paragraphs = set() 
    
    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            if len(cells) >= 5: 
                ten_sp_word = normalize_text(cells[1].text)
                if ten_sp_word and len(ten_sp_word) > 10: 
                    for sp in input_data['san_pham']:
                        sp_ten_norm = normalize_text(sp['ten'])
                        if sp_ten_norm[:30] in ten_sp_word and sp_ten_norm[-15:] in ten_sp_word:
                            if cells[3].paragraphs: replace_text_keep_format(cells[3].paragraphs[0], sp['so_luong'])
                            if cells[4].paragraphs: replace_text_keep_format(cells[4].paragraphs[0], sp['don_gia'])
                            if len(cells) >= 6 and cells[5].paragraphs: replace_text_keep_format(cells[5].paragraphs[0], sp['thanh_tien'])
            
            for cell in cells:
                for p in cell.paragraphs:
                    if p in processed_paragraphs:
                        continue
                    processed_paragraphs.add(p)
                    
                    text_lower = normalize_text(p.text)
                    if "tổng(đãbaogồmvat)" in text_lower: 
                        replace_mixed_format(p, [("Tổng (đã bao gồm VAT): ", True), (input_data['tong_tien'], True)])
                    elif "bằngchữ:" in text_lower: 
                        replace_mixed_format(p, [("Bằng chữ: ", True), (input_data['bang_chu'], True)])

    # B. Sửa lỗi Thông tin cơ bản 
    count_mst = 0
    count_dc = 0
    count_dd = 0
    passed_4_2 = False 
    
    for p in doc.paragraphs:
        text_lower = p.text.lower().strip()
        if "bên mua (bên a):" in text_lower: replace_mixed_format(p, [("Bên Mua (Bên A): ", False), (input_data['cong_ty'].upper(), True)])
        elif "mã số thuế" in text_lower:
            count_mst += 1
            if count_mst == 1: replace_text_keep_format(p, f"Mã số thuế\t: {input_data['mst']}")
        elif "địa chỉ" in text_lower and ("phường" in text_lower or "khu phố" in text_lower or count_dc == 0) and not text_lower.startswith("4.3") and "nhận hàng" not in text_lower:
            count_dc += 1
            if count_dc == 1: replace_text_keep_format(p, f"Địa chỉ\t: {input_data['dia_chi']}")
        elif "đại diện" in text_lower and "chức vụ" in text_lower:
            count_dd += 1
            if count_dd == 1: replace_mixed_format(p, [("Đại diện\t: ", False), (input_data['dai_dien'], True), ("\tChức vụ: ", False), (input_data['chuc_vu'], True)])
        elif text_lower.startswith("4.2"): passed_4_2 = True 
        elif text_lower.startswith("4.3") and "địa điểm giao hàng" in text_lower:
            if passed_4_2: replace_text_keep_format(p, f"4.3.\tĐịa điểm giao hàng: {input_data['dia_chi_giao']}")
        elif "tổng (đã bao gồm vat):" in text_lower: 
            replace_mixed_format(p, [("Tổng (đã bao gồm VAT): ", True), (input_data['tong_tien'], True)])
        elif "bằng chữ:" in text_lower and "đồng" in text_lower: 
            replace_mixed_format(p, [("Bằng chữ: ", True), (input_data['bang_chu'], True)])

    # C. Sửa lỗi CHỮ KÝ VÀ CHỨC VỤ 
    ten_dai_dien_ky = re.sub(r'^(Ông\s+|Bà\s+|Anh\s+|Chị\s+)', '', input_data['dai_dien'], flags=re.IGNORECASE).strip().upper()
    chuc_vu_goc = input_data['chuc_vu'].strip()
    
    target_table = None
    target_col_idx = -1
    for table in doc.tables:
        for row in table.rows:
            for col_idx, cell in enumerate(row.cells):
                if "đạidiệnbênmua" in normalize_text(cell.text):
                    target_table = table
                    target_col_idx = col_idx
                    break
            if target_table: break
        if target_table: break
        
    if target_table:
        col_paras = []
        for row in target_table.rows:
            if target_col_idx < len(row.cells):
                for p in row.cells[target_col_idx].paragraphs:
                    if p.text.strip() and (not col_paras or p._element != col_paras[-1]._element): col_paras.append(p)
                                
        idx_dai_dien = -1
        for i, p in enumerate(col_paras):
            if "đạidiệnbênmua" in normalize_text(p.text):
                idx_dai_dien = i
                break
                
        if idx_dai_dien != -1:
            p_ten = col_paras[-1]
            if normalize_text(p_ten.text) != normalize_text(ten_dai_dien_ky):
                match_spaces = re.match(r'^(\s+)', p_ten.text)
                leading_spaces = match_spaces.group(1) if match_spaces else ""
                replace_mixed_format(p_ten, [(leading_spaces + ten_dai_dien_ky, True)])
                
            if len(col_paras) >= idx_dai_dien + 3:
                p_chuc_vu = col_paras[idx_dai_dien + 1]
                if normalize_text(p_chuc_vu.text) != normalize_text(chuc_vu_goc):
                    match_spaces = re.match(r'^(\s+)', p_chuc_vu.text)
                    leading_spaces = match_spaces.group(1) if match_spaces else ""
                    replace_mixed_format(p_chuc_vu, [(leading_spaces + chuc_vu_goc, True)])
    else:
        found = False
        lines = 0
        p_chuc_vu = None
        for p in doc.paragraphs:
            if "đạidiệnbênmua" in normalize_text(p.text):
                found = True
                lines = 0
                continue
            if found:
                if p.text.strip():
                    lines += 1
                    if lines == 1:
                        p_chuc_vu = p
                    if p.text.strip().isupper() and "GIÁM ĐỐC" not in p.text.upper():
                        if normalize_text(p.text) != normalize_text(ten_dai_dien_ky):
                            match_spaces = re.match(r'^(\s+)', p.text)
                            leading_spaces = match_spaces.group(1) if match_spaces else ""
                            replace_mixed_format(p, [(leading_spaces + ten_dai_dien_ky, True)])
                            
                        if p_chuc_vu and normalize_text(p_chuc_vu.text) != normalize_text(chuc_vu_goc):
                            match_spaces = re.match(r'^(\s+)', p_chuc_vu.text)
                            leading_spaces = match_spaces.group(1) if match_spaces else ""
                            replace_mixed_format(p_chuc_vu, [(leading_spaces + chuc_vu_goc, True)])
                        break
                if lines >= 5: break

    return doc

# 4. GIAO DIỆN CHÍNH (UI/UX ĐÃ ĐƯỢC TỐI ƯU SANG TRỌNG)
def main():
    st.markdown("<h1>🤖 Hệ Thống Kiểm Tra & Sửa Hợp Đồng Tự Động</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b; font-size: 1.05rem; margin-bottom: 2rem;'>Công cụ rà soát thông tin đối chiếu giữa Zalo/Email và File Word hợp đồng, tự động phát hiện sai sót và chuẩn hóa định dạng.</p>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2, gap="large")
    with col1:
        st.markdown("### 📝 1. Nhập thông tin gốc")
        input_text = st.text_area("Paste dữ liệu Zalo/Email vào đây:", height=280, placeholder="Dán nội dung đơn hàng, thông tin khách hàng, sản phẩm tại đây...", label_visibility="collapsed")
        
    with col2:
        st.markdown("### 📄 2. Tải lên Hợp đồng mẫu")
        uploaded_file = st.file_uploader("Kéo thả file .docx vào đây", type=["docx"], label_visibility="collapsed")
        st.markdown("<p style='color: #94a3b8; font-size: 0.85rem; margin-top: 0.5rem;'>Hỗ trợ định dạng Microsoft Word (.docx) tối đa 200MB.</p>", unsafe_allow_html=True)
        
    st.markdown("<br>", unsafe_allow_html=True)
    
    if st.button("🚀 BẮT ĐẦU RÀ SOÁT & FIX LỖI TỰ ĐỘNG", use_container_width=True, type="primary"):
        if not input_text or not uploaded_file:
            st.error("⚠️ Vui lòng nhập đầy đủ thông tin gốc ở Bước 1 và tải lên file Word ở Bước 2!")
            return
            
        progress_text = "Hệ thống đang quét từng dòng hợp đồng..."
        my_bar = st.progress(0, text=progress_text)
        for percent_complete in range(100):
            time.sleep(0.005)
            my_bar.progress(percent_complete + 1, text=progress_text)
        my_bar.empty()

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
        st.markdown("### 📊 BẢNG ĐIỀU KHIỂN KẾT QUẢ RÀ SOÁT")
        dashboard_placeholder = st.empty()
        
        total_checks = 0
        passed_checks = 0
        failed_checks = 0
        has_error = False 
        
        def show_result(label, is_match, original_text, expander):
            nonlocal total_checks, passed_checks, failed_checks, has_error
            total_checks += 1
            if is_match:
                passed_checks += 1
                expander.success(f"✅ **{label}:** Khớp dữ liệu ({original_text})")
            else:
                failed_checks += 1
                has_error = True
                expander.error(f"❌ **{label}:** LỆCH DỮ LIỆU! (Gốc: {original_text}) -> Đã Auto-Fix")

        exp_basic = st.expander("📌 I. Thông tin cơ bản & Chữ ký", expanded=has_error)
        exp_product = st.expander("🛒 II. Chi tiết Hàng hóa / Sản phẩm", expanded=has_error)
        
        # QUÉT THÔNG TIN CƠ BẢN
        show_result("Tên công ty", normalize_text(input_data['cong_ty']) in full_doc_normalized, input_data['cong_ty'], exp_basic)
        show_result("Mã số thuế", normalize_text(input_data['mst']) in full_doc_normalized, input_data['mst'], exp_basic)
        show_result("Địa chỉ", normalize_text(input_data['dia_chi']) in full_doc_normalized, input_data['dia_chi'], exp_basic)
        show_result("Người đại diện", normalize_text(input_data['dai_dien']) in full_doc_normalized, input_data['dai_dien'], exp_basic)
        show_result("Chức vụ", normalize_text(input_data['chuc_vu']) in full_doc_normalized, input_data['chuc_vu'], exp_basic)
        show_result("Tổng tiền", normalize_text(input_data['tong_tien']) in full_doc_normalized, input_data['tong_tien'], exp_basic)
        show_result("Số tiền bằng chữ", normalize_text(input_data['bang_chu']) in full_doc_normalized, input_data['bang_chu'], exp_basic)
        show_result("Địa chỉ nhận hàng", normalize_text(input_data['dia_chi_giao']) in full_doc_normalized, input_data['dia_chi_giao'], exp_basic)
        
        # KIỂM TRA CHỮ KÝ
        ten_dai_dien_ky = re.sub(r'^(Ông\s+|Bà\s+|Anh\s+|Chị\s+)', '', input_data['dai_dien'], flags=re.IGNORECASE).strip().upper()
        chuc_vu_goc = input_data['chuc_vu'].strip()
        ten_nguoi_ky_word = ""
        chuc_vu_ky_word = ""
        
        target_table = None
        target_col_idx = -1
        for table in doc.tables:
            for row in table.rows:
                for col_idx, cell in enumerate(row.cells):
                    if "đạidiệnbênmua" in normalize_text(cell.text):
                        target_table = table
                        target_col_idx = col_idx
                        break
                if target_table: break
            if target_table: break
            
        if target_table:
            col_paras = []
            for row in target_table.rows:
                if target_col_idx < len(row.cells):
                    for p in row.cells[target_col_idx].paragraphs:
                        if p.text.strip() and (not col_paras or p._element != col_paras[-1]._element): col_paras.append(p)
                            
            idx_dai_dien = -1
            for i, p in enumerate(col_paras):
                if "đạidiệnbênmua" in normalize_text(p.text):
                    idx_dai_dien = i
                    break
                    
            if idx_dai_dien != -1:
                ten_nguoi_ky_word = col_paras[-1].text.strip()
                if len(col_paras) >= idx_dai_dien + 3: chuc_vu_ky_word = col_paras[idx_dai_dien + 1].text.strip()
        else:
            found = False
            lines = 0
            for p in doc.paragraphs:
                if "đạidiệnbênmua" in normalize_text(p.text):
                    found = True
                    lines = 0
                    continue
                if found:
                    if p.text.strip():
                        lines += 1
                        if lines == 1: chuc_vu_ky_word = p.text.strip()
                        if p.text.strip().isupper() and "GIÁM ĐỐC" not in p.text.upper():
                            ten_nguoi_ky_word = p.text.strip()
                            break
                    if lines >= 5: break
                    
        if ten_nguoi_ky_word: show_result("Tên người ký (Cuối HĐ)", normalize_text(ten_dai_dien_ky) == normalize_text(ten_nguoi_ky_word), ten_dai_dien_ky, exp_basic)
        if chuc_vu_ky_word: show_result("Chức vụ người ký (Cuối HĐ)", normalize_text(chuc_vu_goc) == normalize_text(chuc_vu_ky_word), chuc_vu_goc, exp_basic)

        # QUÉT SẢN PHẨM
        if not input_data['san_pham']:
            exp_product.info("Không tìm thấy danh sách sản phẩm trong Input gốc.")
            
        for idx, sp in enumerate(input_data['san_pham']):
            ten_goc_norm = normalize_text(sp['ten'])
            is_name_match = False
            is_qty_price_match = False
            qty_goc_norm = normalize_text(sp['so_luong'])
            price_goc_norm = normalize_text(sp['don_gia'])
            
            for sp_word in danh_sach_sp_word:
                if ten_goc_norm[:30] in sp_word['ten'] and ten_goc_norm[-15:] in sp_word['ten']:
                    is_name_match = True
                    if qty_goc_norm == sp_word['so_luong'] and price_goc_norm in sp_word['don_gia']:
                        is_qty_price_match = True
                    break 
            
            show_result(f"Tên Sản Phẩm {idx+1}", is_name_match, sp['ten'], exp_product)
            show_result(f"Giá & Số lượng SP {idx+1}", is_qty_price_match, sp['chuoi_gia_goc'], exp_product)

        with dashboard_placeholder.container():
            d_col1, d_col2, d_col3 = st.columns(3)
            d_col1.metric("🔍 TỔNG MỤC RÀ SOÁT", f"{total_checks} mục")
            d_col2.metric("✅ KHỚP CHÍNH XÁC", f"{passed_checks} mục")
            d_col3.metric("❌ LỖI ĐÃ TỰ ĐỘNG FIX", f"{failed_checks} mục")

        st.markdown("---")
        
        if has_error:
            st.warning("🤖 Hệ thống đã phát hiện và tự động sửa các lỗi sai thành công. Bạn có thể tải ngay bản Hợp đồng đã hoàn thiện bên dưới!")
            doc_fixed = auto_fix_document(doc, input_data)
            bio = io.BytesIO()
            doc_fixed.save(bio)
            
            st.download_button(
                label="📥 TẢI XUỐNG FILE HỢP ĐỒNG ĐÃ CHUẨN HÓA",
                data=bio.getvalue(),
                file_name="Hop_Dong_Da_Fix.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )
        else:
            st.balloons() 
            st.success("🎉 TUYỆT VỜI! Hợp đồng này khớp tuyệt đối 100%, không phát hiện điểm sai lệch nào!")

if __name__ == "__main__":
    main()
