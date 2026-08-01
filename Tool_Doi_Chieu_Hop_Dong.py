import streamlit as st
import docx
import re
import unicodedata

def normalize_text(text):
    return unicodedata.normalize('NFC', text)

def read_docx(file):
    doc = docx.Document(file)
    full_text = []
    blocks = []
    tables_list = [] 
    
    for para in doc.paragraphs:
        txt = normalize_text(para.text).strip()
        if txt:
            full_text.append(txt)
            blocks.append(txt)
            
    for table in doc.tables:
        current_table = []
        for row in table.rows:
            row_cells = []
            seen_txt = set()
            for cell in row.cells:
                txt = normalize_text(cell.text).strip()
                row_cells.append(txt)
                if txt and txt not in seen_txt:
                    seen_txt.add(txt)
                    blocks.append(txt)
                    
            if row_cells:
                current_table.append(row_cells)
                row_unique = [c for c in row_cells if c]
                full_text.append(' '.join(row_unique))
                
        tables_list.append(current_table)
                
    unique_blocks = []
    for b in blocks:
        if b not in unique_blocks:
            unique_blocks.append(b)
            
    return ' '.join(full_text), unique_blocks, tables_list

def parse_input(text):
    data = {}
    lines = text.strip().split('\n')
    product_count = 1
    pending_product_name = None

    for line in lines:
        line = normalize_text(line.strip())
        if not line or line == '.':
            continue
            
        if any(char in line for char in ['x', '=', '*']) and any(char.isdigit() for char in line):
            key_name = f"Giá & Số lượng SP {product_count}"
            if pending_product_name:
                 data[f"Tên Sản Phẩm {product_count}"] = pending_product_name
                 pending_product_name = None
            data[key_name] = line
            product_count += 1
        else:
            if ':' in line:
                key, val = line.split(':', 1)
                key = key.strip()
                val = val.strip()
                
                if len(key) > 40: 
                    pending_product_name = line
                elif key.lower() == 'email':
                    continue
                elif val:
                    if val.startswith('Ông/Bà '):
                        val = val.replace('Ông/Bà ', '').strip()
                    elif val.startswith('Ông ') or val.startswith('Bà '):
                        val = val.replace('Ông ', '').replace('Bà ', '').strip()
                        
                    if val.endswith('đ') or val.endswith('đ '):
                        val = val[:-1].strip()
                        
                    data[key] = val
                else:
                    pending_product_name = line
            else:
                if any(w in line.lower() for w in ['triệu', 'đồng', 'chẵn', 'nghìn']):
                    data['Bằng chữ'] = line
                else:
                    pending_product_name = line
                
    return data

def get_conflicts(key, val, anomalies, doc_blocks):
    if not any(c.isdigit() for c in val):
        return []
    key_lower = key.lower()
    key_words = [w for w in re.findall(r'\w+', key_lower) if len(w) >= 3 or w in ['giá', 'phí']]
    
    conflicts = []
    for anomaly in anomalies:
        for block in doc_blocks:
            if anomaly in block:
                block_lower = block.lower()
                if any(kw in block_lower for kw in key_words):
                    if anomaly not in conflicts:
                        conflicts.append(anomaly)
    return conflicts

# Giao diện Web
st.set_page_config(page_title="Tool Đối Chiếu Hợp Đồng", layout="wide")
st.title("Phần Mềm Kiểm Tra Nhanh Hợp Đồng (Chuyên Bên A) 🔍")

col1, col2 = st.columns(2)

with col1:
    st.header("1. Nhập thông tin gốc Bên A")
    input_text = st.text_area("Paste dữ liệu Input của bạn vào đây:", height=400)

with col2:
    st.header("2. Tải file Hợp đồng")
    uploaded_file = st.file_uploader("Chọn file Word (.docx)", type="docx")

st.markdown("---")
if st.button("Tiến hành đối chiếu", type="primary"):
    if not input_text or not uploaded_file:
        st.warning("Vui lòng điền đủ thông tin vào Form và tải file Word lên.")
    else:
        with st.spinner("Đang quét và đối chiếu dữ liệu Bên A..."):
            doc_text, doc_blocks, tables_list = read_docx(uploaded_file) 
            doc_text_lower = doc_text.lower().replace(" ", "") 
            input_data = parse_input(input_text)
            
            # Cắt tách riêng vùng dữ liệu của Bên Mua (Bên A) để kiểm tra độc lập
            ben_a_text = ""
            for block in doc_blocks:
                if 'bên bán (bên b)' in block.lower():
                    break
                ben_a_text += " " + block
            ben_a_text_clean = ben_a_text.lower().replace(" ", "")

            doc_money_patterns = re.findall(r'\d{1,3}(?:\.\d{3})+', doc_text)
            doc_money_set = set(doc_money_patterns)
            
            input_money_set = set()
            for k, val in input_data.items():
                if k.lower() == 'bằng chữ': continue
                val_no_space = val.replace(" ", "")
                nums = re.findall(r'\d{1,3}(?:\.\d{3})+', val_no_space)
                parts = re.split(r'[x*=]', val_no_space)
                for p in parts:
                    clean_p = ''.join(c for c in p if c.isdigit() or c=='.')
                    if clean_p and '.' in clean_p:
                        input_money_set.add(clean_p)
                input_money_set.update(nums)
                
            clean_expected = {num.replace('.', '') for num in input_money_set if num}
            
            anomalies = []
            for doc_money in doc_money_set:
                clean_doc_money = doc_money.replace('.', '')
                if clean_doc_money not in clean_expected:
                    anomalies.append(doc_money)
            
            st.subheader("Kết quả đối chiếu Bên A")
            
            for key, val in input_data.items():
                if not val:
                    continue
                
                conflicts = get_conflicts(key, val, anomalies, doc_blocks)
                if conflicts:
                    st.error(f"❌ **{key}**: BẤT NHẤT DỮ LIỆU! Phát hiện số gõ sai **{', '.join(conflicts)}** ở mục này trong hợp đồng (Gốc yêu cầu: {val})")
                    continue 
                    
                search_val = re.sub(r'\s+', '', val.lower())
                
                # --- KIỂM TRA THÔNG TIN ĐỊNH DANH BÊN A (Tên, MST, Địa chỉ, Đại diện, Chức vụ) ---
                if key.lower() in ['tên công ty', 'mã số thuế', 'địa chỉ', 'người đại diện', 'chức vụ']:
                    if search_val in ben_a_text_clean:
                        # Riêng Chức vụ: check chặt chẽ chống bị dính từ Tổng/Phó nếu input không yêu cầu
                        if key.lower() == 'chức vụ':
                            match_obj = re.search(r'(tổng|phó)' + search_val, ben_a_text_clean)
                            if match_obj:
                                st.error(f"❌ **{key}**: SAI CHỨC VỤ BÊN A! (Yêu cầu '{val}' nhưng thực tế bị gắn kèm Tổng/Phó)")
                                continue
                        st.success(f"✅ **{key}**: Khớp chính xác tại phần Bên A ({val})")
                    else:
                        st.error(f"❌ **{key}**: KHÔNG TÌM THẤY HOẶC BỊ SAI ở phần thông tin Bên A (Gốc yêu cầu: {val})")
                    continue

                # --- XỬ LÝ BẰNG CHỮ ---
                if key.lower() == 'bằng chữ':
                    doc_words_clean = re.sub(r'[^\w\s]', '', doc_text.lower())
                    val_words_clean = re.sub(r'[^\w\s]', '', val.lower())
                    if val_words_clean in doc_words_clean:
                        st.success(f"✅ **{key}**: Khớp chính xác số tiền bằng chữ")
                    else:
                        st.error(f"❌ **{key}**: BỊ SAI CHỮ SO VỚI THỰC TẾ! (Gốc bạn ghi: {val})")
                    continue

                # Xử lý bảng tính giá
                if any(char in search_val for char in ['x', '=', '*']) and key.startswith("Giá"):
                    parts = re.split(r'[x*=]', search_val)
                    parts = [p for p in parts if p]
                    
                    prod_num = key.split()[-1]
                    prod_name_key = f"Tên Sản Phẩm {prod_num}"
                    prod_name = input_data.get(prod_name_key, "")
                    
                    prod_name_clean = re.sub(r'\W+', '', prod_name.lower())
                    search_prefix = prod_name_clean[:40] 
                    
                    target_row_cells = []
                    headers = []
                    
                    if search_prefix:
                        for table in tables_list:
                            if not table: continue
                            head = [re.sub(r'\s+', '', c.lower()) for c in table[0]]
                            for row in table:
                                row_str_clean = re.sub(r'\W+', '', "".join(row).lower())
                                if search_prefix in row_str_clean:
                                    target_row_cells = row
                                    headers = head
                                    break
                            if target_row_cells: break
                    
                    missing_parts = []
                    if target_row_cells:
                        valid_cells = []
                        for i, cell in enumerate(target_row_cells):
                            h_clean = headers[i] if i < len(headers) else ""
                            if 'môtả' in h_clean or 'mota' in h_clean:
                                continue
                            if 'tt' == h_clean or 'stt' == h_clean or (i == 0 and cell.strip().isdigit()):
                                continue
                            c = re.sub(r'\s+', '', cell.lower()).replace("vnđ", "").replace("đ", "")
                            valid_cells.append(c)
                            
                        for part in parts:
                            found = False
                            for c in valid_cells:
                                if len(part) <= 3 and part.isdigit():
                                    if part == c or re.search(r'^\D*' + part + r'\D*$', c):
                                        found = True
                                        break
                                else:
                                    if part in c:
                                        found = True
                                        break
                            if not found:
                                missing_parts.append(part)
                    else:
                        for part in parts:
                            if len(part) <= 3 and part.isdigit():
                                missing_parts.append(f"{part} (Không định vị được bảng)")
                            else:
                                if part not in doc_text_lower:
                                    missing_parts.append(part)
                                
                    if not missing_parts:
                        st.success(f"✅ **{key}**: Khớp toàn bộ thông số ({val})")
                    else:
                        st.error(f"❌ **{key}**: SAI SỐ LƯỢNG HOẶC ĐƠN GIÁ! Không tìm thấy số: {', '.join(missing_parts)} (Gốc: {val})")
                
                else:
                    occurrences = doc_text_lower.count(search_val)
                    if occurrences > 0:
                        st.success(f"✅ **{key}**: Khớp ({val})")
                    else:
                        st.error(f"❌ **{key}**: KHÔNG TÌM THẤY hoặc BỊ SAI (Gốc: {val})")
            
            st.markdown("---")
            st.subheader("🚨 Radar Quét Sai Sót Ẩn")
            
            if anomalies:
                st.error("**CẢNH BÁO: Phát hiện các con số LẠ xuất hiện trong hợp đồng:**")
                for anomaly in anomalies:
                    contexts = [b for b in doc_blocks if anomaly in b]
                    st.markdown(f"### 🛑 Bị sai số: `{anomaly}`")
                    for ctx in contexts:
                        highlighted_ctx = ctx.replace(anomaly, f"**🔥{anomaly}🔥**")
                        st.warning(f"📍 **Vị trí sai sót:** \"... {highlighted_ctx} ...\"")
            else:
                st.info("Bình thường: Không phát hiện số tiền lạ nào bị gõ nhầm trong hợp đồng.")