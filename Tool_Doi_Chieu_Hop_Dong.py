import streamlit as st
import docx
import re
import unicodedata
import io 
import difflib

# 1. CẤU HÌNH GIAO DIỆN & CSS CAO CẤP
st.set_page_config(
    page_title="App Review & So Sánh Hợp Đồng", 
    page_icon="🔍", 
    layout="wide"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Be Vietnam Pro', sans-serif !important;
    }
    
    .main {
        background-color: #f8fafc;
        padding: 2rem 1.5rem;
    }
    
    h1 {
        color: #1e293b;
        font-weight: 700;
        font-size: 2.1rem;
        letter-spacing: -0.025em;
        margin-bottom: 0.5rem;
    }
    
    div[data-testid="stFileUploader"] {
        border: 2px dashed #cbd5e1 !important;
        border-radius: 12px !important;
        background-color: #ffffff !important;
        padding: 1.25rem !important;
        transition: all 0.3s ease;
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: #10b981 !important;
        background-color: #f1f5f9 !important;
    }
    
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
        border: none !important;
        border-radius: 12px !important;
        font-weight: 600 !important;
        font-size: 1.05rem !important;
        padding: 0.75rem 1.5rem !important;
        color: white !important;
        box-shadow: 0 4px 12px rgba(16, 185, 129, 0.3) !important;
        transition: all 0.3s ease !important;
    }
    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #059669 0%, #047857 100%) !important;
        box-shadow: 0 6px 16px rgba(16, 185, 129, 0.4) !important;
        transform: translateY(-1px);
    }
    
    [data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        padding: 1.25rem;
        border-radius: 16px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.02);
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
    
    details {
        border: 1px solid #e2e8f0 !important;
        border-radius: 12px !important;
        background-color: #ffffff !important;
        margin-bottom: 1rem;
    }
    
    .diff-add {
        background-color: #dcfce7;
        color: #166534;
        padding: 2px 6px;
        border-radius: 4px;
        font-weight: 600;
    }
    .diff-del {
        background-color: #fee2e2;
        color: #991b1b;
        padding: 2px 6px;
        border-radius: 4px;
        text-decoration: line-through;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# 2. HÀM ĐỌC NỘI DUNG WORD (Paragraphs & Tables)
def read_docx_content(uploaded_file):
    doc = docx.Document(uploaded_file)
    full_text = []
    
    # 1. Đọc các đoạn văn
    for p in doc.paragraphs:
        if p.text.strip():
            full_text.append(p.text.strip())
            
    # 2. Đọc các bảng biểu
    for table in doc.tables:
        for row in table.rows:
            row_content = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if row_content:
                full_text.append(" | ".join(row_content))
                
    return full_text, doc

# 3. HÀM SO SÁNH VĂN BẢN
def compare_contracts(list_orig, list_client):
    diffs = []
    
    matcher = difflib.SequenceMatcher(None, list_orig, list_client)
    
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'replace':
            diffs.append({
                "type": "MODIFIED",
                "orig": " \n".join(list_orig[i1:i2]),
                "client": " \n".join(list_client[j1:j2])
            })
        elif tag == 'delete':
            diffs.append({
                "type": "DELETED",
                "orig": " \n".join(list_orig[i1:i2]),
                "client": ""
            })
        elif tag == 'insert':
            diffs.append({
                "type": "ADDED",
                "orig": "",
                "client": " \n".join(list_client[j1:j2])
            })
            
    return diffs

# 4. GIAO DIỆN CHÍNH
def main():
    st.markdown("<h1>🔍 App Review & So Sánh Hợp Đồng Tự Động</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b; font-size: 1.05rem; margin-bottom: 2rem;'>Tự động phát hiện các điều khoản, con số, thông tin đã bị đối tác âm thầm chỉnh sửa hoặc xóa bỏ.</p>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2, gap="large")
    with col1:
        st.markdown("### 📄 1. File Hợp đồng Gốc (Bản Công Ty)")
        file_orig = st.file_uploader("Tải lên file Word bản gốc gửi đi (.docx)", type=["docx"], key="orig")
        
    with col2:
        st.markdown("### 📝 2. File Hợp đồng Khách gửi lại")
        file_client = st.file_uploader("Tải lên file Word khách đã phản hồi (.docx)", type=["docx"], key="client")
        
    st.markdown("<br>", unsafe_allow_html=True)
    
    if st.button("🚀 BẮT ĐẦU REVIEW & SO SÁNH LỖI", use_container_width=True, type="primary"):
        if not file_orig or not file_client:
            st.error("⚠️ Vui lòng tải lên đủ cả 2 file Hợp đồng để tiến hành so sánh!")
            return
            
        with st.spinner("Hệ thống đang quét chi tiết từng điều khoản..."):
            list_orig, doc_orig = read_docx_content(file_orig)
            list_client, doc_client = read_docx_content(file_client)
            
            diff_results = compare_contracts(list_orig, list_client)
            
            count_added = sum(1 for d in diff_results if d['type'] == 'ADDED')
            count_deleted = sum(1 for d in diff_results if d['type'] == 'DELETED')
            count_modified = sum(1 for d in diff_results if d['type'] == 'MODIFIED')
            total_changes = count_added + count_deleted + count_modified
            
            st.markdown("---")
            st.markdown("### 📊 TỔNG QUAN ĐIỂM CHỈNH SỬA")
            
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("🔍 TỔNG ĐIỂM THAY ĐỔI", f"{total_changes} chỗ")
            m2.metric("🟡 ĐIỀU KHOẢN BỊ SỬA", f"{count_modified} mục")
            m3.metric("🔴 ĐOẠN BỊ XÓA BỎ", f"{count_deleted} mục")
            m4.metric("🟢 ĐOẠN KHÁCH THÊM VÀO", f"{count_added} mục")
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            if total_changes == 0:
                st.balloons()
                st.success("🎉 TUYỆT VỜI! Khách hàng giữ nguyên 100% nội dung bản gốc, không chỉnh sửa bất kỳ chữ nào!")
            else:
                st.markdown("### 🚨 CHI TIẾT CÁC VỊ TRÍ KHÁCH HÀNG ĐÃ SỬA")
                
                exp_mod = st.expander(f"🟡 Các điều khoản / Nội dung đã bị thay đổi ({count_modified})", expanded=True)
                exp_del = st.expander(f"🔴 Các nội dung bản gốc bị khách xóa đi ({count_deleted})", expanded=count_deleted > 0)
                exp_add = st.expander(f"🟢 Các điều khoản khách tự gõ thêm vào ({count_added})", expanded=count_add > 0 if 'count_add' in locals() else False)
                
                # 1. Hiển thị nội dung sửa
                if count_modified > 0:
                    for idx, d in enumerate([x for x in diff_results if x['type'] == 'MODIFIED']):
                        exp_mod.markdown(f"**Vị trí {idx+1}:**")
                        exp_mod.markdown(f"- 🔴 **Bản gốc của bạn:** `{d['orig']}`")
                        exp_mod.markdown(f"- 🟢 **Bản khách đã sửa:** `{d['client']}`")
                        exp_mod.markdown("---")
                else:
                    exp_mod.info("Không có nội dung nào bị thay đổi trực tiếp.")
                    
                # 2. Hiển thị nội dung xóa
                if count_deleted > 0:
                    for idx, d in enumerate([x for x in diff_results if x['type'] == 'DELETED']):
                        exp_del.markdown(f"**Vị trí {idx+1}:** Khách đã xóa câu/đoạn:")
                        exp_del.error(f"❌ `{d['orig']}`")
                else:
                    exp_del.info("Khách không xóa đoạn văn bản nào.")
                    
                # 3. Hiển thị nội dung thêm
                if count_added > 0:
                    for idx, d in enumerate([x for x in diff_results if x['type'] == 'ADDED']):
                        exp_add.markdown(f"**Vị trí {idx+1}:** Khách đã bổ sung thêm:")
                        exp_add.success(f"➕ `{d['client']}`")
                else:
                    exp_add.info("Khách không gõ thêm đoạn văn bản mới nào.")

if __name__ == "__main__":
    main()
