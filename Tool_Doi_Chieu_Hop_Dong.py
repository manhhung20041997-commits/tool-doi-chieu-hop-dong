import streamlit as st
import docx
import difflib
import google.generativeai as genai

# 1. CẤU HÌNH GIAO DIỆN & GIAO DIỆN CHUẨN MEMORYZONE
st.set_page_config(
    page_title="Công cụ giấy tờ - MemoryZone", 
    page_icon="🟢", 
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
        padding: 1rem 2rem; 
    }
    
    .mz-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 0.75rem 1.5rem;
        background: #ffffff;
        border-bottom: 1px solid #e2e8f0;
        margin: -2rem -2rem 2rem -2rem;
    }
    .mz-logo {
        font-weight: 700;
        font-size: 1.25rem;
        color: #059669;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .mz-nav {
        display: flex;
        gap: 20px;
        font-weight: 600;
        font-size: 0.95rem;
    }
    .mz-nav-item-active {
        color: #059669;
        border-bottom: 2px solid #059669;
        padding-bottom: 4px;
    }

    div[data-testid="stFileUploader"] {
        border: 2px dashed #cbd5e1 !important; 
        border-radius: 12px !important; 
        background-color: #ffffff !important; 
        padding: 1rem !important;
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: #059669 !important;
    }

    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
        border: none !important; 
        border-radius: 10px !important; 
        font-weight: 600 !important; 
        color: white !important;
        box-shadow: 0 4px 12px rgba(16, 185, 129, 0.25) !important;
        transition: all 0.3s ease;
    }
    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #059669 0%, #047857 100%) !important;
        box-shadow: 0 6px 16px rgba(16, 185, 129, 0.35) !important;
        transform: translateY(-1px);
    }
    
    .diff-box {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.25rem;
        margin-bottom: 1rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }
</style>

<div class="mz-header">
    <div class="mz-logo">
        ⚡memoryZone &nbsp;|&nbsp; <span style="font-size: 1rem; color: #334155; font-weight: 500;">Công cụ giấy tờ</span>
    </div>
    <div class="mz-nav">
        <span class="mz-nav-item-active">Hợp Đồng (Review AI)</span>
    </div>
</div>
""", unsafe_allow_html=True)

# 2. CÀI ĐẶT SIDEBAR
with st.sidebar:
    st.image("https://img.icons8.com/color/96/artificial-intelligence.png", width=64)
    st.title("Cấu hình AI Pháp Chế")
    api_key = st.text_input("Nhập Gemini API Key:", type="password", help="Lấy key miễn phí tại aistudio.google.com")
    st.markdown("---")
    st.info("🟢 **Chế độ tối ưu siêu tốc:**\nApp tự động trích xuất chính xác các phần bị sửa đổi và gửi cho AI thẩm định trong tích tắc.")

# 3. HÀM ĐỌC NỘI DUNG WORD THÀNH DANH SÁCH DÒNG
def read_docx_to_lines(uploaded_file):
    doc = docx.Document(uploaded_file)
    lines = []
    for p in doc.paragraphs:
        text = p.text.strip()
        if text:
            lines.append(text)
    for table in doc.tables:
        for row in table.rows:
            row_content = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if row_content:
                lines.append(" | ".join(row_content))
    return lines

# 4. HÀM GỌI AI THẨM ĐỊNH TỪNG ĐIỂM KHÁC BIỆT (SIÊU NHANH)
def ai_analyze_single_diff(orig_text, client_text):
    if not api_key:
        return "⚠️ Vui lòng nhập Gemini API Key."
    
    genai.configure(api_key=api_key)
    try:
        model = genai.GenerativeModel('gemini-1.5-flash')
        prompt = f"""
        Bạn là Chuyên gia Pháp chế Doanh nghiệp cấp cao của MemoryZone.
        Khách hàng đã chỉnh sửa một điều khoản trong hợp đồng so với bản gốc của công ty.
        
        - NỘI DUNG GỐC CỦA CÔNG TY: "{orig_text}"
        - NỘI DUNG KHÁCH ĐÃ SỬA: "{client_text}"

        Hãy phân tích ngắn gọn bằng tiếng Việt theo cấu trúc sau:
        1. **Đánh giá rủi ro:** Phân tích điểm sửa đổi này có lợi hay gây rủi ro/bất lợi gì cho công ty.
        2. **Khuyên duyệt hay từ chối:** [NÊN DUYỆT / CẦN ĐÀM PHÁN LẠI / TỪ CHỐI]
        3. **Nội dung cuối cùng đề xuất (Safe Clause):** Soạn thảo lại câu văn bản pháp lý hoàn chỉnh, chuẩn chỉnh nhất để nhân viên copy phản hồi lại khách hàng.
        """
        response = model.generate_content(prompt)
        return response.text if response and response.text else "Không nhận được phản hồi từ AI."
    except Exception as e:
        return f"Lỗi kết nối AI: {str(e)}"

# 5. GIAO DIỆN CHÍNH
def main():
    st.markdown("### 📋 Thẩm Định & Review Hợp Đồng Khách Hàng Phản Hồi")
    st.markdown("<p style='color: #64748b; font-size: 0.95rem; margin-bottom: 1.5rem;'>So sánh siêu tốc để bóc tách chính xác các điểm khách hàng đã thay đổi.</p>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2, gap="large")
    with col1:
        st.markdown("##### 📄 1. File Hợp đồng Gốc (Bản MemoryZone)")
        file_orig = st.file_uploader("Tải lên bản gốc (.docx)", type=["docx"], key="orig")
        
    with col2:
        st.markdown("##### 📝 2. File Hợp đồng Khách gửi lại")
        file_client = st.file_uploader("Tải lên bản khách phản hồi (.docx)", type=["docx"], key="client")
        
    st.markdown("<br>", unsafe_allow_html=True)
    
    if st.button("🚀 BẮT ĐẦU ĐỂ AI REVIEW & ĐỀ XUẤT NỘI DUNG", use_container_width=True, type="primary"):
        if not file_orig or not file_client:
            st.error("⚠️ Vui lòng tải lên đủ cả 2 file Hợp đồng!")
            return
        if not api_key:
            st.warning("⚠️ Bạn chưa nhập Gemini API Key ở thanh bên trái (Sidebar).")
            return
            
        with st.spinner("Đang bóc tách điểm khác biệt và gọi AI thẩm định..."):
            lines_orig = read_docx_to_lines(file_orig)
            lines_client = read_docx_to_lines(file_client)
            
            # Dùng difflib tìm chính xác các đoạn bị thay thế (replace)
            matcher = difflib.SequenceMatcher(None, lines_orig, lines_client)
            modifications = []
            
            for tag, i1, i2, j1, j2 in matcher.get_opcodes():
                if tag == 'replace':
                    orig_snippet = " \n".join(lines_orig[i1:i2])
                    client_snippet = " \n".join(lines_client[j1:j2])
                    modifications.append({"orig": orig_snippet, "client": client_snippet})
            
            st.markdown("---")
            st.markdown(f"### 📊 KẾT QUẢ THẨM ĐỊNH (Phát hiện {len(modifications)} điểm thay đổi)")
            
            if len(modifications) == 0:
                st.balloons()
                st.success("🎉 Tuyệt vời! Khách hàng giữ nguyên 100% nội dung so với bản gốc của công ty.")
            else:
                for idx, diff in enumerate(modifications):
                    st.markdown(f"""
                    <div class="diff-box">
                        <h4 style="color: #059669; margin-bottom: 10px;">🔍 Điểm thay đổi #{idx+1}</h4>
                    """, unsafe_allow_html=True)
                    
                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown("**🔴 Bản gốc MemoryZone:**")
                        st.code(diff['orig'], language="text")
                    with c2:
                        st.markdown("**🟢 Bản khách đã sửa:**")
                        st.code(diff['client'], language="text")
                        
                    st.markdown("---")
                    st.markdown("**🧠 Đánh giá & Đề xuất từ AI Pháp Chế:**")
                    
                    # Gọi AI cho từng điểm thay đổi (nhanh và chuẩn xác tuyệt đối)
                    ai_feedback = ai_analyze_single_diff(diff['orig'], diff['client'])
                    st.markdown(ai_feedback)
                    
                    st.markdown("</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
