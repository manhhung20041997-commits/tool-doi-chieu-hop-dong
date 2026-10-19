import streamlit as st
import docx
import re
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
    
    details { 
        border: 1px solid #e2e8f0 !important; 
        border-radius: 12px !important; 
        background-color: #ffffff !important; 
        margin-bottom: 1rem; 
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
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
    st.info("🟢 **Chế độ giao diện MemoryZone:**\nTải lên bản gốc và bản khách sửa để AI tự động trích xuất, soi lỗi lệch điều khoản và đề xuất phương án tối ưu.")

# 3. HÀM ĐỌC NỘI DUNG WORD
def read_docx_to_text(uploaded_file):
    doc = docx.Document(uploaded_file)
    full_text = []
    for p in doc.paragraphs:
        if p.text.strip():
            full_text.append(p.text.strip())
    for table in doc.tables:
        for row in table.rows:
            row_content = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if row_content:
                full_text.append(" | ".join(row_content))
    return "\n".join(full_text)

# 4. HÀM GỌI AI VỚI CƠ CHẾ TỰ ĐỘNG CHUYỂN MODEL DỰ PHÒNG (FAIL-SAFE FALLBACK)
def ai_analyze_and_resolve(orig_full_text, client_full_text, orig_diff, client_diff):
    if not api_key:
        return "⚠️ Vui lòng nhập Gemini API Key ở thanh bên (Sidebar) để AI phân tích."
    
    genai.configure(api_key=api_key)
    
    # Danh sách các model sẽ thử lần lượt để tránh tuyệt đối lỗi 404
    models_to_try = ['gemini-1.5-flash', 'gemini-1.5-pro', 'gemini-pro', 'gemini-2.0-flash']
    
    prompt = f"""
    Bạn là một Chuyên gia Pháp chế Doanh nghiệp cấp cao. 
    Dựa vào BẢN HỢP ĐỒNG GỐC CỦA CÔNG TY (tiêu chuẩn pháp lý chính thức):
    ---
    {orig_full_text[:10000]}
    ---

    Khách hàng đã gửi lại bản hợp đồng đã qua điều chỉnh. Dưới đây là phần nội dung cụ thể bị thay đổi:
    - NỘI DUNG GỐC CỦA CÔNG TY: "{orig_diff}"
    - NỘI DUNG KHÁCH ĐÃ SỬA: "{client_diff}"

    Hãy phân tích ngắn gọn bằng tiếng Việt theo cấu trúc sau:
    1. **Đánh giá rủi ro:** Phân tích xem việc khách hàng sửa đổi điểm này mang lại lợi ích hay rủi ro/bất lợi gì cho công ty.
    2. **Khuyên duyệt hay từ chối:** [NÊN DUYỆT / CẦN ĐÀM PHÁN LẠI / TỪ CHỐI]
    3. **Nội dung cuối cùng đề xuất (Safe Clause):** Soạn thảo lại một câu văn bản pháp lý hoàn chỉnh, chuẩn chỉnh nhất để nhân viên copy phản hồi lại khách hàng.
    """
    
    last_error = ""
    for m_name in models_to_try:
        try:
            model = genai.GenerativeModel(m_name)
            response = model.generate_content(prompt)
            if response and response.text:
                return response.text
        except Exception as e:
            last_error = str(e)
            continue
            
    return f"⚠️ Lỗi kết nối AI với mọi model dự phòng. Chi tiết lỗi hệ thống từ Google: {last_error}"

# 5. GIAO DIỆN CHÍNH
def main():
    st.markdown("### 📋 Thẩm Định & Review Hợp Đồng Khách Hàng Phản Hồi")
    st.markdown("<p style='color: #64748b; font-size: 0.95rem; margin-bottom: 1.5rem;'>So sánh chéo giữa bản hợp đồng gốc của MemoryZone và bản khách hàng đã chỉnh sửa để phát hiện rủi ro.</p>", unsafe_allow_html=True)
    
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
            st.warning("⚠️ Bạn chưa nhập Gemini API Key ở thanh bên trái (Sidebar). Vui lòng nhập để tính năng AI hoạt động.")
            return
            
        with st.spinner("Hệ thống đang đối chiếu chính sách MemoryZone và phân tích các điểm chỉnh sửa..."):
            text_orig_full = read_docx_to_text(file_orig)
            text_client_full = read_docx_to_text(file_client)
            
            list_orig = text_orig_full.split('\n')
            list_client = text_client_full.split('\n')
            
            matcher = difflib.SequenceMatcher(None, list_orig, list_client)
            diff_modifications = []
            
            for tag, i1, i2, j1, j2 in matcher.get_opcodes():
                if tag == 'replace':
                    diff_modifications.append({
                        "orig": " \n".join(list_orig[i1:i2]),
                        "client": " \n".join(list_client[j1:j2])
                    })
            
            st.markdown("---")
            st.markdown(f"### 📊 KẾT QUẢ THẨM ĐỊNH TỪ AI (Phát hiện {len(diff_modifications)} điểm thay đổi)")
            
            if len(diff_modifications) == 0:
                st.balloons()
                st.success("🎉 Tuyệt vời! Khách hàng giữ nguyên 100% nội dung so với bản gốc của công ty.")
            else:
                for idx, diff in enumerate(diff_modifications):
                    with st.expander(f"🔍 Điểm thay đổi #{idx+1}: {diff['orig'][:50]}...", expanded=True):
                        col_a, col_b = st.columns(2)
                        with col_a:
                            st.markdown("**🔴 Nội dung trong Bản gốc MemoryZone:**")
                            st.code(diff['orig'], language="text")
                        with col_b:
                            st.markdown("**🟢 Nội dung khách đã sửa:**")
                            st.code(diff['client'], language="text")
                            
                        st.markdown("---")
                        st.markdown("#### 🧠 AI Phân tích & Đề xuất Nội dung Cuối Cùng:")
                        
                        ai_feedback = ai_analyze_and_resolve(text_orig_full, text_client_full, diff['orig'], diff['client'])
                        st.markdown(ai_feedback)

if __name__ == "__main__":
    main()
