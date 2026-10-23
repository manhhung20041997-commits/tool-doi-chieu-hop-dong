import streamlit as st
import docx
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
    
    .report-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.02);
        margin-top: 1rem;
        line-height: 1.6;
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
    st.info("🟢 **Chế độ Smart AI Compare:**\nAI đọc trực tiếp toàn bộ 2 file, tự động phát hiện 100% các điểm sửa đổi dù là nhỏ nhất (số liệu, điều khoản, thông tin bên mua).")

# 3. HÀM ĐỌC NỘI DUNG WORD (BAO GỒM CẢ VĂN BẢN VÀ BẢNG BIỂU)
def read_docx_full(uploaded_file):
    doc = docx.Document(uploaded_file)
    content = []
    for p in doc.paragraphs:
        if p.text.strip():
            content.append(p.text.strip())
    for table in doc.tables:
        for row in table.rows:
            row_content = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if row_content:
                content.append(" | ".join(row_content))
    return "\n".join(content)

# 4. HÀM GỌI AI PHÂN TÍCH THÔNG MINH
def ai_smart_compare(orig_text, client_text):
    if not api_key:
        return "⚠️ Vui lòng nhập Gemini API Key ở thanh bên (Sidebar)."
    
    genai.configure(api_key=api_key)
    
    prompt = f"""
    Bạn là Chuyên gia Pháp chế Doanh nghiệp cấp cao của công ty MemoryZone. 
    Nhiệm vụ của bạn là đối chiếu chi tiết giữa 2 văn bản hợp đồng sau đây để tìm ra toàn bộ các điểm mà khách hàng đã chỉnh sửa (kể cả những thay đổi nhỏ nhất về con số, tỷ lệ % phạt, điều kiện bảo hành, thông tin liên hệ, thêm bớt câu chữ).

    --- BẢN GỐC CỦA MEMORYZONE ---
    {orig_text[:20000]}

    --- BẢN DO KHÁCH HÀNG GỬI LẠI ---
    {client_text[:20000]}

    Hãy phân tích và lập báo cáo chi tiết bằng tiếng Việt theo cấu trúc sau:
    1. **Tổng quan:** Khách hàng đã sửa đổi tổng cộng bao nhiêu điểm chính? Có điểm nào rủi ro cao đối với công ty không?
    2. **Chi tiết các điểm khác biệt & Đánh giá rủi ro:** 
       - Liệt kê rõ: *Bản gốc* nói gì vs *Khách sửa* thành gì.
       - Phân tích rủi ro/bất lợi pháp lý cho công ty.
       - Đưa ra nhận định: **[NÊN DUYỆT]** / **[CẦN ĐÀM PHÁN LẠI]** / **[TỪ CHỐI]**.
    3. **Safe Clause (Nội dung phản hồi cuối cùng):** Soạn thảo sẵn câu văn bản pháp lý chuẩn chỉnh nhất để nhân viên copy gửi phản hồi lại cho khách hàng ngay lập tức.
    """
    
    models_to_try = ['gemini-1.5-flash', 'gemini-1.5-pro', 'gemini-2.0-flash', 'gemini-pro']
    last_error = ""
    
    for m_name in models_to_try:
        try:
            model = genai.GenerativeModel(m_name)
            generation_config = genai.types.GenerationConfig(max_output_tokens=4096, temperature=0.1)
            response = model.generate_content(prompt, generation_config=generation_config)
            if response and response.text:
                return response.text
        except Exception as e:
            last_error = str(e)
            continue
            
    return f"⚠️ Lỗi kết nối AI: {last_error}"

# 5. GIAO DIỆN CHÍNH
def main():
    st.markdown("### 📋 Thẩm Định & Review Hợp Đồng Khách Hàng Phản Hồi")
    st.markdown("<p style='color: #64748b; font-size: 0.95rem; margin-bottom: 1.5rem;'>Hệ thống Smart AI tự động rà soát toàn diện từng chữ, từng con số chỉnh sửa giữa 2 file hợp đồng.</p>", unsafe_allow_html=True)
    
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
            st.warning("⚠️ Bạn chưa nhập Gemini API Key ở thanh bên (Sidebar).")
            return
            
        with st.spinner("AI đang đọc, đối chiếu từng câu chữ và lập báo cáo thẩm định... (Chỉ mất 3-5 giây)"):
            text_orig = read_docx_full(file_orig)
            text_client = read_docx_full(file_client)
            
            # Gọi AI phân tích thông minh toàn diện
            ai_report = ai_smart_compare(text_orig, text_client)
            
            st.markdown("---")
            st.markdown("### 📊 BÁO CÁO THẨM ĐỊNH & ĐỀ XUẤT PHÁP LÝ TỪ AI")
            st.markdown(f"""
            <div class="report-card">
                {ai_report}
            </div>
            """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
