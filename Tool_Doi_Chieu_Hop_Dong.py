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
    st.info("🟢 **Chế độ giao diện MemoryZone:**\nTải lên bản gốc và bản khách sửa để AI tự động trích xuất toàn diện, phát hiện điểm lệch và đề xuất phương án tối ưu trong 1 chạm.")

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

# 4. HÀM GỌI AI PHÂN TÍCH TOÀN DIỆN (1 LẦN GỌI DUY NHẤT - AN TOÀN, KHÔNG TIMEOUT)
def ai_review_full_contract(orig_text, client_text):
    if not api_key:
        return "⚠️ Vui lòng nhập Gemini API Key ở thanh bên (Sidebar) để AI phân tích."
    
    genai.configure(api_key=api_key)
    models_to_try = ['gemini-1.5-flash', 'gemini-1.5-pro', 'gemini-pro', 'gemini-2.0-flash']
    
    prompt = f"""
    Bạn là một Chuyên gia Pháp chế Doanh nghiệp cấp cao của công ty MemoryZone.
    Dưới đây là BẢN HỢP ĐỒNG GỐC CỦA CÔNG TY (đóng vai trò chuẩn hóa chính sách pháp lý):
    ---
    {orig_text[:15000]}
    ---

    Và đây là BẢN HỢP ĐỒNG DO KHÁCH HÀNG/ĐỐI TÁC ĐÃ CHỈNH SỬA:
    ---
    {client_text[:15000]}
    ---

    Hãy thực hiện phân tích đối chiếu chuyên sâu và trả lời bằng tiếng Việt theo cấu trúc rõ ràng:
    1. **Tổng quan nhận định:** Khách hàng có thay đổi nhiều so với bản gốc không? Có điểm nào nguy hiểm cốt lõi không?
    2. **Chi tiết các điểm khác biệt & Đánh giá rủi ro:** Liệt kê các điều khoản mà khách hàng đã âm thầm sửa đổi (thời hạn thanh toán, đơn giá, phạt vi phạm, tài khoản ngân hàng, v.v.), phân tích rõ rủi ro/bất lợi cho công ty.
    3. **Nội dung cuối cùng đề xuất (Safe Clause):** Soạn thảo lại câu văn bản pháp lý hoàn chỉnh, chuẩn chỉnh nhất theo đúng tinh thần bản gốc để nhân viên copy phản hồi lại khách hàng ngay lập tức.
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
            
    return f"⚠️ Lỗi kết nối AI: {last_error}"

# 5. GIAO DIỆN CHÍNH
def main():
    st.markdown("### 📋 Thẩm Định & Review Hợp Đồng Khách Hàng Phản Hồi")
    st.markdown("<p style='color: #64748b; font-size: 0.95rem; margin-bottom: 1.5rem;'>So sánh toàn diện giữa bản hợp đồng gốc của MemoryZone và bản khách hàng chỉnh sửa để bóc tách rủi ro.</p>", unsafe_allow_html=True)
    
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
            
        with st.spinner("AI đang đọc hiểu toàn bộ 2 file hợp đồng và lập báo cáo thẩm định..."):
            text_orig_full = read_docx_to_text(file_orig)
            text_client_full = read_docx_to_text(file_client)
            
            # Gọi AI phân tích toàn diện 1 lần duy nhất
            ai_report = ai_review_full_contract(text_orig_full, text_client_full)
            
            st.markdown("---")
            st.markdown("### 📊 BÁO CÁO THẨM ĐỊNH & ĐỀ XUẤT PHÁP LÝ TỪ AI")
            
            st.markdown(f"""
            <div class="report-card">
                {ai_report}
            </div>
            """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
