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
    
    .report-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.02);
        margin-top: 1rem;
    }
    .diff-box {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.25rem;
        margin-bottom: 1rem;
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
    st.info("🟢 **Chế độ tối ưu siêu tốc:**\nApp gộp tất cả các điểm thay đổi và gửi 1 request duy nhất đến AI, giúp loại bỏ hoàn toàn lỗi timeout và 503.")

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

# 4. HÀM GỌI AI GỘP CHUNG TẤT CẢ ĐIỂM KHÁC BIỆT (1 REQUEST DUY NHẤT)
def ai_analyze_all_diffs(modifications):
    if not api_key:
        return "⚠️ Vui lòng nhập Gemini API Key."
    
    genai.configure(api_key=api_key)
    
    # Xây dựng danh sách thay đổi gọn gàng
    summary_text = ""
    for idx, diff in enumerate(modifications):
        summary_text += f"\n--- ĐIỂM THAY ĐỔI #{idx+1} ---\n"
        summary_text += f"Bản gốc MemoryZone: {diff['orig']}\n"
        summary_text += f"Bản khách hàng sửa: {diff['client']}\n"

    prompt = f"""
    Bạn là Chuyên gia Pháp chế Doanh nghiệp cấp cao của MemoryZone.
    Dưới đây là danh sách các điểm mà khách hàng đã chỉnh sửa trong hợp đồng so với bản gốc của công ty:
    {summary_text}

    Hãy phân tích lần lượt từng điểm thay đổi trên bằng tiếng Việt theo đúng cấu trúc rõ ràng sau cho mỗi điểm:
    
    ### 🔍 Điểm thay đổi #[Số thứ tự]
    - **Đánh giá rủi ro:** Phân tích ngắn gọn điểm sửa đổi này mang lại lợi ích hay rủi ro/bất lợi gì cho công ty.
    - **Khuyên duyệt hay từ chối:** [NÊN DUYỆT / CẦN ĐÀM PHÁN LẠI / TỪ CHỐI]
    - **Nội dung cuối cùng đề xuất (Safe Clause):** Soạn thảo lại câu văn bản pháp lý hoàn chỉnh, chuẩn chỉnh nhất theo đúng tinh thần bản gốc để nhân viên copy phản hồi lại khách hàng.
    """
    
    models_to_try = ['gemini-1.5-flash', 'gemini-1.5-pro', 'gemini-pro', 'gemini-2.0-flash']
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
            st.warning("⚠️ Bạn chưa nhập Gemini API Key ở thanh bên (Sidebar).")
            return
            
        with st.spinner("Đang bóc tách điểm khác biệt và gọi AI thẩm định siêu tốc..."):
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
                # Hiển thị song song bản gốc và bản khách sửa đối chiếu trước
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
                    st.markdown("</div>", unsafe_allow_html=True)

                st.markdown("---")
                st.markdown("### 🧠 ĐÁNH GIÁ RỦI RO & ĐỀ XUẤT TỪ AI PHÁP CHẾ")
                
                # Gọi AI 1 lần duy nhất cho toàn bộ danh sách thay đổi
                with st.spinner("AI đang tổng hợp báo cáo và viết Safe Clause..."):
                    ai_report = ai_analyze_all_diffs(modifications)
                    st.markdown(f"""
                    <div class="report-card">
                        {ai_report}
                    </div>
                    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
