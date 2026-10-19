import streamlit as st
import docx
import re
import difflib
import google.generativeai as genai

# 1. CẤU HÌNH GIAO DIỆN & CSS CAO CẤP
st.set_page_config(
    page_title="App AI Review & Thẩm Định Hợp Đồng", 
    page_icon="🤖", 
    layout="wide"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Be Vietnam Pro', sans-serif !important; }
    .main { background-color: #f8fafc; padding: 2rem 1.5rem; }
    h1 { color: #1e293b; font-weight: 700; font-size: 2rem; margin-bottom: 0.5rem; }
    div[data-testid="stFileUploader"] {
        border: 2px dashed #cbd5e1 !important; border-radius: 12px !important; background-color: #ffffff !important; padding: 1rem !important;
    }
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
        border: none !important; border-radius: 12px !important; font-weight: 600 !important; color: white !important;
    }
    details { border: 1px solid #e2e8f0 !important; border-radius: 12px !important; background-color: #ffffff !important; margin-bottom: 1rem; }
</style>
""", unsafe_allow_html=True)

# 2. CÀI ĐẶT THANH BÊN (SIDEBAR) CHO API KEY
with st.sidebar:
    st.image("https://img.icons8.com/color/96/artificial-intelligence.png", width=64)
    st.title("Cấu hình AI Pháp Chế")
    api_key = st.text_input("Nhập Gemini API Key:", type="password", help="Lấy key miễn phí tại aistudio.google.com")
    st.markdown("---")
    st.info("💡 **Cách hoạt động mới:**\nAI sẽ tự động đọc Bản gốc của công ty để nắm toàn bộ chính sách, sau đó soi bản khách sửa và tự động chốt phương án tối ưu nhất cho bạn.")

# 3. HÀM ĐỌC NỘI DUNG WORD THÀNH TEXT
def read_docx_to_text(uploaded_file):
    doc = docx.Document(uploaded_file)
    full_text = []
    for p in doc.paragraphs:
        if p.text.strip(): full_text.append(p.text.strip())
    for table in doc.tables:
        for row in table.rows:
            row_content = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if row_content: full_text.append(" | ".join(row_content))
    return "\n".join(full_text)

# 4. HÀM GỌI AI PHÂN TÍCH VÀ ĐỀ XUẤT NỘI DUNG CUỐI CÙNG
def ai_analyze_and_resolve(orig_full_text, client_full_text, orig_diff, client_diff):
    if not api_key:
        return "⚠️ Vui lòng nhập Gemini API Key ở thanh bên (Sidebar) để AI phân tích."
    
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.5-flash')
    
    prompt = f"""
    Bạn là một Chuyên gia Pháp chế Doanh nghiệp cấp cao. 
    Tôi cung cấp cho bạn BẢN HỢP ĐỒNG GỐC CỦA CÔNG TY (đây là tiêu chuẩn và chính sách pháp lý chính thức của chúng tôi):
    ---
    {orig_full_text[:10000]}
    ---

    Bây giờ, khách hàng đã gửi lại bản hợp đồng đã qua điều chỉnh. Dưới đây là phần nội dung cụ thể bị thay đổi:
    - NỘI DUNG GỐC CỦA CÔNG TY: "{orig_diff}"
    - NỘI DUNG KHÁCH ĐÃ SỬA: "{client_diff}"

    Hãy phân tích và trả lời rõ ràng bằng tiếng Việt theo cấu trúc sau:
    1. **Đánh giá rủi ro:** Phân tích ngắn gọn xem việc khách hàng sửa đổi điểm này mang lại lợi ích hay rủi ro/bất lợi gì cho công ty chúng ta.
    2. **Khuyên duyệt hay từ chối:** [NÊN DUYỆT / CẦN ĐÀM PHÁN LẠI / TỪ CHỐI]
    3. **Nội dung cuối cùng đề xuất (Safe Clause):** Soạn thảo lại một câu văn bản pháp lý hoàn chỉnh, chuẩn chỉnh nhất (dựa trên tinh thần bản gốc của công ty nhưng có thể uyển chuyển thương lượng nếu hợp lý) để nhân viên copy phản hồi lại khách hàng.
    """
    
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"Lỗi kết nối AI: {str(e)}"

# 5. GIAO DIỆN CHÍNH
def main():
    st.markdown("<h1>🤖 App AI Review & Thẩm Định Hợp Đồng Thông Minh</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b; font-size: 1.05rem; margin-bottom: 2rem;'>AI tự động đọc bản gốc để nắm chính sách, soi lỗi khách sửa và đề xuất nội dung pháp lý cuối cùng.</p>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2, gap="large")
    with col1:
        st.markdown("### 📄 1. File Hợp đồng Gốc (Bản Công Ty)")
        file_orig = st.file_uploader("Tải lên bản gốc (.docx)", type=["docx"], key="orig")
        
    with col2:
        st.markdown("### 📝 2. File Hợp đồng Khách gửi lại")
        file_client = st.file_uploader("Tải lên bản khách phản hồi (.docx)", type=["docx"], key="client")
        
    st.markdown("<br>", unsafe_allow_html=True)
    
    if st.button("🚀 BẮT ĐẦU ĐỂ AI REVIEW & ĐỀ XUẤT NỘI DUNG", use_container_width=True, type="primary"):
        if not file_orig or not file_client:
            st.error("⚠️ Vui lòng tải lên đủ cả 2 file Hợp đồng!")
            return
        if not api_key:
            st.warning("⚠️ Bạn chưa nhập Gemini API Key ở thanh bên trái (Sidebar). Vui lòng nhập để tính năng AI hoạt động.")
            return
            
        with st.spinner("AI đang đọc hiểu toàn bộ bản gốc công ty và đối chiếu với bản khách sửa..."):
            # Đọc toàn bộ text để AI làm ngữ cảnh (Context)
            text_orig_full = read_docx_to_text(file_orig)
            text_client_full = read_docx_to_text(file_client)
            
            # Tách danh sách dòng để tìm điểm khác biệt (Diff)
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
                            st.markdown("**🔴 Nội dung trong Bản gốc Công ty:**")
                            st.code(diff['orig'], language="text")
                        with col_b:
                            st.markdown("**🟢 Nội dung khách đã sửa:**")
                            st.code(diff['client'], language="text")
                            
                        st.markdown("---")
                        st.markdown("#### 🧠 AI Phân tích & Đề xuất Nội dung Cuối Cùng:")
                        
                        # Gọi AI phân tích dựa trực tiếp vào bản gốc
                        ai_feedback = ai_analyze_and_resolve(text_orig_full, text_client_full, diff['orig'], diff['client'])
                        st.markdown(ai_feedback)

if __name__ == "__main__":
    main()
