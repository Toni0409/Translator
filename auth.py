"""Password gate — chặn truy cập app khi chưa đăng nhập."""
import streamlit as st
from config import APP_PASSWORD


def check_password() -> bool:
    """Trả về True nếu đã đăng nhập. Render màn hình login nếu chưa."""
    if st.session_state.get("authenticated"):
        return True

    st.markdown("""
    <div class='login-box'>
        <div class='login-title'>⬡ Translator</div>
        <div class='login-sub'>PDF & Word — Powered by Gemini — Vi Nguyen</div>
    </div>""", unsafe_allow_html=True)

    _, col_m, _ = st.columns([1, 2, 1])
    with col_m:
        st.markdown("<div style='height:1.2rem'></div>", unsafe_allow_html=True)
        pwd = st.text_input("🔑 Nhập mật khẩu", type="password",
                            key="pwd_input", placeholder="Password...")
        if st.button("Đăng nhập", use_container_width=True):
            if pwd == APP_PASSWORD:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("❌ Sai mật khẩu, thử lại!")
    return False


def logout_button():
    """Nút logout nhỏ ở góc — gọi trong header sau khi đã auth."""
    if st.button("🚪 Logout"):
        st.session_state.pop("authenticated", None)
        st.rerun()
