import streamlit as st
import streamlit.components.v1 as components

# Mở rộng toàn màn hình để giao diện phần mềm hiển thị đầy đủ
st.set_page_config(layout="wide")

# Đọc và hiển thị file qlnt.html
with open("qlnt.html", "r", encoding="utf-8") as f:
    html_code = f.read()

components.html(html_code, height=900, scrolling=True)