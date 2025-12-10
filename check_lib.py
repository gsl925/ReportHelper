import streamlit as st
import streamlit_paste_button as spb

st.header("streamlit_paste_button 模組內容檢查")
st.write("以下是 `spb` (streamlit_paste_button) 模組裡面所有可用的東西：")

# 使用 dir() 函式列出模組的所有屬性和方法
st.write(dir(spb))
