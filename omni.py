import streamlit as st
import json
import os

# Enforce bare UI by hiding default elements
st.set_page_config(page_title="Central Portal", layout="centered")

css_overrides = """
<style>
/* Hide Streamlit furniture */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}

/* Global Typography and Brand Color */
html, body, [class*="css"] {
    font-family: Arial, Helvetica, sans-serif !important;
    color: #003366 !important;
}
h1, h2, h3, h4, h5, h6, p, span, label {
    color: #003366 !important;
    font-family: Arial, Helvetica, sans-serif !important;
}

/* 0.75px Gold Divider */
.gold-divider {
    border-top: 0.75px solid #C9AB4C;
    margin: 25px 0;
}

/* Champagne Tool Boxes */
.tool-box {
    background-color: #E8D494;
    padding: 15px;
    margin-bottom: 12px;
    text-align: center;
    display: block;
    text-decoration: none !important;
    color: #003366 !important;
    font-family: Arial, Helvetica, sans-serif;
    font-weight: bold;
    border-radius: 2px;
}
.tool-box:hover {
    background-color: #D6C282;
}
</style>
"""
st.markdown(css_overrides, unsafe_allow_html=True)

# Tool Input UI
tool_name = st.text_input("Tool Name")
tool_url = st.text_input("URL")

if st.button("Add"):
    if tool_name and tool_url:
        data = []
        if os.path.exists("prime_tools.json"):
            with open("prime_tools.json", "r") as f:
                try:
                    data = json.load(f)
                except json.JSONDecodeError:
                    data = []
        
        data.append({"name": tool_name, "url": tool_url})
        
        with open("prime_tools.json", "w") as f:
            json.dump(data, f)
        st.rerun()

# The Divider
st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)

# The Portal Display
if os.path.exists("prime_tools.json"):
    with open("prime_tools.json", "r") as f:
        try:
            tools = json.load(f)
            for tool in tools:
                st.markdown(
                    f'<a href="{tool["url"]}" target="_self" class="tool-box">{tool["name"]}</a>',
                    unsafe_allow_html=True
                )
        except json.JSONDecodeError:
            pass
