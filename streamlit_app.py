import streamlit as st
import fun as myfun

navisFun=myfun.navisComApi()
st.title("🎈 My new app")
st.write(
    "Let's start building! For help and inspiration, head over to [docs.streamlit.io](https://docs.streamlit.io/)."
)

test_button = st.button("test")
if test_button:
    doc=navisFun.doc_navis_com_api_data()
    print(doc.name)
