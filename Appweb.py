import streamlit as st
import os 
from dotenv import load_dotenv

load_dotenv()

st.title("App Web GitHub")
st.write("AppWeb")

user = os.getenv("DB_Username")
senha = os.getenv("DB_Senha")

st.write(user)
st.write(senha)
