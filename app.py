import streamlit as st
import pandas as pd

st.set_page_config(page_title="CodeAlpha Task 1", layout="wide")
st.title("CodeAlpha Task 1 - Web Scraping - Pranathi")
st.write("100 Books from books.toscrape.com")

# Load Excel
df = pd.read_excel("books.xlsx")
st.dataframe(df, use_container_width=True)

st.success(f"Total Books Scraped: {len(df)}")
