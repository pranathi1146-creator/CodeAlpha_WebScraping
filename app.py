import streamlit as st
import pandas as pd

st.set_page_config(page_title="CodeAlpha Task 1 - Books", layout="wide")

st.title("CodeAlpha Task 1 - Web Scraping")
st.write("100 Books Scraped from books.toscrape.com")

# Load your Excel - same as your 6 photos
try:
    df = pd.read_excel("books.xlsx")
except:
    df = pd.read_excel("scraped_data.xlsx")

# Show exactly like your Excel photo
st.dataframe(
    df,
    use_container_width=True,
    height=600
)

st.success(f"Total Books: {len(df)}")

# Download button
with open("books.xlsx", "rb") as f:
    st.download_button(
        "Download books.xlsx",
        data=f,
        file_name="books.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
