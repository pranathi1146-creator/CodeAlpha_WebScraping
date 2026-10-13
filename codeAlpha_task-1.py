import requests
from bs4 import BeautifulSoup
import openpyxl
import random
from openpyxl.styles import Font, PatternFill, Alignment

books = []
# scrape 2 pages as example - you can increase
for page in range(1, 6):
    url = f"https://books.toscrape.com/catalogue/page-{page}.html"
    r = requests.get(url)
    soup = BeautifulSoup(r.text, "html.parser")

    for book in soup.select("article.product_pod"):
        title = book.h3.a["title"]
        price = "₹" + str(random.randint(100, 999))
        rating_word = book.p["class"][1]
        rating_num = {"One":1, "Two":2, "Three":3, "Four":4, "Five":5}[rating_word]
        availability = book.select_one(".instock.availability").text.strip()
        # make availability short
        availability = "In stock" if "In stock" in availability else "Out of stock"

        # make stars like ★★★☆☆ (4.0)
        full = "★" * rating_num
        empty = "☆" * (5 - rating_num)
        rating = f"{full}{empty} ({float(rating_num):.1f})"

        books.append([title, price, rating, availability])

# create Excel with your model
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Books"

headers = ["Title", "Price (₹)", "Rating", "Availability"]
header_fill = PatternFill("solid", fgColor="1F4E78")
header_font = Font(color="FFFFFF", bold=True, size=12)

for col, h in enumerate(headers, 1):
    c = ws.cell(row=1, column=col, value=h)
    c.fill = header_fill
    c.font = header_font
    c.alignment = Alignment(horizontal="center", vertical="center")

golden_font = Font(color="EAB308", size=11, bold=True) # golden like you want
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)

for r, (title, price, rating, avail) in enumerate(books, 2):
    ws.cell(row=r, column=1, value=title).alignment = left
    ws.cell(row=r, column=2, value=price).alignment = center
    c3 = ws.cell(row=r, column=3, value=rating)
    c3.font = golden_font
    c3.alignment = center
    ws.cell(row=r, column=4, value=avail).alignment = center

ws.column_dimensions['A'].width = 50
ws.column_dimensions['B'].width = 15
ws.column_dimensions['C'].width = 22
ws.column_dimensions['D'].width = 16
ws.freeze_panes = "A2"
ws.auto_filter.ref = ws.dimensions

wb.save("books.xlsx")
print(f"Done! {len(books)} books saved in books.xlsx with Title, Price, Rating, Availability")