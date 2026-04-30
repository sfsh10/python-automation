import requests
from bs4 import BeautifulSoup
import pandas as pd

# スクレイピング練習用サイト（書籍一覧）
url = "https://books.toscrape.com/"
response = requests.get(url)
response.encoding = "utf-8"
soup = BeautifulSoup(response.content, "html.parser", from_encoding="utf-8")

# 本の情報を全部取得する
books = soup.find_all("article", class_="product_pod")

# データを格納するリストを用意する
data = []

for book in books:
    title = book.find("h3").find("a")["title"]
    price = book.find("p", class_="price_color").text.replace("Â", "").strip()
    data.append({"タイトル": title, "価格": price})

# DataFrameに変換する
df = pd.DataFrame(data)

print(df)

# Excelに保存する
df.to_excel("/Users/ryo/Desktop/書籍一覧.xlsx", index=False)
print("\n書籍一覧.xlsx をデスクトップに保存しました")