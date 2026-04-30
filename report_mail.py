import smtplib
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

# Gmail情報を設定する
MY_EMAIL = "sf.sh.10@gmail.com"
APP_PASSWORD = "ivkwzswopzqetoae"
TO_EMAIL = "sf.sh.10@gmail.com"

# 日本語フォントの設定
matplotlib.rc("font", family="Hiragino Sans")

# Excelファイルを読み込む
df = pd.read_excel("/Users/ryo/Desktop/売上データ.xlsx")

# 担当者ごとに売上を集計する
summary = df.groupby("担当者")["売上"].sum().reset_index()
summary.columns = ["担当者", "売上合計"]

# レポートExcelを作成する
output_path = "/Users/ryo/Desktop/自動レポート.xlsx"
summary.to_excel(output_path, index=False, sheet_name="集計")

# グラフを作成して保存する
plt.figure(figsize=(8, 5))
plt.bar(summary["担当者"], summary["売上合計"], color=["#4C72B0", "#DD8452", "#55A868"])
plt.title("担当者別売上グラフ", fontsize=16)
plt.tight_layout()
plt.savefig("/Users/ryo/Desktop/tmp_graph.png")
plt.close()

# ExcelにグラフのPNG画像を埋め込む
wb = load_workbook(output_path)
ws = wb["集計"]
img = Image("/Users/ryo/Desktop/tmp_graph.png")
img.anchor = "E1"
ws.add_image(img)
wb.save(output_path)

print("レポートを作成しました")

# メールを作成する
msg = MIMEMultipart()
msg["From"] = MY_EMAIL
msg["To"] = TO_EMAIL
msg["Subject"] = "【自動送信】売上レポート"

# 売上合計を本文に入れる
total = summary["売上合計"].sum()
body = "お疲れ様です。\n\n本日の売上レポートをお送りします。\n\n売上合計：" + str(total) + "円\n\n詳細は添付ファイルをご確認ください。"
msg.attach(MIMEText(body, "plain"))

# Excelファイルを添付する
with open(output_path, "rb") as f:
    attachment = MIMEBase("application", "octet-stream")
    attachment.set_payload(f.read())
    encoders.encode_base64(attachment)
    attachment.add_header("Content-Disposition", "attachment", filename="売上レポート.xlsx")
    msg.attach(attachment)

# 送信する
with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
    server.login(MY_EMAIL, APP_PASSWORD)
    server.send_message(msg)

print("レポートをメールで送信しました")