import streamlit as st
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt

# 日本語フォントの設定
matplotlib.rc("font", family="Hiragino Sans")

# アプリのタイトル
st.title("売上自動集計ツール")
st.write("Excelファイルをアップロードすると自動で集計・グラフ化します")

# ファイルアップロード欄
uploaded_file = st.file_uploader("Excelファイルを選択してください", type=["xlsx"])

if uploaded_file is not None:
    # Excelを読み込む
    df = pd.read_excel(uploaded_file)

    st.subheader("アップロードされたデータ")
    st.dataframe(df)

    # 担当者ごとに集計する
    summary = df.groupby("担当者")["売上"].sum().reset_index()
    summary.columns = ["担当者", "売上合計"]

    st.subheader("担当者別売上集計")
    st.dataframe(summary)

    # 合計を表示する
    total = summary["売上合計"].sum()
    st.metric("売上合計", str(total) + "円")

    # グラフを表示する
    st.subheader("売上グラフ")
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.bar(summary["担当者"], summary["売上合計"], color=["#4C72B0", "#DD8452", "#55A868"])
    ax.set_title("担当者別売上グラフ", fontsize=16)
    st.pyplot(fig)