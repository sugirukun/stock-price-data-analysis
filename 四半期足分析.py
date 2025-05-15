import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
from datetime import datetime

# 日本語フォントの設定
plt.rcParams['font.family'] = 'Yu Gothic'  # または 'MS Gothic', 'Meiryo' など

# フォルダパスの設定
input_folder = r"C:\Users\rilak\Desktop\株価\株価データ"
output_folder = r"C:\Users\rilak\Desktop\株価\分析結果"

# 出力フォルダが存在しない場合は作成
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# フォルダ内のCSVファイルを取得
all_files = [f for f in os.listdir(input_folder) if f.endswith('.csv')]

# 四半期足データのファイルだけをフィルタリング
# ファイル名に「四半期」「Q」「quarter」などのキーワードが含まれるものを選択
quarterly_files = []
for file in all_files:
    # 四半期足データを示す可能性のあるキーワード
    if any(keyword in file.lower() for keyword in ['四半期', 'q', 'quarter', '3ヶ月', '3か月', '3ヵ月']):
        quarterly_files.append(file)

# 四半期足ファイルが見つからない場合は、ユーザーに確認
if not quarterly_files:
    print("四半期足データと思われるファイルが見つかりませんでした。")
    print("利用可能なCSVファイル:")
    for i, file in enumerate(all_files):
        print(f"{i+1}. {file}")
    
    # ここで分析に使用するファイルを選択するような処理を追加することもできます
    # 例えば、手動で番号を入力してもらうなど
    
    # 今回はサンプルとして、利用可能な全CSVファイルを使用します
    quarterly_files = all_files
    print("すべてのCSVファイルを使用します。")

# データを格納するリスト
stock_data = []
stock_names = []

# ファイルを読み込む
print(f"読み込む四半期足ファイル数: {len(quarterly_files)}")
for file in quarterly_files:
    file_path = os.path.join(input_folder, file)
    try:
        print(f"\n処理中: {file}")
        
        # CSVファイルを読み込む（エンコーディングを試す）
        try:
            df = pd.read_csv(file_path, encoding='shift-jis')
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(file_path, encoding='utf-8')
            except UnicodeDecodeError:
                df = pd.read_csv(file_path, encoding='cp932')
        
        print(f"データ行数: {len(df)}")
        print(f"列名: {df.columns.tolist()}")
        
        # 日付列を特定する
        date_column = None
        for col in df.columns:
            if isinstance(col, str) and ('日付' in col or 'Date' in col or '年月日' in col):
                date_column = col
                break
        
        if date_column is None:
            print(f"警告: {file} には日付列が見つかりませんでした。先頭列を日付として使用します。")
            date_column = df.columns[0]
            
        print(f"使用する日付列: {date_column}")
        
        # 日付を適切な形式に変換
        try:
            df[date_column] = pd.to_datetime(df[date_column])
        except:
            print(f"警告: {file} の日付列をdatetime形式に変換できませんでした。別の形式で試みます。")
            try:
                df[date_column] = pd.to_datetime(df[date_column], format='%Y/%m/%d')
            except:
                try:
                    df[date_column] = pd.to_datetime(df[date_column], format='%Y-%m-%d')
                except:
                    print(f"エラー: {file} の日付列を適切な形式に変換できませんでした。スキップします。")
                    continue
        
        # インデックスとして日付列を設定
        df.set_index(date_column, inplace=True)
        
        # 株価データ列を特定
        close_column = None
        for col in df.columns:
            if isinstance(col, str) and ('終値' in col or 'Close' in col or '株価' in col or '調整後終値' in col):
                close_column = col
                break
        
        if close_column is None:
            # 終値の列名が見つからない場合は、数値データを含む最初の列を使用
            for col in df.columns:
                if pd.api.types.is_numeric_dtype(df[col]):
                    close_column = col
                    break
        
        if close_column is None:
            print(f"警告: {file} には適切な株価データ列が見つかりませんでした。スキップします。")
            continue
        
        print(f"使用する株価データ列: {close_column}")
        
        # 株価データを追加
        stock_data.append(df[close_column])
        
        # ファイル名から銘柄名を取得（拡張子を除く）
        stock_name = os.path.splitext(file)[0]
        stock_names.append(stock_name)
        
        print(f"{file} の読み込みに成功しました。データ数: {len(df)}")
        print(f"期間: {df.index.min()} から {df.index.max()}")
        
    except Exception as e:
        print(f"エラー: {file} の読み込みに失敗しました: {e}")
        import traceback
        traceback.print_exc()

# データが読み込めたかチェック
if len(stock_data) == 0:
    print("読み込めるデータがありませんでした。")
else:
    print("\n読み込んだ銘柄データ:")
    for i, (name, data) in enumerate(zip(stock_names, stock_data)):
        print(f"{i+1}. {name}: データ数 = {len(data)}")
    
    # 横並びでの比較グラフを作成
    plt.figure(figsize=(15, 8))
    
    # 各株価データをプロット
    for i, data in enumerate(stock_data):
        plt.plot(data.index, data.values, label=stock_names[i], linewidth=2)
    
    # グラフの設定
    plt.title('四半期足株価比較', fontsize=16)
    plt.xlabel('日付', fontsize=12)
    plt.ylabel('株価（円）', fontsize=12)
    plt.grid(True, alpha=0.3)
    plt.legend(loc='best')
    
    # x軸の日付表示を調整
    plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y年%m月'))
    plt.gca().xaxis.set_major_locator(mdates.YearLocator())
    plt.xticks(rotation=45)
    
    # グラフを保存
    plt.tight_layout()
    output_path = os.path.join(output_folder, '四半期足株価比較.png')
    plt.savefig(output_path)
    print(f"グラフを保存しました: {output_path}")
    
    # 各株価の変動率を計算して比較グラフを作成
    plt.figure(figsize=(15, 8))
    
    for i, data in enumerate(stock_data):
        if len(data) > 0:
            # 最初の値で正規化して変動率を計算
            normalized_data = data / data.iloc[0] * 100
            plt.plot(normalized_data.index, normalized_data.values, label=stock_names[i], linewidth=2)
    
    # グラフの設定
    plt.title('四半期足株価変動率比較（初期値=100）', fontsize=16)
    plt.xlabel('日付', fontsize=12)
    plt.ylabel('変動率（%）', fontsize=12)
    plt.grid(True, alpha=0.3)
    plt.legend(loc='best')
    
    # グラフに基準線（100%）を追加
    plt.axhline(y=100, color='gray', linestyle='--', alpha=0.5)
    
    # x軸の日付表示を調整
    plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y年%m月'))
    plt.gca().xaxis.set_major_locator(mdates.YearLocator())
    plt.xticks(rotation=45)
    
    # グラフを保存
    plt.tight_layout()
    output_path = os.path.join(output_folder, '四半期足株価変動率比較.png')
    plt.savefig(output_path)
    print(f"変動率比較グラフを保存しました: {output_path}")