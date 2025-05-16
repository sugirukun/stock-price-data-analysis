import pandas as pd
import os
import numpy as np
from datetime import datetime
import japanize_matplotlib
import plotly.graph_objects as go
import plotly.io as pio

# ファイルを読み込むディレクトリとファイル名のパターン
directory = r"C:\Users\rilak\Desktop\株価\株価データ"
# 四半期足のファイルのみをフィルタリング
files = [f for f in os.listdir(directory) if f.endswith('.csv') and '四半期足' in f]

# 結果を保存するディレクトリ
result_dir = r"C:\Users\rilak\Desktop\株価\分析結果"
os.makedirs(result_dir, exist_ok=True)

# ユーザーに日付範囲を問い合わせる
def get_date_range():
    print("分析したい期間を指定してください。")
    
    # デフォルト値を設定
    default_start = '2020-01-01'
    default_end = '2025-01-01'
    
    # 開始日の入力
    while True:
        start_input = input(f"開始日（YYYY-MM-DD形式、例: 2020-01-01）[Enter で {default_start}]: ")
        # デフォルト値を使用
        if start_input == "":
            start_date = default_start
            break
        # 入力値を検証
        try:
            datetime.strptime(start_input, '%Y-%m-%d')
            start_date = start_input
            break
        except ValueError:
            print("無効な日付形式です。YYYY-MM-DD形式で入力してください。")
    
    # 終了日の入力
    while True:
        end_input = input(f"終了日（YYYY-MM-DD形式、例: 2025-01-01）[Enter で {default_end}]: ")
        # デフォルト値を使用
        if end_input == "":
            end_date = default_end
            break
        # 入力値を検証
        try:
            datetime.strptime(end_input, '%Y-%m-%d')
            end_date = end_input
            break
        except ValueError:
            print("無効な日付形式です。YYYY-MM-DD形式で入力してください。")
    
    print(f"分析期間: {start_date} から {end_date} まで")
    return start_date, end_date

# ユーザーから日付範囲を取得
start_date, end_date = get_date_range()

# 四半期足データを処理して変動率を計算する関数
def process_quarterly_data(file_path, start_date, end_date):
    try:
        # データの読み込み
        df = pd.read_csv(file_path, index_col=0, parse_dates=True)
        
        # タイムゾーン情報を削除
        if df.index.tz is not None:
            df.index = df.index.tz_localize(None)
        
        # インデックスがDatetimeIndexであることを確認
        if not isinstance(df.index, pd.DatetimeIndex):
            print(f"警告: {file_path} のインデックスがDatetimeIndexではありません。スキップします。")
            return None
        
        # データが空でないことを確認
        if df.empty:
            print(f"警告: {file_path} のデータが空です。スキップします。")
            return None
            
        # 必須カラムがあることを確認
        if 'Close' not in df.columns:
            print(f"警告: {file_path} に 'Close' カラムがありません。スキップします。")
            return None
        
        # まずデータを期間でフィルタリング
        df = df[(df.index >= start_date) & (df.index <= end_date)]
        
        if df.empty:
            print(f"警告: {file_path} の指定期間内にデータがありません。スキップします。")
            return None
        
        # すでに四半期足データなので、そのまま使用
        quarterly_data = df.copy()
        
        # 開始日の値を取得
        if quarterly_data.empty:
            print(f"警告: {file_path} の四半期データが空です。スキップします。")
            return None
            
        first_value = quarterly_data.iloc[0]['Close']
        
        # 変動率の計算（開始日を1.0として正規化）
        normalized_data = quarterly_data.copy()
        normalized_data['Normalized'] = quarterly_data['Close'] / first_value
        normalized_data['Actual_Close'] = quarterly_data['Close']  # 実際の終値も保存
        
        print(f"処理完了: {file_path} - 四半期データ数: {len(normalized_data)}")
        return normalized_data
        
    except Exception as e:
        print(f"エラー発生: {file_path} の処理中にエラーが発生しました: {str(e)}")
        return None

# 全銘柄の変動率データを格納するリスト
all_normalized_data = []
all_actual_close_data = []
stock_names = []

# 各ファイルを処理
print(f"四半期足ファイル総数: {len(files)}")
for file in files:
    file_path = os.path.join(directory, file)
    stock_name = file.split('.')[0]  # ファイル名から銘柄名を取得
    
    print(f"処理中: {file}")
    normalized_data = process_quarterly_data(file_path, start_date, end_date)
    
    if normalized_data is not None and not normalized_data.empty:
        # 必要なカラムがあることを確認
        if 'Normalized' in normalized_data.columns and 'Actual_Close' in normalized_data.columns:
            all_normalized_data.append(normalized_data['Normalized'])
            all_actual_close_data.append(normalized_data['Actual_Close'])
            stock_names.append(stock_name)
            print(f"追加: {stock_name} - データ数: {len(normalized_data)}")
        else:
            print(f"警告: {file_path} の処理結果に必要なカラムがありません。")

print(f"処理完了した銘柄数: {len(all_normalized_data)}")

# 同じ日付インデックスで全銘柄のデータを結合
if all_normalized_data:
    # 変動率データの結合
    combined_data = pd.concat(all_normalized_data, axis=1)
    combined_data.columns = stock_names
    
    # 実際の終値データの結合
    actual_close_data = pd.concat(all_actual_close_data, axis=1)
    actual_close_data.columns = stock_names
    
    print(f"結合後のデータ形状: {combined_data.shape}")
    print(f"結合後のインデックス: {combined_data.index[:5]} ... {combined_data.index[-5:]}")
    
    # ファイル名に期間情報を含める
    period_str = f"{start_date.replace('-', '')}_{end_date.replace('-', '')}"
    
    # CSV形式での保存
    csv_path = os.path.join(result_dir, f'四半期足株価変動率比較_{period_str}.csv')
    combined_data.to_csv(csv_path)
    
    # インタラクティブグラフの作成（Plotly使用）
    fig = go.Figure()
    
    for column in combined_data.columns:
        # 各銘柄のデータをプロット
        fig.add_trace(
            go.Scatter(
                x=combined_data.index,
                y=combined_data[column],
                mode='lines+markers',
                name=column,
                hovertemplate=
                '<b>%{fullData.name}</b><br>' +
                '日付: %{x}<br>' +
                '変動率: %{y:.2f}<br>' +
                '終値: ' + actual_close_data[column].astype(str) +
                '<extra></extra>'
            )
        )
    
    # グラフのレイアウト設定
    fig.update_layout(
        title=f'四半期足変動率比較（{start_date}～{end_date}、開始日基準：1.0）',
        xaxis_title='四半期',
        yaxis_title='変動率（開始日=1.0）',
        hovermode='closest',
        template='plotly_white'
    )
    
    # HTML形式で保存（インタラクティブなグラフ）
    html_path = os.path.join(result_dir, f'四半期足株価変動率比較_{period_str}.html')
    fig.write_html(html_path)
    
    # 従来のPNG形式でも保存（オプション）
    png_path = os.path.join(result_dir, f'四半期足株価変動率比較_{period_str}.png')
    fig.write_image(png_path)
    
    print(f"分析完了。結果は {result_dir} に保存されました：")
    print(f" - インタラクティブグラフ: {html_path}")
    print(f" - 画像: {png_path}")
    print(f" - CSV: {csv_path}")
else:
    print("指定した期間内のデータが見つかりませんでした。")