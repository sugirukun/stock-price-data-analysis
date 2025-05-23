import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np
from datetime import datetime
import japanize_matplotlib
import plotly.graph_objects as go
import plotly.subplots as sp

# ファイルを読み込むディレクトリとファイル名のパターン
directory = r"C:\Users\rilak\Desktop\株価\株価データ"
# 週足のファイルのみをフィルタリング
files = [f for f in os.listdir(directory) if f.endswith('.csv') and '週足' in f]

# 結果を保存するディレクトリ
result_dir = r"C:\Users\rilak\Desktop\株価\分析結果"
os.makedirs(result_dir, exist_ok=True)

# ユーザーに日付範囲を問い合わせる
def get_date_range():
    print("分析したい期間を指定してください。")
    
    # デフォルト値を設定（週足なので中期間に設定）
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

# テクニカル指標を追加する関数（週足用）
def add_technical_indicators(df):
    # 元のデータのコピーを作成
    result = df.copy()
    
    # 週足用の移動平均線
    result['MA4'] = df['Close'].rolling(window=4).mean()      # 約1ヶ月
    result['MA13'] = df['Close'].rolling(window=13).mean()   # 約3ヶ月（四半期）
    result['MA26'] = df['Close'].rolling(window=26).mean()   # 約6ヶ月（半期）
    result['MA52'] = df['Close'].rolling(window=52).mean()   # 約1年
    
    # 相対力指数 (RSI) - 14週間
    delta = df['Close'].diff()
    gain = delta.where(delta > 0, 0)
    loss = -delta.where(delta < 0, 0)
    avg_gain = gain.rolling(window=14).mean()
    avg_loss = loss.rolling(window=14).mean()
    rs = avg_gain / avg_loss
    result['RSI'] = 100 - (100 / (1 + rs))
    
    # MACD（移動平均収束拡散法）- 週足用パラメータ
    ema12 = df['Close'].ewm(span=12).mean()
    ema26 = df['Close'].ewm(span=26).mean()
    result['MACD'] = ema12 - ema26
    result['MACD_Signal'] = result['MACD'].ewm(span=9).mean()
    result['MACD_Histogram'] = result['MACD'] - result['MACD_Signal']
    
    # ボリンジャーバンド - 20週移動平均線と標準偏差
    ma20 = df['Close'].rolling(window=20).mean()
    std20 = df['Close'].rolling(window=20).std()
    
    # 中心線（移動平均線）
    result['BB_中心線'] = ma20
    
    # 上側バンド（+1σ, +2σ, +3σ）
    result['BB_上側1σ'] = ma20 + (std20 * 1)
    result['BB_上側2σ'] = ma20 + (std20 * 2)
    result['BB_上側3σ'] = ma20 + (std20 * 3)
    
    # 下側バンド（-1σ, -2σ, -3σ）
    result['BB_下側1σ'] = ma20 - (std20 * 1)
    result['BB_下側2σ'] = ma20 - (std20 * 2)
    result['BB_下側3σ'] = ma20 - (std20 * 3)
    
    # ストキャスティクス - 週足用
    low_min = df['Low'].rolling(window=14).min()
    high_max = df['High'].rolling(window=14).max()
    result['%K'] = 100 * (df['Close'] - low_min) / (high_max - low_min)
    result['%D'] = result['%K'].rolling(window=3).mean()
    
    # 出来高移動平均 - 週足用
    result['Volume_MA'] = df['Volume'].rolling(window=20).mean()
    
    # 週足特有の指標：月間高値・安値
    result['Monthly_High'] = df['High'].rolling(window=4).max()  # 約1ヶ月の高値
    result['Monthly_Low'] = df['Low'].rolling(window=4).min()   # 約1ヶ月の安値
    
    # 長期トレンド指標：52週高値・安値からの位置
    result['52W_High'] = df['High'].rolling(window=52).max()
    result['52W_Low'] = df['Low'].rolling(window=52).min()
    result['52W_Position'] = (df['Close'] - result['52W_Low']) / (result['52W_High'] - result['52W_Low']) * 100
    
    return result

# ユーザーから日付範囲を取得
start_date, end_date = get_date_range()

# 週足データを処理して変動率と指標を計算する関数
def process_weekly_data(file_path, start_date, end_date):
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
        required_columns = ['Open', 'High', 'Low', 'Close']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"警告: {file_path} に必須カラム {missing_columns} がありません。スキップします。")
            return None
        
        # まずデータを期間でフィルタリング
        df = df[(df.index >= start_date) & (df.index <= end_date)]
        
        if df.empty:
            print(f"警告: {file_path} の指定期間内にデータがありません。スキップします。")
            return None
        
        # 既に週足データなので、そのまま使用
        weekly_data = df.copy()
        
        # 休場週を補完（連続した週で前方埋め）
        start_date_dt = pd.to_datetime(start_date)
        end_date_dt = pd.to_datetime(end_date)
        
        # 週の開始を月曜日に設定し、週次の連続インデックスを作成
        # 開始日を含む週の月曜日を取得
        start_monday = start_date_dt - pd.Timedelta(days=start_date_dt.weekday())
        # 終了日を含む週の月曜日を取得
        end_monday = end_date_dt - pd.Timedelta(days=end_date_dt.weekday())
        
        # 週次の連続インデックスを作成（月曜日ベース）
        full_week_range = pd.date_range(start=start_monday, end=end_monday, freq='W-MON')
        
        print(f"期間内の総週数: {len(full_week_range)}")
        print(f"実際のデータ週数: {len(weekly_data)}")
        
        # 連続した週インデックスでリインデックス（前方埋め）
        weekly_data_filled = weekly_data.reindex(full_week_range, method='ffill')
        
        # 最初の週のデータが欠損している場合は後方埋め
        weekly_data_filled = weekly_data_filled.fillna(method='bfill')
        
        # 休場週補完後のデータを使用
        weekly_data = weekly_data_filled
        
        # 開始週の値を取得
        if weekly_data.empty:
            print(f"警告: {file_path} の週足データが空です。スキップします。")
            return None
        
        # テクニカル指標を追加
        weekly_data = add_technical_indicators(weekly_data)
            
        first_value = weekly_data.iloc[0]['Close']
        
        # 変動率の計算（開始週を1.0として正規化）
        normalized_data = weekly_data.copy()
        normalized_data['Normalized'] = weekly_data['Close'] / first_value
        
        # 週次変化率も計算
        normalized_data['Weekly_Change'] = weekly_data['Close'].pct_change() * 100
        
        print(f"処理完了: {file_path} - 週足データ数（休場週補完後）: {len(normalized_data)}")
        return normalized_data
        
    except Exception as e:
        print(f"エラー発生: {file_path} の処理中にエラーが発生しました: {str(e)}")
        return None

# 全銘柄のデータを格納するリスト
all_data = {}
all_normalized_data = []
stock_names = []

# 各ファイルを処理
print(f"週足ファイル総数: {len(files)}")
for file in files:
    file_path = os.path.join(directory, file)
    stock_name = file.split('.')[0]  # ファイル名から銘柄名を取得
    
    print(f"処理中: {file}")
    normalized_data = process_weekly_data(file_path, start_date, end_date)
    
    if normalized_data is not None and not normalized_data.empty:
        # データを保存
        all_data[stock_name] = normalized_data
        all_normalized_data.append(normalized_data['Normalized'])
        stock_names.append(stock_name)
        print(f"追加: {stock_name} - データ数: {len(normalized_data)}")

print(f"処理完了した銘柄数: {len(all_normalized_data)}")

# ファイル名用の期間文字列
period_str = f"{start_date.replace('-', '')}_{end_date.replace('-', '')}"

# 同じ日付インデックスで全銘柄のデータを結合
if all_normalized_data:
    # 変動率データの結合
    combined_data = pd.concat(all_normalized_data, axis=1)
    combined_data.columns = stock_names
    
    print(f"結合前のデータ形状: {combined_data.shape}")
    print(f"結合前のインデックス: {combined_data.index[:5]} ... {combined_data.index[-5:]}")
    
    # 休場週を含む連続した週インデックスを作成
    start_date_dt = pd.to_datetime(start_date)
    end_date_dt = pd.to_datetime(end_date)
    
    # 週の開始を月曜日に設定
    start_monday = start_date_dt - pd.Timedelta(days=start_date_dt.weekday())
    end_monday = end_date_dt - pd.Timedelta(days=end_date_dt.weekday())
    
    # 週次の連続インデックスを作成
    full_week_range = pd.date_range(start=start_monday, end=end_monday, freq='W-MON')
    
    print(f"連続した週数: {len(full_week_range)}")
    print(f"実際のデータ週数: {len(combined_data)}")
    print(f"欠損週数: {len(full_week_range) - len(combined_data)}")
    
    # 連続した週インデックスでリインデックス（前方埋め）
    combined_data_filled = combined_data.reindex(full_week_range, method='ffill')
    
    # 最初の週のデータが欠損している場合は後方埋め
    combined_data_filled = combined_data_filled.fillna(method='bfill')
    
    print(f"休場週補完後のデータ形状: {combined_data_filled.shape}")
    print(f"休場週補完後のインデックス: {combined_data_filled.index[:5]} ... {combined_data_filled.index[-5:]}")
    
    # 元のcombined_dataを更新
    combined_data = combined_data_filled
    
    # CSV形式での保存（休場週補完済み）
    csv_path = os.path.join(result_dir, f'週足株価変動率比較_{period_str}.csv')
    combined_data.to_csv(csv_path)
    
    # 休場週補完の詳細情報をCSVと同じディレクトリに保存
    補完情報_path = os.path.join(result_dir, f'休場週補完情報_{period_str}.txt')
    with open(補完情報_path, 'w', encoding='utf-8') as f:
        f.write(f"休場週補完情報\n")
        f.write(f"==================\n")
        f.write(f"分析期間: {start_date} ～ {end_date}\n")
        f.write(f"対象週数: {len(full_week_range)}\n")
        f.write(f"実際のデータ週数: {len(all_normalized_data[0]) if all_normalized_data else 0}\n")
        f.write(f"補完された週数: {len(full_week_range) - (len(all_normalized_data[0]) if all_normalized_data else 0)}\n")
        f.write(f"補完方法: 前週の値で埋める（前方埋め）\n")
        f.write(f"週の基準: 月曜日開始\n")
        f.write(f"注意: 市場休場週（祝日週など）は前週の値で補完されています\n")
    
    print(f"休場週補完情報を保存しました: {補完情報_path}")
    
    # 1. メイングラフ: 週足変動率比較
    fig = go.Figure()
    
    # 最終週のパフォーマンス順に銘柄を並び替え
    final_performance = combined_data.iloc[-1].sort_values(ascending=False)
    sorted_columns = final_performance.index.tolist()
    
    print("最終週パフォーマンス順（上位から）:")
    for i, column in enumerate(sorted_columns, 1):
        performance_pct = (final_performance[column] - 1) * 100
        print(f"{i:2d}. {column}: {performance_pct:+7.2f}%")
    
    for column in sorted_columns:
        # 日付を数字形式でフォーマット
        formatted_dates = [date.strftime('%Y/%m/%d') for date in combined_data.index]
        
        # 対応する終値データを安全に取得
        try:
            close_prices = all_data[column]['Close']
            # NaNや無効な値を処理
            close_prices_str = close_prices.apply(lambda x: f"{x:.2f}" if pd.notna(x) else "N/A")
        except Exception as e:
            print(f"警告: {column} の終値データ取得エラー: {e}")
            close_prices_str = ["N/A"] * len(combined_data.index)
        
        # 各銘柄のデータをプロット
        fig.add_trace(
            go.Scatter(
                x=combined_data.index,
                y=combined_data[column],
                mode='lines+markers',
                name=column,
                customdata=list(zip(formatted_dates, close_prices_str)),
                hovertemplate=
                '<b>' + column + '</b><br>' +
                '週開始日: %{customdata[0]}<br>' +
                '変動率: %{y:.2%}<br>' +
                '終値: %{customdata[1]}<br>' +
                '<extra></extra>'
            )
        )
    
    # グラフのレイアウト設定
    fig.update_layout(
        title=f'週足変動率比較（{start_date}～{end_date}、開始週基準：100%）',
        xaxis_title='週（月曜日開始）',
        yaxis=dict(
            type='log',  # 対数スケールを使用
            title='変動率（開始週=100%、対数スケール）',
            tickformat='.0%'  # Y軸の目盛りを%形式で表示
        ),
        hovermode='closest',
        template='plotly_white',
        legend_title='銘柄',
        # ホバーツールの表示改善
        hoverlabel=dict(
            bgcolor="rgba(255,255,255,0.95)",
            bordercolor="black",
            font_size=12,
            font_family="Arial",
            align="left",  # 左寄せに変更
            namelength=-1  # 銘柄名の表示文字数制限を無効化
        ),
        # マージンを大幅に拡張してホバーツールの表示領域を確保
        margin=dict(l=80, r=150, t=100, b=80),
        # 追加のレイアウト設定
        width=1200,  # 幅を拡張
        height=700,  # 高さも調整
        updatemenus=[
            # 銘柄選択ドロップダウンメニュー
            {
                'buttons': [
                    {
                        'label': '全銘柄表示',
                        'method': 'update',
                        'args': [{'visible': [True] * len(sorted_columns)}]
                    }
                ] + [
                    {
                        'label': stock,
                        'method': 'update',
                        'args': [{'visible': [i == j for j in range(len(sorted_columns))]}]
                    }
                    for i, stock in enumerate(sorted_columns)
                ],
                'direction': 'down',
                'showactive': True,
                'x': 0.1,
                'y': 1.15,
                'xanchor': 'left'
            },
            # ホバーモード切り替えボタン
            {
                'buttons': [
                    {
                        'label': '個別表示',
                        'method': 'relayout',
                        'args': [{'hovermode': 'closest'}]
                    },
                    {
                        'label': '複数比較',
                        'method': 'relayout', 
                        'args': [{'hovermode': 'x unified'}]
                    }
                ],
                'direction': 'down',
                'showactive': True,
                'x': 0.3,
                'y': 1.15,
                'xanchor': 'left'
            }
        ]
    )
    
    # X軸の日付表示を数字形式に変更
    fig.update_xaxes(tickformat="%Y/%m/%d")
    
    # HTMLとPNG形式で保存
    main_graph_path = os.path.join(result_dir, f'週足株価変動率比較_{period_str}.html')
    fig.write_html(main_graph_path)
    fig.write_image(os.path.join(result_dir, f'週足株価変動率比較_{period_str}.png'))
    
    print(f"分析完了。以下のファイルが作成されました：")
    print(f"1. メイングラフ: {main_graph_path}")
    print(f"2. CSV: {csv_path}")
    print(f"3. 休場週補完情報: {補完情報_path}")
    
    print(f"\nメイングラフを開くには、以下のファイルをブラウザで開いてください：")
    print(f"{main_graph_path}")
else:
    print("指定した期間内のデータが見つかりませんでした。")