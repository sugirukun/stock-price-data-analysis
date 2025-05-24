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
# 日足のファイルのみをフィルタリング
files = [f for f in os.listdir(directory) if f.endswith('.csv') and ('日足' in f or 'daily' in f.lower())]

# 結果を保存するディレクトリ
result_dir = r"C:\Users\rilak\Desktop\株価\分析結果"
os.makedirs(result_dir, exist_ok=True)

# ユーザーに日付範囲を問い合わせる
def get_date_range():
    print("分析したい期間を指定してください。")
    
    # デフォルト値を設定（日足なので短期間に設定）
    default_start = '2022-01-01'
    default_end = '2025-01-01'
    
    # 開始日の入力
    while True:
        start_input = input(f"開始日（YYYY-MM-DD形式、例: 2022-01-01）[Enter で {default_start}]: ")
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

# テクニカル指標を追加する関数（日足用）
def add_technical_indicators(df):
    # 元のデータのコピーを作成
    result = df.copy()
    
    # 日足用の移動平均線
    result['MA5'] = df['Close'].rolling(window=5).mean()      # 1週間
    result['MA25'] = df['Close'].rolling(window=25).mean()   # 約1ヶ月
    result['MA75'] = df['Close'].rolling(window=75).mean()   # 約3ヶ月
    result['MA200'] = df['Close'].rolling(window=200).mean() # 約10ヶ月
    
    # 相対力指数 (RSI) - 14日間
    delta = df['Close'].diff()
    gain = delta.where(delta > 0, 0)
    loss = -delta.where(delta < 0, 0)
    avg_gain = gain.rolling(window=14).mean()
    avg_loss = loss.rolling(window=14).mean()
    rs = avg_gain / avg_loss
    result['RSI'] = 100 - (100 / (1 + rs))
    
    # MACD（移動平均収束拡散法）
    ema12 = df['Close'].ewm(span=12).mean()
    ema26 = df['Close'].ewm(span=26).mean()
    result['MACD'] = ema12 - ema26
    result['MACD_Signal'] = result['MACD'].ewm(span=9).mean()
    result['MACD_Histogram'] = result['MACD'] - result['MACD_Signal']
    
    # ボリンジャーバンド - 20日移動平均線と標準偏差
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
    
    # ストキャスティクス
    low_min = df['Low'].rolling(window=14).min()
    high_max = df['High'].rolling(window=14).max()
    result['%K'] = 100 * (df['Close'] - low_min) / (high_max - low_min)
    result['%D'] = result['%K'].rolling(window=3).mean()
    
    # 出来高移動平均
    result['Volume_MA'] = df['Volume'].rolling(window=20).mean()
    
    return result

# ユーザーから日付範囲を取得
start_date, end_date = get_date_range()

# 日足データを処理して変動率と指標を計算する関数
def process_daily_data(file_path, start_date, end_date):
    try:
        # データの読み込み（日付列をパースしない）
        df = pd.read_csv(file_path)
        
        # 日付列をインデックスに設定
        if 'Date' in df.columns:
            try:
                # まずUTCとして日付を解析
                df['Date'] = pd.to_datetime(df['Date'], format='mixed', utc=True)
                # UTCからローカル時間に変換し、時間情報を削除
                df['Date'] = df['Date'].dt.tz_convert(None).dt.normalize()
                df.set_index('Date', inplace=True)
            except Exception as e:
                print(f"日付変換エラー（別の方法を試みます）: {str(e)}")
                try:
                    # バックアップとして、単純な日付変換を試みる
                    df['Date'] = pd.to_datetime(df['Date'], format='mixed')
                    df.set_index('Date', inplace=True)
                except Exception as e:
                    print(f"日付変換の2次試行も失敗: {str(e)}")
                    return None
        
        # インデックスがDatetimeIndexであることを確認
        if not isinstance(df.index, pd.DatetimeIndex):
            print(f"警告: {file_path} のインデックスがDatetimeIndexではありません。スキップします。")
            return None
        
        # データが空でないことを確認
        if df.empty:
            print(f"警告: {file_path} のデータが空です。スキップします。")
            return None
            
        # 必須カラムがあることを確認
        required_columns = ['Open', 'High', 'Low', 'Close', 'Volume']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"警告: {file_path} に必須カラム {missing_columns} がありません。スキップします。")
            return None
        
        # 不要なカラムを削除（Dividends, Stock Splitなど）
        df = df[required_columns]
        
        # 連続した日付インデックスを作成（土日を除く）
        start_date_dt = pd.to_datetime(start_date)
        end_date_dt = pd.to_datetime(end_date)
        all_dates = pd.date_range(start=start_date_dt, end=end_date_dt, freq='B')  # 'B'は営業日を表す
        
        # 元のデータを連続した日付でリインデックス
        df_reindexed = df.reindex(all_dates)
        
        # 欠損値を前の値で埋める（前方補完）
        df_filled = df_reindexed.fillna(method='ffill')
        
        # 最初の日のデータが欠損している場合は、後ろの値で埋める（後方補完）
        df = df_filled.fillna(method='bfill')
        
        # まずデータを期間でフィルタリング
        df = df[(df.index >= start_date) & (df.index <= end_date)]
        
        if df.empty:
            print(f"警告: {file_path} の指定期間内にデータがありません。スキップします。")
            return None
        
        # 既に日足データなので、そのまま使用
        daily_data = df.copy()
        
        # 開始日の値を取得
        if daily_data.empty:
            print(f"警告: {file_path} の日足データが空です。スキップします。")
            return None
        
        # テクニカル指標を追加
        daily_data = add_technical_indicators(daily_data)
            
        first_value = daily_data.iloc[0]['Close']
        
        # 変動率の計算（開始日を1.0として正規化）
        normalized_data = daily_data.copy()
        normalized_data['Normalized'] = daily_data['Close'] / first_value
        
        # 日次変化率も計算
        normalized_data['Daily_Change'] = daily_data['Close'].pct_change() * 100
        
        # 休業日情報を追加
        normalized_data['Is_Market_Holiday'] = normalized_data.index.isin(df.index)
        
        print(f"処理完了: {file_path} - 日足データ数: {len(normalized_data)}")
        return normalized_data
        
    except Exception as e:
        print(f"エラー発生: {file_path} の処理中にエラーが発生しました: {str(e)}")
        import traceback
        print(traceback.format_exc())  # 詳細なエラー情報を表示
        return None

# 全銘柄のデータを格納するリスト
all_data = {}
all_normalized_data = []
stock_names = []

# 各ファイルを処理
print(f"日足ファイル総数: {len(files)}")
for file in files:
    file_path = os.path.join(directory, file)
    stock_name = file.split('.')[0]  # ファイル名から銘柄名を取得
    
    print(f"処理中: {file}")
    normalized_data = process_daily_data(file_path, start_date, end_date)
    
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
    
    # タイムゾーン情報を削除
    combined_data.index = combined_data.index.tz_localize(None)
    
    # 休場日を含む連続した日付インデックスを作成
    start_date_dt = pd.to_datetime(start_date)
    end_date_dt = pd.to_datetime(end_date)
    full_date_range = pd.date_range(start=start_date_dt, end=end_date_dt, freq='D')
    
    # 土日を除外（平日のみ）
    weekdays_only = full_date_range[full_date_range.weekday < 5]
    
    print(f"連続した平日数: {len(weekdays_only)}")
    print(f"実際のデータ日数: {len(combined_data)}")
    print(f"欠損日数: {len(weekdays_only) - len(combined_data)}")
    
    # 連続した日付インデックスでリインデックス（前方埋め）
    combined_data_filled = combined_data.reindex(weekdays_only, method='ffill')
    
    # 最初の日のデータが欠損している場合は後方埋め
    combined_data_filled = combined_data_filled.fillna(method='bfill')
    
    print(f"休場日補完後のデータ形状: {combined_data_filled.shape}")
    print(f"休場日補完後のインデックス: {combined_data_filled.index[:5]} ... {combined_data_filled.index[-5:]}")
    
    # 元のcombined_dataを更新
    combined_data = combined_data_filled
    
    # CSV形式での保存（休場日補完済み）
    csv_path = os.path.join(result_dir, f'日足株価変動率比較_{period_str}.csv')
    combined_data.to_csv(csv_path)
    
    # 休場日補完の詳細情報をCSVと同じディレクトリに保存
    補完情報_path = os.path.join(result_dir, f'休場日補完情報_{period_str}.txt')
    with open(補完情報_path, 'w', encoding='utf-8') as f:
        f.write(f"休場日補完情報\n")
        f.write(f"==================\n")
        f.write(f"分析期間: {start_date} ～ {end_date}\n")
        f.write(f"対象平日数: {len(weekdays_only)}\n")
        f.write(f"実際のデータ日数: {len(all_normalized_data[0]) if all_normalized_data else 0}\n")
        f.write(f"補完された日数: {len(weekdays_only) - (len(all_normalized_data[0]) if all_normalized_data else 0)}\n")
        f.write(f"補完方法: 前日の値で埋める（前方埋め）\n")
        f.write(f"注意: 土日は除外されています\n")
    
    print(f"休場日補完情報を保存しました: {補完情報_path}")
    
    # 1. メイングラフ: 日足変動率比較
    fig = go.Figure()
    
    # 最終日のパフォーマンス順に銘柄を並び替え
    final_performance = combined_data.iloc[-1].sort_values(ascending=False)
    sorted_columns = final_performance.index.tolist()
    
    print("最終日パフォーマンス順（上位から）:")
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
                mode='lines',
                name=column,
                customdata=list(zip(formatted_dates, close_prices_str)),
                hovertemplate=
                '<b>' + column + '</b><br>' +
                '日付: %{customdata[0]}<br>' +
                '変動率: %{y:.2%}<br>' +
                '終値: %{customdata[1]}<br>' +
                '<extra></extra>'
            )
        )
    
    # グラフのレイアウト設定
    fig.update_layout(
        title=f'日足変動率比較（{start_date}～{end_date}、開始日基準：100%）',
        xaxis_title='日付',
        yaxis=dict(
            type='log',  # 対数スケールを使用
            title='変動率（開始日=100%、対数スケール）',
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
    main_graph_path = os.path.join(result_dir, f'日足株価変動率比較_{period_str}.html')
    fig.write_html(main_graph_path)
    fig.write_image(os.path.join(result_dir, f'日足株価変動率比較_{period_str}.png'))
    
    # 2. 各銘柄の詳細テクニカル分析グラフ（日足用）
    for stock in stock_names:
        # 銘柄ごとの詳細グラフを作成
        stock_data = all_data[stock]
        
        # 日付を数字形式でフォーマット
        formatted_dates = [date.strftime('%Y/%m/%d') for date in stock_data.index]
        
        # サブプロット（株価、RSI、MACD、ストキャスティクス、出来高）
        technical_fig = sp.make_subplots(
            rows=5, cols=1,
            shared_xaxes=True,
            vertical_spacing=0.05,
            subplot_titles=(
                f'{stock} - 株価とテクニカル指標', 
                'RSI (相対力指数)', 
                'MACD (移動平均収束拡散法)',
                'ストキャスティクス',
                '出来高'
            ),
            row_heights=[0.4, 0.15, 0.15, 0.15, 0.15]
        )
        
        # ローソク足チャート（上段）
        technical_fig.add_trace(
            go.Candlestick(
                x=stock_data.index,
                open=stock_data['Open'],
                high=stock_data['High'],
                low=stock_data['Low'],
                close=stock_data['Close'],
                name='ローソク足',
                increasing_line_color='blue',
                decreasing_line_color='red',
                increasing_fillcolor='lightblue',
                decreasing_fillcolor='lightcoral',
                xhoverformat='%Y/%m/%d'
            ),
            row=1, col=1
        )
        
        # 移動平均線
        ma_configs = [
            ('MA5', '短期移動平均線 (5日)', 'orange', 1.0),
            ('MA25', '中期移動平均線 (25日)', 'green', 1.5),
            ('MA75', '長期移動平均線 (75日)', 'red', 1.5),
            ('MA200', '超長期移動平均線 (200日)', 'purple', 2.0)
        ]
        
        for ma_col, ma_name, color, width in ma_configs:
            if ma_col in stock_data.columns and not stock_data[ma_col].isna().all():
                technical_fig.add_trace(
                    go.Scatter(
                        x=stock_data.index,
                        y=stock_data[ma_col],
                        mode='lines',
                        name=ma_name,
                        line=dict(color=color, width=width),
                        customdata=formatted_dates,
                        hovertemplate=f'<b>{ma_name}</b><br>' +
                                     '日付: %{customdata}<br>' +
                                     '値: %{y}<extra></extra>'
                    ),
                    row=1, col=1
                )
        
        # ボリンジャーバンド
        bb_configs = [
            ('BB_中心線', 'BB 中心線', 'gray', 'dot'),
            ('BB_上側1σ', '+1σ', 'rgba(0, 176, 246, 0.7)', 'dot'),
            ('BB_上側2σ', '+2σ', 'rgba(0, 176, 246, 0.8)', 'dot'),
            ('BB_下側1σ', '-1σ', 'rgba(255, 100, 100, 0.7)', 'dot'),
            ('BB_下側2σ', '-2σ', 'rgba(255, 100, 100, 0.8)', 'dot')
        ]
        
        for bb_col, bb_name, color, dash in bb_configs:
            if bb_col in stock_data.columns and not stock_data[bb_col].isna().all():
                technical_fig.add_trace(
                    go.Scatter(
                        x=stock_data.index,
                        y=stock_data[bb_col],
                        mode='lines',
                        name=bb_name,
                        line=dict(color=color, width=1, dash=dash),
                        customdata=formatted_dates,
                        hovertemplate=f'<b>{bb_name}</b><br>' +
                                     '日付: %{customdata}<br>' +
                                     '値: %{y}<extra></extra>'
                    ),
                    row=1, col=1
                )
        
        # RSIチャート（2段目）
        if 'RSI' in stock_data.columns:
            technical_fig.add_trace(
                go.Scatter(
                    x=stock_data.index,
                    y=stock_data['RSI'],
                    mode='lines',
                    name='RSI',
                    line=dict(color='purple'),
                    customdata=formatted_dates,
                    hovertemplate='<b>RSI</b><br>' +
                                 '日付: %{customdata}<br>' +
                                 'RSI: %{y:.2f}<extra></extra>'
                ),
                row=2, col=1
            )
            
            # RSIの上限と下限を示す水平線
            technical_fig.add_hline(y=70, line_dash="dash", line_color="red", row=2, col=1)
            technical_fig.add_hline(y=30, line_dash="dash", line_color="green", row=2, col=1)
        
        # MACDチャート（3段目）
        if 'MACD' in stock_data.columns and 'MACD_Signal' in stock_data.columns:
            technical_fig.add_trace(
                go.Scatter(
                    x=stock_data.index,
                    y=stock_data['MACD'],
                    mode='lines',
                    name='MACD',
                    line=dict(color='blue'),
                    customdata=formatted_dates,
                    hovertemplate='<b>MACD</b><br>' +
                                 '日付: %{customdata}<br>' +
                                 'MACD: %{y:.4f}<extra></extra>'
                ),
                row=3, col=1
            )
            
            technical_fig.add_trace(
                go.Scatter(
                    x=stock_data.index,
                    y=stock_data['MACD_Signal'],
                    mode='lines',
                    name='MACD Signal',
                    line=dict(color='red'),
                    customdata=formatted_dates,
                    hovertemplate='<b>MACD Signal</b><br>' +
                                 '日付: %{customdata}<br>' +
                                 'Signal: %{y:.4f}<extra></extra>'
                ),
                row=3, col=1
            )
            
            # MACDヒストグラム
            if 'MACD_Histogram' in stock_data.columns:
                colors = ['green' if x >= 0 else 'red' for x in stock_data['MACD_Histogram']]
                technical_fig.add_trace(
                    go.Bar(
                        x=stock_data.index,
                        y=stock_data['MACD_Histogram'],
                        name='MACD Histogram',
                        marker_color=colors,
                        customdata=formatted_dates,
                        hovertemplate='<b>MACD Histogram</b><br>' +
                                     '日付: %{customdata}<br>' +
                                     'Histogram: %{y:.4f}<extra></extra>'
                    ),
                    row=3, col=1
                )
        
        # ストキャスティクス（4段目）
        if '%K' in stock_data.columns and '%D' in stock_data.columns:
            technical_fig.add_trace(
                go.Scatter(
                    x=stock_data.index,
                    y=stock_data['%K'],
                    mode='lines',
                    name='%K',
                    line=dict(color='blue'),
                    customdata=formatted_dates,
                    hovertemplate='<b>%K</b><br>' +
                                 '日付: %{customdata}<br>' +
                                 '%K: %{y:.2f}<extra></extra>'
                ),
                row=4, col=1
            )
            
            technical_fig.add_trace(
                go.Scatter(
                    x=stock_data.index,
                    y=stock_data['%D'],
                    mode='lines',
                    name='%D',
                    line=dict(color='red'),
                    customdata=formatted_dates,
                    hovertemplate='<b>%D</b><br>' +
                                 '日付: %{customdata}<br>' +
                                 '%D: %{y:.2f}<extra></extra>'
                ),
                row=4, col=1
            )
            
            # ストキャスティクスの上限と下限を示す水平線
            technical_fig.add_hline(y=80, line_dash="dash", line_color="red", row=4, col=1)
            technical_fig.add_hline(y=20, line_dash="dash", line_color="green", row=4, col=1)
        
        # 出来高チャート（5段目）
        if 'Volume' in stock_data.columns:
            technical_fig.add_trace(
                go.Bar(
                    x=stock_data.index,
                    y=stock_data['Volume'],
                    name='出来高',
                    marker_color='lightblue',
                    customdata=formatted_dates,
                    hovertemplate='<b>出来高</b><br>' +
                                 '日付: %{customdata}<br>' +
                                 '出来高: %{y:,}<extra></extra>'
                ),
                row=5, col=1
            )
            
            # 出来高移動平均線
            if 'Volume_MA' in stock_data.columns and not stock_data['Volume_MA'].isna().all():
                technical_fig.add_trace(
                    go.Scatter(
                        x=stock_data.index,
                        y=stock_data['Volume_MA'],
                        mode='lines',
                        name='出来高移動平均線',
                        line=dict(color='red', width=2),
                        customdata=formatted_dates,
                        hovertemplate='<b>出来高移動平均線</b><br>' +
                                     '日付: %{customdata}<br>' +
                                     '平均出来高: %{y:,}<extra></extra>'
                    ),
                    row=5, col=1
                )
        
        # レイアウト設定
        technical_fig.update_layout(
            title=f'{stock} - テクニカル分析（{start_date}～{end_date}）',
            xaxis_title='日付',
            yaxis_title='株価',
            yaxis2_title='RSI',
            yaxis3_title='MACD',
            yaxis4_title='ストキャスティクス',
            yaxis5_title='出来高',
            hovermode='x unified',
            template='plotly_white',
            height=1000,
            width=1200,  # 幅を拡張
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            ),
            # ホバーツールの改善
            hoverlabel=dict(
                bgcolor="rgba(255,255,255,0.95)",
                bordercolor="black",
                font_size=12,
                font_family="Arial",
                align="left",
                namelength=-1
            ),
            margin=dict(l=80, r=150, t=100, b=80)
        )
        
        # X軸の日付表示を数字形式に変更
        technical_fig.update_xaxes(tickformat="%Y/%m/%d")
        
        # y軸の範囲設定
        technical_fig.update_yaxes(range=[0, 100], row=2, col=1)  # RSI
        technical_fig.update_yaxes(range=[0, 100], row=4, col=1)  # ストキャスティクス
        
        # HTMLとして保存
        technical_path = os.path.join(result_dir, f'{stock}_テクニカル分析_{period_str}.html')
        technical_fig.write_html(technical_path)
    
    # 3. パフォーマンス比較表
    performance_data = {}
    for column in combined_data.columns:
        initial_value = combined_data[column].iloc[0]
        final_value = combined_data[column].iloc[-1]
        max_value = combined_data[column].max()
        min_value = combined_data[column].min()
        
        overall_change = ((final_value / initial_value) - 1) * 100
        max_gain = ((max_value / initial_value) - 1) * 100
        max_drawdown = ((min_value / max_value) - 1) * 100
        
        # 日次リターンの統計
        if 'Daily_Change' in all_data[column].columns:
            daily_returns = all_data[column]['Daily_Change'].dropna()
            max_daily_return = daily_returns.max()
            min_daily_return = daily_returns.min()
            volatility = daily_returns.std()
            # シャープレシオ（簡易版、リスクフリーレートを0と仮定）
            avg_daily_return = daily_returns.mean()
            sharpe_ratio = avg_daily_return / volatility if volatility != 0 else 0
        else:
            max_daily_return = np.nan
            min_daily_return = np.nan
            volatility = np.nan
            sharpe_ratio = np.nan
            
        performance_data[column] = {
            'overall_change': overall_change,
            'max_gain': max_gain,
            'max_drawdown': max_drawdown,
            'max_daily_return': max_daily_return,
            'min_daily_return': min_daily_return,
            'volatility': volatility,
            'sharpe_ratio': sharpe_ratio
        }
    
    # パフォーマンスデータをテーブルとして表示
    table_data = [
        {'銘柄': stock, 
         '全期間変動率(%)': f"{performance_data[stock]['overall_change']:.2f}", 
         '最大上昇率(%)': f"{performance_data[stock]['max_gain']:.2f}", 
         '最大下落率(%)': f"{performance_data[stock]['max_drawdown']:.2f}",
         '最大日次リターン(%)': f"{performance_data[stock]['max_daily_return']:.2f}",
         '最小日次リターン(%)': f"{performance_data[stock]['min_daily_return']:.2f}",
         'ボラティリティ': f"{performance_data[stock]['volatility']:.2f}",
         'シャープレシオ': f"{performance_data[stock]['sharpe_ratio']:.3f}"}
        for stock in stock_names
    ]
    
    # 全期間変動率でソート
    table_data = sorted(table_data, key=lambda x: float(x['全期間変動率(%)'].replace(',', '.')), reverse=True)
    
    table_fig = go.Figure(data=[go.Table(
        header=dict(
            values=['銘柄', '全期間変動率(%)', '最大上昇率(%)', '最大下落率(%)', '最大日次リターン(%)', '最小日次リターン(%)', 'ボラティリティ', 'シャープレシオ'],
            fill_color='paleturquoise',
            align='left',
            font=dict(size=12)
        ),
        cells=dict(
            values=[
                [data['銘柄'] for data in table_data],
                [data['全期間変動率(%)'] for data in table_data],
                [data['最大上昇率(%)'] for data in table_data],
                [data['最大下落率(%)'] for data in table_data],
                [data['最大日次リターン(%)'] for data in table_data],
                [data['最小日次リターン(%)'] for data in table_data],
                [data['ボラティリティ'] for data in table_data],
                [data['シャープレシオ'] for data in table_data]
            ],
            fill_color=[['white', 'lightgrey'] * (len(table_data) // 2 + len(table_data) % 2)],
            align='left',
            font=dict(size=11)
        )
    )])
    
    table_fig.update_layout(
        title='銘柄パフォーマンス比較（全期間変動率の降順）',
        height=400 + 25 * len(table_data),
        width=1400  # 幅を拡張（カラムが増えたため）
    )
    
    # HTMLとして保存
    table_path = os.path.join(result_dir, f'銘柄パフォーマンス比較_{period_str}.html')
    table_fig.write_html(table_path)
    
    # 4. ボラティリティ分析グラフ
    volatility_fig = go.Figure()
    
    for column in sorted_columns:
        if 'Daily_Change' in all_data[column].columns:
            daily_returns = all_data[column]['Daily_Change'].dropna()
            # 20日間の移動ボラティリティを計算
            rolling_volatility = daily_returns.rolling(window=20).std()
            
            volatility_fig.add_trace(
                go.Scatter(
                    x=rolling_volatility.index,
                    y=rolling_volatility,
                    mode='lines',
                    name=column,
                    hovertemplate='<b>' + column + '</b><br>' +
                                 '日付: %{x}<br>' +
                                 'ボラティリティ: %{y:.2f}%<br>' +
                                 '<extra></extra>'
                )
            )
    
    volatility_fig.update_layout(
        title=f'20日移動ボラティリティ比較（{start_date}～{end_date}）',
        xaxis_title='日付',
        yaxis_title='ボラティリティ (%)',
        hovermode='x unified',
        template='plotly_white',
        legend_title='銘柄',
        height=600,
        width=1200,
        hoverlabel=dict(
            bgcolor="rgba(255,255,255,0.95)",
            bordercolor="black",
            font_size=12,
            font_family="Arial",
            align="left",
            namelength=-1
        ),
        margin=dict(l=80, r=150, t=100, b=80)
    )
    
    volatility_fig.update_xaxes(tickformat="%Y/%m/%d")
    
    # HTMLとして保存
    volatility_path = os.path.join(result_dir, f'ボラティリティ比較_{period_str}.html')
    volatility_fig.write_html(volatility_path)
    
    # 5. 相関分析ヒートマップ
    # 日次リターンの相関を計算
    returns_data = {}
    for stock in stock_names:
        if 'Daily_Change' in all_data[stock].columns:
            returns_data[stock] = all_data[stock]['Daily_Change'].dropna()
    
    if returns_data:
        returns_df = pd.DataFrame(returns_data)
        correlation_matrix = returns_df.corr()
        
        correlation_fig = go.Figure(data=go.Heatmap(
            z=correlation_matrix.values,
            x=correlation_matrix.columns,
            y=correlation_matrix.index,
            colorscale='RdBu',
            zmid=0,
            text=correlation_matrix.round(3).values,
            texttemplate="%{text}",
            textfont={"size": 10},
            hovertemplate='銘柄1: %{x}<br>銘柄2: %{y}<br>相関係数: %{z:.3f}<extra></extra>'
        ))
        
        correlation_fig.update_layout(
            title='銘柄間の相関分析（日次リターン）',
            xaxis_title='銘柄',
            yaxis_title='銘柄',
            height=600 + 20 * len(stock_names),
            width=600 + 20 * len(stock_names),
            template='plotly_white'
        )
        
        # HTMLとして保存
        correlation_path = os.path.join(result_dir, f'相関分析_{period_str}.html')
        correlation_fig.write_html(correlation_path)
    
    # 6. 総合ダッシュボード作成
    dashboard_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>日足株価分析ダッシュボード</title>
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background-color: #f5f5f5; }}
            .dashboard {{ display: flex; flex-wrap: wrap; gap: 20px; justify-content: space-between; }}
            .chart-container {{ background: white; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 15px; margin-bottom: 20px; }}
            .full-width {{ width: 100%; }}
            .half-width {{ width: calc(50% - 10px); }}
            h1, h2 {{ color: #333; }}
            h1 {{ text-align: center; margin-bottom: 25px; border-bottom: 2px solid #ddd; padding-bottom: 15px; }}
            h2 {{ margin-top: 0; padding-bottom: 10px; border-bottom: 1px solid #eee; }}
            iframe {{ border: none; width: 100%; height: 700px; border-radius: 4px; }}
            .stock-list {{ display: flex; flex-wrap: wrap; gap: 10px; margin-top: 15px; }}
            .stock-link {{ padding: 8px 16px; background-color: #e9f5ff; border-radius: 6px; text-decoration: none; color: #0066cc; transition: all 0.2s; }}
            .stock-link:hover {{ background-color: #cce5ff; }}
            .date-info {{ text-align: center; font-size: 1.2em; margin-bottom: 25px; color: #666; }}
            .table-iframe {{ height: 500px; }}
            .analysis-iframe {{ height: 600px; }}
        </style>
    </head>
    <body>
        <h1>日足株価分析ダッシュボード</h1>
        <div class="date-info">分析期間: {start_date} ～ {end_date}</div>
        
        <div class="dashboard">
            <div class="chart-container full-width">
                <h2>日足変動率比較</h2>
                <iframe src="{os.path.basename(main_graph_path)}"></iframe>
            </div>
            
            <div class="chart-container half-width">
                <h2>ボラティリティ分析</h2>
                <iframe src="{os.path.basename(volatility_path)}" class="analysis-iframe"></iframe>
            </div>
    """
    
    # 相関分析がある場合のみ追加
    if returns_data:
        dashboard_html += f"""
            <div class="chart-container half-width">
                <h2>相関分析</h2>
                <iframe src="{os.path.basename(correlation_path)}" class="analysis-iframe"></iframe>
            </div>
        """
    
    dashboard_html += f"""
            <div class="chart-container full-width">
                <h2>銘柄パフォーマンス比較</h2>
                <iframe src="{os.path.basename(table_path)}" class="table-iframe"></iframe>
            </div>
            
            <div class="chart-container full-width">
                <h2>各銘柄のテクニカル分析</h2>
                <p>以下の銘柄をクリックすると、詳細なテクニカル分析グラフが表示されます：</p>
                <div class="stock-list">
    """
    
    # 各銘柄へのリンクを追加
    for stock in stock_names:
        technical_filename = f'{stock}_テクニカル分析_{period_str}.html'
        dashboard_html += f'<a href="{technical_filename}" class="stock-link" target="_blank">{stock}</a>\n'
    
    dashboard_html += """
                </div>
            </div>
        </div>
        
        <div class="chart-container full-width" style="margin-top: 30px;">
            <h2>分析概要</h2>
            <p><strong>含まれる分析内容：</strong></p>
            <ul>
                <li><strong>変動率比較：</strong> 各銘柄の開始日を基準とした変動率をインタラクティブに比較</li>
                <li><strong>テクニカル分析：</strong> ローソク足、移動平均線、ボリンジャーバンド、RSI、MACD、ストキャスティクス、出来高分析</li>
                <li><strong>ボラティリティ分析：</strong> 20日移動ボラティリティによるリスク評価</li>
                <li><strong>相関分析：</strong> 銘柄間の相関関係をヒートマップで可視化</li>
                <li><strong>パフォーマンス指標：</strong> シャープレシオ、最大ドローダウン等の総合評価</li>
            </ul>
            <p><strong>使用したテクニカル指標：</strong></p>
            <ul>
                <li>移動平均線：5日、25日、75日、200日</li>
                <li>ボリンジャーバンド：20日移動平均±1σ、±2σ</li>
                <li>RSI：14日間の相対力指数</li>
                <li>MACD：12日EMA-26日EMA、9日シグナル線</li>
                <li>ストキャスティクス：14日%K、3日%D</li>
                <li>出来高：実際の出来高と20日移動平均</li>
            </ul>
        </div>
    </body>
    </html>
    """
    
    # ダッシュボードHTMLの保存
    dashboard_path = os.path.join(result_dir, f'日足株価分析ダッシュボード_{period_str}.html')
    with open(dashboard_path, 'w', encoding='utf-8') as f:
        f.write(dashboard_html)
    
    print(f"分析完了。以下のファイルが作成されました：")
    print(f"1. メイングラフ: {main_graph_path}")
    print(f"2. ボラティリティ分析: {volatility_path}")
    if returns_data:
        print(f"3. 相関分析: {correlation_path}")
    print(f"4. 銘柄パフォーマンス比較: {table_path}")
    print(f"5. 各銘柄のテクニカル分析: {result_dir}/*_テクニカル分析_*.html")
    print(f"6. 総合ダッシュボード: {dashboard_path}")
    print(f"7. CSV: {csv_path}")
    
    print(f"\nダッシュボードを開くには、以下のファイルをブラウザで開いてください：")
    print(f"{dashboard_path}")

    # テクニカル分析レポートの生成
    technical_report_path = os.path.join(result_dir, f'テクニカル分析レポート_{period_str}.txt')
    with open(technical_report_path, 'w', encoding='utf-8') as f:
        f.write(f"テクニカル分析レポート（{start_date}～{end_date}）\n")
        f.write("=" * 80 + "\n\n")
        
        for stock in stock_names:
            stock_data = all_data[stock]
            f.write(f"\n{stock}の分析\n")
            f.write("-" * 40 + "\n")
            
            # 価格トレンド分析
            latest_close = stock_data['Close'].iloc[-1]
            ma5_latest = stock_data['MA5'].iloc[-1]
            ma25_latest = stock_data['MA25'].iloc[-1]
            ma75_latest = stock_data['MA75'].iloc[-1]
            ma200_latest = stock_data['MA200'].iloc[-1]
            
            f.write("\n【価格トレンド分析】\n")
            f.write(f"現在値: {latest_close:.2f}\n")
            f.write(f"5日移動平均線: {ma5_latest:.2f}\n")
            f.write(f"25日移動平均線: {ma25_latest:.2f}\n")
            f.write(f"75日移動平均線: {ma75_latest:.2f}\n")
            f.write(f"200日移動平均線: {ma200_latest:.2f}\n")
            
            # トレンド判定
            trend_text = []
            if latest_close > ma5_latest > ma25_latest:
                trend_text.append("短期的に強い上昇トレンド")
            elif latest_close < ma5_latest < ma25_latest:
                trend_text.append("短期的に強い下降トレンド")
            
            if ma5_latest > ma75_latest and ma25_latest > ma75_latest:
                trend_text.append("中期的に上昇トレンド")
            elif ma5_latest < ma75_latest and ma25_latest < ma75_latest:
                trend_text.append("中期的に下降トレンド")
            
            if all(x > ma200_latest for x in [latest_close, ma5_latest, ma25_latest, ma75_latest]):
                trend_text.append("長期的に強い上昇トレンド")
            elif all(x < ma200_latest for x in [latest_close, ma5_latest, ma25_latest, ma75_latest]):
                trend_text.append("長期的に強い下降トレンド")
            
            if trend_text:
                f.write("\nトレンド判定：\n")
                for trend in trend_text:
                    f.write(f"- {trend}\n")
            
            # RSI分析
            latest_rsi = stock_data['RSI'].iloc[-1]
            f.write("\n【RSI分析】\n")
            f.write(f"現在のRSI値: {latest_rsi:.2f}\n")
            if latest_rsi >= 70:
                f.write("※ 売られ過ぎの可能性が高い（70以上）\n")
            elif latest_rsi <= 30:
                f.write("※ 買われ過ぎの可能性が高い（30以下）\n")
            
            # MACD分析
            latest_macd = stock_data['MACD'].iloc[-1]
            latest_signal = stock_data['MACD_Signal'].iloc[-1]
            latest_hist = stock_data['MACD_Histogram'].iloc[-1]
            
            f.write("\n【MACD分析】\n")
            f.write(f"MACD: {latest_macd:.4f}\n")
            f.write(f"シグナル: {latest_signal:.4f}\n")
            f.write(f"ヒストグラム: {latest_hist:.4f}\n")
            
            # MACDのクロス判定
            if latest_hist > 0 and stock_data['MACD_Histogram'].iloc[-2] <= 0:
                f.write("※ 直近でゴールデンクロス（買いシグナル）\n")
            elif latest_hist < 0 and stock_data['MACD_Histogram'].iloc[-2] >= 0:
                f.write("※ 直近でデッドクロス（売りシグナル）\n")
            
            # ボリンジャーバンド分析
            latest_price = stock_data['Close'].iloc[-1]
            latest_bb_upper2 = stock_data['BB_上側2σ'].iloc[-1]
            latest_bb_lower2 = stock_data['BB_下側2σ'].iloc[-1]
            latest_bb_mid = stock_data['BB_中心線'].iloc[-1]
            
            f.write("\n【ボリンジャーバンド分析】\n")
            f.write(f"現在値: {latest_price:.2f}\n")
            f.write(f"上側2σ: {latest_bb_upper2:.2f}\n")
            f.write(f"中心線: {latest_bb_mid:.2f}\n")
            f.write(f"下側2σ: {latest_bb_lower2:.2f}\n")
            
            if latest_price > latest_bb_upper2:
                f.write("※ +2σを上回っており、売られ過ぎの可能性\n")
            elif latest_price < latest_bb_lower2:
                f.write("※ -2σを下回っており、買われ過ぎの可能性\n")
            
            # ストキャスティクス分析
            latest_k = stock_data['%K'].iloc[-1]
            latest_d = stock_data['%D'].iloc[-1]
            
            f.write("\n【ストキャスティクス分析】\n")
            f.write(f"%K: {latest_k:.2f}\n")
            f.write(f"%D: {latest_d:.2f}\n")
            
            if latest_k > 80 and latest_d > 80:
                f.write("※ 売られ過ぎの水準（80以上）\n")
            elif latest_k < 20 and latest_d < 20:
                f.write("※ 買われ過ぎの水準（20以下）\n")
            
            # 出来高分析
            latest_volume = stock_data['Volume'].iloc[-1]
            avg_volume = stock_data['Volume_MA'].iloc[-1]
            volume_ratio = latest_volume / avg_volume if avg_volume != 0 else 0
            
            f.write("\n【出来高分析】\n")
            f.write(f"直近出来高: {int(latest_volume):,}\n")
            f.write(f"20日平均出来高: {int(avg_volume):,}\n")
            f.write(f"出来高乖離率: {(volume_ratio - 1) * 100:.2f}%\n")
            
            if volume_ratio > 2:
                f.write("※ 出来高が平均の2倍を超えており、大きな動きの可能性\n")
            
            # 総合判断
            f.write("\n【総合判断】\n")
            signals = []
            
            # トレンド
            if latest_close > ma5_latest > ma25_latest > ma75_latest:
                signals.append("強い上昇トレンド")
            elif latest_close < ma5_latest < ma25_latest < ma75_latest:
                signals.append("強い下降トレンド")
            
            # RSI
            if latest_rsi > 70:
                signals.append("RSIが売られ過ぎ")
            elif latest_rsi < 30:
                signals.append("RSIが買われ過ぎ")
            
            # MACD
            if latest_hist > 0 and stock_data['MACD_Histogram'].iloc[-2] <= 0:
                signals.append("MACDがゴールデンクロス")
            elif latest_hist < 0 and stock_data['MACD_Histogram'].iloc[-2] >= 0:
                signals.append("MACDがデッドクロス")
            
            # ボリンジャーバンド
            if latest_price > latest_bb_upper2:
                signals.append("ボリンジャーバンド+2σを超過")
            elif latest_price < latest_bb_lower2:
                signals.append("ボリンジャーバンド-2σを超過")
            
            # ストキャスティクス
            if latest_k > 80 and latest_d > 80:
                signals.append("ストキャスティクスが売られ過ぎ")
            elif latest_k < 20 and latest_d < 20:
                signals.append("ストキャスティクスが買われ過ぎ")
            
            # 出来高
            if volume_ratio > 2:
                signals.append("出来高が急増")
            
            if signals:
                f.write("注目すべきシグナル：\n")
                for signal in signals:
                    f.write(f"- {signal}\n")
            else:
                f.write("現時点で特筆すべきシグナルはありません\n")
            
            f.write("\n" + "=" * 80 + "\n")
    
    print(f"8. テクニカル分析レポート: {technical_report_path}")

else:
    print("指定した期間内のデータが見つかりませんでした。")