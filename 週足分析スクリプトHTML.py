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
    if 'Volume' in df.columns:
        result['Volume_MA'] = df['Volume'].rolling(window=20).mean()
    else:
        result['Volume_MA'] = np.nan # 出来高データがない場合はNaN
    
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

# ファイル名用の期間文字列
period_str = f"{start_date.replace('-', '')}_{end_date.replace('-', '')}"

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
        required_columns = ['Open', 'High', 'Low', 'Close'] # Volumeはオプション扱い
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"警告: {file_path} に必須カラム {missing_columns} がありません。スキップします。")
            return None
        
        # Volumeカラムがない場合はNaNで埋める
        if 'Volume' not in df.columns:
            df['Volume'] = np.nan
            print(f"警告: {file_path} にVolumeカラムがありません。NaNで補完します。")

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
        weekly_data_filled = weekly_data_filled.bfill() # bfillに変更
        
        # 休場週補完後のデータを使用
        weekly_data = weekly_data_filled
        
        # 開始週の値を取得
        if weekly_data.empty:
            print(f"警告: {file_path} の週足データが空です。スキップします。")
            return None
        
        # テクニカル指標を追加
        weekly_data = add_technical_indicators(weekly_data) # This needs Volume
            
        first_value = weekly_data.iloc[0]['Close']
        if pd.isna(first_value) or first_value == 0: 
            print(f"警告: {file_path} の最初の終値が無効です。正規化をスキップします。")
            # Return weekly_data without 'Normalized' or handle differently
            weekly_data['Normalized'] = np.nan 
        else:
            # 変動率の計算（開始週を1.0として正規化）
            weekly_data['Normalized'] = weekly_data['Close'] / first_value
        
        # 週次変化率も計算
        weekly_data['Weekly_Change'] = weekly_data['Close'].pct_change() * 100
        
        print(f"処理完了: {file_path} - 週足データ数（休場週補完後）: {len(weekly_data)}")
        return weekly_data
        
    except Exception as e:
        print(f"エラー発生: {file_path} の処理中にエラーが発生しました: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def create_technical_analysis_dashboard(weekly_data, stock_name, output_path):
    """テクニカル分析ダッシュボードを作成する関数"""
    # サブプロットの設定
    fig = sp.make_subplots(
        rows=4, cols=1,
        shared_xaxes=True,
        vertical_spacing=0.05,
        subplot_titles=('株価とボリンジャーバンド', 'RSI', 'MACD', '出来高'),
        row_heights=[0.5, 0.2, 0.2, 0.1]
    )

    # メインチャート（株価とボリンジャーバンド）
    fig.add_trace(go.Candlestick(
        x=weekly_data.index,
        open=weekly_data['Open'],
        high=weekly_data['High'],
        low=weekly_data['Low'],
        close=weekly_data['Close'],
        name='株価'
    ), row=1, col=1)

    # ボリンジャーバンドの追加
    fig.add_trace(go.Scatter(x=weekly_data.index, y=weekly_data['BB_中心線'], 
                            name='BB中心線', line=dict(color='gray')), row=1, col=1)
    fig.add_trace(go.Scatter(x=weekly_data.index, y=weekly_data['BB_上側2σ'], 
                            name='BB+2σ', line=dict(color='lightgray')), row=1, col=1)
    fig.add_trace(go.Scatter(x=weekly_data.index, y=weekly_data['BB_下側2σ'], 
                            name='BB-2σ', line=dict(color='lightgray')), row=1, col=1)

    # RSIの追加
    fig.add_trace(go.Scatter(x=weekly_data.index, y=weekly_data['RSI'], 
                            name='RSI', line=dict(color='blue')), row=2, col=1)
    # RSIの基準線
    fig.add_hline(y=70, line_dash="dash", line_color="red", row=2, col=1)
    fig.add_hline(y=30, line_dash="dash", line_color="green", row=2, col=1)

    # MACDの追加
    fig.add_trace(go.Scatter(x=weekly_data.index, y=weekly_data['MACD'], 
                            name='MACD', line=dict(color='blue')), row=3, col=1)
    fig.add_trace(go.Scatter(x=weekly_data.index, y=weekly_data['MACD_Signal'], 
                            name='シグナル', line=dict(color='orange')), row=3, col=1)
    fig.add_bar(x=weekly_data.index, y=weekly_data['MACD_Histogram'], 
                name='ヒストグラム', marker_color='gray', row=3, col=1)

    # 出来高の追加
    if 'Volume' in weekly_data.columns and not weekly_data['Volume'].isna().all():
        fig.add_bar(x=weekly_data.index, y=weekly_data['Volume'], 
                    name='出来高', marker_color='lightblue', row=4, col=1)
        if 'Volume_MA' in weekly_data.columns and not weekly_data['Volume_MA'].isna().all():
            fig.add_trace(go.Scatter(x=weekly_data.index, y=weekly_data['Volume_MA'], 
                                    name='出来高MA', line=dict(color='navy')), row=4, col=1)
    else:
        # 出来高データがない場合のプレースホルダー
        fig.add_trace(go.Scatter(x=weekly_data.index, y=[0]*len(weekly_data), name='出来高 (データなし)', mode='lines', line=dict(color='rgba(0,0,0,0)')), row=4, col=1)


    # レイアウトの設定
    fig.update_layout(
        title=f'{stock_name} - テクニカル分析チャート',
        xaxis_title='日付',
        height=1000,
        showlegend=True
    )

    # HTMLファイルとして保存
    fig.write_html(output_path)

def create_performance_dashboard(combined_data, all_weekly_data, output_path):
    """パフォーマンス比較ダッシュボードを作成する関数"""
    fig = sp.make_subplots(
        rows=3, cols=1,
        shared_xaxes=True,
        vertical_spacing=0.05,
        subplot_titles=('株価変動率比較', '週次変化率', '出来高比較（相対）'),
        row_heights=[0.5, 0.25, 0.25]
    )

    # 最終日のパフォーマンス順に銘柄を並び替え (combined_dataの列順)
    # combined_dataは変動率(%)なので、そのままソートに使用
    if not combined_data.empty:
        final_performance = combined_data.iloc[-1].sort_values(ascending=False)
        sorted_columns = final_performance.index.tolist()
    else:
        sorted_columns = combined_data.columns.tolist()


    # 1. 株価変動率比較
    for stock_name_key in sorted_columns: # ソートされた順でトレースを追加
        display_name = stock_name_key.replace('_週足','')
        
        # 終値データを取得 (all_weekly_dataから)
        close_prices_str = ["N/A"] * len(combined_data.index)
        if stock_name_key in all_weekly_data and 'Close' in all_weekly_data[stock_name_key].columns:
            # all_weekly_dataのインデックスをcombined_dataのインデックスに合わせる
            # これは、休場補完などでインデックスが異なる可能性があるため
            close_prices_series = all_weekly_data[stock_name_key]['Close'].reindex(combined_data.index, method='ffill').bfill()
            close_prices_str = close_prices_series.apply(lambda x: f"{x:.2f}" if pd.notna(x) else "N/A").tolist()

        fig.add_trace(
            go.Scatter(
                x=combined_data.index,
                y=combined_data[stock_name_key],
                name=display_name,
                mode='lines',
                customdata=list(zip(combined_data.index.strftime('%Y-%m-%d'), close_prices_str)),
                hovertemplate=
                f'<b>{display_name}</b><br>' +
                '日付: %{customdata[0]}<br>' +
                '変動率: %{y:.2f}%<br>' +
                '終値: %{customdata[1]}<br>' +
                '<extra></extra>'
            ),
            row=1, col=1
        )

    # 2. 週次変化率
    for stock_name_key in sorted_columns: # ソートされた順でトレースを追加
        if stock_name_key in all_weekly_data and 'Weekly_Change' in all_weekly_data[stock_name_key].columns:
            weekly_data_df = all_weekly_data[stock_name_key]
            display_name = stock_name_key.replace('_週足','')
            fig.add_trace(
                go.Bar(
                    x=weekly_data_df.index,
                    y=weekly_data_df['Weekly_Change'],
                    name=f"{display_name} (週次変化率)", # Show legend for this subplot
                    legendgroup=display_name, # Group with main plot
                    showlegend=False, # Initially hide to avoid clutter, controlled by main legend
                    hovertemplate=
                    f'<b>{display_name}</b><br>' +
                    '日付: %{x|%Y-%m-%d}<br>' + # Format date in hover
                    '週次変化率: %{y:.2f}%<br>' +
                    '<extra></extra>'
                ),
                row=2, col=1
            )

    # 3. 出来高比較（正規化）
    for stock_name_key in sorted_columns: # ソートされた順でトレースを追加
        if stock_name_key in all_weekly_data and 'Volume' in all_weekly_data[stock_name_key].columns and not all_weekly_data[stock_name_key]['Volume'].isna().all():
            weekly_data_df = all_weekly_data[stock_name_key]
            display_name = stock_name_key.replace('_週足','')
            
            first_volume = weekly_data_df.iloc[0]['Volume']
            if pd.notna(first_volume) and first_volume != 0:
                normalized_volume = (weekly_data_df['Volume'] / first_volume) * 100
            else:
                normalized_volume = pd.Series(np.nan, index=weekly_data_df.index)
            
            fig.add_trace(
                go.Scatter(
                    x=weekly_data_df.index,
                    y=normalized_volume,
                    name=f"{display_name} (出来高)", # Show legend for this subplot
                    legendgroup=display_name, # Group with main plot
                    showlegend=False, # Initially hide
                    mode='lines',
                    hovertemplate=
                    f'<b>{display_name}</b><br>' +
                    '日付: %{x|%Y-%m-%d}<br>' + # Format date in hover
                    '相対出来高: %{y:.2f}%<br>' +
                    '<extra></extra>'
                ),
                row=3, col=1
            )

    # レイアウトの設定
    fig.update_layout(
        title={
            'text': f'株価パフォーマンス総合分析 ({start_date}～{end_date})',
            'y':0.98, # タイトル位置調整
            'x':0.5,
            'xanchor': 'center',
            'yanchor': 'top'
        },
        height=1200,
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.01, # 凡例位置調整
            xanchor="right",
            x=1
        ),
        hovermode='x unified', # Default hovermode
        template='plotly_white',
        updatemenus=[
            # 銘柄選択ドロップダウンメニュー
            {
                'buttons': [
                    {
                        'label': '全銘柄表示',
                        'method': 'update',
                        'args': [{'visible': [True] * len(sorted_columns) * 3}] # 3サブプロット分
                    }
                ] + [
                    {
                        'label': col.replace('_週足',''),
                        'method': 'update',
                        'args': [{'visible': [True if i == j else False for j in range(len(sorted_columns))] * 3}]
                    }
                    for i, col in enumerate(sorted_columns)
                ],
                'direction': 'down',
                'showactive': True,
                'x': 0.01, # 位置調整
                'y': 1.07, # 位置調整
                'xanchor': 'left',
                'yanchor': 'top'
            },
            # ホバーモード切り替えボタン
            {
                'buttons': [
                    {
                        'label': '個別ホバー',
                        'method': 'relayout',
                        'args': [{'hovermode': 'closest'}]
                    },
                    {
                        'label': '全銘柄ホバー',
                        'method': 'relayout', 
                        'args': [{'hovermode': 'x unified'}]
                    }
                ],
                'direction': 'down',
                'showactive': True,
                'x': 0.17, # 位置調整
                'y': 1.07, # 位置調整
                'xanchor': 'left',
                'yanchor': 'top'
            }
        ]
    )

    # Y軸のタイトルを設定
    fig.update_yaxes(title_text="変動率（%）", row=1, col=1)
    fig.update_yaxes(title_text="週次変化率（%）", row=2, col=1)
    fig.update_yaxes(title_text="相対出来高（%）", row=3, col=1)

    # X軸の設定
    fig.update_xaxes(title_text="日付", row=3, col=1, tickformat="%Y-%m-%d")
    fig.update_xaxes(row=1, col=1, tickformat="%Y-%m-%d")
    fig.update_xaxes(row=2, col=1, tickformat="%Y-%m-%d")


    # HTMLファイルとして保存
    fig.write_html(output_path)

# メイン処理部分
print(f"週足ファイル総数: {len(files)}")

# 各銘柄のデータを処理
all_data = {} # Stores normalized percentage change for combined_data
all_weekly_data = {} # Stores full processed weekly data for each stock (with indicators)

for file in files:
    file_path = os.path.join(directory, file)
    print(f"処理中: {file}")
    
    stock_name_from_file = file.replace('.csv', '') # e.g., AAPL_アップル_週足

    # 週足データを処理 (returns full df with indicators and 'Normalized' and 'Weekly_Change')
    processed_df = process_weekly_data(file_path, start_date, end_date)
    
    if processed_df is not None and not processed_df.empty:
        all_weekly_data[stock_name_from_file] = processed_df
        
        # テクニカル分析ダッシュボードの作成
        # Use stock_name_from_file for filename consistency
        technical_output = os.path.join(result_dir, f'テクニカル分析_{stock_name_from_file}_{period_str}.html')
        create_technical_analysis_dashboard(processed_df, stock_name_from_file.replace('_週足',''), technical_output) # Pass cleaned name for title
        
        # 変動率の計算
        # For combined_data, we use the percentage change from the start
        first_close_value = processed_df.iloc[0]['Close']
        if pd.notna(first_close_value) and first_close_value != 0:
            percentage_change = ((processed_df['Close'] - first_close_value) / first_close_value) * 100
        else:
            percentage_change = pd.Series(np.nan, index=processed_df.index) 
            
        all_data[stock_name_from_file] = percentage_change
        
        print(f"テクニカル指標計算済み: {stock_name_from_file}, 変動率データ数: {len(percentage_change)}")

# 全銘柄のデータを結合
print(f"処理完了した銘柄数: {len(all_data)}")
if all_data:
    combined_data = pd.DataFrame(all_data)
    if not isinstance(combined_data.index, pd.DatetimeIndex):
        try:
            combined_data.index = pd.to_datetime(combined_data.index)
        except Exception as e:
            print(f"Combined_dataインデックスの日付変換エラー: {e}")

    print(f"結合前のデータ形状: {combined_data.shape}")
    if not combined_data.empty:
        print(f"結合後のインデックス: {combined_data.index[:5]} ... {combined_data.index[-5:]}")
    
    # パフォーマンス比較表のHTMLファイルパス
    table_html_path = os.path.join(result_dir, f'週足銘柄パフォーマンス比較_{period_str}.html')
    # メイングラフ（変動率比較）のHTMLファイルパス
    main_graph_path = os.path.join(result_dir, f'週足株価変動率比較_{period_str}.html')

    # パフォーマンス比較グラフの作成（変動率比較グラフ）
    create_performance_dashboard(combined_data, all_weekly_data, main_graph_path)
    
    # CSVファイルとして保存 (変動率データ)
    csv_output = os.path.join(result_dir, f'週足株価変動率比較_{period_str}.csv')
    combined_data.to_csv(csv_output)
    
    # 銘柄パフォーマンス比較表の作成
    performance_metrics = {}
    for stock_key in combined_data.columns: 
        overall_change = combined_data[stock_key].iloc[-1] if not combined_data[stock_key].empty else np.nan
        max_gain = combined_data[stock_key].max() if not combined_data[stock_key].empty else np.nan
        min_loss = combined_data[stock_key].min() if not combined_data[stock_key].empty else np.nan
        
        weekly_returns = pd.Series(dtype=float) 
        if stock_key in all_weekly_data and 'Weekly_Change' in all_weekly_data[stock_key].columns:
            weekly_returns = all_weekly_data[stock_key]['Weekly_Change'].dropna()
        
        max_weekly_return = weekly_returns.max() if not weekly_returns.empty else np.nan
        min_weekly_return = weekly_returns.min() if not weekly_returns.empty else np.nan
        volatility = weekly_returns.std() if not weekly_returns.empty else np.nan
            
        performance_metrics[stock_key] = {
            'overall_change': overall_change,
            'max_gain': max_gain,
            'min_loss': min_loss,
            'max_weekly_return': max_weekly_return,
            'min_weekly_return': min_weekly_return,
            'volatility': volatility
        }

    table_data_for_html = [
        {'銘柄': stock.replace('_週足', ''), 
         '全期間変動率(%)': f"{performance_metrics[stock]['overall_change']:.2f}" if pd.notna(performance_metrics[stock]['overall_change']) else "N/A",
         '期間中最大上昇率(%)': f"{performance_metrics[stock]['max_gain']:.2f}" if pd.notna(performance_metrics[stock]['max_gain']) else "N/A",
         '期間中最大下落率(%)': f"{performance_metrics[stock]['min_loss']:.2f}" if pd.notna(performance_metrics[stock]['min_loss']) else "N/A",
         '最大週間リターン(%)': f"{performance_metrics[stock]['max_weekly_return']:.2f}" if pd.notna(performance_metrics[stock]['max_weekly_return']) else "N/A",
         '最小週間リターン(%)': f"{performance_metrics[stock]['min_weekly_return']:.2f}" if pd.notna(performance_metrics[stock]['min_weekly_return']) else "N/A",
         '週間ボラティリティ(%)': f"{performance_metrics[stock]['volatility']:.2f}" if pd.notna(performance_metrics[stock]['volatility']) else "N/A"
        }
        for stock in combined_data.columns
    ]
    
    table_data_for_html = sorted(
        table_data_for_html, 
        key=lambda x: float(x['全期間変動率(%)'].replace('N/A', '-inf')) if x['全期間変動率(%)'] != "N/A" else -float('inf'), 
        reverse=True
    )

    table_fig = go.Figure(data=[go.Table(
        header=dict(
            values=['銘柄', '全期間変動率(%)', '期間中最大上昇率(%)', '期間中最大下落率(%)', '最大週間リターン(%)', '最小週間リターン(%)', '週間ボラティリティ(%)'],
            fill_color='paleturquoise',
            align='left',
            font=dict(size=12)
        ),
        cells=dict(
            values=[
                [data['銘柄'] for data in table_data_for_html],
                [data['全期間変動率(%)'] for data in table_data_for_html],
                [data['期間中最大上昇率(%)'] for data in table_data_for_html],
                [data['期間中最大下落率(%)'] for data in table_data_for_html],
                [data['最大週間リターン(%)'] for data in table_data_for_html],
                [data['最小週間リターン(%)'] for data in table_data_for_html],
                [data['週間ボラティリティ(%)'] for data in table_data_for_html]
            ],
            fill_color=[['white', 'lightgrey'] * (len(table_data_for_html) // 2 + 1)][:len(table_data_for_html)], # Ensure correct length for alternating colors
            align='left',
            font=dict(size=11)
        )
    )])
    
    table_fig.update_layout(
        title='週足銘柄パフォーマンス比較（全期間変動率の降順）',
        height=max(400, 100 + 30 * len(table_data_for_html)), 
        width=1200
    )
    table_fig.write_html(table_html_path)

    # 最終週のパフォーマンスを表示
    if not combined_data.empty:
        print("\n最終週パフォーマンス順（上位から）:")
        final_performance = combined_data.iloc[-1].sort_values(ascending=False)
        for i, (stock, perf) in enumerate(final_performance.items(), 1):
            print(f" {i}. {stock.replace('_週足','')}:  {perf:+.2f}%")
    
    # 総合ダッシュボード作成
    dashboard_html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>週足株価分析ダッシュボード</title>
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background-color: #f5f5f5; }}
            .dashboard {{ display: flex; flex-wrap: wrap; gap: 20px; justify-content: space-between; }}
            .chart-container {{ background: white; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 15px; margin-bottom: 20px; }}
            .full-width {{ width: 100%; }}
            h1, h2 {{ color: #333; }}
            h1 {{ text-align: center; margin-bottom: 25px; border-bottom: 2px solid #ddd; padding-bottom: 15px; }}
            h2 {{ margin-top: 0; padding-bottom: 10px; border-bottom: 1px solid #eee; }}
            iframe {{ border: none; width: 100%; height: 700px; border-radius: 4px; }}
            .stock-list {{ display: flex; flex-wrap: wrap; gap: 10px; margin-top: 15px; }}
            .stock-link {{ padding: 8px 16px; background-color: #e9f5ff; border-radius: 6px; text-decoration: none; color: #0066cc; transition: all 0.2s; }}
            .stock-link:hover {{ background-color: #cce5ff; }}
            .date-info {{ text-align: center; font-size: 1.2em; margin-bottom: 25px; color: #666; }}
            .table-iframe {{ height: 500px; }} 
        </style>
    </head>
    <body>
        <h1>週足株価分析ダッシュボード</h1>
        <div class="date-info">分析期間: {start_date} ～ {end_date}</div>
        
        <div class="dashboard">
            <div class="chart-container full-width">
                <h2>週足変動率比較</h2>
                <iframe src="{os.path.basename(main_graph_path)}"></iframe>
            </div>
            
            <div class="chart-container full-width">
                <h2>銘柄パフォーマンス比較</h2>
                <iframe src="{os.path.basename(table_html_path)}" class="table-iframe"></iframe>
            </div>
            
            <div class="chart-container full-width">
                <h2>各銘柄のテクニカル分析</h2>
                <p>以下の銘柄をクリックすると、詳細なテクニカル分析グラフが表示されます：</p>
                <div class="stock-list">
    """
    
    stock_names_for_dashboard = sorted(list(all_data.keys()))
    for stock_key in stock_names_for_dashboard: 
        technical_filename = f'テクニカル分析_{stock_key}_{period_str}.html' 
        dashboard_html_content += f'<a href="{technical_filename}" class="stock-link" target="_blank">{stock_key.replace("_週足", "")}</a>\n'
    
    dashboard_html_content += """
                </div>
            </div>
        </div>
        
        <div class="chart-container full-width" style="margin-top: 30px; padding-bottom: 30px;">
            <h2>分析概要</h2>
            <p><strong>含まれる分析内容：</strong></p>
            <ul>
                <li><strong>変動率比較：</strong> 各銘柄の開始週を基準とした変動率(%)をインタラクティブに比較</li>
                <li><strong>テクニカル分析：</strong> ローソク足、移動平均線、ボリンジャーバンド、RSI、MACD、出来高分析（各銘柄ごと）</li>
                <li><strong>パフォーマンス指標：</strong> 全期間変動率、最大・最小週間リターン、ボラティリティ等</li>
            </ul>
            <p><strong>使用した主なテクニカル指標（週足ベース）：</strong></p>
            <ul>
                <li>移動平均線：4週, 13週, 26週, 52週</li>
                <li>ボリンジャーバンド：20週移動平均 ±1σ, ±2σ, ±3σ</li>
                <li>RSI：14週間の相対力指数</li>
                <li>MACD：12週EMA-26週EMA、9週シグナル線</li>
                <li>ストキャスティクス：14週%K、3週%D</li>
                <li>出来高：実際の出来高と20週移動平均線</li>
                <li>その他：月間高値・安値（4週）、52週高値・安値からの位置</li>
            </ul>
        </div>
    </body>
    </html>
    """
    
    dashboard_path = os.path.join(result_dir, f'週足株価分析ダッシュボード_{period_str}.html')
    with open(dashboard_path, 'w', encoding='utf-8') as f:
        f.write(dashboard_html_content)

    print("\n分析完了。以下のファイルが作成されました：")
    print(f"1. メイングラフ (変動率比較): {main_graph_path}")
    print(f"2. 銘柄パフォーマンス比較表: {table_html_path}")
    print(f"3. CSV (変動率データ): {csv_output}")
    
    # 各銘柄のテクニカル分析ファイルのパスを表示
    print("\n各銘柄のテクニカル分析チャート：")
    # Sort keys for consistent output order
    sorted_stock_keys = sorted(all_weekly_data.keys())
    for stock_key in sorted_stock_keys:
        technical_file = os.path.join(result_dir, f'テクニカル分析_{stock_key}_{period_str}.html')
        print(f"- {stock_key.replace('_週足','')}: {technical_file}") 
    
    print(f"\n総合ダッシュボード: {dashboard_path}")
    print(f"\nダッシュボードを開くには、上記の「総合ダッシュボード」のHTMLファイルをブラウザで開いてください。")
else:
    print("処理可能なデータがありませんでした。")
