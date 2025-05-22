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

# テクニカル指標を追加する関数（拡張版）
def add_technical_indicators(df):
    # 元のデータのコピーを作成
    result = df.copy()
    
    # 元々あった移動平均線 (4期間 - 約1年)
    result['MA4'] = df['Close'].rolling(window=4).mean()
    
    # 追加の移動平均線
    # 四半期足では25本で約6年分、50本で約12.5年分、75本で約18.75年分に相当
    result['MA_短期'] = df['Close'].rolling(window=25).mean()
    result['MA_中期'] = df['Close'].rolling(window=50).mean()
    result['MA_長期'] = df['Close'].rolling(window=75).mean()
    
    # 相対力指数 (RSI)
    delta = df['Close'].diff()
    gain = delta.where(delta > 0, 0)
    loss = -delta.where(delta < 0, 0)
    avg_gain = gain.rolling(window=4).mean()
    avg_loss = loss.rolling(window=4).mean()
    rs = avg_gain / avg_loss
    result['RSI'] = 100 - (100 / (1 + rs))
    
    # ボリンジャーバンド - 中心線と各標準偏差
    ma = df['Close'].rolling(window=4).mean()
    std = df['Close'].rolling(window=4).std()
    
    # 中心線（移動平均線）
    result['BB_中心線'] = ma
    
    # 上側バンド（+1σ, +2σ, +3σ）
    result['BB_上側1σ'] = ma + (std * 1)
    result['BB_上側2σ'] = ma + (std * 2)
    result['BB_上側3σ'] = ma + (std * 3)
    
    # 下側バンド（-1σ, -2σ, -3σ）
    result['BB_下側1σ'] = ma - (std * 1)
    result['BB_下側2σ'] = ma - (std * 2)
    result['BB_下側3σ'] = ma - (std * 3)
    
    return result

# ユーザーから日付範囲を取得
start_date, end_date = get_date_range()

# 四半期足データを処理して変動率と指標を計算する関数
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
        
        # すでに四半期足データなので、そのまま使用
        quarterly_data = df.copy()
        
        # 開始日の値を取得
        if quarterly_data.empty:
            print(f"警告: {file_path} の四半期データが空です。スキップします。")
            return None
        
        # テクニカル指標を追加
        quarterly_data = add_technical_indicators(quarterly_data)
            
        first_value = quarterly_data.iloc[0]['Close']
        
        # 変動率の計算（開始日を1.0として正規化）
        normalized_data = quarterly_data.copy()
        normalized_data['Normalized'] = quarterly_data['Close'] / first_value
        
        # 四半期ごとの変化率も計算
        normalized_data['Quarterly_Change'] = quarterly_data['Close'].pct_change() * 100
        
        print(f"処理完了: {file_path} - 四半期データ数: {len(normalized_data)}")
        return normalized_data
        
    except Exception as e:
        print(f"エラー発生: {file_path} の処理中にエラーが発生しました: {str(e)}")
        return None

# 全銘柄のデータを格納するリスト
all_data = {}
all_normalized_data = []
stock_names = []

# 各ファイルを処理
print(f"四半期足ファイル総数: {len(files)}")
for file in files:
    file_path = os.path.join(directory, file)
    stock_name = file.split('.')[0]  # ファイル名から銘柄名を取得
    
    print(f"処理中: {file}")
    normalized_data = process_quarterly_data(file_path, start_date, end_date)
    
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
    
    print(f"結合後のデータ形状: {combined_data.shape}")
    print(f"結合後のインデックス: {combined_data.index[:5]} ... {combined_data.index[-5:]}")
    
    # CSV形式での保存
    csv_path = os.path.join(result_dir, f'四半期足株価変動率比較_{period_str}.csv')
    combined_data.to_csv(csv_path)
    
    # 1. メイングラフ: 四半期足変動率比較
    fig = go.Figure()
    
    for column in combined_data.columns:
        # 日付を数字形式でフォーマット
        formatted_dates = [date.strftime('%Y/%m') for date in combined_data.index]
        
        # 各銘柄のデータをプロット
        fig.add_trace(
            go.Scatter(
                x=combined_data.index,
                y=combined_data[column],
                mode='lines+markers',
                name=column,
                customdata=formatted_dates,
                hovertemplate=
                '<b>%{fullData.name}</b><br>' +
                '日付: %{customdata}<br>' +  # customdataを使用して数字形式の日付を表示
                '変動率: %{y:.2%}<br>' +
                '終値: ' + all_data[column]['Close'].astype(str) +
                '<extra></extra>'
            )
        )
    
    # グラフのレイアウト設定
    fig.update_layout(
        title=f'四半期足変動率比較（{start_date}～{end_date}、開始日基準：100%）',  # タイトルを変更
        xaxis_title='四半期',
        yaxis=dict(
            type='log',  # 対数スケールを使用
            title='変動率（開始日=100%、対数スケール）',  # ラベルを変更
            tickformat='.0%'  # Y軸の目盛りを%形式で表示
        ),
        hovermode='closest',  # x unifiedから closestに変更
        template='plotly_white',
        legend_title='銘柄',
        # ホバーツールが画面端で切れないように調整
        hoverlabel=dict(
            bgcolor="white",
            bordercolor="black",
            font_size=12,
            align="auto"  # 自動調整で画面端での切れを防ぐ
        ),
        # マージンを調整してホバーツールの表示領域を確保
        margin=dict(l=50, r=80, t=80, b=50),
        updatemenus=[
            # 銘柄選択ドロップダウンメニュー
            {
                'buttons': [
                    {
                        'label': '全銘柄表示',
                        'method': 'update',
                        'args': [{'visible': [True] * len(combined_data.columns)}]
                    }
                ] + [
                    {
                        'label': stock,
                        'method': 'update',
                        'args': [{'visible': [i == j for j in range(len(combined_data.columns))]}]
                    }
                    for i, stock in enumerate(combined_data.columns)
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
    fig.update_xaxes(tickformat="%Y/%m")
    
    # HTMLとPNG形式で保存
    main_graph_path = os.path.join(result_dir, f'四半期足株価変動率比較_{period_str}.html')
    fig.write_html(main_graph_path)
    fig.write_image(os.path.join(result_dir, f'四半期足株価変動率比較_{period_str}.png'))
    
    # 2. 各銘柄の詳細テクニカル分析グラフ（拡張版）
    for stock in stock_names:
        # 銘柄ごとの詳細グラフを作成
        stock_data = all_data[stock]
        
        # 日付を数字形式でフォーマット
        formatted_dates = [date.strftime('%Y/%m') for date in stock_data.index]
        
        # サブプロット（株価、RSI、ボリンジャーバンド）
        technical_fig = sp.make_subplots(
            rows=2, cols=1,
            shared_xaxes=True,
            vertical_spacing=0.1,
            subplot_titles=(f'{stock} - 株価とテクニカル指標', 'RSI (相対力指数)'),
            row_heights=[0.7, 0.3]
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
                increasing_line_color='blue',      # 上昇時の色：青
                decreasing_line_color='red',       # 下降時の色：赤
                increasing_fillcolor='blue',
                decreasing_fillcolor='red',
                xhoverformat='%Y/%m'  # 日付表示を数字形式に変更
            ),
            row=1, col=1
        )
        
        # 移動平均線（元の4四半期移動平均線）
        technical_fig.add_trace(
            go.Scatter(
                x=stock_data.index,
                y=stock_data['MA4'],
                mode='lines',
                name='移動平均線 (4四半期)',
                line=dict(color='blue', width=1.5),
                customdata=formatted_dates,
                hovertemplate='<b>移動平均線 (4四半期)</b><br>' +
                             '日付: %{customdata}<br>' +
                             '値: %{y}<extra></extra>'
            ),
            row=1, col=1
        )
        
        # 追加の移動平均線
        # データが不足している場合は、Noneが含まれることになり、エラーになる可能性があるため、dropna()を使用
        if not stock_data['MA_短期'].isna().all():
            technical_fig.add_trace(
                go.Scatter(
                    x=stock_data.index,
                    y=stock_data['MA_短期'],
                    mode='lines',
                    name='短期移動平均線 (25)',
                    line=dict(color='orange', width=1.5),
                    customdata=formatted_dates,
                    hovertemplate='<b>短期移動平均線 (25)</b><br>' +
                                 '日付: %{customdata}<br>' +
                                 '値: %{y}<extra></extra>'
                ),
                row=1, col=1
            )
        
        if not stock_data['MA_中期'].isna().all():
            technical_fig.add_trace(
                go.Scatter(
                    x=stock_data.index,
                    y=stock_data['MA_中期'],
                    mode='lines',
                    name='中期移動平均線 (50)',
                    line=dict(color='green', width=1.5),
                    customdata=formatted_dates,
                    hovertemplate='<b>中期移動平均線 (50)</b><br>' +
                                 '日付: %{customdata}<br>' +
                                 '値: %{y}<extra></extra>'
                ),
                row=1, col=1
            )
        
        if not stock_data['MA_長期'].isna().all():
            technical_fig.add_trace(
                go.Scatter(
                    x=stock_data.index,
                    y=stock_data['MA_長期'],
                    mode='lines',
                    name='長期移動平均線 (75)',
                    line=dict(color='purple', width=1.5),
                    customdata=formatted_dates,
                    hovertemplate='<b>長期移動平均線 (75)</b><br>' +
                                 '日付: %{customdata}<br>' +
                                 '値: %{y}<extra></extra>'
                ),
                row=1, col=1
            )
        
        # ボリンジャーバンド（拡張版）
        # 中心線
        technical_fig.add_trace(
            go.Scatter(
                x=stock_data.index,
                y=stock_data['BB_中心線'],
                mode='lines',
                name='BB 中心線',
                line=dict(color='gray', width=1, dash='dot'),
                customdata=formatted_dates,
                hovertemplate='<b>BB 中心線</b><br>' +
                             '日付: %{customdata}<br>' +
                             '値: %{y}<extra></extra>'
            ),
            row=1, col=1
        )
        
        # +1σ
        technical_fig.add_trace(
            go.Scatter(
                x=stock_data.index,
                y=stock_data['BB_上側1σ'],
                mode='lines',
                name='+1σ',
                line=dict(color='rgba(0, 176, 246, 0.7)', width=1, dash='dot'),
                customdata=formatted_dates,
                hovertemplate='<b>+1σ</b><br>' +
                             '日付: %{customdata}<br>' +
                             '値: %{y}<extra></extra>'
            ),
            row=1, col=1
        )
        
        # +2σ
        technical_fig.add_trace(
            go.Scatter(
                x=stock_data.index,
                y=stock_data['BB_上側2σ'],
                mode='lines',
                name='+2σ',
                line=dict(color='rgba(0, 176, 246, 0.8)', width=1, dash='dot'),
                customdata=formatted_dates,
                hovertemplate='<b>+2σ</b><br>' +
                             '日付: %{customdata}<br>' +
                             '値: %{y}<extra></extra>'
            ),
            row=1, col=1
        )
        
        # +3σ
        technical_fig.add_trace(
            go.Scatter(
                x=stock_data.index,
                y=stock_data['BB_上側3σ'],
                mode='lines',
                name='+3σ',
                line=dict(color='rgba(0, 176, 246, 0.9)', width=1, dash='dot'),
                customdata=formatted_dates,
                hovertemplate='<b>+3σ</b><br>' +
                             '日付: %{customdata}<br>' +
                             '値: %{y}<extra></extra>'
            ),
            row=1, col=1
        )
        
        # -1σ
        technical_fig.add_trace(
            go.Scatter(
                x=stock_data.index,
                y=stock_data['BB_下側1σ'],
                mode='lines',
                name='-1σ',
                line=dict(color='rgba(255, 100, 100, 0.7)', width=1, dash='dot'),
                customdata=formatted_dates,
                hovertemplate='<b>-1σ</b><br>' +
                             '日付: %{customdata}<br>' +
                             '値: %{y}<extra></extra>'
            ),
            row=1, col=1
        )
        
        # -2σ
        technical_fig.add_trace(
            go.Scatter(
                x=stock_data.index,
                y=stock_data['BB_下側2σ'],
                mode='lines',
                name='-2σ',
                line=dict(color='rgba(255, 100, 100, 0.8)', width=1, dash='dot'),
                customdata=formatted_dates,
                hovertemplate='<b>-2σ</b><br>' +
                             '日付: %{customdata}<br>' +
                             '値: %{y}<extra></extra>'
            ),
            row=1, col=1
        )
        
        # -3σ
        technical_fig.add_trace(
            go.Scatter(
                x=stock_data.index,
                y=stock_data['BB_下側3σ'],
                mode='lines',
                name='-3σ',
                line=dict(color='rgba(255, 100, 100, 0.9)', width=1, dash='dot'),
                fill='tonexty',  # 最も下の線と+3σの間を塗りつぶし
                fillcolor='rgba(230, 230, 250, 0.1)',
                customdata=formatted_dates,
                hovertemplate='<b>-3σ</b><br>' +
                             '日付: %{customdata}<br>' +
                             '値: %{y}<extra></extra>'
            ),
            row=1, col=1
        )
        
        # RSIチャート（下段）
        technical_fig.add_trace(
            go.Scatter(
                x=stock_data.index,
                y=stock_data['RSI'],
                mode='lines+markers',
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
        
        # レイアウト設定
        technical_fig.update_layout(
            title=f'{stock} - テクニカル分析（{start_date}～{end_date}）',
            xaxis_title='四半期',
            yaxis_title='株価',
            yaxis2_title='RSI',
            hovermode='x unified',
            template='plotly_white',
            height=800,
            width=1000,
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            )
        )
        
        # X軸の日付表示を数字形式に変更
        technical_fig.update_xaxes(tickformat="%Y/%m")
        
        # y軸の範囲設定（RSI用）
        technical_fig.update_yaxes(range=[0, 100], row=2, col=1)
        
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
        
        # 四半期リターンの最大値と最小値
        if 'Quarterly_Change' in all_data[column].columns:
            quarterly_returns = all_data[column]['Quarterly_Change'].dropna()
            max_quarterly_return = quarterly_returns.max()
            min_quarterly_return = quarterly_returns.min()
        else:
            max_quarterly_return = np.nan
            min_quarterly_return = np.nan
            
        performance_data[column] = {
            'overall_change': overall_change,
            'max_gain': max_gain,
            'max_drawdown': max_drawdown,
            'max_quarterly_return': max_quarterly_return,
            'min_quarterly_return': min_quarterly_return
        }
    
    # パフォーマンスデータをテーブルとして表示
    table_data = [
        {'銘柄': stock, 
         '全期間変動率(%)': f"{performance_data[stock]['overall_change']:.2f}", 
         '最大上昇率(%)': f"{performance_data[stock]['max_gain']:.2f}", 
         '最大下落率(%)': f"{performance_data[stock]['max_drawdown']:.2f}",
         '最大四半期リターン(%)': f"{performance_data[stock]['max_quarterly_return']:.2f}",
         '最小四半期リターン(%)': f"{performance_data[stock]['min_quarterly_return']:.2f}"}
        for stock in stock_names
    ]
    
    # 全期間変動率でソート
    table_data = sorted(table_data, key=lambda x: float(x['全期間変動率(%)'].replace(',', '.')), reverse=True)
    
    table_fig = go.Figure(data=[go.Table(
        header=dict(
            values=['銘柄', '全期間変動率(%)', '最大上昇率(%)', '最大下落率(%)', '最大四半期リターン(%)', '最小四半期リターン(%)'],
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
                [data['最大四半期リターン(%)'] for data in table_data],
                [data['最小四半期リターン(%)'] for data in table_data]
            ],
            fill_color=[['white', 'lightgrey'] * (len(table_data) // 2 + len(table_data) % 2)],
            align='left',
            font=dict(size=11)
        )
    )])
    
    table_fig.update_layout(
        title='銘柄パフォーマンス比較（全期間変動率の降順）',
        height=400 + 25 * len(table_data),  # データ数に応じて高さを調整
        width=1000
    )
    
    # HTMLとして保存
    table_path = os.path.join(result_dir, f'銘柄パフォーマンス比較_{period_str}.html')
    table_fig.write_html(table_path)
    
    # 4. 総合ダッシュボード作成
    dashboard_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>株価分析ダッシュボード</title>
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background-color: #f5f5f5; }}
            .dashboard {{ display: flex; flex-wrap: wrap; gap: 20px; justify-content: space-between; }}
            .chart-container {{ background: white; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 15px; margin-bottom: 20px; }}
            .full-width {{ width: 100%; }}
            h1, h2 {{ color: #333; }}
            h1 {{ text-align: center; margin-bottom: 25px; border-bottom: 2px solid #ddd; padding-bottom: 15px; }}
            h2 {{ margin-top: 0; padding-bottom: 10px; border-bottom: 1px solid #eee; }}
            iframe {{ border: none; width: 100%; height: 600px; border-radius: 4px; }}
            .stock-list {{ display: flex; flex-wrap: wrap; gap: 10px; margin-top: 15px; }}
            .stock-link {{ padding: 8px 16px; background-color: #e9f5ff; border-radius: 6px; text-decoration: none; color: #0066cc; transition: all 0.2s; }}
            .stock-link:hover {{ background-color: #cce5ff; }}
            .date-info {{ text-align: center; font-size: 1.2em; margin-bottom: 25px; color: #666; }}
        </style>
    </head>
    <body>
        <h1>株価分析ダッシュボード</h1>
        <div class="date-info">分析期間: {start_date} ～ {end_date}</div>
        
        <div class="dashboard">
            <div class="chart-container full-width">
                <h2>四半期足変動率比較</h2>
                <iframe src="{os.path.basename(main_graph_path)}"></iframe>
            </div>
            
            <div class="chart-container full-width">
                <h2>銘柄パフォーマンス比較</h2>
                <iframe src="{os.path.basename(table_path)}" style="height: 400px;"></iframe>
            </div>
            
            <div class="chart-container full-width">
                <h2>各銘柄のテクニカル分析</h2>
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
    </body>
    </html>
    """
    
    # ダッシュボードHTMLの保存
    dashboard_path = os.path.join(result_dir, f'株価分析ダッシュボード_{period_str}.html')
    with open(dashboard_path, 'w', encoding='utf-8') as f:
        f.write(dashboard_html)
    
    print(f"分析完了。以下のファイルが作成されました：")
    print(f"1. メイングラフ: {main_graph_path}")
    print(f"2. 銘柄パフォーマンス比較: {table_path}")
    print(f"3. 各銘柄のテクニカル分析: {result_dir}/*.html")
    print(f"4. 総合ダッシュボード: {dashboard_path}")
    print(f"5. CSV: {csv_path}")
    
    print(f"\nダッシュボードを開くには、以下のファイルをブラウザで開いてください：")
    print(f"{dashboard_path}")
else:
    print("指定した期間内のデータが見つかりませんでした。")