import yfinance as yf
import pandas as pd
from datetime import datetime
import os
import time

# 現在の日付を取得（ファイル名用）
today = datetime.now().strftime("%Y%m%d")

# ユーザーにティッカーシンボルの入力を求める
print("株価データを取得するティッカーシンボルを入力してください。")
print("複数のティッカーシンボルを入力する場合はカンマ(,)で区切ってください。")
print("例: 7974.T,1387.T,AAPL")
print("※日本株は '.T'をつけてください (例: 7974.T)")
print("\n特殊な銘柄のティッカーシンボル例:")
print("- 為替レート: JPY=X (米ドル/円), EUR=X (ユーロ/米ドル), EURJPY=X (ユーロ/円)")
print("- 株価指数: ^N225 (日経平均), ^TPX (TOPIX), ^DJI (NYダウ), ^SOX (半導体指数), ^GSPC (S&P500)")
print("- ETF: 純金上場信託は '1540.T' または '1540.JP' を試してください")

# ユーザー入力からティッカーシンボルを取得
ticker_input = input("ティッカーシンボル: ").strip()
ticker_list = [t.strip() for t in ticker_input.split(',') if t.strip()]

if not ticker_list:
    print("ティッカーシンボルが入力されていません。例として以下の銘柄を使用します。")
    tickers = {
        "7974.T": "任天堂",         # 任天堂
        "1387.T": "eMAXIS Slim米国株式(S&P500)",  # eMAXIS Slim米国株式
        "AAPL": "アップル"          # Apple
    }
else:
    # ティッカーシンボルから銘柄名を取得（または入力を求める）
    tickers = {}
    for ticker in ticker_list:
        try:
            # 特定の既知のティッカーに対するデフォルト名を設定
            default_names = {
                "JPY=X": "米ドル円",
                "EURJPY=X": "ユーロ円",
                "^N225": "日経平均",
                "^TPX": "TOPIX",
                "TPX.I": "TOPIX",
                "1306.T": "TOPIX連動ETF",
                "1348.T": "MAXIS_TOPIX",
                "^DJI": "NYダウ",
                "^SOX": "SOX半導体指数",
                "^GSPC": "S&P500",
                "1540.T": "純金上場信託",
                "1540.JP": "純金上場信託"
            }
            
            if ticker in default_names:
                suggested_name = default_names[ticker]
                print(f"{ticker}の銘柄名として「{suggested_name}」を使用します。")
                tickers[ticker] = suggested_name
                continue
            
            # Yahoo FinanceのAPIからティッカー情報を取得
            stock_info = yf.Ticker(ticker).info
            if 'shortName' in stock_info and stock_info['shortName']:
                suggested_name = stock_info['shortName']
                print(f"{ticker}の銘柄名として「{suggested_name}」が見つかりました。")
                use_suggested = input(f"この銘柄名を使用しますか？(y/n、nの場合は手動入力): ").strip().lower() == 'y'
                
                if use_suggested:
                    tickers[ticker] = suggested_name
                else:
                    manual_name = input(f"{ticker}の銘柄名を入力してください: ").strip()
                    tickers[ticker] = manual_name
            else:
                manual_name = input(f"{ticker}の銘柄名を入力してください: ").strip()
                tickers[ticker] = manual_name
        except Exception as e:
            print(f"警告: {ticker}の情報取得中にエラーが発生しました: {e}")
            manual_name = input(f"{ticker}の銘柄名を入力してください: ").strip()
            tickers[ticker] = manual_name

# 出力ディレクトリを確保（ローカルフォルダに保存）
output_dir = "C:\\Users\\rilak\\Desktop\\株価\\株価データ"
os.makedirs(output_dir, exist_ok=True)
print(f"データをローカルフォルダに保存します: {output_dir}")

# カラム名は英語のまま使用します
use_japanese_columns = False
print("カラム名は英語表記を使用します。")

# 期間タイプの定義
frequency_types = {
    "日足": "1d",
    "週足": "1wk",
    "月足": "1mo",
    "年足": "1y"
}

# 各銘柄について株価データを取得
for ticker, name in tickers.items():
    print(f"{ticker}（{name}）の株価データを取得中...")
    
    try:
        # yfinanceで銘柄オブジェクトを取得
        stock = yf.Ticker(ticker)
        
        # 各期間タイプのデータを取得
        period_data = {}
        for period_name, period_code in frequency_types.items():
            print(f"{period_name}データを取得中...")
            data = stock.history(period="max", interval=period_code)
            
            # データが空でないか確認
            if data.empty:
                print(f"{period_name}のデータがありません: {ticker}")
                continue
            
            # データの行数と期間を表示
            start_date = data.index[0].strftime('%Y-%m-%d')
            end_date = data.index[-1].strftime('%Y-%m-%d')
            print(f"{period_name}の取得期間: {start_date}から{end_date}まで")
            print(f"{period_name}のデータ点数: {len(data)}日分")
            
            # 期間データを保存
            period_data[period_name] = data
        
        # 安全なファイル名を生成（ティッカー記号からピリオドを除去）
        safe_ticker = ticker.replace('.', '_')
        
        # カラム名を日本語に変更
        japanese_columns = {
            'Open': '始値',
            'High': '高値',
            'Low': '安値',
            'Close': '終値',
            'Volume': '出来高',
            'Dividends': '配当',
            'Stock Splits': '株式分割'
        }
        
        # CSVファイルとして各期間データを保存
        for period_name, data in period_data.items():
            # データのコピーを作成
            processed_data = data.copy()
            
            # ユーザー選択に基づいてカラム名を変更
            if use_japanese_columns:
                processed_data.rename(columns=japanese_columns, inplace=True)
                print(f"{period_name}のカラム名を日本語に変換しました。")
            else:
                print(f"{period_name}のカラム名は英語のままです。")
            
            # 銘柄名からファイル名に使えない文字を削除
            safe_name = name.replace('/', '').replace('\\', '').replace(':', '').replace('*', '').replace('?', '').replace('"', '').replace('<', '').replace('>', '').replace('|', '')
            
            # CSVファイルとして保存（エンコーディングを明示的に指定）
            csv_path = os.path.join(output_dir, f"{safe_ticker}_{safe_name}_{period_name}.csv")
            processed_data.to_csv(csv_path, encoding='utf-8-sig')
            print(f"{period_name}のCSVファイルを保存しました: {csv_path}")
            
            # データの最初と最後の5行を表示
            print(f"{period_name}の最初の5日分のデータ:")
            print(processed_data.head())
            
            print(f"{period_name}の最新の5日分のデータ:")
            print(processed_data.tail())
            print("\n" + "-"*50 + "\n")  # 区切り線
        
        # 銘柄名からファイル名に使えない文字を削除
        safe_name = name.replace('/', '').replace('\\', '').replace(':', '').replace('*', '').replace('?', '').replace('"', '').replace('<', '').replace('>', '').replace('|', '')
        
        # Excelファイルとして全期間データを1つのファイルに保存
        excel_path = os.path.join(output_dir, f"{safe_ticker}_{safe_name}.xlsx")
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            for period_name, data in period_data.items():
                # データのコピーを作成
                excel_data = data.copy()
                
                # ユーザー選択に基づいてカラム名を変更
                if use_japanese_columns:
                    excel_data.rename(columns=japanese_columns, inplace=True)
                
                # タイムゾーン情報を削除
                excel_data.index = excel_data.index.tz_localize(None)
                
                # インデックスの名前を変更
                excel_data.index.name = '日付' if use_japanese_columns else 'Date'
                
                # シート名（31文字以内に制限）
                sheet_name = f"{safe_name[:15]}_{period_name}" if len(safe_name) > 15 else f"{safe_name}_{period_name}"
                sheet_name = sheet_name[:31]  # シート名の制限のため31文字まで
                
                # シートにデータを書き込み
                excel_data.to_excel(writer, sheet_name=sheet_name)
                
            print(f"Excelファイルを保存しました（全期間データ）: {excel_path}")
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
    
    # 連続リクエストによるAPIの制限を避けるため少し待つ
    time.sleep(1)
    
    print("\n" + "="*80 + "\n")  # 区切り線
print("処理が完了しました。全てのデータをローカルフォルダに保存しました。")