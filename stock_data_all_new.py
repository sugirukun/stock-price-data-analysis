import yfinance as yf
import pandas as pd
from datetime import datetime
import os
import time
import json
import sys

# 現在の日付を取得（ファイル名用）
today = datetime.now().strftime("%Y%m%d")

# 設定ファイルのパス - 絶対パスを使用
CONFIG_FILE = "C:\\Users\\rilak\\Desktop\\株価\\stock_config.json"

# 出力ディレクトリを固定値で設定
OUTPUT_DIR = "C:\\Users\\rilak\\Desktop\\株価\\株価データ"

def load_config():
    """設定ファイルを読み込む、存在しない場合はデフォルト設定を返す"""
    default_config = {
        "tickers": [
            {"symbol": "7974.T", "name": "任天堂"},
            {"symbol": "1387.T", "name": "eMAXIS Slim米国株式(S&P500)"},
            {"symbol": "AAPL", "name": "アップル"}
        ],
        "use_japanese_columns": False,
        "output_dir": OUTPUT_DIR
    }
    
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                print(f"設定ファイル '{CONFIG_FILE}' を読み込みました。")
                
                # 出力ディレクトリは常に固定値を使用
                config["output_dir"] = OUTPUT_DIR
                
                return config
        except Exception as e:
            print(f"設定ファイルの読み込み中にエラーが発生しました: {e}")
            print("デフォルト設定を使用します。")
    else:
        print(f"設定ファイル '{CONFIG_FILE}' が見つかりません。デフォルト設定を使用します。")
        print("設定を管理するには stock_config_manager.py を実行してください。")
    
    return default_config

def main():
    # 設定ファイルを読み込む
    config = load_config()
    
    # 設定から情報を取得
    ticker_config = config.get("tickers", [])
    use_japanese_columns = config.get("use_japanese_columns", False)
    
    # 出力ディレクトリを常に固定値に設定
    output_dir = OUTPUT_DIR
    
    # ティッカー情報をディクショナリに変換
    tickers = {item["symbol"]: item["name"] for item in ticker_config}
    
    # 設定情報を表示
    print("\n===== 株価データ取得ツール =====")
    print(f"出力ディレクトリ: {output_dir}")
    print(f"カラム名: {'日本語' if use_japanese_columns else '英語'}")
    print("\n取得対象の銘柄:")
    for symbol, name in tickers.items():
        print(f"- {symbol} ({name})")
    
    # コマンドライン引数で実行モードを確認
    # "-auto"フラグがある場合は自動実行モード
    auto_mode = "-auto" in sys.argv
    
    if not auto_mode and not tickers:
        print("\n銘柄リストが空です。ティッカーシンボルを入力してください。")
        print("複数のティッカーシンボルを入力する場合はカンマ(,)で区切ってください。")
        print("例: 7974.T,1387.T,AAPL")
        print("※日本株は '.T'をつけてください (例: 7974.T)")
        print("\n特殊な銘柄のティッカーシンボル例:")
        print("- 為替レート: JPY=X (米ドル/円), EUR=X (ユーロ/米ドル), EURJPY=X (ユーロ/円)")
        print("- 株価指数: ^N225 (日経平均), ^TPX (TOPIX), ^DJI (NYダウ), ^SOX (半導体指数), ^GSPC (S&P500)")
        print("- ETF: 純金上場信託は '1540.T' または '1540.JP' を試してください")
        
        # ユーザー入力からティッカーシンボルを取得
        ticker_input = input("\nティッカーシンボル: ").strip()
        ticker_list = [t.strip() for t in ticker_input.split(',') if t.strip()]
        
        if ticker_list:
            # ティッカーシンボルから銘柄名を取得（または入力を求める）
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
    
    # 出力ディレクトリを確保
    try:
        os.makedirs(output_dir, exist_ok=True)
        print(f"\nデータをローカルフォルダに保存します: {output_dir}")
    except Exception as e:
        print(f"警告: 出力ディレクトリの作成中にエラーが発生しました: {e}")
    
    # カラム名の表示設定
    if use_japanese_columns:
        print("カラム名は日本語表記を使用します。")
    else:
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
        print(f"\n{ticker}（{name}）の株価データを取得中...")
        
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

if __name__ == "__main__":
    main()
