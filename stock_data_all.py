import yfinance as yf
import pandas as pd
from datetime import datetime
import os
import time
# 現在の日付を取得（ファイル名用）
today = datetime.now().strftime("%Y%m%d")
# 取得したい銘柄のリスト
tickers = {
    "7974.T": "任天堂",         # 任天堂の証券コード (東証)
    "1387.T": "eMAXIS Slim米国株式(S&P500)",  # eMAXIS Slim米国株式(S&P500)のETFコード
    "AAPL": "アップル"          # Apple (米国株)
}
# 出力ディレクトリを確保（ローカルフォルダに保存）
output_dir = "C:\\Users\\rilak\\Desktop\\株価\\株価データ"
os.makedirs(output_dir, exist_ok=True)
print(f"データをローカルフォルダに保存します: {output_dir}")

# カラム名の日本語表示の選択
use_japanese_columns = input("カラムの表示を日本語にしますか？(y/n)\n※日本語表記にすることでこのプログラムの実行は問題は出ませんが、より詳細な分析をするときは支障が出ることがあります: ").strip().lower() == 'y'
# 各銘柄について株価データを取得
for ticker, name in tickers.items():
    print(f"{ticker}（{name}）の株価データを取得中...")
    
    try:
        # yfinanceで株価データを取得（可能な限り古いデータから）
        stock = yf.Ticker(ticker)
        data = stock.history(period="max")
        
        # データが空でないか確認
        if data.empty:
            print(f"データがありません: {ticker}")
            continue
        
        # データの行数と期間を表示
        start_date = data.index[0].strftime('%Y-%m-%d')
        end_date = data.index[-1].strftime('%Y-%m-%d')
        print(f"取得期間: {start_date}から{end_date}まで")
        print(f"データ点数: {len(data)}日分")
        
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
        
        # データのコピーを作成
        processed_data = data.copy()
        
        # ユーザー選択に基づいてカラム名を変更
        if use_japanese_columns:
            processed_data.rename(columns=japanese_columns, inplace=True)
            print("カラム名を日本語に変換しました。")
        else:
            print("カラム名は英語のままです。")
        
        # 安全なファイル名を生成（ティッカー記号からピリオドを除去）
        safe_ticker = ticker.replace('.', '_')
        
        # CSVファイルとして保存
        csv_path = os.path.join(output_dir, f"{safe_ticker}_{name}.csv")
        processed_data.to_csv(csv_path)
        print(f"CSVファイルを保存しました: {csv_path}")
        
        # タイムゾーン情報を削除してからExcelファイルとして保存
        excel_data = processed_data.copy()
        excel_data.index = excel_data.index.tz_localize(None)
        
        # インデックスの名前を変更
        excel_data.index.name = '日付' if use_japanese_columns else 'Date'
        
        # Excelファイルとして保存
        excel_path = os.path.join(output_dir, f"{safe_ticker}_{name}.xlsx")
        excel_data.to_excel(excel_path, sheet_name=name[:31])  # シート名の制限のため31文字まで
        print(f"Excelファイルを保存しました: {excel_path}")
        
        # データの最初と最後の5行を表示
        column_type = "日本語カラム" if use_japanese_columns else "英語カラム"
        print(f"最初の5日分のデータ（{column_type}）:")
        print(processed_data.head())
        
        print(f"最新の5日分のデータ（{column_type}）:")
        print(processed_data.tail())
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
    
    # 連続リクエストによるAPIの制限を避けるため少し待つ
    time.sleep(1)
    
    print("\n" + "="*80 + "\n")  # 区切り線
print("処理が完了しました。全てのデータをローカルフォルダに保存しました。")