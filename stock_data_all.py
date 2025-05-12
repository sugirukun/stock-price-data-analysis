import yfinance as yf
import pandas as pd
from datetime import datetime
import o
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
            
            # CSVファイルとして保存
            csv_path = os.path.join(output_dir, f"{safe_ticker}_{name}_{period_name}.csv")
            processed_data.to_csv(csv_path)
            print(f"{period_name}のCSVファイルを保存しました: {csv_path}")
            
            # データの最初と最後の5行を表示
            column_type = "日本語カラム" if use_japanese_columns else "英語カラム"
            print(f"{period_name}の最初の5日分のデータ（{column_type}）:")
            print(processed_data.head())
            
            print(f"{period_name}の最新の5日分のデータ（{column_type}）:")
            print(processed_data.tail())
            print("\n" + "-"*50 + "\n")  # 区切り線
        
        # Excelファイルとして全期間データを1つのファイルに保存
        excel_path = os.path.join(output_dir, f"{safe_ticker}_{name}.xlsx")
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
                sheet_name = f"{name[:15]}_{period_name}" if len(name) > 15 else f"{name}_{period_name}"
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