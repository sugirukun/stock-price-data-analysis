import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime
import os
import time
import glob
import shutil

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

# カラム名の対応辞書
japanese_columns = {
    'Open': '始値',
    'High': '高値',
    'Low': '安値',
    'Close': '終値',
    'Volume': '出来高',
    'Dividends': '配当',
    'Stock Splits': '株式分割'
}

english_columns = {v: k for k, v in japanese_columns.items()}  # 逆変換用辞書

def get_stock_data():
    """
    stock_data_all.pyの機能を使用して株価データを取得する
    """
    # カラム名の日本語表示の選択
    use_japanese_columns = input("カラムの表示を日本語にしますか？(y/n)\n※日本語表記にすることでこのプログラムの実行は問題は出ませんが、より詳細な分析をするときは支障が出ることがあります: ").strip().lower() == 'y'

    # 期間タイプの定義
    frequency_types = {
        "日足": "1d",
        "週足": "1wk",
        "月足": "1mo"
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
    print("株価データの取得が完了しました。")
    
    return use_japanese_columns

def create_quarterly_data(use_japanese_columns):
    """
    保存された日足データから四半期足データを作成する
    """
    print("\n四半期足データの作成を開始します...")
    
    # 各銘柄に対して処理
    for ticker, name in tickers.items():
        safe_ticker = ticker.replace('.', '_')
        
        try:
            # 日足データを読み込む
            daily_file_path = os.path.join(output_dir, f"{safe_ticker}_{name}_日足.csv")
            if not os.path.exists(daily_file_path):
                print(f"日足データファイルが見つかりません: {daily_file_path}")
                continue
                
            daily_data = pd.read_csv(daily_file_path, index_col=0, parse_dates=True)
            
            # カラム名を確認して必要に応じて変換
            if '始値' in daily_data.columns:
                # 日本語カラム名の場合、Pandas の処理のために一時的に英語に戻す
                daily_data.rename(columns={v: k for k, v in japanese_columns.items() if v in daily_data.columns}, inplace=True)
            
            # 四半期の最後の日を持つデータに絞る
            daily_data['quarter'] = daily_data.index.to_period('Q')
            quarterly_data = daily_data.groupby('quarter').agg({
                'Open': 'first',
                'High': 'max',
                'Low': 'min',
                'Close': 'last',
                'Volume': 'sum'
            })
            
            # インデックスを日付形式に戻す
            quarterly_data.index = quarterly_data.index.to_timestamp()
            
            # カラム名を指定された形式に戻す
            if use_japanese_columns:
                quarterly_data.rename(columns=japanese_columns, inplace=True)
            
            # CSVとして保存
            quarterly_csv_path = os.path.join(output_dir, f"{safe_ticker}_{name}_四半期足.csv")
            quarterly_data.to_csv(quarterly_csv_path)
            print(f"四半期足データを保存しました: {quarterly_csv_path}")
            
            # 四半期データの確認
            print(f"{name}の四半期足データ（先頭5行）:")
            print(quarterly_data.head())
            
        except Exception as e:
            print(f"{name}の四半期足データ作成中にエラーが発生しました: {e}")
    
    print("四半期足データの作成が完了しました。")

def create_yearly_data(use_japanese_columns):
    """
    保存された日足データから年足データを作成する
    """
    print("\n年足データの作成を開始します...")
    
    # 各銘柄に対して処理
    for ticker, name in tickers.items():
        safe_ticker = ticker.replace('.', '_')
        
        try:
            # 日足データを読み込む
            daily_file_path = os.path.join(output_dir, f"{safe_ticker}_{name}_日足.csv")
            if not os.path.exists(daily_file_path):
                print(f"日足データファイルが見つかりません: {daily_file_path}")
                continue
                
            daily_data = pd.read_csv(daily_file_path, index_col=0, parse_dates=True)
            
            # カラム名を確認して必要に応じて変換
            if '始値' in daily_data.columns:
                # 日本語カラム名の場合、Pandas の処理のために一時的に英語に戻す
                daily_data.rename(columns={v: k for k, v in japanese_columns.items() if v in daily_data.columns}, inplace=True)
            
            # 年の最後の日を持つデータに絞る
            daily_data['year'] = daily_data.index.to_period('Y')
            yearly_data = daily_data.groupby('year').agg({
                'Open': 'first',
                'High': 'max',
                'Low': 'min',
                'Close': 'last',
                'Volume': 'sum'
            })
            
            # インデックスを日付形式に戻す
            yearly_data.index = yearly_data.index.to_timestamp()
            
            # カラム名を指定された形式に戻す
            if use_japanese_columns:
                yearly_data.rename(columns=japanese_columns, inplace=True)
            
            # CSVとして保存
            yearly_csv_path = os.path.join(output_dir, f"{safe_ticker}_{name}_年足.csv")
            yearly_data.to_csv(yearly_csv_path)
            print(f"年足データを保存しました: {yearly_csv_path}")
            
            # 年足データの確認
            print(f"{name}の年足データ（先頭5行）:")
            print(yearly_data.head())
            
        except Exception as e:
            print(f"{name}の年足データ作成中にエラーが発生しました: {e}")
    
    print("年足データの作成が完了しました。")

def update_excel_with_quarterly_yearly_data(use_japanese_columns):
    """
    既存のExcelファイルに四半期足と年足のデータを追加する
    """
    print("\nExcelファイルに四半期足と年足のデータを追加しています...")
    
    for ticker, name in tickers.items():
        safe_ticker = ticker.replace('.', '_')
        excel_path = os.path.join(output_dir, f"{safe_ticker}_{name}.xlsx")
        
        if not os.path.exists(excel_path):
            print(f"Excelファイルが見つかりません: {excel_path}")
            continue
        
        try:
            # 一時ファイルとして新しいExcelを作成
            temp_excel_path = os.path.join(output_dir, f"{safe_ticker}_{name}_temp.xlsx")
            
            # 既存のExcelからシートを読み込む
            xls = pd.ExcelFile(excel_path)
            existing_sheets = xls.sheet_names
            
            # 四半期足と年足のデータを読み込む
            quarterly_csv_path = os.path.join(output_dir, f"{safe_ticker}_{name}_四半期足.csv")
            yearly_csv_path = os.path.join(output_dir, f"{safe_ticker}_{name}_年足.csv")
            
            with pd.ExcelWriter(temp_excel_path, engine='openpyxl') as writer:
                # まず既存のシートを保存（四半期足と年足以外）
                for sheet_name in existing_sheets:
                    if not ('四半期足' in sheet_name or '年足' in sheet_name):
                        sheet_data = pd.read_excel(excel_path, sheet_name=sheet_name, index_col=0)
                        sheet_data.to_excel(writer, sheet_name=sheet_name)
                
                # 四半期足データを追加
                if os.path.exists(quarterly_csv_path):
                    quarterly_data = pd.read_csv(quarterly_csv_path, index_col=0, parse_dates=True)
                    sheet_name = f"{name[:15]}_四半期足" if len(name) > 15 else f"{name}_四半期足"
                    sheet_name = sheet_name[:31]  # シート名の制限のため31文字まで
                    quarterly_data.to_excel(writer, sheet_name=sheet_name)
                    print(f"Excelに四半期足シートを追加しました: {sheet_name}")
                
                # 年足データを追加
                if os.path.exists(yearly_csv_path):
                    yearly_data = pd.read_csv(yearly_csv_path, index_col=0, parse_dates=True)
                    sheet_name = f"{name[:15]}_年足" if len(name) > 15 else f"{name}_年足"
                    sheet_name = sheet_name[:31]  # シート名の制限のため31文字まで
                    yearly_data.to_excel(writer, sheet_name=sheet_name)
                    print(f"Excelに年足シートを追加しました: {sheet_name}")
            
            # 一時ファイルを元のファイルに置き換える
            if os.path.exists(temp_excel_path):
                # 元のファイルが残っている場合は削除
                if os.path.exists(excel_path):
                    os.remove(excel_path)
                # 一時ファイルを元のファイル名にリネーム
                os.rename(temp_excel_path, excel_path)
                print(f"更新されたExcelファイルを保存しました: {excel_path}")
            
        except Exception as e:
            print(f"{name}のExcel更新中にエラーが発生しました: {e}")
    
    print("すべてのExcelファイルの更新が完了しました。")

def main():
    print("===== 株価データ統合処理スクリプト =====")
    
    # 1. 株価データの取得
    print("\n【ステップ1】株価データの取得")
    use_japanese_columns = get_stock_data()
    
    # 2. 四半期足データの作成
    print("\n【ステップ2】四半期足データの作成")
    create_quarterly_data(use_japanese_columns)
    
    # 3. 年足データの作成
    print("\n【ステップ3】年足データの作成")
    create_yearly_data(use_japanese_columns)
    
    # 4. Excelファイルの更新
    print("\n【ステップ4】Excelファイルの更新")
    update_excel_with_quarterly_yearly_data(use_japanese_columns)
    
    print("\n===== 株価データ統合処理が完了しました =====")

if __name__ == "__main__":
    main()
