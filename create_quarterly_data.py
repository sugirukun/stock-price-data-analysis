import os
import pandas as pd
import glob
from datetime import datetime

def convert_monthly_to_quarterly(monthly_file_path):
    """
    月足データから四半期足データを作成する関数
    
    Parameters:
    monthly_file_path (str): 月足データファイルのパス
    
    Returns:
    pd.DataFrame: 四半期足データのデータフレーム
    """
    # ファイル名から銘柄名を取得
    file_name = os.path.basename(monthly_file_path)
    ticker_name = file_name.replace('_月足.csv', '')
    
    # 月足データの読み込み
    df_monthly = pd.read_csv(monthly_file_path)
    
    # カラム名をチェックし、英語か日本語かを判断
    if 'Open' in df_monthly.columns:
        # 英語カラム名の場合
        open_col = 'Open'
        high_col = 'High'
        low_col = 'Low'
        close_col = 'Close'
        volume_col = 'Volume'
        dividend_col = 'Dividends'
        split_col = 'Stock Splits'
    else:
        # 日本語カラム名の場合
        open_col = '始値'
        high_col = '高値'
        low_col = '安値'
        close_col = '終値'
        volume_col = '出来高'
        dividend_col = '配当'
        split_col = '株式分割'
    
    # 日付をdatetime型に変換
    df_monthly['Date'] = pd.to_datetime(df_monthly['Date'], utc=True)
    
    # 年と四半期の列を追加
    df_monthly['Year'] = df_monthly['Date'].dt.year
    df_monthly['Quarter'] = (df_monthly['Date'].dt.month - 1) // 3 + 1
    
    # 年と四半期でグループ化
    quarterly_groups = df_monthly.groupby(['Year', 'Quarter'])
    
    # 四半期データの作成
    quarterly_data = []
    
    for (year, quarter), group in quarterly_groups:
        # 四半期の最初の月の始値を取得
        open_price = group.iloc[0][open_col]
        
        # 四半期中の最高値
        high_price = group[high_col].max()
        
        # 四半期中の最安値
        low_price = group[low_col].min()
        
        # 四半期の最後の月の終値
        close_price = group.iloc[-1][close_col]
        
        # 四半期の出来高の合計
        volume = group[volume_col].sum()
        
        # 四半期の配当の合計
        dividend = group[dividend_col].sum()
        
        # 株式分割情報（任意の処理方法）
        # ここでは単純に最大値を取得
        stock_split = group[split_col].max()
        
        # 四半期の日付（最初の月の日を1日にする）
        quarter_start_month = (quarter - 1) * 3 + 1
        quarter_date = datetime(year, quarter_start_month, 1)
        
        # 入力ファイルのカラム名の言語に合わせてデータを追加
        if 'Open' in df_monthly.columns:
            # 英語カラム名の場合
            quarterly_data.append({
                'Date': quarter_date,
                'Open': open_price,
                'High': high_price,
                'Low': low_price,
                'Close': close_price,
                'Volume': volume,
                'Dividends': dividend,
                'Stock Splits': stock_split
            })
        else:
            # 日本語カラム名の場合
            quarterly_data.append({
                'Date': quarter_date,
                '始値': open_price,
                '高値': high_price,
                '安値': low_price,
                '終値': close_price,
                '出来高': volume,
                '配当': dividend,
                '株式分割': stock_split
            })
    
    # データフレームに変換
    df_quarterly = pd.DataFrame(quarterly_data)
    
    # 日付でソート
    df_quarterly = df_quarterly.sort_values('Date')
    
    # データフレームのインデックスをリセット
    df_quarterly = df_quarterly.reset_index(drop=True)
    
    return df_quarterly, ticker_name

def main():
    """
    メイン関数：株価データフォルダの全ての月足データを四半期足に変換
    """
    # 株価データフォルダのパス
    data_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), '株価データ')
    
    # 月足データファイルを検索
    monthly_files = glob.glob(os.path.join(data_folder, '*_月足.csv'))
    
    print(f"変換対象ファイル数: {len(monthly_files)}")
    
    # 各月足ファイルを四半期足に変換
    for monthly_file in monthly_files:
        try:
            df_quarterly, ticker_name = convert_monthly_to_quarterly(monthly_file)
            
            # 出力ファイルパス
            output_path = os.path.join(data_folder, f"{ticker_name}_四半期足.csv")
            
            # CSVファイルに保存（エンコーディングを明示的に指定）
            df_quarterly.to_csv(output_path, index=False, encoding='utf-8-sig')
            
            print(f"変換完了: {ticker_name}")
        except Exception as e:
            print(f"エラー ({os.path.basename(monthly_file)}): {e}")
    
    print("処理完了")

if __name__ == "__main__":
    main()
