import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta
import os
import json
import time
import shutil
# from collections import OrderedDict # OrderedDict は常に必要ではないため、必要な場所でインポート

# 設定ファイルのパス
CONFIG_FILE = "C:\\Users\\rilak\\Desktop\\株価\\stock_config.json"

# 出力ディレクトリを設定
BASE_OUTPUT_DIR = "C:\\Users\\rilak\\Desktop\\株価\\株価データ"
FAVORITES_OUTPUT_DIR = os.path.join(BASE_OUTPUT_DIR, "お気にり銘柄の更新")

def load_config():
    """設定ファイルを読み込む"""
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config = json.load(f)
            print(f"設定ファイル '{CONFIG_FILE}' を読み込みました。")
            return config
    except Exception as e:
        print(f"設定ファイルの読み込み中にエラーが発生しました: {e}")
        return None

def process_data_with_holidays(data, period_name):
    """市場休業日を含む連続したデータを作成する"""
    try:
        # データが空の場合は処理しない
        if data.empty:
            return data
        
        # データのコピーを作成
        processed_data = data.copy()
        
        # タイムゾーン情報を削除
        if processed_data.index.tz is not None:
            processed_data.index = processed_data.index.tz_localize(None)
        
        # 連続した日付インデックスを作成
        start_date = processed_data.index.min()
        end_date = processed_data.index.max()
        
        # 期間タイプに基づいてfrequencyを設定
        if period_name == "日足":
            freq = "D"  # 毎日
            # 開始日から終了日までの連続した日付を作成
            all_dates = pd.date_range(start=start_date, end=end_date, freq=freq)
        elif period_name == "週足":
            # 週足の場合は、開始日と同じ曜日から始まる週ごとの日付を作成
            freq = "W-" + start_date.strftime('%a')  # 例：W-Mon（月曜日）
            all_dates = pd.date_range(start=start_date, end=end_date, freq=freq)
        else:
            # 月足と年足は処理しない（そのまま返す）
            return processed_data
        
        # 元のデータを連続した日付でリインデックス
        reindexed_data = processed_data.reindex(all_dates)
        
        # 欠損値を前の値で埋める（前方補完）
        filled_data = reindexed_data.ffill()
        
        # 最初の日のデータが欠損している場合は、後ろの値で埋める（後方補完）
        filled_data = filled_data.bfill()
        
        # 元のデータに含まれる日付を取得（実際の取引日）
        original_dates = set(processed_data.index)
        
        # 休業日情報を追加（元のデータにある日付は営業日、ない日付は休業日）
        filled_data['Is_Market_Holiday'] = ~filled_data.index.isin(original_dates)
        
        return filled_data
        
    except Exception as e:
        print(f"休業日データ処理中にエラーが発生しました: {str(e)}")
        import traceback
        print(traceback.format_exc())
        return data  # エラー時は元のデータを返す

def update_csv_file(symbol, name):
    """CSVファイルを更新する（複数日のデータ抜け補填対応）"""
    try:
        # 元のCSVファイルのパス（7974_T形式）
        formatted_symbol = symbol.replace('.', '_')
        # 元データが存在するベースディレクトリ
        base_source_file = os.path.join(BASE_OUTPUT_DIR, f"{formatted_symbol}_{name}_日足.csv")
        
        if not os.path.exists(base_source_file):
            print(f"元のデータファイルが見つかりません: {base_source_file}")
            return False
            
        # 出力ディレクトリが存在しない場合は作成
        os.makedirs(FAVORITES_OUTPUT_DIR, exist_ok=True)
        
        # 更新用ファイルのパス（お気に入りディレクトリ内）
        target_file = os.path.join(FAVORITES_OUTPUT_DIR, f"{formatted_symbol}_{name}_日足.csv")
        
        # ファイルが存在しない場合は元ファイルをコピー
        if not os.path.exists(target_file):
            shutil.copy2(base_source_file, target_file)
            print(f"ファイルをコピーしました: {target_file}")
        
        # 既存のデータを読み込む（日付カラムをインデックスとしてtz-naiveで読み込む）
        # 最初のカラム（インデックスとして保存されている日付）をパースしてインデックスに設定
        df_existing = pd.read_csv(target_file, index_col=0, parse_dates=True)
        
        if df_existing.empty:
             print(f"警告: 既存のデータファイル {target_file} が空です。スキップします。")
             return False # 既存データが空の場合は処理しない
             
        # 既存データの最終日付を取得
        last_date_in_csv = df_existing.index.max()
        
        # ダウンロード開始日を計算（CSV最終日の翌日）
        start_date_for_download = last_date_in_csv + timedelta(days=1)
        
        # 今日の日付（タイムゾーンなし比較用）
        today_date_naive = datetime.now().date()
        
        # ダウンロード開始日が今日の日付より後の場合は更新不要
        if start_date_for_download.date() > today_date_naive:
             print(f"更新データはありません。既存データは最新です: {symbol} ({name})")
             return True # 更新データがないが、エラーではないので成功とする
             
        print(f"ダウンロード開始日: {start_date_for_download.strftime('%Y-%m-%d')}")
        
        # Yahoo Financeから不足期間のデータを取得
        try:
            ticker = yf.Ticker(symbol)
            # 開始日と終了日（今日）を指定
            hist = ticker.history(start=start_date_for_download, end=datetime.now() + timedelta(days=1)) # 終了日は含めないため、今日+1日を指定
            
            if hist.empty:
                print(f"警告: {symbol} ({name}) の指定期間のデータが取得できませんでした。")
                # データが取得できない場合でも、市場休業日の処理は行う可能性がある
                pass # 処理を続行して休業日を埋める
            else:
                # 取得したデータにタイムゾーン情報がある場合は削除してtz-naiveに統一
                if hist.index.tz is not None:
                    hist.index = hist.index.tz_localize(None)
            
        except Exception as e:
            print(f"Yahoo Financeからのデータ取得中にエラーが発生しました ({symbol}): {e}")
            import traceback
            print(traceback.format_exc())
            # データ取得に失敗した場合は、休業日のデータ埋め処理のみ試みる
            hist = pd.DataFrame() # 空のデータフレームを作成して後続処理に進む
        
        # 既存データと新しいデータを結合
        # 取得したhistデータが空でない場合のみ結合を試みる
        if not hist.empty:
            combined_data_temp = pd.concat([df_existing, hist])
            # 重複を削除（新しいデータを優先）
            combined_data_temp = combined_data_temp[~combined_data_temp.index.duplicated(keep='last')]
        else:
            # histが空の場合は既存データのみを一時データとする
            combined_data_temp = df_existing.copy()
            
        # インデックス（日付）でソート
        combined_data_temp = combined_data_temp.sort_index()
        
        # 結合したデータ全体に市場休業日を埋める処理を適用
        # 処理範囲は、結合データの開始日〜終了日となる
        combined_data = process_data_with_holidays(combined_data_temp, "日足")
        
        if combined_data.empty:
            print(f"警告: 処理された結合データがありません: {symbol} ({name})")
            return True # データがない場合もエラーではない
            
        # CSVファイルに保存（インデックスも保存）
        # process_data_with_holidaysによってIs_Market_Holidayカラムが追加される
        combined_data.to_csv(target_file)
        print(f"データを更新しました: {symbol} ({name}) - {start_date_for_download.strftime('%Y-%m-%d')} から {combined_data.index.max().strftime('%Y-%m-%d')}")
        
        return True
            
    except Exception as e:
        print(f"ファイル更新中にエラーが発生しました ({symbol}): {e}")
        import traceback
        print(traceback.format_exc()) # エラーの詳細を出力
        return False

def main():
    print("\n===== お気に入り銘柄不足データ補填ツール =====")
    
    # 設定を読み込む
    config = load_config()
    if config is None:
        print("設定ファイルの読み込みに失敗しました。")
        return
    
    # お気に入り銘柄リストを取得
    favorites = config.get("favorites", [])
    if not favorites:
        print("お気に入り銘柄が設定されていません。")
        return
    
    print(f"更新対象銘柄数: {len(favorites)}")
    print(f"出力ディレクトリ: {FAVORITES_OUTPUT_DIR}")
    
    # 処理結果のカウンター
    success_count = 0
    error_count = 0
    
    # 各お気に入り銘柄について処理
    for item in favorites:
        symbol = item["symbol"]
        name = item["name"]
        
        print(f"\n処理中: {symbol} ({name})")
        
        try:
            if update_csv_file(symbol, name):
                success_count += 1
            else:
                error_count += 1
            
            # APIレート制限を考慮して少し待機
            time.sleep(1) # 連続アクセスを防ぐために1秒待機
            
        except Exception as e:
            print(f"エラーが発生しました: {e}")
            error_count += 1
    
    # 処理結果の表示
    print("\n===== 処理結果 =====")
    print(f"総処理件数: {len(favorites)}")
    print(f"成功: {success_count}")
    print(f"エラー: {error_count}")
    
    print("\n処理が完了しました。")

if __name__ == "__main__":
    main() 