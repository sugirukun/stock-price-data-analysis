import pandas as pd
import os
import glob
from openpyxl import load_workbook

# 出力ディレクトリ
data_dir = "C:\\Users\\rilak\\Desktop\\株価\\株価データ"
print(f"処理対象ディレクトリ: {data_dir}")

# 年足CSVファイルを検索
yearly_csv_files = glob.glob(os.path.join(data_dir, "*_年足.csv"))
print(f"見つかった年足CSVファイル: {len(yearly_csv_files)}個")

for yearly_csv_file in yearly_csv_files:
    try:
        # ファイル名から情報を抽出
        basename = os.path.basename(yearly_csv_file)
        ticker_and_name = basename.replace('_年足.csv', '')
        
        print(f"処理中: {basename} -> {ticker_and_name}")
        
        # 元のExcelファイルのパス
        excel_file = os.path.join(data_dir, f"{ticker_and_name}.xlsx")
        
        if not os.path.exists(excel_file):
            print(f"警告: 元のExcelファイルが見つかりません: {excel_file}")
            continue
        
        # 年足CSVデータを読み込む
        yearly_data = pd.read_csv(yearly_csv_file, index_col=0, parse_dates=True)
        
        # シート名の生成（銘柄名_年足）
        if "任天堂" in ticker_and_name:
            sheet_name = "任天堂_年足"
        elif "eMAXIS" in ticker_and_name:
            sheet_name = "eMAXIS_年足"
        elif "アップル" in ticker_and_name:
            sheet_name = "アップル_年足"
        else:
            sheet_name = f"{ticker_and_name.split('_')[-1][:15]}_年足"
        
        # シート名の長さ制限（31文字まで）
        sheet_name = sheet_name[:31]
        
        # シートの存在を確認し、存在する場合は上書き
        try:
            book = load_workbook(excel_file)
            if sheet_name in book.sheetnames:
                print(f"シート '{sheet_name}' はすでに存在します。削除します。")
                std = book[sheet_name]
                book.remove(std)
                book.save(excel_file)
            book.close()
        except Exception as e:
            print(f"Excelファイルの読み込み中にエラーが発生しました: {e}")
            continue
        
        try:
            # Excelファイルに年足データを追加
            with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a') as writer:
                yearly_data.to_excel(writer, sheet_name=sheet_name)
            print(f"年足データを元のExcelファイルに追加しました: {excel_file} (シート: {sheet_name})")
        except Exception as e:
            print(f"エラー: {e}")
    
    except Exception as e:
        print(f"処理中にエラーが発生しました: {e}")
    
    print("=" * 50)

print("\n処理が完了しました。年足データを元のExcelファイルに追加しました。")
