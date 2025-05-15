import os
import sys
import time
import subprocess
import traceback

def main():
    """
    指定されたスクリプトを順番に実行する総合スクリプト
    """
    print("===== 株価データ処理総合スクリプト =====")
    print("このスクリプトは以下のスクリプトを順番に実行します：")
    print("1. stock_data_all.py - 株価データの取得")
    print("2. create_quarterly_data.py - 四半期データの作成")
    print("3. create_yearly_data_fixed.py - 年足データの作成")
    print("4. add_quarterly_to_excel.py - 四半期データのExcelファイルへの追加")
    print("5. add_yearly_data_to_excel.py - 年足データのExcelファイルへの追加")
    print("=" * 50)
    
    # 実行するスクリプトのリスト
    scripts = [
        "stock_data_all.py",
        "create_quarterly_data.py",
        "create_yearly_data_fixed.py",
        "add_quarterly_to_excel.py",
        "add_yearly_data_to_excel.py"
    ]
    
    # カレントディレクトリを取得
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # スクリプトごとに実行
    for i, script in enumerate(scripts, 1):
        script_path = os.path.join(current_dir, script)
        
        if not os.path.exists(script_path):
            print(f"エラー: スクリプト {script_path} が見つかりません。")
            continue
        
        print(f"\n\n{'='*80}")
        print(f"[{i}/{len(scripts)}] {script} を実行中...")
        print(f"{'='*80}\n")
        
        try:
            # Pythonのサブプロセスでスクリプトを実行
            subprocess.run([sys.executable, script_path], check=True)
            print(f"\n{script} の実行が完了しました。")
        except subprocess.CalledProcessError as e:
            print(f"\nエラー: {script} の実行中にエラーが発生しました。")
            print(f"終了コード: {e.returncode}")
            print("\n処理を続行します...\n")
        except Exception as e:
            print(f"\n予期せぬエラーが発生しました: {e}")
            traceback.print_exc()
            print("\n処理を続行します...\n")
        
        # 次のスクリプト実行前に少し待機
        if i < len(scripts):
            print("\n次のスクリプト実行まで5秒待機しています...\n")
            time.sleep(5)
    
    print("\n" + "="*50)
    print("全てのスクリプトの実行が完了しました。")
    print("=" * 50)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nユーザーによって処理が中断されました。")
    except Exception as e:
        print(f"\n\n予期せぬエラーが発生しました: {e}")
        traceback.print_exc()
    
    print("\n処理を終了します。何かキーを押すと終了します...")
    input()
