import json
import os
import yfinance as yf
from time import sleep

# 設定ファイルのパス
CONFIG_FILE = "stock_config.json"

def get_stock_name(ticker_symbol):
    """ティッカーシンボルから銘柄名を自動取得する"""
    try:
        # Yahoo Finance APIを使用して銘柄名を取得
        ticker = yf.Ticker(ticker_symbol)
        # APIリクエストを連続して送ると制限がかかることがあるので少し待機
        sleep(0.5)
        info = ticker.info
        
        # 日本の銘柄ならlongNameを優先し、なければshortNameを使用
        if 'longName' in info:
            return info['longName']
        elif 'shortName' in info:
            return info['shortName']
        else:
            # 名前が取得できない場合はシンボルをそのまま返す
            return ticker_symbol
    except Exception as e:
        print(f"警告: {ticker_symbol} の銘柄名を取得できませんでした: {e}")
        return ticker_symbol  # エラーの場合はシンボルをそのまま返す

# デフォルトの設定
DEFAULT_CONFIG = {
    "tickers": [
        {"symbol": "7974.T", "name": "任天堂"},
        {"symbol": "1387.T", "name": "eMAXIS Slim米国株式(S&P500)"},
        {"symbol": "AAPL", "name": "アップル"},
        {"symbol": "^N225", "name": "日経平均"},
        {"symbol": "^DJI", "name": "NYダウ"}
    ],
    "use_japanese_columns": False,
    "output_dir": "C:\\Users\\rilak\\Desktop\\株価\\株価データ"
}

def load_config():
    """設定ファイルを読み込む、存在しない場合はデフォルト設定を作成して返す"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                print(f"設定ファイル '{CONFIG_FILE}' を読み込みました。")
                return config
        except Exception as e:
            print(f"設定ファイルの読み込み中にエラーが発生しました: {e}")
            print("デフォルト設定を使用します。")
    else:
        print(f"設定ファイル '{CONFIG_FILE}' が見つかりません。デフォルト設定を作成します。")
        save_config(DEFAULT_CONFIG)
    
    return DEFAULT_CONFIG

def save_config(config):
    """設定ファイルを保存する"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
        print(f"設定ファイル '{CONFIG_FILE}' を保存しました。")
    except Exception as e:
        print(f"設定ファイルの保存中にエラーが発生しました: {e}")

def edit_config():
    """対話形式で設定を編集する"""
    config = load_config()
    
    while True:
        print("\n===== 設定管理 =====")
        print("1. 現在の銘柄リストを表示")
        print("2. 銘柄を追加")
        print("3. 銘柄を削除")
        print("4. カラム名の表示設定変更 (日本語/英語)")
        print("5. 出力ディレクトリの変更")
        print("6. 設定を保存して終了")
        print("0. 変更を破棄して終了")
        
        choice = input("\n選択してください (0-6): ").strip()
        
        if choice == '1':
            print("\n===== 現在の銘柄リスト =====")
            for i, ticker in enumerate(config['tickers'], 1):
                print(f"{i}. {ticker['symbol']} ({ticker['name']})")
        
        elif choice == '2':
            symbol = input("\n追加するティッカーシンボルを入力してください: ").strip()
            
            if symbol:
                try:
                    print(f"銘柄名を自動取得中... {symbol}")
                    name = get_stock_name(symbol)
                    print(f"取得した銘柄名: {name}")
                    
                    # 必要なら手動での編集も可能にする
                    edit_name = input("銘柄名を編集しますか？ (デフォルトを使用する場合は空白のままエンター): ").strip()
                    if edit_name:
                        name = edit_name
                except Exception as e:
                    print(f"エラーが発生しました: {e}")
                    name = input("銘柄名を手動で入力してください: ").strip()
                    if not name:
                        print("銘柄名が入力されていません。操作をスキップします。")
                        continue
            else:
                print("ティッカーシンボルが入力されていません。操作をスキップします。")
                continue
                
            # ここでnameが定義されているか確認
            if 'name' not in locals() or not name:
                print("銘柄名が取得できませんでした。操作をスキップします。")
                continue
            
            # すでに存在するかチェック
            exists = any(t['symbol'] == symbol for t in config['tickers'])
            if exists:
                print(f"警告: ティッカーシンボル '{symbol}' はすでに存在します。")
                update = input("更新しますか？ (y/n): ").strip().lower() == 'y'
                if update:
                    # 既存のエントリを更新
                    for t in config['tickers']:
                        if t['symbol'] == symbol:
                            t['name'] = name
                            break
                    print(f"銘柄情報を更新しました: {symbol} ({name})")
            else:
                config['tickers'].append({"symbol": symbol, "name": name})
                print(f"銘柄を追加しました: {symbol} ({name})")
        
        elif choice == '3':
            print("\n===== 銘柄を削除 =====")
            for i, ticker in enumerate(config['tickers'], 1):
                print(f"{i}. {ticker['symbol']} ({ticker['name']})")
            
            try:
                idx = int(input("\n削除する銘柄の番号を入力してください: ").strip()) - 1
                if 0 <= idx < len(config['tickers']):
                    removed = config['tickers'].pop(idx)
                    print(f"銘柄を削除しました: {removed['symbol']} ({removed['name']})")
                else:
                    print("無効な番号です。")
            except ValueError:
                print("有効な番号を入力してください。")
        
        elif choice == '4':
            current = "日本語" if config.get('use_japanese_columns', False) else "英語"
            print(f"\n現在のカラム名表示設定: {current}")
            
            change = input("カラム名を日本語で表示しますか？ (y/n): ").strip().lower() == 'y'
            config['use_japanese_columns'] = change
            
            new_setting = "日本語" if change else "英語"
            print(f"カラム名表示設定を '{new_setting}' に変更しました。")
        
        elif choice == '5':
            current_dir = config.get('output_dir', "")
            print(f"\n現在の出力ディレクトリ: {current_dir}")
            
            new_dir = input("新しい出力ディレクトリを入力してください: ").strip()
            if new_dir:
                # 絶対パスでない場合は相対パスとして処理
                if not os.path.isabs(new_dir):
                    new_dir = os.path.abspath(new_dir)
                
                # ディレクトリが存在するか確認
                if not os.path.exists(new_dir):
                    create = input(f"ディレクトリ '{new_dir}' は存在しません。作成しますか？ (y/n): ").strip().lower() == 'y'
                    if create:
                        try:
                            os.makedirs(new_dir, exist_ok=True)
                            print(f"ディレクトリを作成しました: {new_dir}")
                        except Exception as e:
                            print(f"ディレクトリの作成中にエラーが発生しました: {e}")
                            continue
                    else:
                        print("操作をキャンセルしました。")
                        continue
                
                config['output_dir'] = new_dir
                print(f"出力ディレクトリを変更しました: {new_dir}")
            else:
                print("入力が空です。操作をスキップします。")
        
        elif choice == '6':
            save_config(config)
            print("設定を保存して終了します。")
            return config
        
        elif choice == '0':
            print("変更を破棄して終了します。")
            return load_config()  # 元の設定を読み込み直して返す
        
        else:
            print("無効な選択です。再度試してください。")

if __name__ == "__main__":
    print("===== 株価データ設定管理ツール =====")
    config = edit_config()
    print("\n最終設定:")
    print(json.dumps(config, ensure_ascii=False, indent=2))
    print("\nこの設定ファイルは 'stock_config.json' に保存されています。")
    print("株価データ取得スクリプトはこの設定ファイルを使用します。")
    
    input("\nEnterキーを押して終了...")
