import json
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import yfinance as yf
from time import sleep
import threading

# 設定ファイルのパス
CONFIG_DIR = "C:\\Users\\rilak\\Desktop\\株価"
CONFIG_FILE = os.path.join(CONFIG_DIR, "stock_config.json")

# デフォルトの設定
DEFAULT_CONFIG = {
    "tickers": [
        {"symbol": "7974.T", "name": "任天堂"},
        {"symbol": "1387.T", "name": "eMAXIS Slim米国株式(S&P500)"},
        {"symbol": "AAPL", "name": "アップル"},
        {"symbol": "^N225", "name": "日経平均"},
        {"symbol": "^DJI", "name": "NYダウ"}
    ],
    "favorites": [],
    "portfolios": [
        {"name": "ポートフォリオ1", "tickers": []},
        {"name": "ポートフォリオ2", "tickers": []},
        {"name": "ポートフォリオ3", "tickers": []},
    ],
    "use_japanese_columns": False,
    "output_dir": "C:\\Users\\rilak\\Desktop\\株価\\株価データ"
}

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

def load_config():
    """設定ファイルを読み込む、存在しない場合はデフォルト設定を作成して返す"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                print(f"設定ファイル '{CONFIG_FILE}' を読み込みました。")
                
                # 旧バージョンの設定ファイルの場合、新しいフィールドを追加
                if 'favorites' not in config:
                    config['favorites'] = []
                if 'portfolios' not in config:
                    config['portfolios'] = [
                        {"name": "ポートフォリオ1", "tickers": []},
                        {"name": "ポートフォリオ2", "tickers": []},
                        {"name": "ポートフォリオ3", "tickers": []}
                    ]
                
                return config
        except Exception as e:
            print(f"設定ファイルの読み込み中にエラーが発生しました: {e}")
            print("デフォルト設定を使用します。")
    else:
        print(f"設定ファイル '{CONFIG_FILE}' が見つかりません。デフォルト設定を作成します。")
        # 設定ディレクトリが存在しない場合は作成
        if not os.path.exists(CONFIG_DIR):
            os.makedirs(CONFIG_DIR)
        save_config(DEFAULT_CONFIG)
    
    return DEFAULT_CONFIG

def save_config(config):
    """設定ファイルを保存する"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
        print(f"設定ファイル '{CONFIG_FILE}' を保存しました。")
        return True
    except Exception as e:
        print(f"設定ファイルの保存中にエラーが発生しました: {e}")
        messagebox.showerror("エラー", f"設定ファイルの保存中にエラーが発生しました:\n{e}\n\n保存先: {CONFIG_FILE}")
        return False

class StockConfigApp:
    def __init__(self, root):
        self.root = root
        self.root.title("株価データ設定管理ツール")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # 設定の読み込み
        self.config = load_config()
        
        # メインフレーム
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # タブコントロールの作成
        self.tab_control = ttk.Notebook(main_frame)
        
        # タブの作成
        self.stocks_tab = ttk.Frame(self.tab_control)
        self.favorites_tab = ttk.Frame(self.tab_control)
        self.portfolio_tab = ttk.Frame(self.tab_control)
        self.settings_tab = ttk.Frame(self.tab_control)
        
        self.tab_control.add(self.stocks_tab, text="銘柄管理")
        self.tab_control.add(self.favorites_tab, text="お気に入り")
        self.tab_control.add(self.portfolio_tab, text="ポートフォリオ")
        self.tab_control.add(self.settings_tab, text="一般設定")
        self.tab_control.pack(expand=True, fill=tk.BOTH)
        
        # 各タブの設定
        self.setup_stocks_tab()
        self.setup_favorites_tab()
        self.setup_portfolio_tab()
        self.setup_settings_tab()
        
        # 下部のボタン群
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        save_button = ttk.Button(button_frame, text="設定を保存", command=self.save_settings)
        save_button.pack(side=tk.RIGHT, padx=5)
        
        # ステータスバー
        self.status_var = tk.StringVar()
        self.status_var.set(f"設定ファイル保存先: {CONFIG_FILE}")
        status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 現在のターゲットリスト（複数選択時のコピー先を特定するため）
        self.current_target_list = None
        self.current_portfolio_index = 0
    
    # ここにsetup_stocks_tab, setup_favorites_tab, setup_portfolio_tab, setup_settings_tabメソッドを追加する
    
    def delete_stock(self, tree_view, list_type):
        """選択された銘柄を削除する"""
        selected_items = tree_view.selection()
        if not selected_items:
            messagebox.showinfo("情報", "削除する銘柄を選択してください。")
            return
        
        # 確認ダイアログを表示
        count = len(selected_items)
        if count == 1:
            symbol = tree_view.item(selected_items[0], "values")[0]
            name = tree_view.item(selected_items[0], "values")[1]
            message = f"以下の銘柄を削除しますか？\n\n{symbol} - {name}"
        else:
            message = f"選択された {count} 件の銘柄を削除しますか？"
        
        response = messagebox.askyesno("確認", message)
        if not response:
            return
        
        # 削除処理
        symbols_to_delete = [tree_view.item(item, "values")[0] for item in selected_items]
        
        # リストタイプに応じたデータを更新
        if list_type == 'tickers':
            self.config['tickers'] = [item for item in self.config['tickers'] 
                                    if item['symbol'] not in symbols_to_delete]
        elif list_type == 'favorites':
            self.config['favorites'] = [item for item in self.config['favorites'] 
                                      if item['symbol'] not in symbols_to_delete]
        
        # リストを更新
        self.update_stock_list(tree_view, list_type)
        
        messagebox.showinfo("成功", f"{count}件の銘柄を削除しました。")

    def delete_from_portfolio(self):
        """ポートフォリオから選択された銘柄を削除する"""
        selected_items = self.portfolio_list.selection()
        if not selected_items:
            messagebox.showinfo("情報", "削除する銘柄を選択してください。")
            return
        
        # 確認ダイアログを表示
        count = len(selected_items)
        if count == 1:
            symbol = self.portfolio_list.item(selected_items[0], "values")[0]
            name = self.portfolio_list.item(selected_items[0], "values")[1]
            message = f"以下の銘柄をポートフォリオから削除しますか？\n\n{symbol} - {name}"
        else:
            message = f"選択された {count} 件の銘柄をポートフォリオから削除しますか？"
        
        response = messagebox.askyesno("確認", message)
        if not response:
            return
        
        # 削除処理
        symbols_to_delete = [self.portfolio_list.item(item, "values")[0] for item in selected_items]
        
        # 現在のポートフォリオから指定された銘柄を削除
        current_portfolio = self.config['portfolios'][self.current_portfolio_index]
        current_portfolio['tickers'] = [item for item in current_portfolio['tickers'] 
                                      if item['symbol'] not in symbols_to_delete]
        
        # リストを更新
        self.update_portfolio_list(self.current_portfolio_index)
        
        messagebox.showinfo("成功", f"{count}件の銘柄をポートフォリオから削除しました。")

    def refresh_stock_names(self, tree_view, list_type):
        """選択された銘柄の名前を再取得する"""
        selected_items = tree_view.selection()
        if not selected_items:
            messagebox.showinfo("情報", "名前を再取得する銘柄を選択してください。")
            return
        
        # 確認ダイアログを表示
        count = len(selected_items)
        if count == 1:
            symbol = tree_view.item(selected_items[0], "values")[0]
            message = f"銘柄 {symbol} の名前を再取得しますか？"
        else:
            message = f"選択された {count} 件の銘柄の名前を再取得しますか？"
        
        response = messagebox.askyesno("確認", message)
        if not response:
            return
        
        # 進捗状況を表示するウィンドウ
        progress_window = tk.Toplevel(self.root)
        progress_window.title("処理中")
        progress_window.geometry("300x100")
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        progress_var = tk.StringVar(value="銘柄名を取得中...")
        progress_label = ttk.Label(progress_window, textvariable=progress_var)
        progress_label.pack(pady=10)
        
        progress_bar = ttk.Progressbar(progress_window, mode="indeterminate")
        progress_bar.pack(fill=tk.X, padx=20, pady=10)
        progress_bar.start()
        
        # 銘柄情報の辞書を作成
        symbols_to_update = [tree_view.item(item, "values")[0] for item in selected_items]
        
        # バックグラウンドスレッドで処理を実行
        def update_names():
            updated_count = 0
            
            for i, symbol in enumerate(symbols_to_update):
                # 進捗状況を更新
                progress_var.set(f"銘柄名を取得中... {i+1}/{len(symbols_to_update)}")
                progress_window.update_idletasks()
                
                # 銘柄名を取得
                new_name = get_stock_name(symbol)
                
                # リストタイプに応じたデータを更新
                if list_type == 'tickers':
                    target_list = self.config['tickers']
                elif list_type == 'favorites':
                    target_list = self.config['favorites']
                else:
                    continue
                
                # 対象の銘柄を検索して名前を更新
                for item in target_list:
                    if item['symbol'] == symbol:
                        if item['name'] != new_name:
                            item['name'] = new_name
                            updated_count += 1
                        break
            
            # 処理完了後、UIスレッドで後処理を実行
            self.root.after(0, lambda: finish_update(updated_count))
        
        def finish_update(updated_count):
            # 進捗ウィンドウを閉じる
            progress_window.destroy()
            
            # リストを更新
            self.update_stock_list(tree_view, list_type)
            
            # 結果を表示
            messagebox.showinfo("完了", f"{len(symbols_to_update)}件の銘柄を処理しました。\n{updated_count}件の銘柄名が更新されました。")
        
        # バックグラウンドスレッドを開始
        threading.Thread(target=update_names, daemon=True).start()

    def refresh_portfolio_names(self):
        """ポートフォリオ内の選択された銘柄の名前を再取得する"""
        selected_items = self.portfolio_list.selection()
        if not selected_items:
            messagebox.showinfo("情報", "名前を再取得する銘柄を選択してください。")
            return
        
        # 確認ダイアログを表示
        count = len(selected_items)
        if count == 1:
            symbol = self.portfolio_list.item(selected_items[0], "values")[0]
            message = f"銘柄 {symbol} の名前を再取得しますか？"
        else:
            message = f"選択された {count} 件の銘柄の名前を再取得しますか？"
        
        response = messagebox.askyesno("確認", message)
        if not response:
            return
        
        # 進捗状況を表示するウィンドウ
        progress_window = tk.Toplevel(self.root)
        progress_window.title("処理中")
        progress_window.geometry("300x100")
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        progress_var = tk.StringVar(value="銘柄名を取得中...")
        progress_label = ttk.Label(progress_window, textvariable=progress_var)
        progress_label.pack(pady=10)
        
        progress_bar = ttk.Progressbar(progress_window, mode="indeterminate")
        progress_bar.pack(fill=tk.X, padx=20, pady=10)
        progress_bar.start()
        
        # 銘柄情報の辞書を作成
        symbols_to_update = [self.portfolio_list.item(item, "values")[0] for item in selected_items]
        
        # バックグラウンドスレッドで処理を実行
        def update_names():
            updated_count = 0
            
            for i, symbol in enumerate(symbols_to_update):
                # 進捗状況を更新
                progress_var.set(f"銘柄名を取得中... {i+1}/{len(symbols_to_update)}")
                progress_window.update_idletasks()
                
                # 銘柄名を取得
                new_name = get_stock_name(symbol)
                
                # 現在のポートフォリオから対象の銘柄を検索して名前を更新
                portfolio_tickers = self.config['portfolios'][self.current_portfolio_index]['tickers']
                for item in portfolio_tickers:
                    if item['symbol'] == symbol:
                        if item['name'] != new_name:
                            item['name'] = new_name
                            updated_count += 1
                        break
            
            # 処理完了後、UIスレッドで後処理を実行
            self.root.after(0, lambda: finish_update(updated_count))
        
        def finish_update(updated_count):
            # 進捗ウィンドウを閉じる
            progress_window.destroy()
            
            # リストを更新
            self.update_portfolio_list(self.current_portfolio_index)
            
            # 結果を表示
            messagebox.showinfo("完了", f"{len(symbols_to_update)}件の銘柄を処理しました。\n{updated_count}件の銘柄名が更新されました。")
        
        # バックグラウンドスレッドを開始
        threading.Thread(target=update_names, daemon=True).start()

    def copy_stocks(self, source_tree, target_list_type):
        """選択された銘柄を指定されたリストにコピーする"""
        selected_items = source_tree.selection()
        if not selected_items:
            messagebox.showinfo("情報", "コピーする銘柄を選択してください。")
            return
        
        # 確認ダイアログを表示
        count = len(selected_items)
        
        # コピー先のリスト名を取得
        list_name = ""
        if target_list_type == 'tickers':
            list_name = "銘柄管理"
        elif target_list_type == 'favorites':
            list_name = "お気に入り"
        
        if count == 1:
            symbol = source_tree.item(selected_items[0], "values")[0]
            name = source_tree.item(selected_items[0], "values")[1]
            message = f"以下の銘柄を{list_name}にコピーしますか？\n\n{symbol} - {name}"
        else:
            message = f"選択された {count} 件の銘柄を{list_name}にコピーしますか？"
        
        response = messagebox.askyesno("確認", message)
        if not response:
            return
        
        # コピー処理
        added_count = 0
        updated_count = 0
        
        for item in selected_items:
            values = source_tree.item(item, "values")
            symbol = values[0]
            name = values[1]
            
            # ターゲットリストを決定
            if target_list_type == 'tickers':
                target_list = self.config['tickers']
            elif target_list_type == 'favorites':
                target_list = self.config['favorites']
            else:
                continue
            
            # 既に存在するかチェック
            existing = False
            for ticker in target_list:
                if ticker['symbol'] == symbol:
                    existing = True
                    # 名前が異なる場合は更新
                    if ticker['name'] != name:
                        ticker['name'] = name
                        updated_count += 1
                    break
            
            # 存在しない場合は追加
            if not existing:
                target_list.append({"symbol": symbol, "name": name})
                added_count += 1
        
        # 結果メッセージ
        result_msg = f"{added_count}件の銘柄が追加されました。"
        if updated_count > 0:
            result_msg += f"\n{updated_count}件の銘柄名が更新されました。"
        
        messagebox.showinfo("完了", result_msg)
        
        # リストを更新
        if target_list_type == 'tickers':
            self.update_stock_list(self.stock_list, 'tickers')
        elif target_list_type == 'favorites':
            self.update_stock_list(self.favorites_list, 'favorites')

    def copy_to_portfolio(self, source_tree):
        """選択された銘柄を現在のポートフォリオにコピーする"""
        selected_items = source_tree.selection()
        if not selected_items:
            messagebox.showinfo("情報", "コピーする銘柄を選択してください。")
            return
        
        # 確認ダイアログを表示
        count = len(selected_items)
        portfolio_name = self.config['portfolios'][self.current_portfolio_index]['name']
        
        if count == 1:
            symbol = source_tree.item(selected_items[0], "values")[0]
            name = source_tree.item(selected_items[0], "values")[1]
            message = f"以下の銘柄をポートフォリオ「{portfolio_name}」にコピーしますか？\n\n{symbol} - {name}"
        else:
            message = f"選択された {count} 件の銘柄をポートフォリオ「{portfolio_name}」にコピーしますか？"
        
        response = messagebox.askyesno("確認", message)
        if not response:
            return
        
        # コピー処理
        added_count = 0
        updated_count = 0
        
        target_list = self.config['portfolios'][self.current_portfolio_index]['tickers']
        
        for item in selected_items:
            values = source_tree.item(item, "values")
            symbol = values[0]
            name = values[1]
            
            # 既に存在するかチェック
            existing = False
            for ticker in target_list:
                if ticker['symbol'] == symbol:
                    existing = True
                    # 名前が異なる場合は更新
                    if ticker['name'] != name:
                        ticker['name'] = name
                        updated_count += 1
                    break
            
            # 存在しない場合は追加
            if not existing:
                target_list.append({"symbol": symbol, "name": name})
                added_count += 1
        
        # 結果メッセージ
        result_msg = f"{added_count}件の銘柄が追加されました。"
        if updated_count > 0:
            result_msg += f"\n{updated_count}件の銘柄名が更新されました。"
        
        messagebox.showinfo("完了", result_msg)
        
        # リストを更新
        self.update_portfolio_list(self.current_portfolio_index)
    
    # その他のメソッド（update_stock_list, update_portfolio_list, save_settings など）
    # 必要なメソッドは実装する必要があります

# メインエントリーポイント
if __name__ == "__main__":
    root = tk.Tk()
    app = StockConfigApp(root)
    root.mainloop()
