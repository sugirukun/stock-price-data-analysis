import json
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
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
        return True
    except Exception as e:
        print(f"設定ファイルの保存中にエラーが発生しました: {e}")
        messagebox.showerror("エラー", f"設定ファイルの保存中にエラーが発生しました:\n{e}\n\n保存先: {CONFIG_FILE}")
        return False

class StockConfigApp:
    def __init__(self, root):
        self.root = root
        self.root.title("株価データ設定管理ツール")
        self.root.geometry("800x600")
        self.root.minsize(600, 400)
        
        # 設定の読み込み
        self.config = load_config()
        
        # メインフレーム
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # タブコントロールの作成
        self.tab_control = ttk.Notebook(main_frame)
        
        # タブの作成
        self.stocks_tab = ttk.Frame(self.tab_control)
        self.settings_tab = ttk.Frame(self.tab_control)
        
        self.tab_control.add(self.stocks_tab, text="銘柄管理")
        self.tab_control.add(self.settings_tab, text="一般設定")
        self.tab_control.pack(expand=True, fill=tk.BOTH)
        
        # 銘柄管理タブの設定
        self.setup_stocks_tab()
        
        # 一般設定タブの設定
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
    
    def setup_stocks_tab(self):
        # 左側：銘柄リスト
        list_frame = ttk.LabelFrame(self.stocks_tab, text="銘柄リスト", padding="10")
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # リストビューの作成
        columns = ("シンボル", "銘柄名")
        self.stock_list = ttk.Treeview(list_frame, columns=columns, show="headings")
        
        # ヘッダーの設定
        for col in columns:
            self.stock_list.heading(col, text=col)
            
        # カラム幅の設定
        self.stock_list.column("シンボル", width=100)
        self.stock_list.column("銘柄名", width=250)
        
        # リストビューをグリッドに配置
        self.stock_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # スクロールバーの追加
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.stock_list.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.stock_list.configure(yscrollcommand=scrollbar.set)
        
        # データの表示
        self.update_stock_list()
        
        # 右側：操作フレーム
        operation_frame = ttk.LabelFrame(self.stocks_tab, text="銘柄操作", padding="10")
        operation_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=5, pady=5)
        
        # 銘柄追加フレーム
        add_frame = ttk.Frame(operation_frame)
        add_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(add_frame, text="ティッカーシンボル:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.symbol_entry = ttk.Entry(add_frame)
        self.symbol_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(add_frame, text="銘柄名（オプション）:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.name_entry = ttk.Entry(add_frame)
        self.name_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # チェックボックスオプション
        self.auto_name_var = tk.BooleanVar(value=True)
        auto_name_check = ttk.Checkbutton(add_frame, text="銘柄名を自動取得", variable=self.auto_name_var)
        auto_name_check.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # ボタンフレーム
        button_frame = ttk.Frame(operation_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        add_button = ttk.Button(button_frame, text="銘柄を追加", command=self.add_stock)
        add_button.pack(fill=tk.X, pady=5)
        
        delete_button = ttk.Button(button_frame, text="選択した銘柄を削除", command=self.delete_stock)
        delete_button.pack(fill=tk.X, pady=5)
        
        refresh_button = ttk.Button(button_frame, text="銘柄名を再取得", command=self.refresh_stock_names)
        refresh_button.pack(fill=tk.X, pady=5)
        
        # ヘルプのテキスト
        help_text = "※ティッカーシンボルの例:\n・日経平均: ^N225\n・NYダウ: ^DJI\n・日本株: [証券コード].T (例: 7974.T)\n・米国株: [シンボル] (例: AAPL)"
        help_label = ttk.Label(operation_frame, text=help_text, wraplength=200, justify=tk.LEFT)
        help_label.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
    
    def setup_settings_tab(self):
        settings_frame = ttk.Frame(self.settings_tab, padding="10")
        settings_frame.pack(fill=tk.BOTH, expand=True)
        
        # カラム名の設定
        columns_frame = ttk.LabelFrame(settings_frame, text="カラム表示設定", padding="10")
        columns_frame.pack(fill=tk.X, pady=10)
        
        self.column_var = tk.BooleanVar(value=self.config.get('use_japanese_columns', False))
        columns_check = ttk.Checkbutton(columns_frame, text="カラム名を日本語で表示", variable=self.column_var)
        columns_check.pack(fill=tk.X, pady=5)
        
        # 出力ディレクトリの設定
        dir_frame = ttk.LabelFrame(settings_frame, text="出力ディレクトリ設定", padding="10")
        dir_frame.pack(fill=tk.X, pady=10)
        
        dir_selection_frame = ttk.Frame(dir_frame)
        dir_selection_frame.pack(fill=tk.X)
        
        self.dir_var = tk.StringVar(value=self.config.get('output_dir', ""))
        dir_entry = ttk.Entry(dir_selection_frame, textvariable=self.dir_var, width=50)
        dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        
        dir_button = ttk.Button(dir_selection_frame, text="参照...", command=self.select_directory)
        dir_button.pack(side=tk.RIGHT, padx=5, pady=5)
        
        # 設定ファイルの情報
        config_frame = ttk.LabelFrame(settings_frame, text="設定ファイル情報", padding="10")
        config_frame.pack(fill=tk.X, pady=10)
        
        config_info = ttk.Label(config_frame, text=f"設定ファイルの保存先:\n{CONFIG_FILE}", wraplength=700)
        config_info.pack(fill=tk.X, pady=5)
        
        open_folder_button = ttk.Button(config_frame, text="設定ファイルのフォルダを開く", 
                                      command=self.open_config_folder)
        open_folder_button.pack(pady=5)
    
    def open_config_folder(self):
        """設定ファイルのあるフォルダをエクスプローラーで開く"""
        folder_path = os.path.dirname(CONFIG_FILE)
        try:
            os.startfile(folder_path)
        except Exception as e:
            messagebox.showerror("エラー", f"フォルダを開けませんでした: {e}")
    
    def update_stock_list(self):
        # 既存の項目をクリア
        for item in self.stock_list.get_children():
            self.stock_list.delete(item)
        
        # 新しい項目を追加
        for ticker in self.config['tickers']:
            self.stock_list.insert("", tk.END, values=(ticker['symbol'], ticker['name']))
    
    def add_stock(self):
        symbol = self.symbol_entry.get().strip()
        name = self.name_entry.get().strip()
        
        if not symbol:
            messagebox.showerror("エラー", "ティッカーシンボルを入力してください。")
            return
        
        # 既に存在するかチェック
        existing = False
        for ticker in self.config['tickers']:
            if ticker['symbol'] == symbol:
                existing = True
                break
        
        if existing:
            response = messagebox.askyesno("確認", 
                f"ティッカーシンボル '{symbol}' はすでに存在します。更新しますか？")
            if not response:
                return
        
        # 銘柄名の自動取得が選択されている場合
        if self.auto_name_var.get() and not name:
            self.status_var.set(f"銘柄名を取得中: {symbol}...")
            self.root.update()
            
            # 銘柄名の取得処理を別スレッドで実行
            def fetch_name():
                try:
                    stock_name = get_stock_name(symbol)
                    # UIスレッドで処理を続行
                    self.root.after(0, lambda: self.complete_add_stock(symbol, stock_name, existing))
                except Exception as e:
                    self.root.after(0, lambda: self.show_fetch_error(symbol, str(e)))
            
            threading.Thread(target=fetch_name).start()
        else:
            # 手動入力の名前を使用
            if not name:
                name = symbol  # 名前が空の場合はシンボルを使用
            self.complete_add_stock(symbol, name, existing)
    
    def complete_add_stock(self, symbol, name, existing):
        self.status_var.set(f"設定ファイル保存先: {CONFIG_FILE}")
        
        # 名前が空でないか確認
        if not name:
            name = symbol
        
        if existing:
            # 既存のエントリを更新
            for ticker in self.config['tickers']:
                if ticker['symbol'] == symbol:
                    ticker['name'] = name
                    break
            messagebox.showinfo("成功", f"銘柄情報を更新しました: {symbol} ({name})")
        else:
            # 新しいエントリを追加
            self.config['tickers'].append({"symbol": symbol, "name": name})
            messagebox.showinfo("成功", f"銘柄を追加しました: {symbol} ({name})")
        
        # リストの更新
        self.update_stock_list()
        
        # 入力フィールドをクリア
        self.symbol_entry.delete(0, tk.END)
        self.name_entry.delete(0, tk.END)
    
    def show_fetch_error(self, symbol, error_message):
        self.status_var.set(f"設定ファイル保存先: {CONFIG_FILE}")
        messagebox.showerror("エラー", f"{symbol} の銘柄名を取得できませんでした。\n{error_message}")
        
        # 手動入力用のダイアログを表示
        response = messagebox.askyesno("確認", "銘柄名を手動で入力しますか？")
        if response:
            # 手入力のダイアログ
            manual_name = self.name_entry.get().strip()
            if not manual_name:
                dialog = tk.Toplevel(self.root)
                dialog.title("銘柄名入力")
                dialog.geometry("300x150")
                dialog.transient(self.root)
                dialog.grab_set()
                
                ttk.Label(dialog, text=f"「{symbol}」の銘柄名を入力してください:").pack(pady=10)
                
                name_var = tk.StringVar()
                name_entry = ttk.Entry(dialog, textvariable=name_var, width=40)
                name_entry.pack(fill=tk.X, padx=20, pady=10)
                name_entry.focus_set()
                
                def on_ok():
                    manual_name = name_var.get().strip()
                    dialog.destroy()
                    self.complete_add_stock(symbol, manual_name if manual_name else symbol, False)
                
                def on_cancel():
                    dialog.destroy()
                
                button_frame = ttk.Frame(dialog)
                button_frame.pack(fill=tk.X, pady=10)
                
                ok_button = ttk.Button(button_frame, text="OK", command=on_ok)
                ok_button.pack(side=tk.RIGHT, padx=5)
                
                cancel_button = ttk.Button(button_frame, text="キャンセル", command=on_cancel)
                cancel_button.pack(side=tk.RIGHT, padx=5)
                
                # Enterキーでも確定できるように
                dialog.bind("<Return>", lambda event: on_ok())
                
                # ダイアログが閉じるまで待機
                self.root.wait_window(dialog)
            else:
                self.complete_add_stock(symbol, manual_name, False)
    
    def delete_stock(self):
        selected_items = self.stock_list.selection()
        
        if not selected_items:
            messagebox.showinfo("情報", "削除する銘柄を選択してください。")
            return
        
        confirmation = messagebox.askyesno("確認", "選択した銘柄を削除しますか？")
        if confirmation:
            # 削除する項目の取得
            for item in selected_items:
                symbol = self.stock_list.item(item, "values")[0]
                # 設定から削除
                self.config['tickers'] = [t for t in self.config['tickers'] if t['symbol'] != symbol]
            
            # リストの更新
            self.update_stock_list()
            messagebox.showinfo("成功", "選択した銘柄を削除しました。")
    
    def refresh_stock_names(self):
        selected_items = self.stock_list.selection()
        
        if not selected_items:
            messagebox.showinfo("情報", "銘柄名を再取得する銘柄を選択してください。")
            return
        
        confirmation = messagebox.askyesno("確認", "選択した銘柄の名前を再取得しますか？")
        if confirmation:
            self.status_var.set("銘柄名を再取得中...")
            self.root.update()
            
            def refresh_names():
                # 選択した項目の銘柄名を再取得
                for item in selected_items:
                    symbol = self.stock_list.item(item, "values")[0]
                    
                    try:
                        # 銘柄名の再取得
                        name = get_stock_name(symbol)
                        
                        # 設定を更新
                        for ticker in self.config['tickers']:
                            if ticker['symbol'] == symbol:
                                ticker['name'] = name
                                break
                    except Exception as e:
                        self.root.after(0, lambda e=e, s=symbol: messagebox.showerror(
                            "エラー", f"{s} の銘柄名の再取得中にエラーが発生しました: {e}"))
                
                # UIスレッドで表示を更新
                self.root.after(0, self.complete_refresh)
            
            threading.Thread(target=refresh_names).start()
    
    def complete_refresh(self):
        self.update_stock_list()
        messagebox.showinfo("成功", "銘柄名の再取得が完了しました。")
        self.status_var.set(f"設定ファイル保存先: {CONFIG_FILE}")
    
    def select_directory(self):
        directory = filedialog.askdirectory(
            initialdir=self.dir_var.get(),
            title="出力ディレクトリの選択"
        )
        
        if directory:
            self.dir_var.set(directory)
    
    def save_settings(self):
        # 設定の更新
        self.config['use_japanese_columns'] = self.column_var.get()
        self.config['output_dir'] = self.dir_var.get()
        
        # 設定の保存
        if save_config(self.config):
            messagebox.showinfo("成功", "設定を保存しました。")
        else:
            pass  # エラーメッセージはsave_config内で表示される

if __name__ == "__main__":
    root = tk.Tk()
    app = StockConfigApp(root)
    root.mainloop()
