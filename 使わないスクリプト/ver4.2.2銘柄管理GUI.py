import json
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import yfinance as yf
from time import sleep
import threading
from PIL import Image, ImageTk  # 画像処理用にPILをインポート

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
        
        # 上部の画像エリアとボタン
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 画像の読み込みと表示（左側）
        try:
            # 画像ファイルのパス
            image_path = "C:\\Users\\rilak\\Desktop\\株価\\image\\コイン.jpg"
            
            # PILで画像を読み込み、サイズを調整
            original_image = Image.open(image_path)
            # サイズを2倍に調整
            resized_image = original_image.resize((160, 100), Image.Resampling.LANCZOS)
            
            # Tkinter用のフォーマットに変換
            self.character_image = ImageTk.PhotoImage(resized_image)
            
            # 画像表示用のラベル
            image_label = ttk.Label(top_frame, image=self.character_image)
            image_label.pack(side=tk.LEFT, padx=5)
            
            # 画像の横にキャッチフレーズなどを表示
            catch_phrase = ttk.Label(top_frame, 
                                    text="アクティブファンドの成績はインデックス以下で信託報酬が高いだけ！\n「全部同じじゃないですか」", 
                                    font=("Helvetica", 10), wraplength=600, justify=tk.LEFT)
            catch_phrase.pack(side=tk.LEFT, padx=5)
            
        except Exception as e:
            print(f"画像の読み込みに失敗しました: {e}")
        
        # 保存ボタン（右側）
        save_button = ttk.Button(top_frame, text="設定を保存", command=self.save_settings)
        save_button.pack(side=tk.RIGHT, padx=10)
        
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
        
        # ステータスバー
        self.status_var = tk.StringVar()
        self.status_var.set(f"設定ファイル保存先: {CONFIG_FILE}")
        status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 現在のターゲットリスト（複数選択時のコピー先を特定するため）
        self.current_target_list = None
        self.current_portfolio_index = 0
        
        # ドラッグ＆ドロップ用の変数
        self.drag_item = None
        self.drag_data = {"item": None, "x": 0, "y": 0, "tree": None}
    
    def setup_stocks_tab(self):
        """銘柄管理タブのUIを設定する"""
        # 左側：銘柄リスト
        list_frame = ttk.LabelFrame(self.stocks_tab, text="銘柄リスト", padding="10")
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # リストビューの作成
        columns = ("シンボル", "銘柄名")
        self.stock_list = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="extended")
        
        # ヘッダーの設定
        for col in columns:
            self.stock_list.heading(col, text=col, command=lambda c=col: self.sort_treeview(self.stock_list, c, False))
            
        # カラム幅の設定
        self.stock_list.column("シンボル", width=100)
        self.stock_list.column("銘柄名", width=250)
        
        # リストビューをグリッドに配置
        self.stock_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # スクロールバーの追加
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.stock_list.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.stock_list.configure(yscrollcommand=scrollbar.set)
        
        # ドラッグ＆ドロップの設定
        self.setup_drag_and_drop(self.stock_list)
        
        # 右側：操作フレーム
        operation_frame = ttk.LabelFrame(self.stocks_tab, text="銘柄操作", padding="10")
        operation_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5, pady=5)
        
        # 幅を小さくする
        operation_frame.configure(width=180)  # 幅を小さく設定
        
        # 銘柄追加フレーム
        add_frame = ttk.Frame(operation_frame)
        add_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(add_frame, text="ティッカーシンボル:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.symbol_entry = ttk.Entry(add_frame)
        self.symbol_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2)
        
        ttk.Label(add_frame, text="銘柄名（オプション）:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.name_entry = ttk.Entry(add_frame)
        self.name_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=2)
        
        # チェックボックスオプション
        self.auto_name_var = tk.BooleanVar(value=True)
        auto_name_check = ttk.Checkbutton(add_frame, text="銘柄名を自動取得", variable=self.auto_name_var)
        auto_name_check.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=2)
        
        # ボタンフレーム
        button_frame = ttk.Frame(operation_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        add_button = ttk.Button(button_frame, text="銘柄を追加", command=self.add_stock)
        add_button.pack(fill=tk.X, pady=2)
        
        delete_button = ttk.Button(button_frame, text="選択した銘柄を削除", command=lambda: self.delete_stock(self.stock_list, 'tickers'))
        delete_button.pack(fill=tk.X, pady=2)
        
        refresh_button = ttk.Button(button_frame, text="銘柄名を再取得", command=lambda: self.refresh_stock_names(self.stock_list, 'tickers'))
        refresh_button.pack(fill=tk.X, pady=2)
        
        # コピーボタン
        copy_frame = ttk.LabelFrame(operation_frame, text="銘柄のコピー", padding="5")
        copy_frame.pack(fill=tk.X, pady=5)
        
        copy_to_favorites_button = ttk.Button(copy_frame, text="お気に入りへコピー", 
                                             command=lambda: self.copy_stocks(self.stock_list, 'favorites'))
        copy_to_favorites_button.pack(fill=tk.X, pady=2)
        
        copy_to_portfolio_button = ttk.Button(copy_frame, text="ポートフォリオへコピー", 
                                              command=lambda: self.copy_to_portfolio_dialog(self.stock_list))
        copy_to_portfolio_button.pack(fill=tk.X, pady=2)
        
        # 複数選択のヘルプ
        select_help_frame = ttk.LabelFrame(operation_frame, text="操作ヘルプ", padding="5")
        select_help_frame.pack(fill=tk.X, pady=5)
        select_help_text = "※複数選択のヒント:\n・Shiftキーを押しながらクリックで範囲選択\n・Ctrl+A で全選択"
        select_help_label = ttk.Label(select_help_frame, text=select_help_text, wraplength=170, justify=tk.LEFT)
        select_help_label.pack(fill=tk.X, pady=2)
        
        # ヘルプのテキスト
        help_text = "※ティッカーシンボルの例:\n・日経平均: ^N225\n・NYダウ: ^DJI\n・日本株: [証券コード].T (例: 7974.T)\n・米国株: [シンボル] (例: AAPL)\n・ドル円為替: JPY=X"
        help_label = ttk.Label(operation_frame, text=help_text, wraplength=170, justify=tk.LEFT)
        help_label.pack(side=tk.BOTTOM, fill=tk.X, pady=5)
        
        # データの表示
        self.update_stock_list(self.stock_list, 'tickers')
        
        # 現在のターゲットリストを設定
        self.current_target_list = 'tickers'
    
    def setup_favorites_tab(self):
        """お気に入りタブのUIを設定する"""
        # 左側：お気に入りリスト
        list_frame = ttk.LabelFrame(self.favorites_tab, text="お気に入り銘柄", padding="10")
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # リストビューの作成
        columns = ("シンボル", "銘柄名")
        self.favorites_list = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="extended")
        
        # ヘッダーの設定
        for col in columns:
            self.favorites_list.heading(col, text=col, command=lambda c=col: self.sort_treeview(self.favorites_list, c, False))
            
        # カラム幅の設定
        self.favorites_list.column("シンボル", width=100)
        self.favorites_list.column("銘柄名", width=250)
        
        # リストビューをグリッドに配置
        self.favorites_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # スクロールバーの追加
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.favorites_list.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.favorites_list.configure(yscrollcommand=scrollbar.set)
        
        # ドラッグ＆ドロップの設定
        self.setup_drag_and_drop(self.favorites_list)
        
        # 右側：操作フレーム
        operation_frame = ttk.LabelFrame(self.favorites_tab, text="お気に入り操作", padding="10")
        operation_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5, pady=5)
        
        # お気に入り操作ボタン
        button_frame = ttk.Frame(operation_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        delete_button = ttk.Button(button_frame, text="選択した銘柄を削除", 
                                  command=lambda: self.delete_stock(self.favorites_list, 'favorites'))
        delete_button.pack(fill=tk.X, pady=5)
        
        refresh_button = ttk.Button(button_frame, text="銘柄名を再取得", 
                                   command=lambda: self.refresh_stock_names(self.favorites_list, 'favorites'))
        refresh_button.pack(fill=tk.X, pady=5)
        
        # コピーボタン
        copy_frame = ttk.LabelFrame(operation_frame, text="銘柄のコピー", padding="10")
        copy_frame.pack(fill=tk.X, pady=10)
        
        copy_to_main_button = ttk.Button(copy_frame, text="銘柄管理へコピー", 
                                        command=lambda: self.copy_stocks(self.favorites_list, 'tickers'))
        copy_to_main_button.pack(fill=tk.X, pady=5)
        
        copy_to_portfolio_button = ttk.Button(copy_frame, text="ポートフォリオへコピー", 
                                             command=lambda: self.copy_to_portfolio_dialog(self.favorites_list))
        copy_to_portfolio_button.pack(fill=tk.X, pady=5)
        
        # 複数選択のヘルプ
        select_help_frame = ttk.LabelFrame(operation_frame, text="操作ヘルプ", padding="10")
        select_help_frame.pack(fill=tk.X, pady=10)
        select_help_text = "※複数選択のヒント:\n・Shiftキーを押しながらクリックで範囲選択\n・Ctrl+A で全選択"
        select_help_label = ttk.Label(select_help_frame, text=select_help_text, wraplength=200, justify=tk.LEFT)
        select_help_label.pack(fill=tk.X, pady=5)
        
        # データの表示
        self.update_stock_list(self.favorites_list, 'favorites')
        
        # タブ切り替え時の処理
        self.tab_control.bind("<<NotebookTabChanged>>", self.on_tab_changed)
    
    def setup_portfolio_tab(self):
        """ポートフォリオタブのUIを設定する"""
        # ポートフォリオタブのメインフレーム
        main_frame = ttk.Frame(self.portfolio_tab)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 上部：ポートフォリオ選択
        portfolio_select_frame = ttk.LabelFrame(main_frame, text="ポートフォリオ選択", padding="10")
        portfolio_select_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # ポートフォリオ選択用コンボボックス
        portfolio_names = [p["name"] for p in self.config['portfolios']]
        self.portfolio_var = tk.StringVar()
        if portfolio_names:
            self.portfolio_var.set(portfolio_names[0])
        
        portfolio_combo = ttk.Combobox(portfolio_select_frame, textvariable=self.portfolio_var, 
                                       values=portfolio_names, state="readonly", width=40)
        portfolio_combo.pack(side=tk.LEFT, padx=5)
        portfolio_combo.bind("<<ComboboxSelected>>", self.on_portfolio_selected)
        
        # ポートフォリオ名変更ボタン
        rename_button = ttk.Button(portfolio_select_frame, text="ポートフォリオ名変更", 
                                   command=self.rename_portfolio)
        rename_button.pack(side=tk.LEFT, padx=5)
        
        # ポートフォリオ追加/削除ボタン
        add_portfolio_button = ttk.Button(portfolio_select_frame, text="新規ポートフォリオ", 
                                        command=self.add_new_portfolio)
        add_portfolio_button.pack(side=tk.LEFT, padx=5)
        
        delete_portfolio_button = ttk.Button(portfolio_select_frame, text="ポートフォリオ削除", 
                                           command=self.delete_portfolio)
        delete_portfolio_button.pack(side=tk.LEFT, padx=5)
        
        # 下部：メインコンテンツ
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 左側：ポートフォリオ内容
        list_frame = ttk.LabelFrame(content_frame, text="銘柄リスト", padding="10")
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # リストビューの作成
        columns = ("シンボル", "銘柄名")
        self.portfolio_list = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="extended")
        
        # ヘッダーの設定
        for col in columns:
            self.portfolio_list.heading(col, text=col, command=lambda c=col: self.sort_treeview(self.portfolio_list, c, False))
            
        # カラム幅の設定
        self.portfolio_list.column("シンボル", width=100)
        self.portfolio_list.column("銘柄名", width=250)
        
        # リストビューをグリッドに配置
        self.portfolio_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # スクロールバーの追加
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.portfolio_list.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.portfolio_list.configure(yscrollcommand=scrollbar.set)
        
        # ドラッグ＆ドロップの設定
        self.setup_drag_and_drop(self.portfolio_list)
        
        # 右側：操作フレーム
        operation_frame = ttk.LabelFrame(content_frame, text="ポートフォリオ操作", padding="10")
        operation_frame.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ポートフォリオ操作ボタン
        button_frame = ttk.Frame(operation_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        delete_button = ttk.Button(button_frame, text="選択した銘柄を削除", 
                                  command=self.delete_from_portfolio)
        delete_button.pack(fill=tk.X, pady=5)
        
        refresh_button = ttk.Button(button_frame, text="銘柄名を再取得", 
                                   command=self.refresh_portfolio_names)
        refresh_button.pack(fill=tk.X, pady=5)
        
        # コピーボタン
        copy_frame = ttk.LabelFrame(operation_frame, text="銘柄のコピー", padding="10")
        copy_frame.pack(fill=tk.X, pady=10)
        
        copy_to_main_button = ttk.Button(copy_frame, text="銘柄管理へコピー", 
                                        command=lambda: self.copy_stocks(self.portfolio_list, 'tickers'))
        copy_to_main_button.pack(fill=tk.X, pady=5)
        
        copy_to_favorites_button = ttk.Button(copy_frame, text="お気に入りへコピー", 
                                             command=lambda: self.copy_stocks(self.portfolio_list, 'favorites'))
        copy_to_favorites_button.pack(fill=tk.X, pady=5)
        
        copy_to_other_portfolio_button = ttk.Button(copy_frame, text="別のポートフォリオへコピー", 
                                             command=lambda: self.copy_to_portfolio_dialog(self.portfolio_list))
        copy_to_other_portfolio_button.pack(fill=tk.X, pady=5)
        
        # 複数選択のヘルプ
        select_help_frame = ttk.LabelFrame(operation_frame, text="操作ヘルプ", padding="10")
        select_help_frame.pack(fill=tk.X, pady=10)
        select_help_text = "※複数選択のヒント:\n・Shiftキーを押しながらクリックで範囲選択\n・Ctrl+A で全選択"
        select_help_label = ttk.Label(select_help_frame, text=select_help_text, wraplength=200, justify=tk.LEFT)
        select_help_label.pack(fill=tk.X, pady=5)
        
        # 初期表示
        if self.config['portfolios']:
            self.update_portfolio_list(0)
    
    def setup_settings_tab(self):
        """一般設定タブのUIを設定する"""
        settings_frame = ttk.Frame(self.settings_tab, padding="10")
        settings_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左側のコンテンツフレーム
        left_frame = ttk.Frame(settings_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        # 右側の画像フレーム
        right_frame = ttk.Frame(settings_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 0))
        
        # 右側に画像を配置
        try:
            # 1つ目の画像
            image_path1 = "C:\\Users\\rilak\\Desktop\\株価\\image\\インデックス.png"
            original_image1 = Image.open(image_path1)
            resized_image1 = original_image1.resize((200, 150), Image.Resampling.LANCZOS)
            self.settings_image1 = ImageTk.PhotoImage(resized_image1)
            
            image_label1 = ttk.Label(right_frame, image=self.settings_image1)
            image_label1.pack(pady=10)
            
            # 2つ目の画像
            image_path2 = "C:\\Users\\rilak\\Desktop\\株価\\image\\仮想通貨.png"
            original_image2 = Image.open(image_path2)
            resized_image2 = original_image2.resize((200, 150), Image.Resampling.LANCZOS)
            self.settings_image2 = ImageTk.PhotoImage(resized_image2)
            
            image_label2 = ttk.Label(right_frame, image=self.settings_image2)
            image_label2.pack(pady=10)
            
        except Exception as e:
            print(f"設定タブの画像読み込みに失敗しました: {e}")
        
        # カラム名の設定
        columns_frame = ttk.LabelFrame(left_frame, text="カラム表示設定", padding="10")
        columns_frame.pack(fill=tk.X, pady=10)
        
        self.column_var = tk.BooleanVar(value=self.config.get('use_japanese_columns', False))
        columns_check = ttk.Checkbutton(columns_frame, text="カラム名を日本語で表示", variable=self.column_var)
        columns_check.pack(fill=tk.X, pady=5)
        
        # ポートフォリオ設定
        portfolio_frame = ttk.LabelFrame(left_frame, text="ポートフォリオ設定", padding="10")
        portfolio_frame.pack(fill=tk.X, pady=10)
        
        # ポートフォリオ名変更インターフェース
        ttk.Label(portfolio_frame, text="ポートフォリオ名の一括編集:").pack(anchor=tk.W, pady=5)
        
        portfolio_list_frame = ttk.Frame(portfolio_frame)
        portfolio_list_frame.pack(fill=tk.X, pady=5)
        
        # スクロール可能なキャンバスを作成
        canvas = tk.Canvas(portfolio_list_frame)
        scrollbar = ttk.Scrollbar(portfolio_list_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # ポートフォリオ名のリスト表示
        for i, portfolio in enumerate(self.config['portfolios']):
            frame = ttk.Frame(scrollable_frame)
            frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(frame, text=f"ポートフォリオ {i+1}:").pack(side=tk.LEFT, padx=5)
            
            portfolio_name_var = tk.StringVar(value=portfolio['name'])
            portfolio_name_entry = ttk.Entry(frame, textvariable=portfolio_name_var, width=40)
            portfolio_name_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            
            # 変数をインデックスと関連付けて保存
            portfolio_name_entry.bind("<FocusOut>", lambda event, idx=i, var=portfolio_name_var: 
                                     self.update_portfolio_name(idx, var.get()))
        
        # 出力ディレクトリの設定
        dir_frame = ttk.LabelFrame(left_frame, text="出力ディレクトリ設定", padding="10")
        dir_frame.pack(fill=tk.X, pady=10)
        
        dir_selection_frame = ttk.Frame(dir_frame)
        dir_selection_frame.pack(fill=tk.X)
        
        self.dir_var = tk.StringVar(value=self.config.get('output_dir', ""))
        dir_entry = ttk.Entry(dir_selection_frame, textvariable=self.dir_var, width=50)
        dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)
        
        dir_button = ttk.Button(dir_selection_frame, text="参照...", command=self.select_directory)
        dir_button.pack(side=tk.RIGHT, padx=5, pady=5)
        
        # 設定ファイルの情報
        config_frame = ttk.LabelFrame(left_frame, text="設定ファイル情報", padding="10")
        config_frame.pack(fill=tk.X, pady=10)
        
        # 設定ファイルのパスを表示（絶対パスでハードコード）
        file_path = "C:\\Users\\rilak\\Desktop\\株価\\stock_config.json"
        
        config_path_label = ttk.Label(config_frame, text="設定ファイルの保存先:", anchor=tk.W)
        config_path_label.pack(fill=tk.X, pady=(5, 0))
        
        # パスを別のテキストで表示
        path_text = tk.Text(config_frame, height=2, width=50, wrap=tk.WORD)
        path_text.insert(tk.END, file_path)
        path_text.config(state=tk.DISABLED)  # 読み取り専用
        path_text.pack(fill=tk.X, pady=5)
        
        open_folder_button = ttk.Button(config_frame, text="設定ファイルのフォルダを開く", 
                                      command=self.open_config_folder)
        open_folder_button.pack(pady=5)
    
    def update_stock_list(self, tree_view, list_type):
        """指定されたTreeviewと設定タイプでリストを更新する"""
        # 既存の項目をクリア
        for item in tree_view.get_children():
            tree_view.delete(item)
        
        # リストタイプに応じたデータを取得
        if list_type == 'tickers':
            tickers = self.config['tickers']
        elif list_type == 'favorites':
            tickers = self.config['favorites']
        else:
            return
        
        # 新しい項目を追加
        for ticker in tickers:
            tree_view.insert("", tk.END, values=(ticker['symbol'], ticker['name']))
    
    def update_portfolio_list(self, index):
        """指定されたインデックスのポートフォリオデータでリストを更新する"""
        # 既存の項目をクリア
        for item in self.portfolio_list.get_children():
            self.portfolio_list.delete(item)
        
        # 指定されたポートフォリオのデータを取得
        if 0 <= index < len(self.config['portfolios']):
            tickers = self.config['portfolios'][index]['tickers']
            
            # 新しい項目を追加
            for ticker in tickers:
                self.portfolio_list.insert("", tk.END, values=(ticker['symbol'], ticker['name']))
            
            # 現在のポートフォリオインデックスを保存
            self.current_portfolio_index = index
    
    def add_stock(self):
        """新しい銘柄を追加する"""
        symbol = self.symbol_entry.get().strip()
        name = self.name_entry.get().strip()
        
        if not symbol:
            messagebox.showerror("エラー", "ティッカーシンボルを入力してください。")
            return
        
        # 既に存在するかチェック
        existing = False
        existing_index = -1
        for i, ticker in enumerate(self.config['tickers']):
            if ticker['symbol'] == symbol:
                existing = True
                existing_index = i
                break
        
        if existing:
            response = messagebox.askyesno("確認", 
                f"ティッカーシンボル '{symbol}' はすでに存在します。更新しますか？")
            if not response:
                return
        
        # 銘柄名が入力されていない場合かつ自動取得がチェックされている場合
        if (not name) and self.auto_name_var.get():
            # 銘柄名の自動取得処理
            self.status_var.set(f"銘柄名を取得中: {symbol}")
            self.root.update_idletasks()
            name = get_stock_name(symbol)
            
        # 上記でもnameが空の場合はシンボルをそのまま名前として使用
        if not name:
            name = symbol
        
        # 設定の更新
        stock_info = {"symbol": symbol, "name": name}
        
        if existing:
            # 既存の項目を更新
            self.config['tickers'][existing_index] = stock_info
            messagebox.showinfo("成功", f"銘柄情報を更新しました: {symbol}")
        else:
            # 新規追加
            self.config['tickers'].append(stock_info)
            messagebox.showinfo("成功", f"新しい銘柄を追加しました: {symbol}")
        
        # 入力フィールドをクリア
        self.symbol_entry.delete(0, tk.END)
        self.name_entry.delete(0, tk.END)
        
        # リストの更新
        self.update_stock_list(self.stock_list, 'tickers')
        
        # ステータスバーを元に戻す
        self.status_var.set(f"設定ファイル保存先: {CONFIG_FILE}")
    
    def save_settings(self):
        """設定を保存する"""
        # 設定の更新
        self.config['use_japanese_columns'] = self.column_var.get()
        self.config['output_dir'] = self.dir_var.get()
        
        # 設定の保存
        if save_config(self.config):
            messagebox.showinfo("成功", "設定を保存しました。")
            
            # 設定を保存した後、設定表示を更新
            self.update_config_file_display()
        else:
            pass  # エラーメッセージはsave_config内で表示される
    
    def update_config_file_display(self):
        """設定ファイル情報の表示を更新する"""
        if hasattr(self, 'config_path_label'):
            current_path = CONFIG_FILE
            self.config_path_var.set(f"設定ファイルの保存先:\n{current_path}")
    
    def select_directory(self):
        """出力ディレクトリを選択するダイアログを表示する"""
        directory = filedialog.askdirectory(
            initialdir=self.dir_var.get() if self.dir_var.get() else os.path.expanduser("~"),
            title="出力ディレクトリの選択"
        )
        
        if directory:
            self.dir_var.set(directory)
            # ディレクトリが選択されたら設定ファイル表示も更新
            self.update_config_file_display()
    
    def open_config_folder(self):
        """設定ファイルのあるフォルダをエクスプローラーで開く"""
        folder_path = os.path.dirname(CONFIG_FILE)
        try:
            os.startfile(folder_path)
        except Exception as e:
            messagebox.showerror("エラー", f"フォルダを開けませんでした: {e}")
    
    def on_tab_changed(self, event):
        """タブが切り替えられたときに呼ばれるイベントハンドラ"""
        tab_index = self.tab_control.index(self.tab_control.select())
        
        # 現在のターゲットリストの更新
        if tab_index == 0:  # 銘柄管理タブ
            self.current_target_list = 'tickers'
        elif tab_index == 1:  # お気に入りタブ
            self.current_target_list = 'favorites'
        elif tab_index == 2:  # ポートフォリオタブ
            self.current_target_list = 'portfolio'
    
    def on_portfolio_selected(self, event):
        """ポートフォリオコンボボックスが選択されたときに呼ばれるイベントハンドラ"""
        # 選択されたポートフォリオ名から対応するインデックスを検索
        selected_name = self.portfolio_var.get()
        
        for i, portfolio in enumerate(self.config['portfolios']):
            if portfolio['name'] == selected_name:
                self.current_portfolio_index = i
                self.update_portfolio_list(i)
                break
    
    def rename_portfolio(self):
        """現在選択されているポートフォリオの名前を変更する"""
        if self.current_portfolio_index < 0 or self.current_portfolio_index >= len(self.config['portfolios']):
            messagebox.showerror("エラー", "ポートフォリオが選択されていません。")
            return
        
        current_name = self.config['portfolios'][self.current_portfolio_index]['name']
        
        # 新しい名前を入力するダイアログ
        new_name = simpledialog.askstring(
            "ポートフォリオ名変更",
            "新しいポートフォリオ名を入力してください:",
            initialvalue=current_name
        )
        
        if new_name and new_name.strip():
            # 名前の重複をチェック
            duplicate = False
            for i, portfolio in enumerate(self.config['portfolios']):
                if i != self.current_portfolio_index and portfolio['name'] == new_name.strip():
                    duplicate = True
                    break
            
            if duplicate:
                messagebox.showerror("エラー", f"ポートフォリオ名 '{new_name}' はすでに使用されています。")
                return
            
            # 名前を更新
            self.config['portfolios'][self.current_portfolio_index]['name'] = new_name.strip()
            
            # コンボボックスの更新
            portfolio_names = [p["name"] for p in self.config['portfolios']]
            combo_widget = [w for w in self.portfolio_tab.winfo_children()[0].winfo_children()[0].winfo_children() 
                           if isinstance(w, ttk.Combobox)]
            if combo_widget:
                combo_widget[0]['values'] = portfolio_names
                combo_widget[0].set(new_name.strip())
            
            messagebox.showinfo("成功", f"ポートフォリオ名を '{current_name}' から '{new_name.strip()}' に変更しました。")
    
    def add_new_portfolio(self):
        """新しいポートフォリオを追加する"""
        # 最大数のチェック（30に変更）
        if len(self.config['portfolios']) >= 30:
            messagebox.showerror("エラー", "ポートフォリオは最大30個までしか作成できません。")
            return
        
        # 新しいポートフォリオ名を入力するダイアログ
        new_name = simpledialog.askstring(
            "新規ポートフォリオ",
            "新しいポートフォリオ名を入力してください:"
        )
        
        if not new_name or not new_name.strip():
            return
            
        new_name = new_name.strip()
        
        # 名前の重複をチェック
        for portfolio in self.config['portfolios']:
            if portfolio['name'] == new_name:
                messagebox.showerror("エラー", f"ポートフォリオ名 '{new_name}' はすでに使用されています。")
                return
        
        # 新しいポートフォリオを追加
        self.config['portfolios'].append({
            "name": new_name,
            "tickers": []
        })
        
        # コンボボックスの更新
        portfolio_names = [p["name"] for p in self.config['portfolios']]
        combo_widget = [w for w in self.portfolio_tab.winfo_children()[0].winfo_children()[0].winfo_children() 
                       if isinstance(w, ttk.Combobox)]
        if combo_widget:
            combo_widget[0]['values'] = portfolio_names
            combo_widget[0].set(new_name)
            
            # 新しいポートフォリオを表示
            self.current_portfolio_index = len(self.config['portfolios']) - 1
            self.update_portfolio_list(self.current_portfolio_index)
        
        messagebox.showinfo("成功", f"新しいポートフォリオ '{new_name}' を作成しました。")
        
        # 一般設定タブの更新
        self.update_portfolio_settings_ui()
    
    def update_portfolio_settings_ui(self):
        """一般設定タブのポートフォリオ名一覧を更新する"""
        # 既存のポートフォリオ設定枠を探す
        portfolio_frame = None
        for child in self.settings_tab.winfo_children()[0].winfo_children():
            if isinstance(child, ttk.LabelFrame) and "ポートフォリオ設定" in child.cget("text"):
                portfolio_frame = child
                break
        
        if not portfolio_frame:
            return
        
        # ポートフォリオリストのフレームを探す
        portfolio_list_frame = None
        for child in portfolio_frame.winfo_children():
            if isinstance(child, ttk.Frame):
                portfolio_list_frame = child
                break
        
        if not portfolio_list_frame:
            return
        
        # キャンバスを探す
        canvas = None
        scrollable_frame = None
        for child in portfolio_list_frame.winfo_children():
            if isinstance(child, tk.Canvas):
                canvas = child
                # スクロール可能なフレームを取得
                scrollable_frame = canvas.winfo_children()[0]
                break
        
        if not canvas or not scrollable_frame:
            return
        
        # 既存の項目をクリア
        for child in scrollable_frame.winfo_children():
            child.destroy()
        
        # ポートフォリオ名のリスト表示を更新
        for i, portfolio in enumerate(self.config['portfolios']):
            frame = ttk.Frame(scrollable_frame)
            frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(frame, text=f"ポートフォリオ {i+1}:").pack(side=tk.LEFT, padx=5)
            
            portfolio_name_var = tk.StringVar(value=portfolio['name'])
            portfolio_name_entry = ttk.Entry(frame, textvariable=portfolio_name_var, width=40)
            portfolio_name_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            
            # 変数をインデックスと関連付けて保存
            portfolio_name_entry.bind("<FocusOut>", lambda event, idx=i, var=portfolio_name_var: 
                                     self.update_portfolio_name(idx, var.get()))
        
        # スクロールリージョンを更新
        scrollable_frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
    
    def update_portfolio_name(self, index, new_name):
        """一般設定タブからポートフォリオ名を更新する"""
        if not new_name or not new_name.strip():
            return
        
        if index < 0 or index >= len(self.config['portfolios']):
            return
        
        current_name = self.config['portfolios'][index]['name']
        
        if current_name == new_name.strip():
            return
        
        # 名前の重複をチェック
        duplicate = False
        for i, portfolio in enumerate(self.config['portfolios']):
            if i != index and portfolio['name'] == new_name.strip():
                duplicate = True
                break
        
        if duplicate:
            messagebox.showerror("エラー", f"ポートフォリオ名 '{new_name}' はすでに使用されています。")
            # 元の名前に戻す
            self.settings_tab.after(100, lambda: self.reset_portfolio_name_entry(index, current_name))
            return
        
        # 名前を更新
        self.config['portfolios'][index]['name'] = new_name.strip()
        
        # ポートフォリオタブのコンボボックスも更新
        if hasattr(self, 'portfolio_var'):
            portfolio_names = [p["name"] for p in self.config['portfolios']]
            combo_widget = [w for w in self.portfolio_tab.winfo_children()[0].winfo_children()[0].winfo_children() 
                           if isinstance(w, ttk.Combobox)]
            if combo_widget:
                combo_widget[0]['values'] = portfolio_names
                # 現在表示中のポートフォリオが変更されたなら、表示も更新
                if index == self.current_portfolio_index:
                    combo_widget[0].set(new_name.strip())
    
    def reset_portfolio_name_entry(self, index, original_name):
        """エラー時に設定タブのポートフォリオ名入力欄を元の値に戻す"""
        portfolio_frame = None
        for child in self.settings_tab.winfo_children()[0].winfo_children():
            if isinstance(child, ttk.LabelFrame) and "ポートフォリオ設定" in child.cget("text"):
                portfolio_frame = child
                break
        
        if not portfolio_frame:
            return
        
        portfolio_list_frame = None
        for child in portfolio_frame.winfo_children():
            if isinstance(child, ttk.Frame):
                portfolio_list_frame = child
                break
        
        if not portfolio_list_frame:
            return
        
        canvas = None
        for child in portfolio_list_frame.winfo_children():
            if isinstance(child, tk.Canvas):
                canvas = child
                break
        
        if not canvas:
            return
        
        scrollable_frame = canvas.winfo_children()[0]
        if not scrollable_frame or index >= len(scrollable_frame.winfo_children()):
            return
        
        try:
            frame = scrollable_frame.winfo_children()[index]
            for child in frame.winfo_children():
                if isinstance(child, ttk.Entry):
                    child.delete(0, tk.END)
                    child.insert(0, original_name)
                    break
        except:
            pass
    
    def delete_portfolio(self):
        """現在選択されているポートフォリオを削除する"""
        if len(self.config['portfolios']) <= 1:
            messagebox.showerror("エラー", "最低でも1つのポートフォリオが必要です。")
            return
        
        if self.current_portfolio_index < 0 or self.current_portfolio_index >= len(self.config['portfolios']):
            messagebox.showerror("エラー", "ポートフォリオが選択されていません。")
            return
        
        current_name = self.config['portfolios'][self.current_portfolio_index]['name']
        
        confirmation = messagebox.askyesno(
            "確認",
            f"ポートフォリオ '{current_name}' を削除しますか？\nこの操作は元に戻せません。"
        )
        
        if confirmation:
            # ポートフォリオを削除
            del self.config['portfolios'][self.current_portfolio_index]
            
            # コンボボックスの更新
            portfolio_names = [p["name"] for p in self.config['portfolios']]
            combo_widget = [w for w in self.portfolio_tab.winfo_children()[0].winfo_children()[0].winfo_children() 
                           if isinstance(w, ttk.Combobox)]
            if combo_widget:
                combo_widget[0]['values'] = portfolio_names
                
                # 新しい選択インデックスを設定
                new_index = max(0, min(self.current_portfolio_index, len(self.config['portfolios']) - 1))
                self.current_portfolio_index = new_index
                combo_widget[0].current(new_index)
                
                # リストの更新
                self.update_portfolio_list(new_index)
            
            messagebox.showinfo("成功", f"ポートフォリオ '{current_name}' を削除しました。")
            
            # 一般設定タブの更新
            self.update_portfolio_settings_ui()
    
    def setup_drag_and_drop(self, tree_view):
        """ドラッグ＆ドロップおよび複数選択機能を設定する"""
        # マウス関連のイベントバインド
        tree_view.bind("<ButtonPress-1>", self.on_tree_press)
        tree_view.bind("<B1-Motion>", self.on_tree_motion)
        tree_view.bind("<ButtonRelease-1>", self.on_tree_release)
        
        # 複数選択のためのキーボードショートカット
        tree_view.bind("<Control-a>", lambda e: self.select_all_items(e.widget))
        
        # Shiftキーによる選択範囲を有効にするための特別な処理
        tree_view.bind("<Shift-ButtonPress-1>", self.on_shift_click)
        
    def select_all_items(self, tree_view):
        """ツリービューの全アイテムを選択する"""
        tree_view.selection_set(tree_view.get_children())
        return "break"  # イベント伝播を停止
        
    def on_shift_click(self, event):
        """Shiftキーを押しながらクリックした場合の範囲選択処理"""
        tree = event.widget
        item = tree.identify_row(event.y)
        if not item:
            return
            
        current_selection = tree.selection()
        if not current_selection:
            tree.selection_set(item)
            return "break"
            
        # 最後に選択されたアイテムと新しいアイテムの間を選択
        last_selected = current_selection[-1]
        last_index = tree.index(last_selected)
        new_index = tree.index(item)
        
        # 範囲の開始・終了インデックスを決定
        start_idx = min(last_index, new_index)
        end_idx = max(last_index, new_index)
        
        # 選択をリセットして新しい範囲を選択
        tree.selection_clear()
        for i in range(start_idx, end_idx + 1):
            try:
                item_to_select = tree.get_children()[i]
                tree.selection_add(item_to_select)
            except IndexError:
                pass
        
        return "break"  # イベント伝播を停止
    
    def on_tree_press(self, event):
        """ツリービューでマウスボタンが押されたときの処理"""
        tree = event.widget
        item = tree.identify_row(event.y)
        if not item:
            return
            
        # Ctrlキーが押されているか確認（Windowsでは4、Macでは8）
        ctrl_pressed = (event.state & 0x4) != 0
        
        # 選択状態の処理
        if not ctrl_pressed:
            # 選択されたアイテムをクリックした場合は選択を維持（ドラッグ用）
            if item not in tree.selection():
                tree.selection_set(item)
        else:
            # Ctrlキーが押されている場合は、クリックしたアイテムの選択状態を切り替え
            if item in tree.selection():
                tree.selection_remove(item)
            else:
                tree.selection_add(item)
        
        # ドラッグ開始データを保存
        self.drag_data = {
            "item": item,
            "x": event.x,
            "y": event.y,
            "tree": tree
        }
    
    def on_tree_motion(self, event):
        """ツリービューでマウスがドラッグされているときの処理"""
        if not self.drag_data["item"]:
            return
        
        # 十分な距離が移動されたらドラッグ開始と判断
        if abs(event.x - self.drag_data["x"]) > 5 or abs(event.y - self.drag_data["y"]) > 5:
            self.drag_item = self.drag_data["item"]
            
            # 複数選択時に他の選択項目も表示するとよい (将来拡張用)
            # この時点では特に処理しない
    
    def on_tree_release(self, event):
        """ツリービューでマウスボタンが離されたときの処理"""
        if not self.drag_item:
            return
        
        tree = event.widget
        target_item = tree.identify_row(event.y)
        
        if target_item and target_item != self.drag_item:
            # 移動先のインデックスを計算
            selected_items = tree.selection()
            
            # 移動処理
            self.move_items(tree, selected_items, target_item)
        
        # ドラッグ情報をリセット
        self.drag_item = None
        self.drag_data = {"item": None, "x": 0, "y": 0, "tree": None}
    
    def move_items(self, tree, selected_items, target_item):
        """選択された項目を目標位置に移動する"""
        # 選択項目が空の場合は処理しない
        if not selected_items:
            return
            
        # ターゲットがドラッグされている選択項目に含まれている場合は処理しない
        if target_item in selected_items:
            return
            
        # ターゲットの位置の決定
        target_index = tree.index(target_item)
        
        # 選択された項目の情報
        items_to_move = []
        for item in selected_items:
            values = tree.item(item, "values")
            items_to_move.append({"symbol": values[0], "name": values[1]})
            
        # 選択された項目が存在するリストの決定
        if tree == self.stock_list:
            target_list = self.config['tickers']
            list_type = 'tickers'
        elif tree == self.favorites_list:
            target_list = self.config['favorites']
            list_type = 'favorites'
        elif tree == self.portfolio_list:
            target_list = self.config['portfolios'][self.current_portfolio_index]['tickers']
            list_type = 'portfolio'
        else:
            return
        
        # 選択された項目を元のリストから削除
        symbols_to_move = [item["symbol"] for item in items_to_move]
        
        # 移動ではなく順序を入れ替えるだけなので、項目を削除した後の調整は不要
        target_list_new = [item for item in target_list if item["symbol"] not in symbols_to_move]
        
        # 項目を新しい位置に挿入
        # まず、ターゲットのインデックスを調整（削除によるインデックスの変化を考慮）
        original_indices = []
        for item in selected_items:
            original_indices.append(tree.index(item))
        
        # ターゲットのインデックスの調整（削除前のインデックスより大きいなら減算）
        adjusted_target_index = target_index
        for idx in sorted(original_indices):
            if idx < target_index:
                adjusted_target_index -= 1
        
        # 項目を挿入
        for item in items_to_move:
            target_list_new.insert(adjusted_target_index, item)
            adjusted_target_index += 1
        
        # リストを更新
        if list_type == 'tickers':
            self.config['tickers'] = target_list_new
            self.update_stock_list(tree, list_type)
        elif list_type == 'favorites':
            self.config['favorites'] = target_list_new
            self.update_stock_list(tree, list_type)
        elif list_type == 'portfolio':
            self.config['portfolios'][self.current_portfolio_index]['tickers'] = target_list_new
            self.update_portfolio_list(self.current_portfolio_index)
    
    def sort_treeview(self, tree, col, reverse):
        """ツリービューのカラムをクリックしたときの並び替え処理"""
        # 並び替え対象リストの決定
        if tree == self.stock_list:
            data = self.config['tickers']
            list_type = 'tickers'
        elif tree == self.favorites_list:
            data = self.config['favorites']
            list_type = 'favorites'
        elif tree == self.portfolio_list:
            data = self.config['portfolios'][self.current_portfolio_index]['tickers']
            list_type = 'portfolio'
        else:
            return
        
        # カラムに応じたキーを決定
        if col == "シンボル":
            key = lambda x: x["symbol"]
        else:  # 銘柄名
            key = lambda x: x["name"]
        
        # データを並び替え
        data_sorted = sorted(data, key=key, reverse=reverse)
        
        # リストを更新
        if list_type == 'tickers':
            self.config['tickers'] = data_sorted
            self.update_stock_list(tree, list_type)
        elif list_type == 'favorites':
            self.config['favorites'] = data_sorted
            self.update_stock_list(tree, list_type)
        elif list_type == 'portfolio':
            self.config['portfolios'][self.current_portfolio_index]['tickers'] = data_sorted
            self.update_portfolio_list(self.current_portfolio_index)
        
        # 次回のクリックでは逆順にするために、ヘッダーのコマンドを更新
        tree.heading(col, command=lambda c=col: self.sort_treeview(tree, c, not reverse))
    
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
            symbols = [tree_view.item(item, "values")[0] for item in selected_items[:5]]
            symbol_text = "\n".join(symbols)
            if count > 5:
                symbol_text += f"\n... 他 {count - 5} 件"
            message = f"選択された {count} 件の銘柄を削除しますか？\n\n{symbol_text}"
        
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
            symbols = [self.portfolio_list.item(item, "values")[0] for item in selected_items[:5]]
            symbol_text = "\n".join(symbols)
            if count > 5:
                symbol_text += f"\n... 他 {count - 5} 件"
            message = f"選択された {count} 件の銘柄をポートフォリオから削除しますか？\n\n{symbol_text}"
        
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
            symbols = [source_tree.item(item, "values")[0] for item in selected_items[:5]]
            symbol_text = "\n".join(symbols)
            if count > 5:
                symbol_text += f"\n... 他 {count - 5} 件"
            message = f"選択された {count} 件の銘柄を{list_name}にコピーしますか？\n\n{symbol_text}"
        
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

    def copy_to_portfolio_dialog(self, source_tree):
        """選択した銘柄をポートフォリオにコピーするためのダイアログを表示"""
        selected_items = source_tree.selection()
        if not selected_items:
            messagebox.showinfo("情報", "コピーする銘柄を選択してください。")
            return
        
        # 選択された銘柄の情報を取得
        selected_tickers = []
        for item in selected_items:
            values = source_tree.item(item, "values")
            selected_tickers.append({"symbol": values[0], "name": values[1]})
        
        count = len(selected_tickers)
        
        # ポートフォリオ選択ダイアログの作成
        dialog = tk.Toplevel(self.root)
        dialog.title("ポートフォリオ選択")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # ダイアログの内容を作成
        if count == 1:
            ticker = selected_tickers[0]
            label_text = f"銘柄 {ticker['symbol']} - {ticker['name']} を\nコピー先のポートフォリオを選択してください:"
        else:
            symbols = [ticker["symbol"] for ticker in selected_tickers[:3]]
            symbol_text = ", ".join(symbols)
            if count > 3:
                symbol_text += f" 他 {count - 3} 件"
            label_text = f"選択された {count} 件の銘柄 ({symbol_text})\nのコピー先ポートフォリオを選択してください:"
        
        ttk.Label(dialog, text=label_text, wraplength=380, justify=tk.LEFT).pack(padx=10, pady=10)
        
        # ポートフォリオのリストボックス
        frame = ttk.Frame(dialog)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL)
        listbox = tk.Listbox(frame, yscrollcommand=scrollbar.set, selectmode=tk.SINGLE, height=10)
        scrollbar.config(command=listbox.yview)
        
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ポートフォリオの一覧を表示
        for portfolio in self.config['portfolios']:
            listbox.insert(tk.END, portfolio['name'])
        
        # 現在のポートフォリオを選択状態にする
        if 0 <= self.current_portfolio_index < len(self.config['portfolios']):
            listbox.selection_set(self.current_portfolio_index)
            listbox.see(self.current_portfolio_index)
        
        # ボタンフレーム
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        cancel_button = ttk.Button(button_frame, text="キャンセル", command=dialog.destroy)
        cancel_button.pack(side=tk.RIGHT, padx=5)
        
        # OKボタンのコマンド
        def on_ok():
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("警告", "ポートフォリオを選択してください。")
                return
            
            portfolio_index = selection[0]
            target_portfolio = self.config['portfolios'][portfolio_index]
            
            self.copy_stocks_to_portfolio(selected_tickers, portfolio_index)
            dialog.destroy()
        
        ok_button = ttk.Button(button_frame, text="OK", command=on_ok)
        ok_button.pack(side=tk.RIGHT, padx=5)
    
    def copy_stocks_to_portfolio(self, selected_tickers, portfolio_index):
        """選択された銘柄を指定されたポートフォリオにコピーする"""
        if portfolio_index < 0 or portfolio_index >= len(self.config['portfolios']):
            return
        
        target_portfolio = self.config['portfolios'][portfolio_index]
        portfolio_name = target_portfolio['name']
        
        # コピー処理
        added_count = 0
        updated_count = 0
        
        for ticker in selected_tickers:
            symbol = ticker['symbol']
            name = ticker['name']
            
            # 既に存在するかチェック
            existing = False
            for existing_ticker in target_portfolio['tickers']:
                if existing_ticker['symbol'] == symbol:
                    existing = True
                    # 名前が異なる場合は更新
                    if existing_ticker['name'] != name:
                        existing_ticker['name'] = name
                        updated_count += 1
                    break
            
            # 存在しない場合は追加
            if not existing:
                target_portfolio['tickers'].append({"symbol": symbol, "name": name})
                added_count += 1
        
        # 結果メッセージ
        result_msg = f"ポートフォリオ '{portfolio_name}' に:\n{added_count}件の銘柄が追加されました。"
        if updated_count > 0:
            result_msg += f"\n{updated_count}件の銘柄名が更新されました。"
        
        messagebox.showinfo("完了", result_msg)
        
        # 現在表示中のポートフォリオが変更対象だった場合、リストを更新
        if portfolio_index == self.current_portfolio_index:
            self.update_portfolio_list(self.current_portfolio_index)

# メインエントリーポイント
if __name__ == "__main__":
    root = tk.Tk()
    app = StockConfigApp(root)
    root.mainloop()
