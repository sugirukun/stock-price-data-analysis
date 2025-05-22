import json
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import yfinance as yf
from time import sleep
import threading

def delete_stock(self, tree_view, list_type):
    """銘柄を削除する関数"""
    pass  # 実装を追加する必要があります

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