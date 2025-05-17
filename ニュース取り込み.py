import argparse
import csv
import datetime
import requests
import pandas as pd
import time
import os


class StockNewsScraper:
    def __init__(self, api_key=None):
        # Alpha Vantageの無料APIキーを使用する場合は設定
        # 実際の使用時には自分のAPIキーに置き換えてください
        self.api_key = api_key or "YOUR_API_KEY"
        # APIのエンドポイントを設定
        self.news_endpoint = "https://www.alphavantage.co/query"
        
    def get_news_by_ticker(self, ticker, limit=10):
        """
        銘柄コードに基づいてニュースを取得
        """
        params = {
            "function": "NEWS_SENTIMENT",
            "tickers": ticker,
            "apikey": self.api_key,
            "limit": limit
        }
        
        return self._fetch_news(params)
    
    def get_news_by_sector(self, sector, limit=10):
        """
        セクターに基づいてニュースを取得
        """
        # セクターを英語に変換（必要に応じて拡張）
        sector_mapping = {
            "テクノロジー": "technology",
            "金融": "finance",
            "ヘルスケア": "healthcare",
            "エネルギー": "energy",
            "消費財": "consumer",
            "工業": "industrial",
            "素材": "materials",
            "通信": "communication",
            "公共事業": "utilities",
            "不動産": "real_estate"
        }
        
        # 日本語のセクター名が与えられた場合は英語に変換
        english_sector = sector_mapping.get(sector, sector)
        
        params = {
            "function": "NEWS_SENTIMENT",
            "topics": english_sector,
            "apikey": self.api_key,
            "limit": limit
        }
        
        return self._fetch_news(params)
    
    def _fetch_news(self, params):
        """
        APIからニュースを取得する内部メソッド
        """
        try:
            response = requests.get(self.news_endpoint, params=params)
            response.raise_for_status()  # エラーチェック
            data = response.json()
            
            # APIレスポンスが正しい形式かチェック
            if "feed" not in data:
                if "Note" in data:
                    print(f"API制限エラー: {data['Note']}")
                    return []
                print(f"予期しないAPIレスポンス形式: {data}")
                return []
                
            return data["feed"]
        
        except requests.exceptions.RequestException as e:
            print(f"リクエストエラー: {e}")
            return []
        except ValueError as e:
            print(f"JSONパースエラー: {e}")
            return []
    
    def save_to_csv(self, news_items, output_file):
        """
        ニュース項目をCSVファイルに保存
        """
        if not news_items:
            print("保存するニュースアイテムがありません。")
            return False
            
        try:
            # ニュース項目をDataFrameに変換
            df = pd.DataFrame(news_items)
            
            # 必要な列のみ選択
            columns_to_save = [
                'title', 'summary', 'source', 'url', 'time_published', 
                'overall_sentiment_score', 'overall_sentiment_label'
            ]
            
            # 存在する列のみを選択（APIレスポンスが変わる可能性があるため）
            available_columns = [col for col in columns_to_save if col in df.columns]
            df_to_save = df[available_columns]
            
            # 日付形式を整形（例: 20230101T120000 → 2023-01-01 12:00:00）
            if 'time_published' in available_columns:
                df_to_save['time_published'] = df_to_save['time_published'].apply(
                    lambda x: datetime.datetime.strptime(x, "%Y%m%dT%H%M%S").strftime("%Y-%m-%d %H:%M:%S")
                    if isinstance(x, str) and len(x) >= 15 else x
                )
            
            # CSVに保存
            df_to_save.to_csv(output_file, index=False, encoding='utf-8-sig')
            print(f"{len(df_to_save)}件のニュースを {output_file} に保存しました。")
            return True
            
        except Exception as e:
            print(f"CSVへの保存中にエラーが発生しました: {e}")
            return False


def main():
    parser = argparse.ArgumentParser(description='証券コードまたはセクターによるニュース取得')
    parser.add_argument('--ticker', help='証券コード（例: 7203.T, AAPL）')
    parser.add_argument('--sector', help='セクター（例: テクノロジー, 金融）')
    parser.add_argument('--output', default='stock_news.csv', help='出力CSVファイル名')
    parser.add_argument('--limit', type=int, default=10, help='取得するニュース数')
    parser.add_argument('--api-key', help='Alpha Vantage APIキー')
    
    args = parser.parse_args()
    
    if not args.ticker and not args.sector:
        print("エラー: 証券コード(--ticker)またはセクター(--sector)を指定してください。")
        return
    
    scraper = StockNewsScraper(api_key=args.api_key)
    
    # 出力ディレクトリが存在しなければ作成
    output_dir = os.path.dirname(args.output)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    if args.ticker:
        print(f"証券コード '{args.ticker}' のニュースを取得しています...")
        news = scraper.get_news_by_ticker(args.ticker, args.limit)
    else:
        print(f"セクター '{args.sector}' のニュースを取得しています...")
        news = scraper.get_news_by_sector(args.sector, args.limit)
    
    if news:
        scraper.save_to_csv(news, args.output)
    else:
        print("ニュースが見つかりませんでした。")


if __name__ == "__main__":
    main()
