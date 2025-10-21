import re
import urllib.parse
import json # JSONはデータ送信のために必要
import requests # HTTP通信のために必要

from bs4 import BeautifulSoup
# --- スプレッドシート連携用に追加したgspreadなどのimportは削除 ---

import settings
import os
from pathlib import Path
try:
    # optional: prefer python-dotenv if available for local .env files
    from dotenv import load_dotenv
    _DOTENV_AVAILABLE = True
except Exception:
    _DOTENV_AVAILABLE = False
from webclass_client import WebClassClient
from parser import parse_course_contents

# --- 定数（GAS Webアプリ用） ---
# ★ここをGASをWebアプリとして公開したURLに置き換えてください！
# GAS WebアプリURLは環境変数から取得します。
# ローカル開発時はプロジェクトルートの .env に GAS_WEB_APP_URL=... を書くと便利です。
if _DOTENV_AVAILABLE:
    # load .env from project root if present
    project_root = Path(__file__).resolve().parent
    dotenv_path = project_root / '.env'
    if dotenv_path.exists():
        load_dotenv(dotenv_path)

GAS_WEB_APP_URL = os.environ.get('GAS_WEB_APP_URL')
if not GAS_WEB_APP_URL:
    # フォールバック（必要ならここで空文字や例を設定）
    GAS_WEB_APP_URL = ''
# ------------------------------------

REDIRECT_REGEX = re.compile(r'window.location.href\s*=\s*"([^"]+)"')

def get_course_links(client: WebClassClient):
    """ダッシュボードから科目一覧のリンクを取得する"""
    print("ダッシュボードから科目リンクを取得中...")
    try:
        # クライアントの既存セッションを使用
        top_html = client.get(client.dashboard_url).text
        soup = BeautifulSoup(top_html, "html.parser")
        
        course_links = set()
        for a_tag in soup.find_all('a', href=True, class_='list-group-item course'):
            name = a_tag.get_text(strip=True)
            href = a_tag.get("href")
            if href and '/webclass/course.php/' in href:
                course_links.add((name, href))
                
        print(f"科目リンク抽出数: {len(course_links)}")
        if not course_links:
            print("警告: 科目が見つかりません．")
        return list(course_links)
        
    except Exception as e:
        print(f"科目リンクの取得に失敗しました: {e}")
        return []


def fetch_and_parse_course(course_info, client: WebClassClient):
    """
    単一の科目のページを取得，解析し，データを返す (JSON保存処理を削除)
    """
    course_name, href = course_info
    
    session = client.session 
    base_url = client.base_url
    
    try:
        url = urllib.parse.urljoin(base_url, href)
        
        res = session.get(url)
        html = res.text
        
        # JavaScriptリダイレクト処理 (省略なしで貼り付け)
        soup = BeautifulSoup(html, "html.parser")
        script = soup.find("script")
        if script and script.string and "window.location.href" in script.string:
            m = REDIRECT_REGEX.search(script.string)
            if m:
                redirect_url = urllib.parse.urljoin(base_url, m.group(1))
                res = session.get(redirect_url)
                html = res.text

        # HTMLを解析
        data = parse_course_contents(html)
        
        return course_name, data, "成功" 

    except Exception as e:
        return course_name, None, f"失敗: {e}"

def send_data_to_gas(assignments):
    """
    収集した課題データをJSON形式でGAS Webアプリに送信する
    """
    try:
        headers = {'Content-Type': 'application/json'}
        # JSON形式に変換して送信
        payload = json.dumps({'assignments': assignments})
        
        print(f"GAS Webアプリにデータを送信中... URL: {GAS_WEB_APP_URL}")
        
        response = requests.post(GAS_WEB_APP_URL, headers=headers, data=payload)
        response.raise_for_status() # HTTPエラーチェック

        # GASからの応答を解析
        gas_response = response.json()
        
        if gas_response.get('status') == 'SUCCESS':
            print("GASへのデータ送信と書き込みに成功しました。")
        else:
            print(f"GAS側でエラーが発生しました: {gas_response.get('message', '不明なエラー')}")
            
    except requests.exceptions.RequestException as e:
        print(f"GASへのHTTP通信中にエラーが発生しました: {e}")
    except json.JSONDecodeError:
        print(f"GASからの応答が不正です: {response.text}")


def main():
    # 1. 資格情報取得
    try:
        userdata = settings.load_or_create_credentials()
    except Exception as e:
        print(f"資格情報の読み込みに失敗しました: {e}")
        return

    # 2. WebClassログイン
    try:
        client = WebClassClient(userdata["userid"], userdata["password"])
    except Exception as e:
        print(f"ログインに失敗しました: {e}")
        return

    # 3. 科目一覧を取得
    course_links = get_course_links(client)
    if not course_links:
        print("処理する科目がありません．終了します．")
        return

    # 4. 全科目を処理し、課題データを収集
    all_assignments = []
    print(f"{len(course_links)}件の科目を処理します...")
    
    for i, course in enumerate(course_links):
        course_name, parsed_data, result = fetch_and_parse_course(course, client)
        print(f"  ({i+1}/{len(course_links)}) [{result}] - {course_name}")
        
        if result == "成功" and parsed_data:
            for panel in parsed_data:
                for item in panel['items']:
                    
                    # 締切日時（終了日時）の抽出（データ整形）
                    period_str = item.get('period', '')
                    due_date_str = ''
                    if ' - ' in period_str:
                        due_date_str = period_str.split(' - ')[1].strip()
                    else:
                        due_date_str = period_str

                    # GASの列構成に合わせてデータを整形
                    # [0:ソース, 1:授業名, 2:課題タイトル, 3:締切日時, 4:課題リンク (URL), 5:Tasks ID, 6:登録済みフラグ]
                    row = [
                        'WebClass',
                        course_name,
                        item['title'] + f" ({item['category']})",
                        due_date_str,
                        item['share_link'],
                        '',
                        ''
                    ]
                    all_assignments.append(row)

    if not all_assignments:
        print("WebClassからコンテンツは見つかりませんでした．")
        return
        
    # 5. GAS Webアプリにデータを送信
    send_data_to_gas(all_assignments)
        
    print("すべての処理が完了しました．")

if __name__ == "__main__":
    main()