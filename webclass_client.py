# 標準ライブラリと外部依存
import urllib.parse
import requests
import time
import json
import random
import os
from typing import Optional, Dict
from bs4 import BeautifulSoup
from pathlib import Path
try:
    from dotenv import load_dotenv
    _DOTENV_AVAILABLE = True
except Exception:
    _DOTENV_AVAILABLE = False

# dotenv が利用可能ならプロジェクトルートの .env を読み込む
if _DOTENV_AVAILABLE:
    project_root = Path(__file__).resolve().parent
    dotenv_path = project_root / '.env'
    if dotenv_path.exists():
        load_dotenv(dotenv_path)

# SSO URL と WebClass ベース URL
SSO_URL = os.environ.get('SSO_URL')
MAX_RETRIES = int(os.environ.get('MAX_RETRIES', '5'))
RETRY_DELAY = int(os.environ.get('RETRY_DELAY', '1'))
WEBCLASS_BASE_URL = os.environ.get('WEBCLASS_BASE_URL')
#　サーバに迷惑をかけないように、User-Agentは多めに用意してランダムに
# USER_AGENTS は環境変数 USER_AGENTS に JSON 配列文字列で指定できます。
_env_user_agents = os.environ.get('USER_AGENTS')
if _env_user_agents:
    try:
        USER_AGENTS = json.loads(_env_user_agents)
        if not isinstance(USER_AGENTS, list):
            raise ValueError("USER_AGENTS must be a JSON array of strings")
    except Exception:
        # フォールバックとしてデフォルトリストを使う
        USER_AGENTS = [
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36',
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.1 Safari/605.1.15',
            'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36',
            'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:40.0) Gecko/20100101 Firefox/40.1',
            'Mozilla/5.0 (iPhone; CPU iPhone OS 14_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Mobile/15A372 Safari/604.1'
        ]
else:
    USER_AGENTS = [
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.1 Safari/605.1.15',
        'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36',
        'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:40.0) Gecko/20100101 Firefox/40.1',
        'Mozilla/5.0 (iPhone; CPU iPhone OS 14_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Mobile/15A372 Safari/604.1',
        'Mozilla/5.0 (Linux; Android 10; SM-G975F) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.106 Mobile Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/117.0'
    ]


def build_headers(referer: Optional[str] = None) -> Dict[str, str]:
    ua = random.choice(USER_AGENTS)
    headers = {
        "User-Agent": ua,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "ja-JP,ja;q=0.9,en-US;q=0.8,en;q=0.7",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
    }
    if referer:
        headers["Referer"] = referer
    return headers

def _get_sso_token(userid, password):
    """SSOサーバーから認証トークンを取得する"""
    headers = {'Content-Type': 'application/json'}
    
    try:
        # 1. 認証テンプレートの取得
        source_res = requests.post(SSO_URL, headers=headers)
        source_res.raise_for_status()
        jsn = source_res.json()
        
        # 2. テンプレートに資格情報を設定
        jsn["callbacks"][0]["input"][0]["value"] = userid
        jsn["callbacks"][1]["input"][0]["value"] = password
        
        # 3. 資格情報を送信
        for atdatat in range(MAX_RETRIES):
            token_res = requests.post(SSO_URL, headers=headers, json=jsn)
            if token_res.status_code == 200:
                token_data = token_res.json()
                print(f"SSOトークン取得成功．")
                return token_data["tokenId"]
            
            print(f"認証試行 {atdatat + 1}/{MAX_RETRIES} 失敗 (Status: {token_res.status_code}). {RETRY_DELAY}秒後リトライ...")
            time.sleep(RETRY_DELAY)
            
        raise Exception(f"トークン取得失敗 (Status: {token_res.status_code})")

    except requests.exceptions.RequestException as e:
        print(f"SSOサーバーへの接続に失敗しました: {e}")
        raise
    except (KeyError, IndexError, json.JSONDecodeError) as e:
        print(f"SSOサーバーの応答形式が予期したものではありません: {e}")
        raise


def _get_acs_path(source):
    """HTMLソースからJavaScript内のリダイレクトパスを抽出する"""
    soup = BeautifulSoup(source, "html.parser")
    script_tag = soup.find("script")
    if script_tag and script_tag.string and '"' in script_tag.string:
        exccode = script_tag.string
        parts = exccode.split('"')
        if len(parts) > 1:
            return parts[1].replace('&amp;', "&")
    return None

class WebClassClient:
    def __init__(self, userid, password):
        print("認証セッションを開始します...")
        self.session = requests.Session()
        self.session.headers.update(build_headers())
        self.base_url = WEBCLASS_BASE_URL
        self.dashboard_url = None
        self._login(userid, password)
        # 認証クッキーの保存ロジックを削除

    def _login(self, userid, password):
        try:
            # 1. SSOトークンを取得
            token_id = _get_sso_token(userid, password)
            
            # 2. WebClassのSAMLログインページにアクセス
            login_url = f"{self.base_url}/webclass/login.php?auth_mode=SAML"
            res = self.session.get(login_url, allow_redirects=False)
            sso_location = res.headers["Location"]
            
            # 3. SSOのページにトークンをCookieとして送信
            sso_cookies = {"iPlanetDirectoryPro": token_id}
            sso_res = self.session.get(sso_location, cookies=sso_cookies)
            sso_res.raise_for_status()

            # 4. SAML認証情報をパース
            soup = BeautifulSoup(sso_res.text, "html.parser")
            inputs = soup.find_all("input")
            data = {
                "SAMLResponse": inputs[0].attrs["value"],
                "RelayState": inputs[1].attrs["value"]
            }
            
            # 5. SAML認証情報をWebClassのACSにPOST
            acs_url = f"{self.base_url}/simplesaml/module.php/saml/sp/saml2-acs.php/default-sp"
            self.session.post(acs_url, data=data, allow_redirects=False)
            
            # 6. 再度ログインページにアクセス (SAML Cookieで認証される)
            login_php_res = self.session.get(login_url, allow_redirects=False)
            
            # 7. 最終的なダッシュボードのURLを取得
            acs_path = _get_acs_path(login_php_res.text)
            if not acs_path:
                raise Exception("最終的なリダイレクトパス(acsPath)の取得に失敗しました．")
                
            self.dashboard_url = urllib.parse.urljoin(self.base_url, acs_path)
            
            # 8. 最終ページにアクセスしてセッションを確定
            self.session.get(self.dashboard_url).raise_for_status()
            
            print(f"ログイン成功！")
        
        except Exception as e:
            print(f"ログイン処理中にエラーが発生しました: {e}")
            raise

    # get_auth_cookies メソッドを削除
    
    def get(self, url, referer: Optional[str] = None, **kwargs):
        """認証済みセッションでGETリクエストを送信する。Refererも任意指定可"""
        headers = build_headers(referer=referer)
        return self.session.get(url, headers=headers, **kwargs)