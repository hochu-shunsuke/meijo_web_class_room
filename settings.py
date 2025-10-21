import json
from cryptography.fernet import Fernet
import getpass
from pathlib import Path

KEY_FILE = Path("key.key")
USERDATA_FILE = Path("userdata.txt")

def _get_or_create_key():
    """暗号キーをファイルから読み込む．存在しない場合は生成して保存する．"""
    if KEY_FILE.exists():
        return KEY_FILE.read_bytes()
    else:
        print("暗号キーが見つかりません．新しいキーを生成します...")
        key = Fernet.generate_key()
        KEY_FILE.write_bytes(key)
        return key

def _create_new_userdata(fernet: Fernet):
    """ユーザーに資格情報を要求し，暗号化して保存する．"""
    print("新しい資格情報を作成します．")
    userid = input("userid:")
    password = getpass.getpass("password:")
    
    pass_setting = {"userid": userid, "password": password}
    data_json = json.dumps(pass_setting).encode()
    
    encrypted = fernet.encrypt(data_json)
    USERDATA_FILE.write_bytes(encrypted)
    print("資格情報を暗号化して保存しました．")
    return pass_setting

def load_or_create_credentials():
    """
    保存された資格情報を読み込んで復号する．
    ファイルが存在しないか復号に失敗した場合は，新規作成を促す．
    """
    key = _get_or_create_key()
    fernet = Fernet(key)
    
    if not USERDATA_FILE.exists():
        print(f"{USERDATA_FILE.name} が見つかりません．")
        return _create_new_userdata(fernet)
        
    try:
        encrypted = USERDATA_FILE.read_bytes()
        decrypted = fernet.decrypt(encrypted)
        return json.loads(decrypted.decode())
    except Exception as e: # InvalidToken, JSONDecodeError など
        print(f"資格情報の読み込みに失敗しました: {e}")
        return _create_new_userdata(fernet)