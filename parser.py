import re
from bs4 import BeautifulSoup

# 正規表現をコンパイル
ID_REGEX = re.compile(r'id=([a-f0-9]+)')

def parse_course_contents(html):
    """
    WebClassの各コースのトップページのHTMLを解析し，
    コンテンツリストを抽出する．
    """
    soup = BeautifulSoup(html, 'html.parser')
    items = soup.find_all('section', class_='list-group-item cl-contentsList_listGroupItem')
    
    result = []
    for item in items:
        # Newフラグ
        is_new = bool(item.find('div', class_='cl-contentsList_new'))
        
        # タイトルとURL
        title = ""
        url = ""
        share_link = ""
        
        if name_tag := item.find('h4', class_='cm-contentsList_contentName'):
            # cl-contentsList_newを除去
            for new_tag in name_tag.find_all('div', class_='cl-contentsList_new'):
                new_tag.decompose()
            
            title = name_tag.get_text(strip=True)
            
            # URLの取得
            if a_tag := name_tag.find('a'):
                url = a_tag.get('href', '')
                # IDの抽出とshare_linkの生成
                if id_match := ID_REGEX.search(url):
                    share_link = f"https://rpwebcls.meijo-u.ac.jp/webclass/login.php?id={id_match.group(1)}&page=1&auth_mode=SAML"
        
        # カテゴリ
        category = ""
        if category_tag := item.find('div', class_='cl-contentsList_categoryLabel'):
            category = category_tag.get_text(strip=True)
        
        # 利用可能期間
        period = ""
        for detail_item in item.find_all('div', class_='cm-contentsList_contentDetailListItem'):
            if label := detail_item.find('div', class_='cm-contentsList_contentDetailListItemLabel'):
                if "利用可能期間" in label.get_text():
                    if data := detail_item.find('div', class_='cm-contentsList_contentDetailListItemData'):
                        period = data.get_text(strip=True)
                    break
        
        result.append({
            'title': title,
            'url': url,
            'share_link': share_link,
            'is_new': is_new,
            'category': category,
            'period': period
        })
    
    return result