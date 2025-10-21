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
    result = []
    
    panels = soup.find_all('section', class_='panel panel-default cl-contentsList_folder')
    if not panels:
        # 代替の構造を探す (もしあれば)
        pass

    for panel in panels:
        heading = panel.find('div', class_='panel-heading')
        panel_title = heading.find('h4', class_='panel-title').get_text(strip=True) if heading else ""
        
        list_group = panel.find('div', class_='list-group')
        items = []
        if list_group:
            for item in list_group.find_all('section', class_='list-group-item cl-contentsList_listGroupItem'):
                is_new = item.find('div', class_='cl-contentsList_new') is not None
                name_tag = item.find('h4', class_='cm-contentsList_contentName')
                # cl-contentsList_new（New）を除去
                if name_tag:
                    for new_tag in name_tag.find_all('div', class_='cl-contentsList_new'):
                        new_tag.decompose()
                a_tag = name_tag.find('a') if name_tag else None
                title = name_tag.get_text(strip=True) if name_tag else ""
                url = a_tag['href'] if a_tag and a_tag.has_attr('href') else ""
                
                id_match = ID_REGEX.search(url)
                share_link = f"https://rpwebcls.meijo-u.ac.jp/webclass/login.php?id={id_match.group(1)}&page=1&auth_mode=SAML" if id_match else ""
                
                category_tag = item.find('div', class_='cl-contentsList_categoryLabel')
                category = category_tag.get_text(strip=True) if category_tag else ""
                
                period = ""
                for label, data in zip(item.find_all('div', class_='cm-contentsList_contentDetailListItemLabel'), item.find_all('div', class_='cm-contentsList_contentDetailListItemData')):
                    if "利用可能期間" in label.get_text():
                        period = data.get_text(strip=True)
                
                items.append({
                    'title': title,
                    'url': url,
                    'share_link': share_link,
                    'is_new': is_new,
                    'category': category,
                    'period': period
                })
        result.append({
            'panel_title': panel_title,
            'items': items
        })
    return result