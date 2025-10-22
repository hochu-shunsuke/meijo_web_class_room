// Parser.gs
// Config.gsの定数 (ID_REGEX, WEBCLASS_BASE_URLなど) を利用します。

function parseCourseContents(html) {
  const result = [];
  
  // 課題リストのコンテナ全体を正規表現で抽出
  const itemsMatch = html.match(/<section class="list-group-item cl-contentsList_listGroupItem"[\s\S]*?<\/section>/g);
  
  if (!itemsMatch) return result;
  
  itemsMatch.forEach(itemHtml => {
    // 1. Newフラグのチェック (現在は使用しないが、データとして保持)
    // const isNew = itemHtml.includes('cl-contentsList_new');
    
    // 2. タイトル、URL、カテゴリ、期間の初期値
    let title = "";
    let shareLink = "";
    let category = "";
    let period = "";

    // 3. タイトル要素の抽出
    const nameTagMatch = itemHtml.match(/<h4 class="cm-contentsList_contentName">([\s\S]*?)<\/h4>/);
    if (nameTagMatch) {
      let content = nameTagMatch[1];
      
      // Newタグと空白を除去
      content = content.replace(/<div class="cl-contentsList_new">[\s\S]*?<\/div>/g, '').trim();
      
      // タイトル（タグを除去したテキスト）
      title = content.replace(/<[^>]+>/g, '').trim();

      // URLの抽出
      const urlMatch = content.match(/<a href="([^"]+)">/);
      if (urlMatch) {
        const url = urlMatch[1];
        
        // Share Linkの生成 (Config.gs の ID_REGEX を利用)
        const idMatch = url.match(ID_REGEX); 
        if (idMatch) {
          shareLink = `${WEBCLASS_BASE_URL}/webclass/login.php?id=${idMatch[1]}&page=1&auth_mode=SAML`;
        }
      }
    }
    
    // 4. カテゴリの抽出
    const categoryMatch = itemHtml.match(/<div class="cl-contentsList_categoryLabel">([\s\S]*?)<\/div>/);
    if (categoryMatch) {
      category = categoryMatch[1].replace(/<[^>]+>/g, '').trim();
    }
    
    // 5. 利用可能期間の抽出
    const periodSectionMatch = itemHtml.match(/利用可能期間[\s\S]*?<div class="cm-contentsList_contentDetailListItemData">([\s\S]*?)<\/div>/);
    if (periodSectionMatch) {
      period = periodSectionMatch[1].trim();
    }
    
    // 抽出結果の整理
    if (title && shareLink) {
        result.push({
            'title': title, 'share_link': shareLink,
            'category': category, 'period': period
        });
    }
  });

  return result;
}