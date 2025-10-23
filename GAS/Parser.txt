// Parser.gs
// Config.gsの定数 (ID_REGEX, WEBCLASS_BASE_URLなど) を利用します。

/**
 * WebClassダッシュボードのHTMLから科目一覧のリンクと名前を抽出する
 * @param {string} html - ダッシュボードのHTML
 * @returns {Array<Object>} 科目オブジェクトの配列 {url: string, name: string}
 */
function parseDashboardForCourseLinks(html) {
  const result = [];
  
  // 科目リンクを抽出
  const COURSE_LINK_TIMETABLE_REGEX = /<a href='(\/webclass\/course\.php\/[a-f0-9]+\/login\?acs_=[^']*)' Target='_top'>([\s\S]*?)<\/a>/g;
  
  let match;
  const baseUrl = WEBCLASS_BASE_URL.replace(/\/$/, '');
  
  const uniqueCourses = new Map();

  while ((match = COURSE_LINK_TIMETABLE_REGEX.exec(html)) !== null) {
    const relativeUrl = match[1];
    let rawName = match[2];

    let courseName = rawName.replace(/&raquo;/, '').trim();
    courseName = courseName.replace(/<[^>]+>/g, '').trim(); 
    
    const uniqueKey = relativeUrl.split('?')[0]; 

    if (relativeUrl && !uniqueCourses.has(uniqueKey)) {
        const absoluteUrl = baseUrl + relativeUrl; 
        
        uniqueCourses.set(uniqueKey, {
            url: absoluteUrl,
            name: courseName
        });
    }
  }

  return Array.from(uniqueCourses.values());
}

/**
 * WebClassの各コースのトップページのHTMLを解析し、コンテンツリストを抽出する。
 * @returns {Array<Object>} 課題オブジェクトの配列 {title: string, share_link: string, category: string, start_period: string, end_period: string}
 */
function parseCourseContents(html) {
  const result = [];
  
  // 課題リストのコンテナ全体を正規表現で抽出
  const itemsMatch = html.match(/<section class=\"list-group-item cl-contentsList_listGroupItem\"[\s\S]*?<\/section>/g);
  
  if (!itemsMatch) return result;
  
  const baseUrl = WEBCLASS_BASE_URL;

  itemsMatch.forEach(itemHtml => {
    
    let title = "";
    let shareLink = "";
    let category = "";
    let startPeriod = ""; 
    let endPeriod = "";   

    // 3. タイトル要素の抽出
    const nameTagMatch = itemHtml.match(/<h4\s+[^>]*?class=\"cm-contentsList_contentName\"[^>]*?>([\s\S]*?)<\/h4>/);
    if (nameTagMatch) {
      let content = nameTagMatch[1];
      
      content = content.replace(/<div class=\"cl-contentsList_new\">[\s\S]*?<\/div>/g, '').trim();
      
      title = content.replace(/<[^>]+>/g, '').trim();

      const urlMatch = content.match(/<a href=\"([^\"]+)\">/);
      if (urlMatch) {
        const url = urlMatch[1];
        
        const idMatch = url.match(ID_REGEX); 
        if (idMatch) {
          shareLink = `${baseUrl}/webclass/login.php?id=${idMatch[1]}&page=1&auth_mode=SAML`;
        }
      }
    }
    
    // 4. カテゴリの抽出
    const categoryMatch = itemHtml.match(/<div\s+[^>]*?class=\"cl-contentsList_categoryLabel\"[^>]*?>([\s\S]*?)<\/div>/);
    if (categoryMatch) {
      category = categoryMatch[1].replace(/<[^>]+>/g, '').trim();
    }
    
    // 5. 利用可能期間の抽出と分割 (★修正箇所)
    // '利用可能期間' ラベルの直後にある 'cm-contentsList_contentDetailListItemData' の内容をキャプチャ
    const periodSectionMatch = itemHtml.match(
      /利用可能期間<\/div>\s*<div\s+[^>]*?class=['"]cm-contentsList_contentDetailListItemData['"][^>]*?>\s*([\s\S]*?)\s*<\/div>/
    );

    if (periodSectionMatch && periodSectionMatch[1]) {
      // 抽出された内容からHTMLタグと前後の空白を除去
      const rawPeriod = periodSectionMatch[1].replace(/<[^>]+>/g, '').trim();
      
      // rawPeriodを「 - 」で分割
      const parts = rawPeriod.split(' - ');
      
      if (parts.length === 2) {
        // 開始日時と終了日時が存在
        startPeriod = parts[0].trim(); // 例: 2025/11/04 17:00
        endPeriod = parts[1].trim();   // 例: 2025/11/05 23:59
      } else if (parts.length === 1 && rawPeriod) {
        // 分割できなかったがデータがある場合
        endPeriod = rawPeriod;
      }
    }
    
    // 結果に追加
    if (title && shareLink) {
        result.push({
            title: title,
            share_link: shareLink,
            category: category,
            start_period: startPeriod,
            end_period: endPeriod 
        });
    }
  });
  
  return result;
}