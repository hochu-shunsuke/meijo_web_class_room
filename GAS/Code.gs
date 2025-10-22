// Code.gs

// ====================================================================
//  GAS 課題リマインドシステム - 最終完全版 (WebClass内部取得 & Tasks統合)
//  Config.gs のすべての定数を利用します。
// ====================================================================

// --- メインフロー実行関数 ---

/**
 * 【メイン関数】毎日定刻に実行するトリガーに設定します。
 */
function dailySystemRun() {
  Logger.log('--- 課題リマインドシステム 実行開始 ---');
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 1. WebClassの最新データを取得し、シートに書き込む (WebClass認証が必要)
    fetchWebClassAssignments();
    
    // 2. Classroomの最新データを取得し、シートに書き込む
    fetchClassroomAssignments(); 
    
    // 3. Tasksの完了状態をスプレッドシートに反映させる（同期）
    syncTaskCompletionStatus(); 
    
    // 4. スプレッドシートの統合データから、未登録・期限内の課題をTasksに追加
    integrateAndRegisterTasks();
    
    // 5. 古いデータや完了済みのデータをシートから削除（整理）
    cleanupOldAssignments(); 
    
    Logger.log('--- 課題リマインドシステム 実行完了 ---');
    ui.alert('✅ 全ての課題取得と同期処理が完了しました！');
    
  } catch(e) {
    Logger.log(`致命的なシステムエラー: ${e.toString()}`);
    // 意図的な診断中断エラーの場合もアラートを表示
    if (e.message.includes('WebClass課題抽出診断のため') || e.message.includes('WebClassダッシュボードアクセス失敗')) {
        ui.alert(`🚨 診断中断: ${e.message}`);
    } else if (e.message.includes('SSOトークン取得に失敗しました')) {
      // ログインエラーの場合は、ここでは追加アラートを出さない
    } else {
      ui.alert(`🚨 致命的なシステムエラーが発生しました。\nLoggerをご確認ください。\n${e.message}`);
    }
  }
}

// --- 課題データ収集関数 ---

/**
 * WebClassの課題を取得し、シートに書き込む (GAS内部でWebClass認証を実行)
 * 【重要】この関数内に診断ログ強制出力ロジックを組み込んでいます。
 */
function fetchWebClassAssignments() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_WEBCLASS);
  
  if (!sheet) {
    Logger.log('エラー: WebClass課題シートが見つかりません。');
    logDiagnosisStatus('エラー: WebClass課題シートが見つかりません。', 4);
    return;
  }
  
  // 1. 認証情報の取得
  let credentials;
  try {
      credentials = getCredentials(); // Properties.gs の関数
      if (!credentials) {
        Logger.log('🚨 WebClass認証情報が未設定のため、WebClassの課題取得をスキップしました。');
        logDiagnosisStatus('🚨 1. WebClass認証情報が未設定のため、スキップしました。', 4);
        return;
      }
      logDiagnosisStatus('✅ 1. 認証情報を取得しました。', 4);
  } catch (e) {
      logDiagnosisStatus(`🚨 1. 認証情報取得でエラー: ${e.message}`, 4);
      return;
  }

  // 2. WebClassログインとセッション確立
  let loginResultUrl;
  let primaryDashboardUrl;
  try {
    // Auth.gs の loginWebClass は、最終リダイレクトURLまたはベースURLを返す
    loginResultUrl = loginWebClass(credentials.userid, credentials.password); // Auth.gs の関数
    
    // Pythonクライアントのロジックを反映: 最終リダイレクトURLがダッシュボードである可能性が高い
    primaryDashboardUrl = loginResultUrl;
    
    // ベースURLの整形
    const baseUri = WEBCLASS_BASE_URL.replace(/\/$/, ''); // 末尾のスラッシュを除去

    // ログイン後のURLがベースURL自体だった場合（不完全な場合）の補完はAuth.gs側で処理されるべきだが、
    // ここでは念のため、もしベースURLが返された場合に備えて、最も可能性の高いパスで補完する。
    if (loginResultUrl === baseUri || loginResultUrl.endsWith(baseUri + '/')) {
        primaryDashboardUrl = baseUri + '/webclass/'; // Pythonクライアントが抽出する可能性のあるルート
        logDiagnosisStatus(`✅ 2. SSO成功。URLが不完全だったため、最も可能性の高い ${primaryDashboardUrl} を試行開始点とします。`, 5);
    } else {
        logDiagnosisStatus('✅ 2. WebClassログインとセッション確立に成功しました。（SSO完了）', 5);
    }
    
  } catch(e) {
    Logger.log(`🚨 WebClassログイン失敗: ${e.toString()}`);
    logDiagnosisStatus(`🚨 2. WebClassログイン失敗: ${e.message}`, 5);
    ui.alert(`🚨 WebClassログインに失敗しました。認証情報をご確認ください。\nエラー: ${e.message}`);
    return;
  }
  
  // 3. 科目一覧を取得 (複数URLを試行)
  let courseLinks = [];
  let dashboardUrlFound = false;

  // 試行するダッシュボードパスのリスト（Pythonの挙動から予測されるパス）
  const baseUri = WEBCLASS_BASE_URL.replace(/\/$/, '');
  const dashboardUrls = [
    primaryDashboardUrl,              // Auth.gsが返した最終URL
    baseUri + '/webclass/my/index.php', // よくあるパス1 (前回試したパス)
    baseUri + '/webclass/index.php',    // よくあるパス2
    baseUri + '/webclass/',             // WebClassのトップディレクトリ
    baseUri + '/index.php'              // 最も単純なパス
  ];
  
  // 試行開始
  for (let i = 0; i < dashboardUrls.length; i++) {
    const currentUrl = dashboardUrls[i];
    
    try {
        logDiagnosisStatus(`➡️ 3. 科目リンクの取得試行開始（${i + 1}/${dashboardUrls.length}）。アクセスURL: ${currentUrl}`, 6); 
        
        // getCourseLinks内でUrlFetchApp.fetchが実行される
        courseLinks = getCourseLinks(currentUrl); 
        
        // 404エラーの場合はtry-catchで捕捉されるため、ここに到達したら成功とみなす
        dashboardUrlFound = true; 

        if (courseLinks.length === 0) {
            Logger.log('WebClass: 処理する科目がありません。');
            logDiagnosisStatus(`⚠️ 3. 処理する科目がありませんでした。（ダッシュボード確認: ${i + 1}/${dashboardUrls.length}）`, 6);
        } else {
            logDiagnosisStatus(`✅ 3. 科目一覧（${courseLinks.length}件）を取得しました。`, 6);
        }
        
        // 正しいダッシュボードURLが見つかったので、primaryDashboardUrlを更新し、ループを抜ける
        primaryDashboardUrl = currentUrl;
        break; 
        
    } catch (e) {
        // 404エラーの場合は次のURLを試す
        if (e.message.includes('404')) {
            logDiagnosisStatus(`🚨 3. 取得試行失敗 (${i + 1}/${dashboardUrls.length} - 404エラー)。次のURLを試します。`, 6);
            continue;
        } 
        
        // その他の致命的なエラーの場合はループを中断し、エラーを再スロー
        Logger.log(`WebClass: 科目リンクの取得中に致命的なエラー: ${e.message}`);
        throw new Error(`科目リンク取得エラー: ${e.message}`);
    }
  }

  // すべてのURLを試しても成功しなかった場合
  if (!dashboardUrlFound) {
    logDiagnosisStatus(`🚨 3. 全てのダッシュボードURL試行が失敗。最終レスポンス全文をC列に出力します。`, 6);
    // 最後に試したURLでエラー時のレスポンスを強制取得
    const finalUrl = dashboardUrls[dashboardUrls.length - 1];
    const errorHtml = fetchErrorResponseHtml(finalUrl); 
    logHtmlForDiagnosis(errorHtml, `WebClassダッシュボードアクセス時 最終404エラーHTML`);
    
    throw new Error("WebClassダッシュボードアクセス失敗（404）。正しいURLを特定できませんでした。");
  }
  
  // 4. 全科目を処理し、課題データを収集
  const allAssignments = [];
  let isFirstCourse = true; // ログ出力用フラグを初期化

  courseLinks.forEach(courseInfo => {
    const courseName = courseInfo.name;
    const href = courseInfo.href;
    
    try {
      // fetchAndParseCourseは { assignments: [...], html: "..." } を返す前提
      const { assignments: courseData, html: courseHtml } = fetchAndParseCourse(courseName, href);
      
      // ★★★ 診断ログ強制出力部分（最初の科目で無条件停止） ★★★
      if (isFirstCourse) {
        logDiagnosisStatus(`✅ 4. 最初の科目（${courseName}）のHTMLを取得しました。C3セルを確認してください。`, 7);
        // HTMLコンテンツをC3セルに強制書き込み
        logHtmlForDiagnosis(courseHtml, `WebClass科目HTML診断ログ: ${courseName} (パース結果: ${courseData.length}件)`);
        isFirstCourse = false; 
        
        // 処理を中断し、ユーザーにログ確認を促す
        throw new Error("WebClass課題抽出診断のため、処理を中断しました。");
      }
      // ★★★ 診断ログ強制出力部分 終わり ★★★

      courseData.forEach(item => {
        const periodStr = item.period || '';
        let dueDateStr = periodStr.includes(' - ') ? periodStr.split(' - ')[1].trim() : periodStr;
        
        // Config.gs の HEADER に合わせたデータ整形
        const row = [
          'WebClass',
          courseName,
          item.title + (item.category ? ` (${item.category})` : ''),
          dueDateStr,
          item.share_link,
          '', // Tasks ID
          ''  // 登録済みフラグ
        ];
        allAssignments.push(row);
      });
      
    } catch(e) {
      if (e.message.includes('WebClass課題抽出診断のため')) {
          throw e; // 意図的な中断は再スロー
      }
      Logger.log(`WebClass: [失敗] - ${courseName}: ${e.message}`);
    }
  });

  // 5. シートに書き込み（Tasks IDとフラグを保持しながらマージ）
  writeMergedAssignmentsToSheet(sheet, allAssignments);
  Logger.log(`${SHEET_NAME_WEBCLASS}に${allAssignments.length}件の課題を書き込みました。`);
}

/**
 * ダッシュボードから科目一覧のリンクを取得する
 */
function getCourseLinks(dashboardUrl) {
  Logger.log("WebClass: ダッシュボードから科目リンクを取得中...");
  const courseLinks = [];
  
  // UrlFetchApp.fetch() はエラー発生時に例外をスローする（getCourseLinksの呼び出し元で処理）
  const res = fetchWithSession(dashboardUrl); 
  const topHtml = res.getContentText();
  
  // Pythonのget_course_linksを参考に、より正確な抽出を行う
  const soup = new BeautifulSoup(topHtml, "html.parser"); // BeautifulSoupをGASでシミュレーション（簡易）
  
  // Pythonコードから: class_='list-group-item course' を探している
  // GASで簡易的に正規表現で抽出
  const linkMatches = topHtml.match(/<a[\s\S]*?class="list-group-item course"[\s\S]*?href="([^"]+)"[\s\S]*?>([\s\S]*?)<\/a>/g);
  
  if (linkMatches) {
    linkMatches.forEach(match => {
      const hrefMatch = match.match(/href="([^"]+)"/);
      const nameMatch = match.match(/>([\s\S]*?)</);
      
      if (hrefMatch && nameMatch) {
        const href = hrefMatch[1];
        let name = nameMatch[1].replace(/<[^>]+>/g, '').trim();
        name = name.replace(/^»?\s*\d+\s*/, '').trim(); // Pythonの正規表現をシミュレーション

        // Pythonコードから: if href and '/webclass/course.php/' in href:
        if (href.match(COURSE_LINK_REGEX)) { // Config.gs の正規表現を利用
          courseLinks.push({ name: name, href: href });
        }
      }
    });
  }

  return courseLinks;
}

/**
 * 特定の科目の課題データを取得し、解析する
 * @returns {{assignments: Array<Object>, html: string}}
 */
function fetchAndParseCourse(courseName, href) {
  let url = WEBCLASS_BASE_URL + href; // Config.gs の定数を利用
  
  let res = fetchWithSession(url); // Auth.gs の関数
  let html = res.getContentText();
  
  const redirectMatch = html.match(REDIRECT_REGEX); // Config.gs の正規表現を利用
  if (redirectMatch) {
    // リダイレクトURLが完全なURLの場合もあるため、パスとして結合せず、完全なURLを使用
    const redirectUrl = redirectMatch[1].replace(/&amp;/g, "&");
    
    // リダイレクトURLが相対パスの場合はWEBCLASS_BASE_URLを付加
    const finalUrl = redirectUrl.startsWith('http') ? redirectUrl : WEBCLASS_BASE_URL + redirectUrl;
    
    res = fetchWithSession(finalUrl);
    html = res.getContentText();
  }
  
  // HTMLコンテンツも一緒に返す
  return { 
      assignments: parseCourseContents(html), // Parser.gs の関数
      html: html
  };
}

/**
 * 受信した課題データをシートに書き込む（Tasks IDとフラグを保持しながらマージする）
 */
function writeMergedAssignmentsToSheet(sheet, newAssignments) {
    // 1. 初期設定としてヘッダーを書き込む
    if (sheet.getLastRow() === 0) {
        sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold'); // Config.gs の定数を利用
    }
  
    // 2. 既存データを読み込み、リンクをキーとするマップを作成
    const existingDataMap = new Map(); // Key: 課題リンク(URL), Value: [Tasks ID, 登録済みフラグ]
    const existingData = sheet.getLastRow() > 1 
        ? sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADER.length).getValues()
        : [];
        
    // 既存のTasks IDと登録済みフラグを保持
    existingData.forEach(row => {
        const link = row[4]; // E列: 課題リンク (URL)
        const tasksId = row[5]; // F列: Tasks ID
        const registeredFlag = row[6]; // G列: 登録済みフラグ
        if (link) {
            existingDataMap.set(link, [tasksId, registeredFlag]);
        }
    });

    const mergedAssignments = [];
  
    // 3. 新しい課題データに既存のTasks情報をマージ
    newAssignments.forEach(newRow => {
        const link = newRow[4]; 
        if (existingDataMap.has(link)) {
            const [tasksId, registeredFlag] = existingDataMap.get(link);
            newRow[5] = tasksId;
            newRow[6] = registeredFlag;
        } else {
            newRow[5] = '';
            newRow[6] = '';
        }
        mergedAssignments.push(newRow);
    });
  
    // 4. 既存のデータ行をクリア（ヘッダー行は残す）
    sheet.getRange(2, 1, sheet.getMaxRows(), HEADER.length).clearContent();
        
    // 5. 結合したデータを2行目から書き込む
    if (mergedAssignments.length > 0) {
        sheet.getRange(2, 1, mergedAssignments.length, mergedAssignments[0].length).setValues(mergedAssignments);
    }
    
    // 6. 最終更新日時をA1に記載
    sheet.getRange('A1').setValue('最終更新: ' + new Date().toLocaleString());
    SpreadsheetApp.flush();
}

/**
 * Google Classroomの課題を取得し、シートに書き込む (Classroom APIを利用)
 */
function fetchClassroomAssignments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_CLASSROOM); // Config.gs の定数を利用

  if (!sheet) {
    Logger.log('エラー: Classroom課題シートが見つかりません。');
    return;
  }

  // ヘッダーがなければ作成
  if (sheet.getLastRow() === 0) {
     sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold'); // Config.gs の定数を利用
  }
  
  // 既存データを読み込み、リンクをキーとするマップを作成
  const existingDataMap = new Map();
  const existingData = sheet.getLastRow() > 1 
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADER.length).getValues()
    : [];
    
  existingData.forEach(row => {
    const link = row[4];
    const tasksId = row[5];
    const registeredFlag = row[6];
    if (link) {
      existingDataMap.set(link, [tasksId, registeredFlag]);
    }
  });
  
  let allAssignments = [];

  try {
    const coursesResponse = Classroom.Courses.list({ courseStates: ['ACTIVE'], studentId: 'me' });
    const courses = coursesResponse.courses;
    
    if (!courses || courses.length === 0) return;

    courses.forEach(course => {
      const courseWorkResponse = Classroom.Courses.CourseWork.list(course.id, {
        courseWorkStates: ['PUBLISHED'] 
      });

      if (courseWorkResponse.courseWork) {
        courseWorkResponse.courseWork.forEach(assignment => {
          const dueDate = assignment.dueDate ? getFormattedDueDate(assignment) : '期限なし';
          if (assignment.alternateLink && assignment.title) {
             allAssignments.push([
               'Classroom', course.name, assignment.title, dueDate, assignment.alternateLink, '', ''
             ]);
          }
        });
      }
    });

    if (allAssignments.length > 0) {
      const mergedAssignments = [];
      
      allAssignments.forEach(newRow => {
        const link = newRow[4];
        
        if (existingDataMap.has(link)) {
          const [tasksId, registeredFlag] = existingDataMap.get(link);
          newRow[5] = tasksId;
          newRow[6] = registeredFlag;
        } else {
          newRow[5] = '';
          newRow[6] = '';
        }
        mergedAssignments.push(newRow);
      });
      
      sheet.getRange(2, 1, sheet.getMaxRows(), HEADER.length).clearContent();
      sheet.getRange(2, 1, mergedAssignments.length, mergedAssignments[0].length).setValues(mergedAssignments);
      Logger.log(`${SHEET_NAME_CLASSROOM}に${mergedAssignments.length}件の課題を書き込みました。`);
    }
    
    sheet.getRange('A1').setValue('最終更新: ' + new Date().toLocaleString());
    SpreadsheetApp.flush();

  } catch (e) {
    Logger.log('Classroom課題取得中にエラーが発生しました: ' + e.toString());
  }
}

// --------------------------------------------------------------------

// --- Tasks完了状態の同期関数 ---

function syncTaskCompletionStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = [SHEET_NAME_WEBCLASS, SHEET_NAME_CLASSROOM];
  const taskListId = getTaskListId(TASK_LIST_NAME);

  sheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return;

    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADER.length);
    const data = range.getValues();
    let updated = false;

    data.forEach((row, index) => {
      const tasksId = row[5];
      const registeredFlag = row[6];

      if (tasksId && registeredFlag !== 'COMPLETED' && registeredFlag !== 'DELETED') { 
        try {
          const task = Tasks.Tasks.get(taskListId, tasksId);

          if (task.status === 'completed') {
            data[index][6] = 'COMPLETED';
            updated = true;
          }
        } catch (e) {
          if (e.toString().includes('notFound')) {
            data[index][6] = 'DELETED';
            updated = true;
          }
        }
      }
    });

    if (updated) {
      range.setValues(data);
      Logger.log(`${sheetName} のTasks完了状態を同期しました。`);
    }
  });
}

// --------------------------------------------------------------------

// --- 課題の統合と登録関数 ---

function integrateAndRegisterTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const webclassSheet = ss.getSheetByName(SHEET_NAME_WEBCLASS);
  const classroomSheet = ss.getSheetByName(SHEET_NAME_CLASSROOM);
  
  if (!webclassSheet || !classroomSheet) return;
  
  if (webclassSheet.getLastRow() === 0) {
     webclassSheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold');
  }

  const taskListId = getTaskListId(TASK_LIST_NAME); 

  const webclassData = getDataFromSheet(webclassSheet);
  const classroomData = getDataFromSheet(classroomSheet);
  
  const allAssignments = [...webclassData, ...classroomData];
  const assignmentsToRegister = filterAssignmentsToRegister(allAssignments);

  if (assignmentsToRegister.length === 0) {
    Logger.log('新しくTasksに登録すべき課題はありませんでした。');
    return;
  }
  
  registerAssignmentsToTasks(assignmentsToRegister, taskListId, ss);
}

// --------------------------------------------------------------------

// --- ヘルパー関数群 ---

function getDataFromSheet(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; 
  return sheet.getRange(2, 1, lastRow - 1, HEADER.length).getValues();
}

function filterAssignmentsToRegister(assignments) {
  const unique = [];
  const registeredLinks = new Set();

  assignments.forEach(row => {
    const dueDateStr = row[3];
    const link = row[4];
    const registered = row[6];
    
    if (registered) return; 

    const dueDate = new Date(dueDateStr);
    const now = new Date();
    if (isNaN(dueDate.getTime()) || dueDate < now) return; 

    if (!registeredLinks.has(link)) { 
      unique.push(row);
      registeredLinks.add(link);
    }
  });
  return unique;
}

function registerAssignmentsToTasks(assignments, taskListId, ss) {
  assignments.forEach(row => {
    const [source, courseName, title, dueDateStr, link] = [row[0], row[1], row[2], row[3], row[4]];
    const date = new Date(dueDateStr);
    
    const task = Tasks.newTask();
    task.title = `[${courseName} / ${source}] ${title}`; 
    task.notes = `課題リンク:\n${link}`;
    
    if (date instanceof Date && !isNaN(date)) {
      const rfc3339 = date.toISOString().split('T')[0] + 'T00:00:00.000Z';
      task.due = rfc3339;
    }
    
    try {
      const registeredTask = Tasks.Tasks.insert(task, taskListId);
      Logger.log(`タスク登録成功: ${title} (Tasks ID: ${registeredTask.id})`);
      
      setTaskRegisteredFlag(ss, source, link, registeredTask.id);

    } catch (e) {
      Logger.log(`タスク登録エラー (${title}): ${e.toString()}`);
    }
  });
}

function setTaskRegisteredFlag(ss, source, link, taskId) {
    const sheetName = source === 'WebClass' ? SHEET_NAME_WEBCLASS : SHEET_NAME_CLASSROOM;
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return;

    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADER.length);
    const data = dataRange.getValues();
    
    for (let i = 0; i < data.length; i++) {
        if (data[i][4] === link) { 
            data[i][5] = taskId;
            data[i][6] = 'TRUE';
            
            sheet.getRange(i + 2, 1, 1, HEADER.length).setValues([data[i]]);
            return;
        }
    }
}

function cleanupOldAssignments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = [SHEET_NAME_WEBCLASS, SHEET_NAME_CLASSROOM];
  const today = new Date();
  const retentionDays = 60;
  const retentionThreshold = new Date(today.setDate(today.getDate() - retentionDays));

  sheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return;

    const maxRows = sheet.getLastRow();
    const data = sheet.getRange(2, 1, maxRows - 1, HEADER.length).getValues();
    const rowsToDelete = [];

    for (let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      const dueDate = new Date(row[3]);
      const statusFlag = row[6];
      
      const isCompleted = statusFlag === 'COMPLETED';
      const isDeleted = statusFlag === 'DELETED';
      const isTooOld = !isNaN(dueDate.getTime()) && dueDate < retentionThreshold;

      if (isCompleted || isDeleted || isTooOld) {
        rowsToDelete.push(i + 2);
      }
    }

    rowsToDelete.sort((a, b) => b - a).forEach(rowIndex => {
      sheet.deleteRow(rowIndex);
    });

    Logger.log(`${sheetName} から ${rowsToDelete.length} 行の古いデータを削除しました。`);
  });
}

function getTaskListId(taskListName) {
    const lists = Tasks.Tasklists.list().getItems()
    let taskListId
    for (let i = 0; i < lists.length; i++) {
        if (lists[i].title == taskListName) {
            taskListId = lists[i].id
            break; 
        }
    }
    
    if (!taskListId) {
        throw new Error(`タスクリスト "${taskListName}" が見つかりませんでした。`);
    }
    
    return taskListId
}

function getFormattedDueDate(assignment) {
  const due = assignment.dueDate;
  
  if (!due || !due.year) return '期限なし';
  
  const hour = assignment.dueTime ? assignment.dueTime.hours : 23;
  const minute = assignment.dueTime ? assignment.dueTime.minutes || 59 : 59;
  
  const date = new Date(due.year, due.month - 1, due.day, hour, minute);
  
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
}

/**
 * エラー発生時のHTTPレスポンスを例外を無視して強制取得する
 */
function fetchErrorResponseHtml(url) {
    try {
        const options = {
            'muteHttpExceptions': true, 
        };
        const res = UrlFetchApp.fetch(url, options);
        return `ステータスコード: ${res.getResponseCode()}\n\n` + res.getContentText();
    } catch (e) {
        return `レスポンス取得中に別のエラーが発生しました: ${e.message}`;
    }
}


// --- 診断ログ関数（Code.gs内に定義し、確実な書き込みを保証） ---

/**
 * 診断ログシートにHTMLコンテンツを書き込み、強制的にFlushする（C3セルにHTML）
 */
function logHtmlForDiagnosis(htmlContent, title) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "GAS診断ログ"; 
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    // シートが存在しない場合は作成
    sheet = ss.insertSheet(sheetName);
    sheet.getRange('A1').setValue('診断ログシート');
    sheet.getRange('A2').setValues([['最終更新:', '診断タイトル', 'HTMLコンテンツ', '処理ステータス']]);
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 800); 
    sheet.setColumnWidth(4, 300); 
  }
  
  // HTML (C3) に書き込む
  sheet.getRange('A3').setValue(new Date().toLocaleString());
  sheet.getRange('B3').setValue(title);
  sheet.getRange('C3').setValue(htmlContent);
  
  sheet.getRange('A2').setValue('最終更新: ' + new Date().toLocaleString()); // ヘッダー日付更新
  SpreadsheetApp.flush(); // 強制的に書き込みを完了させる
  Logger.log(`診断ログ（HTML）をシート「${sheetName}」に出力しました: ${title}`);
}

/**
 * 診断ログシートに段階的なステータスを書き込み、強制的にFlushする（D列にステータス）
 */
function logDiagnosisStatus(statusText, rowNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "GAS診断ログ"; 
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    // シートがない場合はまず作成
    sheet = ss.insertSheet(sheetName);
    sheet.getRange('A1').setValue('診断ログシート');
    sheet.getRange('A2').setValues([['最終更新:', '診断タイトル', 'HTMLコンテンツ', '処理ステータス']]);
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 800); 
    sheet.setColumnWidth(4, 300); 
  }
  
  // D列の指定行にステータスを追記
  sheet.getRange(rowNum, 4).setValue(`${new Date().toLocaleTimeString()}: ${statusText}`);
  SpreadsheetApp.flush(); // 強制的に書き込みを完了させる
  Logger.log(`診断ステータス（D${rowNum}）を更新: ${statusText}`);
}

// Code.gs に追加 (既存のコードに追加してください)

/**
 * ログシートにタイムスタンプ付きでメッセージを記録する
 */
function logToSheet(message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // ログシートが存在しない場合は作成（名前は 'ログ'）
  let logSheet = ss.getSheetByName('ログ');
  if (!logSheet) {
    logSheet = ss.insertSheet('ログ');
    logSheet.appendRow(['タイムスタンプ', 'メッセージ']);
  }
  
  // スプレッドシートに追記
  logSheet.appendRow([new Date(), message]);
  // 最新のログが見えるようにスクロール
  logSheet.getRange(logSheet.getLastRow(), 1).activate();
}

/**
 * ログシートをクリアする
 */
function clearLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('ログ');
  if (logSheet) {
    // ヘッダー行を除いてデータをクリア
    logSheet.deleteRows(2, logSheet.getLastRow() - 1);
    SpreadsheetApp.getUi().alert('ログシートをクリアしました。');
  } else {
    SpreadsheetApp.getUi().alert('「ログ」シートが見つかりませんでした。');
  }
}