// Menu.gs

/**
 * スプレッドシートを開いたときにカスタムメニューを作成する
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('✨ 課題自動取得システム')
      .addItem('1. 認証情報を設定', 'showCredentialDialog') 
      .addItem('2. Tasks連携設定', 'setupTasksList')       
      .addSeparator()
      .addItem('3. 今すぐ実行（テスト）', 'dailySystemRun')
      .addToUi();
}

/**
 * 認証情報入力用のカスタムダイアログを表示する
 * (Properties.gsのsaveCredentials関数に依存)
 */
function showCredentialDialog() {
  const html = HtmlService.createHtmlOutput('<div>WebClassのIDとパスワードをScript Propertiesに安全に保存します。<br><input type="text" id="userid" placeholder="ユーザーID"><br><input type="password" id="password" placeholder="パスワード"><br><button onclick="google.script.run.saveCredentials(document.getElementById(\'userid\').value, document.getElementById(\'password\').value); window.close();">保存</button></div>')
      .setWidth(300)
      .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, 'WebClass認証情報の設定');
}

/**
 * Google Tasks連携用のタスクリストIDを取得・設定する
 * (Code.gsのgetTaskListId関数に依存)
 */
function setupTasksList() {
    const ui = SpreadsheetApp.getUi();
    // TASK_LIST_NAME は Config.gs の定数
    const taskListName = TASK_LIST_NAME; 
    
    try {
        const taskListId = getTaskListId(taskListName); 
        ui.alert(`✅ タスクリスト「${taskListName}」のIDが正常に取得できました。\nタスクの登録が可能です。`);
    } catch (e) {
        ui.alert(`🚨 エラー: ${e.message}\nタスクリスト名（Config.gsのTASK_LIST_NAME）が正しいか、Tasks APIが有効になっているか確認してください。`);
    }
}