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
 * ★修正点: Settings.html ファイルを読み込み、ダイアログを表示する
 */
function showCredentialDialog() {
  // Settings.htmlファイルを読み込む
  const htmlOutput = HtmlService.createHtmlOutputFromFile('Settings')
      .setWidth(450)
      .setHeight(250); // サイズを調整
      
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'WebClass認証情報の設定');
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
        const taskListId = getTaskListId(taskListName); // Code.gs の関数を呼び出し
        
        // taskListIdをPropertiesServiceに保存
        PropertiesService.getUserProperties().setProperty('taskListId', taskListId);
        
        ui.alert(`✅ Google Tasks連携設定完了。\n利用タスクリスト: 「${taskListName}」に決定しました。\n初回実行時、Googleアカウントのアクセス許可が必要です。`);
        
    } catch (e) {
        ui.alert(`🚨 Tasksリスト設定エラー:\n${e.message}\n\nTasksサービスが有効化されているか、Config.gsのTASK_LIST_NAMEが正しいか確認してください。`);
    }
}