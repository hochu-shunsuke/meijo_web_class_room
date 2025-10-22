/**
 * WebClassの学籍番号とパスワードをPropertiesServiceに保存する
 * Settings.html から google.script.run 経由で呼び出されます。
 * @param {string} userid - 学籍番号
 * @param {string} password - パスワード
 */
function setCredentials(userid, password) {
  if (!userid || !password) {
    // HTML側の withFailureHandler にエラーを返す
    throw new Error('学籍番号とパスワードの両方を入力してください。'); 
  }
  
  // ユーザープロパティに保存
  PropertiesService.getUserProperties()
      .setProperties({
          'userid': userid,
          'password': password
      });
  
  // 戻り値がないため、HTML側の withSuccessHandler が呼ばれます。
}

/**
 * 保存された認証情報を取得する
 * @returns {Object|null} 認証情報 {userid: string, password: string} または null
 */
function getCredentials() {
  const properties = PropertiesService.getUserProperties().getProperties();
  
  if (properties.userid && properties.password) {
    return {
      userid: properties.userid,
      password: properties.password
    };
  }
  return null;
}