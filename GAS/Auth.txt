// Auth.gs (SAMLResponse完全クリーンアップ版 - 最終確認版)

/**
 * WebClassのセッションを管理するためのグローバル変数（Cookieを保持）
 */
let WEBCLASS_SESSION_COOKIES = {};

// リダイレクト追跡の最大回数
const MAX_REDIRECTS = 15;

// --- ヘルパー関数群 ---

/**
 * URLから明示的なポート番号 :443 を除去する（SSO認証のクッキー不具合対策）
 */
const stripPort443 = (url) => {
    if (url && url.includes(':443/')) {
        logToSheet(`[WARN] URLからポート番号 :443 を除去: ${url.replace(':443/', '/')}`);
        return url.replace(':443/', '/');
    }
    return url;
};

/**
 * HTMLエンティティでエンコードされた文字列をデコードする (簡易版)
 * ※SAML Requestを破損させないよう、URLデコードは行わない
 * @param {string} text エンコードされた文字列
 * @returns {string} デコードされた文字列
 */
const decodeHtmlEntities = (text) => {
    if (!text) return text;
    // URLに使われることが多いエンティティをデコード
    let decoded = text.replace(/&#x3a;/g, ':'); // :
    decoded = decoded.replace(/&#x2f;/g, '/'); // /
    decoded = decoded.replace(/&amp;/g, '&');   // &
    decoded = decoded.replace(/&quot;/g, '"'); // "
    decoded = decoded.replace(/&gt;/g, '>'); // >
    decoded = decoded.replace(/&lt;/g, '<');   // <
    return decoded;
};

/**
 * Base64文字列を完全にクリーンアップする（SAML POST用）
 * (中略)
 */
const cleanBase64String = (str) => {
    if (!str) return '';
    
    // 1. 数値エンティティを手動でデコード（&#xNN; 形式）
    let cleaned = str.replace(/&#x([0-9a-fA-F]+);/g, (match, hex) => {
        return String.fromCharCode(parseInt(hex, 16));
    });
    
    // 2. 数値エンティティをデコード（&#NNN; 形式）
    cleaned = cleaned.replace(/&#(\d+);/g, (match, dec) => {
        return String.fromCharCode(parseInt(dec, 10));
    });
    
    // 3. 名前付きエンティティをデコード
    cleaned = cleaned.replace(/&amp;/g, '&');
    cleaned = cleaned.replace(/&quot;/g, '"');
    cleaned = cleaned.replace(/&lt;/g, '<');
    cleaned = cleaned.replace(/&gt;/g, '>');
    
    // 4. 制御文字、空白文字、改行を完全除去
    cleaned = cleaned.replace(/[\x00-\x1F\x7F-\x9F\s]/g, '');
    
    // 5. Base64文字以外を完全除去
    cleaned = cleaned.replace(/[^A-Za-z0-9+/=]/g, '');
    
    // 6. パディングの正規化（=は末尾のみ）
    const base64Core = cleaned.replace(/=/g, '');
    const paddingNeeded = (4 - (base64Core.length % 4)) % 4;
    cleaned = base64Core + '='.repeat(paddingNeeded);
    
    // 7. 最終検証
    if (!/^[A-Za-z0-9+/]*={0,2}$/.test(cleaned)) {
        logToSheet(`[ERROR] Invalid Base64 format detected: ${cleaned.substring(0, 100)}`);
        throw new Error('Base64 validation failed after cleaning');
    }
    
    return cleaned;
};

/**
 * RelayState文字列をクリーンアップする
 * (中略)
 */
const cleanRelayState = (str) => {
    if (!str) return '';
    
    // 1. 数値エンティティを手動でデコード（&#xNN; 形式）
    let cleaned = str.replace(/&#x([0-9a-fA-F]+);/g, (match, hex) => {
        return String.fromCharCode(parseInt(hex, 16));
    });
    
    // 2. 数値エンティティをデコード（&#NNN; 形式）
    cleaned = cleaned.replace(/&#(\d+);/g, (match, dec) => {
        return String.fromCharCode(parseInt(dec, 10));
    });
    
    // 3. 名前付きエンティティをデコード
    cleaned = cleaned.replace(/&amp;/g, '&');
    cleaned = cleaned.replace(/&quot;/g, '"');
    cleaned = cleaned.replace(/&lt;/g, '<');
    cleaned = cleaned.replace(/&gt;/g, '>');
    
    // 4. 改行と不要な空白を除去
    cleaned = cleaned.replace(/[\r\n]/g, '').trim();
    
    return cleaned;
};

/**
 * URLFetchAppのレスポンスからセッションクッキーを抽出して保持する (強制適用)
 */
const setSessionCookies = (res) => {
    if (!res || typeof res.getAllHeaders !== 'function') return;
    const cookies = res.getAllHeaders()['Set-Cookie'];
    if (cookies) {
        const cookieArray = Array.isArray(cookies) ? cookies : [cookies];
        cookieArray.forEach(cookieStr => {
            const parts = cookieStr.split(';');
            if (parts[0].includes('=')) {
                const [name, value] = parts[0].split('=').map(s => s.trim());
                // 【重要】PathやDomainを無視して、名前と値のみを保持
                WEBCLASS_SESSION_COOKIES[name] = value;
            }
        });
    }
};

/**
 * リクエストヘッダーを構築し、保持しているセッションクッキーを追加する (User-Agent強化)
 * (中略)
 */
const buildRequestHeaders = (url) => {
    // Config.gsのUSER_AGENTSが存在しない場合のフォールバック
    const commonUserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36';
    // USER_AGENTSはConfig.gsで定義されているものとする
    const userAgents = (typeof USER_AGENTS !== 'undefined' && Array.isArray(USER_AGENTS) && USER_AGENTS.length > 0) 
                       ? USER_AGENTS 
                       : [commonUserAgent];
    const userAgent = userAgents[Math.floor(Math.random() * userAgents.length)];

    const headers = {
        'User-Agent': userAgent,
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'ja,en-US;q=0.7,en;q=0.3',
        'Referer': url 
    };
    
    // 保持しているすべてのクッキーを送信
    const cookieString = Object.keys(WEBCLASS_SESSION_COOKIES)
        .map(name => `${name}=${WEBCLASS_SESSION_COOKIES[name]}`)
        .join('; ');
    if (cookieString) {
        headers['Cookie'] = cookieString;
    }
    
    return headers;
};

/**
 * UrlFetchApp.fetchをラップし、通信ログを記録する
 * (中略)
 */
const fetchWrapper = (url, options) => {
    const decodedUrl = decodeHtmlEntities(url);
    const defaultOptions = {
        'method': 'get',
        'headers': buildRequestHeaders(decodedUrl),
        'muteHttpExceptions': true,
        'followRedirects': options && options.followRedirects === true 
    };
    const mergedOptions = Object.assign({}, defaultOptions, options);
    const isAutoRedirect = mergedOptions.followRedirects === true;
    
    try {
        const samlFormRegex = /<input type="hidden" name="SAMLResponse" value="([^"]+)"/;
        
        logToSheet(`[REQ] ${mergedOptions.method.toUpperCase()}: ${decodedUrl} (Redirects: ${isAutoRedirect ? 'AUTO' : 'MANUAL/NONE'})`);
        logToSheet(`[SENT COOKIES] ${JSON.stringify(WEBCLASS_SESSION_COOKIES)}`); 

        const res = UrlFetchApp.fetch(decodedUrl, mergedOptions);
        setSessionCookies(res); 

        const statusCode = res.getResponseCode();
        const locationHeader = res.getAllHeaders()['Location'] || 'N/A';
     
        let samlFormFound = 'NO';
        
        if (statusCode === 200) {
            const body = res.getContentText();
            if (samlFormRegex.test(body)) {
                samlFormFound = 'YES (SAML POST)';
            } else if (body.includes('XUI/#login')) {
                samlFormFound = 'SSO_LOGIN (SPA)';
            }
        }
        
        logToSheet(`[RES] Status: ${statusCode}, Location: ${locationHeader}, SAML Form: ${samlFormFound}`);
        return res;
    } catch (e) {
        logToSheet(`[ERROR] UrlFetchApp.fetch実行エラー: ${decodedUrl} - ${e.message}`);
        throw new Error(`UrlFetchApp.fetch実行エラー: ${decodedUrl} - ${e.message}`);
    }
};

/**
 * 相対URLを絶対URLに修正する
 */
const correctUrl = (url, base) => {
    if (!url || url.startsWith('http')) return url;
    const correctedBase = base.endsWith('/') ? base.slice(0, -1) : base;
    if (url.startsWith('/')) {
        return correctedBase + url;
    }
    return correctedBase + '/' + url;
};

/**
 * 手動でリダイレクトチェーンを追跡し、最終的な200 OKレスポンスと最終URLを返す
 * @param {string} startUrl - SAMLフローの開始URL
 * @param {string} ssoBaseUri - SSOサーバーのベースURL
 * @param {string} webclassBaseUri - WebClassサーバーのベースURL
 * @returns {{response: GoogleAppsScript.URL_Fetch.HTTPResponse, finalUrl: string}} WebClassダッシュボードの最終レスポンスと最終URL
 */
function followManualRedirects(startUrl, ssoBaseUri, webclassBaseUri) {
    let currentUrl = startUrl;
    let res;
    let samlPostData = null; // SAML POSTデータ保持用 (文字列)
    let samlAcsUrl = '';
    
    for (let i = 0; i < MAX_REDIRECTS; i++) {
        logToSheet(`-> リダイレクト追跡 [${i + 1}/${MAX_REDIRECTS}]: ${currentUrl}`);
        
        // SAML POSTが必要な場合
        if (samlPostData) {
            logToSheet(`🚨 SAML POSTフォームを検出。POSTを実行します...`);
            const postOptions = {
                'method': 'post', 
                'contentType': 'application/x-www-form-urlencoded', 
                'payload': samlPostData, // 手動で構築した文字列
                'followRedirects': false 
            };
            res = fetchWrapper(samlAcsUrl, postOptions);
            samlPostData = null; // POST後はデータをクリア

        } else {
            // 通常のGETリクエスト（リダイレクトを自動追跡しない）
            res = fetchWrapper(currentUrl, { 'method': 'get', 'followRedirects': false });
        }

        // 追跡終了条件1: 200 OK (リダイレクト終了)
        if (res.getResponseCode() === 200) {
            const body = res.getContentText();
            
            // WebClassのダッシュボードに到達したか確認 (タイトルタグでの判定を強化)
            const isDashboard = body.includes('<title>コースリスト - WebClass</title>') || 
                                body.includes('cl-courseList_courseLink') || 
                                body.includes('cl-user_name') || 
                                body.includes('/webclass/course/list.php');

            if (isDashboard) {
                logToSheet(`✅ 追跡終了: WebClassダッシュボードに到達 (Status 200)`);
                return { response: res, finalUrl: currentUrl }; 
            }
            
            // JavaScriptリダイレクトの検出（SAML認証成功後）
            const jsRedirectMatch = body.match(/window\.location\.href\s*=\s*["']([^"']+)["']/);
            if (jsRedirectMatch) {
                let redirectPath = jsRedirectMatch[1];
                logToSheet(`✓ JavaScript リダイレクトを検出: ${redirectPath}`);
                
                // 相対パスを絶対URLに変換
                if (!redirectPath.startsWith('http')) {
                    redirectPath = correctUrl(redirectPath, webclassBaseUri);
                }
                
                // リダイレクト先に移動
                currentUrl = redirectPath;
                continue;
            }
            
            // 200 OKだがSAML POSTフォームの場合
            const samlPostMatch = body.match(/<input type="hidden" name="SAMLResponse" value="([^"]+)"/);

            if (samlPostMatch) {
                logToSheet('✓ SAML POSTフォームを検出しました。次のリクエストでPOSTします。');
                
                let samlResponse = cleanBase64String(samlPostMatch[1]);
                const relayStateMatch = body.match(/<input type="hidden" name="RelayState" value="([^"]+)"/);
                let relayState = cleanRelayState(relayStateMatch ? relayStateMatch[1] : '');
                const formActionMatch = body.match(/<form method="post" action="([^"]+)"/);
                let extractedAcsUrl = formActionMatch ? formActionMatch[1] : ACS_URL;
                samlAcsUrl = decodeHtmlEntities(extractedAcsUrl); 
                
                if (!samlAcsUrl.startsWith('http')) {
                    samlAcsUrl = correctUrl(samlAcsUrl, webclassBaseUri);
                }

                const encodedSamlResponse = encodeURIComponent(samlResponse);
                const encodedRelayState = encodeURIComponent(relayState);

                samlPostData = `SAMLResponse=${encodedSamlResponse}&RelayState=${encodedRelayState}`;
                
                currentUrl = res.getHeaders()['Location'] || currentUrl; 
                continue; // 次のループでPOST実行
            }
            
            // 200 OKだが、想定外のページの場合
            logHtmlForDiagnosis(res.getContentText(), `SAML追跡失敗時最終200OK（ホップ${i+1}）`);
            
            // 追跡ホップ数が一定回数を超えていれば（SAMLフローが終わっていれば）、正常終了と見なして返す
            if (i >= 5) { 
                logToSheet(`⚠️ 成功判定はFalseでしたが、SAMLフロー完了後の200 OKページと見なし、処理を続行します。`);
                return { response: res, finalUrl: currentUrl }; 
            }
            
            throw new Error(`SAMLフロー追跡中に200 OKの想定外ページに到達しました。診断ログを確認してください。`);
        }

        // 追跡終了条件2: リダイレクトがない
        let nextLocation = res.getAllHeaders()['Location'];
        if (!nextLocation) {
            logHtmlForDiagnosis(res.getContentText(), `SAML追跡失敗時Locationヘッダーなし（ホップ${i+1}）`);
            throw new Error(`リダイレクトチェーンが途切れました。最終ステータス: ${res.getResponseCode()}、最終URL: ${currentUrl}`);
        }
        
        // リダイレクト先に移動
        nextLocation = decodeHtmlEntities(nextLocation);
        nextLocation = stripPort443(nextLocation); // ポート除去を適用
        const base = nextLocation.includes(webclassBaseUri) ? webclassBaseUri : ssoBaseUri;
        currentUrl = correctUrl(nextLocation, base);
    }
    
    // 最大追跡回数を超えた
    throw new Error(`リダイレクト追跡が最大回数（${MAX_REDIRECTS}回）を超えました。ループの可能性があります。`);
}

// --- ログインフロー本体 ---

/**
 * 認証情報を使用してWebClassのセッションを確立する
 * @param {string} userid - ユーザーID
 * @param {string} password - パスワード
 * @returns {string} 確立されたセッションのWebClassダッシュボードURL
 */
function loginWebClass(userid, password) {
    logToSheet('--- WebClassログイン処理開始 (SAMLResponse完全クリーンアップ版) ---');
    WEBCLASS_SESSION_COOKIES = {}; 
    
    // Config.gsの定数に依存
    const baseUri = WEBCLASS_BASE_URL.replace(/\/$/, '');
    const ssoBaseUriMatch = SSO_URL.match(/^https?:\/\/[^\/]+/);
    const ssoBaseUri = ssoBaseUriMatch ? ssoBaseUriMatch[0] : '';
    
    let res;

    // ステップ1〜3: 認証情報のPOSTとiPlanetDirectoryProクッキー取得 (中略)
    res = fetchWrapper(SSO_URL, { 'method': 'post' });
    let authResponse = JSON.parse(res.getContentText());
    const authId = authResponse.authId;
    
    const authPayload = {
        'authId': authId,
        'callbacks': [
            {'type': 'NameCallback', 'output': [{'name': 'prompt', 'value': 'ユーザー名:'}], 'input': [{'name': 'IDToken1', 'value': userid}]},
            {'type': 'PasswordCallback', 'output': [{'name': 'prompt', 'value': 'パスワード:'}], 'input': [{'name': 'IDToken2', 'value': password}], 'echoPassword': false}
        ]
    };
    const jsonOptions = {'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(authPayload)};
    res = fetchWrapper(SSO_URL, jsonOptions);
    authResponse = JSON.parse(res.getContentText());
    
    if (!authResponse.tokenId && !authResponse.successUrl) {
        let reason = authResponse.reason || authResponse.message || authResponse.errorMessage || '認証失敗';
        throw new Error(`WebClassログインに失敗しました。SSO認証エラー: ${reason}`);
    }
    
    if (authResponse.tokenId) {
        WEBCLASS_SESSION_COOKIES['iPlanetDirectoryPro'] = authResponse.tokenId;
    }
    logToSheet(`3. 認証成功！ iPlanetDirectoryPro クッキーを手動注入しました。`);

    // ステップ4: SAMLフローの完全手動追跡
    const LOGIN_URL = WEBCLASS_BASE_URL + '/webclass/login.php?auth_mode=SAML';
    logToSheet(`4. WebClassセッション確定のため LOGIN_URL にアクセス (SAMLフロー完全手動追跡)...`);
    
    let finalResult; 
    try {
        finalResult = followManualRedirects(LOGIN_URL, ssoBaseUri, baseUri);
    } catch(e) {
        logToSheet('🚨 SAMLフローの追跡に失敗しました。ログと診断シートを確認してください。');
        throw new Error(`ステップ4 (SAMLフロー) 追跡失敗: ${e.message}`);
    }
    
    // 最終チェック
    const finalUrl = finalResult.finalUrl; // 修正済みのロジック: 追跡が停止した最終URLを返す
    
    logToSheet(`✅ WebClassセッション確立成功！最終URL: ${finalUrl}`);
    return finalUrl;
}

// --- セッション付きフェッチ関数 ---

/**
 * 確立されたセッションクッキーを使用してWebClassのページを取得する
 * @param {string} url - 取得するWebClassのURL
 * @param {number} [redirectCount=0] - リダイレクトの追跡回数 (再帰用)
 * @returns {string} 取得したHTMLコンテンツ
 * * Config.gsの WEBCLASS_BASE_URL, REDIRECT_REGEX, MAX_REDIRECTS に依存します。
 */
function fetchWithSession(url, redirectCount = 0) {
    // 無限ループ防止のための追跡回数チェック
    if (redirectCount >= MAX_REDIRECTS) {
        logToSheet(`🚨 リダイレクトが最大追跡回数（${MAX_REDIRECTS}回）を超えました。強制終了します。URL: ${url}`);
        throw new Error(`リダイレクトが最大追跡回数（${MAX_REDIRECTS}回）を超えました。`);
    }

    const options = {
        'method': 'get',
        // headersとcookiesはグローバルな setSessionCookies で処理される
        'headers': buildRequestHeaders(url), 
        'muteHttpExceptions': true, 
        'followRedirects': false // リダイレクトは自前で処理する
    };
    
    let res;
    try {
        res = fetchWrapper(url, options);
    } catch (e) {
        throw new Error(`WebClassアクセスエラー: ${url} のリクエストに失敗しました（UrlFetchAppエラー: ${e.message}）。`);
    }

    // レスポンスからセッションクッキーを更新
    setSessionCookies(res);
    
    const statusCode = res.getResponseCode();
    
    // 1. HTTPリダイレクト（302, 301）の処理
    if (statusCode === 302 || statusCode === 301) {
        const nextUrl = res.getHeaders()['Location'];
        logToSheet(`[HTTP REDIRECT ${statusCode}] URLを ${nextUrl} へ追跡中...`);
        return fetchWithSession(nextUrl, redirectCount + 1); // 再帰呼び出し
    } 

    // 2. JavaScriptリダイレクト（200 OK + <script>）の検出と処理
    if (statusCode >= 200 && statusCode < 300) {
        const html = res.getContentText();
        
        // REDIRECT_REGEX は Config.gs で定義されていると仮定
        const redirectMatch = html.match(REDIRECT_REGEX);
        
        // HTMLが小さく（リダイレクトスクリプトの典型）、かつリダイレクトパターンが見つかった場合
        if (redirectMatch && html.length < 500) { 
            const relativePath = redirectMatch[1];
            
            // WEBCLASS_BASE_URL は Config.gs で定義されていると仮定
            // 絶対URLに変換 (baseUrl + relativePath)
            const baseUrl = WEBCLASS_BASE_URL.replace(/\/$/, '');
            const newUrl = baseUrl + (relativePath.startsWith('/') ? relativePath : '/' + relativePath);
            
            logToSheet(`[JS REDIRECT] ${url} から ${newUrl} へJavaScriptリダイレクトを追跡中...`);
            Logger.log(`[JS REDIRECT] URLを ${newUrl} に変更して再アクセスします。`);
            
            return fetchWithSession(newUrl, redirectCount + 1); // 再帰呼び出し
        }
        
        // 通常のHTMLコンテンツ（課題リストなど）が取得された
        return html;
    }
    
    // 3. その他のエラー
    logToSheet(`🚨 WebClassアクセスエラー: ${url} (Status ${statusCode})。セッション切れの可能性があります。`);
    throw new Error(`WebClassアクセス失敗: ${url} (Status ${statusCode} - ${res.getContentText().substring(0, 100)}...)`);
}