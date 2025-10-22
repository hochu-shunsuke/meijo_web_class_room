// Auth.gs (最終版 - ポート除去再導入と手動追跡)

/**
 * WebClassのセッションを管理するためのグローバル変数（Cookieを保持）
 */
let WEBCLASS_SESSION_COOKIES = {};

// --- 新しいヘルパー関数 ---
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

// --- ヘルパー関数定義 ---
// setSessionCookies, buildRequestHeaders, fetchWrapper, correctUrl は省略せず全て含まれている必要があります

/**
 * URLFetchAppのレスポンスからセッションクッキーを抽出して保持する
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
                WEBCLASS_SESSION_COOKIES[name] = value;
            }
        });
    }
};

/**
 * リクエストヘッダーを構築し、保持しているセッションクッキーを追加する
 */
const buildRequestHeaders = (url) => {
    // USER_AGENTSのフォールバックを追加
    const userAgents = typeof USER_AGENTS !== 'undefined' ? USER_AGENTS : ['Mozilla/5.0 (Windows NT 10.0; Win64; x64)'];
    const userAgent = userAgents[Math.floor(Math.random() * userAgents.length)];

    const headers = {
        'User-Agent': userAgent,
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'ja,en-US;q=0.7,en;q=0.3',
        'Referer': url 
    };
    
    const cookieString = Object.keys(WEBCLASS_SESSION_COOKIES)
        .map(name => `${name}=${WEBCLASS_SESSION_COOKIES[name]}`)
        .join('; ');
        
    if (cookieString) {
        headers['Cookie'] = cookieString;
    }
    
    return headers;
};

/**
 * UrlFetchApp.fetchをラップし、**すべての通信ログ**を記録する
 */
const fetchWrapper = (url, options) => {
    const defaultOptions = {
        'method': 'get',
        'headers': buildRequestHeaders(url),
        'muteHttpExceptions': true,
        'followRedirects': false 
    };
    
    const mergedOptions = Object.assign({}, defaultOptions, options);
    const followRedirects = mergedOptions.followRedirects === true;

    try {
        logToSheet(`[REQ] ${mergedOptions.method.toUpperCase()}: ${url} (Redirects: ${followRedirects ? 'AUTO' : 'MANUAL'})`);
        
        const res = UrlFetchApp.fetch(url, mergedOptions);
        setSessionCookies(res); 

        const statusCode = res.getResponseCode();
        const locationHeader = res.getAllHeaders()['Location'] || 'N/A';
        let samlFormFound = 'NO';
        
        if (statusCode === 200) {
            const body = res.getContentText();
            if (body.includes('SAMLResponse') && body.includes('RelayState')) {
                samlFormFound = 'YES';
            } else if (body.includes('XUI/#login')) {
                 samlFormFound = 'SSO_LOGIN';
            }
        }
        
        logToSheet(`[RES] Status: ${statusCode}, Location: ${locationHeader}, SAML Form: ${samlFormFound}`);
        
        return res;
    } catch (e) {
        logToSheet(`[ERROR] UrlFetchApp.fetch実行エラー: ${url} - ${e.message}`);
        throw new Error(`UrlFetchApp.fetch実行エラー: ${url} - ${e.message}`);
    }
};

// URL補正ヘルパー
const correctUrl = (url, base) => {
    if (!url || url.startsWith('http')) return url;
    const correctedBase = base.endsWith('/') ? base.slice(0, -1) : base;
    if (url.startsWith('/')) {
        return correctedBase + url;
    }
    return correctedBase + '/' + url;
};

// --- ログインフロー本体 ---

function loginWebClass(userid, password) {
    logToSheet('--- WebClassログイン処理開始 (最終版 - ポート除去再導入) ---');
    WEBCLASS_SESSION_COOKIES = {}; 
    const baseUri = WEBCLASS_BASE_URL.replace(/\/$/, '');
    
    const ssoBaseUriMatch = SSO_URL.match(/^https?:\/\/[^\/]+/);
    const ssoBaseUri = ssoBaseUriMatch ? ssoBaseUriMatch[0] : '';
    if (!ssoBaseUri) {
         throw new Error("SSO_URLからベースURLを抽出できませんでした。");
    }
    
    let currentUrl; 
    let res;

    // ステップ1: 初回アクセスでauthIdとコールバック構造を取得
    logToSheet(`1. SSO初回アクセス...`);
    res = fetchWrapper(SSO_URL, { 'method': 'post' });
    let authResponse;
    try {
        authResponse = JSON.parse(res.getContentText());
    } catch (e) {
        throw new Error("SSO初回応答のパースに失敗しました。");
    }
    if (!authResponse.authId) {
        throw new Error("SSO初回応答にauthIdがありません。");
    }
    const authId = authResponse.authId;
    logToSheet(`✓ authId取得成功: ${authId.substring(0, 50)}...`);
    
    // ステップ2: コールバック形式でユーザー名/パスワードを送信
    logToSheet(`2. 認証情報をPOST中...`);
    const authPayload = {
        'authId': authId,
        'callbacks': [
            {'type': 'NameCallback', 'output': [{'name': 'prompt', 'value': 'ユーザー名:'}], 'input': [{'name': 'IDToken1', 'value': userid}]},
            {'type': 'PasswordCallback', 'output': [{'name': 'prompt', 'value': 'パスワード:'}], 'input': [{'name': 'IDToken2', 'value': password}], 'echoPassword': false}
        ]
    };
    const jsonOptions = {'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(authPayload)};
    res = fetchWrapper(SSO_URL, jsonOptions);
    try {
        authResponse = JSON.parse(res.getContentText());
    } catch (e) {
        throw new Error("SSO認証応答のパースに失敗しました。");
    }
    let isSuccess = false;
    if (authResponse.tokenId || authResponse.successUrl) {
        isSuccess = true;
        currentUrl = authResponse.successUrl;
    }
    if (!isSuccess) {
        let reason = authResponse.reason || authResponse.message || authResponse.errorMessage || '認証失敗';
        throw new Error(`WebClassログインに失敗しました。SSO認証エラー: ${reason}`);
    }
    logToSheet(`3. 認証成功！ successUrl: ${currentUrl}`);

    // ステップ3: successUrl追跡（まだSSOドメイン）
    if (currentUrl) {
        currentUrl = correctUrl(currentUrl, ssoBaseUri); 
        
        logToSheet(`4. successUrl追跡開始: ${currentUrl}`);
        res = fetchWrapper(currentUrl, { 'method': 'get' });
        
        let locationFromSuccessUrl = res.getAllHeaders()['Location'];
        if (locationFromSuccessUrl) {
            // ポート除去を適用
            locationFromSuccessUrl = stripPort443(locationFromSuccessUrl); 
            currentUrl = correctUrl(locationFromSuccessUrl, ssoBaseUri);
            logToSheet(`-> リダイレクト(Location): ${currentUrl}`);
        }
    }


    // ステップ4/5: WebClassセッション確定のためのアクセス (SAMLフォーム取得追跡)
    const LOGIN_URL = WEBCLASS_BASE_URL + '/webclass/login.php?auth_mode=SAML';
    logToSheet(`5. WebClassセッション確定のため LOGIN_URL にアクセス: ${LOGIN_URL}`);
    
    res = fetchWrapper(LOGIN_URL, { 'method': 'get' });
    
    let samlResponseHtml = '';
    let nextLocation = res.getAllHeaders()['Location'];
    currentUrl = LOGIN_URL; 

    // Locationヘッダーがある場合、リダイレクト追跡を開始
    if (nextLocation) {
        // ポート除去を適用
        nextLocation = stripPort443(nextLocation);
        currentUrl = correctUrl(nextLocation, nextLocation.includes(baseUri) ? baseUri : ssoBaseUri);
        
        let maxSAMLRedirects = 10;
        while (maxSAMLRedirects-- > 0) {
            logToSheet(`-> SAMLフロー追跡: ${currentUrl}`);
            res = fetchWrapper(currentUrl, { 'method': 'get' });

            // ステータスコードが200ならHTMLコンテンツをチェック
            if (res.getResponseCode() === 200) {
                samlResponseHtml = res.getContentText();
                // 認証が完了したSSOからのSAMLフォームを返すページか？
                if (samlResponseHtml.includes('SAMLResponse') && samlResponseHtml.includes('RelayState')) {
                    logToSheet(`✓ SAML Responseフォームを発見！`);
                    break; 
                }
            }

            nextLocation = res.getAllHeaders()['Location'];
            
            if (nextLocation) {
                // ポート除去を適用
                nextLocation = stripPort443(nextLocation);
                // リダイレクト先をSSOドメインかWebClassドメインかで適切に結合
                currentUrl = correctUrl(nextLocation, nextLocation.includes(baseUri) ? baseUri : ssoBaseUri);
            } else {
                 // リダイレクトが途切れ、かつSAMLフォームがない場合は追跡終了
                 break;
            }
        }
    }

    if (!samlResponseHtml || !samlResponseHtml.includes('SAMLResponse')) {
        throw new Error("SAML Responseフォームの取得に失敗しました。ログを確認してください。");
    }

    // ステップ6: SAML ResponseをパースしてWebClass ACSにPOST
    logToSheet(`6. SAML Responseをパースし、WebClass ACSにPOST中...`);

    const samlMatch = samlResponseHtml.match(/name="SAMLResponse" value="([^"]+)"/);
    const relayMatch = samlResponseHtml.match(/name="RelayState" value="([^"]+)"/);
    const formActionMatch = samlResponseHtml.match(/<form method="post" action="([^"]+)"/); 

    if (!samlMatch || !relayMatch || !formActionMatch) {
        throw new Error("SAML ResponseまたはRelayState、またはPOST先URLのパースに失敗しました。");
    }

    const samlResponse = samlMatch[1];
    const relayState = relayMatch[1];
    const finalAcsPath = formActionMatch[1]; 

    const samlPayload = {
        'SAMLResponse': samlResponse,
        'RelayState': relayState
    };

    const samlPostOptions = {
        'method': 'post',
        'payload': samlPayload,
        'followRedirects': false, 
        'headers': {
            'User-Agent': buildRequestHeaders(currentUrl)['User-Agent'], 
            'Referer': currentUrl 
        }
    };
    
    const correctedAcsUrl = finalAcsPath.startsWith('http') ? finalAcsPath : correctUrl(finalAcsPath, baseUri);
    logToSheet(`-> ACS URLへPOST: ${correctedAcsUrl}`);

    res = fetchWrapper(correctedAcsUrl, samlPostOptions);

    // ステップ7: 最終的なダッシュボードへのリダイレクト追跡
    logToSheet(`7. SAML POST成功。最終リダイレクトを追跡中... (ステータス: ${res.getResponseCode()})`);
    
    let maxRedirects = 10;
    currentUrl = correctedAcsUrl;

    while (maxRedirects-- > 0) {
        let nextLocation = res.getAllHeaders()['Location'];
        
        if (res.getResponseCode() === 200) {
            const body = res.getContentText();
            const jsRedirectMatch = body.match(/window\.location\.href\s*=\s*['"]([^'"]+)['"]/);

            if (jsRedirectMatch) {
                let jsRedirectUrl = jsRedirectMatch[1].replace(/&amp;/g, '&');
                let finalDashboardUrl = jsRedirectUrl.startsWith('http') ? jsRedirectUrl : baseUri + jsRedirectUrl;
                logToSheet(`✅ WebClassセッション確立成功 (JSリダイレクト)。最終URL: ${finalDashboardUrl}`);
                return finalDashboardUrl;
            }

            if (currentUrl.includes(baseUri) && body.includes('cl-courseList_courseLink')) {
                logToSheet(`✅ WebClassセッション確立成功。最終URL: ${currentUrl}`);
                return currentUrl;
            }
            
            throw new Error("SAML認証は成功しましたが、WebClassのダッシュボードに到達できませんでした。");

        } else if (res.getResponseCode() >= 300 && res.getResponseCode() < 400 && nextLocation) {
            currentUrl = correctUrl(nextLocation, baseUri);
            logToSheet(`-> 最終リダイレクト追跡: ${currentUrl}`);
            res = fetchWrapper(currentUrl, { 'method': 'get' });
            continue;
        } else {
            throw new Error(`最終リダイレクト追跡に失敗。ステータス: ${res.getResponseCode()}、URL: ${currentUrl}`);
        }
    }

    throw new Error("最終リダイレクトループを検出しました。セッション確立失敗。");
}

// --- セッション付きフェッチ関数（変更なし） ---

function fetchWithSession(url) {
    const options = {
        'method': 'get',
        'headers': buildRequestHeaders(url),
        'muteHttpExceptions': true, 
        'followRedirects': false
    };

    let res;
    try {
        res = fetchWrapper(url, options); 
    } catch (e) {
        throw new Error(`WebClassアクセスエラー: ${url} のリクエストに失敗しました（UrlFetchAppエラー: ${e.message}）。`);
    }

    setSessionCookies(res);
    
    const statusCode = res.getResponseCode();
    
    if (statusCode >= 400 && statusCode !== 404) {
        const errorText = res.getContentText();
        throw new Error(`WebClassアクセスエラー: ${url} のリクエストに失敗しました（エラー: ${statusCode}）。サーバー応答の一部: ${errorText.substring(0, 100)}`);
    }

    return res;
}