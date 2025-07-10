Office.onReady(() => {
  const btn = document.getElementById("btn-test-dialog");
  if (btn) {
    console.log("✅ ボタン見つかりました");
    btn.onclick = runAuthFlow;
  } else {
    console.error("❌ btn-test-dialog ボタンが見つかりません");
  }
});

// 固定情報（アプリ登録内容に応じて書き換えてください）
const tenantId    = "c7202a3e-8ddf-4149-ba61-30915b2b6188";
const clientId    = "d33ca1e9-0900-4a00-a7c7-634127a47e5d";
const redirectUri = "https://white-forest-07ab38200.1.azurestaticapps.net/auth.html";
const scope       = "openid profile email offline_access User.Read";
const responseType = "code";
const responseMode = "query";

// PKCEに必要なコードベリファイアとチャレンジを生成
let code_verifier = "";
let code_challenge = "";

/**
 * 認証フロー実行：ダイアログ表示 → コード受信 → トークン取得
 */
async function runAuthFlow() {
  console.log("✅ runAuthFlow 開始"); // ← これが出ればイベント発火成功
  // PKCEコード生成
  await generatePKCE();

  // 認可URL作成
  const authUrl =
    "https://login.microsoftonline.com/" + tenantId + "/oauth2/v2.0/authorize" +
    `?client_id=${clientId}` +
    `&response_type=${responseType}` +
    `&redirect_uri=${encodeURIComponent(redirectUri)}` +
    `&response_mode=${responseMode}` +
    `&scope=${encodeURIComponent(scope)}` +
    `&code_challenge=${code_challenge}` +
    `&code_challenge_method=S256` +
    `&state=12345`;

  // セッションストレージに認可URLを保存（auth.html で利用）
  sessionStorage.setItem("authUrl", authUrl);

  const encodedAuthUrl = encodeURIComponent(authUrl);
  // const dialogUrl = authUrl;
  const dialogUrl = `https://white-forest-07ab38200.1.azurestaticapps.net/auth.html?authUrl=${encodedAuthUrl}`;
  console.log("表示するURL: ", dialogUrl);

  // 認証ダイアログを auth.html 経由で表示
  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 50, width: 50 },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("❌ ダイアログ失敗:", asyncResult.error.message);
        return;
      }

      console.log("✅ ダイアログ表示成功");
      const dialog = asyncResult.value;

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
        dialog.close();
        const authCode = arg.message;
        console.log("認可コードを受信:", authCode);
        await exchangeCodeForToken(authCode);
      });
    }
  );
}

/**
 * 認可コードをアクセストークンに交換
 */
async function exchangeCodeForToken(authCode) {
  const body = new URLSearchParams({
    client_id: clientId,
    grant_type: "authorization_code",
    code: authCode,
    redirect_uri: redirectUri,
    code_verifier: code_verifier
  });

  try {
    const tokenRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body
    });

    const tokenJson = await tokenRes.json();
    if (tokenJson.access_token) {
      console.log("✅ アクセストークン取得成功:", tokenJson.access_token);

      // ↓ Graph API 呼び出しなどの後続処理を書く
      await callGraphApi(tokenJson.access_token);

    } else {
      console.error("❌ トークン取得エラー:", tokenJson);
    }
  } catch (err) {
    console.error("❌ トークンリクエスト失敗:", err);
  }
}

/**
 * Microsoft Graph API 呼び出し（例）
 */
async function callGraphApi(accessToken) {
  const res = await fetch("https://graph.microsoft.com/v1.0/me", {
    headers: {
      Authorization: `Bearer ${accessToken}`
    }
  });
  const profile = await res.json();
  console.log("✅ ユーザープロファイル:", profile);
}

/**
 * PKCE用のコードベリファイア＆チャレンジを生成
 */
async function generatePKCE() {
  code_verifier = base64URLEncode(crypto.getRandomValues(new Uint8Array(32)));
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(code_verifier));
  code_challenge = base64URLEncode(new Uint8Array(digest));
}

// Base64URLエンコード
function base64URLEncode(buffer) {
  return btoa(String.fromCharCode.apply(null, buffer))
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}


// Office.onReady((info) => {
//   if (info.host === Office.HostType.Outlook && Office.context.mailbox?.item) {
//     // document.getElementById("btn-send-kintone").onclick = async () => {
//     //   try {
//     //     // const accessToken = await OfficeRuntime.auth.getAccessToken({
//     //     //   allowSignInPrompt: true,
//     //     //   forMSGraphAccess: true
//     //     // });

//     //     // await startPolling();  // ← トークンを渡して処理実行
//     //     //await TestNinshou();
//     //   } catch (e) {
//     //     console.error("SSOエラー:", e);
//     //   }
//     // };
    
//     document.getElementById("btn-send-kintone").onclick = startAuthFlowAndAddContact;

//     document.getElementById("btn-test-dialog").onclick = () => {
//       // Office.context.ui.displayDialogAsync("https://white-forest-07ab38200.1.azurestaticapps.net/test.html", {
//       //   height: 50,
//       //   width: 50
//       // }, function (asyncResult) {
//       //   if (asyncResult.status === Office.AsyncResultStatus.Failed) {
//       //     console.error("ダイアログ失敗:", asyncResult.error.message);
//       //   } else {
//       //     console.log("テストダイアログ成功");
//       //   }
//       // });

//       // 認可エンドポイント情報
//       const tenantId = "c7202a3e-8ddf-4149-ba61-30915b2b6188"; // 例: "common" でもOK
//       const clientId = "d33ca1e9-0900-4a00-a7c7-634127a47e5d";
//       const redirectUri = "https://white-forest-07ab38200.1.azurestaticapps.net/auth.html";
//       const scope = "openid profile email offline_access User.Read"; // 適宜修正
//       const responseType = "code";
//       const responseMode = "query"; // auth.html で受け取れる形式

//       // 認可URL生成
//       const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize` +
//         `?client_id=${clientId}` +
//         `&response_type=${responseType}` +
//         `&redirect_uri=${encodeURIComponent(redirectUri)}` +
//         `&response_mode=${responseMode}` +
//         `&scope=${encodeURIComponent(scope)}` +
//         `&state=12345`; // stateは任意


//       Office.context.ui.displayDialogAsync("https://white-forest-07ab38200.1.azurestaticapps.net/auth.html", {
//         height: 60, width: 60
//       }, function (asyncResult) {
//         if (asyncResult.status === Office.AsyncResultStatus.Failed) {
//         console.error("ダイアログ表示失敗：", asyncResult.error.message);
//         } else {
//             const dialog = asyncResult.value;
//             dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
//                 const data = JSON.parse(args.message);
//                 console.log("アクセストークン:", data.token);
//                 dialog.close();
//             });
//         }
//       });
//     };
//   } else {
//     console.warn("アイテムコンテキストが無いため、SSOは使用できません");
//   }
// });

// // async function TestNinshou() {
// //   //const qs = require('querystring');

// //   const tenant = 'c7202a3e-8ddf-4149-ba61-30915b2b6188';
// //   const clientId = 'f1b7e6ae-5f2b-4a56-ae10-7c6379fa65fb';
// //   const clientSecret = '2ud8Q~XQ1czEn0cp~h_p3FQOzHwtn9vBSWVNqbU_'; // 必須
// //   //900f6a2c-6dd7-492e-9a92-a7de3510e662
// //   const username = 'testshiratori@shiratoripharm.onmicrosoft.com';
// //   const password = 'shiraPass02';

// //   const tokenUrl = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`;
// //   const body = new URLSearchParams({
// //     client_id: clientId,
// //     scope: 'https://graph.microsoft.com/.default',
// //     username: username,
// //     password: password,
// //     grant_type: 'password',
// //     client_secret: clientSecret
// //   }).toString();
// //   // const body = qs.stringify({
// //   //   client_id: clientId,
// //   //   scope: 'https://graph.microsoft.com/.default',
// //   //   username: username,
// //   //   password: password,
// //   //   grant_type: 'password',
// //   //   client_secret: clientSecret
// //   // });

// //   fetch(tokenUrl, {
// //     method: 'POST',
// //     headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
// //     body: body
// //   })
// //   .then(res => res.json())
// //   .then(data => {
// //     console.log(data.access_token);
// //   })
// //   .catch(err => console.error(err));
// // }

// async function startPolling() {
//   const intervalMs = 10 * 1000;//5 * 60 * 1000; // 5分

//   setInterval(async () => {
//     try {
//       // SSOトークン取得（Microsoft 365のユーザーである必要あり）
//       //const accessToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true,forMSGraphAccess: true });

//       // 自分の userId を取得（@より前）+ 特例処理
//       const email = Office.context.mailbox.userProfile.emailAddress;
//       const userIdRaw = email.split("@")[0];
//       const matchUserId = userIdRaw === "ito-mitsuyuki" ? email : userIdRaw;

//       // SharePoint REST APIでフィルタ付きリスト取得
//       const listUrl = `https://shiratoripharm.sharepoint.com/_api/web/lists/getbytitle('trn_card_fetch_status')/items?$filter=(user_id eq '${matchUserId}' and is_fetched ne 1)`;

//       const spResponse = await fetch(listUrl, {
//         method: "GET",
//         headers: {
//           "Accept": "application/json;odata=verbose",
//           "Authorization": `Bearer ${accessToken}`
//         }
//       });

//       if (!spResponse.ok) throw new Error(`SharePoint応答エラー: ${spResponse.status}`);

//       const spData = await spResponse.json();
//       const items = spData.d.results;

//       if (items.length > 0) {
//         console.log("対象データあり:", items);

//         // Graph APIでメール送信（例）
//         const graphResponse = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
//           method: "POST",
//           headers: {
//             "Content-Type": "application/json",
//             "Authorization": `Bearer ${accessToken}`
//           },
//           body: JSON.stringify({
//             message: {
//               subject: "未取得のカードデータがあります",
//               body: {
//                 contentType: "Text",
//                 content: `次のレコードが未取得です（件数: ${items.length}）。最初のID: ${items[0].ID}`
//               },
//               toRecipients: [
//                 { emailAddress: { address: email } } // 自分に送信
//               ]
//             }
//           })
//         });

//         if (!graphResponse.ok) {
//           console.warn("Graph API 呼び出し失敗", await graphResponse.text());
//         }
//       }

//     } catch (error) {
//       console.error("ポーリングエラー:", error);
//     }
//   }, intervalMs);
// }

// function base64URLEncode(str) {
//   return str.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
// }

// function base64URLEncodeFromBuffer(buffer) {
//   return btoa(String.fromCharCode(...buffer))
//     .replace(/\+/g, '-')
//     .replace(/\//g, '_')
//     .replace(/=+$/, '');
// }

// async function sha256(buffer) {
//   const digest = await crypto.subtle.digest('SHA-256', buffer);
//   return new Uint8Array(digest);
// }

// // async function generatePKCE() {
// //   const randomBytes = crypto.getRandomValues(new Uint8Array(32));
// //   const code_verifier = base64URLEncode(btoa(String.fromCharCode(...randomBytes)));
// //   const challenge = await sha256(new TextEncoder().encode(code_verifier));
// //   const code_challenge = base64URLEncode(btoa(String.fromCharCode(...challenge)));
// //   return { code_verifier, code_challenge };
// // }
// async function generatePKCE() {
//   const randomBytes = new Uint8Array(32);
//   crypto.getRandomValues(randomBytes);

//   const code_verifier = base64URLEncodeFromBuffer(randomBytes);
//   const challengeBuffer = await sha256(new TextEncoder().encode(code_verifier));
//   const code_challenge = base64URLEncodeFromBuffer(new Uint8Array(challengeBuffer));

//   return { code_verifier, code_challenge };
// }

// async function startAuthFlowAndAddContact() {
//   const tenantId = "c7202a3e-8ddf-4149-ba61-30915b2b6188";
//   const clientId = "d33ca1e9-0900-4a00-a7c7-634127a47e5d";
//   const redirectUri = "https://white-forest-07ab38200.1.azurestaticapps.net/auth-callback.html";
//   const scope = "https://graph.microsoft.com/Contacts.ReadWrite offline_access";

//   try{
//     const { code_verifier, code_challenge } = await generatePKCE();

//     console.log("認証開始");
//     const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize` +
//       `?client_id=${clientId}` +
//       `&response_type=code` +
//       `&redirect_uri=${encodeURIComponent(redirectUri)}` +
//       `&response_mode=query` +
//       `&scope=${encodeURIComponent(scope)}` +
//       `&code_challenge=${code_challenge}&code_challenge_method=S256`;


//     Office.context.ui.displayDialogAsync(authUrl, {
//       height: 50,
//       width: 50
//     }, function (asyncResult) {
//       if (asyncResult.status === Office.AsyncResultStatus.Failed) {
//         console.error("❌ ダイアログ失敗:", asyncResult.error.message);
//         alert("ダイアログ失敗: " + asyncResult.error.message);
//       } else {
//         console.log("✅ ダイアログ成功");
//         const dialog = asyncResult.value;
        
//         console.log("ダイアログ表示");
//         dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
//           console.log("認証コード受信:", arg.message);  // ← ここに移動
//           dialog.close();
//           const authCode = arg.message;
//           const tokenRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
//             method: "POST",
//             headers: { "Content-Type": "application/x-www-form-urlencoded" },
//             body: new URLSearchParams({
//               client_id: clientId,
//               grant_type: "authorization_code",
//               code: authCode,
//               redirect_uri: redirectUri,
//               code_verifier: code_verifier
//             })
//           });
//           const tokenJson = await tokenRes.json();
//           const accessToken = tokenJson.access_token;

//           console.log("連絡先追加");
//           // ★ TEST 連絡先を追加
//           const res = await fetch("https://graph.microsoft.com/v1.0/me/contacts", {
//             method: "POST",
//             headers: {
//               "Authorization": `Bearer ${accessToken}`,
//               "Content-Type": "application/json"
//             },
//             body: JSON.stringify({
//               givenName: "TEST",
//               surname: "User",
//               emailAddresses: [{ address: "test@example.com", name: "TEST User" }],
//               companyName: "Test Co"
//             })
//           });

//           if (res.ok) {
//             console.log("連絡先を追加しました");
//           } else {
//             console.error("連絡先追加失敗", await res.text());
//           }
//         });
//       }
//     });

//     // Office.context.ui.displayDialogAsync(authUrl, { height: 60, width: 30 }, (asyncResult) => {
//     //   console.log("認証終了");
//     //   const dialog = asyncResult.value;
//     //   console.log("ダイアログ表示");
//     //   dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
//     //     console.log("認証コード受信:", arg.message);  // ← ここに移動
//     //     dialog.close();
//     //     const authCode = arg.message;
//     //     const tokenRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
//     //       method: "POST",
//     //       headers: { "Content-Type": "application/x-www-form-urlencoded" },
//     //       body: new URLSearchParams({
//     //         client_id: clientId,
//     //         grant_type: "authorization_code",
//     //         code: authCode,
//     //         redirect_uri: redirectUri,
//     //         code_verifier: code_verifier
//     //       })
//     //     });
//     //     const tokenJson = await tokenRes.json();
//     //     const accessToken = tokenJson.access_token;

//     //     console.log("連絡先追加");
//     //     // ★ TEST 連絡先を追加
//     //     const res = await fetch("https://graph.microsoft.com/v1.0/me/contacts", {
//     //       method: "POST",
//     //       headers: {
//     //         "Authorization": `Bearer ${accessToken}`,
//     //         "Content-Type": "application/json"
//     //       },
//     //       body: JSON.stringify({
//     //         givenName: "TEST",
//     //         surname: "User",
//     //         emailAddresses: [{ address: "test@example.com", name: "TEST User" }],
//     //         companyName: "Test Co"
//     //       })
//     //     });

//     //     if (res.ok) {
//     //       console.log("連絡先を追加しました");
//     //     } else {
//     //       console.error("連絡先追加失敗", await res.text());
//     //     }
//     //   });
//     // });
//   }
//   catch(error){
//     console.error("連絡先追加処理エラー:", error);
//   }

  
// }