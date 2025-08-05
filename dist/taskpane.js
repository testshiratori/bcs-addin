Office.onReady(() => {
  const btn = document.getElementById("btn-send-kintone");
  if (btn) {
    console.log("✅ ボタン見つかりました");
    btn.onclick = runAuthFlow;
  } else {
    console.error("❌ btn-send-kintone ボタンが見つかりません");
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

  await fetchCardStatusForCurrentUser(accessToken,profile.userPrincipalName);
}

async function fetchCardStatusForCurrentUser(accessToken, userPrincipalName) {
  const siteHostname = "shiratoripharm.sharepoint.com";
  const sitePath = "/sites/" + encodeURIComponent("コミュニケーションサイト");
  // const sitePath = "/sites/コミュニケーションサイト";
  const listName = "trn_card_fetch_status";

  // プリンシパルIDの@より前を抽出
  const userId = userPrincipalName.split("@")[0];

  // // サイトIDを取得
  // const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteHostname}:${sitePath}`, {
  //   headers: { Authorization: `Bearer ${accessToken}` }
  // });
  // const siteJson = await siteRes.json();
  // const siteId = siteJson.id;

  // // リストIDを取得
  // const listRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listName}`, {
  //   headers: { Authorization: `Bearer ${accessToken}` }
  // });
  // const listJson = await listRes.json();
  // const listId = listJson.id;

  const siteId = "320b6b44-40b6-4265-b314-afdba6eb20ba"; // サイトID（同時に取得されている）
  const listId = "e1999f46-92c8-4978-bb3c-0f5826e7143f"; // リストID（今わかった値）

  console.log("アクセストークン:",accessToken);
  // リストアイテムを取得（filter）
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?$expand=fields`;
  // const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?$filter=fields/user_id eq 'naoto-fujiwara'`;

  console.log("最終リクエストURL:", url);
  console.log("使用トークン:", accessToken);

  const itemsRes = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`
    }
  });

  const itemsJson = await itemsRes.json();
  console.log("対象アイテム:", itemsJson.value);

  itemsJson.value.forEach(item => {
    const raw = item.fields?.user_id ?? '';
    const norm = String(raw).normalize('NFKC').trim().toLowerCase();
    // console.log('user_id:', raw, '| normalized:', norm, '| codes:', Array.from(norm).map(c => c.charCodeAt(0)));
  });

  // フィルタ処理（user_idが一致し、is_fetchedがfalse）
  const targetUserId = 'naoto-fujiwara';

  const filteredItems = itemsJson.value.filter(item => {
    const rawUserId = String(item.fields?.user_id ?? '');
    const normUserId = rawUserId.normalize('NFKC').trim().toLowerCase();
    const userMatch = normUserId === targetUserId;

    // const rawFetched = item.fields?.is_fetched;
    // const fetchedMatch = rawFetched === false || rawFetched === 'false';

    // console.log(
    //   `[Check] user_id: "${rawUserId}" → "${normUserId}" | match: ${userMatch}`,
    //   `| is_fetched: ${rawFetched} | match: ${fetchedMatch}`
    // );

    // return userMatch && fetchedMatch;
    return userMatch;
  }); 

  console.log("対象アイテム:", filteredItems);

  // cardListItems: 最初のリスト（trn_card_fetch_status など）から取得した複数のレコード
  const cardListItems = filteredItems; // ← 例：filteredItems や取得済みの配列

  const personListId = "591f8714-ffdc-4787-82ce-8f4be141504c";
  // meishiPersonList: trn_meishi_person から取得した全レコード
  const personUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${personListId}/items?$expand=fields`;
  const personRes = await fetch(personUrl, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
  const personJson = await personRes.json();
  console.log("meishiPersonリスト取得結果：", personJson);

  // value が無い場合は空配列にする
  const personArray = personJson.value || [];

  // person_id でマッチするデータを1つずつ取得
  const mergedResults = cardListItems.map(cardItem => {
    const cardId = String(cardItem.fields?.card_id).trim();

    // person_idと一致する名刺情報を検索
    const person = personJson.value.find(personItem =>
      String(personItem.fields?.person_id).trim() === cardId
    );

    return {
      card: cardItem,
      person: person || null
    };
  });

  console.log("結合済みレコード一覧:", mergedResults);

  await addContactsToBCSFolder(accessToken, mergedResults);

  return itemsJson.value;
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

async function addContactsToBCSFolder(accessToken, personList) {
  // Step 1: 「BCS」フォルダIDを取得
  const folderListRes = await fetch("https://graph.microsoft.com/v1.0/me/contactFolders", {
    headers: {
      "Authorization": `Bearer ${accessToken}`
    }
  });

  if (!folderListRes.ok) {
    console.error("フォルダ一覧取得失敗:", await folderListRes.text());
    return;
  }

  const folderList = await folderListRes.json();
  const bcsFolder = folderList.value.find(f => f.displayName === "BCS");

  if (!bcsFolder) {
    console.error("BCSフォルダが見つかりません");
    return;
  }

  const folderId = bcsFolder.id;

  // Step 2: 各personデータを順番に追加
  for (const item of personList) {
    const person = item.person;
    const card = item.card;

    if (!person || !person.fields?.full_name) {
      console.warn("無効なデータのためスキップ:", item);
      continue;
    }

    const f = person.fields;
    const cf = card.fields;

    if(!cf.old_id || cf.old_id.trim() === "")
    {
      if(item.is_fetched === false){
        const res = await fetch(`https://graph.microsoft.com/v1.0/me/contactFolders/${folderId}/contacts`, {
          method: "POST",
          headers: {
            "Authorization": `Bearer ${accessToken}`,
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            givenName: f.full_name,
            companyName: f.company_name,
            jobTitle: f.position,
            department: f.department,
            businessPhones: [f.phone],
            emailAddresses: [
              {
                address: f.email,
                name: f.full_name
              }
            ],
            yomiGivenName:f.furigana
          })
        });

        if (res.ok) {
          console.log(`登録成功：${f.full_name}`);

          await updateIsFetchedTrue(accessToken, cf.id);
        } else {
          const error = await res.json();
          console.error(`登録失敗：${f.full_name}`, error);
        }
      }
    }else{
      try {
        // BCSフォルダ内の連絡先を取得
        const contactListRes = await fetch(`https://graph.microsoft.com/v1.0/me/contactFolders/${folderId}/contacts`, {
          headers: {
            "Authorization": `Bearer ${accessToken}`
          }
        });

        if (contactListRes.ok) {
          const contactList = await contactListRes.json();
          const oldContact = contactList.value.find(c => {
            const givenName = c.categories[0]?.toString().trim();
            const cfId = String(cf.card_id).trim();
            return givenName === cfId;
          });
          //メールアドレスで比較
          // const oldContact = contactList.value.find(c => {
          //   const emails = (c.emailAddresses || [])
          //     .map(e => (e?.address ? e.address.toLowerCase() : null))
          //     .filter(addr => addr !== null);

          //   return f.email && emails.includes(f.email.toLowerCase());
          // });
          // const oldContact = contactList.value.find(c => {
          //   const emails = c.emailAddresses?.map(e => e.address.toLowerCase()) || [];
          //   return emails.includes(cf.old_id.toLowerCase()); // 例：old_id をメールアドレスと一致とみなす場合
          // });

          if (oldContact) {
            // 削除
            const deleteRes = await fetch(`https://graph.microsoft.com/v1.0/me/contacts/${oldContact.id}`, {
              method: "DELETE",
              headers: {
                "Authorization": `Bearer ${accessToken}`
              }
            });

            if (deleteRes.ok) {
              console.log(`🗑️ 旧連絡先を削除しました: ${cf.card_id}`);
            } else {
              console.warn(`⚠️ 旧連絡先削除失敗: ${cf.card_id}`, await deleteRes.text());
            }
          } else {
            console.log(`ℹ️ 該当する旧連絡先が見つかりませんでした: ${cf.card_id}`);
          }
        } else {
          console.warn("⚠️ BCSフォルダの連絡先取得失敗", await contactListRes.text());
        }
      } catch (err) {
        console.error("❌ 旧連絡先削除処理でエラー:", err);
      }
    }
    

  }
}

async function updateIsFetchedTrue(accessToken, itemId) {
  const siteId = "320b6b44-40b6-4265-b314-afdba6eb20ba"; // サイトID（同時に取得されている）
  const listId = "e1999f46-92c8-4978-bb3c-0f5826e7143f";
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${itemId}/fields`;

  const res = await fetch(url, {
    method: "PATCH",
    headers: {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      is_fetched: true
    })
  });

  if (res.ok) {
    console.log(`✅ is_fetched を true に更新: ID=${itemId}`);
  } else {
    const err = await res.json();
    console.error(`❌ 更新失敗: ID=${itemId}`, err);
  }
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