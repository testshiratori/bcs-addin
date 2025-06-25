Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook && Office.context.mailbox?.item) {
    // document.getElementById("btn-send-kintone").onclick = async () => {
    //   try {
    //     // const accessToken = await OfficeRuntime.auth.getAccessToken({
    //     //   allowSignInPrompt: true,
    //     //   forMSGraphAccess: true
    //     // });

    //     await startPolling();  // ← トークンを渡して処理実行
    //   } catch (e) {
    //     console.error("SSOエラー:", e);
    //   }
    // };
    document.getElementById("btn-send-kintone").onclick = startAuthFlowAndAddContact;
  } else {
    console.warn("アイテムコンテキストが無いため、SSOは使用できません");
  }
});

async function startPolling() {
  const intervalMs = 10 * 1000;//5 * 60 * 1000; // 5分

  setInterval(async () => {
    try {
      // SSOトークン取得（Microsoft 365のユーザーである必要あり）
      //const accessToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true,forMSGraphAccess: true });

      // 自分の userId を取得（@より前）+ 特例処理
      const email = Office.context.mailbox.userProfile.emailAddress;
      const userIdRaw = email.split("@")[0];
      const matchUserId = userIdRaw === "ito-mitsuyuki" ? email : userIdRaw;

      // SharePoint REST APIでフィルタ付きリスト取得
      const listUrl = `https://shiratoripharm.sharepoint.com/_api/web/lists/getbytitle('trn_card_fetch_status')/items?$filter=(user_id eq '${matchUserId}' and is_fetched ne 1)`;

      const spResponse = await fetch(listUrl, {
        method: "GET",
        headers: {
          "Accept": "application/json;odata=verbose",
          "Authorization": `Bearer ${accessToken}`
        }
      });

      if (!spResponse.ok) throw new Error(`SharePoint応答エラー: ${spResponse.status}`);

      const spData = await spResponse.json();
      const items = spData.d.results;

      if (items.length > 0) {
        console.log("対象データあり:", items);

        // Graph APIでメール送信（例）
        const graphResponse = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${accessToken}`
          },
          body: JSON.stringify({
            message: {
              subject: "未取得のカードデータがあります",
              body: {
                contentType: "Text",
                content: `次のレコードが未取得です（件数: ${items.length}）。最初のID: ${items[0].ID}`
              },
              toRecipients: [
                { emailAddress: { address: email } } // 自分に送信
              ]
            }
          })
        });

        if (!graphResponse.ok) {
          console.warn("Graph API 呼び出し失敗", await graphResponse.text());
        }
      }

    } catch (error) {
      console.error("ポーリングエラー:", error);
    }
  }, intervalMs);
}

function base64URLEncode(str) {
  return str.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

async function sha256(buffer) {
  const digest = await crypto.subtle.digest('SHA-256', buffer);
  return new Uint8Array(digest);
}

async function generatePKCE() {
  const randomBytes = crypto.getRandomValues(new Uint8Array(32));
  const code_verifier = base64URLEncode(btoa(String.fromCharCode(...randomBytes)));
  const challenge = await sha256(new TextEncoder().encode(code_verifier));
  const code_challenge = base64URLEncode(btoa(String.fromCharCode(...challenge)));
  return { code_verifier, code_challenge };
}

async function startAuthFlowAndAddContact() {
  const tenantId = "c7202a3e-8ddf-4149-ba61-30915b2b6188";
  const clientId = "d33ca1e9-0900-4a00-a7c7-634127a47e5d";
  const redirectUri = "https://white-forest-07ab38200.1.azurestaticapps.net/auth-callback.html";
  const scope = "https://graph.microsoft.com/Contacts.ReadWrite offline_access";

  try{
    const { code_verifier, code_challenge } = await generatePKCE();

    console.log("認証開始");
    const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize` +
      `?client_id=${clientId}` +
      `&response_type=code` +
      `&redirect_uri=${encodeURIComponent(redirectUri)}` +
      `&response_mode=query` +
      `&scope=${encodeURIComponent(scope)}` +
      `&code_challenge=${code_challenge}&code_challenge_method=S256`;

    Office.context.ui.displayDialogAsync(authUrl, { height: 60, width: 30 }, (asyncResult) => {
      console.log("認証終了");
      const dialog = asyncResult.value;
      console.log("ダイアログ表示");
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
        dialog.close();
        const authCode = arg.message;
        const tokenRes = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: new URLSearchParams({
            client_id: clientId,
            grant_type: "authorization_code",
            code: authCode,
            redirect_uri: redirectUri,
            code_verifier: code_verifier
          })
        });
        const tokenJson = await tokenRes.json();
        const accessToken = tokenJson.access_token;

        console.log("連絡先追加");
        // ★ TEST 連絡先を追加
        const res = await fetch("https://graph.microsoft.com/v1.0/me/contacts", {
          method: "POST",
          headers: {
            "Authorization": `Bearer ${accessToken}`,
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            givenName: "TEST",
            surname: "User",
            emailAddresses: [{ address: "test@example.com", name: "TEST User" }],
            companyName: "Test Co"
          })
        });

        if (res.ok) {
          console.log("連絡先を追加しました");
        } else {
          console.error("連絡先追加失敗", await res.text());
        }
      });
    });
  }
  catch(error){
    console.error("連絡先追加処理エラー:", error);
  }

  
}