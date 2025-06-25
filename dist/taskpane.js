Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook && Office.context.mailbox?.item) {
    document.getElementById("btn-send-kintone").onclick = async () => {
      try {
        const accessToken = await OfficeRuntime.auth.getAccessToken({
          allowSignInPrompt: true,
          forMSGraphAccess: true
        });

        await startPolling(accessToken);  // ← トークンを渡して処理実行
      } catch (e) {
        console.error("SSOエラー:", e);
      }
    };
  } else {
    console.warn("アイテムコンテキストが無いため、SSOは使用できません");
  }
});

async function startPolling(accessToken) {
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
