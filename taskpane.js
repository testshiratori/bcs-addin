// taskpane.js
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("btn-send-kintone").onclick = async () => {
      const company = document.getElementById("txtCompany").value.trim();
      const project = document.getElementById("txtProject").value.trim();
      const phase = document.getElementById("txtPhase").value.trim();

      try {
        const item = Office.context.mailbox.item;
        const subject = item.subject;
        const date = new Date().toISOString(); // 実際は item.dateTimeCreated 等を使っても良い
        const userId = Office.context.mailbox.userProfile.emailAddress.split("@")[0] === "ito-mitsuyuki"
          ? "ito-mitsuyuki@shiratori-pharm.co.jp"
          : Office.context.mailbox.userProfile.emailAddress.split("@")[0];

        item.body.getAsync("text", async (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const bodyText = result.value;

            const payload = {
              subject,
              from: item.from && item.from.emailAddress,
              date,
              body: bodyText,
              userId,
              company,
              project,
              phase
            };

            try {
              const response = await fetch("https://prod-44.japaneast.logic.azure.com:443/workflows/026961ca129249119a3e9c0ed61c35b9/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=FAeF2rSLcH6hJomaCJkdeLT-IxwUSBovmNYBlhnDYxo", {
                method: "POST",
                headers: {
                  "Content-Type": "application/json"
                },
                body: JSON.stringify(payload)
              });

              const msg = document.getElementById("statusMessage");
              if (response.ok) {
                msg.textContent = "kintoneへの登録が完了しました。";
                setTimeout(() => msg.textContent = "", 3000);
              } else {
                msg.textContent = "登録に失敗しました。";
              }
            } catch (err) {
              console.error(err);
              document.getElementById("statusMessage").textContent = "通信エラーが発生しました。";
            }
          }
        });
      } catch (error) {
        console.error("アドイン処理エラー:", error);
      }
    };
  }
});
