<!-- taskpane.html -->
<!DOCTYPE html>
<html>
  <head>
    <meta charset="UTF-8" />
    <title>Outlook to Kintone</title>
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    <link rel="stylesheet" href="https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css" />
  </head>
  <body class="ms-font-m ms-Fabric">
    <main id="app-body" class="ms-welcome__main">
      <label>会社名：</label><br />
      <input id="txtCompany" type="text" /><br /><br />

      <label>案件名：</label><br />
      <input id="txtProject" type="text" /><br /><br />

      <label>フェーズ：</label><br />
      <input id="txtPhase" type="text" /><br /><br />

      <label>添付方法：</label><br />
      <input type="radio" id="modeSimple" name="attachmentMode" value="simple" checked />
      <label for="modeSimple">簡易</label><br />
      <input type="radio" id="modeFull" name="attachmentMode" value="full" />
      <label for="modeFull">フル</label><br /><br />
      
      <button id="btn-send-kintone">kintoneに送信</button><br /><br />

      <div id="statusMessage" style="color: green; font-weight: bold;"></div>
    </main>
    <script src="./taskpane.js"></script>
  </body>
</html>
