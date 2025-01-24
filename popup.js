(async () => {
  await Office.onReady();

  document.getElementById("ok-button").onclick = sendStringToParentPage;

  function sendStringToParentPage() {
    const sheetName = document.getElementById("sheetName").value;
    Office.context.ui.messageParent(sheetName);
  }
})();
