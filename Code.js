function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Mi Web App con CI/CD')
      .setXframeOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}