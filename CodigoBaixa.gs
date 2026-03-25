const SHEET_ID = '1HPuW7JvmO9_ss0rCayi1MnW9hydRMgXTyDopSlkAO9o';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('DashboardVerba')
    .setTitle('Painel Executivo - Suprimentos')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getDashboardData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    // IMPORTANTE: Nomeie a aba da planilha EXATAMENTE como DASHBOARD
    const sheet = ss.getSheetByName('DASHBOARD'); 
    
    if (!sheet) {
      return { error: "Aba 'DASHBOARD' não encontrada. Verifique o nome na planilha." };
    }
    
    const data = sheet.getDataRange().getDisplayValues();
    return { data: data };
  } catch (e) {
    return { error: e.message };
  }
}
