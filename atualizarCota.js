function atualizarCotaNaPlanilha() {
  // Puxa a cota do Google
  var cotaRestante = MailApp.getRemainingDailyQuota();
  
  // Escolhe a aba e a célula onde o valor vai aparecer (Ex: Aba "Editais", célula "J1")
  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Editais");
  
  aba.getRange("J1").setValue("📧 Cota Google: " + cotaRestante);
  
  // Opcional: formata a célula para ficar destacada (fundo preto, letra amarela)
  aba.getRange("J1").setBackground("#202124").setFontColor("#fbbc04").setFontWeight("bold");
}
