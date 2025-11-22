function saveAutomatico(e) {
  const id_sheet_destino = "1UC1OrMuwb9vY3TNdlazro_C55XAn2F5kmOGePPa0X-k";
  const data = "Dados";

  const colStatus = 6;

  const range = e.range;
  const sheet = range.getSheet();
  const linhaAtual = range.getRow();

  if (sheet.getName() !== data || linhaAtual < 2)
  {
    return;
  }

  const dataRow = sheet.getRange(linhaAtual, 2, 1, 5).getValues()[0];

  const op  = dataRow[0];
  const os = dataRow[1];
  const nomeLoja = dataRow[2];
  const cnpj = dataRow[3];
  const status = dataRow[4];

  if (status !== "")
  {
    return;
  }

  if (op === "" || os === "" || nomeLoja === "" || cnpj === "")
  {
    return;
  }
  try
  {
    const spreadsheetDestino = SpreadsheetApp.openById(id_sheet_destino);
    let sheetDestino = spreadsheetDestino.getSheetByName(nomeLoja);

    if (!sheetDestino)
    {
      sheetDestino = spreadsheetDestino.insertSheet(nomeLoja);
      sheetDestino.appendRow(["OS", "OP", "CNPJ", "EMPRESA"]);
      sheetDestino.getRange("A1:D1").setFontWeight("bold");
    }

    sheetDestino.appendRow([os, op, cnpj, nomeLoja]);
    sheet.getRange(linhaAtual, colStatus).setValue("ENVIADO").setFontColor("green").setFontWeight("bold");
    SpreadsheetApp.getActiveSpreadsheet().toast(`Linha ${linhaAtual} foi enviada com sucesso.`);

    if (e && e.range) {
     console.log(`Alteração detectada na linha: ${e.range.getRow()}, Coluna: ${e.range.getColumn()}`);
     console.log(`Valor novo: ${e.value}`);
  } else {
     console.log("Função rodou sem evento (teste manual?)");
  }
  }
  catch (exception)
  {
    SpreadsheetApp.getUi().alert("Erro ao enviar linha para a planilha desejada. Erro: " + exception.message);
  }
}
