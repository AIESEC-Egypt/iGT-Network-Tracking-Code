function AIActual() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "Base actuals term 23_AI"
  );
  const sheetData = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  const lcs = sheetData.flat(1);
  for (let i = 0; i < sprintStartDatenew.length; i++) {
    console.log(sprintStartDatenew[i]);

    for (let j = 0; j < 42; j++) {
      console.log(lcs[rowStartIndexnew[i] - 1 + j]);
      var url =
        "https://analytics.api.aiesec.org/v2/applications/analyze.json?access_token=" +
        "&start_date=" +
        sprintStartDatenew[i] +
        "&end_date=" +
        sprintEndDatenew[i] +
        "&performance_v3%5Boffice_id%5D=" +
        ai_codes[lcs[rowStartIndexnew[i] - 1 + j]];

      var response = UrlFetchApp.fetch(url, { method: "GET" }).getContentText();
      var data = JSON.parse(response);
      var list = [];
      list.push([data.o_approved_8.doc_count, data.o_approved_9.doc_count]);

      sheet
        .getRange(rowStartIndexnew[i] + j, 2, 1, list[0].length)
        .setValues(list);
    }
  }
}
