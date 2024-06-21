function lcsActual() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "Base actuals term 23_LC"
  );
  const sheetData = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  const lcs = sheetData.flat(1);
  for (let i = 0; i < sprintStartDate.length; i++) {
    console.log(sprintStartDate[i]);

    for (let j = 0; j < 25; j++) {
      console.log(lcs[rowStartIndex[i] - 1 + j]);
      var url =
        "https://analytics.api.aiesec.org/v2/applications/analyze.json?access_token=" +
        "&start_date=" +
        sprintStartDate[i] +
        "&end_date=" +
        sprintEndDate[i] +
        "&performance_v3%5Boffice_id%5D=" +
        lcs_codes[lcs[rowStartIndex[i] - 1 + j]];

      var response = UrlFetchApp.fetch(url, { method: "GET" }).getContentText();
      var data = JSON.parse(response);
      var list = [];
      list.push([
        data.open_i_programme_8.doc_count,
        data.i_applied_8.applicants.value,
        data.i_matched_8.doc_count,
        data.i_approved_8.doc_count,
        data.i_realized_8.doc_count,
        data.i_finished_8.doc_count,
        data.i_completed_8.doc_count,
        data.open_i_programme_9.doc_count,
        data.i_applied_9.applicants.value,
        data.i_matched_9.doc_count,
        data.i_approved_9.doc_count,
        data.i_realized_9.doc_count,
        data.i_finished_9.doc_count,
        data.i_completed_9.doc_count,
      ]);

      sheet
        .getRange(rowStartIndex[i] + j, 2, 1, list[0].length)
        .setValues(list);
    }
  }
}
