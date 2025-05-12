function onOpen() {
    SpreadsheetApp.getUi()
      .createMenu('Custom')
      .addItem('Import CSV/TSV + Setup Analysis', 'showImportDialog')
      .addToUi();
  }
  
  function showImportDialog() {
    const html = HtmlService
      .createHtmlOutputFromFile('ImportDialog')
      .setWidth(400)
      .setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(html, 'Upload File & Configure');
  }
  
  /**
   * @param {string} fileText
   * @param {string} rangeName
   * @param {string} filterText
   */
  function importCsv(fileText, rangeName, filterText) {
    if (!rangeName.match(/^[A-Za-z]\w*$/)) {
      throw new Error('Named range must start with a letter and contain only letters, numbers, or _.');
    }
    const ss = SpreadsheetApp.getActive();
  
    // 1) Parse CSV or TSV
    const lines = fileText.trim().split(/\r?\n/);
    const rows  = lines.map(line => line.includes('\t') ? line.split('\t') : line.split(','));
    const numRows = rows.length;
    const numCols = rows[0].length;
  
    // 2) Write Data_<rangeName> sheet
    const dataSheetName = `Data_${rangeName}`;
    let dataSheet = ss.getSheetByName(dataSheetName);
    if (!dataSheet) dataSheet = ss.insertSheet(dataSheetName);
    dataSheet.clearContents();
    dataSheet.getRange(1, 1, numRows, numCols).setValues(rows);
  
    // 3) Add Earning Share column
    const esCol = numCols + 1;
    dataSheet.getRange(1, esCol).setValue('Earning Share');
    const esFormula =
      `=IFERROR(
         D2 /
         SUMIFS(
           $D$2:$D$${numRows},
           $A$2:$A$${numRows}, A2,
           $C$2:$C$${numRows}, "*␟"&RIGHT(C2,1)
         ),
         ""
       )`;
    dataSheet.getRange(2, esCol, numRows - 1).setFormula(esFormula);
  
    // 4) Define named range
    ss.getNamedRanges()
      .filter(n => n.getName() === rangeName)
      .forEach(n => n.remove());
    ss.setNamedRange(rangeName, dataSheet.getRange(1, 1, numRows, esCol));
  
    // --- Helper to build one analysis sheet ---
    function buildAnalysis(sheetSuffix, valueIndex) {
      const sheetName = `Analysis_${rangeName}_${sheetSuffix}`;
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) sheet = ss.insertSheet(sheetName);
      sheet.clearContents();
  
      // a) Query pivot
      const query =
        `=QUERY(
           {
             INDEX(${rangeName},,3),
             INDEX(${rangeName},,1),
             INDEX(${rangeName},,${valueIndex})
           },
           "select Col1, sum(Col3)
            where Col3 is not null
            group by Col1
            pivot Col2
            order by Col1
            label Col1 'Ad source instance', sum(Col3) ''",
           0
         )`;
      sheet.getRange(1, 1).setFormula(query);
  
      // b) wait for calculate
      SpreadsheetApp.flush();
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
  
      // c) Variant & Chart columns
      sheet.insertColumnAfter(lastCol);
      sheet.insertColumnAfter(lastCol + 1);
      const varCol   = lastCol + 1;
      const sparkCol = lastCol + 2;
      sheet.getRange(1, varCol).setValue('Variant');
      sheet.getRange(1, sparkCol).setValue('Chart');
  
      // Variant = RIGHT(A2)
      sheet.getRange(2, varCol, lastRow - 1)
           .setFormulaR1C1('=RIGHT(RC1)');
  
      // Sparkline across all date columns (2..lastCol)
      const startOffset = 2 - sparkCol;
      const endOffset   = lastCol - sparkCol;
      sheet.getRange(2, sparkCol, lastRow - 1)
           .setFormulaR1C1(
             `=IFERROR(
                SPARKLINE(
                  RC[${startOffset}]:RC[${endOffset}],
                  {"charttype","line"}
                ),
                ""
              )`
           );
  
      // d) Apply filter on Col A
      sheet.getDataRange()
           .createFilter()
           .setColumnFilterCriteria(
             1,
             SpreadsheetApp.newFilterCriteria()
               .whenTextContains(filterText)
               .build()
           );
    }
  
    // 5) Build first analysis: Earning Share (column esCol)
    buildAnalysis('EarningShare', esCol);
  
    // 6) Build second analysis: Match Rate (column 7)
    buildAnalysis('MatchRate', 7);
  
    // 7) Build second analysis: eCPM (column 7)
    buildAnalysis('eCPM', 5);
  }
  