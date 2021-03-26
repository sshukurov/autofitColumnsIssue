/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      range.values = [["45043089.64"]];
      range.format.font.bold = true;
      range.numberFormat = [['_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)']];
      range.format.autofitColumns();

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
