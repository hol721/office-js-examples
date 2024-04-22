/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;

  

  if (Office.context.requirements.isSetSupported('SharedRuntime', '1.1')) {
    console.log('This code is using shared runtime version 1.1 requirement set');
  }

  if (Office.context.requirements.isSetSupported('SharedRuntime', '1.2')) {
      console.log('This code is using shared runtime version 1.2 requirement set');
  }

  if (Office.context.requirements.isSetSupported('CustomFunctionsRuntime', '1.1')) {
      console.log('This code is using Custom Functions runtime version 1.1 requirement set');
  }

  if (Office.context.requirements.isSetSupported('ExcelApi', '1.2')) {
      console.log('This code is using Excel.js version 1.2 requirement set');
  }
  if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
      console.log('This code is using Excel.js version 1.3 requirement set');
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
