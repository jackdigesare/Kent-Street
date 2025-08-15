/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

document.addEventListener("DOMContentLoaded", function () {
  Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "block";
      document.getElementById("run").onclick = run;
      document.getElementById("writeCell").onclick = async () => {
        const value = document.getElementById("cellInput").value;
        await Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.values = [[value]];
          await context.sync();
        });
      };
    }
  });
});

async function run() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      await context.sync(); // Wait for address to load
      range.format.fill.color = "yellow";
      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}