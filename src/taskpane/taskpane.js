// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const g = getGlobal();


Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    ensureStateInitialized(true);
    console.log("ensure state initialized from the office.initialize");
    isOfficeInitialized = true;
    monitorSheetChanges();

    document.getElementById("connectService").onclick = connectService; // in office-apis-helpers.js
    document.getElementById("selectFilter").onclick = insertFilteredData;
    
    updateRibbon();
    updateTaskPaneUI();
  }
});


Excel.run(function(context) {
  var sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.onSingleClicked.add(myEventHandler);

  return context.sync();
}).catch(errorHandlerFunction);


let lastClickTime = null;

function myEventHandler(event) {
    console.log("Single click detected.");
    let currentTime = new Date();
    if (lastClickTime !== null && currentTime - lastClickTime < 500) {
        console.log("Double click detected.");
        lastClickTime = null; // reset the click time
    } else {
        lastClickTime = currentTime;
    }
}

async function insertFilteredData() {
  try {
    //Determine which data source the user selected from the radio buttons.
    const radioExcel = document.getElementById("communicationFilter");
    if (radioExcel.checked) {
      generateCustomFunction("Communications");
    } else {
      generateCustomFunction("Groceries");
    }
  } catch (error) {
    console.error(error);
  }
}
