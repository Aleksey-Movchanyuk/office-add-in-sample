// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

function btnConnectService(event) {
  console.log("Connect service button pressed");
  // Your code goes here
  g.state.setConnected(true);
  g.state.isConnectInProgress = true;
  updateRibbon();
  connectService();
  monitorSheetChanges();
  event.completed();
}

function btnDisconnectService(event) {
  console.log("Disconnect service button pressed");
  // Your code goes here
  g.state.setConnected(false);
  updateRibbon();
  updateTaskPaneUI();
  event.completed();
}

function btnUserProfile(event) {
  console.log("User profile button pressed");
  // Your code goes here

  event.completed();
}

function btnHelp(event) {
  console.log("User help button pressed");
  // Your code goes here

  event.completed();
}

function btnInsertData(event) {
  console.log("Insert data button pressed");
  // Mock code that pretends to insert data from a data source
  insertData();
  event.completed();
}

function btnRefreshData(event) {
  console.log("Refresh data button pressed");
  // Mock code that pretends to insert data from a data source

  event.completed();
}

function btnFilterData(event) {
  console.log("Filter data button pressed");
  // Mock code that pretends to insert data from a data source

  event.completed();
}

function btnScope(event) {
  console.log("Scope data button pressed");
  // Mock code that pretends to insert data from a data source

  event.completed();
}

function btnParameters(event) {
  console.log("Parameters data button pressed");
  // Mock code that pretends to insert data from a data source

  event.completed();
}

function btnSimulation(event) {
  console.log("Simulation button pressed");
  // Your code goes here

  event.completed();
}

function btnMassiveSimulation(event) {
  console.log("Massive Simulation button pressed");
  // Your code goes here

  event.completed();
}

function btnOpenTaskpane(event) {
  console.log("Open task pane button pressed");
  // Your code goes here
  SetRuntimeVisibleHelper(true);
  g.state.isTaskpaneOpen = true;
  updateRibbon();
  event.completed();
}

function btnCloseTaskpane(event) {
  console.log("Open task pane button pressed");
  // Your code goes here
  SetRuntimeVisibleHelper(false);
  g.state.isTaskpaneOpen = false;
  updateRibbon();
  event.completed();
}

function btnSettings(event) {
  console.log("Settings button pressed");
  // Your code goes here

  event.completed();
}

async function insertData() {
  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
      expensesTable.name = "ExpensesTable";

      expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

      expensesTable.rows.add(null /*add rows to the end of the table*/, [
        ["1/1/2017", "The Phone Company", "Communications", "$120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "$142"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "$27"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "$33"],
        ["1/11/2017", "Bellows College", "Education", "$350"],
        ["1/15/2017", "Trey Research", "Other", "$135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "$97"],
      ]);

      /*if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
      }*/
      context.sync().then(() => {
        monitorSheetChanges();
        return context.sync();
      });
    });
  } catch (error) {
    console.log(error);
  }
}



const g = getGlobal();

Office.actions.associate("btnConnectService", btnConnectService);
Office.actions.associate("btnDisconnectService", btnDisconnectService);
Office.actions.associate("btnUserProfile", btnUserProfile);
Office.actions.associate("btnHelp", btnHelp);

Office.actions.associate("btnInsertData", btnInsertData);
Office.actions.associate("btnRefreshData", btnRefreshData);
Office.actions.associate("btnFilterData", btnFilterData);
Office.actions.associate("btnScope", btnScope);
Office.actions.associate("btnParameters", btnParameters);

Office.actions.associate("btnSimulation", btnSimulation);
Office.actions.associate("btnMassiveSimulation", btnMassiveSimulation);

Office.actions.associate("btnOpenTaskpane", btnOpenTaskpane);
Office.actions.associate("btnCloseTaskpane", btnCloseTaskpane);

Office.actions.associate("btnSettings", btnSettings);


