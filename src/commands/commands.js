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
      // Adjusted the range to have enough columns for each quarter from Q1_2018 to Q4_2024
      let expensesTable = sheet.tables.add("A1:AG1", true /*hasHeaders*/);
      expensesTable.name = "ExpensesTable";

      expensesTable.getHeaderRowRange().values = [
        [
          "Geography",
          "Product",
          "Brand",
          "Size",
          "Fact",
          "Q1_2018",
          "Q2_2018",
          "Q3_2018",
          "Q4_2018",
          "Q1_2019",
          "Q2_2019",
          "Q3_2019",
          "Q4_2019",
          "Q1_2020",
          "Q2_2020",
          "Q3_2020",
          "Q4_2020",
          "Q1_2021",
          "Q2_2021",
          "Q3_2021",
          "Q4_2021",
          "Q1_2022",
          "Q2_2022",
          "Q3_2022",
          "Q4_2022",
          "Q1_2023",
          "Q2_2023",
          "Q3_2023",
          "Q4_2023",
          "Q1_2024",
          "Q2_2024",
          "Q3_2024",
          "Q4_2024"
        ],
      ];
  
      expensesTable.rows.add(null /*add rows to the end of the table*/, [
        ["▼ Geography Total", "▶ Product Total", "Brand Total", "Size Total", "Fact A", 600, 640, 650, 675, 710, 740, 770, 800, 830, 860, 890, 920, 950, 980, 1010, 1040, 1070, 1100, 1130, 1160, 1190, 1220, 1250, 1280, 1310, 1340, 1370, 1400],
        ["▶ North America", "Product Total", "Brand Total", "Size Total", "Fact A", 100, 120, 110, 115, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350, 360],
        ["▶ Europe", "Product Total", "Brand Total", "Size Total", "Fact A", 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350, 360, 370, 380, 390, 400, 410, 420, 430, 440, 450, 460, 470],
        ["▶ Asia", "Product Total", "Brand Total", "Size Total", "Fact A", 300, 310, 320, 330, 340, 350, 360, 370, 380, 390, 400, 410, 420, 430, 440, 450, 460, 470, 480, 490, 500, 510, 520, 530, 540, 550, 560, 570],
      ]);
  
      // Add the Table Style
      expensesTable.style = "TableStyleLight9";

      // Get the range of cells that allow to edit data.
      let range = sheet.getRange("R2:Y5");
      range.format.fill.color = "#FFF2CC";

      // Get the range of cells of high level data.
      range = sheet.getRange("A2:E2");
      range.format.fill.color = "#8EA9DB";

      // Get the range of cells of high level data.
      range = sheet.getRange("A3:E5");
      range.format.fill.color = "#B4C6E7";

      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
      }

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


