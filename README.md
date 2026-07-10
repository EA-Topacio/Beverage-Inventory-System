# Beverage Inventory & Consumption System

## Overview

A Business Intelligence and Inventory Management System developed using Google Sheets and Google Apps Script to automate banquet beverage operations.

## Features

- Banquet Event Order (BEO) Management
- Beverage Inventory Tracking
- Stock Movement Requests
- Inventory Variance Calculation
- KPI Dashboard
- Automated Reporting

## Technologies

- Google Apps Script
- Google Sheets
- JavaScript

## Skills Demonstrated

- Business Intelligence
- Dashboard Design
- Process Automation
- Database Design
- Data Analysis
- Workflow Automation

## Project Screenshots

📁 screenshots
<img width="485" height="561" alt="image" src="https://github.com/user-attachments/assets/94b9bce3-efbc-4a36-bb7a-eb32517f7355" />
<img width="440" height="451" alt="image" src="https://github.com/user-attachments/assets/3a8f88c7-b2d4-4102-843e-d1e8231eb70c" />
<img width="514" height="581" alt="image" src="https://github.com/user-attachments/assets/d278080f-b93a-4e7d-8951-e5b5efadd8c6" />
<img width="640" height="445" alt="image" src="https://github.com/user-attachments/assets/b262fc0b-218d-417e-bc5e-5faa3b71b058" />
<img width="1066" height="617" alt="image" src="https://github.com/user-attachments/assets/f7270ac4-ad22-45d1-9b5f-32726779de0c" />
<img width="1565" height="152" alt="image" src="https://github.com/user-attachments/assets/712394c9-c2a5-4883-b209-5fa7c0e8e2ea" />
<img width="1158" height="483" alt="image" src="https://github.com/user-attachments/assets/005421a2-52f7-4890-8a50-a9154fbd13c9" />
<img width="1236" height="462" alt="image" src="https://github.com/user-attachments/assets/5fff3724-28fb-4ace-b1d1-bac31f7c45f3" />

📁 documentation
<img width="570" height="630" alt="image" src="https://github.com/user-attachments/assets/fcc8a6a2-0779-4d32-90d7-15f154eddb0c" />
<img width="798" height="655" alt="image" src="https://github.com/user-attachments/assets/9c3ec49f-7f44-489c-a3da-32700f8e90a2" />
<img width="781" height="538" alt="image" src="https://github.com/user-attachments/assets/044aa0c0-0b55-4426-ba33-d26c2e7312c8" />
<img width="689" height="409" alt="image" src="https://github.com/user-attachments/assets/bf769e88-a20b-4578-a2a5-81e0a9af2744" />
<img width="779" height="323" alt="image" src="https://github.com/user-attachments/assets/1e1b59ad-38ca-45f2-9b51-2aa738596fbe" />

📁 screenshots
const ADDITIONAL_BEVERAGE_ID = "ADDBEV"

//Menu
function openCreateBEO() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName("Create_BEO");

  sheet.activate();
}
function openManageBEO() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  ss.getSheetByName("Manage_BEO")
    .activate();

}
function openCreateRequest() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  ss.getSheetByName("Create_StockMovement")
    .activate();
}
function openManageRequest() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  ss.getSheetByName("Manage_StockMovement")
    .activate();
}
function openManageInventory() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  ss.getSheetByName("Manage_Inventory")
    .activate();
}
function openDashboard() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  ss.getSheetByName("Dashboard")
    .activate();
}
function openStockMovement() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  ss.getSheetByName("STOCK_MOVEMENT")
    .activate();
}

//Back 
function backFromCreateBEO() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const currentSheet =
    ss.getSheetByName("Create_BEO");

  const menuSheet =
    ss.getSheetByName("Menu");

  menuSheet.activate();

  currentSheet.hideSheet();

}
function backFromManageBEO() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const currentSheet =
    ss.getSheetByName("Manage_BEO");

  const menuSheet =
    ss.getSheetByName("Menu");

  menuSheet.activate();

  currentSheet.hideSheet();

}
function backFromCreateRequest() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const currentSheet =
    ss.getSheetByName("Create_StockMovement");

  const menuSheet =
    ss.getSheetByName("Menu");

  menuSheet.activate();

  currentSheet.hideSheet();

}
function backFromManageRequest() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const currentSheet =
    ss.getSheetByName("Manage_StockMovement");

  const menuSheet =
    ss.getSheetByName("Menu");

  menuSheet.activate();

  currentSheet.hideSheet();

}
function backFromManageInventory() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const currentSheet =
    ss.getSheetByName("Manage_Inventory");

  const menuSheet =
    ss.getSheetByName("Menu");

  menuSheet.activate();

  currentSheet.hideSheet();

}
function backFromDashboard() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const currentSheet =
    ss.getSheetByName("Dashboard");

  const menuSheet =
    ss.getSheetByName("Menu");

  menuSheet.activate();

  currentSheet.hideSheet();

}
function backFromStockMovement() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const currentSheet =
    ss.getSheetByName("STOCK_MOVEMENT");

  const menuSheet =
    ss.getSheetByName("Menu");

  menuSheet.activate();

  currentSheet.hideSheet();

}

// ====================
// Banquet Event Order
// =====================

//Create_BEO
function createBEO() {
  //Create BEO - Submit: Adds record to the database

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const beoSheet = ss.getSheetByName("Create_BEO");
  const eventSheet = ss.getSheetByName("EVENT");
  const eventPackageSheet = ss.getSheetByName("EVENT_PACKAGE");
  const eventBeverageSheet = ss.getSheetByName("EVENT_BEVERAGE");
  const packageSheet = ss.getSheetByName("PACKAGE");
  const packageBeverageSheet = ss.getSheetByName("PACKAGE_BEVERAGE");
  const beverageSheet = ss.getSheetByName("BEVERAGE");

  // ==========================
  // FORM VALUES
  // ==========================

  const beoNumber =
    beoSheet.getRange("D5")
      .getValue()
      .toString()
      .trim();

  const createdBy =
    beoSheet.getRange("D6")
      .getValue()
      .toString()
      .trim();

  const selectedPackages =
    beoSheet
      .getRange("D7:D11")
      .getValues()
      .flat()
      .map(x => String(x).trim())
      .filter(x => x !== "");

  const additionalBeverages =
    beoSheet
      .getRange("D13:D22")
      .getValues()
      .flat()
      .map(x => String(x).trim())
      .filter(x => x !== "");

  // ==========================
  // VALIDATION
  // ==========================

  if (!beoNumber) {

    SpreadsheetApp.getUi().alert(
      "BEO Number is required."
    );

    return;
  }

  if (!createdBy) {

    SpreadsheetApp.getUi().alert(
      "Created By is required."
    );

    return;
  }

  if (
    selectedPackages.length === 0 &&
    additionalBeverages.length === 0
  ) {

    SpreadsheetApp.getUi().alert(
      "Please select at least one package or one additional beverage."
    );

    return;
  }

  // ==========================
  // CHECK DUPLICATE BEO
  // ==========================

  const existingEvents =
    eventSheet
      .getDataRange()
      .getValues();

  for (let i = 1; i < existingEvents.length; i++) {

    if (
      String(existingEvents[i][1]).trim() ===
      beoNumber
    ) {

      SpreadsheetApp.getUi().alert(
        "BEO Number already exists."
      );

      return;
    }

  }

  // ==========================
  // GENERATE NEXT EVT ID
  // ==========================

  let maxID = 0;

  for (let i = 1; i < existingEvents.length; i++) {

    const id =
      String(existingEvents[i][0]);

    if (id.startsWith("EVT")) {

      const num =
        parseInt(
          id.replace("EVT", ""),
          10
        );

      if (num > maxID) {
        maxID = num;
      }

    }

  }

  const beoID =
    "EVT" +
    String(maxID + 1)
      .padStart(4, "0");

  // ==========================
  // CREATE EVENT RECORD
  // ==========================

  eventSheet.appendRow([
    beoID,
    beoNumber,
    createdBy,
    new Date(),
    0
  ]);

  // ==========================
  // PACKAGE LOOKUP
  // ==========================

  const packageData =
    packageSheet
      .getDataRange()
      .getValues();

  const packageMap = {};

  for (let i = 1; i < packageData.length; i++) {

    const packageID =
      String(packageData[i][0]).trim();

    const packageName =
      String(packageData[i][1]).trim();

    packageMap[packageName] =
      packageID;

  }

  // ==========================
  // BEVERAGE LOOKUP
  // ==========================

  const beverageData =
    beverageSheet
      .getDataRange()
      .getValues();

  const beverageMap = {};
  const beveragePriceMap = {};

  for (let i = 1; i < beverageData.length; i++) {

    const beverageID =
      String(beverageData[i][0]).trim();

    const beverageName =
      String(beverageData[i][2]).trim();

    const unitPrice =
      Number(beverageData[i][4]) || 0;

    beverageMap[beverageName] =
      beverageID;

    beveragePriceMap[beverageID] =
      unitPrice;

  }

  // ==========================
  // EVENT PACKAGE
  // ==========================

  selectedPackages.forEach(packageName => {

    const cleanPackageName =
      String(packageName).trim();

    const packageID =
      packageMap[cleanPackageName];

    if (!packageID) return;
    
    eventPackageSheet.appendRow([
      beoID,
      packageID,
      cleanPackageName
    ]);

  });

  // ==========================
  // PACKAGE BEVERAGE DATA
  // ==========================

  const packageBeverageData =
    packageBeverageSheet
      .getDataRange()
      .getValues();

  const addedBeverages =
    new Set();

  // ==========================
  // PACKAGE BEVERAGES
  // ==========================

  selectedPackages.forEach(packageName => {

    const cleanPackageName =
      String(packageName).trim();

    const packageID =
      packageMap[cleanPackageName];

    if (!packageID) return;

    for (
      let i = 1;
      i < packageBeverageData.length;
      i++
    ) {

      const rowPackageID =
        String(
          packageBeverageData[i][0]
        ).trim();

      if (
        rowPackageID === packageID
      ) {

        const beverageID =
          String(
            packageBeverageData[i][1]
          ).trim();

        const beverageName =
          packageBeverageData[i][2];

        if (
          !addedBeverages.has(
            beverageID
          )
        ) {

          eventBeverageSheet.appendRow([
            beoID,
            beoNumber,
            beverageID,
            beverageName,
            packageID,
            cleanPackageName,
            "",
            "",
            Number(
              beveragePriceMap[
                beverageID
              ]
            ) || 0,
            0
          ]);

          addedBeverages.add(
            beverageID
          );

        }

      }

    }

  });

  // ==========================
  // ADDITIONAL BEVERAGES
  // ==========================

  additionalBeverages.forEach(
    beverageName => {

      const beverageID =
        beverageMap[
          String(beverageName).trim()
        ];

      if (!beverageID) return;

      if (
        !addedBeverages.has(beverageID)
      ) {

        eventBeverageSheet.appendRow([
          beoID,
          beoNumber,
          beverageID,
          beverageName,
          ADDITIONAL_BEVERAGE_ID,
          "Additional Beverage",
          "",
          "",
          Number(
            beveragePriceMap[beverageID]
          ) || 0,
          0
        ]);

        addedBeverages.add(
          beverageID
        );

      }

    }
  );
  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Create_BEO");

  sheet.getRange("D5").clearContent();
  sheet.getRange("D6").clearContent();
  sheet.getRange("D7:D11").clearContent();
  sheet.getRange("D13:D22").clearContent();

  SpreadsheetApp.getUi().alert(
    "BEO Created Successfully!\n\n" +
    "BEO Number: " + beoNumber + "\n" +
    "Created By: " + createdBy + "\n\n" +

    "Selected Packages:\n" +
    (selectedPackages.length
      ? selectedPackages.join("\n")
      : "None") +

    "\n\nAdditional Beverages:\n" +

    (additionalBeverages.length
      ? additionalBeverages.join("\n")
      : "None")
  );

}
function clearBEO() {
  // Clear Form - Deletes all written text on the BEO form to prevent manual deletion.

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Create_BEO");

  sheet.getRange("D5").clearContent();
  sheet.getRange("D6").clearContent();
  sheet.getRange("D7:D11").clearContent();
  sheet.getRange("D13:D22").clearContent();

  SpreadsheetApp.getUi().alert(
    "BEO Form Cleared!" 
  );

}

//Manage_BEO
function clearManageBEO() {
  // Clear Form - Deletes all written text on the Manage BEO form to prevent manual deletion.

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Manage_BEO");

  sheet.getRange("K1").clearContent();
  sheet.getRange("D5:D12").clearContent();
  sheet.getRange("B14:I48").clearContent();
  sheet.getRange("I49").clearContent();

  SpreadsheetApp.getUi().alert(
    "BEO Management Form Cleared!" 
  );

}
function searchBEO() {
  // Enter a BEO Number and it will automatically display all details related to that number.

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const manageSheet =
    ss.getSheetByName("Manage_BEO");

  const eventSheet =
    ss.getSheetByName("EVENT");

  const eventPackageSheet =
    ss.getSheetByName("EVENT_PACKAGE");

  const eventBeverageSheet =
    ss.getSheetByName("EVENT_BEVERAGE");

  const beoNumber =
    manageSheet.getRange("D5")
      .getValue()
      .toString()
      .trim();

  if (!beoNumber) {
    SpreadsheetApp.getUi().alert(
      "BEO Number is blank." +
      "\n\nPlease enter a valid BEO Number."
    );
    return;
  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Manage_BEO");

  sheet.getRange("K1").clearContent();
  sheet.getRange("D5:D12").clearContent();
  sheet.getRange("B14:I48").clearContent();
  sheet.getRange("I49").clearContent();

  manageSheet
    .getRange("D5")
    .setValue(beoNumber);

  const events =
    eventSheet
      .getDataRange()
      .getValues();

  let beoID = "";

  for (let i = 1; i < events.length; i++) {

    if (
      String(events[i][1]).trim() ===
      beoNumber
    ) {

      beoID = events[i][0];

      manageSheet
        .getRange("K1")
        .setValue(beoID);

      manageSheet
        .getRange("D6")
        .setValue(events[i][2]);

      manageSheet
        .getRange("D7")
        .setValue(events[i][3]);

      break;

    }
  }

  if (!beoID) {
    SpreadsheetApp.getUi().alert(
      "BEO Number: " + beoNumber + " not found." +
      "\n\nPlease enter a valid BEO Number."
    );
    sheet.getRange("K1").clearContent();
    sheet.getRange("D5:D12").clearContent();
    sheet.getRange("B14:I48").clearContent();
    sheet.getRange("I49").clearContent();
    return;
  }

  let packageRow = 8;

  const packageData =
    eventPackageSheet
      .getDataRange()
      .getValues();

  for (let i = 1; i < packageData.length; i++) {

    if (packageData[i][0] === beoID) {

      manageSheet
        .getRange("D" + packageRow)
        .setValue(packageData[i][2]);

      packageRow++;

    }
  }

  const beverageData =
    eventBeverageSheet
      .getDataRange()
      .getValues();

  let beverageRow = 14;
  let totalPrice = 0;

  for (let i = 1; i < beverageData.length; i++) {
    
    // Change beverageRow, if more rows are to be added in the beverage table.
    if (
      beverageData[i][0] === beoID &&
      beverageRow <= 48
      ) {

      manageSheet.getRange("B" + beverageRow)
        .setValue(beverageData[i][3]);

      manageSheet.getRange("D" + beverageRow)
        .setValue(beverageData[i][5]);

      manageSheet.getRange("F" + beverageRow)
        .setValue(beverageData[i][6]);

      manageSheet.getRange("G" + beverageRow)
        .setValue(beverageData[i][7]);

      manageSheet.getRange("H" + beverageRow)
        .setValue(beverageData[i][8]);

      manageSheet.getRange("I" + beverageRow)
        .setValue(beverageData[i][9]);

      totalPrice +=
        Number(beverageData[i][9]) || 0;

      beverageRow++;

    }

  }

  manageSheet
    .getRange("I49")
    .setValue(totalPrice);

}
function saveChangesBEO() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const manageSheet =
    ss.getSheetByName("Manage_BEO");

  const eventSheet =
    ss.getSheetByName("EVENT");

  const eventBeverageSheet =
    ss.getSheetByName("EVENT_BEVERAGE");

  const beoID =
    manageSheet.getRange("K1").getValue();

  if (!beoID) return;

  // Change Range, if added more rows.
  const manageData =
    manageSheet
      .getRange("B14:I48")
      .getValues();

  const beverageData =
    eventBeverageSheet
      .getDataRange()
      .getValues();

  let eventTotal = 0;

  for (let i = 1; i < beverageData.length; i++) {

    if (beverageData[i][0] !== beoID)
      continue;

    const beverageName =
      beverageData[i][3];

    const packageName =
      beverageData[i][5];

    for (let j = 0; j < manageData.length; j++) {

      if (
        manageData[j][0] === beverageName &&
        manageData[j][2] === packageName
      ) {

        const allocatedQty =
          Number(manageData[j][4]) || 0;

        const actualQty =
          Number(manageData[j][5]) || 0;

        const unitPrice =
          Number(beverageData[i][8]) || 0;

        const totalPrice =
          actualQty * unitPrice;

        beverageData[i][6] =
          allocatedQty;

        beverageData[i][7] =
          actualQty;

        beverageData[i][9] =
          totalPrice;

        eventTotal += totalPrice;

        break;

      }

    }

  }

  eventBeverageSheet
    .getRange(
      1,
      1,
      beverageData.length,
      beverageData[0].length
    )
    .setValues(beverageData);

  const eventData =
    eventSheet
      .getDataRange()
      .getValues();

  for (let i = 1; i < eventData.length; i++) {

    if (eventData[i][0] === beoID) {

      eventSheet
        .getRange(i + 1, 5)
        .setValue(eventTotal);

      break;

    }

  }

  searchBEO();
  updateEventConsumption();

  SpreadsheetApp.getUi().alert(
    "Changes Saved Successfully!"
  );

}
function deleteBEO() {
  // Deletes BEO record from database

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const manageSheet =
    ss.getSheetByName("Manage_BEO");

  const beoID =
    manageSheet.getRange("K1").getValue();

  const beoNumber =
    manageSheet.getRange("D5").getValue();

  if (!beoID) {

    SpreadsheetApp.getUi().alert(
      "Please search a BEO first."
    );

    return;

  }

  const response =
    SpreadsheetApp.getUi().alert(
      "Delete BEO",
      "Are you sure you want to delete BEO Number: " +
      beoNumber +
      "?\n\nThis action cannot be undone. Continue?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

  if (
    response !==
    SpreadsheetApp.getUi().Button.YES
  ) return;

  removeRowsByBEO(
    ss.getSheetByName("EVENT"),
    beoID
  );

  removeRowsByBEO(
    ss.getSheetByName("EVENT_PACKAGE"),
    beoID
  );

  removeRowsByBEO(
    ss.getSheetByName("EVENT_BEVERAGE"),
    beoID
  );

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Manage_BEO");

  sheet.getRange("K1").clearContent();
  sheet.getRange("D5:D12").clearContent();
  sheet.getRange("B14:I48").clearContent();
  sheet.getRange("I49").clearContent();

  SpreadsheetApp.getUi().alert(
    "BEO Deleted Successfully."
  );

}
function removeRowsByBEO(sheet, beoID) {
  // Delete rows helper for BEO
  const data =
    sheet.getDataRange().getValues();

  for (
    let i = data.length - 1;
    i >= 1;
    i--
  ) {

    if (
      String(data[i][0]).trim() ===
      String(beoID).trim()
    ) {

      sheet.deleteRow(i + 1);

    }

  }

}
function updateEventConsumption() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const stockSheet =
    ss.getSheetByName(
      "STOCK_COUNT"
    );

  const eventBeverageSheet =
    ss.getSheetByName(
      "EVENT_BEVERAGE"
    );

  const stockData =
    stockSheet
      .getDataRange()
      .getValues();

  const beverageData =
    eventBeverageSheet
      .getDataRange()
      .getValues();

  // Reset Event Consumption

  for (let i = 1; i < stockData.length; i++) {

    stockData[i][4] = 0;

  }

  // Sum Actual Consumed Qty

  for (let i = 1; i < beverageData.length; i++) {

    const beverageName =
      String(
        beverageData[i][3]
      ).trim();

    const actualConsumed =
      Number(
        beverageData[i][7]
      ) || 0;

    for (
      let j = 1;
      j < stockData.length;
      j++
    ) {

      if (
        String(
          stockData[j][1]
        ).trim() ===
        beverageName
      ) {

        stockData[j][4] =
          (Number(
            stockData[j][4]
          ) || 0) +
          actualConsumed;

        break;

      }

    }

  }

  stockSheet
    .getRange(
      1,
      1,
      stockData.length,
      stockData[0].length
    )
    .setValues(
      stockData
    );

}

// ====================
// Stock Movement
// =====================

//Create_StockMovement
function clearStockMovementForm() {

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("Create_StockMovement");

  sheet.getRange("D5:D9")
    .clearContent();

    SpreadsheetApp.getUi().alert(
    "Stock Movement Request Form Cleared!" 
  );

}
function createStockMovementRequest() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const formSheet =
    ss.getSheetByName("Create_StockMovement");

  const stockMovementSheet =
    ss.getSheetByName("STOCK_MOVEMENT");

  const beverageSheet =
    ss.getSheetByName("BEVERAGE");

  // ==========================
  // FORM VALUES
  // ==========================

  const createdBy =
    formSheet.getRange("D5")
      .getValue()
      .toString()
      .trim();

  const reference =
    formSheet.getRange("D6")
      .getValue()
      .toString()
      .trim();

  const beverageName =
    formSheet.getRange("D7")
      .getValue()
      .toString()
      .trim();

  const movementType =
    formSheet.getRange("D8")
      .getValue()
      .toString()
      .trim();

  const quantity =
    Number(
      formSheet.getRange("D9")
        .getValue()
    );

  // ==========================
  // VALIDATION
  // ==========================

  if (!createdBy) {
    SpreadsheetApp.getUi().alert(
      "Created By is required."
    );
    return;
  }

  if (!reference) {
    SpreadsheetApp.getUi().alert(
      "Reference is required."
    );
    return;
  }

  if (!beverageName) {
    SpreadsheetApp.getUi().alert(
      "Beverage is required."
    );
    return;
  }

  if (!movementType) {
    SpreadsheetApp.getUi().alert(
      "Movement Type is required."
    );
    return;
  }

  if (!quantity || quantity <= 0) {
    SpreadsheetApp.getUi().alert(
      "Quantity must be greater than zero."
    );
    return;
  }

  // ==========================
  // GENERATE MOVEMENT ID
  // ==========================

  const existingData =
    stockMovementSheet
      .getDataRange()
      .getValues();

  let maxID = 0;

  for (let i = 1; i < existingData.length; i++) {

    const id =
      String(existingData[i][0]);

    if (id.startsWith("SM")) {

      const num =
        parseInt(
          id.replace("SM", ""),
          10
        );

      if (num > maxID) {
        maxID = num;
      }

    }

  }

  const movementID =
    "SM" +
    String(maxID + 1)
      .padStart(4, "0");

  // ==========================
  // GET BEVERAGE ID
  // ==========================

  const beverageData =
    beverageSheet
      .getDataRange()
      .getValues();

  let beverageID = "";

  for (let i = 1; i < beverageData.length; i++) {

    if (
      String(beverageData[i][2]).trim() ===
      beverageName
    ) {

      beverageID =
        beverageData[i][0];

      break;

    }

  }

  // ==========================
  // SAVE REQUEST
  // ==========================
  const movementQuantity =
    movementType === "REMOVAL"
      ? -quantity
      : quantity;
      
  stockMovementSheet.appendRow([
    movementID,          // A
    new Date(),          // B Creation Date
    createdBy,           // C
    beverageID,          // D
    beverageName,        // E
    movementType,        // F
    quantity,            // G Quantity Entered
    movementQuantity,    // H Movement Quantity
    reference,           // I
    "PENDING",           // J Status
    "",                  // K Approved By
    ""                   // L Approved Date
  ]);

  const sheet =
    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("Create_StockMovement");

  sheet.getRange("D5:D9")
    .clearContent();

  SpreadsheetApp.getUi().alert(
    "Stock Movement Request Created Successfully!" + "\n" +
    "\nCreated By: " + createdBy +
    "\nReference: " + reference + 
    "\nBeverage: " + beverageName + 
    "\nMovement Type: " + movementType + 
    "\nQuantity: " + quantity
  );

}

//Manage_StockMovement
function loadPendingRequests(showPopup = true) {
  // Displays all pending stock movement requests
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const manageSheet =
    ss.getSheetByName(
      "Manage_StockMovement"
    );

  const stockMovementSheet =
    ss.getSheetByName(
      "STOCK_MOVEMENT"
    );

  const lastRow =
    manageSheet.getMaxRows();

  manageSheet.getRange("D5").clearContent();
  manageSheet.getRange("B7:I" + lastRow).clearContent();

  const data =
    stockMovementSheet
      .getDataRange()
      .getValues();

  let row = 7;
  let pendingCount = 0;

  for (let i = 1; i < data.length; i++) {

    if (
      String(data[i][9]).trim() ===
      "PENDING"
    ) {

      manageSheet
        .getRange("B" + row)
        .setValue(data[i][1]);

      manageSheet
        .getRange("C" + row)
        .setValue(data[i][8]);

      manageSheet
        .getRange("D" + row)
        .setValue(data[i][4]);

      manageSheet
        .getRange("E" + row)
        .setValue(data[i][5]);

      manageSheet
        .getRange("F" + row)
        .setValue(data[i][6]);

      manageSheet
        .getRange("G" + row)
        .setValue(data[i][9]);

      manageSheet
        .getRange("H" + row)
        .insertCheckboxes();

      manageSheet
        .getRange("I" + row)
        .setValue(data[i][0]);

      row++;
      pendingCount++;

    }

  }

  if (showPopup) {

    SpreadsheetApp.getUi().alert(
      pendingCount +
      " pending request" +
      (pendingCount === 1 ? "" : "s") +
      " loaded successfully!"
    );

  }

}
function selectAllRequests() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        "Manage_StockMovement"
      );

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 7) return;

  sheet
    .getRange(
      7,
      8,
      lastRow - 6,
      1
    )
    .setValue(true);

}
function approveSelectedRequests() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const manageSheet =
    ss.getSheetByName(
      "Manage_StockMovement"
    );

  const stockMovementSheet =
    ss.getSheetByName(
      "STOCK_MOVEMENT"
    );

  const stockCountSheet =
    ss.getSheetByName(
      "STOCK_COUNT"
    );

  const approvedBy =
    manageSheet
      .getRange("D5")
      .getValue()
      .toString()
      .trim();

  if (!approvedBy) {

    SpreadsheetApp.getUi().alert(
      "Approved By is required."
    );

    return;

  }

  const lastRow =
    manageSheet.getLastRow();

  if (lastRow < 7) {

    SpreadsheetApp.getUi().alert(
      "No requests loaded."
    );

    return;

  }

  const manageData =
    manageSheet
      .getRange(
        7,
        2,
        lastRow - 6,
        8
      )
      .getValues();

  const stockMovementData =
    stockMovementSheet
      .getDataRange()
      .getValues();

  const stockCountData =
    stockCountSheet
      .getDataRange()
      .getValues();

  // ==========================
  // CHECK IF ANY REQUESTS
  // WERE SELECTED
  // ==========================

  let selectedCount = 0;

  for (
    let r = 0;
    r < manageData.length;
    r++
  ) {

    if (
      manageData[r][6] === true
    ) {

      selectedCount++;

    }

  }

  if (selectedCount === 0) {

    SpreadsheetApp.getUi().alert(
      "Please select stock movement requests to be approved."
    );

    return;

  }

  // ==========================
  // APPROVE REQUESTS
  // ==========================

  let approvedCount = 0;

  for (
    let r = 0;
    r < manageData.length;
    r++
  ) {

    const selected =
      manageData[r][6];

    const movementID =
      manageData[r][7];

    if (
      selected !== true ||
      !movementID
    ) {

      continue;

    }

    for (
      let i = 1;
      i < stockMovementData.length;
      i++
    ) {

      if (
        stockMovementData[i][0] ===
        movementID
      ) {

        const beverageID =
          stockMovementData[i][3];

        const movementQty =
          Number(
            stockMovementData[i][7]
          ) || 0;

        // ==========================
        // UPDATE STOCK MOVEMENT
        // ==========================

        stockMovementData[i][9] =
          "APPROVED";

        stockMovementData[i][10] =
          approvedBy;

        stockMovementData[i][11] =
          new Date();

        // ==========================
        // UPDATE STOCK COUNT
        // ==========================

        for (
          let j = 1;
          j < stockCountData.length;
          j++
        ) {

          if (
            stockCountData[j][0] ===
            beverageID
          ) {

            stockCountData[j][3] =
              (Number(
                stockCountData[j][3]
              ) || 0) +
              movementQty;

            break;

          }

        }

        approvedCount++;

        break;

      }

    }

  }

  // ==========================
  // SAVE CHANGES
  // ==========================

  stockMovementSheet
    .getRange(
      1,
      1,
      stockMovementData.length,
      stockMovementData[0].length
    )
    .setValues(
      stockMovementData
    );

  stockCountSheet
    .getRange(
      1,
      1,
      stockCountData.length,
      stockCountData[0].length
    )
    .setValues(
      stockCountData
    );

  // ==========================
  // RELOAD PENDING REQUESTS
  // ==========================

  loadPendingRequests(false);

  SpreadsheetApp.getUi().alert(
    approvedCount +
    " request(s) approved successfully."
  );

}

// ====================
// Inventory
// =====================
function clearInventoryForm() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        "Manage_Inventory"
      );

  const maxRows =
    sheet.getMaxRows();

  sheet
    .getRange(
      "B9:H" + maxRows
    )
    .clearContent();

}
function loadInventory(showPopup = true) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const inventorySheet =
    ss.getSheetByName(
      "Manage_Inventory"
    );

  const stockSheet =
    ss.getSheetByName(
      "STOCK_COUNT"
    );

  clearInventoryForm();

  const data =
    stockSheet
      .getDataRange()
      .getValues();

  let row = 9;

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const openingStock =
      Number(data[i][2]) || 0;

    const movementQty =
      Number(data[i][3]) || 0;

    const eventConsumption =
      Number(data[i][4]) || 0;

    const expectedStock =
      openingStock +
      movementQty -
      eventConsumption;

    // Save expected stock back to STOCK_COUNT
    data[i][5] =
      expectedStock;

    inventorySheet
      .getRange(
        row,
        2,
        1,
        7
      )
      .setValues([
        [
          data[i][1],       // Beverage
          openingStock,     // Opening Stock
          movementQty,      // Movement Qty
          eventConsumption, // Event Consumption
          expectedStock,    // Expected Stock
          data[i][6],       // Actual Stock
          data[i][7]        // Variance
        ]
      ]);

    row++;

  }

  // Update STOCK_COUNT
  stockSheet
    .getRange(
      1,
      1,
      data.length,
      data[0].length
    )
    .setValues(data);

  if (showPopup) {

    SpreadsheetApp.getUi().alert(
      (data.length - 1) +
      " inventory record(s) loaded successfully."
    );

  }

}
function saveInventoryChanges() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const inventorySheet =
    ss.getSheetByName(
      "Manage_Inventory"
    );

  const stockSheet =
    ss.getSheetByName(
      "STOCK_COUNT"
    );

  const stockData =
    stockSheet
      .getDataRange()
      .getValues();

  const inventoryData =
    inventorySheet
      .getRange(
        9,
        2,
        stockData.length - 1,
        7
      )
      .getValues();

  for (
    let i = 1;
    i < stockData.length;
    i++
  ) {

    const openingStock =
      Number(
        inventoryData[i - 1][1]
      ) || 0;

    const movementQty =
      Number(
        inventoryData[i - 1][2]
      ) || 0;

    const eventConsumption =
      Number(
        inventoryData[i - 1][3]
      ) || 0;

    const expectedStock =
      openingStock +
      movementQty -
      eventConsumption;

    const actualStockCell =
      inventoryData[i - 1][5];

    const actualStock =
      actualStockCell === ""
        ? ""
        : Number(
            actualStockCell
          );

    const variance =
      actualStock === ""
        ? ""
        : actualStock -
          expectedStock;

    // STOCK_COUNT

    stockData[i][5] =
      expectedStock;

    stockData[i][6] =
      actualStock;

    stockData[i][7] =
      variance;

    // Refresh form values

    inventoryData[i - 1][4] =
      expectedStock;

    inventoryData[i - 1][6] =
      variance;

  }

  stockSheet
    .getRange(
      1,
      1,
      stockData.length,
      stockData[0].length
    )
    .setValues(
      stockData
    );

  inventorySheet
    .getRange(
      9,
      2,
      inventoryData.length,
      7
    )
    .setValues(
      inventoryData
    );

  SpreadsheetApp.getUi().alert(
    "Inventory updated successfully."
  );

}
function startNewMonth() {

  const response =
    SpreadsheetApp
      .getUi()
      .alert(
        "Start New Month",
        "This will reset monthly inventory data.\n It is recommended to save a copy of this file before resetting.\n\nThis action cannot be undone. Continue?",
        SpreadsheetApp
          .getUi()
          .ButtonSet
          .YES_NO
      );

  if (
    response !==
    SpreadsheetApp
      .getUi()
      .Button.YES
  ) {

    return;

  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const stockSheet =
    ss.getSheetByName(
      "STOCK_COUNT"
    );

  const data =
    stockSheet
      .getDataRange()
      .getValues();

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const actualStock =
      Number(
        data[i][6]
      ) || 0;

    data[i][2] =
      actualStock;

    data[i][3] = 0;

    data[i][4] = 0;

    data[i][5] = 0;

    data[i][6] = 0;

    data[i][7] = 0;

  }

  stockSheet
    .getRange(
      1,
      1,
      data.length,
      data[0].length
    )
    .setValues(
      data
    );

  loadInventory();

  SpreadsheetApp.getUi().alert(
    "New inventory period started successfully."
  );

}



## Author

Ethel Anne B. Topacio
