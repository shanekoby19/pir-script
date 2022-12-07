// GLOBAL VARIABLES
const program = "ALCC Enrolled"
const ss = SpreadsheetApp.getActiveSpreadsheet()
const staticColumns = [1,2,3,4,5]

/**
 * Creates a ui dropdown menu whenever the google sheet is opened by a user.
 */
function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('PIR Audit Reports')
  .addItem(`Run ${getMonth()} Information Audit`, `thisMonthsInformationAudit`)
  .addItem(`Run ${getMonth()} Service Audit`, `thisMonthsServiceAudit`)
  .addItem('Send Emails to Advocates', 'sendAdvocateEmails')
  .addItem('Send Emails to Leadership', 'sendLeadershipEmails')
  .addToUi();
}

/**
 * Serves as the entry point for the scripts execution. This function will be used for automation only.
 * Manual execution can be found in the google sheet "PIR Audit Reports" dropdown menu.
 */
function main() {
  let sheet = ss.getSheetByName(`dataIn`); // Select a tab
  storeFileOnDrive(getAttachment(`${program} PIR Audit`, 1), `${program} - PIR Audit - ${getStrDate(0)}`);
  convertExceltoGoogleSpreadsheet(`${program} - PIR Audit - ${getStrDate(0)}`); // Live Links
  writeFileToSpreadsheet(sheet, `${program} - PIR Audit - ${getStrDate(0)}`); // Store data on selected tab.
  removeMetaData(sheet, 11, 2);
}

/**
 * Runs an audit report for the service metrics.
 */
function thisMonthsServiceAudit() {
  // Create the audit report sheet and figure out which metrics should be on this report.
  const auditMonth = createNewSheet(`${getMonth()} Service Audit Report`);
  const auditColumns = getAuditColumns("Service");

  // Import the PIR metrics that correspond to this months audit and set a completion percentage.
  importColumns(auditMonth, auditColumns);
  setAuditPercentage(auditMonth);

  // If the service metric master sheet already exists, delete it.
  const exists = ss.getSheets().filter(sheet => sheet.getName() === 'Service Metric Master Clean-up').length === 1;
  if(exists) {
    ss.deleteSheet(ss.getSheetByName('Service Metric Master Clean-up'));
  }

  // Define the service metric master sheet and create a new clean up report.
  const serviceMetricMaster = createNewSheet('Service Metric Master Clean-up');
  // Insert an additional 9000 rows, so we have enough space for any month audit.
  serviceMetricMaster.insertRowsAfter(1000, 9000);
  serviceMetricMaster.protect();
  createCleanUpReport(auditMonth, serviceMetricMaster);
  setLiveLinks(serviceMetricMaster);

  // Track audit explanations
  appendUniqueKeys(serviceMetricMaster);
  portUniqueKeys(serviceMetricMaster);
  createAdvocateSheets();

  // Create New Historical Record
  updateHistoricalDataSheet(); 
  
  // Hide anything that isn't an FA sheet.
  ss.getSheets().filter(sheet => !sheet.getName().includes('FA') && !sheet.getName().includes('Service Metric Master Clean-up')).forEach(sheet => sheet.hideSheet());

  // Add data back-ups
  backupUniqueKeys();

  // Autofit all columns in the workbook
  ss.getSheets().forEach(sheet => {
    if(!sheet.isSheetHidden()) {
      autoFitColumns(sheet);
    }
  })
}

/**
 * Runs an audit report for the information metrics.
 */
function thisMonthsInformationAudit() {
  // Create the audit report sheet and figure out which metrics should be on this report.
  const auditMonth = createNewSheet(`${getMonth()} Information Audit Report`);
  const auditColumns = getAuditColumns("Information");

  // Import the PIR metrics that correspond to this months audit and set a completion percentage.
  importColumns(auditMonth, auditColumns);
  setAuditPercentage(auditMonth);

  // Create a new clean up sheet and start to normallize the data from the current audit report.
  const cleanUpMonth = createNewSheet(`${getMonth()}  Clean-up`);
  createCleanUpReport(auditMonth, cleanUpMonth);
  // Delete the Explanation Option column --> this is not need for the information metrics.
  cleanUpMonth.deleteColumn(cleanUpMonth.getLastColumn());
  setLiveLinks(cleanUpMonth);

  // Autofit all columns in the workbook
  ss.getSheets().forEach(sheet => {
    if(!sheet.isSheetHidden()) {
      autoFitColumns(sheet);
    }
  })
}


/**
 * @param auditType - A string that represents the type of audit you are performing. Either "Service" or "Information"
 * @return - An array of indexes to columns that will be audited in the generated report. 
 */
function getAuditColumns(auditType) {
  const dataIn = ss.getSheetByName("dataIn");
  const currentMonth = getMonth();
  const auditColumns = [];

  // Push the first five columns of the PIR Report onto our data sheet.
  // [Name,	Program,	Center,	Classroom,	Family Advocate] // Index is pushed not the name.
  auditColumns.push(...staticColumns);

  // Filter the audit data by the type of audit you are performing.
  const auditSchedule = ss.getSheetByName("Audit Schedule");
  const auditData = auditSchedule.getRange(2,1,auditSchedule.getLastRow()-1,5).getValues()
                                 .filter(row => row[4] === auditType);

  // row[1] === "Monthly PIR Audit Distribution" Column B on the "Audit Schedule" tab.
  const filteredAuditData = auditData.filter(row => row[1] === currentMonth);

  // A list of all the PIR metrics from the "Data In" tab. (Data In headers)
  const allCols = dataIn.getRange(1, 1, 1,dataIn.getLastColumn()).getValues().flat(1)
  filteredAuditData.forEach(row => {
    // Push the column index of the columns to be audited onto an array.
    auditColumns.push(allCols.indexOf(row[0]) + 1);
  });

  return auditColumns;
}

/**
 * Import all columns that correspond to this months audit. Check the audit schedule for more information.
 * @param auditMonth - A sheet object that represents the current audit month tab.
 * @param colList - A list of column indexes that will be included on the generate audit report. 
 */
function importColumns(auditMonth, colList) {
  const dataIn = ss.getSheetByName("dataIn");

  // Loop through the dataIn tab and retrieve all the data needed to compose this months audit report.
  colList.forEach((value,index) => {
    // Copies the Shine live link to child that is in column 1 of the "dataIn" tab.
    if (index === 0) {
      // Copy and paste entire columns from dataIn to the current audit sheet.
      let copyRange = dataIn.getRange(2,value,dataIn.getLastRow()).getRichTextValues()
      let pasteRange = auditMonth.getRange(2, index+1,dataIn.getLastRow())
      pasteRange.setRichTextValues(copyRange)
      auditMonth.getRange(1,1).setValue("Child Name");
    }
    else {
      let copyRange = dataIn.getRange(1,value,dataIn.getLastRow()).getValues()
      let pasteRange = auditMonth.getRange(1, index+1,dataIn.getLastRow())
      pasteRange.setValues(copyRange)
    }
  })
}

/**
 * Add a percentage column to this months audit sheet. This percentage represents how complete each child's audit is for the given 
 * audit report.
 * @param auditMonth - A sheet object that represents the current audit month tab.
 */
function setAuditPercentage(auditMonth) {
  // Get all the data that exists of this months audit spreadsheet.
  const auditData = auditMonth.getRange(2,1,auditMonth.getLastRow()-1,auditMonth.getLastColumn()).getValues();

  // Find the last column to place the percentage in.
  const destColumn = auditMonth.getLastColumn() + 1;

  // Set the header for that column.
  auditMonth.getRange(1,destColumn).setValue("Percentage Incomplete");

  // Loop through each row of the current audit month and set a completed percentage. Yes / Number of metrics for the month.
  auditData.forEach((row, index) => {
    auditMonth.getRange(index + 2, destColumn)
              .setValue(`${Math.floor((row.filter(value => value === "Yes").length) / (row.length - 5) * 100)}%`);
  })
}

/**
 * Creates the clean up report for the month.
 * @param auditSheet - A sheet object that represents the current audit month tab.
 * @param cleanUpSheet - A sheet object that reporesents the current clean up report tab.
 */
function createCleanUpReport(auditSheet, cleanUpSheet) {
  // Set the header rows for the clean up sheet.
  setHeaderRows(cleanUpSheet);

  // Define the data.
  const richAuditColumns = auditSheet.getRange(1, 1, auditSheet.getLastRow()).getRichTextValues().flat(1);
  const auditData = auditSheet.getRange(1, 1, auditSheet.getLastRow(), auditSheet.getLastColumn()).getValues();
  let counter = 1;

  // Loop through 
  auditData.forEach((row, auditColumnIndex) => {
    // Exclude the header.
    if (auditColumnIndex !== 0) {
      // If the child has any PIR metric due, append them to the clean up report.
      if (row.filter(value => value === "Yes").length > 0) {
        // Loop through every past due metric and append it to the clean up report. NORMALIZATION.
        row.forEach((value,index) => {
          if (value === "Yes") {
            counter++
            cleanUpSheet.getRange(counter,1).setRichTextValue(richAuditColumns[auditColumnIndex])
            cleanUpSheet.getRange(counter,2).setValue(row[1])
            cleanUpSheet.getRange(counter,3).setValue(row[2])
            cleanUpSheet.getRange(counter,4).setValue(row[3])
            cleanUpSheet.getRange(counter,5).setValue(row[4])
            cleanUpSheet.getRange(counter,6).setValue(auditData[0][index]); 
            cleanUpSheet.getRange(counter,7).setValue(`=vlookup(F${counter},'Audit Schedule'!A:D,4,false)`);
            cleanUpSheet.getRange(counter,7).setValue(cleanUpSheet.getRange(counter, 7).getValue());
            cleanUpSheet.getRange(counter,6).setValue(`=HYPERLINK(G${counter}, "${auditData[0][index]}")`); 
          }
        }) 
      }
    }
  });
}

function appendUniqueKeys(sheet) {
  const uniqueKeys = ss.getSheetByName("Unique Keys")

  const uniqueKeyData = uniqueKeys.getRange(2,1, sheet.getLastRow()-1).getValues().flat(1)
  const data = sheet.getRange(2,1, sheet.getLastRow()-1, sheet.getLastColumn()).getValues()



  data.forEach(row => {
    const key = `${row[0]} - ${row[5]}`
    if (!uniqueKeyData.includes(key)) {
      uniqueKeys.getRange(uniqueKeys.getLastRow()+1,1).setValue(key)
    }
  })
}

function portUniqueKeys(serviceMetricMaster) {
  // FROM ADVOCATES TO UNIQUE KEYS
  const uniqueKeys = ss.getSheetByName("Unique Keys")

  let uniqueKeyData = uniqueKeys.getRange(2, 1, uniqueKeys.getLastRow()-1).getValues().flat(1);
  let uniqueKeyExplanations = uniqueKeys.getRange(2, 2, uniqueKeys.getLastRow()-1).getValues().flat(1);


  ss.getSheets().forEach(sheet => {
    if(sheet.getName().slice(0, 2) === "FA") {
      const data = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
      data.forEach(row => {
        if(row[7] !== "") {
          const key = `${row[0]} - ${row[5]}`;
          uniqueKeys.getRange(uniqueKeyData.indexOf(key) + 2, 2).setValue(row[7]);
        }
      })
    };
  })



  // FROM UNIQUE KEYS TO MASTER
  uniqueKeyData = uniqueKeys.getRange(2, 1, uniqueKeys.getLastRow()-1).getValues().flat(1);
  uniqueKeyExplanations = uniqueKeys.getRange(2, 2, uniqueKeys.getLastRow()-1).getValues().flat(1);
  const data = serviceMetricMaster.getRange(2,1, serviceMetricMaster.getLastRow()-1, serviceMetricMaster.getLastColumn()).getValues()


  data.forEach((row, index) => {
    const key = `${row[0]} - ${row[5]}`;
    const explanationIndex = uniqueKeyData.indexOf(key);
    
    // INDEX is found
    if(uniqueKeyExplanations[explanationIndex] !== "") {
      serviceMetricMaster.getRange(index + 2, 8).setValue(uniqueKeyExplanations[explanationIndex]);
    }

  })
}

/**
 * Creates a tab for each advocate consisting of only the clean up data they need to receive.
 */
function createAdvocateSheets() {
  const serviceMetricMaster = ss.getSheetByName('Service Metric Master Clean-up');

  // Remove any pre-existing data valadations before re-writing to the sheet.
  serviceMetricMaster.getRange(1, 1, serviceMetricMaster.getLastRow(), serviceMetricMaster.getLastColumn()).clearDataValidations();

  // Get all the data on the service metric master sheet.
  const data = serviceMetricMaster.getRange(1, 1, serviceMetricMaster.getLastRow(), serviceMetricMaster.getLastColumn()).getValues();

  // Get a unique list of advocates for this report.
  const uniqueAdvocates = [... new Set(serviceMetricMaster.getRange(2, 5, serviceMetricMaster.getLastRow() - 1).getValues().flat(1))];

  // Loop through each advocate and filter for only their data.
  uniqueAdvocates.forEach(uniqueAdvocate => {
    const filteredData = data.filter((row, index) => {
      //HEADER
      if(index === 0) {
        return true;
      }

      if(row[4] === uniqueAdvocate) {
        return true;
      }
    });

    // CREATE NEW ADVOCATE SHEET
    const advocateSheet = createNewSheet(`FA - ${uniqueAdvocate}`);
    advocateSheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);



    // REPLACE PLAIN TEXT NAME WITH LIVE LINK CHILD NAME

    // Get the child names from the advocate sheet for comparison.
    const childNames = advocateSheet.getRange(2, 1, advocateSheet.getLastRow() - 1).getValues().flat(1);

    // Get the advocate names from the service metric master clean up sheet as plain text..
    const advocates = serviceMetricMaster.getRange(2, 5, serviceMetricMaster.getLastRow() - 1).getValues().flat(1);

    // Get the rich text data from the service metric master clean up sheet.
    const richTextData = serviceMetricMaster.getRange(2, 1, serviceMetricMaster.getLastRow() - 1).getRichTextValues();

    // Filter the rich text data for any child under the given advocate caseload.
    // Criteria 1: Does the child name exist on this advocates sheet?
    // Criteria 2: Is the child assigned to this advocate. 
    // Each criteria corresponds to a filter condition.
    const filteredRichTextData = richTextData.filter((row, index) => childNames.indexOf(row[0].getText()) >= 0 && 
                                                                     advocates[index] === uniqueAdvocate);

    // Paste the rich text data.
    advocateSheet.getRange(2, 1, advocateSheet.getLastRow() - 1).setRichTextValues(filteredRichTextData);
  


    // REPLACE THE PLAIN TEXT METRIC WITH A HELPFUL LINK.

    // Get the child names from the advocate sheet for comparison.
    const missingItems = advocateSheet.getRange(2, 6, advocateSheet.getLastRow() - 1).getValues().flat(1);

    const formulaData = serviceMetricMaster.getRange(2, 6, serviceMetricMaster.getLastRow() - 1).getFormulas();
    const filteredFormulaData = formulaData.filter((row, index) => {
      // Parse the formula to get the start and end of the metric.
      const metricStart = row[0].indexOf('"') + 1;
      const metricEnd = row[0].indexOf('")');

      // Slice the string to get only the metric name.
      const metric = row[0].slice(metricStart, metricEnd);

      // Filter the data on service metric master where
      // Criteria 1: The metric name must match.
      // Criteria 2: The advocate name must match.
      return missingItems.indexOf(metric) >= 0 && advocates[index] === uniqueAdvocate
    })
    // Replace the old row definition with the new row definition.
    .map((row, index) => {
      // Get the start and end of the row number to replace from service metric master.
      const rangeStart = row[0].indexOf("G") + 1;
      const rangeEnd = row[0].indexOf(", ");
      // Return the new value to be stored in the mapped array.
      return [row[0].replace(row[0].slice(rangeStart, rangeEnd), `${index + 2}`)];
    });

    // Paste the formula data.
    advocateSheet.getRange(2, 6, advocateSheet.getLastRow() - 1).setFormulas(filteredFormulaData);


    // HIDE UNUSED DATA - Family Advocate, Clean-up Guidance (columns e and g)
    advocateSheet.hideColumns(5);
    advocateSheet.hideColumns(7);

    
    determineExplanationOptions(advocateSheet);
  })
}


function determineExplanationOptions(advocateSheet) {
  const pirExplanations = ss.getSheetByName("PIR Explanations");
  const serviceMetricIndicatorsNeedingCleanup = advocateSheet.getRange(2, 6, advocateSheet.getLastRow() - 1).getValues().flat();
  const pirExplanationData = pirExplanations.getRange(2,1,pirExplanations.getLastRow()-1, 2).getValues();
  
  serviceMetricIndicatorsNeedingCleanup.forEach((metric, index) => {
    let pirExplanationOptions = [];
    const filteredData = pirExplanationData.filter(row => row[0] === metric);
    
    // Get all the pir explanation options for the given metric.
    filteredData.forEach(row => {
      pirExplanationOptions.push(row[1]);
    });

    const dropdown = advocateSheet.getRange(`H${index + 2}`);
    const rule = SpreadsheetApp.newDataValidation()
                               .requireValueInList(pirExplanationOptions)
                               .setAllowInvalid(false)
                               .setHelpText('Please select an option from the dropdown list.')
                               .build();
    dropdown.setDataValidation(rule);

  });
}

/**
 * Creates a new record in the historical data sheet with this weeks values.
 */
function updateHistoricalDataSheet() {
  const historicalDataSheet = ss.getSheetByName('Historical Data');
  const serviceMetricMaster = ss.getSheetByName('Service Metric Master Clean-up');
  const explanationData = serviceMetricMaster.getRange(2, 8, serviceMetricMaster.getLastRow() - 1).getValues().flat(1);
  const currentRow = historicalDataSheet.getLastRow() + 1;

  // Set the Week Of
  historicalDataSheet.getRange(currentRow, 1).setValue(getStrDate(0));

  // Set the Total Explanations
  historicalDataSheet.getRange(currentRow, 2).setValue(explanationData.length);
  
  // Set the Completed Explanations for the current week
  let lastWeekValue = historicalDataSheet.getRange(currentRow - 1, 3).getValue();
  let totalExplanations = explanationData.filter(explanation => explanation !== '').length;
  if(typeof lastWeekValue !== 'string') {
    historicalDataSheet.getRange(currentRow, 3).setValue(totalExplanations - lastWeekValue);
  }
  else {
    historicalDataSheet.getRange(currentRow, 3).setValue(0);
  }

  // Set the Number of explanations still needed
  historicalDataSheet.getRange(currentRow, 4).setValue(explanationData.filter(explanation => explanation === '').length);

  // Set the change from last week
  lastWeekValue = historicalDataSheet.getRange(currentRow - 1, 4).getValue();
  thisWeekValue = historicalDataSheet.getRange(currentRow, 4).getValue();
  if(typeof lastWeekValue !== 'string') {
    historicalDataSheet.getRange(currentRow, 5).setValue(lastWeekValue - thisWeekValue);
  }
  else {
    historicalDataSheet.getRange(currentRow, 5).setValue(0);
  }

  // SET THE TOP 3 FOCUS AREAS FOR THE WEEK
  
  // Get the unique metrics
  const uniqueMetrics = [ ...new Set(serviceMetricMaster.getRange(2, 6, serviceMetricMaster.getLastRow() - 1).getValues().flat(1))];

  // Get all the data from "Service Metric Master Clean-up"
  const data = serviceMetricMaster.getRange(2, 1, serviceMetricMaster.getLastRow() - 1, serviceMetricMaster.getLastColumn())
                                  .getValues();

  // Filter for metrics with no explanations.
  const filteredData = data.filter(row => row[7] === '');
  
  // Return an object for each metric with a metric name and metric count.
  const metricData = uniqueMetrics.map(uniqueMetric => {
    return {
      name: uniqueMetric,
      count: filteredData.filter(row => row[5] === uniqueMetric).length
    }
  });

  // Rank the objects from first to last by importance.
  metricData.sort((obj1, obj2) => obj2.count - obj1.count);

  // Place them in the correct spots on the Historical Data sheet.
  // Set an offset so the columns will align with the index
  metricData.slice(0, 3).forEach((dataPoint, index) => {
    historicalDataSheet.getRange(currentRow, index + 6).setValue(`${dataPoint.name} - ${dataPoint.count} needing explanation`);
  });
}

/**
 * Creates a new data back-up file for the given partner.
 * This ensures that if the database (Unique Keys) is deleted we can restore to a last saved value
 */
function backupUniqueKeys() {
  // Location the Automated - MBI folder on the users drive or create it if it doesn't exists.
  const convertedSheetFolders = DriveApp.getFoldersByName(`${program} - Unique Keys Backups`);
  const folderFound = convertedSheetFolders.hasNext();
  let convertedSheetFolder = folderFound ? convertedSheetFolders.next() : DriveApp.createFolder(`${program} - Unique Keys Backups`);

  // Create a new google sheet to store the back-up data.
  const sheet = SpreadsheetApp.create(`Backup - ${getStrDate(0)}`);
  // Create a file for that sheet.
  const file = DriveApp.getFileById(sheet.getId());
  // Move the file to designatied folder
  file.moveTo(convertedSheetFolder);

  // Add the sheet to the newly created file and unhide it.
  const uniqueKeysBackup = ss.getSheetByName('Unique Keys').copyTo(sheet).setName(`Backup - ${getStrDate(0)}`);
  uniqueKeysBackup.showSheet();

  // Delete the default sheet created by the file creation. (Sheet 1)
  sheet.deleteSheet(sheet.getSheets()[0]);
}





