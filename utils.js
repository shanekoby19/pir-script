/**
 * Returns the current month as a string.
 */
 function getMonth() {
    const months = ["January","February","March","April","May","June","July","August","September","October","November","December"]
    const month = new Date().getMonth(); 
    return months[month]
  }
  
  /**
   * Sets the header rows for a given sheet.
   */
  function setHeaderRows(sheet) {
    sheet.getRange(1,1).setValue("Child Name")
    sheet.getRange(1,2).setValue("Program")
    sheet.getRange(1,3).setValue("Center")
    sheet.getRange(1,4).setValue("Classroom")
    sheet.getRange(1,5).setValue("Family Advocate")
    sheet.getRange(1,6).setValue("Missing Item")
    sheet.getRange(1,7).setValue("Clean-up Guidance");
    sheet.getRange(1,8).setValue("Explanation Options")
  }
  
  /**
   * Creates a new sheet with the given name.
   * @returns - A sheet object that represent the newly created sheet.
   */
  function createNewSheet(name) {
    if (ss.getSheetByName(name)) {
      ss.deleteSheet(ss.getSheetByName(name))
    }
    return ss.insertSheet().setName(name)
  }
  
  /**
   * Returns the calculated date as a string with the form MM-DD-YYYY
   * @param daysOffset - The number of days in the past you want the date to represent. (Use 0 for today's date).
   */
  function getStrDate(daysOffset) {
    // Gets the current date given the offset of days.
    const date = new Date(Date.now() - (daysOffset * 24 * 60 * 60 *1000));
    const day = date.getDate() > 9 ? date.getDate() : date.getDate().toString().padStart(2, `0`)
    const month = date.getMonth() + 1 > 9 ? date.getMonth() + 1: (date.getMonth() + 1).toString().padStart(2, `0`)
    const year = date.getFullYear();
    return `${month}/${day}/${year}`;
  }
  
  /**
   * Automatically resizes the columns of the spreadsheet.
   * @param sheet - The sheet you want to resize columns on.
   */
  const autoFitColumns = function(sheet) {
    for(let i=1; i <= sheet.getLastColumn(); i++) {
      sheet.autoResizeColumn(i);
    }
  }
  
  /**
   * Creates a unique link to the child's profile in Shine. Includes the corresponding insight tab where the PIR metric lives.
   * @param cleanUpMonth - A sheet that represents the NORMALIZED pir data that needs to be cleaned up by staff.
   */
  function setLiveLinks(cleanUpSheet) {
    const auditSchedule = ss.getSheetByName(`Audit Schedule`);
  
    const links = cleanUpSheet.getRange(2,1,cleanUpSheet.getLastRow() - 1).getRichTextValues().flat(1);
    const missingItems = cleanUpSheet.getRange(2, 6, cleanUpSheet.getLastRow() - 1).getValues().flat(1);
    const auditMetrics = auditSchedule.getRange(1, 1, auditSchedule.getLastRow()).getValues().flat(1);
    const insightTabs = auditSchedule.getRange(1, 3, auditSchedule.getLastRow()).getValues().flat(1);
    links.forEach((link, index) => {
      const itemIndex = auditMetrics.indexOf(missingItems[index]);
      const insightTab = insightTabs[itemIndex];
      const linkUrl = link.getLinkUrl().replace(`Enrollment`, insightTab);
      const text = link.getText()
      const value = SpreadsheetApp.newRichTextValue()
      .setText(text)
      .setLinkUrl(0, text.length-1, linkUrl)
      .build();
      cleanUpSheet.getRange(index + 2, 1).setRichTextValue(value);
    });
  }
  
  
  /**
   * Sends an email to the advocate letting them know what service metrics need cleaning up.
   */
  function sendAdvocateEmails() {
    // Define the data sheets
    const advocateEmailerList = ss.getSheetByName('Advocate Emailer List');
    const advocateSheets = ss.getSheets().filter(sheet => sheet.getName().includes('FA') && !sheet.getName().includes('(i)'));
    const participants = advocateEmailerList.getRange(2, 1, advocateEmailerList.getLastRow() - 1).getValues().flat(1);
  
    // Validate that the user wants to send emails.
    const response = validateEmailSend(participants);
  
    // If the user responds with a "YES", send emails.
    if(response === SpreadsheetApp.getUi().Button.YES) {
      advocateSheets.forEach(advocateSheet => {
        const sheetName = advocateSheet.getName();
  
        // Remove the "FA - " prefix from the sheet name.
        const advocateName = advocateSheet.getName().slice(5, sheetName.length);
  
        // Get the advocate first name. String Start until the first space.
        const advocateFirstName = advocateName.slice(0, advocateName.indexOf(' '));
  
        // Get the number of completed explanations.
        const explanations = advocateSheet.getRange(2, 8, advocateSheet.getLastRow() - 1).getValues().flat(1);
        const completedExplanations = explanations.filter(explanation => explanation !== '').length;
        const incompleteExplanations = explanations.length - completedExplanations;
  
        const completionPercentage = (completedExplanations / explanations.length) * 100;
  
        // Get the url of the current advocate tab.
        const url = (`https://docs.google.com/spreadsheets/d/${ss.getId()}/edit#gid=${advocateSheet.getSheetId()}`);
  
  
        const emailDataPoints = advocateEmailerList.getRange(2, 1, advocateEmailerList.getLastRow() - 1, advocateEmailerList.getLastColumn()).getValues();
  
        // Filter data points from "Advocate Emailer List" to only include the current advocate.
        const emailDataPoint = emailDataPoints.filter(emailDataPoint => emailDataPoint[0] === advocateName);
  
  
        // If the current advocate was found on the "Advocate Emailer List" send the email to everyone in the "Email" column.
        if(emailDataPoint.length > 0) {
          // Create a new HTML Template 
          const template = HtmlService.createTemplateFromFile('advocateTemplate');
          const data = {
            advocateFirstName: advocateFirstName,
            explanationCount: explanations.length,
            completedExplanations: completedExplanations,
            className: completionPercentage > 80 ? 'good': 'bad',
            completionPercentage: completionPercentage.toFixed(2),
            url: url
          }
          
          template.data = data;
  
          // Evaluate the template with the given data.
          const message = template.evaluate().getContent();
  
          // Send the email.
          GmailApp.sendEmail(emailDataPoint[0][1], 
                            `${advocateName} - PIR Audit Clean-Up`, 
                            '', 
          {
            htmlBody: message,
            from: Session.getActiveUser().getEmail()
          });
        }
      }); 
    }
  };
  
  /**
   * Sends an email to the leadership team with the current weeks data.
   */
  function sendLeadershipEmails() {
    // Define the data sheets
    const leadershipEmailerList = ss.getSheetByName('Leadership Email List');
    const historicalDataSheet = ss.getSheetByName('Historical Data');
    const serviceMetricMaster = ss.getSheetByName('Service Metric Master Clean-up');
  
    // Get the email data and create the emailTo field.
    const emailData = leadershipEmailerList.getRange(2, 1, leadershipEmailerList.getLastRow() - 1).getValues().flat(1);
    let emailTo = '';
    emailData.forEach(email => emailTo += `${email},`);
  
    // Validate that the user wants to send emails.
    const response = validateEmailSend(emailData);
    const lastRow = historicalDataSheet.getLastRow();
    // If the user responds with a "YES", send emails.
    if(response === SpreadsheetApp.getUi().Button.YES) {
      // Get the number of completed explanations.
        const missingExplanations = historicalDataSheet.getRange(lastRow, 4).getValue();
        const thisWeeksExplanations = historicalDataSheet.getRange(lastRow, 3).getValue();
        const focusArea1 = historicalDataSheet.getRange(lastRow, 6).getValue();
        const focusArea2 = historicalDataSheet.getRange(lastRow, 7).getValue();
        const focusArea3 = historicalDataSheet.getRange(lastRow, 8).getValue();
  
        // Get the url of the current advocate tab.
        const url = (`https://docs.google.com/spreadsheets/d/${ss.getId()}/edit#gid=${serviceMetricMaster.getSheetId()}`);
  
        // If anyone is on the email list send the email to all participants.
        if(emailTo !== '') {
          // Create a new HTML Template 
          const template = HtmlService.createTemplateFromFile('leadershipTemplate');
          const data = {
            date: getStrDate(0),
            program,
            missingExplanations,
            thisWeeksExplanations,
            focusArea1,
            focusArea2,
            focusArea3,
            url,
          }
          
          template.data = data;
  
          // Evaluate the template with the given data.
          const message = template.evaluate().getContent();
  
          // Send the email.
          GmailApp.sendEmail(emailTo, 
                            `${program} - ${getStrDate(0)} - PIR Audit Progress`,
                            '',
          {
            htmlBody: message,
            from: Session.getActiveUser().getEmail()
          });
        }
    }
  }
  
  /**
   * Validates email sending by showing the user a confirmation popup. If yes, send email if no then don't.
   * @param participants - A list of the names of participants who will receive emails.
   * @returns - The response from the popup. This will dictact if emails are sent or not.
   */
  function validateEmailSend(participants) {
    // Define the alert message given the participants.
    let message = `You are about to send out ${participants.length} emails to:\n\n`;
    participants.forEach(participant => message += `${participant}\n`);
    message += `\nAre you sure you want to do this?`
  
    // Return the users response.
    return SpreadsheetApp.getUi().alert(message, SpreadsheetApp.getUi().ButtonSet.YES_NO);
  }
  
  
  /**
   * Deletes any sheet that contains the letters 'FA'
   */
  function deleteAdvocateSheets() {
    ss.getSheets().forEach(sheet => sheet.getName().includes('FA') ? ss.deleteSheet(sheet) : undefined);
  } 
  
  
  
  
  
  
  
  
  
  
  
  