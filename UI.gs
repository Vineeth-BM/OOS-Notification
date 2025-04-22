/**
 * Combined Email and Slack Out-of-Stock Notification System - ui.gs
 * Custom menu and UI interactions
 * @version 2.0
 * @lastModified 2025-04-17
 */

/**
 * Creates a custom menu in the spreadsheet UI when the document is opened
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('OOS Notifications')
      .addItem('Run OOS Notification Process', 'main')
      .addSeparator()
      .addSubMenu(ui.createMenu('Setup')
          .addItem('Initial Setup', 'initialSetup')
          .addItem('Set Slack Token', 'promptForSlackToken')
          .addItem('Get Slack Channel IDs', 'getSlackChannelIds')
          .addItem('Create Weekly Trigger', 'createWeeklyTrigger'))
      .addSeparator()
      .addItem('Test System', 'testSystem')
      .addItem('View Email Tracking Stats', 'showTrackingStats')
      .addToUi();
}

/**
 * Shows an alert if running from UI
 */
function showAlert(title, message) {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(title, message, ui.ButtonSet.OK);
  } catch (e) {
    // Running from trigger, just log the message
    Logger.log(`Alert (not shown): ${title} - ${message}`);
  }
}

/**
 * Prompts for Slack token and saves it
 */
function promptForSlackToken() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Set Slack API Token',
    'Please enter your Slack Bot User OAuth Token (starts with xoxb-):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() == ui.Button.OK) {
    var token = response.getResponseText();
    if (token && token.startsWith('xoxb-')) {
      PropertiesService.getScriptProperties().setProperty("SLACK_API_TOKEN", token);
      ui.alert('Success', 'Slack token saved successfully', ui.ButtonSet.OK);
    } else {
      ui.alert('Error', 'Invalid token format. Token should start with "xoxb-"', ui.ButtonSet.OK);
    }
  }
}

/**
 * Displays tracking statistics in a modal dialog
 */
function showTrackingStats() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const trackingSheet = spreadsheet.getSheetByName(CONFIG.TRACKING_SHEET_NAME);
    
    if (!trackingSheet) {
      showAlert('Error', 'Tracking sheet not found. Please run the system at least once to create it.');
      return;
    }
    
    // Get all tracking data
    const data = trackingSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      showAlert('Info', 'No tracking data found. Send some emails first.');
      return;
    }
    
    // Calculate statistics
    const totalEmails = data.length - 1; // Subtract header row
    let openedEmails = 0;
    let totalViews = 0;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][5] === 'Yes') { // "Opened" column
        openedEmails++;
        totalViews += (parseInt(data[i][6]) || 0); // "Views" column
      }
    }
    
    const openRate = (openedEmails / totalEmails * 100).toFixed(2);
    const avgViews = (totalViews / (openedEmails || 1)).toFixed(2);
    
    // Create HTML for the modal
    const htmlOutput = HtmlService
      .createHtmlOutput(`
        <h2>Email Tracking Statistics</h2>
        <div style="margin: 20px 0;">
          <p><strong>Total Emails Sent:</strong> ${totalEmails}</p>
          <p><strong>Emails Opened:</strong> ${openedEmails}</p>
          <p><strong>Open Rate:</strong> ${openRate}%</p>
          <p><strong>Total Views:</strong> ${totalViews}</p>
          <p><strong>Average Views per Opened Email:</strong> ${avgViews}</p>
        </div>
        <button onclick="google.script.host.close()">Close</button>
      `)
      .setWidth(400)
      .setHeight(300);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Email Tracking Statistics');
    
  } catch (error) {
    showAlert('Error', `Failed to display tracking stats: ${error.message}`);
  }
}
