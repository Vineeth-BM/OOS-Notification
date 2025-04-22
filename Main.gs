/**
 * Combined Email and Slack Out-of-Stock Notification System - main.gs
 * Integrates all modules together
 * @version 2.0
 * @lastModified 2025-04-17
 */

/**
 * Main function to send out-of-stock notifications
 */
function main() {
  try {
    Logger.log('Starting out-of-stock notification process');
    
    // Access the OOS data sheet
    const oosData = getSheetData(CONFIG.OOS_SHEET_NAME);
    if (!oosData || oosData.length <= CONFIG.HEADER_ROW_COUNT) {
      showAlert('Warning', 'No OOS data found or only header row exists.');
      return;
    }
    
    // Access the sellers data sheet
    const sellersData = getSheetData(CONFIG.SELLERS_SHEET_NAME);
    if (!sellersData || sellersData.length <= CONFIG.HEADER_ROW_COUNT) {
      showAlert('Warning', 'No seller data found or only header row exists.');
      return;
    }
    
    // Log some information about the data
    Logger.log(`OOS data has ${oosData.length} rows (including headers)`);
    if (oosData.length > 1) {
      Logger.log(`Sample OOS item: ${JSON.stringify(oosData[1])}`);
    }
    
    // Ensure we have all necessary columns
    Logger.log('OOS data header: ' + oosData[0].join(', '));
    
    // Get the Slack channel mapping data
    const slackChannelsData = getSheetData(CONFIG.SLACK_CHANNELS_SHEET_NAME);
    const slackChannels = [];
    
    if (slackChannelsData && slackChannelsData.length > CONFIG.HEADER_ROW_COUNT) {
      // Skip header row and get channel mappings
      for (let i = CONFIG.HEADER_ROW_COUNT; i < slackChannelsData.length; i++) {
        const row = slackChannelsData[i];
        const recipientName = row[SLACK_COLUMNS.RECIPIENT_NAME];
        const channelId = row[SLACK_COLUMNS.CHANNEL_ID];
        
        if (channelId) {
          slackChannels.push({
            name: recipientName,
            id: channelId
          });
        }
      }
      Logger.log(`Found ${slackChannels.length} Slack channels for notifications`);
    } else {
      Logger.log('No Slack channels found or Slack_Channels sheet is missing');
    }
    
    // Check if Slack token is set
    if (CONFIG.SLACK.ENABLED) {
      const tokenExists = checkSlackToken();
      if (!tokenExists) {
        Logger.log("Warning: Slack integration is enabled but no token is set. Slack notifications will be skipped.");
        CONFIG.SLACK.ENABLED = false;
      }
    }
    
    // Process the data and send emails
    const stats = processData(oosData, sellersData);
    
    // Send Slack notifications to channels if enabled and channels are available
    if (CONFIG.SLACK.ENABLED && slackChannels.length > 0) {
      sendOOSNotificationsToSlack(oosData, slackChannels, stats);
    } else {
      // Remove other Slack notifications entirely
      // Don't send notifications for individual emails or summaries
      Logger.log("Only sending direct channel notifications, skipping other Slack messages");
    }
    
    // Log summary statistics
    Logger.log(`Process completed. Results: ${stats.processed} rows processed, ${stats.emailsSent} emails sent, ${stats.errors} errors`);
    
    showAlert('Success', 
              `Process completed:\n${stats.processed} OOS items processed\n${stats.emailsSent} emails sent\n${stats.errors} errors`);
    
    return stats;
    
  } catch (error) {
    Logger.log(`Critical error in main function: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    
    showAlert('Error', `An error occurred: ${error.message}`);
    throw error;
  }
}

/**
 * Processes data and sends emails for out-of-stock items
 * @param {Array} oosData 2D array of OOS data
 * @param {Array} sellersData 2D array of sellers data
 * @return {Object} Statistics about the process
 */
function processData(oosData, sellersData) {
  // Enable debug logging to see more information
  if (CONFIG.DEBUG_MODE) {
    Logger.log("Starting processing with " + oosData.length + " OOS items and " + sellersData.length + " sellers");
  }
  
  // Statistics tracking
  const stats = {
    processed: 0,
    emailsSent: 0,
    errors: 0
  };
  
  // Create template object
  let htmlTemplate;
  try {
    htmlTemplate = HtmlService.createTemplateFromFile(CONFIG.EMAIL_TEMPLATE);
    Logger.log('Email template loaded successfully');
  } catch (error) {
    Logger.log(`Error loading email template: ${error.message}`);
    stats.errors++;
    return stats;
  }
  
  // Get the business update from sheet
  const businessUpdate = getBusinessUpdate();
  
  // Create a map of sellers for quick lookup
  const sellers = {};
  for (let i = CONFIG.HEADER_ROW_COUNT; i < sellersData.length; i++) {
    const row = sellersData[i];
    const sellerId = row[SELLER_COLUMNS.SELLER_ID];
    
    if (!sellerId) continue;
    
    sellers[sellerId] = {
      id: sellerId,
      name: row[SELLER_COLUMNS.SELLER_NAME] || '',
      fullName: row[SELLER_COLUMNS.SELLER_FULL_NAME] || '',
      email: row[SELLER_COLUMNS.SELLER_EMAIL] || ''
    };
  }
  
  // For demonstration, we'll associate OOS items with sellers
  // In a real scenario, you'd need actual product-seller relationships
  // For now, we'll distribute OOS items evenly among sellers
  const sellerIds = Object.keys(sellers);
  if (sellerIds.length === 0) {
    Logger.log("No sellers found to assign OOS items to");
    return stats;
  }
  
  // Group OOS items by seller ID
  const sellerOOSItems = {};
  
  // Initialize empty array for each seller
  for (const sellerId of sellerIds) {
    sellerOOSItems[sellerId] = [];
  }
  
  // Group OOS items by subcategory for proper ranking
  const subcategories = {};
  
  // First pass: Group items by subcategory
  for (let i = CONFIG.HEADER_ROW_COUNT; i < oosData.length; i++) {
    const subcategory = oosData[i][OOS_COLUMNS.CATEGORY_SUB_CLUSTER] || 'Uncategorized';
    if (!subcategories[subcategory]) {
      subcategories[subcategory] = [];
    }
    
    subcategories[subcategory].push({
      rowIndex: i
    });
  }
  
  // Second pass: Assign subcategory ranks
  for (const subcategory in subcategories) {
    // Assign ranks
    subcategories[subcategory].forEach((item, index) => {
      // Rank starts from 1
      const subcategoryRank = index + 1;
      
      // Store the rank for later use when creating product objects
      oosData[item.rowIndex][3] = subcategoryRank; // Using index 3 for subcategory rank
    });
  }
  
  // Process all OOS items at once to create product objects
  const allProducts = [];
  
  // Create product objects for all OOS items
  for (let i = CONFIG.HEADER_ROW_COUNT; i < oosData.length; i++) {
    try {
      stats.processed++;
      
      // Extract product details
      const product = {
        rank: i - CONFIG.HEADER_ROW_COUNT + 1, // Overall rank based on row order
        subcategoryRank: oosData[i][3] || i - CONFIG.HEADER_ROW_COUNT + 1, // Use computed subcategory rank
        name: oosData[i][OOS_COLUMNS.PRODUCT_NAME] || 'Unknown Product',
        category: oosData[i][OOS_COLUMNS.CATEGORY_SUB_CLUSTER] || 'Uncategorized',
        grade: oosData[i][OOS_COLUMNS.BACKBOX_GRADE] || 'N/A'
      };
      
      allProducts.push(product);
      
    } catch (error) {
      Logger.log(`Error processing OOS row ${i+1}: ${error.message}`);
      stats.errors++;
      continue;
    }
  }
  
  // Distribute all products to all sellers (each seller gets ALL products)
  // This ensures no products are lost when processing multiple sellers
  for (const sellerId of sellerIds) {
    sellerOOSItems[sellerId] = [...allProducts]; // Give each seller a copy of all products
  }
  
  // Log how many products each seller has
  for (const sellerId in sellerOOSItems) {
    Logger.log(`Seller ${sellerId} has ${sellerOOSItems[sellerId].length} products`);
  }
  
  // Send emails to each seller with their assigned OOS items
  for (const sellerId in sellerOOSItems) {
    try {
      const seller = sellers[sellerId];
      const oosItems = sellerOOSItems[sellerId];
      
      // Skip if no email or no items
      if (!seller.email || !oosItems || oosItems.length === 0) {
        Logger.log(`Skipping seller ${sellerId}: No email or no items to process`);
        continue;
      }
      
      // Send the email
      const emailSent = sendNotificationEmail(htmlTemplate, seller, oosItems, businessUpdate);
      if (emailSent) {
        stats.emailsSent++;
        
        // Don't send Slack notifications for individual emails
        // Comment out the Slack notification code to remove these messages
        /*
        // Send a Slack notification after successful email
        if (CONFIG.SLACK && CONFIG.SLACK.ENABLED) {
          // Extract subcategories for Slack notification
          const subcategories = {};
          oosItems.forEach(product => {
            if (!subcategories[product.category]) {
              subcategories[product.category] = [];
            }
            subcategories[product.category].push(product);
          });
          
          // Send notification to Slack about email
          sendSlackNotification(seller, oosItems.length, subcategories);
        }
        */
      } else {
        stats.errors++;
      }
      
    } catch (error) {
      Logger.log(`Error sending email to seller ${sellerId}: ${error.message}`);
      stats.errors++;
    }
  }
  
  // Send a summary notification to Slack
  // if (CONFIG.SLACK && CONFIG.SLACK.ENABLED) {
  //   sendSlackSummary(stats);
  // }
  
  return stats;
}

/**
 * Run this once to set up everything
 */
function initialSetup() {
  // Create necessary sheets if they don't exist
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Check and create OOS List sheet
  if (!spreadsheet.getSheetByName(CONFIG.OOS_SHEET_NAME)) {
    const sheet = spreadsheet.insertSheet(CONFIG.OOS_SHEET_NAME);
    sheet.appendRow([
      'CATEGORY_SUB_CLUSTER', 
      'PRODUCT_NAME', 
      'BACKBOX_GRADE'
    ]);
    Logger.log(`Created ${CONFIG.OOS_SHEET_NAME} sheet`);
  }

  // Check and create Sellers List sheet
  if (!spreadsheet.getSheetByName(CONFIG.SELLERS_SHEET_NAME)) {
    const sheet = spreadsheet.insertSheet(CONFIG.SELLERS_SHEET_NAME);
    sheet.appendRow([
      'SELLER_ID', 
      'SELLER_NAME', 
      'SELLER_FULL_NAME', 
      'SELLER_EMAIL'
    ]);
    Logger.log(`Created ${CONFIG.SELLERS_SHEET_NAME} sheet`);
  }

  // Check and create Slack Channels sheet
  if (!spreadsheet.getSheetByName(CONFIG.SLACK_CHANNELS_SHEET_NAME)) {
    const sheet = spreadsheet.insertSheet(CONFIG.SLACK_CHANNELS_SHEET_NAME);
    sheet.appendRow([
      'RECIPIENT_NAME', 
      'CHANNEL_ID'
    ]);
    Logger.log(`Created ${CONFIG.SLACK_CHANNELS_SHEET_NAME} sheet`);
  }

  // Check and create Business Update sheet
  if (!spreadsheet.getSheetByName(CONFIG.BUSINESS_UPDATE_SHEET_NAME)) {
    const sheet = spreadsheet.insertSheet(CONFIG.BUSINESS_UPDATE_SHEET_NAME);
    Logger.log(`Created empty ${CONFIG.BUSINESS_UPDATE_SHEET_NAME} sheet. Please add your business update content.`);
  }

  Logger.log("Initial setup complete. Please run the following functions in order:");
  Logger.log("1. setSlackToken() - to set your Slack API token");
  Logger.log("2. getSlackChannelIds() - to get a list of available Slack channels");
  Logger.log("3. testSystem() - to verify everything is working correctly");
  Logger.log("4. setDailySchedule() or createWeeklyTrigger() - to set up automatic execution");

  showAlert('Setup Complete', 'Initial setup complete. Check the logs for next steps.');
}

/**
 * Test the notification system with comprehensive checks
 */
function testSystem() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Test System',
    'This will perform a complete system check and send test notifications. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response == ui.Button.YES) {
    try {
      // Create results array for reporting
      const testResults = [];
      
      // Check 1: Verify all required sheets exist
      testResults.push(checkRequiredSheets());
      
      // Check 2: Verify email template exists
      testResults.push(checkEmailTemplate());
      
      // Check 3: Check Slack token and connectivity
      testResults.push(checkSlackIntegration());
      
      // Check 4: Test sending an email
      testResults.push(testEmailSending());
      
      // Check 5: Test Slack notification
      testResults.push(testSlackNotification());
      
      // Display test results
      showTestResults(testResults);
      
    } catch (error) {
      ui.alert('Test Failed', `Unexpected error: ${error.message}`, ui.ButtonSet.OK);
    }
  }
}

/**
 * Check if all required sheets exist
 * @return {Object} Test result object
 */
function checkRequiredSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = [
    CONFIG.OOS_SHEET_NAME,
    CONFIG.SELLERS_SHEET_NAME,
    CONFIG.SLACK_CHANNELS_SHEET_NAME,
    CONFIG.BUSINESS_UPDATE_SHEET_NAME
  ];
  
  const missingSheets = [];
  
  for (const sheetName of requiredSheets) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      missingSheets.push(sheetName);
    }
  }
  
  if (missingSheets.length > 0) {
    return {
      name: "Sheet Check",
      passed: false,
      details: `Missing sheets: ${missingSheets.join(', ')}`
    };
  }
  
  return {
    name: "Sheet Check",
    passed: true,
    details: "All required sheets exist"
  };
}

/**
 * Check if email template exists
 * @return {Object} Test result object
 */
function checkEmailTemplate() {
  try {
    HtmlService.createTemplateFromFile(CONFIG.EMAIL_TEMPLATE);
    return {
      name: "Email Template",
      passed: true,
      details: `Template '${CONFIG.EMAIL_TEMPLATE}' found`
    };
  } catch (error) {
    return {
      name: "Email Template",
      passed: false,
      details: `Template '${CONFIG.EMAIL_TEMPLATE}' error: ${error.message}`
    };
  }
}

/**
 * Check Slack token and test connectivity
 * @return {Object} Test result object
 */
function checkSlackIntegration() {
  if (!CONFIG.SLACK.ENABLED) {
    return {
      name: "Slack Integration",
      passed: true,
      details: "Slack integration is disabled in configuration"
    };
  }
  
  // Check if token exists
  const token = PropertiesService.getScriptProperties().getProperty("SLACK_API_TOKEN");
  if (!token) {
    return {
      name: "Slack Integration",
      passed: false,
      details: "No Slack API token found. Please run setSlackToken() first."
    };
  }
  
  // Test Slack API connectivity
  try {
    const options = {
      "method": "get",
      "headers": {
        "Authorization": "Bearer " + token
      },
      "muteHttpExceptions": true
    };
    
    const response = UrlFetchApp.fetch("https://slack.com/api/auth.test", options);
    const responseData = JSON.parse(response.getContentText());
    
    if (responseData.ok) {
      return {
        name: "Slack Integration",
        passed: true,
        details: `Connected to Slack as: ${responseData.user} in team: ${responseData.team}`
      };
    } else {
      return {
        name: "Slack Integration",
        passed: false,
        details: `Error testing Slack connection: ${responseData.error}`
      };
    }
  } catch (error) {
    return {
      name: "Slack Integration",
      passed: false,
      details: `Exception testing Slack: ${error.message}`
    };
  }
}

/**
 * Test sending an email to the current user
 * @return {Object} Test result object
 */
function testEmailSending() {
  try {
    // Get template
    let htmlTemplate = HtmlService.createTemplateFromFile(CONFIG.EMAIL_TEMPLATE);
    
    // Create test data
    const testSeller = {
      id: 'TEST001',
      name: 'Test Seller',
      fullName: 'Test Seller Full Name',
      email: Session.getActiveUser().getEmail() // Send to current user
    };
    
    const testProducts = [
      {
        rank: 1,
        subcategoryRank: 1,
        name: 'Test Product 1',
        category: 'Test Category',
        grade: 'A'
      },
      {
        rank: 2,
        subcategoryRank: 1,
        name: 'Test Product 2',
        category: 'Another Category',
        grade: 'B'
      }
    ];
    
    // Send test email
    const emailSent = sendNotificationEmail(htmlTemplate, testSeller, testProducts, '<h3>Test Business Update</h3>');
    
    if (emailSent) {
      return {
        name: "Email Sending",
        passed: true,
        details: `Test email sent to ${testSeller.email}`
      };
    } else {
      return {
        name: "Email Sending",
        passed: false,
        details: "Failed to send test email"
      };
    }
  } catch (error) {
    return {
      name: "Email Sending",
      passed: false,
      details: `Error sending test email: ${error.message}`
    };
  }
}

/**
 * Test sending a Slack notification
 * @return {Object} Test result object
 */
function testSlackNotification() {
  if (!CONFIG.SLACK.ENABLED) {
    return {
      name: "Slack Notification",
      passed: true,
      details: "Slack integration is disabled in configuration"
    };
  }
  
  try {
    // Create test data
    const testSeller = {
      id: 'TEST001',
      name: 'Test Seller',
      fullName: 'Test Seller Full Name'
    };
    
    // Create test subcategories
    const testSubcategories = {
      "Test Category": [
        { name: "Test Product 1", grade: "A" },
        { name: "Test Product 2", grade: "A+" }
      ],
      "Another Category": [
        { name: "Test Product 3", grade: "B" }
      ]
    };
    
    // Check if we have any channels configured
    const slackChannelsData = getSheetData(CONFIG.SLACK_CHANNELS_SHEET_NAME);
    if (!slackChannelsData || slackChannelsData.length <= CONFIG.HEADER_ROW_COUNT) {
      return {
        name: "Slack Notification",
        passed: false,
        details: "No Slack channels configured in the Slack_Channels sheet"
      };
    }
    
    // Send test notification
    const notificationSent = sendSlackNotification(testSeller, 3, testSubcategories);
    
    if (notificationSent) {
      return {
        name: "Slack Notification",
        passed: true,
        details: "Test notification sent to Slack"
      };
    } else {
      return {
        name: "Slack Notification",
        passed: false,
        details: "Failed to send Slack notification"
      };
    }
  } catch (error) {
    return {
      name: "Slack Notification",
      passed: false,
      details: `Error sending Slack notification: ${error.message}`
    };
  }
}

/**
 * Show test results in a modal dialog
 * @param {Array} results Array of test result objects
 */
function showTestResults(results) {
  const ui = SpreadsheetApp.getUi();
  
  // Count passed/failed tests
  const passCount = results.filter(r => r.passed).length;
  const failCount = results.length - passCount;
  
  // Build HTML for the results
  let html = '<div style="font-family: Arial, sans-serif; padding: 10px;">';
  html += `<h2>System Test Results: ${passCount}/${results.length} Tests Passed</h2>`;
  
  // Add a summary section
  html += '<div style="margin-bottom: 20px;">';
  if (failCount === 0) {
    html += '<p style="color: green; font-weight: bold;">✅ All tests passed successfully!</p>';
  } else {
    html += `<p style="color: red; font-weight: bold;">❌ ${failCount} test(s) failed. See details below.</p>`;
  }
  html += '</div>';
  
  // Add details for each test
  for (const result of results) {
    const color = result.passed ? 'green' : 'red';
    const icon = result.passed ? '✅' : '❌';
    
    html += `<div style="margin-bottom: 15px; border-left: 4px solid ${color}; padding-left: 10px;">`;
    html += `<h3 style="margin: 0; color: ${color};">${icon} ${result.name}</h3>`;
    html += `<p style="margin-top: 5px;">${result.details}</p>`;
    html += '</div>';
  }
  
  html += '</div>';
  
  // Show the results in a modal dialog
  const htmlOutput = HtmlService
    .createHtmlOutput(html)
    .setWidth(500)
    .setHeight(400);
  
  ui.showModalDialog(htmlOutput, 'System Test Results');
}

/**
 * Creates a weekly trigger to run the notification system
 * RUN THIS FUNCTION ONCE to set up automatic execution
 */
function createWeeklyTrigger() {
  // Delete any existing triggers
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  // Create a new weekly trigger (runs every Monday at 9 AM)
  ScriptApp.newTrigger('main')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY) // Run on Mondays
    .atHour(9) // Run at 9 AM
    .create();
    
  Logger.log("Weekly trigger created to run on Mondays at 9 AM");
}
