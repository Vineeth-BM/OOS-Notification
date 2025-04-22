/**
 * Combined Email and Slack Out-of-Stock Notification System - core.gs
 * Contains configuration, utilities, and data access functions
 * @version 2.0
 * @lastModified 2025-04-17
 */

//--------------------------------------------------------------------
// CONFIGURATION SECTION
//--------------------------------------------------------------------

// Main configuration object
const CONFIG = {
  // Sheet names
  OOS_SHEET_NAME: 'OOS List',
  SELLERS_SHEET_NAME: 'Sellers List',
  BUSINESS_UPDATE_SHEET_NAME: 'Business Update',
  SLACK_CHANNELS_SHEET_NAME: 'Slack_Channels',
  
  // Email settings
  EMAIL_TEMPLATE: 'email',
  EMAIL_SUBJECT: '【バックマーケット_在庫切れ商品のご案内】{sellerName}様',  // Updated to Japanese
  EMAIL_SENDER_NAME: 'Back Market Support',
  EMAIL_SENDER_ADDRESS: 'support@backmarket.com',
  
  // Tracking configuration
  TRACKING_SHEET_NAME: 'OOS Email Tracking',
  WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbyu_xKGyVEERax9t21VG3-10oJyojoEKevoEEDBxOV_BTLDunOFTHzA-CPnuya_adO-/exec',
  
  // General settings
  HEADER_ROW_COUNT: 1,
  DEBUG_MODE: true,
  DEFAULT_TIMEZONE: 'Asia/Tokyo',
  
  // Slack configuration
  SLACK: {
    ENABLED: true,
    MAX_ITEMS_PER_MESSAGE: 30,
    NOTIFICATION_DELAY_MS: 1000, // Delay between Slack messages to avoid rate limiting
    VERBOSE_LOGGING: true // Enable detailed logging for Slack operations
  }
};

// Column indices for OOS List sheet
const OOS_COLUMNS = {
  CATEGORY_SUB_CLUSTER: 0,
  PRODUCT_NAME: 1,
  BACKBOX_GRADE: 2
};

// Column indices for Sellers List sheet
const SELLER_COLUMNS = {
  SELLER_ID: 0,
  SELLER_NAME: 1,
  SELLER_FULL_NAME: 2,
  SELLER_EMAIL: 3
};

// Column indices for Slack Channels sheet
const SLACK_COLUMNS = {
  RECIPIENT_NAME: 0,
  CHANNEL_ID: 1
};

//--------------------------------------------------------------------
// UTILITY FUNCTIONS
//--------------------------------------------------------------------

/**
 * Retrieves data from the specified sheet
 * @param {String} sheetName Name of the sheet
 * @return {Array} 2D array of sheet data
 */
function getSheetData(sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log(`Sheet "${sheetName}" not found!`);
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    Logger.log(`Retrieved ${data.length} rows from sheet ${sheetName}`);
    
    return data;
  } catch (error) {
    Logger.log(`Error retrieving sheet data from ${sheetName}: ${error.message}`);
    return null;
  }
}

/**
 * Gets business update from the Business Update sheet
 * @return {String} The business update content as HTML
 */
function getBusinessUpdate() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(CONFIG.BUSINESS_UPDATE_SHEET_NAME);
    
    if (!sheet) {
      Logger.log("Business Update sheet not found.");
      return '';
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 1) {
      Logger.log("Business Update sheet is empty.");
      return '';
    }
    
    // Get content from the first cell (assuming it contains HTML)
    const updateContent = data[0][0] || '';
    
    return updateContent;
  } catch (error) {
    Logger.log(`Error getting business update: ${error.message}`);
    return '';
  }
}

/**
 * Validates an email address format
 * @param {String} email The email address to validate
 * @return {Boolean} Whether the email is valid
 */
function isValidEmail(email) {
  if (!email || typeof email !== 'string') return false;
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Find the Slack channel ID for a seller
 * @param {string} sellerName The name of the seller
 * @return {string|null} The channel ID or null if not found
 */
function findSlackChannelForSeller(sellerName) {
  try {
    const slackChannelsData = getSheetData(CONFIG.SLACK_CHANNELS_SHEET_NAME);

    if (!slackChannelsData || slackChannelsData.length <= CONFIG.HEADER_ROW_COUNT) {
      return null;
    }

    // Look for an exact match first
    for (let i = CONFIG.HEADER_ROW_COUNT; i < slackChannelsData.length; i++) {
      const row = slackChannelsData[i];
      if (row[SLACK_COLUMNS.RECIPIENT_NAME] === sellerName && row[SLACK_COLUMNS.CHANNEL_ID]) {
        return row[SLACK_COLUMNS.CHANNEL_ID];
      }
    }

    // If no exact match, try a case-insensitive match
    for (let i = CONFIG.HEADER_ROW_COUNT; i < slackChannelsData.length; i++) {
      const row = slackChannelsData[i];
      if (row[SLACK_COLUMNS.RECIPIENT_NAME]?.toLowerCase() === sellerName?.toLowerCase() && 
          row[SLACK_COLUMNS.CHANNEL_ID]) {
        return row[SLACK_COLUMNS.CHANNEL_ID];
      }
    }

    return null;
  } catch (error) {
    Logger.log(`Error finding Slack channel for ${sellerName}: ${error.message}`);
    return null;
  }
}
