/**
 * Combined Email and Slack Out-of-Stock Notification System - webhook.gs
 * Handles web app requests for email tracking
 * @version 2.0
 * @lastModified 2025-04-17
 */

/**
 * Set this script as a web app to handle tracking pixel requests
 * Run this script from the Script Editor and deploy as a web app with these settings:
 * - Execute the app as: Me (your account)
 * - Who has access to the app: Anyone, even anonymous
 */

/**
 * Handles GET requests from tracking pixels
 * This is the main entry point for the web app
 * @param {Object} e The event object containing request parameters
 * @return {HtmlOutput} A transparent response for the tracking pixel
 */
function doGet(e) {
  try {
    // Parse the request parameters
    const params = e.parameter;
    
    // Get the tracking ID and action
    const trackingId = params.id;
    const action = params.action || 'view';
    
    // Log the tracking request
    Logger.log(`Tracking request received: ID=${trackingId}, Action=${action}`);
    
    if (trackingId && action === 'open') {
      // Record the email open
      recordEmailOpen(trackingId);
    }
    
    // Return a transparent 1x1 pixel gif
    return createTransparentPixel();
    
  } catch (error) {
    Logger.log(`Error handling tracking request: ${error.message}`);
    return createTransparentPixel();
  }
}

/**
 * Creates a transparent 1x1 pixel image response
 * @return {HtmlOutput} An HTML response with a transparent pixel
 */
function createTransparentPixel() {
  // Base64 encoded 1x1 transparent GIF
  const transparentGif = 'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7';
  
  // Convert the base64 string to binary
  const decoded = Utilities.base64Decode(transparentGif);
  
  // Create a blob with the GIF data
  const blob = Utilities.newBlob(decoded, 'image/gif', 'transparent.gif');
  
  // Return the image response
  return HtmlService.createHtmlOutput('<img src="data:image/gif;base64,' + transparentGif + '" width="1" height="1" alt="">')
    .setContentType('image/gif');
}

/**
 * Record that an email was opened
 * @param {string} trackingId The tracking ID from the pixel request
 */
function recordEmailOpen(trackingId) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const trackingSheet = spreadsheet.getSheetByName(CONFIG.TRACKING_SHEET_NAME);

    if (!trackingSheet) {
      Logger.log('Tracking sheet not found');
      return;
    }

    // Get all tracking data
    const data = trackingSheet.getDataRange().getValues();

    // Skip the header row and find the tracking ID
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === trackingId) {
        // Only update if this is the first time the email was opened
        if (data[i][5] === 'No') {
          const now = new Date();
          const formattedDate = Utilities.formatDate(now, 
                                                CONFIG.DEFAULT_TIMEZONE, 
                                                "yyyy/MM/dd HH:mm:ss");
          
          // Update the row to indicate the email was opened
          trackingSheet.getRange(i + 1, 5).setValue(formattedDate); // Open Date with time in Tokyo format
          trackingSheet.getRange(i + 1, 6).setValue('Yes');         // Opened
          Logger.log(`Email with tracking ID ${trackingId} marked as opened at ${formattedDate}`);
        } else {
          // Increment view count
          let currentViews = data[i][6] || 0;
          trackingSheet.getRange(i + 1, 7).setValue(currentViews + 1);
          Logger.log(`Email with tracking ID ${trackingId} view count incremented to ${currentViews + 1}`);
        }
        break;
      }
    }
  } catch (error) {
    Logger.log(`Error recording email open: ${error.message}`);
  }
}
