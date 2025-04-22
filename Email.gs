/**
 * Combined Email and Slack Out-of-Stock Notification System - email.gs
 * Email generation and sending functionality
 * @version 2.0
 * @lastModified 2025-04-17
 */

/**
 * Sends a notification email using the provided template and data
 * @param {Object} htmlTemplate The HTML template to use
 * @param {Object} seller Seller information
 * @param {Array} products List of OOS products
 * @param {String} businessUpdate Business update HTML content
 * @return {Boolean} Whether the email was sent successfully
 */
function sendNotificationEmail(htmlTemplate, seller, products, businessUpdate) {
  try {
    // Group products by subcategory
    const productsBySubcategory = {};
    products.forEach(product => {
      if (!productsBySubcategory[product.category]) {
        productsBySubcategory[product.category] = [];
      }
      productsBySubcategory[product.category].push(product);
    });
    
    // Sort each subcategory by the subcategory rank
    for (const category in productsBySubcategory) {
      productsBySubcategory[category].sort((a, b) => a.subcategoryRank - b.subcategoryRank);
    }
    
    // Flatten back to a single array but maintain subcategory grouping
    const sortedProducts = [];
    for (const category in productsBySubcategory) {
      sortedProducts.push(...productsBySubcategory[category]);
    }
    
    // Prepare data for the template
    const templateData = {
      sellerName: seller.name || seller.fullName,
      sellerEmail: seller.email,
      productCount: products.length,
      products: sortedProducts, // Use the products sorted by subcategory
      subcategories: productsBySubcategory, // Also provide grouped products if template needs them
      businessUpdate: businessUpdate || '<h3>【直近の販売傾向および人気商品のご案内】</h3><p>いつもお世話になっております。</p>',
      currentDate: new Date(),
      // For compatibility with the template format
      orderCount: products.length,
      allOrderIds: products.map((p, i) => `商品 ${i+1}`).join(', '),
      earliestOrderCreation: new Date(),
      latestOrderCreation: new Date(),
      shippingDeadline: new Date(new Date().getTime() + 24 * 60 * 60 * 1000) // Tomorrow
    };
    
    // Populate the template
    for (const key in templateData) {
      htmlTemplate[key] = templateData[key];
    }
    
    // Add helper function to format date using Tokyo timezone
    htmlTemplate.formatDate = function(date) {
      if (!date) return '';
      try {
        return Utilities.formatDate(new Date(date), CONFIG.DEFAULT_TIMEZONE, "yyyy年MM月dd日");
      } catch (e) {
        return String(date);
      }
    };
    
    // Evaluate the template and get the HTML content
    let htmlForEmail = htmlTemplate.evaluate().getContent();
    
    // Add tracking pixel to the HTML content
    const trackingPixel = generateTrackingPixel(seller.email, seller.id);
    htmlForEmail += trackingPixel;
    
    // Create email subject with seller name
    const emailSubject = CONFIG.EMAIL_SUBJECT.replace('{sellerName}', seller.name || seller.fullName);
    
    // Handle multiple email addresses
    const emailAddresses = seller.email.split(',').map(email => email.trim());
    let successCount = 0;
    
    // Send email to each recipient
    for (const email of emailAddresses) {
      if (isValidEmail(email)) {
        GmailApp.sendEmail(
          email,
          emailSubject,
          `Please view this email in an HTML-capable email client to see the full content.`,
          {
            htmlBody: htmlForEmail,
            name: CONFIG.EMAIL_SENDER_NAME,
            replyTo: CONFIG.EMAIL_SENDER_ADDRESS
          }
        );
        Logger.log(`Email sent successfully to ${email} (${seller.name || seller.fullName})`);
        successCount++;
      } else {
        Logger.log(`Invalid email address skipped: ${email}`);
      }
    }
    
    return successCount > 0;
    
  } catch (error) {
    Logger.log(`Error sending email for ${seller.name || seller.fullName}: ${error.message}`);
    return false;
  }
}

/**
 * Generate a tracking pixel for email open tracking
 * @param {string} sellerEmail The email address of the recipient
 * @param {string} sellerId The ID of the seller
 * @return {string} HTML for a tracking pixel
 */
function generateTrackingPixel(sellerEmail, sellerId) {
  // Create a unique tracking ID
  const trackingId = Utilities.getUuid();
  
  // Store this tracking ID in a separate sheet for later reference
  recordTrackingId(trackingId, sellerEmail, sellerId);
  
  // Build the tracking URL
  const trackingUrl = `${CONFIG.WEB_APP_URL}?id=${trackingId}&action=open`;
  
  // Return HTML for a 1x1 transparent pixel with the tracking URL
  return `<img src="${trackingUrl}" width="1" height="1" alt="" style="display:none">`;
}

/**
 * Record the tracking ID and email details in a separate sheet
 * @param {string} trackingId The unique tracking ID
 * @param {string} email The recipient email
 * @param {string} sellerId The seller ID
 */
function recordTrackingId(trackingId, email, sellerId) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Try to get the tracking sheet, or create it if it doesn't exist
    let trackingSheet = spreadsheet.getSheetByName(CONFIG.TRACKING_SHEET_NAME);
    if (!trackingSheet) {
      trackingSheet = spreadsheet.insertSheet(CONFIG.TRACKING_SHEET_NAME);
      // Add headers
      trackingSheet.appendRow(['Tracking ID', 'Email', 'Seller ID', 'Send Date', 'Open Date', 'Opened', 'Views']);
    }
    
    // Add the new tracking record with Tokyo time for the send date
    const now = new Date();
    const tokyoSendDate = Utilities.formatDate(now, CONFIG.DEFAULT_TIMEZONE, "yyyy/MM/dd HH:mm:ss");
    
    trackingSheet.appendRow([
      trackingId,
      email,
      sellerId,
      tokyoSendDate, // Using Tokyo-formatted date instead of plain Date object
      '',
      'No',
      0 // Initialize Views count to 0
    ]);
    
    Logger.log(`Tracking ID ${trackingId} recorded for ${email}`);
  } catch (error) {
    Logger.log(`Error recording tracking ID: ${error.message}`);
  }
}
