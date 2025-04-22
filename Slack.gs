/**
 * Combined Email and Slack Out-of-Stock Notification System - slack.gs
 * Slack integration functions
 * @version 2.0
 * @lastModified 2025-04-17
 */

/**
 * Securely stores your Slack Bot token
 * RUN THIS FUNCTION ONCE, then comment it out or delete it
 */
function setSlackToken() {
  var token = " insert the token here "; // Replace with your actual Bot User OAuth Token (starts with xoxb-)
  PropertiesService.getScriptProperties().setProperty("SLACK_API_TOKEN", token);
  Logger.log("Token saved successfully");
}

/**
 * Check if Slack token is set properly
 */
function checkSlackToken() {
  const token = PropertiesService.getScriptProperties().getProperty("SLACK_API_TOKEN");
  if (!token) {
    Logger.log("No Slack API token found. Please run setSlackToken() first.");
    return false;
  }
  
  Logger.log("Slack token found. First 5 characters: " + token.substring(0, 5) + "...");
  return true;
}

/**
 * Creates a sheet with all Slack channel IDs for reference
 * Helpful for setting up your channel mapping
 */
function getSlackChannelIds() {
  var slackApiToken = PropertiesService.getScriptProperties().getProperty("SLACK_API_TOKEN");
  
  if (!slackApiToken) {
    Logger.log("No Slack API token found. Please run setSlackToken() first.");
    return;
  }
  
  var options = {
    "method": "get",
    "headers": {
      "Authorization": "Bearer " + slackApiToken
    },
    "muteHttpExceptions": true
  };
  
  var response = UrlFetchApp.fetch("https://slack.com/api/conversations.list?types=public_channel,private_channel", options);
  var responseData = JSON.parse(response.getContentText());
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Channel_IDs") || 
              SpreadsheetApp.getActiveSpreadsheet().insertSheet("Channel_IDs");
  
  // Clear the sheet
  sheet.clear();
  
  // Add headers
  sheet.appendRow(["Channel Name", "Channel ID"]);
  
  // Add channel data
  if (responseData.ok) {
    responseData.channels.forEach(function(channel) {
      sheet.appendRow([channel.name, channel.id]);
    });
    Logger.log("Channel IDs have been added to the Channel_IDs sheet");
  } else {
    Logger.log("Error fetching channels: " + responseData.error);
  }
}

/**
 * Sends a notification to Slack about an OOS email that was sent
 * NOTE: This function is currently disabled in main.gs to avoid extra notifications
 * @param {Object} seller The seller information
 * @param {Number} productCount The number of products in the email
 * @param {Object} subcategories The products grouped by subcategory
 * @return {Boolean} Whether the Slack notification was sent successfully
 */
function sendSlackNotification(seller, productCount, subcategories) {
  try {
    // Get Slack API token
    const slackApiToken = PropertiesService.getScriptProperties().getProperty("SLACK_API_TOKEN");

    if (!CONFIG.SLACK.ENABLED || !slackApiToken) {
      Logger.log('Slack notifications are disabled or API token is not set');
      return false;
    }

    // Count products by subcategory
    const subcategoryCounts = {};
    for (const category in subcategories) {
      subcategoryCounts[category] = subcategories[category].length;
    }

    // Format category counts as text
    let categoryText = "";
    for (const category in subcategoryCounts) {
      categoryText += `• ${category}: ${subcategoryCounts[category]} 商品\n`;
    }
    
    if (!categoryText) {
      categoryText = "分類された商品はありません";
    }

    // Look up the Slack channel ID from the Slack_Channels sheet
    let channelId = findSlackChannelForSeller(seller.name || seller.fullName);
    if (!channelId) {
      // If no specific channel found, try to get the first available channel
      const slackChannelsData = getSheetData(CONFIG.SLACK_CHANNELS_SHEET_NAME);
      if (slackChannelsData && slackChannelsData.length > CONFIG.HEADER_ROW_COUNT) {
        channelId = slackChannelsData[CONFIG.HEADER_ROW_COUNT][SLACK_COLUMNS.CHANNEL_ID];
        Logger.log(`No specific channel found for ${seller.name || seller.fullName}, using first available channel: ${channelId}`);
      }
      
      // If still no channel found, use general channel
      if (!channelId) {
        channelId = "general";
        Logger.log(`No channels found in sheet, using default channel: ${channelId}`);
      }
    }

    // Create Slack message payload - Updated for Japanese
    const payload = {
      "channel": channelId,
      "blocks": [
        {
          "type": "header",
          "text": {
            "type": "plain_text",
            "text": "📧 【在庫切れ商品のご案内】",
            "emoji": true
          }
        },
        {
          "type": "section",
          "text": {
            "type": "mrkdwn",
            "text": `*${seller.name || seller.fullName}* 様宛に*${productCount}点*の在庫切れ商品に関するメールを送信いたしました。詳細は送信メールをご確認ください。`
          }
        },
        {
          "type": "section",
          "text": {
            "type": "mrkdwn",
            "text": `*商品カテゴリ:*\n${categoryText}`
          }
        },
        {
          "type": "section",
          "text": {
            "type": "mrkdwn",
            "text": "ご質問がございましたら、ご連絡ください。"
          }
        }
      ],
      "text": `${seller.name || seller.fullName}様宛在庫切れ商品のお知らせ - ${productCount}商品` // Fallback text
    };

    // Send the message to Slack
    const options = {
      "method": "post",
      "contentType": "application/json",
      "headers": {
        "Authorization": "Bearer " + slackApiToken
      },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    const response = UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", options);
    const responseData = JSON.parse(response.getContentText());

    if (responseData.ok) {
      Logger.log(`Slack notification sent successfully for ${seller.name || seller.fullName}`);
      return true;
    } else {
      Logger.log(`Error from Slack API: ${responseData.error}`);
      return false;
    }

  } catch (error) {
    Logger.log(`Error sending Slack notification: ${error.message}`);
    return false;
  }
}

/**
 * Sends a summary notification to Slack after all emails have been sent
 * NOTE: This function is currently disabled in main.gs to avoid extra notifications
 * @param {Object} stats Statistics about the email sending process
 * @return {Boolean} Whether the summary notification was sent successfully
 */
function sendSlackSummary(stats) {
  try {
    // Get Slack API token
    const slackApiToken = PropertiesService.getScriptProperties().getProperty("SLACK_API_TOKEN");

    if (!CONFIG.SLACK.ENABLED || !slackApiToken) {
      Logger.log('Slack summary notifications are disabled or API token is not set');
      return false;
    }

    // Find a suitable channel for the summary
    let summaryChannel;
    const slackChannelsData = getSheetData(CONFIG.SLACK_CHANNELS_SHEET_NAME);
    if (slackChannelsData && slackChannelsData.length > CONFIG.HEADER_ROW_COUNT) {
      // Use the first available channel
      summaryChannel = slackChannelsData[CONFIG.HEADER_ROW_COUNT][SLACK_COLUMNS.CHANNEL_ID];
    } else {
      // Fallback to general
      summaryChannel = "general";
    }
    Logger.log(`Using channel ${summaryChannel} for summary notifications`);
    
    // Create Slack message payload
    const payload = {
      "channel": summaryChannel, // Use the determined channel
      "blocks": [
        {
          "type": "header",
          "text": {
            "type": "plain_text",
            "text": "📊 在庫切れ商品のお知らせ（サマリー）",
            "emoji": true
          }
        },
        {
          "type": "section",
          "fields": [
            {
              "type": "mrkdwn",
              "text": `*処理した商品数:*\n${stats.processed}`
            },
            {
              "type": "mrkdwn",
              "text": `*送信したメール数:*\n${stats.emailsSent}`
            },
            {
              "type": "mrkdwn",
              "text": `*エラー数:*\n${stats.errors}`
            }
          ]
        },
        {
          "type": "context",
          "elements": [
            {
              "type": "mrkdwn",
              "text": `プロセス完了時間: ${Utilities.formatDate(new Date(), CONFIG.DEFAULT_TIMEZONE, "yyyy/MM/dd HH:mm:ss")}`
            }
          ]
        }
      ],
      "text": `在庫切れ商品のお知らせ: ${stats.emailsSent} 通のメールが送信されました。エラー: ${stats.errors} 件` // Fallback text
    };

    // Send the message to Slack
    const options = {
      "method": "post",
      "contentType": "application/json",
      "headers": {
        "Authorization": "Bearer " + slackApiToken
      },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    const response = UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", options);
    const responseData = JSON.parse(response.getContentText());

    if (responseData.ok) {
      Logger.log(`Slack summary notification sent successfully`);
      return true;
    } else {
      Logger.log(`Error sending Slack summary: ${responseData.error}`);
      return false;
    }

  } catch (error) {
    Logger.log(`Error sending Slack summary: ${error.message}`);
    return false;
  }
}

/**
 * Send OOS notifications to Slack channels
 * This is the main Slack function that is still active
 * @param {Array} oosData 2D array of OOS data
 * @param {Array} channels Array of channel objects
 * @param {Object} stats Statistics object
 */
function sendOOSNotificationsToSlack(oosData, channels, stats) {
  try {
    // Get Slack API token
    const slackApiToken = PropertiesService.getScriptProperties().getProperty("SLACK_API_TOKEN");

    if (!slackApiToken) {
      Logger.log("Error: No Slack API token found. Please run the setSlackToken function first.");
      return;
    }

    // Calculate total count of OOS items (excluding header row)
    const itemCount = oosData.length - CONFIG.HEADER_ROW_COUNT;
    
    // Count products per category
    const categoryCounts = {};
    for (let i = CONFIG.HEADER_ROW_COUNT; i < oosData.length; i++) {
      const category = oosData[i][OOS_COLUMNS.CATEGORY_SUB_CLUSTER] || "未分類";
      
      if (!categoryCounts[category]) {
        categoryCounts[category] = 0;
      }
      
      categoryCounts[category]++;
    }
    
    // Format category counts as text
    let categoryText = "";
    for (const category in categoryCounts) {
      categoryText += `• ${category}: ${categoryCounts[category]} 商品\n`;
    }
    
    if (!categoryText) {
      categoryText = "分類された商品はありません";
    }

    // Log what we're about to do
    Logger.log(`Sending OOS notifications to ${channels.length} Slack channels for ${itemCount} products`);
    
    // Send notification to each channel
    for (const channel of channels) {
      // Format the Slack message for out-of-stock products
      const blocks = [
        {
          "type": "header",
          "text": {
            "type": "plain_text",
            "text": "📢 【在庫切れ商品のご案内】",
            "emoji": true
          }
        },
        {
          "type": "section",
          "text": {
            "type": "mrkdwn",
            "text": `お世話になっております。カートが空いている商品についてご共有させていただきます。\n\n現在、*${itemCount}点*の在庫切れ商品があります。商品名、グレードなどの詳細は別途送信されたメールに記載がございますので、ぜひご確認ください。`
          }
        },
        {
          "type": "section",
          "text": {
            "type": "mrkdwn",
            "text": `*商品カテゴリ:*\n${categoryText}`
          }
        },
        {
          "type": "section",
          "text": {
            "type": "mrkdwn",
            "text": "よろしくお願いいたします。"
          }
        }
      ];

      // Create payload
      const payload = {
        "channel": channel.id,
        "blocks": blocks,
        "text": `在庫切れ商品のお知らせ - ${itemCount}商品` // Fallback text
      };

      const options = {
        "method": "post",
        "contentType": "application/json",
        "headers": {
          "Authorization": "Bearer " + slackApiToken
        },
        "payload": JSON.stringify(payload),
        "muteHttpExceptions": true
      };

      try {
        const response = UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", options);
        const responseData = JSON.parse(response.getContentText());

        if (responseData.ok) {
          Logger.log(`Slack notification sent to ${channel.name} (${channel.id})`);
        } else {
          Logger.log(`Error sending message to channel ${channel.id}: ${responseData.error}`);
        }
      } catch (error) {
        Logger.log(`Exception when sending to channel ${channel.id}: ${error.toString()}`);
      }
      
      // Add a small delay between messages to avoid rate limiting
      Utilities.sleep(CONFIG.SLACK.NOTIFICATION_DELAY_MS);
    }

    // Log completion
    Logger.log(`OOS notifications sent to all ${channels.length} channels`);

  } catch (error) {
    Logger.log(`Error sending OOS notifications to Slack: ${error.message}`);
  }
}
