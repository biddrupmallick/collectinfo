function extractInformation() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  // Iterate through each row
  for (var i = 0; i < values.length; i++) {
    var url = values[i][0]; // Assuming the URLs are in column A

    if (url !== "") {
      try {
        var htmlContent = fetchUrlContent(url);

        // Extract email addresses and social links
        var email = extractEmail(htmlContent);
        var facebook = extractSocialLink(htmlContent, 'facebook');
        var linkedin = extractSocialLink(htmlContent, 'linkedin');
        var instagram = extractSocialLink(htmlContent, 'instagram');
        var twitter = extractSocialLink(htmlContent, 'twitter');
        var youtube = extractSocialLink(htmlContent, 'youtube');
        var pinterest = extractSocialLink(htmlContent, 'pinterest');
        var bsky = extractSocialLink(htmlContent, 'bsky');
        var tiktok = extractSocialLink(htmlContent, 'tiktok');

        // Write the extracted information to the corresponding columns
        sheet.getRange(i + 1, 2).setValue(email);       // Column B for email
        sheet.getRange(i + 1, 3).setValue(facebook);    // Column C for Facebook
        sheet.getRange(i + 1, 4).setValue(linkedin);    // Column D for LinkedIn
        sheet.getRange(i + 1, 5).setValue(instagram);   // Column E for Instagram
        sheet.getRange(i + 1, 6).setValue(twitter);     // Column F for Twitter
        sheet.getRange(i + 1, 7).setValue(youtube);     // Column G for YouTube
        sheet.getRange(i + 1, 8).setValue(pinterest);   // Column H for Pinterest
        sheet.getRange(i + 1, 9).setValue(bsky);        // Column I for Bluesky
        sheet.getRange(i + 1, 10).setValue(tiktok);     // Column J for TikTok

        // Pause to avoid rate limiting
        Utilities.sleep(1000);
      } catch (error) {
        // Log the error and continue to the next iteration
        sheet.getRange(i + 1, 11).setValue(`Error: ${error}`); // Column K for errors
      }
    }
  }
}

function fetchUrlContent(url) {
  try {
    var response = UrlFetchApp.fetch(url);
    return response.getContentText();
  } catch (error) {
    throw `Error fetching content from ${url}: ${error}`;
  }
}

function extractEmail(content) {
  var emailRegex = /mailto:([^\n\s"']+)|[\w._%+-]+@[\w.-]+\.[a-zA-Z]{2,}/g;
  var matches = content.match(emailRegex);
  if (matches && matches.length > 0) {
    return matches[0].startsWith("mailto:") ? matches[0].substring(7) : matches[0];
  } else {
    return "";
  }
}

function extractSocialLink(content, platform) {
  var domainMap = {
    facebook: "facebook.com",
    linkedin: "linkedin.com",
    instagram: "instagram.com",
    twitter: "twitter.com",
    youtube: "youtube.com",
    pinterest: "pinterest.com",
    bsky: "bsky.app",
    tiktok: "tiktok.com"
  };

  var domain = domainMap[platform];
  var regex = new RegExp(`https?:\\/\\/(?:www\\.)?${domain}\\/[^\\s"']+`, 'i');
  var match = content.match(regex);
  return match ? match[0] : "";
}
