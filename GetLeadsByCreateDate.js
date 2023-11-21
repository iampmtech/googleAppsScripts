// Method with main logic
function getLeadsFromKOMMO() {
  IAMPMKOMMOAuthToken.getPermissionToRefreshTokenSS(); // Get Permission to DB of Access Token
  var accessToken = IAMPMKOMMOAuthToken.getKOMMOAuthToken(); // Get New Access Token (making a GET request to KOMMO)

  var page = 1; // Page number (If there is a lot of leads)

  // GET request to retrieve leads
  do{
    var responseData = makeGETRequestToKOMMO(accessToken, date1, date2, page); // Response data from KOMMO that is converted from JSON

    if(responseData == null) break; // Check if there is an response from KOMMO (If we still taking a leads)

    page++; // Increment page variable to get more leads

    putLeadsToSS(responseData._embedded.leads); // Calling a method to put leads to Spreadsheet
  } while(responseData != null);

  // You can do further processing or logging here
}

// Method to make GET request to KOMMO to get leads by create date
function makeGETRequestToKOMMO(accessToken, date1, date2, page)
{
  // Set up the headers with the access token
  var headers = {
    "Authorization": "Bearer " + accessToken,
    "Content-Type": "application/json"
  };

  // Construct the API endpoint URL
  var apiUrl = "https://iampm.kommo.com/api/v4/leads?filter[pipeline_id]=3447550&filter[created_at][from]=" + dateToUnix(date1) + "&filter[created_at][to]=" + dateToUnix(date2) + "&limit=250" + "&page=" + page;

  var response = UrlFetchApp.fetch(apiUrl, { headers: headers });
  if(response.getResponseCode() == 200) {
    return JSON.parse(response.getContentText());
  } else {
    Logger.log("Error: " + response.getContentText());
    return null;
  }
}

// Method to put leads to Spreadsheet
function putLeadsToSS(leads) {
  var nextRow = sheet.getLastRow() + 1; // Writing in a new rows

  for (var i = 0; i < leads.length; i++) {
    var lead = leads[i];
    sheet.getRange(nextRow + i, 1).setValue(lead.id);
    sheet.getRange(nextRow + i, 2).setValue(lead.name);
    var tags = lead._embedded.tags.map(tag => tag.name).join(', ');
    sheet.getRange(nextRow + i, 3).setValue(tags);
    sheet.getRange(nextRow + i, 4).setValue(lead.status_id);
    sheet.getRange(nextRow + i, 5).setValue(unixTimestampToDate(lead.created_at));
    Logger.log(leads[i]);
  }
}
