
var SheetHttp = {
  sendRequest: function (url, options) {
    try {
      var response = UrlFetchApp.fetch(url, options);
      return JSON.parse(response.getContentText());
    } catch (e) {
      SheetLog.logDebug("error: " + e);
      return null;
    }
  },
  sendGraphQLRequest: function (url, query, variables) {
    const payload = JSON.stringify({ query, variables });
    const options = {
      headers: {
        "Content-Type": "application/json;charset=utf-8",
      },
      payload,
      method: "POST"
    };

    return this.sendRequest(url, options);
  }
}