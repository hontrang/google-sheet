var SheetHttp = {
  URL_GRAPHQL_CAFEF: "https://msh-data.cafef.vn/graphql",
  OPTIONS_POST: {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
    }
  },
  OPTIONS_GET: OPTIONS = {
    method: "GET",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
    }
  },
  sendRequest: function (url, options) {
    try {
      var response = UrlFetchApp.fetch(url, options);
      return JSON.parse(response.getContentText());
    } catch (e) {
      SheetLog.logDebug("error: " + e);
      return null;
    }
  },
  sendPostRequest: function (url) {
    try {
      var response = UrlFetchApp.fetch(url, this.OPTIONS_POST);
      return JSON.parse(response.getContentText());
    } catch (e) {
      SheetLog.logDebug("error: " + e);
      return null;
    }
  },
  sendGetRequest: function (url) {
    try {
      var response = UrlFetchApp.fetch(url, this.OPTIONS_GET);
      return JSON.parse(response.getContentText());
    } catch (e) {
      SheetLog.logDebug("error: " + e);
      return null;
    }
  },
  sendGraphQLRequest: function (url, query, variables) {
    const payload = JSON.stringify({
      query: query,
      variables: variables
    });
    const OPTIONS = {
      "headers": { "Content-Type": "application/json; charset=utf-8" },
      "payload": payload,
      "method": "POST"
    };

    return this.sendRequest(url, OPTIONS);
  }
}