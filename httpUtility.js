const SheetHttp = {
  URL_GRAPHQL_CAFEF: "https://msh-data.cafef.vn/graphql",
  OPTIONS_POST: {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
    }
  },
  OPTIONS_GET: {
    method: "GET",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
    }
  },
  sendRequest: (url, options) => {
    try {
      const response = UrlFetchApp.fetch(url, options);
      return JSON.parse(response.getContentText());
    } catch (e) {
      SheetLog.logDebug(`error: ${e}`);
      return null;
    }
  },
  sendPostRequest: (url) => {
    try {
      const response = UrlFetchApp.fetch(url, SheetHttp.OPTIONS_POST);
      return JSON.parse(response.getContentText());
    } catch (e) {
      SheetLog.logDebug(`error: ${e}`);
      return null;
    }
  },
  sendGetRequest: (url) => {
    try {
      const response = UrlFetchApp.fetch(url, SheetHttp.OPTIONS_GET);
      return JSON.parse(response.getContentText());
    } catch (e) {
      SheetLog.logDebug(`error: ${e}`);
      return null;
    }
  },
  sendGraphQLRequest: (url, query, variables) => {
    const payload = JSON.stringify({
      query: query,
      variables: variables
    });
    const options = {
      headers: { "Content-Type": "application/json; charset=utf-8" },
      method: "POST",
      payload: payload
    };
    return SheetHttp.sendRequest(url, options);
  }
}