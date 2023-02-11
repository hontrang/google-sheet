
var SheetHttp = {
    sendRequest: function (url, options) {
      var response = UrlFetchApp.fetch(url, options);
      return JSON.parse(response.getContentText());
    }
  }