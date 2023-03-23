
var SheetHttp = {
    sendRequest: function (url, options) {
      try{
        var response = UrlFetchApp.fetch(url, options);
        return JSON.parse(response.getContentText());
      }catch(e){
        SheetLog.logDebug("error: "+ e);
        return null;
      }
    }
  }