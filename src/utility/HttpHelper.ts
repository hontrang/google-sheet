import { IHttp } from '@src/types/generic';

export class HttpHelper implements IHttp {
  sendRequest(url: string, option?: object) {
    try {
      const appliedOption = option ?? HttpHelper.OPTIONS_GET;
      const response = UrlFetchApp.fetch(url, appliedOption);
      return JSON.parse(response.getContentText());
    } catch (e) {
      console.log(`error: ${e}`);
      return null;
    }
  }

  sendGetRequest(url: string) {
    try {
      const response = UrlFetchApp.fetch(url);
      return JSON.parse(response.getContentText());
    } catch (e) {
      console.log(`error: ${e}`);
      return null;
    }
  }

  sendPostRequest(url: string, options?: object) {
    try {
      const effectiveOptions = options ?? HttpHelper.OPTIONS_POST;
      const response = UrlFetchApp.fetch(url, effectiveOptions);
      return JSON.parse(response.getContentText());
    } catch (e) {
      console.log(`error: ${e}`);
      return null;
    }
  }

  public static readonly OPTIONS_POST = {
    method: 'post',
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' }
  };

  public static readonly OPTIONS_GET = {
    method: 'get',
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' }
  };


}
