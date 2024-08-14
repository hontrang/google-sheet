/* eslint-disable @typescript-eslint/no-extraneous-class */
/* eslint-disable @typescript-eslint/no-explicit-any */
import { SheetHelper } from "./SheetHelper";
import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;
import { Http } from "../types/types";

export class HttpHelper implements Http {
  sendRequest(url: string, option?: URLFetchRequestOptions) {
    try {
      const appliedOption = option || HttpHelper.OPTIONS_GET;
      const response = UrlFetchApp.fetch(url, appliedOption);
      return JSON.parse(response.getContentText());
    } catch (e) {
      console.log(`error: ${e}`);
      return null;
    }
  }
  sendPostRequest(url: string, options?: URLFetchRequestOptions) {
    try {
      const effectiveOptions = options || HttpHelper.OPTIONS_POST;
      const response = UrlFetchApp.fetch(url, effectiveOptions);
      return JSON.parse(response.getContentText());
    } catch (e) {
      console.log(`error: ${e}`);
      return null;
    }
  }
  sendGetRequest(url: string) {
    try {
      const response = UrlFetchApp.fetch(url, HttpHelper.OPTIONS_GET);
      return JSON.parse(response.getContentText());
    } catch (e) {
      console.log(`error: ${e}`);
      return null;
    }
  }
  sendGraphQLRequest(url: string, query: string, variables?: any) {
    const PAYLOAD = JSON.stringify({
      query: query,
      variables: variables
    });
    const OPTIONS: URLFetchRequestOptions = {
      method: 'post',
      payload: PAYLOAD
    };
    return this.sendPostRequest(url, OPTIONS);
  }
  getToken(): string {
    const sheetHelper = new SheetHelper();
    if (HttpHelper.token !== undefined) return HttpHelper.token;
    else {
      const consumerID = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B7');
      const consumerSecret = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B8');
      const OPTIONS_POST_TOKEN_SSI: URLFetchRequestOptions = {
        method: 'post',
        // eslint-disable-next-line @typescript-eslint/naming-convention
        headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
        payload: JSON.stringify({
          consumerID: consumerID,
          consumerSecret: consumerSecret
        })
      };
      const URL = `https://fc-data.ssi.com.vn/api/v2/Market/AccessToken`;
      const response = this.sendPostRequest(URL, OPTIONS_POST_TOKEN_SSI);
      HttpHelper.token = 'Bearer ' + response.data.accessToken;
      return HttpHelper.token;
    }
  }
  public static token: string | undefined;

  // eslint-disable-next-line @typescript-eslint/naming-convention
  public static readonly OPTIONS_POST: URLFetchRequestOptions = {
    method: 'post',
    // eslint-disable-next-line @typescript-eslint/naming-convention
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' }
  };

  // eslint-disable-next-line @typescript-eslint/naming-convention
  public static readonly OPTIONS_GET: URLFetchRequestOptions = {
    method: 'get',
    // eslint-disable-next-line @typescript-eslint/naming-convention
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' }
  };
}
