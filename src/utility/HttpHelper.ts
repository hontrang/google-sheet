/* eslint-disable @typescript-eslint/no-explicit-any */
import { SheetHelper } from "./SheetHelper";
import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;

export class HttpHelper {
  private static URL_GRAPHQL_CAFEF = 'https://msh-data.cafef.vn/graphql';
  private static TOKEN: string | undefined;

  private static OPTIONS_POST: URLFetchRequestOptions = {
    method: 'post',
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' }
  };

  private static OPTIONS_GET = {
    method: 'get',
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' }
  };

  public static sendRequest(url: string, option?: URLFetchRequestOptions): any {
    try {
      const appliedOption = option || this.OPTIONS_GET;
      const response = UrlFetchApp.fetch(url, appliedOption);
      return JSON.parse(response.getContentText());
    } catch (e) {
      console.log(`error: ${e}`);
      return null;
    }
  }

  public static sendPostRequest(url: string, options?: URLFetchRequestOptions): any {
    try {
      const effectiveOptions = options || this.OPTIONS_POST;
      const response = UrlFetchApp.fetch(url, effectiveOptions);
      return JSON.parse(response.getContentText());
    } catch (e) {
      console.log(`error: ${e}`);
      return null;
    }
  }

  public static sendGetRequest(url: string): any {
    try {
      const response = UrlFetchApp.fetch(url, this.OPTIONS_GET);
      return JSON.parse(response.getContentText());
    } catch (e) {
      console.log(`error: ${e}`);
      return null;
    }
  }

  public static sendGraphQLRequest(url: string, query: string, variables?: any): any {
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

  public static getToken(): string {
    if (this.TOKEN !== undefined) return this.TOKEN;
    else {
      const consumerID = SheetHelper.layDuLieuTrongO(SheetHelper.SheetName.SHEET_CAU_HINH, 'B7');
      const consumerSecret = SheetHelper.layDuLieuTrongO(SheetHelper.SheetName.SHEET_CAU_HINH, 'B8');
      const OPTIONS_POST_TOKEN_SSI: URLFetchRequestOptions = {
        method: 'post',
        headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
        payload: JSON.stringify({
          consumerID: consumerID,
          consumerSecret: consumerSecret
        })
      };
      const URL = `https://fc-data.ssi.com.vn/api/v2/Market/AccessToken`;
      const response = this.sendPostRequest(URL, OPTIONS_POST_TOKEN_SSI);
      this.TOKEN = 'Bearer ' + response.data.accessToken;
      return this.TOKEN;
    }
  }
}
