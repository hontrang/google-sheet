/* eslint-disable @typescript-eslint/no-explicit-any */
import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;
// eslint-disable-next-line @typescript-eslint/no-unused-vars
namespace HttpHelper {
  export const URL_GRAPHQL_CAFEF = 'https://msh-data.cafef.vn/graphql';
  export let TOKEN: string | undefined;

  export const OPTIONS_POST: URLFetchRequestOptions = {
    method: 'post',
    // eslint-disable-next-line @typescript-eslint/naming-convention
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' }
  };

  export const OPTIONS_GET: URLFetchRequestOptions = {
    method: 'get',
    // eslint-disable-next-line @typescript-eslint/naming-convention
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' }
  };

  export function sendRequest(url: string, option?: URLFetchRequestOptions): any {
    try {
      const appliedOption = option || OPTIONS_GET;
      const response = UrlFetchApp.fetch(url, appliedOption);
      return JSON.parse(response.getContentText());
    } catch (e) {
      LogHelper.logDebug(`error: ${e}`);
      return null;
    }
  }

  export function sendPostRequest(url: string, options?: URLFetchRequestOptions): any {
    try {
      const effectiveOptions = options || OPTIONS_POST;
      const response = UrlFetchApp.fetch(url, effectiveOptions);
      return JSON.parse(response.getContentText());
    } catch (e) {
      LogHelper.logDebug(`error: ${e}`);
      return null;
    }
  }

  export function sendGetRequest(url: string): any {
    try {
      const response = UrlFetchApp.fetch(url, OPTIONS_GET);
      return JSON.parse(response.getContentText());
    } catch (e) {
      LogHelper.logDebug(`error: ${e}`);
      return null;
    }
  }

  export function sendGraphQLRequest(url: string, query: string, variables?: any): any {
    const PAYLOAD = JSON.stringify({
      query: query,
      variables: variables
    });
    const OPTIONS: URLFetchRequestOptions = {
      method: 'post',
      payload: PAYLOAD
    };
    return sendPostRequest(url, OPTIONS);
  }

  export function getToken(): string {
    if (TOKEN !== undefined) return TOKEN;
    else {
      const consumerID = SheetHelper.layDuLieuTrongO(SheetHelper.SHEET_CAU_HINH, 'B7');
      const consumerSecret = SheetHelper.layDuLieuTrongO(SheetHelper.SHEET_CAU_HINH, 'B8');
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
      const response = sendPostRequest(URL, OPTIONS_POST_TOKEN_SSI);
      TOKEN = 'Bearer ' + response.data.accessToken;
      return TOKEN;
    }
  }
}