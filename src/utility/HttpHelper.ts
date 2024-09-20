/* eslint-disable @typescript-eslint/no-extraneous-class */
import { SheetHelper } from './SheetHelper';
import { Http } from '@src/types/types';

export class HttpHelper implements Http {
  sendRequest(url: string, option?: any) {
    try {
      const appliedOption = option || HttpHelper.OPTIONS_GET;
      const response = UrlFetchApp.fetch(url, appliedOption);
      return JSON.parse(response.getContentText());
    } catch (e) {
      console.log(`error: ${e}`);
      return null;
    }
  }
  sendPostRequest(url: string, options?: any) {
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
      const response = UrlFetchApp.fetch(url);
      return JSON.parse(response.getContentText());
    } catch (e) {
      console.log(`error: ${e}`);
      return null;
    }
  }
  async getToken(): Promise<string> {
    const sheetHelper = new SheetHelper();
    if (this.token !== undefined) return this.token;
    else {
      const consumerID = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B7');
      const consumerSecret = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B8');
      const OPTIONS_POST_TOKEN_SSI = {
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
      this.token = 'Bearer ' + response.data.accessToken;
      return this.token;
    }
  }
  public token: string | undefined;

  // eslint-disable-next-line @typescript-eslint/naming-convention
  public static readonly OPTIONS_POST = {
    method: 'post',
    // eslint-disable-next-line @typescript-eslint/naming-convention
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' }
  };

  // eslint-disable-next-line @typescript-eslint/naming-convention
  public static readonly OPTIONS_GET = {
    method: 'get',
    // eslint-disable-next-line @typescript-eslint/naming-convention
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' }
  };
}
