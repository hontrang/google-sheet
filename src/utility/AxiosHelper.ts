/* eslint-disable @typescript-eslint/no-extraneous-class */
/* eslint-disable @typescript-eslint/no-explicit-any */
import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;
import { Http } from "../types/types";
import { HttpHelper } from "./HttpHelper";
import axios from 'axios';
import { SheetHelper } from "./SheetHelper";

export class AxiosHelper implements Http {
  async sendRequest(url: string, option?: URLFetchRequestOptions) {
    try {
      const appliedOption = option || HttpHelper.OPTIONS_GET;
      const response = await axios.get(url, appliedOption);
      return JSON.parse(response.data);
    } catch (e) {
      console.log(`error: ${e}`);
      return null;
    }
  }
  async sendPostRequest(url: string, options?: URLFetchRequestOptions) {
    try {
      const response = await axios.post(url, options);
      return response;
    } catch (e) {
      console.log(`error: ${e}`);
      return null;
    }
  }
  async sendGetRequest(url: string) {
    try {
      const response = await axios.get(url);
      return response.data;
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
  async getToken(): Promise<string> {
    if (HttpHelper.token !== undefined) return HttpHelper.token;
    else {
      const consumerID = "35e723ae7e4c426694226ed4649379d5";
      const consumerSecret = "a54573a4a78949a0b59befc59523aee2";
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
      const response = await axios.post(URL, { consumerID: consumerID, consumerSecret: consumerSecret });
      HttpHelper.token = 'Bearer ' + response.data.accessToken;
      return HttpHelper.token;
    }
  }
  private static token: string | undefined;

  // eslint-disable-next-line @typescript-eslint/naming-convention
  private static readonly OPTIONS_POST: URLFetchRequestOptions = {
    method: 'post',
    // eslint-disable-next-line @typescript-eslint/naming-convention
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' }
  };

  // eslint-disable-next-line @typescript-eslint/naming-convention
  private static readonly OPTIONS_GET: URLFetchRequestOptions = {
    method: 'get',
    // eslint-disable-next-line @typescript-eslint/naming-convention
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' }
  };
}
