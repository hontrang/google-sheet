/* eslint-disable @typescript-eslint/no-extraneous-class */
/* eslint-disable @typescript-eslint/no-explicit-any */
import { Configuration } from "../configuration/Configuration";
import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;
import { Http, HttpResponse } from "@src/types/types";
import { HttpHelper } from "@utils/HttpHelper";
import axios from 'axios';

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
  async sendGetRequest(url: string): Promise<HttpResponse> {
    const response = await axios.get(url);
    return (response as HttpResponse);
  }
  async getToken(): Promise<string> {
    if (this.token !== undefined) return this.token;
    else {
      const consumerID = Configuration.consumerID;
      const consumerSecret = Configuration.consumerSecret;
      const URL = `https://fc-data.ssi.com.vn/api/v2/Market/AccessToken`;
      const response = await axios.post(URL, { consumerID: consumerID, consumerSecret: consumerSecret });
      this.token = 'Bearer ' + response.data.data.accessToken;
      return this.token;
    }
  }
  private token: string | undefined;

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
