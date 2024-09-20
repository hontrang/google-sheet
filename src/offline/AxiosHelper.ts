/* eslint-disable @typescript-eslint/no-extraneous-class */
/* eslint-disable @typescript-eslint/no-explicit-any */
import { Http, HttpResponse } from '@src/types/types';
import axios from 'axios';

export class AxiosHelper implements Http {
  async sendRequest(url: string, option?: any) {
    const response = await axios.get(url, option);
    return response as HttpResponse;
  }
  async sendPostRequest(url: string, options?: any) {
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
    return response as HttpResponse;
  }
  async getToken(): Promise<string> {
    if (this.token !== undefined) return this.token;
    else {
      const consumerID = 'replaced';
      const consumerSecret = 'replaced';
      const URL = `https://fc-data.ssi.com.vn/api/v2/Market/AccessToken`;
      const response = await axios.post(URL, { consumerID: consumerID, consumerSecret: consumerSecret });
      this.token = 'Bearer ' + response.data.data.accessToken;
      return this.token;
    }
  }
  private token: string | undefined;
}
