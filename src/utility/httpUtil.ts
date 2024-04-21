import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions
namespace SheetHttp {
    export const URL_GRAPHQL_CAFEF = "https://msh-data.cafef.vn/graphql";
    export let TOKEN: string | undefined;
    export const CONSUMERID = '35e723ae7e4c426694226ed4649379d5';
    export const CONSUMER_SECRET = 'a54573a4a78949a0b59befc59523aee2';

    export const OPTIONS_POST: URLFetchRequestOptions = {
        method: "post",
        headers: {
            "Content-Type": "application/json",
            "Accept": "application/json",
        }
    };

    export const OPTIONS_GET: URLFetchRequestOptions = {
        method: "get",
        headers: {
            "Content-Type": "application/json",
            "Accept": "application/json",
        }
    };


    export const OPTIONS_POST_TOKEN_SSI: URLFetchRequestOptions = {
        method: "post",
        headers: {
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        payload: JSON.stringify({
            "consumerID": SheetHttp.CONSUMERID,
            "consumerSecret": SheetHttp.CONSUMER_SECRET
        })
    };

    export function sendRequest(url: string, option?: URLFetchRequestOptions): any {
        try {
            const appliedOption = option || OPTIONS_GET;
            const response = UrlFetchApp.fetch(url, appliedOption);
            return JSON.parse(response.getContentText());
        } catch (e) {
            SheetLog.logDebug(`error: ${e}`);
            return null;
        }
    }


    export function sendPostRequest(url: string, options?: URLFetchRequestOptions): any {
        try {
            const effectiveOptions = options || OPTIONS_POST;
            const response = UrlFetchApp.fetch(url, effectiveOptions);
            return JSON.parse(response.getContentText());
        } catch (e) {
            SheetLog.logDebug(`error: ${e}`);
            return null;
        }
    }

    export function sendGetRequest(url: string): any {
        try {
            const response = UrlFetchApp.fetch(url, OPTIONS_GET);
            return JSON.parse(response.getContentText());
        } catch (e) {
            SheetLog.logDebug(`error: ${e}`);
            return null;
        }
    }

    export function sendGraphQLRequest(url: string, query: string, variables?: any): any {
        const PAYLOAD = JSON.stringify({
            query: query,
            variables: variables
        });
        const OPTIONS: URLFetchRequestOptions = {
            method: "post",
            payload: PAYLOAD
        }
        return sendPostRequest(url, OPTIONS); // Sửa lại để gọi sendPostRequest với url và options đã chỉnh sửa
    }

    export function getToken(): string {
        if (TOKEN !== undefined) return TOKEN;
        else {
            const URL = `https://fc-data.ssi.com.vn/api/v2/Market/AccessToken`;
            const response = sendPostRequest(URL, OPTIONS_POST_TOKEN_SSI);
            TOKEN = "Bearer " + response.data.accessToken;
            return TOKEN;
        }
    }
}