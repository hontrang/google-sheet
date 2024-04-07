import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions

namespace SheetHttp {
    export const URL_GRAPHQL_CAFEF = "https://msh-data.cafef.vn/graphql";

    export const OPTIONS_POST: URLFetchRequestOptions = {
        method: "post",
        headers: {
            "Content-Type": "application/json",
            Accept: "application/json",
        }
    };

    export const OPTIONS_GET: URLFetchRequestOptions = {
        method: "get",
        headers: {
            "Content-Type": "application/json",
            Accept: "application/json",
        }
    };

    export function sendRequest(url: string): any {
        try {
            const response = UrlFetchApp.fetch(url);
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
}