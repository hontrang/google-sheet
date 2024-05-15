import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;

class SheetHttp {
    static URL_GRAPHQL_CAFEF = "https://msh-data.cafef.vn/graphql";
    static TOKEN: string | undefined;

    static OPTIONS_POST: URLFetchRequestOptions = {
        method: "post",
        headers: {
            "Content-Type": "application/json",
            "Accept": "application/json",
        }
    };

    static OPTIONS_GET: URLFetchRequestOptions = {
        method: "get",
        headers: {
            "Content-Type": "application/json",
            "Accept": "application/json",
        }
    };

    static sendRequest(url: string, option?: URLFetchRequestOptions): any {
        try {
            const appliedOption = option || SheetHttp.OPTIONS_GET;
            const response = UrlFetchApp.fetch(url, appliedOption);
            return JSON.parse(response.getContentText());
        } catch (e) {
            SheetLog.logDebug(`error: ${e}`);
            return null;
        }
    }

    static sendPostRequest(url: string, options?: URLFetchRequestOptions): any {
        try {
            const effectiveOptions = options || SheetHttp.OPTIONS_POST;
            const response = UrlFetchApp.fetch(url, effectiveOptions);
            return JSON.parse(response.getContentText());
        } catch (e) {
            SheetLog.logDebug(`error: ${e}`);
            return null;
        }
    }

    static sendGetRequest(url: string): any {
        try {
            const response = UrlFetchApp.fetch(url, SheetHttp.OPTIONS_GET);
            return JSON.parse(response.getContentText());
        } catch (e) {
            SheetLog.logDebug(`error: ${e}`);
            return null;
        }
    }

    static sendGraphQLRequest(url: string, query: string, variables?: any): any {
        const PAYLOAD = JSON.stringify({
            query: query,
            variables: variables
        });
        const OPTIONS: URLFetchRequestOptions = {
            method: "post",
            payload: PAYLOAD
        };
        return SheetHttp.sendPostRequest(url, OPTIONS);
    }

    static getToken(): string {
        if (SheetHttp.TOKEN !== undefined) return SheetHttp.TOKEN;
        else {
            const consumerID = SheetUtil.layDuLieuTrongO(SheetUtil.SHEET_CAU_HINH, 'B7');
            const consumerSecret = SheetUtil.layDuLieuTrongO(SheetUtil.SHEET_CAU_HINH, 'B8');
            const OPTIONS_POST_TOKEN_SSI: URLFetchRequestOptions = {
                method: "post",
                headers: {
                    "Content-Type": "application/json",
                    "Accept": "application/json",
                },
                payload: JSON.stringify({
                    "consumerID": consumerID,
                    "consumerSecret": consumerSecret
                })
            };
            const URL = `https://fc-data.ssi.com.vn/api/v2/Market/AccessToken`;
            const response = SheetHttp.sendPostRequest(URL, OPTIONS_POST_TOKEN_SSI);
            SheetHttp.TOKEN = "Bearer " + response.data.accessToken;
            return SheetHttp.TOKEN;
        }
    }
}