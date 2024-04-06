import { SheetLog } from "./logUtil";

interface Options {
    method: string;
    headers: {
        "Content-Type": string;
        Accept: string;
    };
    payload?: string;
}

const SheetHttp = {
    URL_GRAPHQL_CAFEF: "https://msh-data.cafef.vn/graphql",
    OPTIONS_POST: {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            Accept: "application/json",
        }
    } as Options,
    OPTIONS_GET: {
        method: "GET",
        headers: {
            "Content-Type": "application/json",
            Accept: "application/json",
        }
    } as Options,
    sendRequest: (url: string): any => {
        try {
            const response = UrlFetchApp.fetch(url);
            return JSON.parse(response.getContentText());
        } catch (e) {
            SheetLog.logDebug(`error: ${e}`);
            return null;
        }
    },
    sendPostRequest: (url: string): any => {
        try {
            const response = UrlFetchApp.fetch(url);
            return JSON.parse(response.getContentText());
        } catch (e) {
            SheetLog.logDebug(`error: ${e}`);
            return null;
        }
    },
    sendGetRequest: (url: string): any => {
        try {
            const response = UrlFetchApp.fetch(url);
            return JSON.parse(response.getContentText());
        } catch (e) {
            SheetLog.logDebug(`error: ${e}`);
            return null;
        }
    },
    sendGraphQLRequest: (url: string, query: string, variables?: any): any => {
        const payload = JSON.stringify({
            query: query,
            variables: variables
        });
        const options = {
            ...SheetHttp.OPTIONS_POST, // Kế thừa các OPTIONS_POST và ghi đè payload
            payload: payload
        };
        return SheetHttp.sendRequest(url);
    }
};
export { SheetHttp }