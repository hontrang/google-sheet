import { IResponseDC, IResponseSimplize, IResponseVietStock } from "@src/types/generic";
import { DateHelper } from "@utils/DateHelper";
import { HttpHelper } from "@utils/HttpHelper";

function TestRequest() {
    QUnit.module('Http helper');
    QUnit.test('Test sendRequest', function (assert) {
        const helper = new HttpHelper();
        const URL = 'https://api.simplize.vn/api/company/separate-share/list-tickers';
        const OPTIONS_CO_TUC: URLFetchRequestOptions = {
            method: 'post',
            headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
            payload: JSON.stringify({
                tickers: ['HPG'],
                page: 0,
                size: 10
            })
        };
        const response = helper.sendRequest(URL, OPTIONS_CO_TUC);
        const datas: [IResponseSimplize] = response.data;
        assert.equal(response.status, 200, 'sendRequest status 200')
        assert.notEqual(datas.length, 0, 'sendRequest có dữ liệu nhận về')
    });
    QUnit.test('Test sendGetRequest', function (assert) {
        const URL = 'https://api.simplize.vn/api/company/analysis-report/list?ticker=HPG&isWl=false&page=0&size=10';
        const httpHelper = new HttpHelper();
        const response = httpHelper.sendGetRequest(URL);
        const datas: [IResponseSimplize] = response.data;
        assert.equal(response.status, 200, 'sendGetRequest status 200')
        assert.notEqual(datas.length, 0, 'sendGetRequest có dữ liệu nhận về')
    });
    QUnit.test('Test sendPostRequest', function (assert) {
        const httpHelper = new HttpHelper();
        const URL = `https://www.dragoncapital.com.vn/individual/vi/webruntime/api/apex/execute?language=vi&asGuest=true&htmlEncode=false`;
        const option = JSON.stringify({
            "namespace": "",
            "classname": "@udd/01pJ2000000CgR7",
            "method": "getDocumentContentsV2",
            "isContinuation": false,
            "params": {
                "siteId": "0DMJ2000000oLukOAE",
                "fundCodeOrReportCode": "VF1",
                "documentType": null,
                "targetYear": DateHelper.layNamHienTaiAsString(),
                "language": "vi"
            },
            "cacheable": false
        });

        const configs = {
            "method": "post",
            "contentType": "application/json",
            "payload": option,
            "headers": {
                'Content-Type': 'application/json',
                'Cookie': 'CookieConsentPolicy=0:1; LSKey-c$CookieConsentPolicy=0:1'
            },
            "maxBodyLength": "Infinity",
        };

        const response = httpHelper.sendPostRequest(URL, configs);
        const datas: [IResponseDC] = response.returnValue;
        // assert.equal(response.status, 200, 'sendPostRequest status 200')
        assert.notEqual(datas.length, 0, 'sendPostRequest có dữ liệu nhận về')
    });
}