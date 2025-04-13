import axios from 'axios';
import { test, expect } from '@playwright/test';
import { ResponseDC, ResponseSimplize, ResponseSsi, ResponseTCBS, ResponseVndirect, ResponseVPS } from '@src/types/types';

let TOKEN: string | undefined;
test.describe('kiểm tra url vndirect chạy chính xác', () => {
  test('kiểm tra phản hồi giá chứng khoán từ api vndirect', async () => {
    const tenMa = 'HPG';
    const fromDate = '2024-04-08';
    const toDate = '2024-04-08';
    const URL = `https://api-finfo.vndirect.com.vn/v4/stock_prices?sort=date&q=code:${tenMa}~date:gte:${fromDate}~date:lte:${toDate}&size=1000`;
    const response = await axios.get(URL);
    const datas: [ResponseVndirect] = response.data.data;
    expect(datas[0].code).toEqual(tenMa);
    expect(response.status).toBe(200);
  });
  test('kiểm tra thông tin đơn vị kiểm toán từ api vndirect', async () => {
    const tenMa = 'HPG';
    const url = `https://api-finfo.vndirect.com.vn/v4/company_relations?q=code:${tenMa}~relationType:AUDITOR&size=100&sort=year:DESC`;
    const response = await axios.get(url);
    const datas: [ResponseVndirect] = response.data.data;
    expect(datas[0].code).toEqual(tenMa);
    expect(response.status).toBe(200);
  });
  test('kiểm tra các chỉ số tài chính từ api vndirect', async () => {
    const tenMa = 'HPG';
    const fiscalDate = '2023-09-30';
    const url = `https://api-finfo.vndirect.com.vn/v4/financial_statements?q=code:${tenMa}~reportType:QUARTER~modelType:1,89,3,91~fiscalDate:${fiscalDate}&sort=fiscalDate&size=2000`;
    const response = await axios.get(url);
    const datas: [ResponseVndirect] = response.data.data;
    expect(datas[0].code).toEqual(tenMa);
    expect(response.status).toBe(200);
  });
  test('kiểm tra hệ số beta và free float từ vndirect', async () => {
    const tenMa = 'HPG';
    const fromDate = '2023-09-30';
    const url = `https://api-finfo.vndirect.com.vn/v4/ratios/latest?filter=ratioCode:MARKETCAP,NMVOLUME_AVG_CR_10D,PRICE_HIGHEST_CR_52W,PRICE_LOWEST_CR_52W,OUTSTANDING_SHARES,FREEFLOAT,BETA,PRICE_TO_EARNINGS,PRICE_TO_BOOK,DIVIDEND_YIELD,BVPS_CR,&where=code:${tenMa}~reportDate:gt:${fromDate}&order=reportDate&fields=ratioCode,value`;
    const response = await axios.get(url);
    const datas: [ResponseVndirect] = response.data.data;
    expect(datas[0].ratioCode).toEqual('MARKETCAP');
    expect(response.status).toBe(200);
  });
  test('kiểm tra thông tin phái sinh từ vndirect', async () => {
    const url = `https://api-finfo.vndirect.com.vn/v4/derivative_mappings`;
    const response = await axios.get(url);
    const datas: ResponseVndirect[] = response.data.data;
    expect(datas[0].code).not.toBeNull();
    expect(response.status).toBe(200);
  });
});

test.describe('kiểm tra url simplize chạy chính xác', () => {
  test('kiểm tra phản hồi từ api lấy cổ tức', async () => {
    const tenMa = 'HPG';
    const URL = `https://api.simplize.vn/api/company/separate-share/list-tickers`;
    const response = await axios.post(URL, { tickers: [`${tenMa}`], page: 0, size: 10 });
    const datas: [ResponseSimplize] = response.data.data;
    expect(datas[0].ticker).toEqual(tenMa);
    expect(response.status).toBe(200);
  });

  test('kiểm tra phản hồi từ api báo cáo phân tích', async () => {
    const tenMa = 'HPG';
    const url = `https://api.simplize.vn/api/company/analysis-report/list?ticker=${tenMa}&isWl=false&page=0&size=10`;
    const response = await axios.get(url);
    const datas: [ResponseSimplize] = response.data.data;
    expect(datas[0].ticker).toEqual(tenMa);
  });

  test('kiểm tra danh sách cổ đông lớn', async () => {
    const tenMa = 'HPG';
    const url = `https://api.simplize.vn/api/company/ownership/shareholder-fund-details/${tenMa}`;
    const response = await axios.get(url);
    const datas: ResponseSimplize[] = response.data.data.shareholderDetails;
    expect(datas[0]).toHaveProperty(`investorFullName`);
  });
});

test.describe('kiểm tra url vps chạy chính xác', () => {
  test('kiểm tra phản hồi từ api vps', async () => {
    const ma1 = 'HPG';
    const ma2 = 'STB';
    const DANH_SACH_MA = [ma1, ma2];
    const URL = `https://bgapidatafeed.vps.com.vn/getliststockdata/${DANH_SACH_MA.join(',')}`;
    const response = await axios.get(URL);
    const datas: [ResponseVPS, ResponseVPS] = response.data;
    expect(datas[0].sym).toBe(ma1);
    expect(datas[1].sym).toBe(ma2);
    expect(response.status).toBe(200);
  });
});

test.describe('kiểm tra url cafef chạy chính xác', () => {
  test('kiểm tra sự kiện mới từ api cafef', async () => {
    const tenMa = 'HPG';
    const url = `https://s.cafef.vn/Ajax/Events_RelatedNews_New.aspx?symbol=${tenMa}&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2`;
    const response = await axios.get(url);
    const data = response.data;
    expect(data).toContain('id="divEvents"');
  });

  test('kiểm tra api cafef báo cáo tài chính', async () => {
    const tenMa = 'HPG';
    const url = `https://s.cafef.vn/Ajax/CongTy/BaoCaoTaiChinh.aspx?sym=${tenMa}`;
    const response = await axios.get(url);
    const data = response.data;
    expect(data).toContain('Báo cáo tài chính');
  });
});

test.skip('kiểm tra url ssi chạy chính xác', () => {
  test('kiểm tra url lấy tên mã trên HOSE', async () => {
    const market = 'HOSE';
    const pageIndex = 1;
    const pageSize = 1000;
    const token = await getToken();
    const url = `https://fc-data.ssi.com.vn/api/v2/Market/Securities?lookupRequest.market=${market}&lookupRequest.pageIndex=${pageIndex}&lookupRequest.pageSize=${pageSize}`;
    // eslint-disable-next-line @typescript-eslint/naming-convention
    const response = await axios.get(url, { headers: { Authorization: token } });
    const datas: [ResponseSsi] = response.data.data;
    expect(datas[0].Market).toBe(market);
  });

  test('kiểm tra url SecuritiesDetails', async () => {
    const market = 'HOSE';
    const tenMa = 'HPG';
    const pageIndex = 1;
    const pageSize = 1000;
    const token = await getToken();
    const url = `https://fc-data.ssi.com.vn/api/v2/Market/SecuritiesDetails?lookupRequest.market=${market}&lookupRequest.pageIndex=${pageIndex}&lookupRequest.pageSize=${pageSize}&lookupRequest.symbol=${tenMa}`;
    // eslint-disable-next-line @typescript-eslint/naming-convention
    const response = await axios.get(url, { headers: { Authorization: token } });
    const data = response.data.data[0].RepeatedInfo;
    expect(data[0].Exchange).toBe(market);
    expect(data[0].Symbol).toBe(tenMa);
  });

  test('kiểm tra url GetIndexList', async () => {
    const exchange = 'HOSE';
    const pageIndex = 1;
    const pageSize = 1000;
    const token = await getToken();
    const url = `https://fc-data.ssi.com.vn/api/v2/Market/IndexComponents?lookupRequest.pageIndex=${pageIndex}&lookupRequest.pageSize=${pageSize}&lookupRequest.Exchange=${exchange}`;
    // eslint-disable-next-line @typescript-eslint/naming-convention
    const response = await axios.get(url, { headers: { Authorization: token } });
    const data = response.data;
    expect(data.status).toBe('Success');
  });

  test('kiểm tra url Daily OHLC', async () => {
    const tenMa = 'HPG';
    const pageIndex = 1;
    const pageSize = 1000;
    const fromDate = '04/05/2020';
    const toDate = '04/05/2020';
    const ascending = true;
    const token = await getToken();
    const url = `https://fc-data.ssi.com.vn/api/v2/Market/DailyOhlc?lookupRequest.pageIndex=${pageIndex}&lookupRequest.pageSize=${pageSize}&lookupRequest.fromDate=${fromDate}&lookupRequest.toDate=${toDate}&lookupRequest.ascending=${ascending}&lookupRequest.Symbol=${tenMa}`;
    // eslint-disable-next-line @typescript-eslint/naming-convention
    const response = await axios.get(url, { headers: { Authorization: token } });
    const data = response.data;
    expect(data.status).toBe('Success');
  });

  test('kiểm tra url GetIntradayOHLC', async () => {
    const tenMa = 'HPG';
    const pageIndex = 1;
    const pageSize = 1000;
    const fromDate = '04/05/2020';
    const toDate = '04/05/2020';
    const ascending = true;
    const token = await getToken();
    const url = `https://fc-data.ssi.com.vn/api/v2/Market/IntradayOhlc?lookupRequest.pageIndex=${pageIndex}&lookupRequest.pageSize=${pageSize}&lookupRequest.fromDate=${fromDate}&lookupRequest.toDate=${toDate}&lookupRequest.ascending=${ascending}&lookupRequest.Symbol=${tenMa}`;
    // eslint-disable-next-line @typescript-eslint/naming-convention
    const response = await axios.get(url, { headers: { Authorization: token } });
    const data = response.data;
    expect(data.status).toBe('Success');
  });

  test('kiểm tra url GetStockPrice', async () => {
    const fromDate = '04/05/2020';
    const toDate = '04/05/2020';
    const market = 'HOSE';
    const token = await getToken();
    const url = `https://fc-data.ssi.com.vn/api/v2/Market/DailyStockPrice?&lookupRequest.fromDate=${fromDate}&lookupRequest.toDate=${toDate}&lookupRequest.market=${market}`;
    // eslint-disable-next-line @typescript-eslint/naming-convention
    const response = await axios.get(url, { headers: { Authorization: token } });
    const data = response.data;
    expect(data.status).toBe('Success');
  });
});

test.describe('kiểm tra url vietstock chạy chính xác', () => {
  test('kiểm tra phản hồi về vietstock token', async () => {
    const tenMa = `VND`;
    const headers = {
      Accept: '*/*',
      'Accept-Language': 'vi',
      Connection: 'keep-alive',
      'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
      Cookie: 'language=vi-VN; __RequestVerificationToken=HOrcaOV9AyD405rXuf5nsBOVwtGAq27usAhYlUknoiTKB9BeVyBMRxdfbnJSEF-fqekKEKIYAJHq6rPRoQz5H0Sz90s7OdoD7WSNLgfcnUc1',
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36'
    };

    const options = {
      method: 'get',
      headers: headers,
      muteHttpExceptions: true
    };
    const response = await axios.get(`https://finance.vietstock.vn/${tenMa}/trai-phieu-lien-quan.htm`, options);
    const htmlContent = response.data;
    const regex = `<input.*name=__RequestVerificationToken.*value=([^ >]+)`;
    const match = RegExp(regex).exec(htmlContent);
    expect(response.status).toBe(200);
    expect(match).toHaveLength(2);
  });
  test('kiểm tra phản hồi danh mục trái phiếu', async () => {
    const url = 'https://finance.vietstock.vn/Data/GetBondRelated';
    const tenMa = `VND`;
    const headers = {
      Accept: '*/*',
      'Accept-Language': 'vi',
      Connection: 'keep-alive',
      'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
      Cookie: 'language=vi-VN; __RequestVerificationToken=HOrcaOV9AyD405rXuf5nsBOVwtGAq27usAhYlUknoiTKB9BeVyBMRxdfbnJSEF-fqekKEKIYAJHq6rPRoQz5H0Sz90s7OdoD7WSNLgfcnUc1',
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36'
    };

    const data = {
      __RequestVerificationToken: 'CXImyVGvqSldrS6OJYhhxioGCJYHWNA3Z5DaNLDrwRFZMEKaSTeJWi21Utfue-GpXm3Yb4poiDsRmvrjs01y1rCtZCmd3mfbx7WgjnIMB5Q1',
      code: `${tenMa}`,
      orderBy: 'ReleaseDate',
      orderDir: 'DESC',
      page: 1,
      pageSize: 20
    };

    const options = {
      url: url,
      method: 'post',
      headers: headers,
      data: data,
      muteHttpExceptions: true
    };
    const response = await axios.request(options);
    expect(response.status).toBe(200);
    expect(response.data[0]).toHaveProperty('KeyCode');
  });
});

test.describe('kiểm tra url dragon capital chạy chính xác', () => {
  test('kiểm tra phản hồi báo cáo danh mục', async () => {
    const URL = `https://www.dragoncapital.com.vn/individual/vi/webruntime/api/apex/execute?language=vi&asGuest=true&htmlEncode=false`;
    let option = JSON.stringify({
      namespace: '',
      classname: '@udd/01pJ2000000CgR7',
      method: 'getDocumentContentsV2',
      isContinuation: false,
      params: {
        siteId: '0DMJ2000000oLukOAE',
        fundCodeOrReportCode: 'VF1',
        documentType: null,
        targetYear: '2024',
        language: 'vi'
      },
      cacheable: false
    });
    let config = {
      method: 'post',
      maxBodyLength: Infinity,
      url: `${URL}`,
      headers: {
        'Content-Type': 'application/json',
        Cookie: 'CookieConsentPolicy=0:1; LSKey-c$CookieConsentPolicy=0:1'
      },
      data: option
    };
    const response = await axios.request(config);
    const datas = response.data.returnValue;
    const baoCao: ResponseDC = datas[5].files[0];
    expect(response.status).toBe(200);
    expect(baoCao.activeFileName__c).toContain('Báo cáo');
    expect(baoCao.downloadUrl__c).toContain('dragoncapitalprod');
    expect(baoCao.displayDate__c).toContain('2024');
  });
});

async function getToken(): Promise<string> {
  if (TOKEN !== undefined) return TOKEN;
  else {
    const consumerID = process.env.consumerID;
    const consumerSecret = process.env.consumerSecret;
    const url = `https://fc-data.ssi.com.vn/api/v2/Market/AccessToken`;
    const response = await axios.post(url, { consumerID: consumerID, consumerSecret: consumerSecret });
    TOKEN = 'Bearer ' + response.data.data.accessToken;
    return TOKEN;
  }
}
