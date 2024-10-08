import axios from 'axios';
import { test, expect } from '@playwright/test';
import { ResponseSimplize, ResponseSsi, ResponseTCBS, ResponseVndirect, ResponseVPS } from '@src/types/types';

let TOKEN: string | undefined;
test.describe('kiểm tra url vndirect chạy chính xác', () => {
  test('kiểm tra phản hồi giá chứng khoán từ api vndirect', async () => {
    const tenMa = 'HPG';
    const fromDate = '2024-04-08';
    const toDate = '2024-04-08';
    const URL = `https://finfo-api.vndirect.com.vn/v4/stock_prices?sort=date&q=code:${tenMa}~date:gte:${fromDate}~date:lte:${toDate}&size=1000`;
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
    const url = `https://finfo-api.vndirect.com.vn/v4/financial_statements?q=code:${tenMa}~reportType:QUARTER~modelType:1,89,3,91~fiscalDate:${fiscalDate}&sort=fiscalDate&size=2000`;
    const response = await axios.get(url);
    const datas: [ResponseVndirect] = response.data.data;
    expect(datas[0].code).toEqual(tenMa);
    expect(response.status).toBe(200);
  });
  test('kiểm tra hệ số beta và free float từ vndirect', async () => {
    const tenMa = 'HPG';
    const fromDate = '2023-09-30';
    const url = `https://finfo-api.vndirect.com.vn/v4/ratios/latest?filter=ratioCode:MARKETCAP,NMVOLUME_AVG_CR_10D,PRICE_HIGHEST_CR_52W,PRICE_LOWEST_CR_52W,OUTSTANDING_SHARES,FREEFLOAT,BETA,PRICE_TO_EARNINGS,PRICE_TO_BOOK,DIVIDEND_YIELD,BVPS_CR,&where=code:${tenMa}~reportDate:gt:${fromDate}&order=reportDate&fields=ratioCode,value`;
    const response = await axios.get(url);
    const datas: [ResponseVndirect] = response.data.data;
    expect(datas[0].ratioCode).toEqual('MARKETCAP');
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

test.describe('kiểm tra url tcbs chạy chính xác', () => {
  test('kiểm tra phản hồi từ api tcbs', async () => {
    const tenMa = 'HPG';
    const url = `https://apipubaws.tcbs.com.vn/tcanalysis/v1/company/${tenMa}/large-share-holders`;
    const response = await axios.get(url);
    const datas: [ResponseTCBS] = response.data.listShareHolder;
    expect(datas[0].ticker).toBe(tenMa);
  });
});

test.describe('kiểm tra url ssi chạy chính xác', () => {
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
