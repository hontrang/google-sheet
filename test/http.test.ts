import axios from 'axios';
import { Configuration } from 'src/configuration/Configuration';

let TOKEN: string | undefined;
describe('kiểm tra url vndirect chạy chính xác', () => {
  test('kiểm tra phản hồi từ api vndirect', async () => {
    const tenMa = 'HPG';
    const fromDate = '2024-04-08';
    const toDate = '2024-04-08';
    const URL = `https://finfo-api.vndirect.com.vn/v4/stock_prices?sort=date&q=code:${tenMa}~date:gte:${fromDate}~date:lte:${toDate}&size=1000`;
    const response = await axios.get(URL);
    const data = response.data;
    expect(data).not.toBeNull();
  });
});

describe('kiểm tra url simplize chạy chính xác', () => {
  test('kiểm tra phản hồi từ api lấy cổ tức', async () => {
    const tenMa = 'HPG';
    const URL = `https://api.simplize.vn/api/company/separate-share/list-tickers`;
    const response = await axios.post(URL, { tickers: [`${tenMa}`], page: 0, size: 10 });
    const data = response.data;
    expect(data).not.toBeNull();
    expect(response.status).toBe(200);
  });
});

describe('kiểm tra url vps chạy chính xác', () => {
  test('kiểm tra phản hồi từ api vps', async () => {
    const ma1 = 'HPG';
    const ma2 = 'STB';
    const DANH_SACH_MA = [ma1, ma2];
    const URL = `https://bgapidatafeed.vps.com.vn/getliststockdata/${DANH_SACH_MA.join(',')}`;
    const response = await axios.get(URL);
    const data = response.data;
    expect(data[0].sym).toBe(ma1);
    expect(data[1].sym).toBe(ma2);
  });
});

describe('kiểm tra url cafef chạy chính xác', () => {
  test('kiểm tra phản hồi từ api báo cáo phân tích', async () => {
    const tenMa = 'HPG';
    const URL = `https://edocs.vietstock.vn/Home/Report_ReportAll_Paging?xml=Keyword:${tenMa}&pageIndex=1&pageSize=9`;
    const response = await axios.post(URL);
    const data = response.data.Data;
    expect(data[0].Content).toContain(tenMa);
  });

  test('kiểm tra phản hồi từ api cafef', async () => {
    const tenMa = 'HPG';
    const URL = `https://s.cafef.vn/Ajax/Events_RelatedNews_New.aspx?symbol=${tenMa}&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2`;
    const response = await axios.get(URL);
    const data = response.data;
    expect(data).toContain('id="divEvents"');
  });

  test('kiểm tra phản hồi từ api cafef báo cáo tài chính', async () => {
    const tenMa = 'HPG';
    const URL = `https://s.cafef.vn/Ajax/CongTy/BaoCaoTaiChinh.aspx?sym=${tenMa}`;
    const response = await axios.get(URL);
    const data = response.data;
    expect(data).toContain('Báo cáo tài chính');
  });
});

describe('kiểm tra url tcbs chạy chính xác', () => {
  test('kiểm tra phản hồi từ api vps', async () => {
    const tenMa = 'HPG';
    const URL = `https://apipubaws.tcbs.com.vn/tcanalysis/v1/company/${tenMa}/large-share-holders`;
    const response = await axios.get(URL);
    const data = response.data.listShareHolder;
    expect(data[0].ticker).toBe(tenMa);
  });
});

describe('kiểm tra url ssi chạy chính xác', () => {
  test('kiểm tra url lấy tên mã trên HOSE', async () => {
    const market = 'HOSE';
    const pageIndex = 1;
    const pageSize = 1000;
    const token = await getToken();
    const URL = `https://fc-data.ssi.com.vn/api/v2/Market/Securities?lookupRequest.market=${market}&lookupRequest.pageIndex=${pageIndex}&lookupRequest.pageSize=${pageSize}`;
    // eslint-disable-next-line @typescript-eslint/naming-convention
    const response = await axios.get(URL, { headers: { Authorization: token } });
    const data = response.data.data;
    expect(data[0].Market).toBe(market);
  });

  test('kiểm tra url SecuritiesDetails', async () => {
    const market = 'HOSE';
    const tenMa = 'HPG';
    const pageIndex = 1;
    const pageSize = 1000;
    const token = await getToken();
    const URL = `https://fc-data.ssi.com.vn/api/v2/Market/SecuritiesDetails?lookupRequest.market=${market}&lookupRequest.pageIndex=${pageIndex}&lookupRequest.pageSize=${pageSize}&lookupRequest.symbol=${tenMa}`;
    // eslint-disable-next-line @typescript-eslint/naming-convention
    const response = await axios.get(URL, { headers: { Authorization: token } });
    const data = response.data.data[0].RepeatedInfo;
    expect(data[0].Exchange).toBe(market);
    expect(data[0].Symbol).toBe(tenMa);
  });

  // test("kiểm tra url IndexComponents", async () => {
  //     const indexCode = 'HNX';
  //     const pageIndex = 1;
  //     const pageSize = 1000;
  //     const token = await getToken();
  //     const URL = `https://fc-data.ssi.com.vn/api/v2/Market/IndexComponents?lookupRequest.pageIndex=${pageIndex}&lookupRequest.pageSize=${pageSize}&lookupRequest.indexCode=${indexCode}`;
  //     const response = await axios.get(URL, { headers: { Authorization: token } });
  //     const data = response.data;
  //     expect(data.status).toBe("Success");
  // });

  test('kiểm tra url GetIndexList', async () => {
    const exchange = 'HOSE';
    const pageIndex = 1;
    const pageSize = 1000;
    const token = await getToken();
    const URL = `https://fc-data.ssi.com.vn/api/v2/Market/IndexComponents?lookupRequest.pageIndex=${pageIndex}&lookupRequest.pageSize=${pageSize}&lookupRequest.Exchange=${exchange}`;
    // eslint-disable-next-line @typescript-eslint/naming-convention
    const response = await axios.get(URL, { headers: { Authorization: token } });
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
    const URL = `https://fc-data.ssi.com.vn/api/v2/Market/DailyOhlc?lookupRequest.pageIndex=${pageIndex}&lookupRequest.pageSize=${pageSize}&lookupRequest.fromDate=${fromDate}&lookupRequest.toDate=${toDate}&lookupRequest.ascending=${ascending}&lookupRequest.Symbol=${tenMa}`;
    // eslint-disable-next-line @typescript-eslint/naming-convention
    const response = await axios.get(URL, { headers: { Authorization: token } });
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
    const URL = `https://fc-data.ssi.com.vn/api/v2/Market/IntradayOhlc?lookupRequest.pageIndex=${pageIndex}&lookupRequest.pageSize=${pageSize}&lookupRequest.fromDate=${fromDate}&lookupRequest.toDate=${toDate}&lookupRequest.ascending=${ascending}&lookupRequest.Symbol=${tenMa}`;
    // eslint-disable-next-line @typescript-eslint/naming-convention
    const response = await axios.get(URL, { headers: { Authorization: token } });
    const data = response.data;
    expect(data.status).toBe('Success');
  });

  // test("kiểm tra url GetDailyIndex", async () => {
  //     const indexCode = 'HNX';
  //     const pageIndex = 1;
  //     const pageSize = 1000;
  //     const fromDate = "04/05/2020";
  //     const toDate = "04/05/2020";
  //     const orderBy = "Tradingdate";
  //     const order = "desc";
  //     const token = await getToken();
  //     const URL = `https://fc-data.ssi.com.vn/api/v2/Market/DailyIndex?lookupRequest.pageIndex=${pageIndex}&lookupRequest.pageSize=${pageSize}&lookupRequest.fromDate=${fromDate}&lookupRequest.toDate=${toDate}&lookupRequest.OrderBy=${orderBy}&lookupRequest.Order=${order}&lookupRequest.IndexCode=${indexCode}`;
  //     const response = await axios.get(URL, { headers: { Authorization: token } });
  //     const data = response.data;
  //     expect(data.status).toBe('Success');
  // });

  test('kiểm tra url GetStockPrice', async () => {
    const fromDate = '04/05/2020';
    const toDate = '04/05/2020';
    const market = 'HOSE';
    const token = await getToken();
    const URL = `https://fc-data.ssi.com.vn/api/v2/Market/DailyStockPrice?&lookupRequest.fromDate=${fromDate}&lookupRequest.toDate=${toDate}&lookupRequest.market=${market}`;
    // eslint-disable-next-line @typescript-eslint/naming-convention
    const response = await axios.get(URL, { headers: { Authorization: token } });
    const data = response.data;
    expect(data.status).toBe('Success');
  });
});

async function getToken(): Promise<string> {
  if (TOKEN !== undefined) return TOKEN;
  else {
    const consumerID = Configuration.consumerID;
    const consumerSecret = Configuration.consumerSecret;
    const URL = `https://fc-data.ssi.com.vn/api/v2/Market/AccessToken`;
    const response = await axios.post(URL, { consumerID: consumerID, consumerSecret: consumerSecret });
    TOKEN = 'Bearer ' + response.data.data.accessToken;
    return TOKEN;
  }
}
