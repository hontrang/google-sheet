/* eslint-disable @typescript-eslint/naming-convention */
import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;
export interface ISheetSpread {
  layDuLieuTrongO(sheetName: string, cell: string): string;
  layDuLieuTrongCot(sheetName: string, column: string): string[];
  laySoHangTrongSheet(sheetName: string): number;
  layViTriCotThamChieu(tenMa: string, duLieuCotThamChieu: string[], hangBatDau: number): number;
  layDuLieuTrongHang(sheetName: string, rowIndex: number): string[];
  ghiDuLieuVaoDay(data: any[][], sheetName: string, rowIndex: number, columnIndex: number): void;
  ghiDuLieuVaoDayTheoVung(data: any[][], sheetName: string, range: string): void;
  ghiDuLieuVaoDayTheoTen(data: any[][], sheetName: string, rowNumber: number, columnName: string): void;
  ghiDuLieuVaoO(data: any, sheetName: string, cell: string): boolean;
  doiTenCotThanhChiSo(columnName: string): number;
  chen1HangVaoDauSheet(sheetName: string): boolean;
  xoaCot(sheetName: string, column: string, numOfCol: number): boolean;
  xoaDuLieuTrongCot(sheetName: string, column: string, numOfCol: number, startRow: number): boolean;
}

export interface IHttp {
  sendRequest(url: string, option?: URLFetchRequestOptions): any;
  sendPostRequest(url: string, options?: URLFetchRequestOptions): any;
  sendGetRequest(url: string): Promise<IHttpResponse>;
  getToken(): Promise<string>;
}

export interface IResponseVndirect {
  code?: string;
  interface?: string;
  tradingDate?: string;
  floor?: string;
  buyVal?: number;
  sellVal?: number;
  netVal?: number;
  buyVol?: number;
  sellVol?: number;
  netVol?: number;
  totalRoom?: number;
  currentRoom?: number;
  nmVolume?: number;
  value?: number;
  itemName?: string;
  reportDate?: string;
  itemCode?: string | number;
  relationNameVn: string;
  year: string;
  fiscalDate: string;
  numericValue: number;
  date?: string;
  close?: number;
  ratioCode?: string;
  underlyingType?: string;
  expirationDate?: string;
  expiryDate?: string;
}

export interface IResponseSsi {
  Symbol?: string;
  ClosePrice?: number;
  Market?: string;
  Open?: number;
  High?: number;
  Low?: number;
  Close?: number;
  Volume?: number;
  Time?: string;
  TradingDate?: string;
}

export interface IResponseDC {
  id?: number;
  fund_id?: number;
  created?: string;
  modified?: string;
  assetId?: string;
  translation?: IDCTranslation;
  sector_en?: string;
  exchange?: string;
  bourse_en?: string;
  per_nav?: number;
  data_index_id?: number;
  shares?: number;
  market_value?: number;
  foreign_ownership?: number;
  name_vi?: string;
  weight?: number;
  fundWeight?: IDCFundWeight;
  activeFileName__c?: string;
  downloadUrl__c?: string;
  displayDate__c?: string;
}

export interface IDCTranslation {
  vi?: IDCSectorLevel & IDCIndustryLevel2;
}
export interface IDCSectorLevel {
  sectorLevel?: string;
}
export interface IDCIndustryLevel2 {
  industryLevel2?: string;
}

export interface IDCFundWeight {
  VF1?: string;
  VF4?: string;
}

export interface IHttpResponse {
  data?: any;
  status?: number;
  statusText?: string;
}

export interface IResponseSimplize {
  date?: string;
  priceClose?: number;
  priceOpen?: number;
  priceLow?: number;
  priceHigh?: number;
  netChange?: number;
  pctChange?: number;
  volume?: number;
  cfr?: number;
  bfq?: number;
  sfq?: number;
  fnbsq?: number;
  bfv?: number;
  sfv?: number;
  fnbsv?: number;
  ticker?: string;
  issueDate?: string;
  title?: string;
  attachedLink?: string;
  fileName?: string;
  targetPrice?: number;
  recommend?: RecommendSimplize;
  source?: string;
  content?: string;
  investorFullName?: string;
  pctOfSharesOutHeld?: number;
  changeValue?: number;
  countryOfInvestor?: string;
  totalValue?: number
}

export type RecommendSimplize = 'TRUNG LẬP' | 'MUA' | 'KHÁC';

export interface IResponseVPS {
  sym?: string;
  lastVolume?: number;
  lastPrice?: number;
}
export interface IResponseTCBS {
  ticker?: string;
}
export interface IResponseVietStock {
  KeyCode?: string;
  StockCode?: string;
  BondCode?: string;
  ReleaseDate?: string;
  DueDate?: string;
  FaceValue?: number;
  IssueRate?: number;
  IssuaVolume?: number;
  InterestRateType?: string;
  InterestPeriod?: string;
  TotalRecord?: number;
}
