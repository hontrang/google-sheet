/* eslint-disable @typescript-eslint/no-extraneous-class */
export class Configuration {
    static readonly consumerID = process.env.CONSUMERID;
    static readonly consumerSecret = process.env.CONSUMERSECRET;
    static readonly mode = 'offline';
    static readonly xlsxInput = './dw.xlsx';
    static readonly xlsxOutput = './build/dw_new.xlsx';
    static readonly dwClasp = '/Users/hontrang/code/clasp/dw';
    static readonly mainClasp = '/Users/hontrang/code/clasp/main';
}