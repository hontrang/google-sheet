import { ExcelHelper } from "./ExcelHelper";


async function main() {
    const excel = new ExcelHelper();
    await excel.initializeWorkbook();
    const value = excel.layDuLieuTrongO("KN Bán", "B1");
    // const value1 = excel.layDuLieuTrongCot("hose", "B");
    console.log(value);
}

main();

