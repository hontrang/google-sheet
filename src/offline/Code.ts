import { ExcelHelper } from "./ExcelHelper";


function main() {
    const excel = new ExcelHelper();
    const value = excel.layDuLieuTrongCot("hose", "B");
    console.log(value);
}

main();

