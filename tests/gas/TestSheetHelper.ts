/* eslint-disable @typescript-eslint/no-unused-vars */
import { SheetHelper } from '@utils/SheetHelper';

function TestSheet() {
    QUnit.module('Sheet helper');

    QUnit.test('Test taoSheetMoi', function (assert) {
        const helper = new SheetHelper();
        helper.taoSheetMoi('Sheet0');
        assert.ok(helper.kiemTraSheetTonTai('Sheet0'), 'taoSheetMoi');
    });
    QUnit.test('Test kiemTraSheetTonTai', function (assert) {
        const helper = new SheetHelper();
        const actual = helper.kiemTraSheetTonTai('Sheet0');
        assert.ok(actual, 'kiemTraSheetTonTai');
    });
    QUnit.test('Test xoaSheet', function (assert) {
        const helper = new SheetHelper();
        helper.xoaSheet('Sheet0');
        assert.notOk(helper.kiemTraSheetTonTai('Sheet0'), 'xoaSheet');
    });

    QUnit.test('Test ghiDuLieuVaoO', function (assert) {
        preCondition();
        const helper = new SheetHelper();
        const expected = 'a';
        const data: string[][] = [[expected]];
        helper.ghiDuLieuVaoO(data, 'Sheet1', 'A1');
        const actual = helper.layDuLieuTrongO('Sheet1', 'A1');
        assert.equal(actual, expected, 'ghiDuLieuVaoO');
    });
    QUnit.test('Test ghiDuLieuVaoDay', function (assert) {
        const helper = new SheetHelper();
        const expected: number[] = [1, 2, 3, 4];
        const data: number[][] = [expected];
        helper.ghiDuLieuVaoDay(data, 'Sheet1', 2, 1);
        const actual = helper.layDuLieuTrongHang('Sheet1', 2);
        assert.deepEqual(actual, expected, 'ghiDuLieuVaoDay');
    });
    QUnit.test('Test ghiDuLieuVaoDayTheoVung', function (assert) {
        const helper = new SheetHelper();
        const expected: number[] = [5, 6, 7, 8];
        const data: number[][] = [expected];
        helper.ghiDuLieuVaoDayTheoVung(data, 'Sheet1', 'A3:D3');
        const actual = helper.layDuLieuTrongHang('Sheet1', 3);
        assert.deepEqual(actual, expected, 'ghiDuLieuVaoDayTheoVung');
    });
    QUnit.test('Test ghiDuLieuVaoDayTheoTen', function (assert) {
        const helper = new SheetHelper();
        const expected: number[] = [9, 10, 11, 12];
        const data: number[][] = [expected];
        helper.ghiDuLieuVaoDayTheoTen(data, 'Sheet1', 4, 'A');
        const actual = helper.layDuLieuTrongHang('Sheet1', 4);
        assert.deepEqual(actual, expected, 'ghiDuLieuVaoDayTheoTen');
    });
    QUnit.test('Test doiTenCotThanhChiSo', function (assert) {
        const helper = new SheetHelper();
        const actual = helper.doiTenCotThanhChiSo('D');
        assert.equal(actual, 4, 'doiTenCotThanhChiSo');
    });
    QUnit.test('Test chen1HangVaoDauSheet', function (assert) {
        const helper = new SheetHelper();
        assert.equal(helper.laySoHangTrongSheet('Sheet1'), 4, 'trước khi chèn hàng là 4')
        helper.chen1HangVaoDauSheet('Sheet1');
        assert.equal(helper.laySoHangTrongSheet('Sheet1'), 5, 'sau khi chèn hàng là 5');
    });
    QUnit.test('Test xoaHang', function (assert) {
        const helper = new SheetHelper();
        assert.equal(helper.laySoHangTrongSheet('Sheet1'), 5, 'trước khi xoá hàng là 5')
        helper.xoaHang('Sheet1', 1);
        assert.equal(helper.laySoHangTrongSheet('Sheet1'), 4, 'sau khi xoá hàng là 4');
    });
    QUnit.test('Test xoaCot', function (assert) {
        const helper = new SheetHelper();
        helper.ghiDuLieuVaoO([[0]], 'Sheet1', 'E1');
        helper.ghiDuLieuVaoO([[0]], 'Sheet1', 'F1');
        helper.ghiDuLieuVaoO([[0]], 'Sheet1', 'G1');
        helper.xoaCot('Sheet1', 'E', 3);
        assert.deepEqual(helper.layDuLieuTrongCot('Sheet1', 'E'), [], 'cột E trống dữ liệu');
        assert.deepEqual(helper.layDuLieuTrongCot('Sheet1', 'F'), [], 'cột F trống dữ liệu');
        assert.deepEqual(helper.layDuLieuTrongCot('Sheet1', 'G'), [], 'cột G trống dữ liệu');
    });
    QUnit.test('Test xoaDuLieuTrongCot', function (assert) {
        const helper = new SheetHelper();
        helper.ghiDuLieuVaoO([[0]], 'Sheet1', 'E10');
        helper.ghiDuLieuVaoO([[0]], 'Sheet1', 'E11');
        helper.ghiDuLieuVaoO([[0]], 'Sheet1', 'F10');
        helper.ghiDuLieuVaoO([[0]], 'Sheet1', 'F11');
        helper.xoaDuLieuTrongCot('Sheet1', 'E', 2, 10, 2);
        assert.equal(helper.layDuLieuTrongO('Sheet1', 'E10'), '', 'cột E10 trống dữ liệu');
        assert.equal(helper.layDuLieuTrongO('Sheet1', 'E11'), '', 'cột E11 trống dữ liệu');
        assert.equal(helper.layDuLieuTrongO('Sheet1', 'F10'), '', 'cột F10 trống dữ liệu');
        assert.equal(helper.layDuLieuTrongO('Sheet1', 'F11'), '', 'cột F11 trống dữ liệu');
    });

    QUnit.test('Test layDuLieuTrongO', function (assert) {
        const helper = new SheetHelper();
        const actual = helper.layDuLieuTrongO('Sheet1', 'A1');

        const expected = 'a';
        assert.equal(actual, expected, 'layDuLieuTrongO');
    });
    QUnit.test('Test layDuLieuTrongHang', function (assert) {
        const helper = new SheetHelper();
        const expected: number[] = [1, 2, 3, 4];
        const actual = helper.layDuLieuTrongHang('Sheet1', 2);
        assert.deepEqual(actual, expected, 'layDuLieuTrongHang');
    });
    QUnit.test('Test layViTriCotThamChieu', function (assert) {
        const helper = new SheetHelper();
        const data: string[] = ['a', 'b', 'c'];
        const actual = helper.layViTriCotThamChieu('b', data, 2);
        assert.equal(actual, 3, 'layDuLieuTrongHang');
    });
    QUnit.test('Test layDuLieuTrongCot', function (assert) {
        const helper = new SheetHelper();
        const expected: unknown[] = ['a', 1, 5, 9];
        const actual = helper.layDuLieuTrongCot('Sheet1', 'A');
        assert.deepEqual(actual, expected, 'layDuLieuTrongCot');
    });
    QUnit.test('Test laySoHangTrongSheet', function (assert) {
        const helper = new SheetHelper();
        const expected = 4;
        const actual = helper.laySoHangTrongSheet('Sheet1');
        assert.equal(actual, expected, 'laySoHangTrongSheet');
        postCondition();
    });

}

function preCondition() {
    const helper = new SheetHelper();
    helper.taoSheetMoi('Sheet1');
}
function postCondition() {
    const helper = new SheetHelper();
    helper.xoaSheet('Sheet1');
}
