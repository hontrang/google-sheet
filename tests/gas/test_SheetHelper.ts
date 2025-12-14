/* eslint-disable @typescript-eslint/no-unused-vars */
import { SheetHelper } from '@utils/SheetHelper';

function sheet() {
    QUnit.module('Sheet helper');

    QUnit.test('Test layDuLieuTrongO', function (assert) {
        const helper = new SheetHelper();
        const actual = helper.layDuLieuTrongO('Sheet1', 'A1');

        const expected = 11111;
        assert.equal(actual, expected, 'layDuLieuTrongO');
    });
    QUnit.test('Test layDuLieuTrongHang', function (assert) {
        const helper = new SheetHelper();
        const expected: string[] = ['a', 'b', 'c'];
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
        const expected: unknown[] = [11111, 'a'];
        const actual = helper.layDuLieuTrongCot('Sheet1', 'A');
        assert.deepEqual(actual, expected, 'layDuLieuTrongCot');
    });

}

