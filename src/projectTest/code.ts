/* eslint-disable @typescript-eslint/no-unused-vars */
// Optional for easier use.
const QUnit = QUnitGS2.QUnit;
// The title used at the top of the web app
const MAIN_TEST_TITLE_ = 'Stock Test'
const TESTS_ = [
    sheet, // Small set of misc tests used in dev of library, with forced faliures 
    //  step,    // async - wont complete
    //  timeout, // async - wont complete 
    //  assert,  // async - wont complete
    //  async,   // async - wont complete
]
function doGet() {
    QUnitGS2.init(); // Initializes the library.

    QUnit.config.title = MAIN_TEST_TITLE_

    /**
    * Add your test functions here.
    */
    QUnit.module('basic testing');
    TESTS_.forEach((testFunction) => {
        testFunction()
    })

    QUnit.start(); // Starts running tests, notice QUnit vs QUnitGS2.
    return QUnitGS2.getHtml();
}

function getResultsFromServer() {
    return QUnitGS2.getResultsFromServer();
}