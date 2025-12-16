/**
 * Tests
 * =====
 *
 * List the test functions to run in the TESTS_ array
 *
 * !!! NOTE !!!! V8 looks for definitions in the order the files 
 * are created, so if a new file is created this file - tests.gs - 
 * needs to be deleted and recreated to pick up the contents of 
 * the new file before this file is parsed!
 */

// The title used at the top of the web app
const MAIN_TEST_TITLE_ = 'QUnitGS2 Test'

// The commented out tests mainly rely on being able to run asynchronously, 
// (see trello.com/c/HSDV5BPC) via Promises, "async" or setTimeout(), 
// so will run with limited success, and can only be run individually. As the 
// tests are never deemed to have finished the totals are never returned
// and the subsequent tests never run.  

const TESTS_ = [
    TestSheet,
    TestRequest,
]

// Setup
// =====
// 
// -------- No changes should be need below -------------

// QUnitGS2 is a wrapper for the main QUnit object, which can be accessed directly 
// with "QUnit"
var QUnit = QUnitGS2.QUnit

/**
 * The doGet is the main function that displays the QUnit test web app once it
 * is deployed.
 *
 * It processes the GET from the user's browser and returns the 
 * a HTML template that is populated once the tests are complete
 *
 * Its main purpose is to call the QUnitLib methods:
 * 
 *  - QUnitGS2.init() - Initialize the Apps Script wrapper of the main QUnit library
 *
 *  - TESTS_ - An array of test functions to run, that themselves include calls to QUnit tests
 *
 *  - QUnit.start() - Start the tests running
 *
 *  - QUnitGS2.getHtml() - Return the HTML version of the tests results to
 *                        the web app
 */

function doGet() {

    QUnitGS2.init()

    QUnit.config.title = MAIN_TEST_TITLE_

    TESTS_.forEach((testFunction) => {
        testFunction()
    })

    QUnit.start()

    return QUnitGS2.getHtml()
}

/**
 * The library user has to provide this function to allow the QUnit library
 * to retrieve the test results once they are finished. 
 *
 * The library will then format the results and display them in the web app
 */

function getResultsFromServer() {
    return QUnitGS2.getResultsFromServer()
}
