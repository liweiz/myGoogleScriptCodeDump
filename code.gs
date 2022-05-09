var pasteBoardSheet = "paste_board";
var dbSheet = "raw_close_data";
var stockDbSheet = "stock_db";
var watchlistsSheet = "watchlists";
var watchlistsSheetContentStartRowNum = 6;
var stockDbSheetTickerStartRowNum = 5;
var numOfWls = 18;

var double1PctRawCSV = "paste_board";

var test_use_ticker = "GT";

// Assuming there the non-empty cells in the given col are all continuous after the input row num.
function numOfContinuousRowsStart(startRowNum, sheetNameStr, colNameStr) {
  var aSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(sheetNameStr);

  var a1Notatoin = colNameStr + startRowNum.toString() + ":" + colNameStr;
  var colNum = aSheet
    .getRange(a1Notatoin)
    .getValues()
    .flat()
    .filter(String)
    .length;

  return colNum - startRowNum + 1;
}

// sort selected range by string
function sortPartialCol(sheetNameStr, colNameStr, startRowNum) {
  var aSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(sheetNameStr);

  var a1Notatoin = colNameStr + startRowNum.toString() + ":" + colNameStr;
  var colNum = aSheet
    .getRange(a1Notatoin)
    .getColumn();

  aSheet
    .getRange(a1Notatoin)
    .sort(colNum);
}

function test_sortPartialCol() {
  Logger.log("START test_sortPartialCol");

  sortPartialCol("Copy of watchlists", "C", watchlistsSheetContentStartRowNum);

  Logger.log("DONE test_sortPartialCol");
}

function sortAllWatchlists(startColNum) {
  var wlSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(watchlistsSheet);

  var aRange = wlSheet.getRange(watchlistsSheetContentStartRowNum, startColNum, 1, numOfWls);

  aRange
    .getValues()[0]
    .forEach((x, i) => sortPartialCol(watchlistsSheet, aRange.getCell(1, i + 1).getA1Notation().substring(0, 1), watchlistsSheetContentStartRowNum));
}

function test_sortAllWatchlists() {
  Logger.log("START test_sortAllWatchlists");

  sortAllWatchlists(2);

  Logger.log("DONE test_sortAllWatchlists");
}

function autoSortAllWatchlists() {
  sortAllWatchlists(2);
}


// returns Sheet
function ensureSheetExist(sheetNameStr) {
  var tickerSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(sheetNameStr);

  if (tickerSheet != null) {
    return tickerSheet;
  }

  return SpreadsheetApp
    .getActiveSpreadsheet()
    .insertSheet(sheetNameStr);
}

function test_ensureSheetExist() {
  Logger.log("START test_ensureSheetExist");

  ensureSheetExist("x");

  Logger.log("DONE test_ensureSheetExist");
}

function getTickerHistoricalQuandlData(tickerStr) {
  // used to build the URL
	var protocol = "https"
	var domain = "www.quandl.com";
	var api_version = "3";
	var root_controller = "datasets";
  var data_base_code = "WIKI" // For Free End of Day US Stock Prices
  var dataset_code = tickerStr;
  var format = "json";
  var api_unique = "";

  // construct the url
	var full_url = protocol + "://" + domain + "/api/v" + api_version + "/" + root_controller + "/" + data_base_code + "/" + dataset_code + "/data." + format + "?api_key=" + api_unique + "&start_date=2018-01-01";

  var result_text = UrlFetchApp.fetch(full_url);

  if (format === "json") {

		// pull out the JSON content 
		var result_json = result_text.getContentText();

		// extract the required data into variables
		var column_names = JSON.parse(result_json)["column_names"];
		var data = JSON.parse(result_json)["data"];

		// append the column headers and the data
		active_sheet_object.appendRow(column_names);
		data.forEach(function(row) {
			active_sheet_object.appendRow(row);
		});

	} else {

		throw "Non-JSON from Quandl NOT supported at this time.";

	};
}

function test_getTickerHistoricalQuandlData() {
  Logger.log("START test_getTickerHistoricalQuandlData");

  getTickerHistoricalQuandlData("X");

  Logger.log("DONE test_getTickerHistoricalQuandlData");
}

// As of Mar 16, 2021, attribute is one of the 5 for historical data.
var googleFinanceAttrs = ["open", "close", "high", "low", "volume"];

function refreshGoogleFinanceData() {

  throw "DO NOT RUN THIS TOO OFTEN";

  var stockSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(stockDbSheet);

  // stackoverflow.com/a/17637159/1373296
  var tickers = stockSheet
    .getRange("A" + stockDbSheetTickerStartRowNum.toString() + ":A")
    .getValues()
    .filter(String)
    .map(x => x[0]);

  var db = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(dbSheet);

  // one col for time info from Google Finance
  var row1ColVals = tickers
    .map(function(x, i) {
      return [x + "_0", x + "_1", x + "_2", x + "_3", x + "_4", x + "_5"];
    })
    .flat();

  var row2ColVals = tickers
    .map(function(x, i) {
      return [`=GOOGLEFINANCE("${x}", "all", DATE(2019,1,1), TODAY(), "DAILY")`, "", "", "", "", ""];
    })
    .flat();

  db.getRange(1, 1, 2, row1ColVals.length).setValues([row1ColVals, row2ColVals]);
}

/**
 * Calculates the EMA of the range.
 *
 * @param {range} range The range of cells to calculate.
 * @param {number} n The number of trailing samples to give higer weight to, e.g. 30.
 * @return The EMA of the range.
 * @customfunction
 */
function EMA(range, n) {
  if (!range.reduce) {
    return range;
  }

  n = Math.min(n, range.length);
  var a = 2 / (n + 1);

  return range.reduce(function(accumulator, currentValue, index, array) {
    return currentValue != "" ? (currentValue * a + accumulator * (1-a)) : accumulator;
  }, 0);
}

// Based on EMA above, Smoothed Moving Average
function SMMA(range, n) {
  if (!range.reduce) {
    return range;
  }

  return range.reduce(function(accumulator, currentValue, index, array) {
    return currentValue != "" ? ((currentValue + accumulator * (n - 1)) / n) : accumulator;
  }, 0);
}

// stackoverflow.com/a/30399727/1373296
function diff(ary) {
    var newA = [];
    for (var i = 1; i < ary.length; i++)  newA.push(ary[i] - ary[i - 1])
    return newA;
}

// Searching is case sensitive
// returns col num of the ticker, according to Google Sheets, Start from 1, not 0.
function tickerIsInRawDb(tickerStr) {
  var db = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(dbSheet);

  var tickerStrs = db.getRange("A1:1").getValues()[0];

  for (var i = 0; i < tickerStrs.length; i++) {
    if (tickerStrs[i] === tickerStr) {
      return i + 1;
    }
  }

  Logger.log(`Ticker "${tickerStr}" NOT found in db`);
}

function test_tickerIsInRawDb() {
  var ticker = test_use_ticker;
  var r = tickerIsInRawDb(ticker);
  if (r === undefined) {
    Logger.log(`test_tickerIsInRawDb DONE: Ticker "${ticker}" not found`);
  } else {
    Logger.log(`test_tickerIsInRawDb DONE: Ticker "${ticker}" Col num: ${r}`);
  }
}

// returns first col num to set value
// rowNum: the row num to operate on
// colSpaceNum: the num of cols each occupied unit needs, e.g.: 2 indicats 2 cells in each unit
function firstAvailableColNum(sheetNameStr, rowNum, colSpaceNum) {
  var s = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(sheetNameStr);

  var rowContents = s.getRange("A"+ rowNum.toString() + ":" + rowNum.toString()).getValues()[0];

  var l = rowContents.length;

  for (var i = 0; i < l; i++) {
    var j = colSpaceNum * i;
    if (rowContents[j] === "") {
      return j + 1;
    } else if (j * i >l) {
      // beyond current row contents length
      return j + 1;
    }
  }
}

// returns existing or newly appended ticker's col num
function ensureTickerInRawDbExist(ticker) {
  var rowNum = 1;
  var colNum = tickerIsInRawDb(ticker);
  if (colNum === undefined) {
    // append
    colNum = firstAvailableColNum(dbSheet, rowNum, 2);

    Logger.log(`colNum: "${colNum}"`);

    var s = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(dbSheet)
    .getRange(rowNum, colNum, 1, 1)
    .setValue(ticker);

  }

  return colNum;
}

function test_ensureTickerInRawDbExist() {
  var colNum = ensureTickerInRawDbExist("OK");
  Logger.log(`test_ensureTickerInRawDbExist DONE: Ticker exists at column "${colNum}"`);
  var colNum = ensureTickerInRawDbExist(test_use_ticker);
  Logger.log(`test_ensureTickerInRawDbExist DONE: Ticker exists at column "${colNum}"`);
}

// range params: [rowNum, colNum, rowSum, colSum] in JS Array.
function copyThenPaste(sheetNameStrToCopy, sheetNameStrToPaste, rangeToCopy, rangeToPaste) {
  var s = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(sheetNameStrToPaste);
  SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(sheetNameStrToCopy)
    .getRange(rangeToCopy[0], rangeToCopy[1], rangeToCopy[2], rangeToCopy[3])
    .copyValuesToRange(s.getSheetId(), rangeToPaste[0], rangeToPaste[1], rangeToPaste[2], rangeToPaste[3]);
}

function test_copyThenPaste() {
  copyThenPaste(pasteBoardSheet, dbSheet, [5, 6, 2, 2], [3, 3 + 3 - 1, 4, 4 + 3 - 1]);
  Logger.log(`test_copyThenPaste DONE"`);
}

// returns first result of given ticker string in watchlist
// returns undefined if not found
function watchlistHasTicker(tickerStr, colNameStr, startRowNum) {
  var watchlists = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("watchlists");

  // stackoverflow.com/a/17637159/1373296
  var tickers = watchlists
    .getRange(colNameStr + startRowNum.toString() + ":" + colNameStr)
    .getValues()
    .filter(String)
    .map(x => x[0]);

  return (tickers.indexOf(tickerStr) > -1);
}

function test_watchlistHasTicker() {

  Logger.log("START test_watchlistHasTicker");

  var r1 = watchlistHasTicker(test_use_ticker, "B", watchlistsSheetContentStartRowNum);
  var r2 = watchlistHasTicker(test_use_ticker, "C", watchlistsSheetContentStartRowNum);
  var r3 = watchlistHasTicker("K" + test_use_ticker, "B", watchlistsSheetContentStartRowNum);

  Logger.log(`existing ticker returns ${r1}`);
  Logger.log(`non existing ticker in empty watchlist returns ${r2}`);
  Logger.log(`non existing ticker returns ${r3}`);

  Logger.log("DONE test_watchlistHasTicker");
}

// returns first watchlist having given ticker string in watchlists
// returns undefined if not found
function watchlistForTicker(tickerStr) {
  var watchlists = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(watchlistsSheet);

  var wlCells = watchlists
    .getRange(1, 2, 1, numOfWls);

  var numsOfWL = wlCells
    .getValues();

  for (var rowIndex = 0; rowIndex < numsOfWL.length; rowIndex++) {
    for (var colIndex = 0; colIndex < numsOfWL[rowIndex].length; colIndex++) {
      var colNameStr = wlCells
        .getCell(rowIndex + 1, colIndex + 1)
        .getA1Notation()
        .substring(0, 1);

      if (watchlistHasTicker(tickerStr, colNameStr, watchlistsSheetContentStartRowNum)) {
        var watchlistNameRowNum = watchlistsSheetContentStartRowNum - 1;
        return watchlists.getRange(colNameStr + watchlistNameRowNum.toString()).getValue().toString();
      }
    }
  }
}

function test_watchlistForTicker() {

  Logger.log("START test_watchlistHasTicker");

  var r1 = watchlistForTicker(test_use_ticker);
  var r2 = watchlistForTicker(test_use_ticker + "K");

  Logger.log(`existing ticker returns ${r1}`);
  Logger.log(`non existing ticker returns ${r2}`);

  Logger.log("DONE test_watchlistForTicker");
}

function getNewDouble1Pct(rawSheetNameStr) {
  // get sheet with raw data to check
  var rawSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(rawSheetNameStr);

  var rangeToFilter = rawSheet.getRange("A:I");

  var blackListedTickers = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("blacklist")
    .getRange("A2:R")
    .getValues()
    .flat()
    .filter(String);

  var watchedTickers = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(watchlistsSheet)
    .getRange("B6:R")
    .getValues()
    .flat()
    .filter(String);

  var tickersToExclude = [blackListedTickers, watchedTickers].flat();

  return rangeToFilter
    .getValues()
    .filter(row => row[7] >= 0.008 && row[7] <= 0.013 && row[8] >= 0.008 && row[8] <= 0.013)
    .filter(row => !row[2].toLowerCase().includes("etf") && !row[2].toLowerCase().includes("fund") && !row[2].toLowerCase().includes("etn") && !row[2].toLowerCase().includes("trust") && !row[2].toLowerCase().includes("shares"))
    .filter(row => tickersToExclude.indexOf(row[1]) === -1)
    .map(row => [row[1], row[2], row[7], row[8]]);
}

function test_getNewDouble1Pct() {

  Logger.log("START test_getNewDouble1Pct");

  var r = getNewDouble1Pct(double1PctRawCSV);

  Logger.log(`getNewDouble1Pct: ${r}`);

  var sheetNameStr = "temp_imported_pool";

  var poolSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(sheetNameStr);

  poolSheet.clear();

  poolSheet
    .getRange(1, 1, r.length, 4)
    .setValues(r);

  Logger.log("DONE test_getNewDouble1Pct");
}

// returns an array, each element is an array having a watchlist name as first elem and col num as second one.
function getStockDbWatchlistNameToColNumMap(startColNum, rowNum, watchlistSum) {
  var dbSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(stockDbSheet);

  return dbSheet
    .getRange(rowNum, startColNum, 1, watchlistSum)
    .getValues()[0]
    .map((x, i) => [x, startColNum + i]);
}

function test_getStockDbWatchlistNameToColNumMap() {
  Logger.log("START test_getStockDbWatchlistNameToColNumMap");

  var r = getStockDbWatchlistNameToColNumMap(2, stockDbSheetTickerStartRowNum - 1, numOfWls);

  r.forEach(x => Logger.log(`${x[0]}: ${x[1]}`));

  Logger.log("DONE test_getStockDbWatchlistNameToColNumMap");
}

function setOneStockDbWatchlistName(tickerStr, wlName, wlNameToColNumPairs, rangeSelected) {
  rangeSelected.clear();
  var i = wlNameToColNumPairs.findIndex(x => x[0] === wlName);
  if (i > -1) {
    var colNum = rangeSelected.getColumn();
    rangeSelected.getCell(1, wlNameToColNumPairs[i][1] - colNum + 1).setValue(1);
  } else {
    throw `ERROR [setOneStockDbWatchlistName] ${tickerStr}'s wishlist name ${wlName} not in db.`;
  }
}

function test_setOneStockDbWatchlistName() {
  Logger.log("START test_setOneStockDbWatchlistName");

  var rangeSelected = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(stockDbSheet)
    .getRange(5, 2, 1, 18);

  var pairs = getStockDbWatchlistNameToColNumMap(2, stockDbSheetTickerStartRowNum - 1, numOfWls);

  setOneStockDbWatchlistName(test_use_ticker, "持有", pairs, rangeSelected);

  Logger.log("DONE test_setOneStockDbWatchlistName");
}

function refreshAllStockDbWatchlistNames() {
  var db = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(stockDbSheet);

  // stackoverflow.com/a/17637159/1373296
  var tickers = db
    .getRange("A" + stockDbSheetTickerStartRowNum.toString() + ":A")
    .getValues()
    .filter(String)
    .map(x => x[0]);

  var pairs = getStockDbWatchlistNameToColNumMap(2, stockDbSheetTickerStartRowNum - 1, numOfWls);

  tickers.forEach(function(x, i) {
    var wlName = watchlistForTicker(x);
    var rangeSelected = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(stockDbSheet)
      .getRange(stockDbSheetTickerStartRowNum + i, 2, 1, numOfWls);

    setOneStockDbWatchlistName(x, wlName, pairs, rangeSelected);
  });
}

function refreshRSI() {
  var watchlists = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("watchlists");

  var wlCells = watchlists
    .getRange(1, 2, 1, 18);

  var numsOfWL = wlCells
    .getValues();

  for (var rowIndex = 0; rowIndex < numsOfWL.length; rowIndex++) {
    for (var colIndex = 0; colIndex < numsOfWL[rowIndex].length; colIndex++) {
      var numOfWL = numsOfWL[rowIndex][colIndex];

      var draftSheetName = "wl_" + numOfWL;

      Logger.log(draftSheetName);

      Logger.log(rowIndex + "," + colIndex);

      Logger.log(wlCells);

      var colNameStr = wlCells
        .getCell(rowIndex + 1, colIndex + 1)
        .getA1Notation()
        .substring(0, 1);

      Logger.log(colNameStr);

      if (colIndex + 1 <= 26) {
        // Col from "A" to "Z"

        // stackoverflow.com/a/17637159/1373296
        var tickers = watchlists
          .getRange(colNameStr + watchlistsSheetContentStartRowNum.toString() + ":" + colNameStr)
          .getValues()
          .filter(String);

        Logger.log("tickers: " + tickers);

        // done getting a watchlist, start calculating in raw sheet

        var draftSheet = SpreadsheetApp
          .getActiveSpreadsheet()
          .getSheetByName(draftSheetName);

        draftSheet.clear();

        for (var i = 0; i < tickers.length; i++) {
          var ticker = tickers[i];
          
          draftSheet
            .getRange(1, i * 2 + 1, 1, 1)
            .setValue(`${ticker}`);

          draftSheet
            .getRange(2, i * 2 + 1, 1, 1)
            .setValue(`=GoogleFinance("${ticker}","price","01/01/2020",TODAY(),"daily")`);

          var isPopulated = false;

          do {
            Utilities.sleep(10);
            var newValue = draftSheet
            .getRange(3, i * 2 + 1, 1, 1)
            .getValue();
            if (newValue != '') {
              isPopulated = true;
            }
          } while (!isPopulated);

          

          // stackoverflow.com/a/65593042/1373296
          var lastRowNum = draftSheet
            .getRange(1, i * 2 + 1, 500, 1)
            .getValues()
            .map(x => x[0])
            .indexOf('');

          Logger.log("lastRowNum: " + lastRowNum);

          draftSheet
            .getRange(1, i * 2 + 2, 1, 1)
            .setValue(lastRowNum);

          var numOfDays = 3;
          var numOfDiffs = numOfDays - 1;

          var firstRowNum = lastRowNum - numOfDays + 1;

          // stackoverflow.com/q/2876536/1373296

          var closes = draftSheet
            .getRange(firstRowNum, i * 2 + 2, numOfDays, 1)
            .getValues()
            .map(x => x[0] * 100);

          var closeDiffs = diff(closes);

          var ups = closeDiffs.map(x => x < 0 ? 0 : x);

          var bothAbs = closeDiffs.map(x => x < 0 ? -x : x);

Logger.log("closes: " + closes);
Logger.log("closeDiffs: " + closeDiffs);
Logger.log("ups: " + ups);
Logger.log("bothAbs ups: " + bothAbs);

          // Test if sum is correct, debug only
          // draftSheet
          //   .getRange(1, i * 2 + 2, 1, 1)
          //   .setValue(closes.reduce((a, b) => a + b, 0));

          Logger.log("SMMA ups: " + SMMA(ups, numOfDiffs));

          Logger.log("SMMA both abs: " + SMMA(bothAbs, numOfDiffs));

          var modifiedRSI = ((100 * SMMA(ups, numOfDiffs) / SMMA(bothAbs, numOfDiffs) * 5000 / 5455) * 100 - 7 - 5000) / 100;

          draftSheet
            .getRange(1, i * 2 + 2, 1, 1)
            .setValue(modifiedRSI);
          }

        

      } else {
        Logger.log("Some watchlists in col 'AA' or beyond.");
      }
    }
  }
}

// webapps.stackexchange.com/questions/134812/list-all-sheet-names-google-sheets
function sheetnames() {
  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheets()
    .map(s => s.getName())
  ;
}

// Assuming most important wl in sheet "watchlists" on the left and least on the right
function getWatchlistNamesWithPriorities() {
  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(watchlistsSheet)
    .getRange("B5:5")
    .getValues()
    .flat()
    .filter(String);
}

var wlNamesWithPriorities = getWatchlistNamesWithPriorities();

// returns a 2D Array, each root element is content of a watchlist.
// the first element in a root element is the name of the corresponding watchlist followed by each member tickers.
function getEachWatchlistTickers() {

  var initialValue = [[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[]];

  if (initialValue.length !== wlNamesWithPriorities.length) {
    throw Error("func getEachWatchlistTickers: initialValue to collect all watchlist items NOT the same size as number of watchlists")
  }

  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(watchlistsSheet)
    .getRange("B5:R")
    .getValues()
    .reduce((accumulator, currentValue) => accumulator.map(function (col, i) { 
      col.push(currentValue[i])
      return col;
      })
      , initialValue)
    .map(x => x.filter(String));
}

function test_getEachWatchlistTickers() {
  Logger.log(getEachWatchlistTickers());
}

var rootDirName = "dashboard试验";
var rawCsvDirName = "csv_raw";
var outputDirName = "output";
var sheetNameToSetCSVData = "Data";

// Assuming "template" Spreadsheet is in root dir
var boilerplateSpreadsheetName = "template";
var boilerplateSpreadsheet;

var csvStartRowNum = 3;
var csvStartColNum = 2;

// Assuming csv is downloaded from Yahoo Finance historical data
var csvNumOfCols = 7;

var longRangeChartNameStr = "Char-L";
var smallRangeChartNameStr = "Char-S";
var zoomRangeChartNameStr = "Char-Z";

function chartsInLSZOrder(aSpreadsheet) {
  function compare(a, b) {
    if (a[1] > b[1]) {
      return -1;
    }
    if (a[1] < b[1]) {
      return 1;
    }
    return 0;
  }

  var charts =  aSpreadsheet
    .getSheetByName(sheetNameToSetCSVData)
    .getCharts()
    .map(chart => [chart, chart.getRanges()[0].getHeight()])
    .sort(compare);
    
  return charts.map(x => x[0]);
}

var zoomStartRowNum = 506;
var smallStartRowNum = 301;
var longStartRowNum = 3;
var chartDataEndRowNum = 1001;
var zoomStartDate = "04/01/2021";
var smallStartRowNum = "03/11/2020";

// Add mark at the beginning of original file name, e.g, GOOG => 1_GOOG.
// 1: There is at least one double action in the last date record
// 9: There is no double action in the last date record
// aFile: Class File
function renameBasedOnLastlineDoubleActionData(aFile) {
  var aSpreadSheet = SpreadsheetApp.open(aFile);
  var dataSheet = aSpreadSheet.getSheetByName(sheetNameToSetCSVData);
  var numOfDates = dataSheet
    .getRange("B" + longStartRowNum.toString() + ":B")
    .getValues()
    .flat()
    .filter(String)
    .length;
  var lastRowNum = longStartRowNum - 1 + numOfDates;
  var values = dataSheet
    .getRange("Q" + lastRowNum.toString() + ":V" + lastRowNum.toString())
    .getValues()[0];
  var numOfValues = values
    .filter(String)
    .length;
  var newName = aFile.getName();
  if (numOfValues === 0) {
    newName = "9_" + newName;
  } else {
    if (values[0] !== "" || values[3] !== "") {
      newName = "1_" + newName;
    } else if (values[1] !== "" || values[4] !== "") {
      newName = "2_" + newName;
    } else if (values[2] !== "" || values[5] !== "") {
      newName = "3_" + newName;
    }
  }
  aFile.setName(newName);
}

// NOT WORKING: keep getting ERROR "Those columns are out of bounds." when updateChart called. No luck on googling solution.
// aSpreadsheet: Class Spreadsheet
function adaptChartsInNewCopy(aSpreadsheet, longRangeMax, smallRangeMax, zoomRangeMax) {
  var chartsInOrder = chartsInLSZOrder(aSpreadsheet);
  [longRangeChartNameStr, smallRangeChartNameStr, zoomRangeChartNameStr]
    .forEach(function (charName, i) {

      var chart = chartsInOrder[i];

      chart
        .modify()
        .setOption("vAxes.viewWindow.max", [longRangeMax, smallRangeMax, zoomRangeMax][i])
        // .setOption("vAxis.maxValue", [longRangeMax, smallRangeMax, zoomRangeMax][i])
        .build();

      aSpreadsheet
        .getSheetByName(sheetNameToSetCSVData)
        .updateChart(chart);

      var chartSheet = aSpreadsheet.moveChartToObjectSheet(chart);
      chartSheet.setName(charName);
      aSpreadsheet.setActiveSheet(chartSheet);
      aSpreadsheet.moveActiveSheet(1);
    });
}

// get y axis max value, given a set of data's max value
function roundUpMaxY(y) {
  var str = Number.parseFloat(y).toExponential();
  var numOfTens = str.charAt(str.length - 1);
  var part1 = Number(str.split("e")[0]);
  var part2 = Number(numOfTens);
  return Math.ceil(part1) * Math.pow(10, part2); 
}

function dataRangeStartRowNum(startDateStr, aSpreadsheet, dataSheetName, dateColName) {
  var i = aSpreadsheet
    .getSheetByName(dataSheetName)
    .getRange(dateColName + "1:" + dateColName)
    .getValues()
    .flat(2)
    .indexOf(startDateStr);

  if (i === -1) {
    return csvStartRowNum;
  }
  
  return i + 1;
}

function createChartsInSheet(aSpreadsheet, sheetName) {
  var dateStrColName = "AA";

  var dataRangeStartRowNums = [csvStartRowNum, smallStartRowNum, zoomStartDate].map(function (x, i) {
    if (i === 0) {
      return x;
    }

    return dataRangeStartRowNum(x, aSpreadsheet, sheetName, dateStrColName);
  });

  // stackoverflow.com/a/6102340/1373296
  var lszMaxes = dataRangeStartRowNums.map(function(x) {
    var targetRange = "Q" + x.toString() + ":" + "V" + chartDataEndRowNum.toString();
    return Math.max(...(aSpreadsheet.getSheetByName(sheetName).getRange(targetRange).getValues().flat().filter(String).map(x => parseFloat(x))));
  });

  [longRangeChartNameStr, smallRangeChartNameStr, zoomRangeChartNameStr]
    .forEach(function (charNamePart1, i) {
      ["-k", ""]
        .map(charNamePart2 => charNamePart1 + charNamePart2)
        .forEach(function (charName, j) {
          if (!(i === 2 && j === 1)) {
            var startRangeRowNum = dataRangeStartRowNums[i];
            var pSize = [2, 3, 7][i];

            var aSheet = aSpreadsheet.getSheetByName(sheetName);
            var chartBuilder = aSheet.newChart();
            var dateRange = aSheet.getRange("J" + startRangeRowNum.toString() + ":J" + chartDataEndRowNum.toString());
            var squash1Range = aSheet.getRange("Q" + startRangeRowNum.toString() + ":Q" + chartDataEndRowNum.toString());
            var squash2Range = aSheet.getRange("R" + startRangeRowNum.toString() + ":R" + chartDataEndRowNum.toString());
            var squash3Range = aSheet.getRange("S" + startRangeRowNum.toString() + ":S" + chartDataEndRowNum.toString());
            var push1Range = aSheet.getRange("T" + startRangeRowNum.toString() + ":T" + chartDataEndRowNum.toString());
            var push2Range = aSheet.getRange("U" + startRangeRowNum.toString() + ":U" + chartDataEndRowNum.toString());
            var push3Range = aSheet.getRange("V" + startRangeRowNum.toString() + ":V" + chartDataEndRowNum.toString());
            var colors = ["#01DF3A", "#A9F5A9", "#E0F8E0", "#FF0000", "#F5A9A9", "#F8E0E0"];

            var chartBase = chartBuilder
              .asScatterChart()
              .addRange(dateRange)
              .addRange(push1Range)
              .addRange(push2Range)
              .addRange(push3Range)
              .addRange(squash1Range)
              .addRange(squash2Range)
              .addRange(squash3Range)
              .setOption("series", {
                0: { pointSize: pSize, lineWidth: 0 },
                1: { pointSize: pSize, lineWidth: 0 },
                2: { pointSize: pSize, lineWidth: 0 },
                3: { pointSize: pSize, lineWidth: 0 },
                4: { pointSize: pSize, lineWidth: 0 },
                5: { pointSize: pSize, lineWidth: 0 }
              });

            // To make same-period chart with and without daily high/low have same vAxis height, which leads to them look overlapped so that easier to use
            // only use the dailyHighMaxes for both with and without daily low/high chart.
            // var yMax = roundUpMaxY(lszMaxes[i]);
            var dailyHighMaxes = dataRangeStartRowNums.map(function(x) {
                var targetRange = "W" + x.toString() + ":" + "W" + chartDataEndRowNum.toString();
                return Math.max(...(aSpreadsheet.getSheetByName(sheetName).getRange(targetRange).getValues().flat().filter(String).map(x => parseFloat(x))));
              });
            var yMax = roundUpMaxY(dailyHighMaxes[i]);

            if (j === 0) {
              var dailyHighRange = aSheet.getRange("W" + startRangeRowNum.toString() + ":W" + chartDataEndRowNum.toString());
              var dailyLowRange = aSheet.getRange("X" + startRangeRowNum.toString() + ":X" + chartDataEndRowNum.toString());
              var smallPointSize = 0.7;

              colors.push("#D8D8D8");
              colors.push("#6E6E6E");
              
              chartBase
                .addRange(dailyHighRange)
                .addRange(dailyLowRange);

              if (i === 2) {
                chartBase.setOption("series", {
                  6: { pointSize: pSize * 0.4, lineWidth: 0 },
                  7: { pointSize: pSize * 0.4, lineWidth: 0 }
                });
              } else {
                chartBase.setOption("series", {
                  6: { pointSize: smallPointSize, lineWidth: 0 },
                  7: { pointSize: smallPointSize, lineWidth: 0 }
                })
              }
            }

            var chart = chartBase
              .setYAxisRange(0, yMax)
              .setOption("colors", colors)
              .setPosition(2, 2, 0, 0)
              .setOption("hAxis.gridlines.count", 10)
              .setOption("hAxis.gridlines.color", "#BDBDBD")
              .setOption("hAxis.minorGridlines.count", 10)
              .setOption("vAxis.gridlines.count", 10)
              .setOption("vAxis.gridlines.color", "#BDBDBD")
              .setOption("vAxis.minorGridlines.count", 10)
              .setOption("hAxis.viewWindowMode", "maximized")
              .setOption("vAxis.viewWindowMode", "maximized")
              // .setOption("vAxis.format", "#.##") // NOT WORKING
              .setOption("legend", "none")
              .setOption("focusTarget", "category")
              .build();

            aSheet.insertChart(chart);

            var chartSheet = aSpreadsheet.moveChartToObjectSheet(chart);
            chartSheet.setName(charName);
            aSpreadsheet.setActiveSheet(chartSheet);
            aSpreadsheet.moveActiveSheet(1);
          }
        })
  });
}

/* 
  Clear all existing folders in output dir and create a new one with lastest trading date as
  folder name.

  Under the new folder, based on tickerNameArray's watchlists, create new folders for each
  watchlist.

  Based on existing available CSVs in raw csv storage folder, combine with template by 
  duplicating, dumping csv data, all steps in corresponding watchlist folder.

  Watchlisted without CSV available and non-watchlisted with CSV available tickers will not be processed.

  Assuming dir path: rootDirStr/csvDirStr/

  tickerNameArray: 
  a 2D Array, each root element is content of a watchlist.
  the first element in a root element is the name of the corresponding watchlist
  followed by each member tickers.
 */
function turnCSVs2ItemsInWatchlists(tickerNameArray, rootDirStr) {
  var processedTickers = [];
  var unprocessedTickers = [];
  var unlistedTickers = [];
  var outputDir;
  var outputSubDirName; // Set the last trading date as folder name
  var outputSubDir;
  var wlFolders;
  var csvData;

  var wlDirNames = tickerNameArray.map(x => x[0]).flat();

  var selectedRootFolders = DriveApp.getFoldersByName(rootDirStr);

  if (selectedRootFolders.hasNext()) {
    var rootDir = selectedRootFolders.next();
    var subDirs = rootDir.getFolders();

    while (subDirs.hasNext()) {
      var folder = subDirs.next();

      if (folder.getName() === rawCsvDirName) {
        var csvDir = folder;
        var csvs = csvDir.getFilesByType("text/csv");

        var everything = tickerNameArray.flat();
        
        // work based on available CSV files
        while (csvs.hasNext()) {
          var csvFile = csvs.next();
          var ticker = csvFile.getName().split(".")[0];

          if (everything.indexOf(ticker) === -1) {
            // csv file not in watchlist
            unlistedTickers.push(ticker);

          } else {
            // csv file in watchlist, process

            // stackoverflow.com/a/51331448/1373296
            csvData = Utilities.parseCsv(csvFile.getBlob().getDataAsString());

            if (typeof outputSubDirName === "undefined") {
              outputSubDirName = csvData[csvData.length - 1][0];
              
              // Get output folder
              var kids = rootDir.getFoldersByName(outputDirName);
              while (kids.hasNext()) {
                outputDir = kids.next();
                break;
              }

              // Clean all existing output
              var outputSubFiles = outputDir.getFiles();
              while (outputSubFiles.hasNext()) {
                outputSubFiles.next().setTrashed(true);
              }

              var outputSubFolders = outputDir.getFolders();
              while (outputSubFolders.hasNext()) {
                outputSubFolders.next().setTrashed(true);
              }

              // Create new date-named folder
              outputSubDir = outputDir.createFolder(outputSubDirName);

              // Get template
              var spreadsheetFiles = rootDir.getFilesByType("application/vnd.google-apps.spreadsheet");
              while (spreadsheetFiles.hasNext()) {
                var spreadsheetFile = spreadsheetFiles.next();
                if (spreadsheetFile.getName() === boilerplateSpreadsheetName) {
                  boilerplateSpreadsheet = spreadsheetFile;
                  break;
                }
              }

              if (typeof outputSubDir != "undefined" && typeof boilerplateSpreadsheet != "undefined") {
                // Create new folders for each watchlist under the date-named folder
                wlFolders = wlDirNames.map(x => outputSubDir.createFolder(x));
              }
            }

            tickerNameArray.forEach(function (tickers, i) {
              if (tickers.indexOf(ticker) !== -1) {
                var newCopy = boilerplateSpreadsheet.makeCopy(ticker, wlFolders[i]);

                SpreadsheetApp
                  .open(newCopy)
                  .getSheetByName(sheetNameToSetCSVData)
                  .getRange(csvStartRowNum, csvStartColNum, csvData.length, csvNumOfCols)
                  .setValues(csvData);
                        
                processedTickers.push(ticker);
              }
            });
          }
        }
      }
    }
  }

  tickerNameArray
    .map(x => x.slice(1))
    .flat()
    .forEach(function (x) {
      if (processedTickers.indexOf(x) === -1) {
        unprocessedTickers.push(x);
      }
    });

  Logger.log("Tickers in Watchlists w/o CSV files, total " + unprocessedTickers.length.toString());
  Logger.log(unprocessedTickers);
  Logger.log("Tickers NOT in Watchlists w/ CSV files, not processed, total " + unlistedTickers.length.toString());
  Logger.log(unlistedTickers);
  Logger.log("Tickers processed in Watchlists w/ CSV files, total " + processedTickers.length.toString());
  Logger.log(processedTickers);
}

function test_turnCSVs2ItemsInWatchlists() {
  var l = getEachWatchlistTickers();
  turnCSVs2ItemsInWatchlists(l, rootDirName);
}

/*
  Sample URL:
  finance.yahoo.com/quote/NEON/history?period1=1546300800&period2=1619827200&interval=1d&filter=history&frequency=1d&includeAdjustedClose=true
  query1.finance.yahoo.com/v7/finance/download/XYL?period1=1546300800&period2=1619827200&interval=1d&events=history&includeAdjustedClose=true
*/

// Tue Jan 01 2019 00:00:00 GMT+0000
let startDateStr = "1546300800";

// stackoverflow.com/a/28431880/1373296
// stackoverflow.com/a/23081260/1373296
var tomorrow = new Date();
tomorrow.setDate(new Date().getDate()+1);
var tomorrowDateStr = tomorrow.toISOString().substring(0, 10).split("-").join(".");

// stackoverflow.com/a/28683720/1373296
let endDateStr = (new Date(tomorrowDateStr).getTime() / 1000).toFixed(0);

Logger.log("using end date string for Yahoo Finance: " + endDateStr);

// gist.github.com/kannaiah/53881607c9eba63099689591ad6e949e

function getYahooFinanceHistoricalData(symbol) {

    // Utilities.sleep(Math.floor(Math.random() * 5000))

    var url = "https://query1.finance.yahoo.com/v7/finance/download/"+ symbol + "?period1=" + startDateStr + "&period2=" + endDateStr + "&interval=1d&events=history&includeAdjustedClose=true";
    
    Logger.log(url);
    
    var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    if (response.getResponseCode() === 200) {
      var selectedRootFolders = DriveApp.getFoldersByName(rootDirName);

      if (selectedRootFolders.hasNext()) {
        var rootDir = selectedRootFolders.next();
        var subDirs = rootDir.getFolders();

        while (subDirs.hasNext()) {
          var folder = subDirs.next();

          if (folder.getName() === rawCsvDirName) {
            folder.createFile(response.getBlob());
          }
        }
      }
      
    } else {
      Logger.log("ERROR: " + symbol + " download request NOT OK, response code " + response.getResponseCode().toString());
      var textFile = response.getContentText();
      Logger.log("error response content: " + textFile);
    }
}

function test_getYahooFinanceHistoricalData() {
  getYahooFinanceHistoricalData("XYL");
}

/*
  # micro-batch workflow: turn large workload into small pieces to meet Google quotas
  # workflow start check
  # workflow end check
  # find unfinished individual tasks, redo, over-write existing output in output set
  # trigger check, if workflow end check feedback not true, finish current function call

*/

// Assuming top row is for description text for each data field and they are continuous.
// sheetObj: Class sheet, sheet to work on
// rowNum: Int, row number of target row
function getContinuousRowContents(sheetObj, rowNum) {
  var a1Notatoin = "A" + rowNum.toString() + ":" + rowNum.toString();
  return sheetObj
    .getRange(a1Notatoin)
    .getValues()
    .flat()
    .filter(String);
}

// Assuming top row is for description text for each data field and they are continuous.
// sheetObj: Class sheet, sheet to work on
// colName: String, col name of target col
function getContinuousColContents(sheetObj, colName) {
  var a1Notatoin = colName + "1:" + colName;

  return sheetObj
    .getRange(a1Notatoin)
    .getValues()
    .flat()
    .filter(String);
}

// get work state, which is like a db
function getWorkState(sheetObj) {
  var numOfCol = getContinuousRowContents(sheetObj, 1).length;
  var numOfRow = getContinuousColContents(sheetObj, "A").length;

  return sheetObj.getSheetValues(1, 1, numOfRow, numOfCol);
}

// wlInPriorities: Array with each watchlist as element. highest priority wl stored first, lowest wl stored last.
// state: WorkState, a 2D Array w/ top row containing descriptions for each col, each row is for an input in db
// returns 2D Array, each root element contains tickers with same priority.
function getTickersPriorities(state, wlInPriorities, tickerColNum, wlColNum) {
  return state.slice(1).reduce(function (accu, row) {
    var i = wlInPriorities.indexOf(row[wlColNum - 1]);
    if (i > -1) {
      accu[i].push(row[tickerColNum - 1]);
    } else {
      throw Error(row[wlColNum - 1].toString() + " NOT found in watchlist array");
    }
    return accu
  }, wlInPriorities.map(x => []));
}

function getWorkStateRowIndex(state, ticker, tickerColNum) {
  var r = state.findIndex(row => row[tickerColNum - 1] === ticker);

  if (r === -1) {
    throw Error(ticker + " NOT found in state, colNum: " + tickerColNum.toString());
  }

  return r;
}

// doneCheckColNum is based on array Start index 1 instead 0 to comply with Google Sheets' range feature
// tickersPriorities: Array with each watchlist as element. highest priority wl stored first, lowest wl stored last.
function getRowIndiceToWork(state, tickersPriorities, doneCheckColNum, numOfJobs, tickerColNum) {
  var r = [];

  for (i = 0; i < tickersPriorities.length; i++) {
    if (r.length === numOfJobs) {
      break;
    }

    var tickers = tickersPriorities[i];

    for (j = 0; j < tickers.length; j++) {
      var ticker = tickers[j];

      var rowIndex = getWorkStateRowIndex(state, ticker, tickerColNum);
      
      if (state[rowIndex][doneCheckColNum - 1] !== 1) {
        r.push(rowIndex);
        if (r.length === numOfJobs) {
          break;
        }
      }
    }
  }

  return r;
}

// removes all existing Google Sheets files with same name in given name list
// also removes all files w/ target formatted names containing same names, e.g., "2_AAPL"
// current format is single digit string with "_" followed by ticker
// folder: Class Folder
function removeAllSameNameSpreadsheet(folder, names) {
  var files = folder.getFilesByType("application/vnd.google-apps.spreadsheet")
  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    var nameToSearch;
    if (fileName.includes("_")) {
      nameToSearch = fileName.split("_")[1];
    } else {
      nameToSearch = fileName;
    }
    
    if (names.indexOf(nameToSearch) > -1) {
      file.setTrashed(true);
    }
  }
}

// get latest number of work setting
// work process adjusts its value based on whether lastest process exceeds time quota or not
function getNumOfJobs(sheet, a1NotationStr) {
  return sheet.getRange(a1NotationStr).getValues()[0][0];
}

function setNumOfJobs(sheet, a1NotationStr, value) {
  return sheet.getRange(a1NotationStr).setValues([[value]]);
}

// next: function (csvContent, ticker), csvContent is tabular 2D array representation of a CSV string.
function yahooFinanceHistoricalData(ticker, next) {
    // Utilities.sleep(Math.floor(Math.random() * 5000))

    var url = "https://query1.finance.yahoo.com/v7/finance/download/"+ ticker + "?period1=" + startDateStr + "&period2=" + endDateStr + "&interval=1d&events=history&includeAdjustedClose=true";
    
    var response = UrlFetchApp.fetch(url);
    if (response.getResponseCode() === 200) {
      var csvContent = Utilities.parseCsv(response.getContentText());

      next(csvContent, ticker);
    } else {
      Logger.log("ERROR: " + ticker + " download request NOT OK, response code " + response.getResponseCode().toString());
      var textFile = response.getContentText();
      Logger.log("error response content: " + textFile);
    }
}

// modified from turnCSVs2ItemsInWatchlists
// templateSpreadsheet: Class File
// toFolder: Class Folder, folder to place sheet
// data: tabular 2D array representation of a CSV string, top row is for titles
function csvData2Spreadsheet(templateSpreadsheet, toFolder, data, ticker) {
  // There is a case tciker name "TRUE" leading to state getter comes back with boolean value true instead of string "TRUE".
  // But ticker "TRUE" is in blacklist after being examined.
  var newCopy = templateSpreadsheet.makeCopy(ticker, toFolder);
  var pureData = data.slice(1);

  SpreadsheetApp
    .open(newCopy)
    .getSheetByName(sheetNameToSetCSVData)
    .getRange(csvStartRowNum, csvStartColNum, pureData.length, csvNumOfCols)
    .setValues(pureData);
  
  // adaptChartsInNewCopy NOT WORKING
  // adaptChartsInNewCopy(SpreadsheetApp.open(newCopy), l, s, z);

  createChartsInSheet(SpreadsheetApp.open(newCopy), sheetNameToSetCSVData);
  renameBasedOnLastlineDoubleActionData(newCopy);
}

function getWorkStateSpreadsheet() {
  var files = DriveApp.getFilesByType("application/vnd.google-apps.spreadsheet");
  while (files.hasNext()) {
    var file = files.next();
    if (workStateSpreadsheetName === file.getName()) {
      return file;
    }
  }
  return undefined;
}

var workStateSpreadsheetName = "2021_stock_tracker";
var workStateSheetName = "csv2charTaskState";

// assuming fields: ticker, watchlist, job started, csv downloaded, spreadsheet done
var stateFields = ["ticker", "watchlist", "job started", "csv downloaded", "spreadsheet done"];

// stateSheet: Class Sheet
// returns state, 2D Array
function resetWorkState(stateSheet) {
  var r = getEachWatchlistTickers()
   .map(x => x.map((y, j, a) => [y, a[0]]).slice(1))
   .flat()
   .map(z => z.concat([0, 0, 0]));
  
  var state = [stateFields].concat(r);

  stateSheet
    .getRange(1, 1, state.length, state[0].length)
    .setValues(state);

  return state;
}

function test_resetWorkState() {
  var stateSpreadsheet = getWorkStateSpreadsheet();
  
  if (typeof stateSpreadsheet === "undefined") {
    throw Error("Spreadsheet " + workStateSpreadsheetName + " NOT found");
  }

  var stateSheet = SpreadsheetApp
    .open(stateSpreadsheet)
    .getSheetByName(workStateSheetName);

  resetWorkState(stateSheet);
}

// stateSheet: Class Sheet
function clearWorkState(stateSheet) {
  var state = getWorkState(stateSheet);

  stateSheet
    .getRange(1, 1, state.length, state[0].length)
    .clearContents();
}

// returns Class Folder instance that stores all outputs
// returns undefined, if no folder not exists
function getOutputFolder(rootDirStr, outputDirStr) {
  var selectedRootFolders = DriveApp.getFoldersByName(rootDirStr);

  if (selectedRootFolders.hasNext()) {
    var rootDir = selectedRootFolders.next();
    var subDirs = rootDir.getFolders();

    while (subDirs.hasNext()) {
      var folder = subDirs.next();

      if (folder.getName() === outputDirStr) {
        return folder;
      }
    }
  }

  return undefined;
}

// ticker: only here to meet next callback function signiture in yahooFinanceHistoricalData
function getSampleCSV2CreateDatedFolder(csvData, ticker) {
  var outputFolder = getOutputFolder(rootDirName, outputDirName);

  outputSubDirName = csvData[csvData.length - 1][0];

  // Create new date-named folder
  outputSubDir = outputFolder.createFolder(outputSubDirName);
}

// outputFolder: Class Folder
// folder strusture: lv1: output, lv2: data-last-date-named, lv3: each watchlist
// returns lv3(watchlist) folders: Array of Class Folder, in same order as wlNamesWithPriorities
function resetAllFoldersInOutput(outputFolder) {
  // Clean all existing output
  var outputFiles = outputFolder.getFiles();

  while (outputFiles.hasNext()) {
    outputFiles.next().setTrashed(true);
  }

  var outputSubFolders = outputFolder.getFolders();

  while (outputSubFolders.hasNext()) {
    outputSubFolders.next().setTrashed(true);
  }

  // create new required folders
  yahooFinanceHistoricalData("AAPL", getSampleCSV2CreateDatedFolder);

  // Assuming at any given time, only max=1 dated folder exists, since there is clear operation before any whole work
  var datedFolder = outputFolder.getFolders().next();

  return wlNamesWithPriorities.map(x => datedFolder.createFolder(x));
}

// Assuming at any given time, only max=1 dated folder exists, since there is clear operation before any whole work
// returns lv3(watchlist) folders: Array of Class Folder, in same order as wlNamesWithPriorities
function getWatchlistFolders(outputFolder) {
  var outputSubFolders = outputFolder.getFolders();

  var r = [];

  while (outputSubFolders.hasNext()) {
    var datedFolderKids = outputSubFolders.next().getFolders();

    while (datedFolderKids.hasNext()) {
      r.push(datedFolderKids.next());
    }
  }

  if (r.length !== wlNamesWithPriorities.length) {
     throw Error("number of wlNamesWithPriorities " + wlNamesWithPriorities + " NOT equals to number of folders created");
  }

  return wlNamesWithPriorities.map(x => r.find(y => x === y.getName()))
}

var numOfJobsRangeStr = "H2";
var isWorkingRangeStr = "H3";
var jobTotalRangeStr = "H4";
var jobDoneTotalRangeStr = "H5";
var latestDisposableTriggerUniqueIdRange = "H6";

var allRecordedDisposableTriggerIdsStartCell = "G9";

function getAllRecordedDisposableTriggerIds() {
  var stateSpreadsheet = getWorkStateSpreadsheet();
  
  if (typeof stateSpreadsheet === "undefined") {
    throw Error("Spreadsheet " + workStateSpreadsheetName + " NOT found");
  }

  return SpreadsheetApp
    .open(stateSpreadsheet)
    .getSheetByName(workStateSheetName)
    .getRange(allRecordedDisposableTriggerIdsStartCell + ":" + allRecordedDisposableTriggerIdsStartCell.charAt(0))
    .flat()
    .filter(String);
}

// Assuming all ids' ranges are continuous
function appendToAllRecordedDisposableTriggerIds(triggerIdStr) {
  var stateSpreadsheet = getWorkStateSpreadsheet();
  
  if (typeof stateSpreadsheet === "undefined") {
    throw Error("Spreadsheet " + workStateSpreadsheetName + " NOT found");
  }

  var sheet = SpreadsheetApp
    .open(stateSpreadsheet)
    .getSheetByName(workStateSheetName);

  var l = sheet
    .getRange(allRecordedDisposableTriggerIdsStartCell + ":" + allRecordedDisposableTriggerIdsStartCell.charAt(0))
    .getValues()
    .flat()
    .filter(String)
    .length;

  sheet
    .getRange(allRecordedDisposableTriggerIdsStartCell.charAt(0) + (parseInt(allRecordedDisposableTriggerIdsStartCell.substring(1)) + l).toString())
    .setValue(triggerIdStr);

  Logger.log("trigger id appended: " + triggerIdStr);
}

// existingTriggers: Array of Class Triger, each element is a Trigger standing for a trigger to remove, if matched
// any id without existing trigger match will be cleared as well
function clearNRemoveMatchedNTriggersInAllRecordedDisposableTriggerIds(existingTriggers) {
  var stateSpreadsheet = getWorkStateSpreadsheet();
  
  if (typeof stateSpreadsheet === "undefined") {
    throw Error("Spreadsheet " + workStateSpreadsheetName + " NOT found");
  }

  var stateSheet = SpreadsheetApp
    .open(stateSpreadsheet)
    .getSheetByName(workStateSheetName);

  // keep latest disposable trigger, if set
  var idToKeepForNow = stateSheet.getRange(latestDisposableTriggerUniqueIdRange).getValue();

  var dataRange = stateSheet.getRange(allRecordedDisposableTriggerIdsStartCell + ":" + allRecordedDisposableTriggerIdsStartCell.charAt(0));

  var tIds = existingTriggers.map(t => t.getUniqueId());

  var values = dataRange
    .getValues()
    .flat()
    .filter(String)
    .map(x => [x]);

  if (values.length > 0) {
    var newValues = values.map(function (x) {
      var storedId = x[0];

      Logger.log("trigger id stored: " + storedId);

      if (storedId !== "") {
        // remove matched
        var i = tIds.indexOf(storedId);
        
        // clear unmatched id, if not the latest disposable trigger id
        if (idToKeepForNow !== storedId) {
          if (tIds.indexOf(storedId) > -1) {
            Logger.log("DELETED trigger id stored W/ existing trigger: " + storedId);

            ScriptApp.deleteTrigger(existingTriggers[i]);
          }
          Logger.log("DELETED trigger id stored W/O existing trigger: " + storedId);
          return [""];
        }
      }

      return x;
    });

    var lastRowNum = parseInt(allRecordedDisposableTriggerIdsStartCell.substring(1)) + newValues.length - 1;

    var newRange = stateSheet.getRange(allRecordedDisposableTriggerIdsStartCell + ":" + allRecordedDisposableTriggerIdsStartCell.charAt(0) + lastRowNum.toString());

    newRange
      .setValues(newValues)
      .sort(newRange.getColumn());
  }
}

function exploreYahooF2SpreadsheetJobs() {
  var stateSpreadsheet = getWorkStateSpreadsheet();
  
  if (typeof stateSpreadsheet === "undefined") {
    throw Error("Spreadsheet " + workStateSpreadsheetName + " NOT found");
  }

  var stateSheet = SpreadsheetApp
    .open(stateSpreadsheet)
    .getSheetByName(workStateSheetName);

  // process starts, mark "is working" status to 1
  stateSheet.getRange(isWorkingRangeStr).setValue(1);
  SpreadsheetApp.flush();

  var state;
  var wlFolders;

  var outputFolder = getOutputFolder(rootDirName, outputDirName);

  if (stateSheet.getRange("A1").getValue() === "") {
    state = resetWorkState(stateSheet);

    // fresh start, reset folders as well
    wlFolders = resetAllFoldersInOutput(outputFolder)
  } else {
    state = getWorkState(stateSheet);
    wlFolders = getWatchlistFolders(outputFolder);
  }
  
  var fields = state[0];

  var numOfJobs = getNumOfJobs(stateSheet, numOfJobsRangeStr);

  var priorityList = getTickersPriorities(state, wlNamesWithPriorities, fields.indexOf("ticker") + 1, fields.indexOf("watchlist") + 1);

  var tickerIndiceInJobs = getRowIndiceToWork(state, priorityList, fields.indexOf("spreadsheet done") + 1, numOfJobs, fields.indexOf("ticker") + 1);

  var tickersInJobs = tickerIndiceInJobs.map(x => state[x][fields.indexOf("ticker")]);

  // mark job started
  var indexOfJobStartedCol = fields.indexOf("job started");
  tickerIndiceInJobs.forEach(x => stateSheet.getRange(x + 1, indexOfJobStartedCol + 1, 1, 1).setValue(1));
  SpreadsheetApp.flush();

  // Get template
  var spreadsheetFiles = outputFolder
    .getParents()
    .next()
    .getFilesByType("application/vnd.google-apps.spreadsheet");

  while (spreadsheetFiles.hasNext()) {
    var spreadsheetFile = spreadsheetFiles.next();
    if (spreadsheetFile.getName() === boilerplateSpreadsheetName) {
      boilerplateSpreadsheet = spreadsheetFile;
      break;
    }
  }

  i = 0;

  var indexOfCSVDownloadedCol = fields.indexOf("csv downloaded");
  var indexOfDoneCol = fields.indexOf("spreadsheet done");

  function workThenMarkDone(data, ticker) {
    var indexOfTickerRow = getWorkStateRowIndex(state, ticker, fields.indexOf("ticker") + 1);

    // mark CSV downloaded
    stateSheet.getRange(indexOfTickerRow + 1, indexOfCSVDownloadedCol + 1, 1, 1).setValue(1);
    SpreadsheetApp.flush();

    var wlName = state[indexOfTickerRow][fields.indexOf("watchlist")];
    var wlFolder = wlFolders[wlNamesWithPriorities.indexOf(wlName)];

    removeAllSameNameSpreadsheet(wlFolder, [ticker]);
    csvData2Spreadsheet(boilerplateSpreadsheet, wlFolder, data, ticker);

    // mark job done
    stateSheet.getRange(indexOfTickerRow + 1, indexOfDoneCol + 1, 1, 1).setValue(1);
    SpreadsheetApp.flush()

    i++;
    
    var jobTotal = stateSheet.getRange(jobTotalRangeStr).getValue();

    if (i === numOfJobs || (jobTotal === stateSheet.getRange(jobDoneTotalRangeStr).getValue() && jobTotal > 0)) {
      // process done, mark "is working" status to 0
      stateSheet.getRange(isWorkingRangeStr).setValue(0);
      SpreadsheetApp.flush();

      // set timer to trigger next run
      var t = ScriptApp
        .newTrigger("checkToDecideExploringYahooF2SpreadsheetJobs")
        .timeBased()
        .after(100)
        .create();

      Logger.log("new trigger id: " + t.getUniqueId());        

      // overwrite latestDisposableTriggerUniqueIdRange to keep it from removal, then append to all
      stateSheet.getRange(latestDisposableTriggerUniqueIdRange).setValue(t.getUniqueId());
      appendToAllRecordedDisposableTriggerIds(t.getUniqueId());
      SpreadsheetApp.flush();
    }
  }

  tickersInJobs.forEach(function (ticker) {
    yahooFinanceHistoricalData(ticker, workThenMarkDone);
  });
}

function test_exploreYahooF2SpreadsheetJobs() {
  exploreYahooF2SpreadsheetJobs();
}

function checkToDecideExploringYahooF2SpreadsheetJobs() {
  try {
    var stateSpreadsheet = getWorkStateSpreadsheet();
  
    if (typeof stateSpreadsheet === "undefined") {
      throw Error("Spreadsheet " + workStateSpreadsheetName + " NOT found");
    }

    var stateSheet = SpreadsheetApp
      .open(stateSpreadsheet)
      .getSheetByName(workStateSheetName);

    var jobTotal = stateSheet.getRange(jobTotalRangeStr).getValue();

    // check if all jobs are done each time
    if (jobTotal === stateSheet.getRange(jobDoneTotalRangeStr).getValue() && jobTotal > 0) {
      Logger.log("All jobs already done last time");
    } else {
      // process starts, if "is working" status is 0
      if (stateSheet.getRange(isWorkingRangeStr).getValue() === 0) {
        clearNRemoveMatchedNTriggersInAllRecordedDisposableTriggerIds(ScriptApp.getUserTriggers(SpreadsheetApp.open(stateSpreadsheet)));
        exploreYahooF2SpreadsheetJobs();
      } else {
        Logger.log("NOT to run exploreYahooF2SpreadsheetJobs, 'is working' status is 1");
      }
    }
  } catch (err) {
    Logger.log("ERROR: " + err.message);
    Logger.log(err.stack);
    Logger.log("Will reset 'is working' back to 0 then retry...");

    var stateSpreadsheet = getWorkStateSpreadsheet();

    if (typeof stateSpreadsheet === "undefined") {
      throw Error("Spreadsheet " + workStateSpreadsheetName + " NOT found");
    }

    var stateSheet = SpreadsheetApp
      .open(stateSpreadsheet)
      .getSheetByName(workStateSheetName);

    // process NOT done but terminated, mark "is working" status back to 0 to clear barrier for next execution
    stateSheet.getRange(isWorkingRangeStr).setValue(0);
    SpreadsheetApp.flush();

    Utilities.sleep(800);
    checkToDecideExploringYahooF2SpreadsheetJobs();
  }
}

function test_checkToDecideExploringYahooF2SpreadsheetJobs() {
  checkToDecideExploringYahooF2SpreadsheetJobs();
}

// 2021-05-31 just did NOT work. Further debug needed.
var marketHoliday2021Strs = ["2021-05-31", "2021-07-05", "2021-09-06", "2021-11-25", "2021-12-24"];

// date String format: "2020-05-04", "2020-05-19"
// only triggered on every weekday, so no need to check weekends, just market holidays
function getSampleCSV2CheckLastTradingDate(csvData, ticker) {
  var lastTradingDateStr = csvData[csvData.length - 1][0];

  // marketHoliday2021Strs.indexOf(lastTradingDateStr) === -1 NOT correct. REWRITE NEEDED.
  if (marketHoliday2021Strs.indexOf(lastTradingDateStr) === -1) {
    // clear previous state and reset working status to make reset possible
    var stateSpreadsheet = getWorkStateSpreadsheet();
  
    if (typeof stateSpreadsheet === "undefined") {
      throw Error("Spreadsheet " + workStateSpreadsheetName + " NOT found");
    }

    SpreadsheetApp
      .open(stateSpreadsheet)
      .getSheetByName(workStateSheetName)
      .getRange(1, 1, 1001, stateFields.length)
      .clearContent();

    SpreadsheetApp
      .open(stateSpreadsheet)
      .getSheetByName(workStateSheetName)
      .getRange(isWorkingRangeStr)
      .setValue(0);
    
    SpreadsheetApp.flush();

    checkToDecideExploringYahooF2SpreadsheetJobs();
  }
}

function startJobsAfterMarketClose() {
  yahooFinanceHistoricalData("AAPL", getSampleCSV2CheckLastTradingDate);
}

function test_startJobsAfterMarketClose() {
  startJobsAfterMarketClose();
}

// this is used for a weekday trigger
function dailyStartJobsAfterMarketClose() {
  var dayIndex = (new Date()).getDay();

  if (dayIndex !== 0 && dayIndex !== 6) {
    startJobsAfterMarketClose();
  }
}

// return Array, 1st element is Array of tickers jobs have but not in wls, 2nd is vice versa
function findWlVsJobsDoneTickerDifferences() {
  var stateSpreadsheet = getWorkStateSpreadsheet();

  if (typeof stateSpreadsheet === "undefined") {
    throw Error("Spreadsheet " + workStateSpreadsheetName + " NOT found");
  }

  var stateSheet = SpreadsheetApp
    .open(stateSpreadsheet)
    .getSheetByName(workStateSheetName);

  var allTickers = stateSheet
    .getRange("A2:A1001")
    .getValues()
    .flat()
    .filter(String);
  
  var wlSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(watchlistsSheet);

  var aRange = wlSheet.getRange(watchlistsSheetContentStartRowNum, 2, 1001, numOfWls);

  var tickersInWl = aRange
    .getValues()
    .flat()
    .filter(String);

  return [allTickers.filter(x => tickersInWl.indexOf(x) === -1), tickersInWl.filter(x => allTickers.indexOf(x) === -1)];
}

function test_findWlVsJobsDoneTickerDifferences() {
  var r = findWlVsJobsDoneTickerDifferences();
  Logger.log(r[0].length.toString() + " ticker(s) in jobs but not in watchlists: " + r[0]);
  Logger.log(r[1].length.toString() + " ticker(s) in watchlists but not in jobs: " + r[1]);
}

var tickerRecordSheetName = "tickersWTags";

// return Array, 1st element is Array of tickers w tags have but not in wls, 2nd is vice versa
function findWlVsRecordedTickerDifferences() {
  var stateSpreadsheet = getWorkStateSpreadsheet();

  if (typeof stateSpreadsheet === "undefined") {
    throw Error("Spreadsheet " + workStateSpreadsheetName + " NOT found");
  }

  var stateSheet = SpreadsheetApp
    .open(stateSpreadsheet)
    .getSheetByName(tickerRecordSheetName);

  var allTickers = stateSheet
    .getRange("E1:1")
    .getValues()[0]
    .filter(String);
  
  var wlSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(watchlistsSheet);

  var aRange = wlSheet.getRange(watchlistsSheetContentStartRowNum, 4, 1001, 12);

  var tickersInWl = aRange
    .getValues()
    .flat()
    .filter(String);

  return [allTickers.filter(x => tickersInWl.indexOf(x) === -1), tickersInWl.filter(x => allTickers.indexOf(x) === -1)];
}

function test_findWlVsRecordedTickerDifferences() {
  var r = findWlVsRecordedTickerDifferences();
  Logger.log(r[0].length.toString() + " ticker(s) in records but not in watchlists: " + r[0]);
  Logger.log(r[1].length.toString() + " ticker(s) in watchlists but not in records: " + r[1]);
}

// Assuming watchList contains exactly the same tickers as in records.
// Based on current watchlist, excluding tickers in "未审", get tags for each ticker then check against existing ticker tag records,
// if there is any change, append in the same column with date.
// Each record starts at row 1, followed by date info of first entry at row 2, then followed by tag info in calculated number of first entry at row 2.
// Each entry occupies two rows, as mentioned above.
// In case of two same date entries, since new entries are always appended in the column, each entry's order indicates which one is inputed earlier.
// If there is duplicate or incorrect inputs, it has to be solved manually. It's out of this function's scope.
function checkToUpdateTickerTags() {
  var today = new Date();
  var starColName = "E";
  var endColName = "IH";

  var stateRangeMaxLength = 4 + SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(watchlistsSheet)
    .getRange("B5:O")
    .getValues()
    .map(x => x.filter(String))
    .filter(y => y.length > 0)
    .length;

  var stateRange = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(watchlistsSheet)
    .getRange("B5:O" + stateRangeMaxLength);

  var stateRangesCells = stateRange
    .getValues()
    .map(function (r, i) {
      return r.map(function (c, j) {
        return stateRange.getCell(i + 1, j +1);
      });
    });

  var stateRowNum = stateRange.getNumRows();
  var stateColNum = stateRange.getNumColumns();
  var titles = stateRange.getValues()[0];

  var tickersInRecords = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(tickerRecordSheetName)
    .getRange(starColName + "1:" + endColName + "1")
    .getValues()
    .flat()
    .filter(String);

  var recordStartColNum = 5;

  var colorTagColors = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(tickerRecordSheetName)
    .getRange("A2:A4")
    .getBackgrounds()
    .flat();

  var colorTagNames = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(tickerRecordSheetName)
    .getRange("A2:A4")
    .getValues()
    .flat();

  tickersInRecords.forEach(function (ticker, i) {
    // records will not reach 500, so 1000 is good and in scope.
    var rangeToCheck = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(tickerRecordSheetName)
      .getRange(1, recordStartColNum + i, 1000, 1);
    var l = rangeToCheck
      .getValues()
      .flat()
      .filter(String);
    var nonEmptyLength = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(tickerRecordSheetName)
      .getRange(l.length + 1, recordStartColNum + i, 1000 - l.length, 1)
      .getValues()
      .flat()
      .filter(String)
      .length;
    if (nonEmptyLength > 0) {
      // There are contents in remaining cells that shoud be empty.
      throw ("Empty cell in records for ticker in tickerRecordSheetName: " + ticker);
    }
    var lastRecordContent = rangeToCheck.getValues()[l.length - 1][0];
    var appendAnyway = false;
    if (l.length === 1) {
      appendAnyway = true;
    }
    checkTicker:
    for (var j = 1; j < stateRowNum; j++) {
      for (var z = 0; z < stateColNum; z++) {
        var cellObj = stateRangesCells[j][z];
        if (cellObj.getValue() === ticker) {
          var title = titles[z];
          var colorStr = cellObj.getBackground();
          var c = colorTagColors.indexOf(colorStr);
          
          var content;
          if (c > -1) {
            var colorTagName = colorTagNames[c];
            content = colorTagName + " + " + title;
          } else {
            content = title;
          }
          var isNew = false;
          if (appendAnyway === false) {
            if (lastRecordContent !== content) {
              isNew = true;
            }
          }
          if (appendAnyway || isNew) {
            SpreadsheetApp
            .getActiveSpreadsheet()
            .getSheetByName(tickerRecordSheetName)
            .getRange(l.length + 1, recordStartColNum + i, 2, 1)
            .setValues([[today], [content]]);
          }
          break checkTicker;
        }
      }
    }
  });
}

function test_checkToUpdateTickerTags() {
  checkToUpdateTickerTags();
}
