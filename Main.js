const ID = '<スプレッドシートID>';
const SHEET_NAME = '<シート名>';
const DATE_CELL_START = 9;
const ROW_IMAGE = 1;
const ROW_URL = 2;
const ROW_CAR_NAME = 3;
const ROW_OVERVIEW = 4;
const ROW_MODEL_YEAR = 5;
const ROW_DISTANCE = 6;
const ROW_BODY_COLOR = 7;
const ROW_DRIVE_SYSTEM = 8;
const ROW_NAME_OVERVIEW = 9;
const WK_START_COL = 2;
var sheet;

var wkColumn;  // 作業用列番号を保持する

function rootFunction() {
  getMySheet();
  setFormat();
  let lastCol = sheet.getRange(ROW_URL, 1).getNextDataCell(SpreadsheetApp.Direction.NEXT).getColumn();
  for (wkColumn=WK_START_COL; wkColumn<=lastCol; wkColumn++) {
    //スクレイピングしたいWebページのURLを変数で定義する
    let url = sheet.getRange(ROW_URL, wkColumn).getValue();
    //URLに対しフェッチを行ってHTMLデータを取得する
    let html = UrlFetchApp.fetch(url).getContentText("EUC-JP");

    writeImage(html);     // 画像
    writeCarName(html);   // 車種
    writeOverview(html);  // 概要
    writePrice(html);     // 価格
    writeModelYear(html); // 年式
    writeDistance(html);  // 走行距離
    writeBodyColor(html); // 車体色
    writeDriveSystem(html); // 駆動方式
    writeDataTitle();     // 車種と概要を連結
  }
}

function getToday() {
  let today = new Date();
  return Utilities.formatDate(today, "JST", "YYYY/MM/dd");
}

/**
 * 書き出すスプレッドシートを取得
 */
function getMySheet() {
  var spreadsheet = SpreadsheetApp.openById(ID);
  sheet = spreadsheet.getSheetByName(SHEET_NAME);
}

/**
 * 見出し等
 */
function setFormat() {
  sheet.getRange(ROW_URL, 1).setValue('URL');
  sheet.getRange(ROW_CAR_NAME, 1).setValue('車種');
  sheet.getRange(ROW_OVERVIEW, 1).setValue('概要');
  sheet.getRange(ROW_MODEL_YEAR, 1).setValue('年式');
  sheet.getRange(ROW_DISTANCE, 1).setValue('走行距離');
  sheet.getRange(ROW_BODY_COLOR, 1).setValue('車体色');
  sheet.getRange(ROW_DRIVE_SYSTEM, 1).setValue('駆動方式');
  sheet.getRange(DATE_CELL_START, 1).setValue('日付');
  sheet.getRange(DATE_CELL_START + ':' + DATE_CELL_START).setBackground("#c9daf8");
}

/**
 * 画像を書き出す
 */
function writeImage(html) {
  let imageURLs = html.match(/<div class="item image"><img src="[^""]+"/g);
  imageURL = imageURLs[0].replace(/<div class="item image"><img src="/g,"").replace(/"/g,"");
  console.log(imageURL);
  sheet.getRange(ROW_IMAGE, wkColumn).setValue('=IMAGE("' + imageURL + '")');
}

/**
 * 本体価格・乗り出し価格を書き出す
 */
function writePrice(html) {
  let priceList = html.match(/<td><span class="num".*>/g);

  let price = "";
  for (i=0; i<priceList.length; i++) {
    tmp = priceList[i].length > 0 ? tagCutter(priceList[i]) : "-";
    price = price + (i==1 ? '/' : '') + tmp;
  }
  sheet.getRange(getTodaysRow(), wkColumn).setValue(price.replace('万円',''));
  return;
}

/**
 * 車種を書き出す
 */
function writeCarName(html) {
  let carNames = Parser.data(html).from('<p class="tit">').to("</p>").iterate();
  let carName = tagCutter2(carNames[0]);
  sheet.getRange(ROW_CAR_NAME, wkColumn).setValue(carName);
}

/**
 * 概要を書き出す
 */
function writeOverview(html) {
  let overview = tagCutter(html.match(/<p.*class="hdBlockTop_txt".*>/g,"")[0]);
  sheet.getRange(ROW_OVERVIEW, wkColumn).setValue(overview);
  return;
}

/**
 * 年式を書き出す
 */
function writeModelYear(html) {
  let modelYears = Parser.data(html).from('年式</th>').to('</td>').iterate();
  let modelYear = tagCutter(modelYears[0]);
  sheet.getRange(ROW_MODEL_YEAR, wkColumn).setValue(modelYear);
}

/**
 * 走行距離を書き出す
 */
function writeDistance(html) {
  let distances = Parser.data(html).from('走行</th>').to('</td>').iterate();
  let distance = tagCutter(distances[0]);
  sheet.getRange(ROW_DISTANCE, wkColumn).setValue(distance);
}

/**
 * 車体色を書き出す
 */
function writeBodyColor(html) {
  let bodyColors = Parser.data(html).from('車体色</th>').to('</td>').iterate();
  let bodyColor = tagCutter(bodyColors[0]);
  sheet.getRange(ROW_BODY_COLOR, wkColumn).setValue(bodyColor);
}

/**
 * 駆動方式を書き出す
 */
function writeDriveSystem(html) {
  let driveSystems = Parser.data(html).from('駆動方式</th>').to('</td>').iterate();
  let driveSystem = tagCutter(driveSystems[0]);
  sheet.getRange(ROW_DRIVE_SYSTEM, wkColumn).setValue(driveSystem);
}

/**
 * 車種と概要を合体したものを書き出す
 */
function writeDataTitle() {
  let title = "";
  title = title + sheet.getRange(ROW_CAR_NAME, wkColumn).getValue();
  title = title + sheet.getRange(ROW_OVERVIEW, wkColumn).getValue();
  sheet.getRange(ROW_NAME_OVERVIEW, wkColumn).setValue(title);
}

/**
 * 不要なタグを削って返す
 */
function tagCutter(arg) {
  return arg.replace(/<[^>]+>/g,"").trim();
}
/**
 * 不要なタグを削った返す part2
 */
function tagCutter2(arg) {
  return arg.replace(/<.*>/g,"").replace(/\s/g,"");
}

/**
 * 今日の行を取得する
 */
function getTodaysRow() {
  let today = getToday();
  let range = getDateRange();
  let dateRange = range.getValues();
  for (i=1; i<dateRange.length; i++) {
    if (Utilities.formatDate(new Date(dateRange[i]), "JST", "YYYY/MM/dd") == today) {
      return DATE_CELL_START + i;
    }
  }
  sheet.getRange(DATE_CELL_START + dateRange.length, 1).setValue(today);
  return DATE_CELL_START + dateRange.length;
}

/**
 * 日付エリアを取得
 */
function getDateRange() {
  let lastRow = DATE_CELL_START;
  if (sheet.getRange(DATE_CELL_START + 1,1).getValue() != '') {
    lastRow = sheet.getRange(DATE_CELL_START, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  }
  let range = sheet.getRange(DATE_CELL_START,1,lastRow - (DATE_CELL_START - 1));
  return range;
}