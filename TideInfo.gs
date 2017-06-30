//
// TideInfo.gs
// Created on 2017-06-30 10:00
// Created by sustny(http://sustny.me/)
//

getInfo = function(DATE, SPOT) {
  //1. 日付と行く場所を指定する(行く場所は特定の観測地点である必要はなく、純粋に自身の目的地を入力すればよい)
  var YEAR = parseInt(DATE.substr(0,4));
  var MONTH = parseInt(DATE.substr(4,2));
  var DAY = parseInt(DATE.substr(6,2));
  
  //1-1. 頭が0の2桁(ex:08)をparseIntするとNaNになるのでその処理
  if( parseInt(DATE.substr(4,1)) == 0 ) {
    MONTH = parseInt(DATE.substr(5,1));
  }
  if( parseInt(DATE.substr(6,1)) == 0 ) {
    DAY = parseInt(DATE.substr(7,1));
  }
  
  //2. 入力されたSPOTを緯度・経度に分解する
  for(var i=0;i<SPOT.length;i++) {
    if(SPOT.substr(i,1) == ',') {
      var lat = parseFloat(SPOT.substr(0,i));
      var lon = parseFloat(SPOT.substr(i+1));
    }
  }
  
  //3. 気象庁が公表している潮位表の掲載地点一覧を呼び出し、行く場所から最も近い観測地点を探す
  var FILE = SpreadsheetApp.getActiveSpreadsheet();
  var SHEET = FILE.getSheetByName('Data');
  var CELL = SHEET.getDataRange().getValues();
  var DIST = [];
  for(i=1;i<CELL.length;i++) {
    //3-1. 経緯度差を三平方の定理から求めて、
    var Dlat = parseFloat(CELL[i][3].substr(0,2)) + (parseFloat(CELL[i][3].substr(3)) / 60) - lat;
    var Dlon = parseFloat(CELL[i][4].substr(0,3)) + (parseFloat(CELL[i][4].substr(4)) / 60) - lon;
    DIST[i-1] = Math.sqrt( (Dlat*Dlat) + (Dlon*Dlon) );
    //3-2. 最も近いところを探す
    if(i == 1) {
      var min = DIST[i-1];
      var j = i;
    } else {
      if(min>DIST[i-1]) {
        min = DIST[i-1];
        j = i;
      }
    }
  }
  
  //4. 最も近い観測地点の情報から当該日付の情報(1行)を抜き出す
  var URL = UrlFetchApp.fetch('http://www.data.jma.go.jp/kaiyou/data/db/tide/suisan/txt/' + YEAR +'/' + CELL[j][1] + '.txt');
  var TXT = URL.getContentText();
  
  if(YEAR % 4 != 0) {
    var arrMonth = [0,31,59,90,120,151,181,212,243,273,304,334]; //各月の開始行メモ
    var Row = TXT.substr(( (137*arrMonth[MONTH-1]) + (137*(DAY-1)) ),136);
  } else {
    var arrMonth = [0,31,60,91,121,152,182,213,244,274,305,335]; //各月の開始行メモ
    var Row = TXT.substr(( (137*arrMonth[MONTH-1]) + (137*(DAY-1)) ),136);
  }
  
  //5. 抜き出した情報を分割して配列に格納する
  //Info = [地名、満潮時刻1、満潮潮位1、満潮時刻2、満潮潮位2、満潮時刻3、満潮潮位3、満潮時刻4、満潮潮位4、干潮時刻1、干潮潮位1、干潮時刻2、干潮潮位2、干潮時刻3、干潮潮位3、干潮時刻4、干潮潮位4]
  /*
  var Info = [CELL[j][2],
              '' + parseInt(Row.substr(80,2)) + ':' + parseInt(Row.substr(82,2)), parseInt(Row.substr(84,3)),
              '' + parseInt(Row.substr(87,2)) + ':' + parseInt(Row.substr(89,2)), parseInt(Row.substr(91,3)),
              '' + parseInt(Row.substr(94,2)) + ':' + parseInt(Row.substr(96,2)), parseInt(Row.substr(98,3)),
              '' + parseInt(Row.substr(101,2)) + ':' + parseInt(Row.substr(103,2)), parseInt(Row.substr(105,3)),
              '' + parseInt(Row.substr(108,2)) + ':' + parseInt(Row.substr(110,2)), parseInt(Row.substr(112,3)),
              '' + parseInt(Row.substr(115,2)) + ':' + parseInt(Row.substr(117,2)), parseInt(Row.substr(119,3)),
              '' + parseInt(Row.substr(122,2)) + ':' + parseInt(Row.substr(124,2)), parseInt(Row.substr(126,3)),
              '' + parseInt(Row.substr(129,2)) + ':' + parseInt(Row.substr(131,2)), parseInt(Row.substr(133,3))];
              */
  var Info = [];
  Info[0] = CELL[j][2];
  
  j=0;
  for(i=1;i<=16;i+=2) {
    if(parseInt(Row.substr(82+(j*7),2)) < 10) {
      Info[i] = '' + parseInt(Row.substr(80+(j*7),2)) + ':0' + parseInt(Row.substr(82+(j*7),2));
    } else {
      Info[i] = '' + parseInt(Row.substr(80+(j*7),2)) + ':' + parseInt(Row.substr(82+(j*7),2));
    }
    Info[i+1] = parseInt(Row.substr(84+(j*7),3))
    j++;
  }
  
  //6. 情報が格納された配列を返す
  return Info;
}

function TideInfo_Main() {
  var FILE = SpreadsheetApp.getActiveSpreadsheet();
  var SHEET = FILE.getSheetByName('Info');
  var CELL = SHEET.getDataRange().getValues();
  var date = CELL[0][2];
  if(date == '') { //入力がないなら今日
    date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  }
  var spot = CELL[0][6];
  if(spot == '') { //入力がないなら東京
    spot = '35.640342, 139.855059';
  }
  var info = getInfo(date, spot);
  
  SHEET.getRange(2,1).setValue('');
  SHEET.getRange(5,1,1,16).setValue('');
  
  SHEET.getRange(2,1).setValue(info[0]);
  for(i=1;i<=16;i+=2) {
    if(info[i+1] != 999) {
      SHEET.getRange(5,i).setValue(info[i]);
      SHEET.getRange(5,i+1).setValue(info[i+1]);
    }
  }
}
