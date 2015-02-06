// GA APIと連携
function init() {
  try {
    // データの取得する
    var GAProfileData = getGAProfile();

    // スプレッドシートにデータを挿入
    outputToSpreadsheet(GAProfileData);
  } catch(error) {
    Browser.msgBox(error.message);
  }
}

// 期間の日付を取得
var getLastNdays = function(nDaysAgo) {
  var today = new Date();
  var before = new Date();
  before.setDate(today.getDate() - nDaysAgo);
  return Utilities.formatDate(before, 'GMT', 'yyyy-MM-dd');
}
// 2 weeks (a fortnight) ago.
var startDate = getLastNdays(31);
// Today.
var endDate = getLastNdays(0);

// GAのアカウント情報（Eメールやプロパティ・ビューなど）
// Management APIは、許可されたユーザのためのGoogle Analyticsの設定データへのアクセスを提供する
// Accountsは、すべてのアカウントを一覧表示することができる
var accounts = Analytics.Management.Accounts.list();
// 5番目のアカウントを指定
var firstAccountId = accounts.getItems()[5].getId();

// Webproperties.listは、プロパティの一覧を取得する
var webProperties = Analytics.Management.Webproperties.list(firstAccountId);

// GAのプロパティIDを取得（UA-xxxxxxxx-x）
var firstWebPropertyId = webProperties.getItems()[1].getId();
// アカウントでトラッキングIDが一致するデータの情報
var profiles = Analytics.Management.Profiles.list(firstAccountId, firstWebPropertyId);


// セグメントのデータを格納
var segmentQuery = [
  {
    segment: '全体',
    query: {}
  },
  {
    segment: 'Referral',
    query: {
      'dimensions': 'ga:medium',
      'filters': 'ga:medium==referral'
    }
  }
];


// GAのビューデータを取得する
var getGAProfile = function() {
  if (accounts.getItems()) {
    if (webProperties.getItems()) {
      if (profiles.getItems()) {
        var GAProfileData = profiles.getItems()[0];
        return GAProfileData;
      } else { // 指定のプロパティのエラー処理
        throw new Error('No views (profiles) found.');
      }
    } else { // プロパティ全体がなかったときのエラー処理
      throw new Error('No webproperties found.');
    }
  } else { // アカウントがなかったときのエラー処理
    throw new Error('No accounts found.');
  }
}

// スプレッドシートにデータを挿入
var outputToSpreadsheet = function(GAProfileData) {

  // tableのID
  var profileId = GAProfileData.getId();
  
  // IDに文字列「ga:」をつける
  var tableId = 'ga:' + profileId;

  // セグメントされたデータの出力・シートの生成
  segmentQuery.map(function(element){
    var segment = element.segment;
    var query = element.query;
    _getSegmentData(segment, query, tableId, startDate, endDate);
  });
}

// セグメントされたデータの取得する
var _getSegmentData = function(segment, query, tableId, startDate, endDate) {
    // GAデータの取得
    var results = _getReportDataForProfile(query, tableId);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(segment + '：' + startDate + '-' + endDate);

    // 見出しの出力
    var headerNames = [];
    var resHeader = results.columnHeaders;
    var square = resHeader.map(function(element, index, array) {
      var header = results.getColumnHeaders()[index];
      headerNames.push(header.getName());
    });
    sheet.getRange(1, 1, 1, headerNames.length)
      .setValues([headerNames])
      .setBackgroundColor('#eeeeee')
      .setFontWeight('bold');

    // データの出力
    sheet.getRange(2, 1, results.getRows().length, headerNames.length)
      .setValues(results.getRows());
};

// セグメントしたGAデータの取得
// 引数の「GAProfileData」はGAのビュー
var _getReportDataForProfile = function(query, tableId) {
  // Make a request to the API.
  var results = Analytics.Data.Ga.get(
      tableId,                    // Table id (format ga:xxxxxx).
      startDate,                  // Start-date (format yyyy-MM-dd).
      endDate,                    // End-date (format yyyy-MM-dd).
      'ga:sessions,ga:pageviews,ga:users,ga:bounces,ga:newUsers,ga:avgSessionDuration',
      query);

  if (results.getRows()) {
    return results;
  } else {
    throw new Error('No views (profiles) found');
  }
}
