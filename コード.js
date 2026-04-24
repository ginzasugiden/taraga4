/**
 * ====================================================
 * TARA SISTER GA4 自動レポート
 * ====================================================
 * 
 * 【設定手順】
 * 1. レポート用のスプレッドシートを作成
 * 2. 拡張機能 → Apps Script を開く
 * 3. このコードを貼り付け
 * 4. 左メニュー「サービス」→「+」→「Google Analytics Data API」を追加
 * 5. 一度 exportAllReports() を手動実行して認証を許可
 * 6. トリガー設定（後述）
 * 
 * 【トリガー設定】
 * 左メニュー「トリガー」→「トリガーを追加」
 * - 実行する関数: exportAllReports
 * - イベントのソース: 時間主導型
 * - 時間ベースのトリガーのタイプ: 日付ベースのタイマー
 * - 時刻: 午前6時〜7時（推奨）
 */

// ============================================
// 設定
// ============================================
const CONFIG = {
  GA4_PROPERTY_ID: '506183623',  // TARA SISTER GA4
  DATE_RANGE_DAYS: 30,           // 取得する日数（過去30日）
  
  // シート名
  SHEETS: {
    DAILY_SUMMARY: 'GA4_日別サマリ',
    PAGE_REPORT: 'GA4_ページ別',
    CHANNEL_REPORT: 'GA4_集客チャネル別',
    CONVERSION_REPORT: 'GA4_コンバージョン別'
  },
  
  // コンバージョンイベント名（GA4で設定済みのもの）
  CONVERSION_EVENTS: [
    'purchase'
  ],

  SPREADSHEET_ID: '1gDVD9q8ZujkU-YLVnxA5_0fSFSLWv3hXVAXTnqAnpfM'
};


// ============================================
// メイン実行関数（トリガーで呼び出す）
// ============================================
/**
 * 全レポートを一括実行
 * トリガーにはこの関数を設定してください
 */
function exportAllReports() {
  console.log('=== GA4レポート取得開始 ===');
  
  try {
    exportDailySummary();
    console.log('✓ 日別サマリ完了');
    
    exportPageReport();
    console.log('✓ ページ別完了');
    
    exportChannelReport();
    console.log('✓ 集客チャネル別完了');
    
    exportConversionReport();
    console.log('✓ コンバージョン別完了');
    
    console.log('=== 全レポート取得完了 ===');
    
  } catch (error) {
    console.error('エラー発生:', error);
    // エラー通知メールを送る場合はここに追加
  }
}


// ============================================
// 1. 日別サマリレポート
// ============================================
/**
 * 日別の基本指標を取得
 * - 日付、ユーザー数、セッション数、PV、エンゲージメント率
 */
function exportDailySummary() {
  const sheetName = CONFIG.SHEETS.DAILY_SUMMARY;
  const sheet = getOrCreateSheet(sheetName);
  
  const header = [
    '取得日時',
    '日付',
    'ユーザー数',
    '新規ユーザー数',
    'セッション数',
    'ページビュー数',
    'エンゲージメント率(%)',
    '平均セッション時間(秒)',
    'コンバージョン数'
  ];
  
  const property = 'properties/' + CONFIG.GA4_PROPERTY_ID;
  
  const request = {
    dateRanges: [
      { startDate: CONFIG.DATE_RANGE_DAYS + 'daysAgo', endDate: 'yesterday' }
    ],
    dimensions: [
      { name: 'date' }
    ],
    metrics: [
      { name: 'totalUsers' },
      { name: 'newUsers' },
      { name: 'sessions' },
      { name: 'screenPageViews' },
      { name: 'engagementRate' },
      { name: 'averageSessionDuration' },
      { name: 'conversions' }
    ],
    orderBys: [
      { dimension: { dimensionName: 'date' }, desc: false }
    ]
  };
  
  const response = AnalyticsData.Properties.runReport(request, property);
  
  if (!response.rows || response.rows.length === 0) {
    console.log('日別サマリ: データなし');
    return;
  }
  
  const now = new Date();
  const values = response.rows.map(row => {
    const dim = row.dimensionValues;
    const met = row.metricValues;
    
    // 日付を YYYY-MM-DD 形式に変換
    const dateStr = dim[0].value;
    const formattedDate = dateStr.substring(0, 4) + '-' + 
                          dateStr.substring(4, 6) + '-' + 
                          dateStr.substring(6, 8);
    
    return [
      now,
      formattedDate,
      Number(met[0].value),
      Number(met[1].value),
      Number(met[2].value),
      Number(met[3].value),
      (Number(met[4].value) * 100).toFixed(1),
      Number(met[5].value).toFixed(1),
      Number(met[6].value)
    ];
  });
  
  // シートをクリアして書き込み
  sheet.clear();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(2, 1, values.length, header.length).setValues(values);
  
  // ヘッダー行を固定・太字
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
}


// ============================================
// 2. ページ別レポート
// ============================================
/**
 * ページ別のアクセス状況
 * どのページが見られているか把握
 */
function exportPageReport() {
  const sheetName = CONFIG.SHEETS.PAGE_REPORT;
  const sheet = getOrCreateSheet(sheetName);
  
  const header = [
    '取得日時',
    'ページパス',
    'ページタイトル',
    'ページビュー数',
    'ユーザー数',
    'セッション数',
    '平均滞在時間(秒)',
    '直帰率(%)'
  ];
  
  const property = 'properties/' + CONFIG.GA4_PROPERTY_ID;
  
  const request = {
    dateRanges: [
      { startDate: CONFIG.DATE_RANGE_DAYS + 'daysAgo', endDate: 'yesterday' }
    ],
    dimensions: [
      { name: 'pagePath' },
      { name: 'pageTitle' }
    ],
    metrics: [
      { name: 'screenPageViews' },
      { name: 'totalUsers' },
      { name: 'sessions' },
      { name: 'averageSessionDuration' },
      { name: 'bounceRate' }
    ],
    orderBys: [
      { metric: { metricName: 'screenPageViews' }, desc: true }
    ],
    limit: 100  // 上位100ページ
  };
  
  const response = AnalyticsData.Properties.runReport(request, property);
  
  if (!response.rows || response.rows.length === 0) {
    console.log('ページ別: データなし');
    return;
  }
  
  const now = new Date();
  const values = response.rows.map(row => {
    const dim = row.dimensionValues;
    const met = row.metricValues;
    
    return [
      now,
      dim[0].value,
      dim[1].value,
      Number(met[0].value),
      Number(met[1].value),
      Number(met[2].value),
      Number(met[3].value).toFixed(1),
      (Number(met[4].value) * 100).toFixed(1)
    ];
  });
  
  sheet.clear();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(2, 1, values.length, header.length).setValues(values);
  
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
}


// ============================================
// 3. 集客チャネル別レポート
// ============================================
/**
 * どこからお客様が来ているか
 * - 検索、SNS、直接、参照元など
 */
function exportChannelReport() {
  const sheetName = CONFIG.SHEETS.CHANNEL_REPORT;
  const sheet = getOrCreateSheet(sheetName);
  
  const header = [
    '取得日時',
    'チャネルグループ',
    '参照元',
    'メディア',
    'セッション数',
    'ユーザー数',
    '新規ユーザー数',
    'コンバージョン数',
    'エンゲージメント率(%)'
  ];
  
  const property = 'properties/' + CONFIG.GA4_PROPERTY_ID;
  
  const request = {
    dateRanges: [
      { startDate: CONFIG.DATE_RANGE_DAYS + 'daysAgo', endDate: 'yesterday' }
    ],
    dimensions: [
      { name: 'sessionDefaultChannelGroup' },
      { name: 'sessionSource' },
      { name: 'sessionMedium' }
    ],
    metrics: [
      { name: 'sessions' },
      { name: 'totalUsers' },
      { name: 'newUsers' },
      { name: 'conversions' },
      { name: 'engagementRate' }
    ],
    orderBys: [
      { metric: { metricName: 'sessions' }, desc: true }
    ],
    limit: 50
  };
  
  const response = AnalyticsData.Properties.runReport(request, property);
  
  if (!response.rows || response.rows.length === 0) {
    console.log('集客チャネル別: データなし');
    return;
  }
  
  const now = new Date();
  const values = response.rows.map(row => {
    const dim = row.dimensionValues;
    const met = row.metricValues;
    
    return [
      now,
      dim[0].value,
      dim[1].value,
      dim[2].value,
      Number(met[0].value),
      Number(met[1].value),
      Number(met[2].value),
      Number(met[3].value),
      (Number(met[4].value) * 100).toFixed(1)
    ];
  });
  
  sheet.clear();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(2, 1, values.length, header.length).setValues(values);
  
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
}


// ============================================
// 4. コンバージョン別レポート
// ============================================
/**
 * コンバージョンイベントごとの発生状況
 * - 購入、お問合せ、電話リンクなど
 */
function exportConversionReport() {
  const sheetName = CONFIG.SHEETS.CONVERSION_REPORT;
  const sheet = getOrCreateSheet(sheetName);
  
  const header = [
    '取得日時',
    '日付',
    'イベント名',
    'イベント数',
    'ユーザー数'
  ];
  
  const property = 'properties/' + CONFIG.GA4_PROPERTY_ID;
  
  // イベント名でフィルタリング
  const eventFilter = {
    filter: {
      fieldName: 'eventName',
      inListFilter: {
        values: CONFIG.CONVERSION_EVENTS
      }
    }
  };
  
  const request = {
    dateRanges: [
      { startDate: CONFIG.DATE_RANGE_DAYS + 'daysAgo', endDate: 'yesterday' }
    ],
    dimensions: [
      { name: 'date' },
      { name: 'eventName' }
    ],
    metrics: [
      { name: 'eventCount' },
      { name: 'totalUsers' }
    ],
    dimensionFilter: eventFilter,
    orderBys: [
      { dimension: { dimensionName: 'date' }, desc: false }
    ]
  };
  
  const response = AnalyticsData.Properties.runReport(request, property);
  
  if (!response.rows || response.rows.length === 0) {
    console.log('コンバージョン別: データなし');
    return;
  }
  
  const now = new Date();
  const values = response.rows.map(row => {
    const dim = row.dimensionValues;
    const met = row.metricValues;
    
    const dateStr = dim[0].value;
    const formattedDate = dateStr.substring(0, 4) + '-' + 
                          dateStr.substring(4, 6) + '-' + 
                          dateStr.substring(6, 8);
    
    return [
      now,
      formattedDate,
      dim[1].value,
      Number(met[0].value),
      Number(met[1].value)
    ];
  });
  
  sheet.clear();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(2, 1, values.length, header.length).setValues(values);
  
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
}


// ============================================
// ユーティリティ関数
// ============================================
/**
 * シートを取得、なければ作成
 */
function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    console.log('シート作成: ' + sheetName);
  }
  
  return sheet;
}


// ============================================
// 追加機能：履歴蓄積版（上書きではなく追記）
// ============================================
/**
 * 日別サマリを追記モードで蓄積
 * 長期のトレンド分析用
 */
function appendDailySummary() {
  const sheetName = 'GA4_日別サマリ_履歴';
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  
  const header = [
    '取得日時',
    '日付',
    'ユーザー数',
    '新規ユーザー数',
    'セッション数',
    'ページビュー数',
    'コンバージョン数'
  ];
  
  // シートがなければ作成してヘッダーを入れる
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
  }
  
  const property = 'properties/' + CONFIG.GA4_PROPERTY_ID;
  
  // 昨日のデータのみ取得
  const request = {
    dateRanges: [
      { startDate: 'yesterday', endDate: 'yesterday' }
    ],
    dimensions: [
      { name: 'date' }
    ],
    metrics: [
      { name: 'totalUsers' },
      { name: 'newUsers' },
      { name: 'sessions' },
      { name: 'screenPageViews' },
      { name: 'conversions' }
    ]
  };
  
  const response = AnalyticsData.Properties.runReport(request, property);
  
  if (!response.rows || response.rows.length === 0) {
    console.log('履歴追記: データなし');
    return;
  }
  
  const now = new Date();
  const row = response.rows[0];
  const dim = row.dimensionValues;
  const met = row.metricValues;
  
  const dateStr = dim[0].value;
  const formattedDate = dateStr.substring(0, 4) + '-' + 
                        dateStr.substring(4, 6) + '-' + 
                        dateStr.substring(6, 8);
  
  const newRow = [
    now,
    formattedDate,
    Number(met[0].value),
    Number(met[1].value),
    Number(met[2].value),
    Number(met[3].value),
    Number(met[4].value)
  ];
  
  // 最終行に追記
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);
  
  console.log('履歴追記完了: ' + formattedDate);
}


// ============================================
// Web API（doGet）
// ============================================
/**
 * Webアプリとしてデプロイ後、URLパラメータでデータを取得
 * ?type=daily    → 日別サマリ
 * ?type=pages    → ページ別
 * ?type=channels → 集客チャネル別
 * ?type=conversions → コンバージョン別
 */
function doGet(e) {
  var type = (e && e.parameter && e.parameter.type) ? e.parameter.type : '';
  var result;

  try {
    switch (type) {
      case 'daily':
        result = getDataDailySummary();
        break;
      case 'pages':
        result = getDataPageReport();
        break;
      case 'channels':
        result = getDataChannelReport();
        break;
      case 'conversions':
        result = getDataConversionReport();
        break;
      default:
        result = { error: 'パラメータ type に daily, pages, channels, conversions のいずれかを指定してください' };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 日別サマリデータを取得してJSON用オブジェクトで返す
 */
function getDataDailySummary() {
  var property = 'properties/' + CONFIG.GA4_PROPERTY_ID;
  var request = {
    dateRanges: [{ startDate: CONFIG.DATE_RANGE_DAYS + 'daysAgo', endDate: 'yesterday' }],
    dimensions: [{ name: 'date' }],
    metrics: [
      { name: 'totalUsers' },
      { name: 'newUsers' },
      { name: 'sessions' },
      { name: 'screenPageViews' },
      { name: 'engagementRate' },
      { name: 'averageSessionDuration' },
      { name: 'conversions' }
    ],
    orderBys: [{ dimension: { dimensionName: 'date' }, desc: false }]
  };

  var response = AnalyticsData.Properties.runReport(request, property);
  if (!response.rows || response.rows.length === 0) return [];

  return response.rows.map(function(row) {
    var d = row.dimensionValues[0].value;
    var m = row.metricValues;
    return {
      date: d.substring(0, 4) + '-' + d.substring(4, 6) + '-' + d.substring(6, 8),
      users: Number(m[0].value),
      newUsers: Number(m[1].value),
      sessions: Number(m[2].value),
      pageviews: Number(m[3].value),
      engagementRate: Number((Number(m[4].value) * 100).toFixed(1)),
      avgSessionDuration: Number(Number(m[5].value).toFixed(1)),
      conversions: Number(m[6].value)
    };
  });
}

/**
 * ページ別データを取得
 */
function getDataPageReport() {
  var property = 'properties/' + CONFIG.GA4_PROPERTY_ID;
  var request = {
    dateRanges: [{ startDate: CONFIG.DATE_RANGE_DAYS + 'daysAgo', endDate: 'yesterday' }],
    dimensions: [{ name: 'pagePath' }, { name: 'pageTitle' }],
    metrics: [
      { name: 'screenPageViews' },
      { name: 'totalUsers' },
      { name: 'sessions' },
      { name: 'averageSessionDuration' },
      { name: 'bounceRate' }
    ],
    orderBys: [{ metric: { metricName: 'screenPageViews' }, desc: true }],
    limit: 100
  };

  var response = AnalyticsData.Properties.runReport(request, property);
  if (!response.rows || response.rows.length === 0) return [];

  return response.rows.map(function(row) {
    var d = row.dimensionValues;
    var m = row.metricValues;
    return {
      pagePath: d[0].value,
      pageTitle: d[1].value,
      pageviews: Number(m[0].value),
      users: Number(m[1].value),
      sessions: Number(m[2].value),
      avgDuration: Number(Number(m[3].value).toFixed(1)),
      bounceRate: Number((Number(m[4].value) * 100).toFixed(1))
    };
  });
}

/**
 * 集客チャネル別データを取得
 */
function getDataChannelReport() {
  var property = 'properties/' + CONFIG.GA4_PROPERTY_ID;
  var request = {
    dateRanges: [{ startDate: CONFIG.DATE_RANGE_DAYS + 'daysAgo', endDate: 'yesterday' }],
    dimensions: [
      { name: 'sessionDefaultChannelGroup' },
      { name: 'sessionSource' },
      { name: 'sessionMedium' }
    ],
    metrics: [
      { name: 'sessions' },
      { name: 'totalUsers' },
      { name: 'newUsers' },
      { name: 'conversions' },
      { name: 'engagementRate' }
    ],
    orderBys: [{ metric: { metricName: 'sessions' }, desc: true }],
    limit: 50
  };

  var response = AnalyticsData.Properties.runReport(request, property);
  if (!response.rows || response.rows.length === 0) return [];

  return response.rows.map(function(row) {
    var d = row.dimensionValues;
    var m = row.metricValues;
    return {
      channel: d[0].value,
      source: d[1].value,
      medium: d[2].value,
      sessions: Number(m[0].value),
      users: Number(m[1].value),
      newUsers: Number(m[2].value),
      conversions: Number(m[3].value),
      engagementRate: Number((Number(m[4].value) * 100).toFixed(1))
    };
  });
}

/**
 * コンバージョン別データを取得
 */
function getDataConversionReport() {
  var property = 'properties/' + CONFIG.GA4_PROPERTY_ID;
  var request = {
    dateRanges: [{ startDate: CONFIG.DATE_RANGE_DAYS + 'daysAgo', endDate: 'yesterday' }],
    dimensions: [{ name: 'date' }, { name: 'eventName' }],
    metrics: [
      { name: 'eventCount' },
      { name: 'totalUsers' }
    ],
    dimensionFilter: {
      filter: {
        fieldName: 'eventName',
        inListFilter: { values: CONFIG.CONVERSION_EVENTS }
      }
    },
    orderBys: [{ dimension: { dimensionName: 'date' }, desc: false }]
  };

  var response = AnalyticsData.Properties.runReport(request, property);
  if (!response.rows || response.rows.length === 0) return [];

  return response.rows.map(function(row) {
    var d = row.dimensionValues;
    var m = row.metricValues;
    var dateStr = d[0].value;
    return {
      date: dateStr.substring(0, 4) + '-' + dateStr.substring(4, 6) + '-' + dateStr.substring(6, 8),
      eventName: d[1].value,
      eventCount: Number(m[0].value),
      users: Number(m[1].value)
    };
  });
}


// ============================================
// テスト用関数
// ============================================
/**
 * 接続テスト
 * まずこれを実行して認証とAPI接続を確認
 */
function testConnection() {
  const property = 'properties/' + CONFIG.GA4_PROPERTY_ID;
  
  const request = {
    dateRanges: [
      { startDate: '7daysAgo', endDate: 'yesterday' }
    ],
    metrics: [
      { name: 'totalUsers' }
    ]
  };
  
  try {
    const response = AnalyticsData.Properties.runReport(request, property);
    console.log('接続成功！');
    console.log('過去7日間のユーザー数: ' + response.rows[0].metricValues[0].value);
  } catch (error) {
    console.error('接続エラー:', error);
    console.log('【確認事項】');
    console.log('1. Google Analytics Data API が有効になっているか');
    console.log('2. Apps Script のサービスに AnalyticsData が追加されているか');
    console.log('3. GA4プロパティへのアクセス権限があるか');
  }
}