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
    CONVERSION_REPORT: 'GA4_コンバージョン別',
    AI_ANALYSIS: 'GA4_AI分析',
    AI_ANALYSIS_LOG: 'GA4_AI分析_ログ'
  },

  GEMINI_MODEL: 'gemini-2.0-flash',
  
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
// AI分析（Gemini API）
// ============================================
/**
 * 本日のAI分析を取得。キャッシュがあればそれを返し、無ければ生成して保存。
 * 呼び出しごとに実行ログシート (GA4_AI分析_ログ) に1行追記する。
 */
function getOrGenerateAIAnalysis() {
  var startedAt = new Date();
  var today = Utilities.formatDate(startedAt, 'Asia/Tokyo', 'yyyy-MM-dd');
  var sheet = getOrCreateSheet(CONFIG.SHEETS.AI_ANALYSIS);

  // ヘッダー初期化
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 3).setValues([['日付', '生成時刻', '分析内容']]);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // 今日のキャッシュがあるか確認
  var cachedRow = null;
  if (sheet.getLastRow() >= 2) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < data.length; i++) {
      var cellDate = data[i][0];
      var dateStr = (cellDate instanceof Date)
        ? Utilities.formatDate(cellDate, 'Asia/Tokyo', 'yyyy-MM-dd')
        : String(cellDate);
      if (dateStr === today) {
        cachedRow = data[i];
        break;
      }
    }
  }

  try {
    if (cachedRow) {
      var cachedResult = {
        cached: true,
        date: today,
        generatedAt: cachedRow[1],
        analysis: cachedRow[2]
      };
      appendAIAnalysisLog({
        timestamp: new Date(),
        kind: 'キャッシュ返却',
        targetDate: today,
        elapsedMs: new Date().getTime() - startedAt.getTime(),
        summary: truncateForLog(cachedRow[2])
      });
      return cachedResult;
    }

    // 新規生成
    var summaryData = collectSummaryDataForAI();
    var prompt = buildAIPrompt(summaryData);
    var analysis = callGeminiAPI(prompt);

    var now = new Date();
    sheet.appendRow([today, now, analysis]);

    appendAIAnalysisLog({
      timestamp: now,
      kind: '新規生成',
      targetDate: today,
      elapsedMs: now.getTime() - startedAt.getTime(),
      summary: truncateForLog(analysis)
    });

    return {
      cached: false,
      date: today,
      generatedAt: now,
      analysis: analysis
    };

  } catch (e) {
    try {
      appendAIAnalysisLog({
        timestamp: new Date(),
        kind: 'エラー',
        targetDate: today,
        elapsedMs: new Date().getTime() - startedAt.getTime(),
        summary: String(e && e.message ? e.message : e).substring(0, 120)
      });
    } catch (logErr) {
      // ログ書き込み失敗は握りつぶす（本エラーを優先）
    }
    throw e;
  }
}

/**
 * AI分析の実行ログを GA4_AI分析_ログ シートに1行追記
 */
function appendAIAnalysisLog(logRow) {
  var sheet = getOrCreateSheet(CONFIG.SHEETS.AI_ANALYSIS_LOG);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 5).setValues([[
      'タイムスタンプ',
      '実行種別',
      '対象日付',
      '処理時間(ms)',
      '分析内容サマリ(冒頭120文字)'
    ]]);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 90);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 600);
  }

  sheet.appendRow([
    logRow.timestamp,
    logRow.kind,
    logRow.targetDate,
    logRow.elapsedMs,
    logRow.summary
  ]);
}

/**
 * ログ用に文字列を整形（空白圧縮 + 120文字で切り詰め）
 */
function truncateForLog(text) {
  if (!text) return '';
  var s = String(text).replace(/\s+/g, ' ').trim();
  return s.length > 120 ? s.substring(0, 120) + '...' : s;
}

/**
 * スプレッドシートの4シートから要約データを読み取り
 */
function collectSummaryDataForAI() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var result = {
    daily: [],
    pages: [],
    channels: [],
    conversions: []
  };

  var dailySheet = ss.getSheetByName(CONFIG.SHEETS.DAILY_SUMMARY);
  if (dailySheet && dailySheet.getLastRow() > 1) {
    result.daily = dailySheet.getRange(2, 1, Math.min(30, dailySheet.getLastRow() - 1), dailySheet.getLastColumn()).getValues();
  }

  var pagesSheet = ss.getSheetByName(CONFIG.SHEETS.PAGE_REPORT);
  if (pagesSheet && pagesSheet.getLastRow() > 1) {
    result.pages = pagesSheet.getRange(2, 1, Math.min(20, pagesSheet.getLastRow() - 1), pagesSheet.getLastColumn()).getValues();
  }

  var chSheet = ss.getSheetByName(CONFIG.SHEETS.CHANNEL_REPORT);
  if (chSheet && chSheet.getLastRow() > 1) {
    result.channels = chSheet.getRange(2, 1, chSheet.getLastRow() - 1, chSheet.getLastColumn()).getValues();
  }

  var cvSheet = ss.getSheetByName(CONFIG.SHEETS.CONVERSION_REPORT);
  if (cvSheet && cvSheet.getLastRow() > 1) {
    result.conversions = cvSheet.getRange(2, 1, cvSheet.getLastRow() - 1, cvSheet.getLastColumn()).getValues();
  }

  return result;
}

/**
 * Gemini API用プロンプトを構築
 */
function buildAIPrompt(data) {
  var lines = [];

  // ロール設定
  lines.push('あなたはEC・D2Cマーケティングに10年以上の経験を持つデータアナリストです。');
  lines.push('GA4データから具体的かつ実践的な売上改善提案を行うプロフェッショナルとして、経営判断に直結する分析レポートを作成してください。');
  lines.push('');

  // コンテキスト
  lines.push('## 対象サイトの基本情報');
  lines.push('- ブランド名: TARA SISTER');
  lines.push('- URL: https://www.tarasister.com');
  lines.push('- 業態: スキンケア・化粧品のD2Cブランド（BASEで構築）');
  lines.push('- 特徴: 自然派志向、コンセプトストーリー重視、リピート購入が売上の柱');
  lines.push('- 想定顧客: 肌に気を使う30-50代女性');
  lines.push('- 主要商品: Cleansing Massage Oil (クレンジングオイル)、Body wash DANA、Mary care oil など');
  lines.push('');

  lines.push('## 分析対象データ（過去30日）');
  lines.push('');
  lines.push('### 日別サマリ (日付 / ユーザー / セッション / PV / CV)');
  data.daily.forEach(function(row) { lines.push(row.join('\t')); });
  lines.push('');
  lines.push('### ページ別TOP20 (タイトル / パス / PV / ユーザー / 滞在秒 / 直帰率)');
  data.pages.forEach(function(row) { lines.push(row.join('\t')); });
  lines.push('');
  lines.push('### 集客チャネル別');
  data.channels.forEach(function(row) { lines.push(row.join('\t')); });
  lines.push('');
  lines.push('### コンバージョンイベント');
  data.conversions.forEach(function(row) { lines.push(row.join('\t')); });
  lines.push('');

  // 出力指示
  lines.push('## 出力指示');
  lines.push('');
  lines.push('以下の7セクションで構成された分析レポートを、Markdown形式・日本語で作成してください。');
  lines.push('全体で2500-3500文字程度を目安に、具体的な数値と根拠を伴う実践的な内容にしてください。');
  lines.push('');
  lines.push('### 📈 今週のトレンド');
  lines.push('- 過去30日の数値の推移、急増急減したポイント、曜日傾向を指摘');
  lines.push('- 直近7日と前週を比較して、改善/悪化している指標を明示');
  lines.push('- 具体的な日付と数値を引用');
  lines.push('');
  lines.push('### 🎯 注目すべきコンテンツ');
  lines.push('- PV/ユーザー/滞在時間/直帰率のバランスから、真に好調なページを特定');
  lines.push('- なぜそのページが読まれているのか、コンテンツの特性から仮説を立てる');
  lines.push('- 他ページへの横展開アイデア');
  lines.push('');
  lines.push('### ⚠️ 改善が必要なページ');
  lines.push('- 直帰率が高い/滞在時間が短い/ユーザー数の割にPVが少ないページを指摘');
  lines.push('- 各ページに対して具体的な改善案を1つ以上（見出し改善、画像追加、CTA配置、関連商品リンク、等）');
  lines.push('- 優先度を 高/中/低 で明示');
  lines.push('');
  lines.push('### 🚪 集客チャネル分析');
  lines.push('- 流入チャネル別のセッション数・割合');
  lines.push('- 各チャネルの強み・弱み（自然流入依存か広告依存か、持続性はどうか）');
  lines.push('- 投資すべきチャネル、縮小してよいチャネル');
  lines.push('');
  lines.push('### 📊 転換率 (CVR) の詳細分析');
  lines.push('- 全体CVRの算出: (CV数 ÷ ユーザー数) × 100 を計算して%で表示');
  lines.push('- ECサイトの業界平均（1-2%程度）との比較');
  lines.push('- CVRを上げるために、今のサイトで最も効果が見込める改善点を2-3個');
  lines.push('- ファネル（認知→閲覧→カート→購入）のどこで離脱しているかの推測');
  lines.push('');
  lines.push('### 💡 優先度付き改善アクション');
  lines.push('以下の3つの時間軸で、実行可能な具体アクションを提示:');
  lines.push('- **今週中に着手**: 即効性◎、工数小（例: タイトル改善、ボタン配置変更）');
  lines.push('- **今月中に実施**: 効果○、工数中（例: 人気商品のLP強化、FAQ追加）');
  lines.push('- **3ヶ月スパン**: 効果◎、工数大（例: リピート施策、メール自動化、Instagram運用強化）');
  lines.push('各アクションに「なぜこれが効くのか」の根拠を付けること');
  lines.push('');
  lines.push('### 🔍 次回チェックすべき定点観測指標');
  lines.push('- 今回特定した課題に対して、翌週/翌月に見るべき指標を3-5個');
  lines.push('- 「この数字がこう変化したら施策成功」の基準値');
  lines.push('');
  lines.push('---');
  lines.push('');
  lines.push('【厳守ルール】');
  lines.push('- 数値は必ず元データから正確に引用する（計算した数値には "(独自算出)" と付記）');
  lines.push('- 推測や仮説には「〜と思われる」「〜の可能性がある」を付けて事実と区別する');
  lines.push('- 具体例・具体策のない抽象的な助言は避ける');
  lines.push('- 「頑張りましょう」「期待しています」のような情緒表現は一切使わない');
  lines.push('- 経営判断に資する客観的・定量的なレポートにする');

  return lines.join('\n');
}

/**
 * Gemini API を呼び出してテキストを返す
 */
function callGeminiAPI(prompt) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY が Script Properties に設定されていません');
  }

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + CONFIG.GEMINI_MODEL + ':generateContent?key=' + apiKey;

  var payload = {
    contents: [{
      parts: [{ text: prompt }]
    }],
    generationConfig: {
      temperature: 0.8,
      maxOutputTokens: 8192,
      topP: 0.95,
      topK: 40
    }
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();
  var body = response.getContentText();

  if (code !== 200) {
    throw new Error('Gemini API エラー: HTTP ' + code + ' ' + body);
  }

  var json = JSON.parse(body);
  if (!json.candidates || !json.candidates[0] || !json.candidates[0].content || !json.candidates[0].content.parts) {
    throw new Error('Gemini API 応答形式エラー: ' + body);
  }

  var candidate = json.candidates[0];
  if (candidate.finishReason && candidate.finishReason !== 'STOP') {
    console.log('Gemini finishReason:', candidate.finishReason);
    console.log('Gemini full response:', JSON.stringify(json).substring(0, 500));
  }

  return candidate.content.parts[0].text;
}


// ============================================
// Web API（doGet）
// ============================================
/**
 * Webアプリとしてデプロイ後、URLパラメータでデータを取得
 * ?type=daily       → 日別サマリ
 * ?type=pages       → ページ別
 * ?type=channels    → 集客チャネル別
 * ?type=conversions → コンバージョン別
 * ?type=ai_analysis → Gemini API によるAI分析（1日1回キャッシュ）
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
      case 'ai_analysis':
        result = getOrGenerateAIAnalysis();
        break;
      default:
        result = { error: 'パラメータ type に daily, pages, channels, conversions, ai_analysis のいずれかを指定してください' };
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