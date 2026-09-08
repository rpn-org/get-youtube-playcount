// ==========================================
// 動画ごとの処理・シート操作・サンプリング・個別グラフ
// ==========================================

/**
 * 1本の動画を処理する（取得 → 記録 → 補完 → 間引き → グラフ更新）。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string}      id              YouTube 動画 ID
 * @param {number}      index           現在のインデックス（ログ用）
 * @param {number}      total           総件数（ログ用）
 * @param {Date}        now             実行日時
 * @param {object|null} preloadedVideo  fetchAllVideoData_ で取得済みの動画オブジェクト
 */
function processVideo_(ss, id, index, total, now, preloadedVideo) {
  try {
    const video = preloadedVideo ?? fetchVideoData_(id);
    if (!video) {
      console.warn(`[${index + 1}/${total}] 動画が見つかりません (id: ${id})`);
      return null;
    }

    const { fullTitle, viewCount, publishedAt, channelId, channelTitle } = parseVideoData_(video);
    const sheetName = buildSheetName_(fullTitle, id);
    const sheet     = getOrCreateVideoSheet_(ss, sheetName, fullTitle, publishedAt, channelId, channelTitle, id);

    // 同一タイムスタンプが既に存在する場合はスキップ（冪等性の担保）
    // CONFIG.FORCE_WRITE = true のときはスキップしない
    const lastRow = sheet.getLastRow();
    if (!CONFIG.FORCE_WRITE && lastRow >= 4) {
      const existing = sheet.getRange(4, 1, lastRow - 3, 1).getValues();
      if (existing.some(r => r[0] instanceof Date && r[0].getTime() === now.getTime())) {
        console.log(`[${index + 1}/${total}] ${sheetName}: スキップ（同一タイムスタンプ既存）`);
        return sheetName;
      }
    }

    sheet.appendRow([now, viewCount]);
    console.log(`[${index + 1}/${total}] ${sheetName}: ${viewCount.toLocaleString()} 回`);

    fillInitialGrowthCurve_(sheet, publishedAt);
    runSampling_(sheet);
    // グラフ更新は updateAllCharts() に分離（実行時間超過対策）
    sortVideoSheetDescending_(sheet);
    normalizeTimestampFormat_(sheet);
    // ソート後のデータを1回読んで渡す（updateGrowthSummary_ 内の重複読み込みを省く）
    const lastRow_ = sheet.getLastRow();
    const allData_ = lastRow_ >= 4 ? sheet.getRange(4, 1, lastRow_ - 3, 2).getValues() : [];
    updateGrowthSummary_(sheet, viewCount, now, allData_);

    return sheetName;
  } catch (e) {
    console.error(`[${index + 1}/${total}] エラー (id: ${id}): ${e.message}\n${e.stack}`);
    return null;
  }
}

/**
 * 動画タイトルからシート名を生成する。
 * Google Sheets の制約: 31 文字以内、バックスラッシュ等の特殊文字禁止。
 * @param {string} fullTitle  動画タイトル
 * @param {string} fallback   タイトルが空の場合に使う動画 ID
 * @returns {string}
 */
function buildSheetName_(fullTitle, fallback) {
  const INVALID_CHARS = /[\\\/\?\*\:\[\]]/g;
  const MAX_LEN = 31;

  let name = fullTitle.replace(/【.*?】/g, '').trim().replace(INVALID_CHARS, '_');
  if (!name) return fallback;

  if (name.length > MAX_LEN) {
    const half = Math.floor((MAX_LEN - 1) / 2);
    name = name.substring(0, half) + '…' + name.substring(name.length - half);
  }

  return name;
}

/**
 * 動画シートを取得する。存在しなければ新規作成してヘッダーを書き込む。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetName
 * @param {string} fullTitle
 * @param {Date}   publishedAt
 * @param {string} channelId
 * @param {string} channelTitle
 * @param {string} videoId       YouTube 動画 ID（D2 に記録。サムネイル・リンクに使う）
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateVideoSheet_(ss, sheetName, fullTitle, publishedAt, channelId, channelTitle, videoId) {
  let sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    if (!sheet.getRange('A2').getValue()) {
      sheet.getRange('A2').setValue(publishedAt);
    }
    // 既存シートにチャンネル情報が未記入の場合のみ補完
    if (channelId && !sheet.getRange('B2').getValue()) {
      sheet.getRange('B2').setValue(channelId);
      sheet.getRange('C2').setValue(channelTitle || '');
    }
    // D2 は後から追加した項目のため、既存シートにも補完する
    if (videoId && !sheet.getRange('D2').getValue()) {
      sheet.getRange('D2').setValue(videoId);
    }
    return sheet;
  }

  console.log(`シートを作成: ${sheetName}`);
  sheet = ss.insertSheet(sheetName);

  sheet.getRange('A1').setValue(fullTitle).setFontWeight('bold');
  sheet.getRange('A2').setValue(publishedAt);
  sheet.getRange('B2').setValue(channelId   || '');
  sheet.getRange('C2').setValue(channelTitle || '');
  sheet.getRange('D2').setValue(videoId     || '');
  sheet.getRange('A3:B3').setValues([['日時', '再生数']]).setBackground('#eeeeee');
  sheet.setFrozenRows(3);
  sheet.appendRow([publishedAt, 0]); // 投稿日時点の起点レコード（再生数 = 0）

  return sheet;
}

/**
 * 動画シートのデータ行（row4以降）をタイムスタンプ降順に並び替える。
 * 最新データが常に行4（先頭）になる。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function sortVideoSheetDescending_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 4) return;
  sheet.getRange(4, 1, lastRow - 3, 2).sort({ column: 1, ascending: false });
}

/**
 * 動画シートの日時列（A列）の表示書式を統一する。
 *
 * Sheets は 00:00 ちょうどの Date を書き込むとそのセルを「日付のみ」書式に自動設定する。
 * Range.sort() は値だけを並び替えて書式をセル位置に残すため、日付のみ書式のセルに
 * 別の時刻の記録が入り込み、FORMATTED_VALUE で読むと時刻が 00:00 に潰れてしまう。
 * （WebUI の再生数グラフが単調増加のはずの箇所で山になっていた原因）
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function normalizeTimestampFormat_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return;
  sheet.getRange(4, 1, lastRow - 3, 1).setNumberFormat(TS_NUMBER_FORMAT);
}

/**
 * 動画シートをチャンネル ID ごとにグループ化して返す。
 * B2 = channelId、C2 = channelTitle を使用する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet[]} videoSheets
 * @returns {{ channelId: string, channelTitle: string, sheets: GoogleAppsScript.Spreadsheet.Sheet[] }[]}
 */
function groupSheetsByChannel_(videoSheets) {
  const groups = {};
  videoSheets.forEach(sh => {
    const channelId    = sh.getRange('B2').getValue() || '_unknown';
    const channelTitle = sh.getRange('C2').getValue() || channelId;
    if (!groups[channelId]) groups[channelId] = { channelTitle, sheets: [] };
    groups[channelId].sheets.push(sh);
  });
  return Object.entries(groups)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([channelId, { channelTitle, sheets }]) => ({ channelId, channelTitle, sheets }));
}

/**
 * CONFIG.SAMPLING.RULES に従い、古いデータを間引いてシートの肥大化を抑制する。
 * バケツ方式で均等に間引くため、記録タイミングのズレに強い。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function runSampling_(sheet) {
  const dataCount = sheet.getLastRow() - 3;
  if (dataCount < CONFIG.SAMPLING.MIN_ROWS_TO_SAMPLE) return;

  const values     = sheet.getRange(4, 1, dataCount, 2).getValues();
  const seenBucket = new Set();
  const now        = new Date();

  const keepRows = values.filter((row, index) => {
    if (index === 0 || index === values.length - 1) return true; // 先頭・末尾は必ず保持
    if (!(row[0] instanceof Date)) return false;

    // 現在時刻からの経過日数で判定（直近1か月は必ず全件保持）
    const ageDays = (now.getTime() - row[0].getTime()) / MS_PER_DAY;
    const rule    = CONFIG.SAMPLING.RULES.find(r => ageDays <= r.maxDays);

    if (!rule || rule.keepEveryHours === null) return true;

    const ageHours  = (now.getTime() - row[0].getTime()) / MS_PER_HOUR;
    const bucketKey = `${rule.keepEveryHours}_${Math.floor(ageHours / rule.keepEveryHours)}`;

    if (seenBucket.has(bucketKey)) return false;
    seenBucket.add(bucketKey);
    return true;
  });

  if (keepRows.length < CONFIG.SAMPLING.MIN_ROWS_TO_SAMPLE) {
    console.log(`間引きスキップ [${sheet.getName()}]: 間引き後 ${keepRows.length} 行 < 最低 ${CONFIG.SAMPLING.MIN_ROWS_TO_SAMPLE} 行`);
    return;
  }

  if (keepRows.length < values.length) {
    console.log(`間引き [${sheet.getName()}]: ${values.length} → ${keepRows.length} 行`);
    sheet.getRange(4, 1, sheet.getLastRow() - 3, 2).clearContent();
    sheet.getRange(4, 1, keepRows.length, 2).setValues(keepRows);
  }
}

/**
 * 各動画シートに再生数推移の折れ線グラフを作成または更新する。
 * 複数グラフが混在しても「このシートのデータのみ参照するグラフ」を再生数グラフと識別する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function updateIndividualChart_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return;

  const dataRange = sheet.getRange(`A3:B${lastRow}`);
  const title     = sheet.getRange('A1').getValue() || sheet.getName();

  // 再生数グラフ: このシート自身のデータのみ参照するグラフ
  const viewCountChart = sheet.getCharts().find(c =>
    c.getRanges().every(r => r.getSheet().getName() === sheet.getName())
  ) ?? null;

  const builder = (viewCountChart ? viewCountChart.modify() : sheet.newChart())
    .asLineChart()
    .clearRanges()
    .addRange(dataRange)
    .setPosition(10, 4, 0, 0)
    .setOption('title', title)
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { slantedText: true, slantedTextAngle: 45 })
    .setOption('vAxis', { format: '#,###' })
    .setOption('pointSize', 2)
    .setOption('lineWidth', 2);

  if (viewCountChart) {
    sheet.updateChart(builder.build());
  } else {
    sheet.insertChart(builder.build());
  }
}

/**
 * 順位データ（負値格納）から Y 軸の viewWindow を計算する。
 *   上端: bestRank を 10 の倍数に切り捨て（例: rank 3 → 0, rank 13 → 10）
 *         ただし 0 のときは GAS が setOption(key, 0) を無視するバグを回避するため 1 を返す
 *   下端: worstRank を 10 の倍数に切り上げ（例: rank 25 → 30, rank 20 → 20）
 * データが空なら { min: null, max: null } を返す（GAS 自動スケール）。
 * @param {number[]} storedValues  シートに格納された負値の順位データ
 * @returns {{ min: number|null, max: number|null }}
 */
function computeRankViewWindow_(storedValues) {
  const valid = storedValues.filter(v => typeof v === 'number' && v < 0);
  if (valid.length === 0) return { min: null, max: null };

  const bestRank  = Math.round(-Math.max(...valid)); // max(-3,-25) = -3 → 3
  const worstRank = Math.round(-Math.min(...valid)); // min(-3,-25) = -25 → 25

  // (bestRank - 1) で計算することで rank が 10 の倍数ちょうどの場合でも
  // 1グリッド上の余白を確保し、より良い順位が描画範囲外になるのを防ぐ
  // 例: rank 10 → floor(9/10)*10 = 0, rank 11 → floor(10/10)*10 = 10
  const topBoundaryRank    = Math.floor((bestRank - 1) / 10) * 10;
  const bottomBoundaryRank = Math.ceil (worstRank       / 10) * 10;

  return {
    // topBoundaryRank=0 のとき vMax=0 を渡すと GAS が falsy として無視するため 1 を使用
    max: topBoundaryRank === 0 ? 1 : -topBoundaryRank,
    min: -bottomBoundaryRank,
  };
}

/**
 * 各動画シートにチャンネル内順位の推移グラフを作成または更新する。
 * 「チャンネル内順位」シートから該当列を参照してグラフを描画する。
 * 順位シートにデータがなければ何もしない。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {GoogleAppsScript.Spreadsheet.Sheet}       sheet  動画シート
 */
function updateVideoRankChart_(ss, sheet) {
  const rankSheet = ss.getSheetByName(CONFIG.RANK_SHEET_NAME);
  if (!rankSheet) return;

  const rankLastRow = rankSheet.getLastRow();
  const rankLastCol = rankSheet.getLastColumn();
  if (rankLastRow < 3 || rankLastCol < 2) return;

  // 動画タイトルと一致する列を行2（タイトル行）から探す
  const videoTitle = sheet.getRange('A1').getValue();
  if (!videoTitle) return;

  const titleRow  = rankSheet.getRange(2, 2, 1, rankLastCol - 1).getValues()[0];
  const colOffset = titleRow.findIndex(t => String(t) === String(videoTitle));
  if (colOffset === -1) return; // この動画の順位データがまだない

  const rankCol  = colOffset + 2; // 1-based (B列=2から)
  const dataRows = rankLastRow - 1; // 行2ヘッダー〜lastRow

  // Y軸範囲を実データから計算（行3以降が実データ）
  const dataRowCount = rankLastRow - 2;
  const storedValues = dataRowCount > 0
    ? rankSheet.getRange(3, rankCol, dataRowCount, 1).getValues().flat()
    : [];
  const { min: vMin, max: vMax } = computeRankViewWindow_(storedValues);

  // 順位グラフ: チャンネル内順位シートのデータを参照するグラフ
  const rankChart = sheet.getCharts().find(c =>
    c.getRanges().some(r => r.getSheet().getName() === CONFIG.RANK_SHEET_NAME)
  ) ?? null;

  const builder = (rankChart ? rankChart.modify() : sheet.newChart())
    .asLineChart()
    .clearRanges()
    .addRange(rankSheet.getRange(2, 1, dataRows, 1))       // 日時列（行2ヘッダー含む）
    .addRange(rankSheet.getRange(2, rankCol, dataRows, 1)) // 順位列
    .setNumHeaders(1)
    .setPosition(10, 11, 0, 0)
    .setOption('title', 'チャンネル内順位')
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { slantedText: true, slantedTextAngle: 45 })
    .setOption('vAxis.format', '#,##0;#,##0;0')
    .setOption('vAxis.title', '順位（1位が上）')
    .setOption('interpolateNulls', true)
    .setOption('pointSize', 3)
    .setOption('lineWidth', 2)
    .setOption('width', 600)
    .setOption('height', 370);

  if (vMin !== null) {
    builder
      .setOption('vAxis.viewWindowMode', 'explicit')
      .setOption('vAxis.viewWindow.min', vMin)
      .setOption('vAxis.viewWindow.max', vMax);
  }

  if (rankChart) {
    sheet.updateChart(builder.build());
  } else {
    sheet.insertChart(builder.build());
  }
}

/**
 * 各動画シートの D4:H9 に直近の再生数増加量サマリーを書き込む。
 * データは降順ソート済みを前提とする（sortVideoSheetDescending_ 後に呼ぶこと）。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} currentViewCount  最新の再生数
 * @param {Date}   now               計測日時
 * @param {any[][]} allData          [[Date, viewCount], ...] 降順
 */
function updateGrowthSummary_(sheet, currentViewCount, now, allData) {
  if (!allData || allData.length === 0) return;

  const nowMs = now.getTime();

  function calcIncrease(fromMs, toMs) {
    const result = calcIncrease_(allData, currentViewCount, nowMs, fromMs, toMs);
    if (!result) return '---';
    return result.isEstimate ? `~${result.value}` : result.value;
  }

  const makeWindows = (unit, count) =>
    Array.from({ length: count }, (_, i) => [(i + 1) * unit, i * unit]);

  const N = CONFIG.SUMMARY_WINDOWS;
  const PERIODS = [
    { label: '1時間', windows: makeWindows(MS_PER_HOUR,     N) },
    { label: '1日',   windows: makeWindows(MS_PER_DAY,      N) },
    { label: '1週間', windows: makeWindows(7  * MS_PER_DAY, N) },
    { label: '1ヶ月', windows: makeWindows(30  * MS_PER_DAY, N) },
    { label: '1年',   windows: makeWindows(365 * MS_PER_DAY, N) },
  ];

  const headers = ['期間', '直近', ...Array.from({ length: N - 1 }, (_, i) => `${i + 1}期前`)];
  const tableData = [
    headers,
    ...PERIODS.map(({ label, windows }) => {
      const vals = windows.map(([from, to]) => calcIncrease(from, to));
      while (vals.length < N) vals.push('---');
      return [label, ...vals];
    }),
  ];

  const numRows  = tableData.length;
  const numCols  = headers.length;
  const startRow = 4;
  sheet.getRange(startRow, 4, numRows, numCols).setValues(tableData);

  const fmtKey = `summary_fmt_v2_${sheet.getName()}`;
  if (!PropertiesService.getScriptProperties().getProperty(fmtKey)) {
    sheet.getRange(startRow, 4, 1, numCols).setBackground('#eeeeee').setFontWeight('bold');
    sheet.getRange(startRow + 1, 5, numRows - 1, numCols - 1).setNumberFormat('#,##0');
    PropertiesService.getScriptProperties().setProperty(fmtKey, 'true');
  }
}

/**
 * 投稿日(0再生)と最初の計測値の間を、べき乗則曲線で補完する。
 *
 * alpha の推定には最初の 1〜3 計測点を log-log 空間で重み付き最小二乗法により求める。
 * 重み [1, 2, 4] は後の計測点ほど大きく設定し、最新の勾配に合わせる。
 * 計測点が 1 点のみの場合は平方根曲線（alpha = 0.5）にフォールバックする。
 *
 * ScriptProperties でシートごとに実行済みフラグを管理し、一度だけ実行する。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Date} publishedAt
 */
function fillInitialGrowthCurve_(sheet, publishedAt) {
  const propKey = `curve_filled_${sheet.getName()}`;
  if (PropertiesService.getScriptProperties().getProperty(propKey)) return;

  const dataCount = sheet.getLastRow() - 3;
  if (dataCount < 2) return; // 起点行 + 最低1計測が必要

  // 降順ソート済みデータへの誤適用を防止:
  // row4 > row5（新→古の順）なら、以前に実行済みだが propKey が失われた状態。
  // この場合、row5〜7 は最新データであり補完に使うと全期間に誤値を挿入してしまうためスキップ。
  const row4Date = sheet.getRange(4, 1).getValue();
  const row5Date = sheet.getRange(5, 1).getValue();
  if (row4Date instanceof Date && row5Date instanceof Date && row4Date > row5Date) {
    PropertiesService.getScriptProperties().setProperty(propKey, 'recovered');
    console.warn(`[${sheet.getName()}] 成長曲線: 降順ソート済みを検出 → 再実行をスキップ`);
    return;
  }

  const t0 = publishedAt.getTime();

  // 起点行(row4)直後の最大3計測点を読み込む（row5〜row7）
  const nSamples = Math.min(dataCount - 1, 3);
  const rawRows  = sheet.getRange(5, 1, nSamples, 2).getValues();

  const points = rawRows
    .filter(r => r[0] instanceof Date && r[1] > 0)
    .map(r => ({ d: r[0].getTime() - t0, v: r[1] }))
    .filter(p => p.d > 0);

  if (points.length === 0) return;

  // log-log 空間での重み付き最小二乗法で alpha を推定
  // Y = log(v) = log(C) + alpha * log(d) = log(C) + alpha * X
  let alpha = 0.5; // フォールバック: 平方根曲線
  if (points.length >= 2) {
    const ws = points.map((_, i) => Math.pow(2, i)); // [1, 2, 4]: 後の点ほど重く
    const xs = points.map(p => Math.log(p.d));
    const ys = points.map(p => Math.log(p.v));

    const W    = ws.reduce((s, w)    => s + w, 0);
    const xBar = ws.reduce((s, w, i) => s + w * xs[i], 0) / W;
    const yBar = ws.reduce((s, w, i) => s + w * ys[i], 0) / W;
    const num  = ws.reduce((s, w, i) => s + w * (xs[i] - xBar) * (ys[i] - yBar), 0);
    const den  = ws.reduce((s, w, i) => s + w * Math.pow(xs[i] - xBar, 2), 0);

    if (den > 0) alpha = Math.max(0.1, Math.min(1.5, num / den));
  }

  const { d: d1, v: v1 } = points[0];
  const C = v1 / Math.pow(d1, alpha);

  // SAMPLING.RULES と同じ頻度で補完点を生成する
  const newRows = [];
  let curMs = t0;
  while (true) {
    const ageDays = (curMs - t0) / MS_PER_DAY;
    const rule    = CONFIG.SAMPLING.RULES.find(r => ageDays <= r.maxDays);
    const stepMs  = (rule && rule.keepEveryHours ? rule.keepEveryHours : 1) * MS_PER_HOUR;
    curMs += stepMs;
    if (curMs >= t0 + d1) break;
    const v = Math.round(C * Math.pow(curMs - t0, alpha));
    if (v <= 0) continue;
    newRows.push([new Date(curMs), v]);
  }
  if (newRows.length === 0) return;

  sheet.insertRowsAfter(4, newRows.length);
  sheet.getRange(5, 1, newRows.length, 2).setValues(newRows);
  PropertiesService.getScriptProperties().setProperty(propKey, 'true');
  console.log(`成長曲線を補完 [${sheet.getName()}]: ${newRows.length} 点追加 (alpha=${alpha.toFixed(2)})`);
}

/**
 * 再生数が単調増加でない行（成長曲線の誤挿入等）を1シートから削除する。
 *
 * 再生数は単調増加しかあり得ないため、昇順で並べたときに
 * 直前行より再生数が小さい行は不正データとして除去する。
 * 処理後はデータを降順に並び替えて書き戻す。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {number} 削除した行数
 */
function removeNonMonotonicRows_(sheet) {
  const lastRow   = sheet.getLastRow();
  const dataCount = lastRow - 3;
  if (dataCount < 2) return 0;

  const values = sheet.getRange(4, 1, dataCount, 2).getValues();

  // 昇順ソートして単調増加チェック
  const sorted = [...values].sort((a, b) => {
    if (!(a[0] instanceof Date)) return 1;
    if (!(b[0] instanceof Date)) return -1;
    return a[0].getTime() - b[0].getTime();
  });

  let maxViewCount = -1;
  const keepRows = sorted.filter(row => {
    if (!(row[0] instanceof Date)) return false;
    const v = row[1];
    if (typeof v !== 'number' || v < maxViewCount) return false;
    maxViewCount = v;
    return true;
  });

  const removed = values.length - keepRows.length;
  if (removed === 0) return 0;

  console.log(`[${sheet.getName()}] 非単調行を削除: ${values.length} → ${keepRows.length} 行 (${removed} 件除去)`);

  // 降順に並び直して書き戻す
  keepRows.sort((a, b) => b[0].getTime() - a[0].getTime());
  sheet.getRange(4, 1, dataCount, 2).clearContent();
  sheet.getRange(4, 1, keepRows.length, 2).setValues(keepRows);
  return removed;
}

/**
 * 「チャンネル内順位」シートにチャンネルごとの順位推移グラフを作成する。
 *
 * - 既存グラフをすべて削除して再作成（オプション継承の問題を回避）
 * - チャンネルごとに1グラフ（行1のチャンネル名でグループ化）
 * - Y 軸反転: 1位が上部に表示される
 * - 凡例: 動画タイトルを右側に表示
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function updateRankHistoryChart_(ss) {
  const sheet = ss.getSheetByName(CONFIG.RANK_SHEET_NAME);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 3 || lastCol < 2) return; // 行1=チャンネル, 行2=タイトル, 行3+=データ

  // 既存グラフをすべて削除して再作成
  sheet.getCharts().forEach(c => sheet.removeChart(c));

  // 行1 (B1 以降) のチャンネル名を読み込んでグループ化
  const row1 = sheet.getRange(1, 2, 1, lastCol - 1).getValues()[0];
  const channelGroups = {}; // { channelTitle: [colIndex (1-based)...] }
  row1.forEach((channelTitle, i) => {
    const key = String(channelTitle || '_unknown');
    if (!channelGroups[key]) channelGroups[key] = [];
    channelGroups[key].push(i + 2); // 1-based (B列=2, C列=3, ...)
  });

  const CHART_WIDTH    = 1800;
  const CHART_HEIGHT   = 600;
  const ROW_HEIGHT_PX  = 21;
  const headerDataRows = lastRow - 1; // 行2（タイトルヘッダー）〜lastRow の行数
  let nextRow = 1; // 行1から縦積み

  // 列インデックス → 動画ID のマップ（ライブ判定に使用）
  const colMap      = getRankSheetColMap_();
  const colToVidId  = Object.fromEntries(Object.entries(colMap).map(([vid, col]) => [col, vid]));
  const liveIdSet   = new Set(CONFIG.LIVE_VIDEO_IDS);

  // チャンネルグループごと、さらにライブ/通常動画に分けてグラフを生成
  const dataRowCount = lastRow - 2; // 行3以降が実データ
  let chartCount = 0;
  Object.entries(channelGroups).forEach(([channelTitle, cols]) => {
    [
      { label: 'ライブ',   filter: col =>  liveIdSet.has(colToVidId[col] ?? '') },
      { label: '動画',     filter: col => !liveIdSet.has(colToVidId[col] ?? '') },
    ].forEach(({ label, filter }) => {
      const groupCols = cols.filter(filter);
      if (groupCols.length === 0) return;

      const storedValues = [];
      if (dataRowCount > 0) {
        groupCols.forEach(col => {
          sheet.getRange(3, col, dataRowCount, 1).getValues()
            .forEach(([v]) => { if (typeof v === 'number' && v < 0) storedValues.push(v); });
        });
      }
      const { min: vMin, max: vMax } = computeRankViewWindow_(storedValues);

      const builder = sheet.newChart()
        .asLineChart()
        .setNumHeaders(1)
        .addRange(sheet.getRange(2, 1, headerDataRows, 1));

      groupCols.forEach(col => {
        builder.addRange(sheet.getRange(2, col, headerDataRows, 1));
      });

      builder
        .setPosition(nextRow, 1, 0, 0)
        .setOption('title', `チャンネル内順位の推移 — ${channelTitle}（${label}）`)
        .setOption('legend', { position: 'right', textStyle: { fontSize: 9 } })
        .setOption('hAxis', { slantedText: true, slantedTextAngle: 45 })
        .setOption('vAxis.format', '#,##0;#,##0;0')
        .setOption('vAxis.title', '順位（1位が上）')
        .setOption('interpolateNulls', true)
        .setOption('pointSize', 3)
        .setOption('lineWidth', 2)
        .setOption('width', CHART_WIDTH)
        .setOption('height', CHART_HEIGHT);

      if (vMin !== null) {
        builder
          .setOption('vAxis.viewWindowMode', 'explicit')
          .setOption('vAxis.viewWindow.min', vMin)
          .setOption('vAxis.viewWindow.max', vMax);
      }

      sheet.insertChart(builder.build());
      nextRow += Math.ceil(CHART_HEIGHT / ROW_HEIGHT_PX) + 3;
      chartCount++;
    });
  });

  console.log(`${CONFIG.RANK_SHEET_NAME}: ${chartCount} グラフを更新しました`);
}

/**
 * Script Properties から「チャンネル内順位」シートの列マッピングを読み込む。
 * @returns {Object<string, number>}  { videoId: colIndex (1-based) }
 */
function getRankSheetColMap_() {
  const raw = PropertiesService.getScriptProperties().getProperty('rank_sheet_col_map');
  return raw ? JSON.parse(raw) : {};
}

/**
 * 「チャンネル内順位」シートの列マッピングを Script Properties に保存する。
 * @param {Object<string, number>} colMap
 */
function saveRankSheetColMap_(colMap) {
  PropertiesService.getScriptProperties().setProperty('rank_sheet_col_map', JSON.stringify(colMap));
}

/**
 * 「チャンネル内順位」シートに順位を記録する（ピボット形式）。
 *
 * レイアウト:
 *   行1（チャンネル行）: チャンネル | チャンネル名A | チャンネル名A | チャンネル名B | ...
 *   行2（タイトル行）: 日時 | 動画タイトルA | 動画タイトルB | ...
 *   行3以降（新しい順）: タイムスタンプ | 順位A | 順位B | ...
 *
 * 旧フォーマット（フラット追記型: B1='動画タイトル'、ピボット旧版: A1='日時'）は
 * 自動検出してクリア・再構築する。
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object<string, number>}  rankMap   { videoId: rank }
 * @param {Object<string, {title: string, channelTitle: string}>} metaMap
 * @param {Date} now
 */
function updateRankHistorySheet_(ss, rankMap, metaMap, now) {
  if (Object.keys(rankMap).length === 0) return;

  let sheet = ss.getSheetByName(CONFIG.RANK_SHEET_NAME);

  // 旧フォーマット検出 → クリアして再構築
  // ① 旧フラット追記型: B1 = '動画タイトル'
  // ② 旧ピボット（チャンネル行なし）: A1 = '日時'
  if (sheet) {
    const a1 = sheet.getRange('A1').getValue();
    const b1 = sheet.getRange('B1').getValue();
    if (b1 === '動画タイトル' || a1 === '日時') {
      sheet.clearContents();
      sheet.clearFormats();
      saveRankSheetColMap_({});
      console.log(`${CONFIG.RANK_SHEET_NAME}: 旧フォーマットを検出、クリアして再構築します`);
    }
  }

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.RANK_SHEET_NAME);
    console.log(`シートを作成: ${CONFIG.RANK_SHEET_NAME}`);
  }

  // 初期化: 行1=チャンネル行（水色）、行2=動画タイトル行（グレー）
  if (!sheet.getRange('A1').getValue()) {
    sheet.getRange('A1').setValue('チャンネル').setBackground('#cfe2f3').setFontWeight('bold');
    sheet.getRange('A2').setValue('日時').setBackground('#eeeeee').setFontWeight('bold');
    sheet.setFrozenRows(2);
    sheet.setColumnWidth(1, 160);
  }

  // 列マッピングを常にシートの行2から再構築する（Script Propertiesキャッシュの不整合を防ぐ）
  // 先頭列を優先して採用することで、clearRankCache後に追加された重複列を無視する
  let colMap = {};
  let mapChanged = false;
  const existingLastCol = sheet.getLastColumn();
  if (existingLastCol > 1) {
    const titleToId = {};
    Object.entries(metaMap).forEach(([id, meta]) => {
      if (meta.title) titleToId[meta.title] = id;
    });
    const row2vals = sheet.getRange(2, 2, 1, existingLastCol - 1).getValues()[0];
    row2vals.forEach((title, i) => {
      const id = titleToId[String(title)];
      if (id && !colMap[id]) colMap[id] = i + 2; // 先頭列を優先（重複列は無視）
    });
  }

  // 初回登録時は現在の順位昇順でカラムを並べる
  const sortedIds = Object.keys(rankMap).sort((a, b) => (rankMap[a] ?? Infinity) - (rankMap[b] ?? Infinity));

  sortedIds.forEach(id => {
    if (colMap[id]) return;
    const newCol       = sheet.getLastColumn() + 1;
    const title        = metaMap[id]?.title        || id;
    const channelTitle = metaMap[id]?.channelTitle || '';
    sheet.getRange(1, newCol).setValue(channelTitle).setBackground('#cfe2f3').setFontWeight('bold');
    sheet.getRange(2, newCol).setValue(title).setBackground('#eeeeee').setFontWeight('bold');
    sheet.setColumnWidth(newCol, 250);
    colMap[id]  = newCol;
    mapChanged  = true;
  });

  if (mapChanged) saveRankSheetColMap_(colMap);

  // データ行を行2の直後（行3）に挿入（新しいデータが先頭に来る）
  const totalCols = sheet.getLastColumn();
  const rowData   = new Array(totalCols).fill('');
  rowData[0]      = now;
  // GAS の direction:-1 が LINE チャートで機能しないため、負値を格納して Y 軸を自然反転させる
  // セルフォーマットで正値として表示し、チャート側でも符号なし書式を指定する
  Object.entries(rankMap).forEach(([id, rank]) => {
    if (colMap[id] && rank != null) rowData[colMap[id] - 1] = -rank;
  });

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) sheet.insertRowAfter(2);
  const dataRow = sheet.getRange(3, 1, 1, totalCols);
  dataRow.setValues([rowData]);
  // 行2（グレーヘッダー）の書式が継承されないよう明示的にリセット
  dataRow.setBackground(null).setFontWeight('normal');
  sheet.getRange(3, 1).setNumberFormat('yyyy/MM/dd HH:mm');
  // 順位列（B列以降）は負値を正値として表示
  if (totalCols >= 2) {
    sheet.getRange(3, 2, 1, totalCols - 1).setNumberFormat('#,##0;#,##0;0');
  }

  console.log(`チャンネル内順位シートに記録: ${now}`);
}
