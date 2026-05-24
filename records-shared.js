"use strict";

/* =========================================================
 * 杉ヶ平キャンプ場 月計表 記録機能 共通ロジック
 * - 保存・読込（LocalStorage）
 * - 料金計算（index.html と同じロジックを再掲）
 * - 月計表 Excel 出力用の列マッピング
 * ========================================================= */

var STORAGE_KEY = "sugigadaira-records-v1";

// 料金マスタ（index.html と同期）
var PRICES = {
  cottageA:   { overnight: 13200, day: 6600 },
  cottageB:   { overnight:  8800, day: 4400 },
  cottageC:   { overnight:  8800, day: 4400 },
  cottageD:   { overnight:  8800, day: 4400 },
  cottageE:   { overnight:  8800, day: 4400 },
  bungalow1:  { overnight:  3150, day: 1580 },
  bungalow2:  { overnight:  3150, day: 1580 },
  tentRental: { overnight:   880, day:  440 },
  tentBring:  { overnight:   660, day:  330 },
  sheet:      { overnight:  1100, day:  550 },
  perPerson:  { overnight:   220, day:  110 }
};

var COTTAGE_CAPACITY = {
  cottageA: { base: 6, max: 8 },
  cottageB: { base: 4, max: 6 },
  cottageC: { base: 4, max: 6 },
  cottageD: { base: 4, max: 6 },
  cottageE: { base: 4, max: 4 }
};

var FACILITY_LABEL = {
  cottageA:  "A棟ニッコウキスゲ",
  cottageB:  "B棟コメツツジ",
  cottageC:  "C棟シャクナゲ",
  cottageD:  "D棟チングルマ",
  cottageE:  "E棟イワカガミ",
  bungalow1: "バンガロー1番",
  bungalow2: "バンガロー2番"
};

/* ---------- 計算（index.html と同一ロジック） ---------- */
function calculate(s) {
  var isOvernight = s.stay === "overnight";
  var key = isOvernight ? "overnight" : "day";
  var nights = isOvernight ? Math.max(1, s.nights) : 1;
  var nightsMul = isOvernight ? (1 + (nights - 1) * 0.5) : 1;

  var total = 0;

  var cottages = s.facilities.filter(function (f) { return f.indexOf("cottage") === 0; });
  var bungalows = s.facilities.filter(function (f) { return f.indexOf("bungalow") === 0; });
  var hasTent = s.tentRental > 0 || s.tentBring > 0;

  cottages.concat(bungalows).forEach(function (f) {
    total += Math.round(PRICES[f][key] * nightsMul);
  });
  if (s.tentRental > 0) total += Math.round(PRICES.tentRental[key] * s.tentRental * nightsMul);
  if (s.tentBring > 0)  total += Math.round(PRICES.tentBring[key]  * s.tentBring  * nightsMul);

  var totalPeople = s.adults + s.children;
  var sheetCount = computeSheetCount(s);
  if (sheetCount > 0) total += Math.round(PRICES.sheet[key] * sheetCount * nightsMul);
  if (hasTent && s.tentPeople > 0) total += Math.round(PRICES.perPerson[key] * s.tentPeople * nightsMul);

  return { total: total, sheetCount: sheetCount };
}

function computeSheetCount(s) {
  var cottages = s.facilities.filter(function (f) { return f.indexOf("cottage") === 0; });
  if (cottages.length === 0) return 0;
  var bungalows = s.facilities.filter(function (f) { return f.indexOf("bungalow") === 0; });
  var hasTent = s.tentRental > 0 || s.tentBring > 0;
  var totalPeople = s.adults + s.children;

  var baseCap = cottages.reduce(function (acc, c) { return acc + COTTAGE_CAPACITY[c].base; }, 0);
  var maxCap  = cottages.reduce(function (acc, c) { return acc + COTTAGE_CAPACITY[c].max; }, 0);

  var cottageSleepers;
  if (bungalows.length > 0) cottageSleepers = 0;
  else if (hasTent) cottageSleepers = Math.max(0, totalPeople - s.tentPeople);
  else cottageSleepers = totalPeople;

  if (cottageSleepers <= baseCap) return 0;
  var addable = maxCap - baseCap;
  return Math.min(cottageSleepers - baseCap, addable);
}

/* ---------- 保存・読込 ---------- */
function loadRecords() {
  try {
    var raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return [];
    var arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    return [];
  }
}

function saveRecords(records) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(records));
}

function addRecord(record) {
  var records = loadRecords();
  records.push(record);
  saveRecords(records);
  return record;
}

function deleteRecord(id) {
  var records = loadRecords().filter(function (r) { return r.id !== id; });
  saveRecords(records);
}

function newId() {
  return String(Date.now()) + "-" + Math.random().toString(36).slice(2, 8);
}

/* ---------- 月計表用：1記録 → 列マトリクスへの振り分け ----------
 * 雛形 〇月キャンプ場利用集計(報告用).xlsx の列構成（0始まりインデックス）：
 *  0  : A 使用日
 *  1-4 : B-E コテージ6人用（宿泊/連泊/超過/日帰）
 *  5-8 : F-I コテージ4人用（宿泊/連泊/超過/日帰）
 *  9-12: J-M コテージ超過人数加算（宿泊/連泊/超過/日帰）
 * 13-16: N-Q バンガロー（宿泊/連泊/超過/日帰）
 * 17-19: R-T 貸しテント（宿泊/連泊/日帰）
 * 20-22: U-W 持込テント（宿泊/連泊/日帰）
 * 23-25: X-Z 使用人数加算（宿泊/連泊/日帰）
 * 26   : AA 小計
 * 27   : AB 減免額
 * 28   : AC 合計
 * 29   : AD 県内人数
 * 30   : AE 県外人数
 * 31   : AF 備考
 */
function recordToMatrixRow(record) {
  var s = record.state;
  var isOvernight = s.stay === "overnight";
  var nights = isOvernight ? Math.max(1, s.nights) : 1;
  var firstNight = isOvernight ? 1 : 0;
  var extraNights = isOvernight && nights > 1 ? (nights - 1) : 0;
  var dayCount = isOvernight ? 0 : 1;

  var cells = {};
  function add(col, v) {
    if (!v) return;
    cells[col] = (cells[col] || 0) + v;
  }

  s.facilities.forEach(function (f) {
    var base;
    if (f === "cottageA") base = 1;            // 6人用（B=1）
    else if (f === "cottageB" || f === "cottageC" || f === "cottageD" || f === "cottageE") base = 5; // 4人用（F=5）
    else if (f === "bungalow1" || f === "bungalow2") base = 13;   // バンガロー（N=13）
    else return;
    add(base + 0, firstNight);   // 宿泊（棟数=1）
    add(base + 1, extraNights);  // 連泊（2泊目以降の泊数）
    // 超過時間加算（base+2）は calc に概念がない → 0 のまま
    add(base + 3, dayCount);     // 日帰
  });

  // コテージ超過人数加算（J=9）
  var sheetCount = computeSheetCount(s);
  if (sheetCount > 0) {
    add(9,  sheetCount * firstNight);   // 宿泊・人数
    add(10, sheetCount * extraNights);  // 連泊・人数
    add(12, sheetCount * dayCount);     // 日帰・人数
  }

  // 貸しテント（R=17, 3列構成）
  if (s.tentRental > 0) {
    add(17, s.tentRental * firstNight);
    add(18, s.tentRental * extraNights);
    add(19, s.tentRental * dayCount);
  }
  // 持込テント（U=20, 3列構成）
  if (s.tentBring > 0) {
    add(20, s.tentBring * firstNight);
    add(21, s.tentBring * extraNights);
    add(22, s.tentBring * dayCount);
  }
  // 使用人数加算（X=23, 3列構成）
  if (s.tentPeople > 0) {
    add(23, s.tentPeople * firstNight);
    add(24, s.tentPeople * extraNights);
    add(25, s.tentPeople * dayCount);
  }

  // 県内/県外の人数振り分け
  var totalPeople = (s.adults || 0) + (s.children || 0) + (s.infants || 0);
  var peopleInside = record.region === "inside" ? totalPeople : 0;
  var peopleOutside = record.region === "outside" ? totalPeople : 0;

  return {
    useDate: record.useDate,
    cells: cells,
    subtotal: record.subtotal || 0,
    discount: record.discount || 0,
    peopleInside: peopleInside,
    peopleOutside: peopleOutside,
    nameKana: record.nameKana || "",
    memo: record.memo || ""
  };
}

/* ---------- 月計表 Excel 生成（雛形テンプレ方式）----------
 * 雛形 tsukikei-template.xlsx を fetch → 値だけ書き込み → 出力。
 * 既存の罫線・書式・数式（AA小計, AC合計, 74行目の SUM）はそのまま温存される。
 *
 * データ行は雛形の 7〜73 行目（67件ぶん）。行6は AA6 数式が欠落しているため使わない。
 * 月あたり最大67件で警告。
 */
var TEMPLATE_URL = "./tsukikei-template.xlsx";
var DATA_ROW_START = 7;   // 1始まりの行番号
var DATA_ROW_END = 73;
var DATA_ROW_CAPACITY = DATA_ROW_END - DATA_ROW_START + 1;

function setCell(ws, col, rowNum, value, type) {
  // rowNum は 1始まり、col は 0始まり
  var addr = XLSX.utils.encode_cell({ r: rowNum - 1, c: col });
  if (ws[addr]) {
    // 既存セル：書式(.s)を保ちつつ値と型を更新、数式(.f)が無いことを前提に v を上書き
    ws[addr].t = type;
    ws[addr].v = value;
    if (type !== "n") delete ws[addr].w; // 表示テキストキャッシュをクリア
  } else {
    ws[addr] = { t: type, v: value };
  }
}

function clearCellValue(ws, col, rowNum) {
  var addr = XLSX.utils.encode_cell({ r: rowNum - 1, c: col });
  if (ws[addr] && !ws[addr].f) {
    // 数式が無いセルだけ値を消す。書式は残す
    delete ws[addr].v;
    delete ws[addr].w;
    ws[addr].t = "z"; // 空セル型
  }
}

function buildMonthlyWorkbook(ymKey, records) {
  if (typeof XLSX === "undefined") throw new Error("SheetJS (XLSX) が読み込まれていません");

  return fetch(TEMPLATE_URL)
    .then(function (res) {
      if (!res.ok) throw new Error("雛形ファイルが読み込めません (HTTP " + res.status + ")");
      return res.arrayBuffer();
    })
    .then(function (buf) {
      var wb = XLSX.read(buf, { type: "array", cellStyles: true, cellNF: true });

      var ym = ymKey.split("-");
      var yearNum = parseInt(ym[0], 10);
      var monthNum = parseInt(ym[1], 10);
      var reiwa = yearNum - 2018; // 2019→1, 2026→8

      var recsInMonth = records
        .filter(function (r) { return r.useDate && r.useDate.indexOf(ymKey) === 0; })
        .sort(function (a, b) { return a.useDate < b.useDate ? -1 : 1; });

      // --- シート③：月計表 ---
      var ws3 = wb.Sheets["③"] || wb.Sheets[wb.SheetNames[0]];
      fillMonthlyMatrix(ws3, reiwa, monthNum, recsInMonth);

      // --- シート④：利用状況一覧 ---
      var ws4 = wb.Sheets["④"] || wb.Sheets[wb.SheetNames[1]];
      if (ws4) fillStatusSheet(ws4, reiwa, monthNum, recsInMonth);

      return wb;
    });
}

function fillMonthlyMatrix(ws, reiwa, monthNum, recsInMonth) {
  var title = "令和" + reiwa + "年" + monthNum + "月分 キャンプ場使用料 月計表";
  setCell(ws, 0, 1, title, "s");

  var rowsInMonth = recsInMonth.map(recordToMatrixRow);
  if (rowsInMonth.length > DATA_ROW_CAPACITY) {
    throw new Error("1ヶ月の記録が雛形の容量(" + DATA_ROW_CAPACITY + "件)を超えています：" + rowsInMonth.length + "件");
  }

  // 既存の値をクリア（数式セル(AA, AC)は触らない）
  var writableCols = [0]; // A: 使用日
  for (var c = 1; c <= 25; c++) writableCols.push(c); // B-Z
  writableCols.push(27); // AB: 減免額
  writableCols.push(29); // AD: 県内人数
  writableCols.push(30); // AE: 県外人数
  writableCols.push(31); // AF: 名前
  for (var r = DATA_ROW_START; r <= DATA_ROW_END; r++) {
    writableCols.forEach(function (col) { clearCellValue(ws, col, r); });
  }

  rowsInMonth.forEach(function (row, i) {
    var rowNum = DATA_ROW_START + i;
    setCell(ws, 0, rowNum, row.useDate, "s");
    Object.keys(row.cells).forEach(function (k) {
      setCell(ws, parseInt(k, 10), rowNum, row.cells[k], "n");
    });
    if (row.discount > 0) setCell(ws, 27, rowNum, row.discount, "n");
    if (row.peopleInside > 0) setCell(ws, 29, rowNum, row.peopleInside, "n");
    if (row.peopleOutside > 0) setCell(ws, 30, rowNum, row.peopleOutside, "n");
    if (row.nameKana) setCell(ws, 31, rowNum, row.nameKana, "s"); // AG列に「様」あり
  });
}

/* ---------- シート④（利用状況一覧）の集計＆書き込み ---------- */
function fillStatusSheet(ws, reiwa, monthNum, recsInMonth) {
  // タイトル：「令和X年Y月分 利用状況一覧」（雛形は「年度」表記だが暦月で統一）
  setCell(ws, 0, 1, "令和" + reiwa + "年" + monthNum + "月分 利用状況一覧", "s");

  var summary = aggregateForStatusSheet(recsInMonth);

  // [行番号, 集計キー, 単価, 連泊適用列あり]
  var mapping = [
    [3,  "cottage6_over", 13200, true],
    [4,  "cottage6_day",  6600,  false],
    [5,  "cottage4_over", 8800,  true],
    [6,  "cottage4_day",  4400,  false],
    [7,  "addon_over",    1100,  true],
    [8,  "addon_day",     550,   false],
    [9,  "bungalow_over", 3150,  true],
    [10, "bungalow_day",  1580,  false],
    [11, "tentR_over",    880,   true],
    [12, "tentR_day",     440,   false],
    [13, "tentB_over",    660,   true],
    [14, "tentB_day",     330,   false],
    [15, "ppl_over",      220,   true],
    [16, "ppl_day",       110,   false]
  ];

  // 雛形④の列インデックス（0始まり）：G=6(利用数), I=8(金額), J=9(連泊数), L=11(連泊適用額), M=12(超過), O=14(超過適用額), P=15(小計)
  var sumI = 0, sumL = 0, sumO = 0;

  mapping.forEach(function (m) {
    var rowNum = m[0], key = m[1], price = m[2], hasExtra = m[3];
    var data = summary.agg[key] || { use: 0, extraNights: 0, overage: 0 };

    var amountI = data.use * price;                        // 1泊（または日帰）料金
    var amountL = Math.round(data.extraNights * price * 0.5); // 連泊（2泊目以降）半額
    var amountO = Math.round(data.overage * price * 0.5);     // 超過時間加算（calcに概念なし＝0）
    var subtotal = amountI + amountL + amountO;

    setCell(ws, 6, rowNum, data.use, "n");        // G: 利用数
    setCell(ws, 8, rowNum, amountI, "n");          // I: 金額
    if (hasExtra) {
      setCell(ws, 9,  rowNum, data.extraNights, "n");  // J: 連泊数
      setCell(ws, 11, rowNum, amountL, "n");            // L: 連泊適用額
      setCell(ws, 12, rowNum, data.overage, "n");       // M: 超過時間加算
      setCell(ws, 14, rowNum, amountO, "n");            // O: 超過適用額
    }
    setCell(ws, 15, rowNum, subtotal, "n");        // P: 小計

    sumI += amountI;
    sumL += amountL;
    sumO += amountO;
  });

  // 行17: 小計
  setCell(ws, 8,  17, sumI, "n");
  setCell(ws, 11, 17, sumL, "n");
  setCell(ws, 14, 17, sumO, "n");
  setCell(ws, 15, 17, sumI + sumL + sumO, "n");

  // 行18: 減免額（記録ごとの discount の合計）
  setCell(ws, 15, 18, summary.totalDiscount, "n");

  // 行19: 合計
  setCell(ws, 15, 19, sumI + sumL + sumO - summary.totalDiscount, "n");
}

function aggregateForStatusSheet(records) {
  var agg = {};
  var totalDiscount = 0;

  function bump(key, field, n) {
    if (!n) return;
    if (!agg[key]) agg[key] = { use: 0, extraNights: 0, overage: 0 };
    agg[key][field] += n;
  }

  records.forEach(function (r) {
    var s = r.state;
    var isOver = s.stay === "overnight";
    var nights = isOver ? Math.max(1, s.nights) : 1;
    var firstNight = isOver ? 1 : 0;
    var extra = isOver && nights > 1 ? nights - 1 : 0;
    var day = isOver ? 0 : 1;

    totalDiscount += (r.discount || 0);

    s.facilities.forEach(function (f) {
      var keyOver, keyDay;
      if (f === "cottageA") { keyOver = "cottage6_over"; keyDay = "cottage6_day"; }
      else if (f === "cottageB" || f === "cottageC" || f === "cottageD" || f === "cottageE") {
        keyOver = "cottage4_over"; keyDay = "cottage4_day";
      } else if (f === "bungalow1" || f === "bungalow2") {
        keyOver = "bungalow_over"; keyDay = "bungalow_day";
      } else return;
      bump(keyOver, "use", firstNight);
      bump(keyOver, "extraNights", extra);
      bump(keyDay, "use", day);
    });

    var sheetCount = computeSheetCount(s);
    if (sheetCount > 0) {
      bump("addon_over", "use", sheetCount * firstNight);
      bump("addon_over", "extraNights", sheetCount * extra);
      bump("addon_day", "use", sheetCount * day);
    }
    if (s.tentRental > 0) {
      bump("tentR_over", "use", s.tentRental * firstNight);
      bump("tentR_over", "extraNights", s.tentRental * extra);
      bump("tentR_day", "use", s.tentRental * day);
    }
    if (s.tentBring > 0) {
      bump("tentB_over", "use", s.tentBring * firstNight);
      bump("tentB_over", "extraNights", s.tentBring * extra);
      bump("tentB_day", "use", s.tentBring * day);
    }
    if (s.tentPeople > 0) {
      bump("ppl_over", "use", s.tentPeople * firstNight);
      bump("ppl_over", "extraNights", s.tentPeople * extra);
      bump("ppl_day", "use", s.tentPeople * day);
    }
  });

  return { agg: agg, totalDiscount: totalDiscount };
}

/* ---------- バックアップ：全件 JSON ダウンロード ---------- */
function exportAllJson() {
  var blob = new Blob([JSON.stringify(loadRecords(), null, 2)], { type: "application/json" });
  var url = URL.createObjectURL(blob);
  var a = document.createElement("a");
  a.href = url;
  a.download = "sugigadaira-records-" + new Date().toISOString().slice(0, 10) + ".json";
  a.click();
  URL.revokeObjectURL(url);
}

function importJson(file, mode, cb) {
  // mode: 'replace' | 'merge'
  var reader = new FileReader();
  reader.onload = function () {
    try {
      var arr = JSON.parse(reader.result);
      if (!Array.isArray(arr)) throw new Error("配列形式ではありません");
      if (mode === "replace") {
        saveRecords(arr);
      } else {
        var existing = loadRecords();
        var existingIds = {};
        existing.forEach(function (r) { existingIds[r.id] = true; });
        arr.forEach(function (r) { if (!existingIds[r.id]) existing.push(r); });
        saveRecords(existing);
      }
      cb(null, arr.length);
    } catch (e) {
      cb(e);
    }
  };
  reader.readAsText(file);
}

/* ---------- グローバル公開 ---------- */
window.SugiRecords = {
  STORAGE_KEY: STORAGE_KEY,
  PRICES: PRICES,
  FACILITY_LABEL: FACILITY_LABEL,
  calculate: calculate,
  loadRecords: loadRecords,
  saveRecords: saveRecords,
  addRecord: addRecord,
  deleteRecord: deleteRecord,
  newId: newId,
  recordToMatrixRow: recordToMatrixRow,
  buildMonthlyWorkbook: buildMonthlyWorkbook,
  aggregateForStatusSheet: aggregateForStatusSheet,
  exportAllJson: exportAllJson,
  importJson: importJson
};
