/**
 * GAS共同家計簿（Kakeibo）Webアプリ - サーバーサイド処理
 * (複数人対応・精算タイプ分類 ＆ 負担者マルチセレクト統合モデル)
 */

// ==========================================
// 【設定】スプレッドシートのIDを設定してください
// ==========================================
var SPREADSHEET_ID = ""; // ここにスプレッドシートIDを貼り付けるか、スクリプトプロパティで設定します。

var cachedSpreadsheet = null;

/**
 * スプレッドシートを取得するヘルパー関数
 */
function getSpreadsheet() {
  if (cachedSpreadsheet) {
    return cachedSpreadsheet;
  }
  var id = SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (!id) {
    throw new Error("スプレッドシートのIDが設定されていません。GAS設定画面のスクリプトプロパティに「SPREADSHEET_ID」を登録してください。");
  }
  try {
    cachedSpreadsheet = SpreadsheetApp.openById(id);
    return cachedSpreadsheet;
  } catch (e) {
    throw new Error("スプレッドシートを開くことができませんでした。権限が共有されているか、IDが正しいか確認してください。エラー: " + e.message);
  }
}

/**
 * Webアプリにアクセスがあった際にHTMLを表示する
 */
function doGet(e) {
  var dbTimestamp = "";
  try {
    checkAndInitializeSpreadsheet();
    try {
      var cache = CacheService.getScriptCache();
      if (cache) {
        cache.put("db_initialized_v2", "true", 21600); // 6時間キャッシュ
      }
    } catch (cacheErr) {
      console.warn("キャッシュ保存中にエラーが発生しました: " + cacheErr.message);
    }
    dbTimestamp = getDatabaseTimestamp();
  } catch (err) {
    console.warn("初期化中にエラーが発生しました: " + err.message);
  }

  var template = HtmlService.createTemplateFromFile('index');
  template.dbLastUpdated = dbTimestamp;

  return template.evaluate()
    .setTitle('共同家計簿 App')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 外部HTMLファイルを読み込むためのヘルパー（インクルード用）
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * スプレッドシートの構造をチェックし、存在しない場合は自動作成する
 */
function checkAndInitializeSpreadsheet() {
  var ss = getSpreadsheet();
  var isNewCreated = false;
  
  // 1. transactions シートの作成
  var txSheet = ss.getSheetByName("transactions");
  if (!txSheet) {
    isNewCreated = true;
    txSheet = ss.insertSheet("transactions");
    var headers = [
      "id",
      "date",
      "title",
      "category",
      "amount",
      "paid_by",
      "split_type",     // "割り勘", "立て替え", "個別"
      "beneficiaries",   // コマ区切りの負担メンバーリスト (例: "ユーザーB,ゲスト")
      "memo",
      "created_at"
    ];
    txSheet.appendRow(headers);
    txSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
  } else {
    // 既存スプレッドシートのマイグレーション（旧カラム of 復元・ヘッダー整合）
    var headers = txSheet.getRange(1, 1, 1, txSheet.getLastColumn()).getValues()[0];
    
    // コラムの有無を確認し、なければ追加
    if (headers.indexOf("split_type") === -1) {
      // 一時退避された旧カラムから再構築、またはヘッダーの修正
      txSheet.insertColumnAfter(6); // paid_byの後ろにカラム追加
      txSheet.getRange(1, 7).setValue("split_type");
    }
    
    // 2度目のチェックで確実にインデックスを取得
    headers = txSheet.getRange(1, 1, 1, txSheet.getLastColumn()).getValues()[0];
    var benIdx = headers.indexOf("beneficiaries");
    if (benIdx === -1) {
      txSheet.insertColumnAfter(7); // split_typeの後ろに追加
      txSheet.getRange(1, 8).setValue("beneficiaries");
    }
  }
  
  // 2. masters シートの作成
  var masterSheet = ss.getSheetByName("masters");
  if (!masterSheet) {
    isNewCreated = true;
    masterSheet = ss.insertSheet("masters");
    
    // ヘッダーの設定
    var headers = ["categories", "users", "split_types", "payment_methods", "settlement_statuses", "payment_patterns"];
    masterSheet.appendRow(headers);
    masterSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e2e2e2");
    
    // デフォルトマスター値の設定
    var defaultCategories = ["食費", "日用品・雑貨", "外食", "水道・光熱費・通信費", "娯楽・交際費", "雑費・その他", "家賃"];
    var defaultUsers = ["ユーザーA", "ユーザーB", "ゲスト"];
    var defaultSplitTypes = ["割り勘", "立て替え", "個別"];
    var defaultPaymentMethods = ["振込", "現金", "別の月で帳消し", "翌月の家賃支払いで帳消し"];
    var defaultSettlementStatuses = ["未払い", "支払い済み"];
    
    // 支払いパターンの初期組み合わせ生成
    var defaultPaymentPatterns = [];
    for (var i = 0; i < defaultUsers.length; i++) {
      for (var j = 0; j < defaultUsers.length; j++) {
        if (i !== j) {
          defaultPaymentPatterns.push(defaultUsers[i] + "から" + defaultUsers[j] + "へ");
        }
      }
    }
    
    var maxRows = Math.max(
      defaultCategories.length,
      defaultUsers.length,
      defaultSplitTypes.length,
      defaultPaymentMethods.length,
      defaultSettlementStatuses.length,
      defaultPaymentPatterns.length
    );
    
    var masterData = [];
    for (var i = 0; i < maxRows; i++) {
      masterData.push([
        defaultCategories[i] || "",
        defaultUsers[i] || "",
        defaultSplitTypes[i] || "",
        defaultPaymentMethods[i] || "",
        defaultSettlementStatuses[i] || "",
        defaultPaymentPatterns[i] || ""
      ]);
    }
    masterSheet.getRange(2, 1, masterData.length, 6).setValues(masterData);
  } else {
    // 既存マスターシートの列追加マイグレーション
    var mLC = masterSheet.getLastColumn();
    var mHeaders = [];
    if (mLC > 0) {
      mHeaders = masterSheet.getRange(1, 1, 1, mLC).getValues()[0];
    }
    
    // payment_methods
    if (mHeaders.indexOf("payment_methods") === -1) {
      masterSheet.getRange(1, 4).setValue("payment_methods");
      masterSheet.getRange(1, 4).setFontWeight("bold").setBackground("#e2e2e2");
      var defaultPaymentMethods = ["振込", "現金", "別の月で帳消し", "翌月の家賃支払いで帳消し"];
      for (var i = 0; i < defaultPaymentMethods.length; i++) {
        masterSheet.getRange(2 + i, 4).setValue(defaultPaymentMethods[i]);
      }
    }
    
    // settlement_statuses
    if (mHeaders.indexOf("settlement_statuses") === -1) {
      masterSheet.getRange(1, 5).setValue("settlement_statuses");
      masterSheet.getRange(1, 5).setFontWeight("bold").setBackground("#e2e2e2");
      var defaultSettlementStatuses = ["未払い", "支払い済み"];
      for (var i = 0; i < defaultSettlementStatuses.length; i++) {
        masterSheet.getRange(2 + i, 5).setValue(defaultSettlementStatuses[i]);
      }
    }
    
    // payment_patterns
    if (mHeaders.indexOf("payment_patterns") === -1) {
      masterSheet.getRange(1, 6).setValue("payment_patterns");
      masterSheet.getRange(1, 6).setFontWeight("bold").setBackground("#e2e2e2");
      
      // 既存のusersをロードしてペアワイズ組み合わせを自動生成
      var mLR = masterSheet.getLastRow();
      var mValues = mLR > 1 ? masterSheet.getRange(2, 2, mLR - 1, 1).getValues() : [];
      var existingUsers = [];
      for (var r = 0; r < mValues.length; r++) {
        if (mValues[r][0]) existingUsers.push(mValues[r][0]);
      }
      if (existingUsers.length === 0) {
        existingUsers = ["ユーザーA", "ユーザーB", "ゲスト"];
      }
      var defaultPaymentPatterns = [];
      for (var i = 0; i < existingUsers.length; i++) {
        for (var j = 0; j < existingUsers.length; j++) {
          if (i !== j) {
            defaultPaymentPatterns.push(existingUsers[i] + "から" + existingUsers[j] + "へ");
          }
        }
      }
      for (var i = 0; i < defaultPaymentPatterns.length; i++) {
        masterSheet.getRange(2 + i, 6).setValue(defaultPaymentPatterns[i]);
      }
    }
  }
  
  // 3. settlements シートの作成
  var settleSheet = ss.getSheetByName("settlements");
  if (!settleSheet) {
    isNewCreated = true;
    settleSheet = ss.insertSheet("settlements");
    var headers = [
      "id",               // ユニークID
      "month",            // YYYY-MM
      "status",           // "未払い", "支払い済み" など
      "payment_date",     // YYYY-MM-DD
      "payment_pattern",  // "ユーザーAからユーザーBへ" など
      "amount",           // 金額
      "payment_method",   // "振込", "現金" など
      "memo",
      "updated_at"
    ];
    settleSheet.appendRow(headers);
    settleSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
  } else {
    // settlements シートのマイグレーション（id カラムの追加）
    var sLC = settleSheet.getLastColumn();
    var sHeaders = [];
    if (sLC > 0) {
      sHeaders = settleSheet.getRange(1, 1, 1, sLC).getValues()[0];
    }
    if (sHeaders.indexOf("id") === -1) {
      settleSheet.insertColumnBefore(1);
      settleSheet.getRange(1, 1).setValue("id");
      settleSheet.getRange(1, 1).setFontWeight("bold").setBackground("#f3f3f3");
      
      var sLR = settleSheet.getLastRow();
      if (sLR > 1) {
        var uuids = [];
        for (var i = 0; i < sLR - 1; i++) {
          uuids.push([Utilities.getUuid()]);
        }
        settleSheet.getRange(2, 1, sLR - 1, 1).setValues(uuids);
      }
    }
  }

  if (isNewCreated) {
    updateDatabaseTimestamp();
  }
}

/**
 * マスタ設定データと、指定された月の家計簿データを一括取得する
 */
function getAppData(monthStr) {
  var step = "開始";
  try {
    step = "スプレッドシートオープン";
    var ss = getSpreadsheet();
    
    step = "キャッシュチェック";
    var cache = null;
    try {
      cache = CacheService.getScriptCache();
    } catch (e) {
      console.warn("CacheService の取得に失敗しました: " + e.message);
    }
    
    var checked = false;
    if (cache) {
      try {
        checked = cache.get("db_initialized_v2") === "true";
      } catch (e) {
        console.warn("キャッシュの取得に失敗しました: " + e.message);
      }
    }
    
    if (!checked) {
      step = "スプレッドシート構造チェック・初期化";
      checkAndInitializeSpreadsheet();
      if (cache) {
        try {
          cache.put("db_initialized_v2", "true", 21600); // 6時間キャッシュ
        } catch (e) {
          console.warn("キャッシュの書き込みに失敗しました: " + e.message);
        }
      }
    }
    
    step = "シート取得（transactions/masters）";
    var txSheet = ss.getSheetByName("transactions");
    var masterSheet = ss.getSheetByName("masters");
    
    step = "マスターデータ読み込み";
    var mLR = masterSheet.getLastRow();
    var mLC = masterSheet.getLastColumn();
    var masterValues = (mLR > 0 && mLC > 0) ? masterSheet.getRange(1, 1, mLR, mLC).getValues() : [[]];
    var categories = [];
    var users = [];
    var splitTypes = [];
    var paymentMethods = [];
    var settlementStatuses = [];
    var paymentPatterns = [];
    
    for (var r = 1; r < masterValues.length; r++) {
      if (masterValues[r][0]) categories.push(masterValues[r][0]);
      if (masterValues[r][1]) users.push(masterValues[r][1]);
      if (masterValues[r][2]) splitTypes.push(masterValues[r][2]);
      if (masterValues[r][3]) paymentMethods.push(masterValues[r][3]);
      if (masterValues[r][4]) settlementStatuses.push(masterValues[r][4]);
      if (masterValues[r][5]) paymentPatterns.push(masterValues[r][5]);
    }
    
    step = "取引明細データ読み込み";
    var txLR = txSheet.getLastRow();
    var txLC = txSheet.getLastColumn();
    var txValues = (txLR > 0 && txLC > 0) ? txSheet.getRange(1, 1, txLR, txLC).getValues() : [[]];
    var headers = txValues[0];
    
    var colIndexes = {};
    headers.forEach(function(h, idx) {
      colIndexes[h] = idx;
    });
    
    step = "取引明細データパース・クリーンアップ";
    var rawTransactions = [];
    for (var i = 1; i < txValues.length; i++) {
      var row = txValues[i];
      var tx = {};
      headers.forEach(function(h) {
        var val = row[colIndexes[h]];
        if (h === "date" && val instanceof Date) {
          if (!isNaN(val.getTime())) {
            tx[h] = Utilities.formatDate(val, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
          } else {
            tx[h] = "";
          }
        } else if (h === "created_at" && val instanceof Date) {
          if (!isNaN(val.getTime())) {
            tx[h] = val.toISOString();
          } else {
            tx[h] = "";
          }
        } else {
          var cleanVal = val;
          if (cleanVal === undefined) {
            cleanVal = null;
          } else if (cleanVal instanceof Date) {
            cleanVal = (!isNaN(cleanVal.getTime())) ? cleanVal.toISOString() : "";
          }
          if (h === "amount") {
            cleanVal = Number(val) || 0;
          } else if (typeof cleanVal === "number" && (isNaN(cleanVal) || !isFinite(cleanVal))) {
            cleanVal = 0;
          }
          tx[h] = cleanVal;
        }
      });
      
      // 旧バージョンのバグによるカラムズレの補正（memoにcreatedAtが入っていて、created_atが空の場合）
      if (!tx.created_at && tx.memo) {
        var memoStr = String(tx.memo);
        // ISO 8601 または yyyy/MM/dd 形式のタイムスタンプかどうかチェック
        if (/^\d{4}[-/]\d{2}[-/]\d{2}/.test(memoStr)) {
          tx.created_at = memoStr;
          tx.memo = "";
        }
      }
      
      rawTransactions.push(tx);
    }
    
    step = "月次フィルタリング";
    var filteredTransactions = rawTransactions.filter(function(tx) {
      if (!tx.date) return false;
      return tx.date.substring(0, 7) === monthStr;
    });
    
    step = "日付ソート";
    filteredTransactions.sort(function(a, b) {
      var dateA = a.date || "";
      var dateB = b.date || "";
      var createdA = a.created_at || "";
      var createdB = b.created_at || "";
      return dateB.localeCompare(dateA) || createdB.localeCompare(createdA);
    });
    
    step = "精算計算処理";
    var analysis = calculateSettlements(filteredTransactions, users);
    
    step = "精算ステータスデータ読み込み";
    var settleSheet = ss.getSheetByName("settlements");
    var settlementRecords = [];
    if (settleSheet) {
      var sLR = settleSheet.getLastRow();
      var sLC = settleSheet.getLastColumn();
      var sValues = (sLR > 0 && sLC > 0) ? settleSheet.getRange(1, 1, sLR, sLC).getValues() : [[]];
      var sHeaders = sValues[0];
      var sColIndexes = {};
      sHeaders.forEach(function(h, idx) {
        sColIndexes[h] = idx;
      });
      
      var tz = ss.getSpreadsheetTimeZone();
      for (var r = 1; r < sValues.length; r++) {
        var sRow = sValues[r];
        var cellMonth = convertMonthToString(sRow[sColIndexes["month"]], tz);
        if (cellMonth === monthStr) {
          var record = {};
          sHeaders.forEach(function(h) {
            var val = sRow[sColIndexes[h]];
            if (h === "month") {
              record[h] = convertMonthToString(val, tz);
            } else if (h === "payment_date") {
              record[h] = convertDateStringToISO(val, tz);
            } else if (h === "updated_at") {
              if (val instanceof Date) {
                record[h] = (!isNaN(val.getTime())) ? val.toISOString() : "";
              } else {
                record[h] = val ? String(val) : "";
              }
            } else if (h === "amount") {
              record[h] = Number(val) || 0;
            } else {
              if (val instanceof Date) {
                record[h] = (!isNaN(val.getTime())) ? Utilities.formatDate(val, tz, "yyyy-MM-dd HH:mm:ss") : "";
              } else if (val === undefined || val === null) {
                record[h] = "";
              } else {
                record[h] = String(val);
              }
            }
          });
          settlementRecords.push(record);
        }
      }
    }
    
    step = "データ返却用JSONオブジェクト構築";
    return {
      success: true,
      month: monthStr,
      transactions: filteredTransactions,
      masters: {
        categories: categories,
        users: users,
        split_types: splitTypes,
        payment_methods: paymentMethods,
        settlement_statuses: settlementStatuses,
        payment_patterns: paymentPatterns
      },
      analysis: analysis,
      settlementRecords: settlementRecords,
      hasGeminiKey: !!getApiKey(),
      dbLastUpdated: getDatabaseTimestamp(ss)
    };
  } catch (err) {
    return {
      success: false,
      message: "ステップ「" + step + "」でエラーが発生しました: " + err.message + "\n" + err.stack
    };
  }
}

/**
 * 複数人に対応した精算およびカテゴリ別集計の計算ロジック（タイプ分類 ＆ 負担メンバーマルチセレクトモデル）
 */
function calculateSettlements(transactions, users) {
  var paidMap = {};
  var usedMap = {};
  users.forEach(function(u) {
    paidMap[u] = 0;
    usedMap[u] = 0;
  });
  
  var categoryTotals = {};
  var totalExpense = 0;
  
  transactions.forEach(function(tx) {
    var amount = Number(tx.amount) || 0;
    var payer = tx.paid_by;
    var splitType = tx.split_type || "割り勘";
    var beneficiariesStr = tx.beneficiaries || "";
    var cat = tx.category || "その他";
    
    // カテゴリ別合計
    categoryTotals[cat] = (categoryTotals[cat] || 0) + amount;
    totalExpense += amount;
    
    // 存在しないユーザーが支払者の場合はスキップ
    if (users.indexOf(payer) === -1) return;
    
    // 1. 支払額を加算
    paidMap[payer] += amount;
    
    // 2. 負担対象メンバーの配列を決定
    var targets = [];
    if (splitType === "個別") {
      // 個別の場合は支払者本人のみ負担
      targets = [payer];
    } else {
      // 割り勘・立て替えの場合は、カンマ区切り文字列をパース
      targets = beneficiariesStr ? beneficiariesStr.split(',') : [];
      if (targets.length === 0) {
        if (splitType === "割り勘") {
          targets = users; // デフォルトは全員
        } else if (splitType === "立て替え") {
          targets = users.filter(function(u) { return u !== payer; }); // デフォルトは自分以外
        }
      }
    }
    
    // 3. 負担対象者の使用額に加算
    var share = targets.length > 0 ? (amount / targets.length) : 0;
    targets.forEach(function(t) {
      if (users.indexOf(t) !== -1) {
        usedMap[t] += share;
      }
    });
  });
  
  // 各自の収支差額（バランス = 支払額 - 使用額）を計算
  var balances = [];
  var balanceMap = {};
  users.forEach(function(u) {
    var bal = paidMap[u] - usedMap[u];
    balanceMap[u] = bal;
    balances.push({
      user: u,
      paid: Math.round(paidMap[u]),
      used: Math.round(usedMap[u]),
      balance: Math.round(bal)
    });
  });
  
  // 精算取引の導出
  var settlements = [];
  var debts = [];
  var creditors = [];
  
  users.forEach(function(u) {
    var b = balanceMap[u];
    if (b < -0.1) {
      debts.push({ user: u, amount: -b });
    } else if (b > 0.1) {
      creditors.push({ user: u, amount: b });
    }
  });
  
  debts.sort(function(a, b) { return b.amount - a.amount; });
  creditors.sort(function(a, b) { return b.amount - a.amount; });
  
  var dIdx = 0, cIdx = 0;
  while (dIdx < debts.length && cIdx < creditors.length) {
    var d = debts[dIdx];
    var c = creditors[cIdx];
    
    var settleAmount = Math.min(d.amount, c.amount);
    if (settleAmount > 0.5) {
      settlements.push({
        from: d.user,
        to: c.user,
        amount: Math.round(settleAmount)
      });
    }
    
    d.amount -= settleAmount;
    c.amount -= settleAmount;
    
    if (d.amount < 0.1) dIdx++;
    if (c.amount < 0.1) cIdx++;
  }
  
  return {
    totalExpense: totalExpense,
    categoryTotals: categoryTotals,
    userBalances: balances,
    settlementInstructions: settlements
  };
}

/**
 * 新しい取引を追加する
 */
function addTransaction(tx) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("transactions");
    
    var id = Utilities.getUuid();
    var createdAt = new Date();
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    var txData = {
      id: id,
      date: tx.date,
      title: tx.title,
      category: tx.category,
      amount: Number(tx.amount) || 0,
      paid_by: tx.paid_by,
      split_type: tx.split_type,
      beneficiaries: tx.beneficiaries || "",
      memo: tx.memo || "",
      created_at: createdAt
    };
    
    var rowData = headers.map(function(h) {
      return txData.hasOwnProperty(h) ? txData[h] : "";
    });
    
    sheet.appendRow(rowData);
    
    // データベースの更新タイムスタンプを更新
    updateDatabaseTimestamp();
    
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * 既存の取引データを更新する
 */
function updateTransaction(tx) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("transactions");
    var lr = sheet.getLastRow();
    var lc = sheet.getLastColumn();
    var values = (lr > 0 && lc > 0) ? sheet.getRange(1, 1, lr, lc).getValues() : [[]];
    var headers = values[0];
    
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === tx.id) {
        var rowNum = i + 1;
        var createdAtIdx = headers.indexOf("created_at");
        var originalCreatedAt = createdAtIdx !== -1 ? values[i][createdAtIdx] : new Date();
        
        var originalRow = values[i];
        var txData = {
          id: tx.id,
          date: tx.date,
          title: tx.title,
          category: tx.category,
          amount: Number(tx.amount) || 0,
          paid_by: tx.paid_by,
          split_type: tx.split_type,
          beneficiaries: tx.beneficiaries || "",
          memo: tx.memo || "",
          created_at: originalCreatedAt
        };
        
        var rowData = headers.map(function(h, idx) {
          return txData.hasOwnProperty(h) ? txData[h] : originalRow[idx];
        });
        
        sheet.getRange(rowNum, 1, 1, rowData.length).setValues([rowData]);
        
        // データベースの更新タイムスタンプを更新
        updateDatabaseTimestamp();
        
        return { success: true };
      }
    }
    return { success: false, message: "更新対象の取引が見つかりませんでした。" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * 指定したIDの取引を削除する
 */
function deleteTransaction(id) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("transactions");
    var lr = sheet.getLastRow();
    var lc = sheet.getLastColumn();
    var values = (lr > 0 && lc > 0) ? sheet.getRange(1, 1, lr, lc).getValues() : [[]];
    
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        sheet.deleteRow(i + 1);
        
        // データベースの更新タイムスタンプを更新
        updateDatabaseTimestamp();
        
        return { success: true };
      }
    }
    return { success: false, message: "指定された取引が見つかりませんでした。" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * 指定された年の家計簿データを集計する
 */
function getYearlyData(yearStr) {
  try {
    var ss = getSpreadsheet();
    
    var cache = null;
    try {
      cache = CacheService.getScriptCache();
    } catch (e) {
      console.warn("CacheService の取得に失敗しました: " + e.message);
    }
    
    var checked = false;
    if (cache) {
      try {
        checked = cache.get("db_initialized_v2") === "true";
      } catch (e) {
        console.warn("キャッシュの取得に失敗しました: " + e.message);
      }
    }
    
    if (!checked) {
      checkAndInitializeSpreadsheet();
      if (cache) {
        try {
          cache.put("db_initialized_v2", "true", 21600); // 6時間キャッシュ
        } catch (e) {
          console.warn("キャッシュの書き込みに失敗しました: " + e.message);
        }
      }
    }
    
    var txSheet = ss.getSheetByName("transactions");
    var masterSheet = ss.getSheetByName("masters");
    
    // 1. マスタデータ取得
    var mLR = masterSheet.getLastRow();
    var mLC = masterSheet.getLastColumn();
    var masterValues = (mLR > 0 && mLC > 0) ? masterSheet.getRange(1, 1, mLR, mLC).getValues() : [[]];
    var categories = [];
    var users = [];
    for (var r = 1; r < masterValues.length; r++) {
      if (masterValues[r][0]) categories.push(masterValues[r][0]);
      if (masterValues[r][1]) users.push(masterValues[r][1]);
    }
    
    // 2. 全取引データ取得
    var txLR = txSheet.getLastRow();
    var txLC = txSheet.getLastColumn();
    var txValues = (txLR > 0 && txLC > 0) ? txSheet.getRange(1, 1, txLR, txLC).getValues() : [[]];
    var headers = txValues[0];
    
    var colIndexes = {};
    headers.forEach(function(h, idx) {
      colIndexes[h] = idx;
    });
    
    var rawTransactions = [];
    for (var i = 1; i < txValues.length; i++) {
      var row = txValues[i];
      var tx = {};
      headers.forEach(function(h) {
        var val = row[colIndexes[h]];
        if (h === "date" && val instanceof Date) {
          if (!isNaN(val.getTime())) {
            tx[h] = Utilities.formatDate(val, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
          } else {
            tx[h] = "";
          }
        } else if (h === "created_at" && val instanceof Date) {
          if (!isNaN(val.getTime())) {
            tx[h] = val.toISOString();
          } else {
            tx[h] = "";
          }
        } else {
          var cleanVal = val;
          if (cleanVal === undefined) {
            cleanVal = null;
          } else if (cleanVal instanceof Date) {
            cleanVal = (!isNaN(cleanVal.getTime())) ? cleanVal.toISOString() : "";
          }
          if (h === "amount") {
            cleanVal = Number(val) || 0;
          } else if (typeof cleanVal === "number" && (isNaN(cleanVal) || !isFinite(cleanVal))) {
            cleanVal = 0;
          }
          tx[h] = cleanVal;
        }
      });
      
      // 旧バージョンのバグによるカラムズレの補正（memoにcreatedAtが入っていて、created_atが空の場合）
      if (!tx.created_at && tx.memo) {
        var memoStr = String(tx.memo);
        // ISO 8601 または yyyy/MM/dd 形式のタイムスタンプかどうかチェック
        if (/^\d{4}[-/]\d{2}[-/]\d{2}/.test(memoStr)) {
          tx.created_at = memoStr;
          tx.memo = "";
        }
      }
      
      rawTransactions.push(tx);
    }
    
    // 当年のデータのみフィルタリング
    var yearTransactions = rawTransactions.filter(function(tx) {
      if (!tx.date) return false;
      return tx.date.substring(0, 4) === yearStr;
    });
    
    // 3. 月別推移（1月〜12月）の集計
    var months = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"];
    var monthlyData = months.map(function(m) {
      var monthStr = yearStr + "-" + m;
      var filtered = yearTransactions.filter(function(tx) {
        return tx.date.substring(0, 7) === monthStr;
      });
      
      var total = 0;
      var paidByMap = {};
      users.forEach(function(u) { paidByMap[u] = 0; });
      
      filtered.forEach(function(tx) {
        var amount = Number(tx.amount) || 0;
        total += amount;
        if (paidByMap.hasOwnProperty(tx.paid_by)) {
          paidByMap[tx.paid_by] += amount;
        }
      });
      
      return {
        month: Number(m) + "月",
        total: total,
        paidBy: paidByMap
      };
    });
    
    // 4. カテゴリ別の年間合計
    var categoryTotals = {};
    yearTransactions.forEach(function(tx) {
      var cat = tx.category || "その他";
      var amount = Number(tx.amount) || 0;
      categoryTotals[cat] = (categoryTotals[cat] || 0) + amount;
    });
    
    // 5. 個人別の年間累計バランス（精算ロジック流用）
    var analysis = calculateSettlements(yearTransactions, users);
    
    return {
      success: true,
      year: yearStr,
      monthlyData: monthlyData,
      categoryTotals: categoryTotals,
      analysis: analysis,
      users: users
    };
  } catch (err) {
    return {
      success: false,
      message: err.message
    };
  }
}

// ==========================================
// Gemini API 連携処理
// ==========================================

function getApiKey() {
  var userKey = PropertiesService.getUserProperties().getProperty("GEMINI_API_KEY");
  if (userKey) return userKey;
  
  return PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
}

function saveGeminiApiKey(key) {
  try {
    if (!key) {
      PropertiesService.getUserProperties().deleteProperty("GEMINI_API_KEY");
    } else {
      PropertiesService.getUserProperties().setProperty("GEMINI_API_KEY", key);
    }
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function analyzeReceipt(base64Data, mimeType, clientApiKey) {
  try {
    var apiKey = clientApiKey || getApiKey();
    if (!apiKey) {
      return { 
        success: false, 
        message: "Gemini APIキーが設定されていません。右上の設定マークからAPIキーを入力してください。" 
      };
    }
    
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName("masters");
    var mLR = masterSheet.getLastRow();
    var mLC = masterSheet.getLastColumn();
    var masterValues = (mLR > 0 && mLC > 0) ? masterSheet.getRange(1, 1, mLR, mLC).getValues() : [[]];
    var categories = [];
    for (var r = 1; r < masterValues.length; r++) {
      if (masterValues[r][0]) categories.push(masterValues[r][0]);
    }
    var categoriesListStr = categories.join(", ");

    var model = "gemini-3.1-pro-preview"; // 精度が高いモデル
    var url = "https://generativelanguage.googleapis.com/v1beta/models/" + model + ":generateContent?key=" + apiKey;
    
    var today = new Date();
    var todayStr = Utilities.formatDate(today, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    
    var prompt = 
      "あなたは有能な家計簿入力アシスタントです。添付されたファイル（画像またはPDF。PDFの場合は複数ページにまたがる可能性があります）から、1枚または複数枚のすべてのレシート情報を読み取り、以下のJSON配列形式で出力してください。\n" +
      "```json\n" +
      "[\n" +
      "  {\n" +
      "    \"date\": \"YYYY-MM-DD形式の購入日。レシートから読み取れない場合は今日の日付(" + todayStr + ")にしてください\",\n" +
      "    \"title\": \"店舗名または会社名。短くわかりやすい名称にしてください\",\n" +
      "    \"amount\": 合計支払金額（カンマなしの数値データ）,\n" +
      "    \"category\": \"次のリストの中から最も適したカテゴリを1つ選んでください: [" + categoriesListStr + "]\"\n" +
      "  }\n" +
      "]\n" +
      "```\n" +
      "【制約ルール】\n" +
      "- 出力はプレーンなJSON配列の文字列だけを返してください。マークダウンの ```json ... ``` や前後の説明文は一切含めないでください。\n" +
      "- ファイル内に複数のレシート（または複数ページ）がある場合は、検出されたすべてのレシート分を配列オブジェクトとして出力してください。\n" +
      "- カテゴリは必ず指定されたリストの中から完全に一致するものを選んでください。どれにも当てはまらない場合は「その他」または「雑費・その他」に分類してください。";
    
    var payload = {
      contents: [{
        parts: [
          { text: prompt },
          {
            inline_data: {
              mime_type: mimeType,
              data: base64Data
            }
          }
        ]
      }]
    };
    
    var options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    
    // 429(Too Many Requests) またはエラーの場合は、軽量なFlashモデルにフォールバック
    if (responseCode === 429 || responseCode >= 500) {
      console.warn(model + " で制限(またはエラー)が発生したため、gemini-3.5-flash で再試行します。コード: " + responseCode);
      var fallbackModel = "gemini-3.5-flash";
      var fallbackUrl = "https://generativelanguage.googleapis.com/v1beta/models/" + fallbackModel + ":generateContent?key=" + apiKey;
      response = UrlFetchApp.fetch(fallbackUrl, options);
      responseCode = response.getResponseCode();
      responseText = response.getContentText();
    }
    
    if (responseCode !== 200) {
      return { 
        success: false, 
        message: "Gemini APIの呼び出しに失敗しました。ステータスコード: " + responseCode + ", レスポンス: " + responseText 
      };
    }
    
    var resultJson = JSON.parse(responseText);
    var rawText = resultJson.candidates[0].content.parts[0].text;
    
    var cleanJsonText = rawText.replace(/```json/g, "").replace(/```/g, "").trim();
    var extractedData = JSON.parse(cleanJsonText);
    
    // 配列でない場合は配列化する
    var dataList = Array.isArray(extractedData) ? extractedData : [extractedData];
    var formattedData = dataList.map(function(item) {
      return {
        date: item.date || "",
        title: item.title || "",
        amount: Number(item.amount) || 0,
        category: item.category || "その他"
      };
    });
    
    return {
      success: true,
      data: formattedData
    };
    
  } catch (err) {
    return {
      success: false,
      message: "レシート解析中にエラーが発生しました: " + err.message
    };
  }
}

/**
 * 複数のレシート画像を1回のAPIコールで一括解析する
 * @param {Array} images - [{base64Data: string, mimeType: string}, ...]
 * @param {string} [clientApiKey] - クライアント側から渡されるAPIキー
 * @return {Object} {success: boolean, data: Array}
 */
function analyzeReceipts(images, clientApiKey) {
  try {
    var apiKey = clientApiKey || getApiKey();
    if (!apiKey) {
      return { 
        success: false, 
        message: "Gemini APIキーが設定されていません。右上の設定マークからAPIキーを入力してください。" 
      };
    }
    
    var ss = getSpreadsheet();
    var masterSheet = ss.getSheetByName("masters");
    var mLR = masterSheet.getLastRow();
    var mLC = masterSheet.getLastColumn();
    var masterValues = (mLR > 0 && mLC > 0) ? masterSheet.getRange(1, 1, mLR, mLC).getValues() : [[]];
    var categories = [];
    for (var r = 1; r < masterValues.length; r++) {
      if (masterValues[r][0]) categories.push(masterValues[r][0]);
    }
    var categoriesListStr = categories.join(", ");
    
    var imageCount = images.length;
    
    var today = new Date();
    var todayStr = Utilities.formatDate(today, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");

    var prompt = 
      "あなたは有能な家計簿入力アシスタントです。添付された " + imageCount + " 枚のレシート画像それぞれから必要な情報を読み取り、以下のJSON配列形式で出力してください。\n" +
      "画像の順番通りに、1枚につき1つのオブジェクトを配列に含めてください。\n" +
      "```json\n" +
      "[\n" +
      "  {\n" +
      "    \"date\": \"YYYY-MM-DD形式の購入日。レシートから読み取れない場合は今日の日付(" + todayStr + ")にしてください\",\n" +
      "    \"title\": \"店舗名または会社名。短くわかりやすい名称にしてください\",\n" +
      "    \"amount\": 合計支払金額（カンマなしの数値データ）,\n" +
      "    \"category\": \"次のリストの中から最も適したカテゴリを1つ選んでください: [" + categoriesListStr + "]\"\n" +
      "  }\n" +
      "]\n" +
      "```\n" +
      "【制約ルール】\n" +
      "- 出力はプレーンなJSON配列の文字列だけを返してください。マークダウンの ```json ... ``` や前後の説明文は一切含めないでください。\n" +
      "- 必ず " + imageCount + " 個のオブジェクトを含む配列を返してください。\n" +
      "- カテゴリは必ず指定されたリストの中から完全に一致するものを選んでください。どれにも当てはまらない場合は「その他」または「雑費・その他」に分類してください。";
    
    // partsの構築: テキストプロンプト + 各画像
    var parts = [{ text: prompt }];
    for (var i = 0; i < images.length; i++) {
      parts.push({
        inline_data: {
          mime_type: images[i].mimeType,
          data: images[i].base64Data
        }
      });
    }
    
    var model = "gemini-3.1-pro-preview"; // 精度が高いモデル
    var url = "https://generativelanguage.googleapis.com/v1beta/models/" + model + ":generateContent?key=" + apiKey;
    
    var payload = {
      contents: [{ parts: parts }]
    };
    
    var options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    
    // 429(Too Many Requests) またはエラーの場合は、軽量なFlashモデルにフォールバック
    if (responseCode === 429 || responseCode >= 500) {
      console.warn(model + " で制限(またはエラー)が発生したため、gemini-3.5-flash で再試行します。コード: " + responseCode);
      var fallbackModel = "gemini-3.5-flash";
      var fallbackUrl = "https://generativelanguage.googleapis.com/v1beta/models/" + fallbackModel + ":generateContent?key=" + apiKey;
      response = UrlFetchApp.fetch(fallbackUrl, options);
      responseCode = response.getResponseCode();
      responseText = response.getContentText();
    }
    
    if (responseCode !== 200) {
      return { 
        success: false, 
        message: "Gemini APIの呼び出しに失敗しました。ステータスコード: " + responseCode + ", レスポンス: " + responseText 
      };
    }
    
    var resultJson = JSON.parse(responseText);
    var rawText = resultJson.candidates[0].content.parts[0].text;
    
    var cleanJsonText = rawText.replace(/```json/g, "").replace(/```/g, "").trim();
    var extractedArray = JSON.parse(cleanJsonText);
    
    // 配列でない場合（1枚だけの場合にオブジェクトで返るケース）は配列に包む
    if (!Array.isArray(extractedArray)) {
      extractedArray = [extractedArray];
    }
    
    var results = extractedArray.map(function(item) {
      return {
        date: item.date || todayStr,
        title: item.title || "",
        amount: Number(item.amount) || 0,
        category: item.category || "その他"
      };
    });
    
    return {
      success: true,
      data: results
    };
    
  } catch (err) {
    return {
      success: false,
      message: "レシート一括解析中にエラーが発生しました: " + err.message
    };
  }
}

/**
 * 複数の取引を一括追加する（CSVインポート用）
 */
function importTransactions(txList) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("transactions");
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    var rowsData = [];
    var createdAt = new Date();
    
    txList.forEach(function(tx, idx) {
      // 1ミリ秒ずつずらしてIDと作成日時を付与
      var id = Utilities.getUuid();
      var txCreatedAt = new Date(createdAt.getTime() + idx);
      
      var txData = {
        id: id,
        date: tx.date,
        title: tx.title,
        category: tx.category,
        amount: Number(tx.amount) || 0,
        paid_by: tx.paid_by,
        split_type: tx.split_type,
        beneficiaries: tx.beneficiaries || "",
        memo: tx.memo || "",
        created_at: txCreatedAt
      };
      
      var row = headers.map(function(h) {
        return txData.hasOwnProperty(h) ? txData[h] : "";
      });
      rowsData.push(row);
    });
    
    // 1. 取引明細シートへ一括書き込み
    if (rowsData.length > 0) {
      var nextRow = sheet.getLastRow() + 1;
      sheet.getRange(nextRow, 1, rowsData.length, headers.length).setValues(rowsData);
    }
    
    // 2. マスターデータの同期（マスターシートにないユーザーやカテゴリーがあれば追加する）
    var masterSheet = ss.getSheetByName("masters");
    var mLR = masterSheet.getLastRow();
    var mLC = masterSheet.getLastColumn();
    var masterValues = (mLR > 0 && mLC > 0) ? masterSheet.getRange(1, 1, mLR, mLC).getValues() : [[]];
    
    var existingCategories = [];
    var existingUsers = [];
    for (var r = 1; r < masterValues.length; r++) {
      if (masterValues[r][0]) existingCategories.push(masterValues[r][0]);
      if (masterValues[r][1]) existingUsers.push(masterValues[r][1]);
    }
    
    var newCategories = [];
    var newUsers = [];
    txList.forEach(function(tx) {
      if (tx.category && existingCategories.indexOf(tx.category) === -1 && newCategories.indexOf(tx.category) === -1) {
        newCategories.push(tx.category);
      }
      if (tx.paid_by && existingUsers.indexOf(tx.paid_by) === -1 && newUsers.indexOf(tx.paid_by) === -1) {
        newUsers.push(tx.paid_by);
      }
      // 負担メンバーもマスター照合
      if (tx.beneficiaries) {
        tx.beneficiaries.split(',').forEach(function(u) {
          var trimmedU = u.trim();
          if (trimmedU && existingUsers.indexOf(trimmedU) === -1 && newUsers.indexOf(trimmedU) === -1) {
            newUsers.push(trimmedU);
          }
        });
      }
    });
    
    // もし新しいカテゴリーかユーザーがあればマスターシートを更新
    if (newCategories.length > 0 || newUsers.length > 0) {
      var updatedCategories = existingCategories.concat(newCategories);
      var updatedUsers = existingUsers.concat(newUsers);
      
      var splitTypesCol = [];
      var paymentMethodsCol = [];
      var settlementStatusesCol = [];
      for (var r = 1; r < masterValues.length; r++) {
        splitTypesCol.push(masterValues[r][2] || "");
        paymentMethodsCol.push(masterValues[r][3] || "");
        settlementStatusesCol.push(masterValues[r][4] || "");
      }
      
      // 支払いパターンを全組み合わせで再生成
      var updatedPaymentPatterns = [];
      for (var i = 0; i < updatedUsers.length; i++) {
        for (var j = 0; j < updatedUsers.length; j++) {
          if (i !== j) {
            updatedPaymentPatterns.push(updatedUsers[i] + "から" + updatedUsers[j] + "へ");
          }
        }
      }
      
      var maxRows = Math.max(
        updatedCategories.length,
        updatedUsers.length,
        splitTypesCol.length,
        paymentMethodsCol.length,
        settlementStatusesCol.length,
        updatedPaymentPatterns.length
      );
      
      var newMasterData = [];
      for (var i = 0; i < maxRows; i++) {
        newMasterData.push([
          updatedCategories[i] || "",
          updatedUsers[i] || "",
          splitTypesCol[i] || "",
          paymentMethodsCol[i] || "",
          settlementStatusesCol[i] || "",
          updatedPaymentPatterns[i] || ""
        ]);
      }
      
      if (mLR > 1) {
        masterSheet.getRange(2, 1, mLR - 1, 6).clearContent();
      }
      masterSheet.getRange(2, 1, newMasterData.length, 6).setValues(newMasterData);
      
      // キャッシュのクリア（マスター再読み込みを誘発）
      try {
        var cache = CacheService.getScriptCache();
        if (cache) {
          cache.remove("db_initialized_v2");
        }
      } catch (cacheErr) {
        console.warn("キャッシュクリア中にエラーが発生しました: " + cacheErr.message);
      }
    }
    
    // データベースの更新タイムスタンプを更新
    updateDatabaseTimestamp();
    
    return { success: true, count: txList.length };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * 月次精算ステータスを保存・更新する
 */
function saveSettlementStatus(data) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("settlements");
    if (!sheet) {
      checkAndInitializeSpreadsheet();
      sheet = ss.getSheetByName("settlements");
    }
    
    var lr = sheet.getLastRow();
    var lc = sheet.getLastColumn();
    var values = (lr > 0 && lc > 0) ? sheet.getRange(1, 1, lr, lc).getValues() : [[]];
    var headers = values[0];
    
    var id = data.id || "";
    var rowNum = -1;
    var idColIdx = headers.indexOf("id");
    
    if (id && idColIdx !== -1) {
      for (var i = 1; i < values.length; i++) {
        if (values[i][idColIdx] === id) {
          rowNum = i + 1;
          break;
        }
      }
    }
    
    if (!id) {
      id = Utilities.getUuid();
    }
    
    var pDate = null;
    if (data.payment_date) {
      var dateParts = data.payment_date.split("-");
      pDate = new Date(Number(dateParts[0]), Number(dateParts[1]) - 1, Number(dateParts[2]));
    }
    
    var rowData = {
      id: id,
      month: data.month,
      status: data.status || "未払い",
      payment_date: pDate || "",
      payment_pattern: data.payment_pattern || "",
      amount: Number(data.amount) || 0,
      payment_method: data.payment_method || "",
      memo: data.memo || "",
      updated_at: new Date()
    };
    
    var newRow = headers.map(function(h) {
      return rowData.hasOwnProperty(h) ? rowData[h] : "";
    });
    
    if (rowNum !== -1) {
      sheet.getRange(rowNum, 1, 1, newRow.length).setValues([newRow]);
    } else {
      sheet.appendRow(newRow);
    }
    
    updateDatabaseTimestamp();
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * 精算記録を削除する
 */
function deleteSettlementStatus(id) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("settlements");
    if (!sheet) return { success: false, message: "精算シートが見つかりません。" };
    
    var lr = sheet.getLastRow();
    var lc = sheet.getLastColumn();
    var values = (lr > 0 && lc > 0) ? sheet.getRange(1, 1, lr, lc).getValues() : [[]];
    var headers = values[0];
    var idColIdx = headers.indexOf("id");
    
    if (idColIdx === -1) return { success: false, message: "IDカラムが見つかりません。" };
    
    var rowNum = -1;
    for (var i = 1; i < values.length; i++) {
      if (values[i][idColIdx] === id) {
        rowNum = i + 1;
        break;
      }
    }
    
    if (rowNum !== -1) {
      sheet.deleteRow(rowNum);
      updateDatabaseTimestamp();
      return { success: true };
    }
    return { success: false, message: "指定された記録が見つかりません。" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * スプレッドシート内のmonth列の値（日付または文字列）を "YYYY-MM" の文字列に変換する
 */
function convertMonthToString(val, timezone) {
  if (val instanceof Date) {
    if (!isNaN(val.getTime())) {
      return Utilities.formatDate(val, timezone, "yyyy-MM");
    }
    return "";
  }
  var str = String(val).trim();
  var match = str.match(/^(\d{4})[-/](\d{1,2})/);
  if (match) {
    var y = match[1];
    var m = ("0" + match[2]).slice(-2);
    return y + "-" + m;
  }
  return str;
}

/**
 * 支払い日などの日付列の値（日付または文字列）を "YYYY-MM-DD" の文字列に変換する
 */
function convertDateStringToISO(val, timezone) {
  if (val instanceof Date) {
    if (!isNaN(val.getTime())) {
      return Utilities.formatDate(val, timezone, "yyyy-MM-dd");
    }
    return "";
  }
  var str = String(val).trim();
  var match = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (match) {
    var y = match[1];
    var m = ("0" + match[2]).slice(-2);
    var d = ("0" + match[3]).slice(-2);
    return y + "-" + m + "-" + d;
  }
  return str;
}

/**
 * データベース全体の更新タイムスタンプを取得する（infoシートから取得）
 */
function getDatabaseTimestamp(ss) {
  try {
    ss = ss || getSpreadsheet();
    var infoSheet = ss.getSheetByName("info");
    var ts = "";
    if (infoSheet) {
      // A1がlast_updatedの場合、B1から値を取得。それ以外の場合はA列からlast_updatedを探す
      var data = infoSheet.getDataRange().getValues();
      for (var i = 0; i < data.length; i++) {
        if (data[i][0] === "last_updated") {
          ts = data[i][1];
          break;
        }
      }
    }
    
    if (!ts) {
      ts = new Date().getTime().toString();
      if (!infoSheet) {
        infoSheet = ss.insertSheet("info");
      }
      // もし既存のinfoシートでlast_updatedがなかった場合は、1行目に挿入または上書き
      var data = infoSheet.getDataRange().getValues();
      var found = false;
      for (var i = 0; i < data.length; i++) {
        if (data[i][0] === "last_updated") {
          infoSheet.getRange(i + 1, 2).setValue(ts);
          found = true;
          break;
        }
      }
      if (!found) {
        infoSheet.getRange("A1").setValue("last_updated");
        infoSheet.getRange("B1").setValue(ts);
      }
    }
    return String(ts);
  } catch (e) {
    Logger.log("Failed to get timestamp from info sheet: " + e.toString());
    return new Date().getTime().toString();
  }
}

/**
 * データベースの更新タイムスタンプを強制的に現在時刻で更新する（infoシートを更新）
 */
function updateDatabaseTimestamp() {
  try {
    var ss = getSpreadsheet();
    var infoSheet = ss.getSheetByName("info");
    var ts = new Date().getTime().toString();
    
    if (!infoSheet) {
      infoSheet = ss.insertSheet("info");
      infoSheet.getRange("A1").setValue("last_updated");
      infoSheet.getRange("B1").setValue(ts);
    } else {
      var data = infoSheet.getDataRange().getValues();
      var found = false;
      for (var i = 0; i < data.length; i++) {
        if (data[i][0] === "last_updated") {
          infoSheet.getRange(i + 1, 2).setValue(ts);
          found = true;
          break;
        }
      }
      if (!found) {
        infoSheet.getRange("A1").setValue("last_updated");
        infoSheet.getRange("B1").setValue(ts);
      }
    }
    
    Logger.log("Database timestamp updated to: " + ts);
    return ts;
  } catch (e) {
    Logger.log("Failed to update timestamp in info sheet: " + e.toString());
    return new Date().getTime().toString();
  }
}

function forceSyncDatabaseTimestamp() {
  return updateDatabaseTimestamp();
}

/**
 * Googleドライブから直近の画像・PDFファイルを取得する
 */
function getRecentDriveFiles() {
  try {
    var files = [];
    var query = "(mimeType = 'application/pdf' or mimeType = 'image/jpeg' or mimeType = 'image/png' or mimeType = 'image/gif') and trashed = false";
    var fileIterator = DriveApp.searchFiles(query);
    
    while (fileIterator.hasNext() && files.length < 50) {
      var file = fileIterator.next();
      files.push({
        id: file.getId(),
        name: file.getName(),
        mimeType: file.getMimeType(),
        lastUpdated: file.getLastUpdated().getTime(),
        url: file.getUrl()
      });
    }
    
    // 更新日時（降順）でソート
    files.sort(function(a, b) {
      return b.lastUpdated - a.lastUpdated;
    });
    
    return {
      success: true,
      files: files.slice(0, 20)
    };
  } catch (e) {
    return {
      success: false,
      message: "Googleドライブのファイル一覧取得に失敗しました: " + e.toString()
    };
  }
}

/**
 * Googleドライブのファイルを名前で検索する
 */
function searchDriveFiles(queryText) {
  try {
    var files = [];
    var query = "(mimeType = 'application/pdf' or mimeType = 'image/jpeg' or mimeType = 'image/png' or mimeType = 'image/gif') and trashed = false";
    if (queryText) {
      var escapedQuery = queryText.replace(/'/g, "\\'");
      query += " and title contains '" + escapedQuery + "'";
    }
    var fileIterator = DriveApp.searchFiles(query);
    
    while (fileIterator.hasNext() && files.length < 50) {
      var file = fileIterator.next();
      files.push({
        id: file.getId(),
        name: file.getName(),
        mimeType: file.getMimeType(),
        lastUpdated: file.getLastUpdated().getTime(),
        url: file.getUrl()
      });
    }
    
    // 更新日時（降順）でソート
    files.sort(function(a, b) {
      return b.lastUpdated - a.lastUpdated;
    });
    
    return {
      success: true,
      files: files.slice(0, 20)
    };
  } catch (e) {
    return {
      success: false,
      message: "Googleドライブの検索に失敗しました: " + e.toString()
    };
  }
}

/**
 * Googleドライブ上のファイルを読み込み、Geminiで解析する
 */
function analyzeDriveReceipt(fileId, clientApiKey) {
  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var mimeType = blob.getContentType();
    
    // サポート対象MimeTypeチェック
    var allowedTypes = ["application/pdf", "image/jpeg", "image/png", "image/gif"];
    if (allowedTypes.indexOf(mimeType) === -1) {
      return {
        success: false,
        message: "このファイル形式 (" + mimeType + ") は解析に対応していません。画像かPDFを選択してください。"
      };
    }
    
    var base64Data = Utilities.base64Encode(blob.getBytes());
    return analyzeReceipt(base64Data, mimeType, clientApiKey);
  } catch (e) {
    return {
      success: false,
      message: "Googleドライブのファイル取得・解析に失敗しました: " + e.toString()
    };
  }
}

/**
 * スプレッドシートの最新の更新タイムスタンプを返す軽量API
 */
function getLatestTimestamp() {
  try {
    return getDatabaseTimestamp();
  } catch (e) {
    return new Date().getTime().toString();
  }
}

