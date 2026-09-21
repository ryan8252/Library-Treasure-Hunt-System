// 老師體驗版：小隊與紀錄存放於 DemoTeams、DemoBorrowRecords。
// 總共書的數量
const MAX_TREASURE_BOOKS = 100;

// 每個小隊需要找到幾本
const REQUIRED_BOOKS_PER_STUDENT = 10;


function getTreasureConfig() {

  return {

    maxBooks:
      MAX_TREASURE_BOOKS,

    requiredBooks:
      REQUIRED_BOOKS_PER_STUDENT

  };

}

function doGet(e) {

  const page =
    e && e.parameter
      ? e.parameter.page
      : "";


  // 管理尋寶書籍
  if (page === "admin") {

    return HtmlService
      .createHtmlOutputFromFile("Admin")
      .setTitle("尋寶書籍管理");

  }


  // 小隊註冊
  if (page === "register") {

    return HtmlService
      .createHtmlOutputFromFile("Register")
      .setTitle("尋寶活動小隊註冊");

  }


  // 排行榜
  if (page === "leaderboard") {

    return HtmlService
      .createHtmlOutputFromFile("Leaderboard")
      .setTitle("圖書館尋寶排行榜");

  }


  // 學生尋寶
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("圖書館尋寶活動");

}



// ======================================================
// 把全形數字轉成半形
// 例如：０００３３９０６ → 00033906
// ======================================================

function normalizeBarcode_(value) {

  return String(value || "")
    .trim()
    .replace(/[０-９]/g, function(char) {
      return String.fromCharCode(
        char.charCodeAt(0) - 0xFEE0
      );
    });

}


// ======================================================
// 取得 Teams 工作表
// 如果不存在就自動建立
// ======================================================

function getTeamsSheet() {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();


  let sheet =
    spreadsheet.getSheetByName(
      "DemoTeams"
    );


  if (!sheet) {

    sheet =
      spreadsheet.insertSheet(
        "DemoTeams"
      );


    sheet
      .getRange(1, 1, 1, 5)
      .setValues([[
        "library_card",
        "team_name",
        "member1_name",
        "member2_name",
        "registered_at"
      ]]);

    sheet.getRange("A:D").setNumberFormat("@");
  }





  return sheet;

}


// ======================================================
// 用借書證尋找小隊
// ======================================================

function getTeamByCard(libraryCard) {
  return findTeam_(getDemoState_(), libraryCard);
}

function registerTeamUncached_(
  libraryCard,
  teamName,
  member1Id,
  member2Id
) {


  libraryCard =
    String(libraryCard || "")
      .trim()
      .toUpperCase();


  teamName =
    String(teamName || "")
      .trim();


  member1Id =
    String(member1Id || "")
      .trim();


  member2Id =
    String(member2Id || "")
      .trim();


  if (
    !/^[A-Z][12][0-9]{8}$/.test(
      libraryCard
    )
  ) {

    throw new Error(
      "借書證格式不正確"
    );

  }


  if (!teamName) {

    throw new Error(
      "請輸入小隊名稱"
    );

  }


  if (!member1Id) {
    throw new Error("請輸入隊員一的姓名或編號");
  }

  const sheet =
    getTeamsSheet();


  {

    const lastRow =
      sheet.getLastRow();


    if (lastRow >= 2) {

      const values =
        sheet
          .getRange(
            2,
            1,
            lastRow - 1,
            5
          )
          .getDisplayValues();


      for (
        let i = 0;
        i < values.length;
        i++
      ) {

        const existingCard =
          String(values[i][0] || "")
            .trim()
            .toUpperCase();


        // 同一張卡
        // → 更新原本小隊
        if (
          existingCard === libraryCard
        ) {

          const row =
            i + 2;


          sheet.getRange(row, 1, 1, 4).setNumberFormat("@");


          sheet
            .getRange(
              row,
              1,
              1,
              5
            )
            .setValues([[
              libraryCard,
              teamName,
              member1Id,
              member2Id,
              new Date()
            ]]);


          return {

            success: true,

            updated: true,

            libraryCard:
              libraryCard,

            teamName:
              teamName,

            member1Id:
              member1Id,

            member2Id:
              member2Id

          };

        }

      }

    }


    const row =
      sheet.getLastRow() + 1;
    sheet.getRange(row, 1, 1, 4).setNumberFormat("@");


    sheet
      .getRange(
        row,
        1,
        1,
        5
      )
      .setValues([[
        libraryCard,
        teamName,
        member1Id,
        member2Id,
        new Date()
      ]]);


    return {

      success: true,

      updated: false,

      libraryCard:
        libraryCard,

      teamName:
        teamName,

      member1Id:
        member1Id,

      member2Id:
        member2Id

    };


  }

}


// ======================================================
// 取得尋寶書籍工作表
// ======================================================

function getTreasureBooksSheet() {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();


  let sheet =
    spreadsheet.getSheetByName("TreasureBooks");


  // 如果不存在，自動建立
  if (!sheet) {

    sheet =
      spreadsheet.insertSheet("TreasureBooks");

    sheet
      .getRange(1, 1, 1, 2)
      .setValues([[
        "barcode",
        "updated_at"
      ]]);

    sheet
      .getRange("A:A")
      .setNumberFormat("@");

    return sheet;

  }



  // ==================================================
  // 如果是上一版的格式：
  //
  // slot | barcode | title | cabinet | updated_at
  //
  // 自動轉成新版：
  //
  // barcode | updated_at
  // ==================================================

  const firstHeader =
    String(
      sheet.getRange(1, 1).getValue()
    ).trim();


  if (firstHeader === "slot") {

    const oldBarcodes =
      sheet
        .getRange(2, 2, 10, 1)
        .getDisplayValues()
        .flat()
        .map(normalizeBarcode_)
        .filter(Boolean)
        .slice(0, 10);


    sheet.clear();


    sheet
      .getRange(1, 1, 1, 2)
      .setValues([[
        "barcode",
        "updated_at"
      ]]);


    sheet
      .getRange("A:A")
      .setNumberFormat("@");


    oldBarcodes.forEach(
      function(barcode, index) {

        const row =
          index + 2;

        sheet
          .getRange(row, 1)
          .setNumberFormat("@")
          .setValue(barcode);

        sheet
          .getRange(row, 2)
          .setValue(new Date());

      }
    );

  }


  return sheet;

}



// ======================================================
// 取得目前所有指定書
// ======================================================

function getTreasureBooks() {

  const sheet =
    getTreasureBooksSheet();


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    return [];

  }


  return sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      1
    )
    .getDisplayValues()
    .flat()
    .map(normalizeBarcode_)
    .filter(Boolean);

}



// ======================================================
// 新增一本指定書
// ======================================================

function addTreasureBookUncached_(barcode) {

  barcode =
    normalizeBarcode_(barcode);


  if (!barcode) {

    throw new Error(
      "沒有讀取到書籍條碼"
    );

  }


  // 書籍條碼只能是數字
  if (!/^[0-9]+$/.test(barcode)) {

    throw new Error(
      "書籍條碼必須全部都是數字"
    );

  }


  {

    const books =
      getTreasureBooks();


    // 重複
    if (books.includes(barcode)) {

      throw new Error(
        "這本書已經在尋寶書單裡了"
      );

    }


    // 最大 100 本
    if (books.length >= MAX_TREASURE_BOOKS) {
      throw new Error(
        "已達 " + MAX_TREASURE_BOOKS + " 本上限，請先刪除一本舊書"
      );
    }


    const sheet =
      getTreasureBooksSheet();


    const row =
      sheet.getLastRow() + 1;


    // 條碼強制文字
    sheet
      .getRange(row, 1)
      .setNumberFormat("@")
      .setValue(barcode);


    sheet
      .getRange(row, 2)
      .setValue(new Date());


    return {

      success: true,

      barcode: barcode,

      count:
        books.length + 1

    };


  }

}



// ======================================================
// 刪除一本指定書
// ======================================================

function deleteTreasureBookUncached_(barcode) {

  barcode =
    normalizeBarcode_(barcode);


  const sheet =
    getTreasureBooksSheet();


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    throw new Error(
      "目前沒有設定任何書籍"
    );

  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getDisplayValues();


  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    const current =
      normalizeBarcode_(
        values[i][0]
      );


    if (current === barcode) {

      // +2 因為第一列是標題
      sheet.deleteRow(i + 2);


      return {
        success: true
      };

    }

  }


  throw new Error(
    "找不到這本書"
  );

}



// ======================================================
// 判斷是不是指定書
// ======================================================

function findTreasureBook(bookBarcode) {

  bookBarcode =
    normalizeBarcode_(
      bookBarcode
    );


  const books =
    getTreasureBooks();


  return books.includes(
    bookBarcode
  );

}



// ======================================================
// 學生成功刷書後儲存紀錄
// ======================================================

function saveBorrowRecord(studentId, bookBarcode) {
  studentId = String(studentId || "").trim().toUpperCase();
  bookBarcode = normalizeBarcode_(bookBarcode);
  if (!studentId) throw new Error("讀者證號不能為空");
  if (!bookBarcode) throw new Error("書籍條碼不能為空");
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // Lock covers snapshot loading, duplicate checks, persistence and cache update.
    const state = readDemoStateLocked_();
    const team = findTeam_(state, studentId);
    if (!team) throw new Error("這張借書證尚未完成小隊註冊");
    if (!state.books.includes(bookBarcode)) {
      throw new Error("這不是本次尋寶活動指定的書籍");
    }
    const found = foundBooks_(state, studentId);
    const isDuplicate = found.has(bookBarcode);
    const alreadyCompleted = found.size >= REQUIRED_BOOKS_PER_STUDENT;
    if (!isDuplicate && !alreadyCompleted) {
      // Remove the old snapshot before writing so a failed write cannot leave stale data.
      CacheService.getScriptCache().remove(DEMO_CACHE_KEY);
      const sheet = getBorrowRecordsSheet();
      const row = sheet.getLastRow() + 1;
      sheet.getRange(row, 1, 1, 2).setNumberFormat("@");
      sheet.getRange(row, 1, 1, 4).setValues([[
        studentId, bookBarcode, new Date(), "尋寶成功"
      ]]);
      SpreadsheetApp.flush();
      state.records.push([studentId, bookBarcode]);
      putDemoState_(state);
      found.add(bookBarcode);
    }
    return Object.assign(teamProgress_(team, found), {
      success: true, studentId, bookBarcode, isDuplicate, alreadyCompleted
    });
  } finally {
    lock.releaseLock();
  }
}

function getStudentProgress(studentId) {
  const state = getDemoState_();
  const team = findTeam_(state, studentId);
  if (!team) return {registered: false, foundCount: 0,
    totalCount: REQUIRED_BOOKS_PER_STUDENT, completed: false};
  return teamProgress_(team, foundBooks_(state, team.libraryCard));
}

function getLeaderboardData() {
  const state = getDemoState_();
  const books = new Set(state.books);
  const foundByCard = new Map();
  state.records.forEach(([card, barcode]) => {
    if (!books.has(barcode)) return;
    if (!foundByCard.has(card)) foundByCard.set(card, new Set());
    foundByCard.get(card).add(barcode);
  });
  const teams = state.teams.map(team =>
    teamProgress_(team, foundByCard.get(team.libraryCard) || new Set()));
  teams.sort((a, b) => b.foundCount - a.foundCount ||
    a.teamName.localeCompare(b.teamName, "zh-Hant"));
  let previousCount = null;
  let rank = 0;
  teams.forEach((team, index) => {
    if (team.foundCount !== previousCount) rank = index + 1;
    team.rank = rank;
    previousCount = team.foundCount;
  });
  return {teams, teamCount: teams.length,
    requiredCount: REQUIRED_BOOKS_PER_STUDENT,
    candidateBookCount: state.books.length,
    updatedAt: Utilities.formatDate(new Date(), state.timeZone, "HH:mm:ss")};
}

function getWebAppUrl() {

  return ScriptApp
    .getService()
    .getUrl();

}


// ======================================================
// ======================================================

function getBorrowRecordsSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName("DemoBorrowRecords");
  if (!sheet) {
    sheet = spreadsheet.insertSheet("DemoBorrowRecords");
    sheet.getRange(1, 1, 1, 4).setValues([[
      "library_card", "barcode", "found_at", "status"
    ]]);
    sheet.getRange("A:B").setNumberFormat("@");
  }
  return sheet;
}

// One shared snapshot for card verification, scoring and polling.
// Absolute expiry bounds visibility of manual spreadsheet edits even during active scanning.
const DEMO_CACHE_KEY = "teacher-demo-state-v1";
const DEMO_CACHE_SECONDS = 10;

function cachedDemoState_() {
  try {
    const raw = CacheService.getScriptCache().get(DEMO_CACHE_KEY);
    if (!raw) return null;
    const state = JSON.parse(raw);
    return Date.now() - state.loadedAt < DEMO_CACHE_SECONDS * 1000 ? state : null;
  } catch (error) {
    return null;
  }
}

function putDemoState_(state) {
  const remaining = Math.floor((state.loadedAt + DEMO_CACHE_SECONDS * 1000 - Date.now()) / 1000);
  if (remaining < 1) return;
  const json = JSON.stringify(state);
  // UTF-8 uses at most three bytes per UTF-16 code unit; stay below CacheService's 100 KB limit.
  if (json.length * 3 > 90000) return;
  try {
    CacheService.getScriptCache().put(DEMO_CACHE_KEY, json, remaining);
  } catch (error) {
    // Cache eviction or a failed put must not turn a persisted scan into a reported failure.
  }
}

function readRows_(sheet, width) {
  const lastRow = sheet.getLastRow();
  return lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, width).getDisplayValues();
}

// Caller must hold the script lock; prevents cold readers publishing an older snapshot after a write.
function readDemoStateLocked_() {
  const cached = cachedDemoState_();
  if (cached) return cached;
  const state = {
    loadedAt: Date.now(),
    teams: readRows_(getTeamsSheet(), 4).filter(row => row[0] && row[1]).map(row => ({
      libraryCard: String(row[0]).trim().toUpperCase(),
      teamName: row[1], member1Id: row[2], member2Id: row[3]
    })),
    books: getTreasureBooks(),
    records: readRows_(getBorrowRecordsSheet(), 2).map(row => [
      String(row[0]).trim().toUpperCase(), normalizeBarcode_(row[1])
    ]),
    timeZone: SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()
  };
  SpreadsheetApp.flush();
  putDemoState_(state);
  return state;
}

function getDemoState_() {
  const cached = cachedDemoState_();
  if (cached) return cached;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return readDemoStateLocked_();
  } finally {
    lock.releaseLock();
  }
}

function findTeam_(state, libraryCard) {
  const card = String(libraryCard || "").trim().toUpperCase();
  return state.teams.find(team => team.libraryCard === card) || null;
}

function foundBooks_(state, libraryCard) {
  const valid = new Set(state.books);
  return new Set(state.records.filter(row => row[0] === libraryCard && valid.has(row[1]))
    .map(row => row[1]));
}

function teamProgress_(team, found) {
  const count = Math.min(found.size, REQUIRED_BOOKS_PER_STUDENT);
  return Object.assign({}, team, {registered: true, foundCount: count,
    totalCount: REQUIRED_BOOKS_PER_STUDENT, completed: count >= REQUIRED_BOOKS_PER_STUDENT});
}

function mutateDemo_(operation) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    CacheService.getScriptCache().remove(DEMO_CACHE_KEY);
    return operation();
  } finally {
    try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
  }
}

function registerTeam(libraryCard, teamName, member1Id, member2Id) {
  return mutateDemo_(() => registerTeamUncached_(libraryCard, teamName, member1Id, member2Id));
}

function addTreasureBook(barcode) {
  return mutateDemo_(() => addTreasureBookUncached_(barcode));
}

function deleteTreasureBook(barcode) {
  return mutateDemo_(() => deleteTreasureBookUncached_(barcode));
}
