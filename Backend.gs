// 總共書的數量
const MAX_TREASURE_BOOKS = 100;

// 每位學生需要找到幾本
const REQUIRED_BOOKS_PER_STUDENT = 10;


const VALID_CLASSES = [
  "501", "502", "503", "504", "505", "506", "507",
  "601", "602", "603", "604", "605", "606", "607"
];


const CLASS_NAMES = {

  "501": "五年一班",
  "502": "五年二班",
  "503": "五年三班",
  "504": "五年四班",
  "505": "五年五班",
  "506": "五年六班",
  "507": "五年七班",

  "601": "六年一班",
  "602": "六年二班",
  "603": "六年三班",
  "604": "六年四班",
  "605": "六年五班",
  "606": "六年六班",
  "607": "六年七班"

};


function normalizeClassId_(classId) {

  return String(classId || "")
    .trim();

}


function getClassInfo(classId) {

  classId =
    normalizeClassId_(classId);


  if (
    !VALID_CLASSES.includes(classId)
  ) {

    return {

      valid: false

    };

  }


  return {

    valid: true,

    classId:
      classId,

    className:
      CLASS_NAMES[classId]

  };

}


function assertValidClass_(classId) {

  classId =
    normalizeClassId_(classId);


  if (
    !VALID_CLASSES.includes(classId)
  ) {

    throw new Error(
      "班級代碼不正確"
    );

  }


  return classId;

}

// ======================================================
// 讀取設定
// ======================================================
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
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName("Teams");
  if (!sheet) {
    sheet = spreadsheet.insertSheet("Teams");
    sheet.getRange(1, 1, 1, 6).setValues([[
      "library_card", "team_name", "member1_id", "member2_id", "registered_at", "class_id"
    ]]);
    sheet.getRange("A:D").setNumberFormat("@");
    sheet.getRange("F:F").setNumberFormat("@");
  } else if (sheet.getRange(1, 6).getValue() === "") {
    sheet.getRange(1, 6).setValue("class_id");
  }
  return sheet;
}

function getTeamByCard(classId, libraryCard) {
  classId = assertValidClass_(classId);
  return findTeam_(getStudentState_(classId), classId, libraryCard);
}

function registerTeamUncached_(
  classId,
  libraryCard,
  teamName,
  member1Id,
  member2Id
) {

  classId =
    assertValidClass_(classId);


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


  if (
    Array.from(teamName).length > 5
  ) {

    throw new Error(
      "小隊名稱最多 5 個字"
    );

  }


  if (
    !/^[0-9]{5}$/.test(
      member1Id
    )
  ) {

    throw new Error(
      "隊員一學號格式不正確"
    );

  }


  if (
    !/^[0-9]{5}$/.test(
      member2Id
    )
  ) {

    throw new Error(
      "隊員二學號格式不正確"
    );

  }


  if (
    member1Id === member2Id
  ) {

    throw new Error(
      "兩位隊員學號不能相同"
    );

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
            6
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


        const existingClass =
          String(values[i][5] || "")
            .trim();


        // 同一班、同一張卡
        // → 更新原本小隊
        if (
          existingCard === libraryCard &&
          existingClass === classId
        ) {

          const row =
            i + 2;


          sheet.getRange(row, 1, 1, 4).setNumberFormat("@");


          sheet.getRange(row, 6).setNumberFormat("@");


          sheet
            .getRange(
              row,
              1,
              1,
              6
            )
            .setValues([[
              libraryCard,
              teamName,
              member1Id,
              member2Id,
              new Date(),
              classId
            ]]);


          return {

            success: true,

            updated: true,

            classId:
              classId,

            className:
              CLASS_NAMES[classId],

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


    sheet.getRange(row, 6).setNumberFormat("@");


    sheet
      .getRange(
        row,
        1,
        1,
        6
      )
      .setValues([[
        libraryCard,
        teamName,
        member1Id,
        member2Id,
        new Date(),
        classId
      ]]);


    return {

      success: true,

      updated: false,

      classId:
        classId,

      className:
        CLASS_NAMES[classId],

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

function saveBorrowRecord(classId, studentId, bookBarcode) {
  classId = assertValidClass_(classId);
  studentId = String(studentId || "").trim().toUpperCase();
  bookBarcode = normalizeBarcode_(bookBarcode);
  if (!studentId) throw new Error("讀者證號不能為空");
  if (!bookBarcode) throw new Error("書籍條碼不能為空");
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // Lock covers snapshot loading, duplicate checks, persistence and cache update.
    const state = readStudentStateLocked_(classId);
    const team = findTeam_(state, classId, studentId);
    if (!team) throw new Error("這張借書證尚未完成小隊註冊");
    if (!state.books.includes(bookBarcode)) {
      throw new Error("這不是本次尋寶活動指定的書籍");
    }
    const found = foundBooks_(state, classId, studentId);
    const isDuplicate = found.has(bookBarcode);
    const alreadyCompleted = found.size >= REQUIRED_BOOKS_PER_STUDENT;
    if (!isDuplicate && !alreadyCompleted) {
      // Remove the old snapshot before writing so a failed write cannot leave stale data.
      CacheService.getScriptCache().remove(studentCacheKey_(classId));
      // Snapshot loading already checked the schema; avoid rereading the header for each scan.
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BorrowRecords");
      if (!sheet) throw new Error("找不到 BorrowRecords 工作表");
      const row = sheet.getLastRow() + 1;
      sheet.getRange(row, 1, 1, 2).setNumberFormat("@");
      sheet.getRange(row, 5).setNumberFormat("@");
      sheet.getRange(row, 1, 1, 5).setValues([[
        studentId, bookBarcode, new Date(), "尋寶成功", classId
      ]]);
      SpreadsheetApp.flush();
      state.records.push([studentId, bookBarcode, classId]);
      putStudentState_(state);
      found.add(bookBarcode);
    }
    return Object.assign(teamProgress_(team, found), {
      success: true, studentId, bookBarcode, isDuplicate, alreadyCompleted
    });
  } finally {
    lock.releaseLock();
  }
}


function getStudentProgress(classId, studentId) {
  classId = assertValidClass_(classId);
  const state = getStudentState_(classId);
  const team = findTeam_(state, classId, studentId);
  if (!team) return {registered: false, classId, className: CLASS_NAMES[classId],
    foundCount: 0, totalCount: REQUIRED_BOOKS_PER_STUDENT, completed: false};
  return teamProgress_(team, foundBooks_(state, classId, team.libraryCard));
}

function getLeaderboardData(classId) {
  classId = assertValidClass_(classId);
  const state = getStudentState_(classId);
  const books = new Set(state.books);
  const foundByCard = new Map();
  state.records.forEach(([card, barcode, recordClass]) => {
    if (recordClass !== classId || !books.has(barcode)) return;
    if (!foundByCard.has(card)) foundByCard.set(card, new Set());
    foundByCard.get(card).add(barcode);
  });
  const teams = state.teams.filter(team => team.classId === classId).map(team =>
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
  return {classId, className: CLASS_NAMES[classId], teams, teamCount: teams.length,
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
// BorrowRecords 增加 class_id 的 helper
// ======================================================

function getBorrowRecordsSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BorrowRecords");
  if (!sheet) throw new Error("找不到 BorrowRecords 工作表");
  if (sheet.getRange(1, 5).getValue() === "") {
    sheet.getRange(1, 5).setValue("class_id");
  }
  return sheet;
}

// Each class shares its own snapshot for card verification, scoring and polling.
// Absolute expiry bounds visibility of manual spreadsheet edits even during active scanning.
const STUDENT_CACHE_PREFIX = "student-state-v2:";

function studentCacheKey_(classId) {
  return STUDENT_CACHE_PREFIX + assertValidClass_(classId);
}
const STUDENT_CACHE_SECONDS = 10;

function cachedStudentState_(classId) {
  try {
    const raw = CacheService.getScriptCache().get(studentCacheKey_(classId));
    if (!raw) return null;
    const state = JSON.parse(raw);
    return state.classId === classId && Date.now() - state.loadedAt < STUDENT_CACHE_SECONDS * 1000 ? state : null;
  } catch (error) {
    return null;
  }
}

function putStudentState_(state) {
  const remaining = Math.floor((state.loadedAt + STUDENT_CACHE_SECONDS * 1000 - Date.now()) / 1000);
  if (remaining < 1) return;
  const json = JSON.stringify(state);
  // UTF-8 uses at most three bytes per UTF-16 code unit; stay below CacheService's 100 KB limit.
  if (json.length * 3 > 90000) return;
  try {
    CacheService.getScriptCache().put(studentCacheKey_(state.classId), json, remaining);
  } catch (error) {
    // Cache eviction or a failed put must not turn a persisted scan into a reported failure.
  }
}

function readRows_(sheet, width) {
  const lastRow = sheet.getLastRow();
  return lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, width).getDisplayValues();
}

// Caller must hold the script lock; prevents cold readers publishing an older snapshot after a write.
function readStudentStateLocked_(classId) {
  const cached = cachedStudentState_(classId);
  if (cached) return cached;
  const state = {
    classId,
    loadedAt: Date.now(),
    teams: readRows_(getTeamsSheet(), 6).filter(row => row[0] && row[1] && String(row[5]).trim() === classId).map(row => ({
      libraryCard: String(row[0]).trim().toUpperCase(),
      teamName: row[1], member1Id: row[2], member2Id: row[3],
      classId: String(row[5]).trim(), className: CLASS_NAMES[String(row[5]).trim()] || ""
    })),
    books: getTreasureBooks(),
    records: readRows_(getBorrowRecordsSheet(), 5).filter(row => String(row[4]).trim() === classId).map(row => [
      String(row[0]).trim().toUpperCase(), normalizeBarcode_(row[1]), String(row[4]).trim()
    ]),
    timeZone: SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()
  };
  SpreadsheetApp.flush();
  putStudentState_(state);
  return state;
}

function getStudentState_(classId) {
  const cached = cachedStudentState_(classId);
  if (cached) return cached;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return readStudentStateLocked_(classId);
  } finally {
    lock.releaseLock();
  }
}

function findTeam_(state, classId, libraryCard) {
  const card = String(libraryCard || "").trim().toUpperCase();
  return state.teams.find(team => team.libraryCard === card && team.classId === classId) || null;
}

function foundBooks_(state, classId, libraryCard) {
  const valid = new Set(state.books);
  return new Set(state.records.filter(row => row[0] === libraryCard && row[2] === classId && valid.has(row[1]))
    .map(row => row[1]));
}

function teamProgress_(team, found) {
  const count = Math.min(found.size, REQUIRED_BOOKS_PER_STUDENT);
  return Object.assign({}, team, {registered: true, foundCount: count,
    totalCount: REQUIRED_BOOKS_PER_STUDENT, completed: count >= REQUIRED_BOOKS_PER_STUDENT});
}

function mutateStudent_(operation, classId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // Shared book-list edits affect every class; registration affects only its class.
    const classes = classId ? [assertValidClass_(classId)] : VALID_CLASSES;
    classes.forEach(id => CacheService.getScriptCache().remove(studentCacheKey_(id)));
    return operation();
  } finally {
    try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
  }
}

function registerTeam(classId, libraryCard, teamName, member1Id, member2Id) {
  classId = assertValidClass_(classId);
  return mutateStudent_(() => registerTeamUncached_(classId, libraryCard, teamName, member1Id, member2Id), classId);
}

function addTreasureBook(barcode) {
  return mutateStudent_(() => addTreasureBookUncached_(barcode));
}

function deleteTreasureBook(barcode) {
  return mutateStudent_(() => deleteTreasureBookUncached_(barcode));
}
