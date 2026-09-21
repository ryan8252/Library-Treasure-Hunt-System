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

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();


  let sheet =
    spreadsheet.getSheetByName(
      "Teams"
    );


  if (!sheet) {

    sheet =
      spreadsheet.insertSheet(
        "Teams"
      );


    sheet
      .getRange(1, 1, 1, 6)
      .setValues([[
        "library_card",
        "team_name",
        "member1_id",
        "member2_id",
        "registered_at",
        "class_id"
      ]]);

  }


  // 舊版 Teams 自動補 class_id
  if (
    sheet.getRange(1, 6).getValue() === ""
  ) {

    sheet
      .getRange(1, 6)
      .setValue("class_id");

  }


  sheet
    .getRange("A:A")
    .setNumberFormat("@");


  sheet
    .getRange("C:D")
    .setNumberFormat("@");


  sheet
    .getRange("F:F")
    .setNumberFormat("@");


  return sheet;

}


// ======================================================
// 用借書證尋找小隊
// ======================================================

function getTeamByCard(
  classId,
  libraryCard
) {

  classId =
    assertValidClass_(classId);


  libraryCard =
    String(libraryCard || "")
      .trim()
      .toUpperCase();


  const sheet =
    getTeamsSheet();


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    return null;

  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        6
      )
      .getDisplayValues();


  for (const row of values) {

    const card =
      String(row[0] || "")
        .trim()
        .toUpperCase();


    const rowClass =
      String(row[5] || "")
        .trim();


    if (
      card === libraryCard &&
      rowClass === classId
    ) {

      return {

        classId:
          classId,

        className:
          CLASS_NAMES[classId],

        libraryCard:
          card,

        teamName:
          row[1],

        member1Id:
          row[2],

        member2Id:
          row[3]

      };

    }

  }


  return null;

}

// ======================================================
// 註冊小隊
// ======================================================

function registerTeam(
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


  const lock =
    LockService.getScriptLock();


  lock.waitLock(10000);


  try {

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


  } finally {

    lock.releaseLock();

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

function addTreasureBook(barcode) {

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


  const lock =
    LockService.getScriptLock();


  lock.waitLock(10000);


  try {

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


  } finally {

    lock.releaseLock();

  }

}



// ======================================================
// 刪除一本指定書
// ======================================================

function deleteTreasureBook(barcode) {

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

function saveBorrowRecord(classId,studentId,bookBarcode) {

  studentId = String(studentId || "").trim().toUpperCase();


  bookBarcode = normalizeBarcode_(bookBarcode);


  if (!studentId) {

    throw new Error(
      "讀者證號不能為空"
    );

  }


  if (!bookBarcode) {

    throw new Error(
      "書籍條碼不能為空"
    );

  }


  // 尋找有沒有這個小隊
  classId = assertValidClass_(classId);
  const team = getTeamByCard(classId, studentId);

  if (!team) {

    throw new Error(
      "這張借書證尚未完成小隊註冊"
    );

  }


  // ======================================
  // 取得所有候選書
  // 最多可以有 100 本
  // ======================================

  const treasureBooks =
    getTreasureBooks();


  const requiredCount =
    REQUIRED_BOOKS_PER_STUDENT;



  // ======================================
  // 掃到的書必須在候選書庫內
  // ======================================

  if (
    !treasureBooks.includes(
      bookBarcode
    )
  ) {

    throw new Error(
      "這不是本次尋寶活動指定的書籍"
    );

  }



  const spreadsheet =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheet =
    spreadsheet.getSheetByName(
      "BorrowRecords"
    );


  if (!sheet) {

    throw new Error(
      "找不到 BorrowRecords 工作表"
    );

  }



  const lock =
    LockService.getScriptLock();


  lock.waitLock(10000);


  try {


    // ======================================
    // 找出這個學生以前找到過的書
    // ======================================

    const foundBooks =
      new Set();


    const lastRow =
      sheet.getLastRow();


    if (lastRow >= 2) {

      const records =
        sheet
          .getRange(
            2,
            1,
            lastRow - 1,
            2
          )
          .getDisplayValues();


      records.forEach(
        function(row) {


          const recordStudent =
            String(row[0] || "")
              .trim()
              .toUpperCase();


          const recordBook =
            normalizeBarcode_(
              row[1]
            );


          if (
            recordStudent === studentId &&
            treasureBooks.includes(
              recordBook
            )
          ) {

            foundBooks.add(
              recordBook
            );

          }

        }
      );

    }



    // ======================================
    // 如果這個學生早就完成 10 本
    // 就不要再增加紀錄
    // ======================================

    if (
      foundBooks.size >=
      requiredCount
    ) {

      return {

        success: true,

        studentId:
          studentId,

        teamName:
          team.teamName,

        member1Id:
          team.member1Id,

        member2Id:
          team.member2Id,

        bookBarcode:
          bookBarcode,

        isDuplicate:
          foundBooks.has(
            bookBarcode
          ),

        alreadyCompleted:
          true,

        foundCount:
          requiredCount,

        totalCount:
          requiredCount,

        completed:
          true

      };

    }



    // ======================================
    // 判斷是不是重複的書
    // ======================================

    const isDuplicate =
      foundBooks.has(
        bookBarcode
      );



    // ======================================
    // 沒找過才新增
    // ======================================

    if (!isDuplicate) {


      const row =
        sheet.getLastRow() + 1;


      sheet
        .getRange(row, 1)
        .setNumberFormat("@")
        .setValue(studentId);


      sheet
        .getRange(row, 2)
        .setNumberFormat("@")
        .setValue(bookBarcode);


      sheet
        .getRange(row, 3)
        .setValue(
          new Date()
        );


      sheet
        .getRange(
          row,
          1,
          1,
          5
        )
        .setValues([[
          studentId,
          bookBarcode,
          new Date(),
          "尋寶成功",
          classId
        ]]);


      foundBooks.add(
        bookBarcode
      );

    }



    // ======================================
    // 最大只顯示到 10
    // ======================================

    const foundCount =
      Math.min(
        foundBooks.size,
        requiredCount
      );


    const completed =
      foundCount >=
      requiredCount;



    return {

      success: true,

      studentId:
        studentId,

      teamName:
        team.teamName,

      member1Id:
        team.member1Id,

      member2Id:
        team.member2Id,

      bookBarcode:
        bookBarcode,

      isDuplicate:
        isDuplicate,

      alreadyCompleted:
        false,

      foundCount:
        foundCount,

      totalCount:
        requiredCount,

      completed:
        completed

    };


  } finally {

    lock.releaseLock();

  }

}

// ======================================================
// 取得學生目前進度的 function
// ======================================================
function getStudentProgress(
  classId,
  studentId
) {

  classId =
    assertValidClass_(classId);


  studentId =
    String(studentId || "")
      .trim()
      .toUpperCase();


  const team =
    getTeamByCard(
      classId,
      studentId
    );


  if (!team) {

    return {

      registered: false,

      classId:
        classId,

      className:
        CLASS_NAMES[classId],

      foundCount: 0,

      totalCount:
        REQUIRED_BOOKS_PER_STUDENT,

      completed: false

    };

  }


  const treasureBooks =
    getTreasureBooks();


  const treasureSet =
    new Set(treasureBooks);


  const sheet =
    getBorrowRecordsSheet();


  const foundBooks =
    new Set();


  const lastRow =
    sheet.getLastRow();


  if (lastRow >= 2) {

    const records =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          5
        )
        .getDisplayValues();


    records.forEach(function(row) {

      const card =
        String(row[0] || "")
          .trim()
          .toUpperCase();


      const barcode =
        normalizeBarcode_(
          row[1]
        );


      const recordClass =
        String(row[4] || "")
          .trim();


      if (
        recordClass === classId &&
        card === studentId &&
        treasureSet.has(barcode)
      ) {

        foundBooks.add(
          barcode
        );

      }

    });

  }


  const foundCount =
    Math.min(
      foundBooks.size,
      REQUIRED_BOOKS_PER_STUDENT
    );


  return {

    registered: true,

    classId:
      classId,

    className:
      CLASS_NAMES[classId],

    teamName:
      team.teamName,

    member1Id:
      team.member1Id,

    member2Id:
      team.member2Id,

    foundCount:
      foundCount,

    totalCount:
      REQUIRED_BOOKS_PER_STUDENT,

    completed:
      foundCount >=
      REQUIRED_BOOKS_PER_STUDENT

  };

}


// ======================================================
// 排行榜後端
// ======================================================

function getLeaderboardData(classId) {

  classId =
    assertValidClass_(classId);


  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();


  const teamsSheet =
    getTeamsSheet();


  const recordsSheet =
    spreadsheet.getSheetByName(
      "BorrowRecords"
    );


  if (!recordsSheet) {

    throw new Error(
      "找不到 BorrowRecords 工作表"
    );

  }


  // ==========================================
  // 目前活動所有有效候選書
  // ==========================================

  const treasureBooks =
    getTreasureBooks();


  const treasureSet =
    new Set(
      treasureBooks
    );


  const requiredCount =
    REQUIRED_BOOKS_PER_STUDENT;



  // ==========================================
  // 讀取目前班級的所有已註冊小隊
  // ==========================================

  const teams = [];


  const teamsLastRow =
    teamsSheet.getLastRow();


  if (teamsLastRow >= 2) {

    const teamValues =
      teamsSheet
        .getRange(
          2,
          1,
          teamsLastRow - 1,
          6
        )
        .getDisplayValues();


    teamValues.forEach(
      function(row) {


        // A：借書證
        const libraryCard =
          String(row[0] || "")
            .trim()
            .toUpperCase();


        // B：小隊名稱
        const teamName =
          String(row[1] || "")
            .trim();


        // C：隊員一
        const member1Id =
          String(row[2] || "")
            .trim();


        // D：隊員二
        const member2Id =
          String(row[3] || "")
            .trim();


        // F：班級
        const rowClass =
          String(row[5] || "")
            .trim();



        // ======================================
        // 只把目前班級的小隊加入排行榜
        // ======================================

        if (
          libraryCard &&
          teamName &&
          rowClass === classId
        ) {

          teams.push({

            libraryCard:
              libraryCard,

            teamName:
              teamName,

            member1Id:
              member1Id,

            member2Id:
              member2Id,

            foundBooks:
              new Set(),

            foundCount:
              0

          });

        }

      }
    );

  }



  // ==========================================
  // 建立快速查詢
  //
  // libraryCard → team
  // ==========================================

  const teamMap =
    new Map();


  teams.forEach(
    function(team) {

      teamMap.set(
        team.libraryCard,
        team
      );

    }
  );



  // ==========================================
  // 讀取尋寶紀錄
  // ==========================================

  const recordLastRow =
    recordsSheet.getLastRow();


  if (recordLastRow >= 2) {

    const records =
      recordsSheet
        .getRange(
          2,
          1,
          recordLastRow - 1,
          5
        )
        .getDisplayValues();


    records.forEach(
      function(row) {


        // A：借書證
        const libraryCard =
          String(row[0] || "")
            .trim()
            .toUpperCase();


        // B：書籍條碼
        const barcode =
          normalizeBarcode_(
            row[1]
          );


        // E：班級
        const recordClass =
          String(row[4] || "")
            .trim();


        const team =
          teamMap.get(
            libraryCard
          );



        // ======================================
        // 必須同時符合：
        //
        // 1. 是目前班級已註冊的小隊
        // 2. 紀錄也是目前班級
        // 3. 書是目前候選書之一
        // ======================================

        if (
          team &&
          recordClass === classId &&
          treasureSet.has(
            barcode
          )
        ) {

          // Set 可以防止同一本到重複計算
          team.foundBooks.add(
            barcode
          );

        }

      }
    );

  }



  // ==========================================
  // 計算每一隊進度
  // ==========================================

  teams.forEach(
    function(team) {


      team.foundCount =
        Math.min(
          team.foundBooks.size,
          requiredCount
        );


      // Set 不需要送給網頁
      delete team.foundBooks;

    }
  );



  // ==========================================
  // 排序
  //
  // 找到越多本 → 越前面
  //
  // 相同進度 → 小隊名稱排序
  // ==========================================

  teams.sort(
    function(a, b) {


      if (
        b.foundCount !==
        a.foundCount
      ) {

        return (
          b.foundCount -
          a.foundCount
        );

      }


      return a.teamName.localeCompare(
        b.teamName,
        "zh-Hant"
      );

    }
  );



  // ==========================================
  // 設定名次
  //
  // 例如：
  //
  // 10本 → 第1
  // 8本  → 第2
  // 8本  → 第2
  // 6本  → 第4
  // ==========================================

  let previousCount =
    null;


  let previousRank =
    0;


  teams.forEach(
    function(team, index) {


      if (
        team.foundCount !==
        previousCount
      ) {

        previousRank =
          index + 1;

      }


      team.rank =
        previousRank;


      team.totalCount =
        requiredCount;


      team.completed =
        team.foundCount >=
        requiredCount;


      previousCount =
        team.foundCount;

    }
  );



  // ==========================================
  // 回傳排行榜
  // ==========================================

  return {

    classId:
      classId,

    className:
      CLASS_NAMES[classId],

    teams:
      teams,

    teamCount:
      teams.length,

    requiredCount:
      requiredCount,

    candidateBookCount:
      treasureBooks.length,

    updatedAt:
      Utilities.formatDate(
        new Date(),
        spreadsheet
          .getSpreadsheetTimeZone(),
        "HH:mm:ss"
      )

  };

}

// ======================================================
// 取得網站網址的 function
// ======================================================

function getWebAppUrl() {

  return ScriptApp
    .getService()
    .getUrl();

}


// ======================================================
// BorrowRecords 增加 class_id 的 helper
// ======================================================

function getBorrowRecordsSheet() {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();


  const sheet =
    spreadsheet.getSheetByName(
      "BorrowRecords"
    );


  if (!sheet) {

    throw new Error(
      "找不到 BorrowRecords 工作表"
    );

  }


  if (
    sheet.getRange(1, 5).getValue() === ""
  ) {

    sheet
      .getRange(1, 5)
      .setValue("class_id");

  }


  sheet
    .getRange("A:B")
    .setNumberFormat("@");


  sheet
    .getRange("E:E")
    .setNumberFormat("@");


  return sheet;

}


