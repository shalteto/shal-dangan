function doGet() {
    return HtmlService.createTemplateFromFile('index')
        .evaluate()
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setTitle('弾台帳アプリ');
}

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 環境構築用スクリプト
 * 必要なシートが存在しない場合に作成し、ヘッダーを設定します。
 */
function initialSetup() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = [
        { name: 'main', headers: ['date', 'use', 'bullet_type', 'category', 'quantity', 'place', 'gun', 'note'] },
        { name: 'gun', headers: ['gun', 'type', 'size'] },
        { name: 'bullet_type', headers: ['bullet_type', 'size', 'type', 'category'] },
        { name: 'place', headers: ['place', 'type'] },
        { name: 'use', headers: ['use'] }
    ];

    sheets.forEach(setting => {
        let sheet = ss.getSheetByName(setting.name);
        if (!sheet) {
            sheet = ss.insertSheet(setting.name);
            sheet.appendRow(setting.headers);
        }
    });
}

/**
 * マスタデータと現在の在庫状況を取得する
 */
function getMetaData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Helper to get data from a sheet as an array of objects
    const getSheetData = (sheetName) => {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) return [];
        const rows = sheet.getDataRange().getValues();
        if (rows.length < 2) return [];
        const headers = rows.shift();
        return rows.map(row => {
            let obj = {};
            headers.forEach((header, i) => obj[header] = row[i]);
            return obj;
        });
    };

    const guns = getSheetData('gun');
    const bulletTypes = getSheetData('bullet_type');
    const places = getSheetData('place');
    const uses = getSheetData('use');

    // Calculate inventory from main sheet
    const mainData = getSheetData('main');
    const inventory = {}; // Key: "use|bullet_type", Value: Quantity

    mainData.forEach(record => {
        if (!record.use || !record.bullet_type || typeof record.quantity !== 'number') return;
        const key = `${record.use}|${record.bullet_type}`;
        if (!inventory[key]) inventory[key] = 0;
        inventory[key] += record.quantity;
    });

    return {
        guns: guns,
        bulletTypes: bulletTypes,
        places: places,
        uses: uses,
        inventory: inventory
    };
}

/**
 * データを登録する
 */
function registerData(data) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('main');

    const timestamp = new Date();
    const date = data.date ? new Date(data.date) : timestamp;

    // bullet_typeテーブルからcategoryを取得するヘルパー関数
    const getCategory = (bulletType) => {
        if (!bulletType) return '';
        const bulletSheet = ss.getSheetByName('bullet_type');
        if (!bulletSheet) return '';
        const bulletData = bulletSheet.getDataRange().getValues();
        if (bulletData.length < 2) return '';
        const headers = bulletData[0];
        const categoryIndex = headers.indexOf('category');
        const bulletTypeIndex = headers.indexOf('bullet_type');
        if (categoryIndex < 0 || bulletTypeIndex < 0) return '';
        for (let i = 1; i < bulletData.length; i++) {
            if (bulletData[i][bulletTypeIndex] === bulletType) {
                return bulletData[i][categoryIndex] || '';
            }
        }
        return '';
    };

    if (data.mode === 'transfer') {
        // 用途変更: 2レコード追加
        const category = getCategory(data.bulletType);
        // 1. 出庫 (From)
        sheet.appendRow([
            date,
            data.fromUse,
            data.bulletType,
            category,
            -1 * Math.abs(data.quantity), // 確実にマイナス
            '-', // place
            data.gun,
            '用途変更(出)'
        ]);

        // 2. 入庫 (To)
        sheet.appendRow([
            date,
            data.toUse,
            data.bulletType,
            category,
            Math.abs(data.quantity), // 確実にプラス
            '-', // place
            data.gun,
            '用途変更(入)'
        ]);

    } else {
        // 購入 または 消費
        // 消費の場合はUI側でマイナス値を送ってくる前提だが、念のためモードで制御してもよい
        // ここでは送られてきた数値をそのまま信じる（UIで制御）
        // 弾消費なしの場合（狩猟・有害での出動記録）はcategoryをUIから受け取る
        const category = data.bulletType
            ? (getCategory(data.bulletType) || data.category || '')
            : (data.category || '');
        sheet.appendRow([
            date,
            data.use,
            data.bulletType || '',
            category,
            data.quantity !== undefined ? data.quantity : '',
            data.place,
            data.gun,
            data.results || ''
        ]);
    }

    return { success: true };
}

/**
 * mainテーブルのデータを取得する
 * @param {Object} filter - オプションのフィルター条件
 * @param {string} filter.startDate - 開始日 (yyyy-MM-dd形式、空欄で下限なし)
 * @param {string} filter.endDate - 終了日 (yyyy-MM-dd形式、空欄で上限なし)
 */
function getMainData(filter) {
    try {
        console.log('[getMainData] Starting...');
        console.log('[getMainData] Filter:', JSON.stringify(filter));
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        console.log('[getMainData] Spreadsheet:', ss ? ss.getName() : 'null');

        const sheet = ss.getSheetByName('main');
        console.log('[getMainData] Sheet:', sheet ? 'found' : 'not found');

        if (!sheet) return [];

        const rows = sheet.getDataRange().getValues();
        console.log('[getMainData] Total rows:', rows.length);

        if (rows.length < 2) return [];

        const headers = rows.shift();
        console.log('[getMainData] Headers:', headers);

        // 日付フィルター用の変換
        let startDate = null;
        let endDate = null;
        if (filter && filter.startDate) {
            startDate = new Date(filter.startDate);
            startDate.setHours(0, 0, 0, 0);
        }
        if (filter && filter.endDate) {
            endDate = new Date(filter.endDate);
            endDate.setHours(23, 59, 59, 999);
        }

        const dateColumnIndex = headers.indexOf('date');

        const result = [];
        rows.forEach((row, index) => {
            // 日付フィルター適用
            if (dateColumnIndex >= 0 && (startDate || endDate)) {
                const rowDate = row[dateColumnIndex];
                if (rowDate instanceof Date) {
                    if (startDate && rowDate < startDate) return;
                    if (endDate && rowDate > endDate) return;
                } else if (typeof rowDate === 'string' && rowDate) {
                    const parsedDate = new Date(rowDate);
                    if (!isNaN(parsedDate.getTime())) {
                        if (startDate && parsedDate < startDate) return;
                        if (endDate && parsedDate > endDate) return;
                    }
                }
            }

            let obj = { rowIndex: index };  // 0-based index from data
            headers.forEach((header, i) => {
                let value = row[i];
                // 日付オブジェクトを文字列に変換
                if (value instanceof Date) {
                    value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
                }
                // undefinedをnullまたは空文字に変換
                if (value === undefined) {
                    value = '';
                }
                obj[header] = value;
            });
            result.push(obj);
        });

        console.log('[getMainData] Result count:', result.length);

        // JSON.stringify + JSON.parseで確実にシリアライズ可能な形式に変換
        const serialized = JSON.parse(JSON.stringify(result));
        console.log('[getMainData] Serialized successfully');
        return serialized;
    } catch (e) {
        console.error('[getMainData] Error:', e.message);
        throw e;
    }
}

/**
 * mainテーブルのレコードを削除する
 */
function deleteMainRecord(rowIndex) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('main');

    // rowIndex is 0-based from data array
    // Actual sheet row = rowIndex + 2 (header is row 1)
    if (typeof rowIndex === 'number' && rowIndex >= 0) {
        sheet.deleteRow(rowIndex + 2);
    }

    return { success: true };
}

/**
 * マスタデータを更新する (CRUD)
 */
function updateMasterData(sheetName, action, data) {
    const allowedSheets = ['gun', 'bullet_type', 'place', 'use'];
    if (!allowedSheets.includes(sheetName)) {
        throw new Error('Invalid sheet name');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    switch (action) {
        case 'add':
            // data is an object corresponding to headers
            // We need to map object values to header order
            const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
            const row = headers.map(h => data[h] !== undefined ? String(data[h]) : '');
            const newRowIndex = sheet.getLastRow() + 1;
            const range = sheet.getRange(newRowIndex, 1, 1, row.length);
            range.setNumberFormat('@'); // テキスト形式に設定
            range.setValues([row]);
            break;

        case 'delete':
            // data contains 'rowIndex' (0-based index from the data array, so actual row is index + 2)
            // client side sends 0-based index of the data list
            // The sheet has headers, so:
            // Data index 0 -> Row 2
            // Data index N -> Row N+2
            if (typeof data.rowIndex === 'number') {
                sheet.deleteRow(data.rowIndex + 2);
            }
            break;

        case 'update':
            // Simple update: We might just delete and re-add, or update specific cell
            // For simplicity/robustness in this app, we can receive rowIndex and new data
            if (typeof data.rowIndex === 'number') {
                const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
                const rowValues = [headers.map(h => data.values[h])];
                sheet.getRange(data.rowIndex + 2, 1, 1, rowValues[0].length).setValues(rowValues);
            }
            break;
    }

    return { success: true };
}

/**
 * mainテーブルのレコードを更新する
 */
function updateMainRecord(data) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('main');

    // rowIndex checks
    if (typeof data.rowIndex !== 'number' || data.rowIndex < 0) {
        throw new Error('Invalid rowIndex');
    }

    const rowNumber = data.rowIndex + 2; // header is row 1, data starts at row 2, rowIndex 0 is row 2

    // date
    const date = data.date ? new Date(data.date) : new Date();

    // get category
    // bullet_typeテーブルからcategoryを取得する helper (reuse logic from registerData or duplicate)
    // defined inside registerData, let's redefine or make a helper function. 
    // Since it's inside registerData, I can't call it. I'll duplicate logic for now to keep it simple or make a shared helper.
    // Making a shared helper is better but requires refactoring registerData. 
    // I will duplicate strictly the necessary part for now to minimize diff, or just implementing it here.

    const getCategory = (bulletType) => {
        if (!bulletType) return '';
        const bulletSheet = ss.getSheetByName('bullet_type');
        if (!bulletSheet) return '';
        const bulletData = bulletSheet.getDataRange().getValues();
        if (bulletData.length < 2) return '';
        const headers = bulletData[0];
        const categoryIndex = headers.indexOf('category');
        const bulletTypeIndex = headers.indexOf('bullet_type');
        if (categoryIndex < 0 || bulletTypeIndex < 0) return '';
        for (let i = 1; i < bulletData.length; i++) {
            if (bulletData[i][bulletTypeIndex] === bulletType) {
                return bulletData[i][categoryIndex] || '';
            }
        }
        return '';
    };

    const category = data.bullet_type
        ? (getCategory(data.bullet_type) || data.category || '')
        : (data.category || '');

    // Headers: ['date', 'use', 'bullet_type', 'category', 'quantity', 'place', 'gun', 'note']
    const rowValues = [
        date,
        data.use,
        data.bullet_type || '',
        category,
        data.quantity,
        data.place,
        data.gun,
        data.note || ''
    ];

    sheet.getRange(rowNumber, 1, 1, 8).setValues([rowValues]);

    return { success: true };
}
