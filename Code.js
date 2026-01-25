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
        { name: 'main', headers: ['date', 'use', 'bullet_type', 'quantity', 'place', 'gun', 'hunting results'] },
        { name: 'gun', headers: ['gun', 'type', 'size'] },
        { name: 'bullet_type', headers: ['bullet_type', 'size', 'type'] },
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

    if (data.mode === 'transfer') {
        // 用途変更: 2レコード追加
        // 1. 出庫 (From)
        sheet.appendRow([
            date,
            data.fromUse,
            data.bulletType,
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
            Math.abs(data.quantity), // 確実にプラス
            '-', // place
            data.gun,
            '用途変更(入)'
        ]);

    } else {
        // 購入 または 消費
        // 消費の場合はUI側でマイナス値を送ってくる前提だが、念のためモードで制御してもよい
        // ここでは送られてきた数値をそのまま信じる（UIで制御）
        sheet.appendRow([
            date,
            data.use,
            data.bulletType,
            data.quantity,
            data.place,
            data.gun,
            data.results || ''
        ]);
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
            const row = headers.map(h => data[h]);
            sheet.appendRow(row);
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
