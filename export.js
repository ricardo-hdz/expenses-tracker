/**
 * Apps Script to process and export personal expenses from AndroMoney to Expense Tracker.
 */

const MONTHS = [
    'Test',
    'Jan',
    'Feb',
    'Mar',
    'Apr',
    'May',
    'Jun',
    'Jul',
    'Aug',
    'Sep',
    'Oct',
    'Nov',
    'Dec'
];

// 2016_ID
// const TRACKER_ID = '1AweSw57sw-a33ijYq9n7aQOR3xuizXeJoYExPKfMTQE';
// 2017_ID
// const TRACKER_ID = '1NZWrWFLTJYJmvA_T0gq4--7DJc1L_GaF5XcS0N3Pgps';
// Europe_2017
// const TRACKER_ID = '1FyMr67nUW6yFL8NJFpguUCft9Q6f5RbC4CcsTIXfuME';
// 2018 ID
// const TRACKER_ID = '16ch-qMh1XGIVHZkRoAA4Rw4HUNqe70oigoGiVNyu8JQ';
// 2019 ID
// const TRACKER_ID = '1HxKjp5JYU-GOtrnuJZk6OYJucPt_RkhytqUL9TZMazQ';
// 2020 ID
// const TRACKER_ID = '1DS3NPcwPnoEpMV_34rTAMQ6SXYMc4HpDvR1vdntSPAE';
// 2021 ID
// const TRACKER_ID = '1LOU4wgl0ukA7xUJ7ga_n0IuCP7poqxuYJFEW7se0U34';
// 2022 ID
// const TRACKER_ID = '1ix0yIQzJf_NhRZorctvb7IF5KylRwmOmOOayahA7v8Q';
// 2023 ID
// const TRACKER_ID = '1in6WsDLWRnifcI2EWw1Ie-4G0-3yz3HuVJ15Tb1z1tM';
// 2024 ID
// const TRACKER_ID = '1BjOdIrgJzRntCwl189c4_of_9k6KcjxdWyZWPwMhiI4';
// 2025 ID
// const TRACKER_ID = '1CUGKlMVP38Lw_UuyY2pz-mgTM4Qryx1iblhBdaY-mCw';
// 2026 ID
const TRACKER_ID = '17yWFIKrKGFWDCyY4lcsFwICZAk5b9x3X_BtYVRpHWbY';

const SHEET_SOURCE = 'AndroMoney';
const SHEET_FORMATTED = 'Formatted';

const FORMATTED_HEADERS = [
    'Date',
    'Type',
    'Category',
    'Subcategory',
    'Vendor',
    'Payment',
    'Currency',
    'Amount',
    'Note'
];

// 0-indexed column positions in raw AndroMoney sheet (data starts on Row 3)
const COL_INDEX = {
    ID: 0,          // Col A (1)
    CURRENCY: 1,    // Col B (2)
    AMOUNT: 2,      // Col C (3)
    CATEGORY: 3,    // Col D (4)
    SUBCATEGORY: 4, // Col E (5)
    DATE: 5,        // Col F (6)
    PAYMENT: 6,     // Col G (7)
    NOTE: 8,        // Col I (9)
    PROJECT: 10,    // Col K (11)
    VENDOR: 11,     // Col L (12)
    TYPE: 14        // Col O (15)
};

const ENDPOINT_RATES = 'https://apilayer.net/api/live?access_key=b9f923b9b69e956ea34daa10694fc9b1&source=USD&currencies={currency}&format=1';

const exchangeRates = {};

// ==========================================
// Application Triggers & UI Menu
// ==========================================

function onOpen() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    getOrCreateFormattedSheet(spreadsheet);

    const sourceSheet = spreadsheet.getSheetByName(SHEET_SOURCE);
    if (sourceSheet) {
        spreadsheet.setActiveSheet(sourceSheet);
    }

    const menuEntries = [
        { name: "Format & Transfer Data", functionName: "formatAndTransferData" },
        { name: "Format Data", functionName: "formatData" },
    ];
    spreadsheet.addMenu("AndroidMoney", menuEntries);
}

// ==========================================
// Main Workflow Functions
// ==========================================

/**
 * Formats raw AndroMoney data and transfers it directly to the annual Tracker spreadsheet.
 */
function formatAndTransferData() {
    const result = formatData();
    if (result && result.formattedRows && result.formattedRows.length > 0) {
        copyDataToTracker(result.formattedRows, result.detectedMonth);
    }
}

/**
 * Reads raw AndroMoney sheet, filters out business rows, normalizes dates,
 * handles currency conversion, and writes cleaned data to 'Formatted' sheet in a single batch.
 * @returns {{ formattedRows: Array<Array>, detectedMonth: string|null } | null}
 */
function formatData() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = spreadsheet.getSheetByName(SHEET_SOURCE);
    if (!sourceSheet) {
        SpreadsheetApp.getUi().alert(`Sheet '${SHEET_SOURCE}' not found. Please ensure your raw CSV is imported there.`);
        return null;
    }

    const rawData = sourceSheet.getDataRange().getValues();
    if (!rawData || rawData.length < 3) {
        SpreadsheetApp.getUi().alert(`No data rows found in '${SHEET_SOURCE}'. Expected data starting at row 3.`);
        return null;
    }

    const result = processRawData(rawData);
    if (!result.formattedRows || result.formattedRows.length === 0) {
        SpreadsheetApp.getUi().alert('No personal expenses found after filtering.');
        return null;
    }

    const formattedSheet = getOrCreateFormattedSheet(spreadsheet);
    formattedSheet.clearContents();

    const outputData = [FORMATTED_HEADERS, ...result.formattedRows];
    formattedSheet.getRange(1, 1, outputData.length, FORMATTED_HEADERS.length).setValues(outputData);

    const monthLabel = result.detectedMonth ? ` (${result.detectedMonth})` : '';
    spreadsheet.toast(`Formatted ${result.formattedRows.length} transactions${monthLabel}.`, 'Formatting Complete', 5);

    return result;
}

/**
 * Copies formatted data into the target month sheet of the annual Tracker spreadsheet.
 * Clears old transaction entries in columns A:I while preserving row 1 headers and columns J:L formulas/charts.
 * @param {Array<Array>} [formattedRows] Optional pre-processed rows. If omitted, reads from 'Formatted' sheet.
 * @param {string} [targetMonth] Optional target month name. If omitted, derives from data.
 */
function copyDataToTracker(formattedRows, targetMonth) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    if (!formattedRows) {
        const formattedSheet = spreadsheet.getSheetByName(SHEET_FORMATTED);
        if (!formattedSheet) {
            throw new Error(`Sheet '${SHEET_FORMATTED}' not found.`);
        }
        const data = formattedSheet.getDataRange().getValues();
        if (data.length <= 1) {
            SpreadsheetApp.getUi().alert('No formatted data available to transfer.');
            return;
        }
        formattedRows = data.slice(1);
    }

    if (!formattedRows || formattedRows.length === 0) {
        SpreadsheetApp.getUi().alert('No transaction rows to copy to tracker.');
        return;
    }

    const monthName = targetMonth || getCurrentMonth();
    if (!monthName) {
        SpreadsheetApp.getUi().alert('Unable to determine the target month from transactions.');
        return;
    }

    const trackerSpreadsheet = SpreadsheetApp.openById(TRACKER_ID);
    const targetSheet = trackerSpreadsheet.getSheetByName(monthName);
    if (!targetSheet) {
        throw new Error(`Target sheet '${monthName}' not found in Tracker spreadsheet (ID: ${TRACKER_ID}).`);
    }

    // Clear existing data in columns A to I (rows 2..lastRow)
    const existingLastRow = targetSheet.getLastRow();
    if (existingLastRow >= 2) {
        targetSheet.getRange(2, 1, existingLastRow - 1, FORMATTED_HEADERS.length).clearContent();
    }

    // Write all new transactions in a single batch starting at Row 2, Column 1
    targetSheet.getRange(2, 1, formattedRows.length, FORMATTED_HEADERS.length).setValues(formattedRows);

    spreadsheet.toast(
        `Successfully transferred ${formattedRows.length} transactions to '${monthName}'.`,
        'Tracker Updated',
        5
    );
}

// ==========================================
// In-Memory Data Transformation
// ==========================================

/**
 * Processes raw CSV 2D array in memory: filters Business rows, normalizes dates,
 * handles income/expense signs, and converts currencies.
 * @param {Array<Array>} rawData
 * @returns {{ formattedRows: Array<Array>, detectedMonth: string|null }}
 */
function processRawData(rawData) {
    // Data rows start from row index 2 (row 3 in sheet)
    const dataRows = rawData.slice(2);
    const formattedRows = [];
    let detectedMonth = null;

    for (const row of dataRows) {
        if (!row || row.length === 0) continue;

        const rawDate = row[COL_INDEX.DATE];
        const rawAmount = row[COL_INDEX.AMOUNT];

        // Skip completely empty rows
        if ((rawDate === '' || rawDate === undefined) && (rawAmount === '' || rawAmount === undefined)) {
            continue;
        }

        // Filter out business expenses
        const project = String(row[COL_INDEX.PROJECT] || '').trim();
        if (project.toLowerCase() === 'business') {
            continue;
        }

        // Normalize date to MM/dd/yyyy
        const formattedDate = formatDateString(rawDate);
        if (!detectedMonth) {
            detectedMonth = getMonthNameFromDate(rawDate);
        }

        // Parse amount
        let numAmount = typeof rawAmount === 'number'
            ? rawAmount
            : parseFloat(String(rawAmount).replace(/[^0-9.-]/g, ''));
        if (isNaN(numAmount)) {
            numAmount = 0;
        }

        // Expenses are negative, income (empty payment field) is positive
        const payment = String(row[COL_INDEX.PAYMENT] || '').trim();
        let finalAmount = numAmount === 0
            ? 0
            : (payment !== '' ? -Math.abs(numAmount) : Math.abs(numAmount));

        // Convert currency to USD if necessary
        let currency = String(row[COL_INDEX.CURRENCY] || 'USD').trim().toUpperCase();
        if (currency !== 'USD' && currency !== '') {
            const rate = getLatestExchangeRate(currency);
            if (rate && rate > 0) {
                finalAmount = Math.round((finalAmount / rate) * 100) / 100;
                currency = 'USD';
            } else {
                Logger.log(`Skipping currency conversion for ${currency} due to unavailable rate.`);
            }
        } else {
            currency = 'USD';
        }

        formattedRows.push([
            formattedDate,
            row[COL_INDEX.TYPE] || 'Personal',
            row[COL_INDEX.CATEGORY] || '',
            row[COL_INDEX.SUBCATEGORY] || '',
            row[COL_INDEX.VENDOR] || '',
            row[COL_INDEX.PAYMENT] || '',
            currency,
            finalAmount,
            row[COL_INDEX.NOTE] || ''
        ]);
    }

    return { formattedRows, detectedMonth };
}

// ==========================================
// Date & Currency Helpers
// ==========================================

/**
 * Safely parses various date formats (Date object, YYYYMMDD, YYYY-MM-DD, MM/DD/YYYY)
 * without timezone midnight shifts.
 * @param {*} rawDate
 * @returns {{ year: number, month: number, day: number } | null}
 */
function parseDateComponents(rawDate) {
    if (!rawDate) return null;

    if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
        return {
            year: rawDate.getFullYear(),
            month: rawDate.getMonth() + 1,
            day: rawDate.getDate()
        };
    }

    const str = String(rawDate).trim();

    // 8-digit YYYYMMDD
    const yyyymmddMatch = str.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (yyyymmddMatch) {
        return {
            year: parseInt(yyyymmddMatch[1], 10),
            month: parseInt(yyyymmddMatch[2], 10),
            day: parseInt(yyyymmddMatch[3], 10)
        };
    }

    // YYYY-MM-DD or YYYY/MM/DD
    const isoMatch = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (isoMatch) {
        return {
            year: parseInt(isoMatch[1], 10),
            month: parseInt(isoMatch[2], 10),
            day: parseInt(isoMatch[3], 10)
        };
    }

    // MM/DD/YYYY
    const usMatch = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
    if (usMatch) {
        return {
            year: parseInt(usMatch[3], 10),
            month: parseInt(usMatch[1], 10),
            day: parseInt(usMatch[2], 10)
        };
    }

    const fallback = new Date(str);
    if (!isNaN(fallback.getTime())) {
        return {
            year: fallback.getFullYear(),
            month: fallback.getMonth() + 1,
            day: fallback.getDate()
        };
    }

    return null;
}

/**
 * Returns formatted date string MM/dd/yyyy.
 * @param {*} rawDate
 * @returns {string}
 */
function formatDateString(rawDate) {
    const parts = parseDateComponents(rawDate);
    if (!parts) return String(rawDate || '');
    const mm = String(parts.month).padStart(2, '0');
    const dd = String(parts.day).padStart(2, '0');
    return `${mm}/${dd}/${parts.year}`;
}

/**
 * Returns 3-letter month abbreviation (e.g. 'Jan', 'Feb') from date.
 * @param {*} rawDate
 * @returns {string|null}
 */
function getMonthNameFromDate(rawDate) {
    const parts = parseDateComponents(rawDate);
    if (!parts || parts.month < 1 || parts.month > 12) {
        return null;
    }
    return MONTHS[parts.month];
}

/**
 * Determines current month from the 'Formatted' sheet.
 * @returns {string|null}
 */
function getCurrentMonth() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_FORMATTED);
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
        const monthName = getMonthNameFromDate(data[i][0]);
        if (monthName) return monthName;
    }
    return null;
}

/**
 * Fetches live exchange rate for currency (source USD). Caches results in memory.
 * @param {string} currency
 * @returns {number|null}
 */
function getLatestExchangeRate(currency) {
    if (exchangeRates[currency]) {
        return exchangeRates[currency];
    }

    const url = ENDPOINT_RATES.replace('{currency}', encodeURIComponent(currency));
    try {
        const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        const data = JSON.parse(response.getContentText());
        const key = 'USD' + currency;

        if (data && data.success && data.quotes && typeof data.quotes[key] === 'number') {
            const rate = data.quotes[key];
            if (rate > 0) {
                exchangeRates[currency] = rate;
                SpreadsheetApp.getActiveSpreadsheet().toast(
                    `Exchange rate for USD - ${currency}: ${rate}`,
                    'Exchange Rate',
                    5
                );
                return rate;
            }
        }
        Logger.log(`Failed to fetch valid exchange rate for ${currency}. Response: ${response.getContentText()}`);
    } catch (e) {
        Logger.log(`Error fetching exchange rate for ${currency}: ${e.message}`);
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(
        `Unable to fetch exchange rate for ${currency}. Keeping original amount.`,
        'Exchange Rate Warning',
        5
    );
    return null;
}

/**
 * Retrieves or inserts the 'Formatted' sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateFormattedSheet(spreadsheet) {
    let sheet = spreadsheet.getSheetByName(SHEET_FORMATTED);
    if (!sheet) {
        sheet = spreadsheet.insertSheet(SHEET_FORMATTED);
    }
    return sheet;
}

// ==========================================
// Backwards Compatibility Stubs
// ==========================================

function transformAmount() {
    Logger.log('transformAmount() is deprecated. Data transformation is now performed in-memory by formatData().');
}

function copyFormattedData() {
    formatData();
}