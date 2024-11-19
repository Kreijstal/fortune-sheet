import _ from "lodash";
import { addSheet as addSheetInternal, deleteSheet as deleteSheetInternal, updateSheet as updateSheetInternal, } from "../modules";
import { getSheet } from "./common";
import { INVALID_PARAMS } from "./errors";
export function addSheet(ctx, settings, newSheetID, isPivotTable, sheetname, sheetData) {
    if (isPivotTable === void 0) { isPivotTable = false; }
    if (sheetname === void 0) { sheetname = undefined; }
    if (sheetData === void 0) { sheetData = undefined; }
    addSheetInternal(ctx, settings, newSheetID, isPivotTable, sheetname, sheetData);
}
export function deleteSheet(ctx, options) {
    if (options === void 0) { options = {}; }
    var sheet = getSheet(ctx, options);
    deleteSheetInternal(ctx, sheet.id);
}
export function updateSheet(ctx, data) {
    updateSheetInternal(ctx, data);
}
export function activateSheet(ctx, options) {
    if (options === void 0) { options = {}; }
    var sheet = getSheet(ctx, options);
    ctx.currentSheetId = sheet.id;
}
export function setSheetName(ctx, name, options) {
    if (options === void 0) { options = {}; }
    var sheet = getSheet(ctx, options);
    sheet.name = name;
}
export function setSheetOrder(ctx, orderList) {
    var _a;
    (_a = ctx.luckysheetfile) === null || _a === void 0 ? void 0 : _a.forEach(function (sheet) {
        if (sheet.id in orderList) {
            sheet.order = orderList[sheet.id];
        }
    });
    // re-order starting from 0
    _.sortBy(ctx.luckysheetfile, ["order"]).forEach(function (sheet, i) {
        sheet.order = i;
    });
}
export function scroll(ctx, scrollbarX, scrollbarY, options) {
    if (options.scrollLeft != null) {
        if (!_.isNumber(options.scrollLeft)) {
            throw INVALID_PARAMS;
        }
        if (scrollbarX) {
            scrollbarX.scrollLeft = options.scrollLeft;
        }
    }
    else if (options.targetColumn != null) {
        if (!_.isNumber(options.targetColumn)) {
            throw INVALID_PARAMS;
        }
        var col_pre = options.targetColumn <= 0
            ? 0
            : ctx.visibledatacolumn[options.targetColumn - 1];
        if (scrollbarX) {
            scrollbarX.scrollLeft = col_pre;
        }
    }
    if (options.scrollTop != null) {
        if (!_.isNumber(options.scrollTop)) {
            throw INVALID_PARAMS;
        }
        if (scrollbarY) {
            scrollbarY.scrollTop = options.scrollTop;
        }
    }
    else if (options.targetRow != null) {
        if (!_.isNumber(options.targetRow)) {
            throw INVALID_PARAMS;
        }
        var row_pre = options.targetRow <= 0 ? 0 : ctx.visibledatarow[options.targetRow - 1];
        if (scrollbarY) {
            scrollbarY.scrollTop = row_pre;
        }
    }
}
