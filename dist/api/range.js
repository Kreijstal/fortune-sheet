import _ from "lodash";
import { getdatabyselection, getFlowdata, getRangetxt } from "..";
import { normalizeSelection, rangeValueToHtml } from "../modules";
import { setCellFormat, setCellValue } from "./cell";
import { getSheet } from "./common";
import { INVALID_PARAMS } from "./errors";
export function getSelection(ctx) {
    var _a;
    return (_a = ctx.luckysheet_select_save) === null || _a === void 0 ? void 0 : _a.map(function (selection) { return ({
        row: selection.row,
        column: selection.column,
    }); });
}
export function getFlattenRange(ctx, range) {
    range = range || getSelection(ctx);
    var result = [];
    range === null || range === void 0 ? void 0 : range.forEach(function (ele) {
        var rs = ele.row;
        var cs = ele.column;
        for (var r = rs[0]; r <= rs[1]; r += 1) {
            for (var c = cs[0]; c <= cs[1]; c += 1) {
                result.push({ r: r, c: c });
            }
        }
    });
    return result;
}
export function getCellsByFlattenRange(ctx, range) {
    range = range || getFlattenRange(ctx);
    var flowdata = getFlowdata(ctx);
    if (!flowdata)
        return [];
    return range.map(function (item) { var _a; return (_a = flowdata[item.r]) === null || _a === void 0 ? void 0 : _a[item.c]; });
}
export function getSelectionCoordinates(ctx) {
    var result = [];
    var rangeArr = _.cloneDeep(ctx.luckysheet_select_save);
    var sheetId = ctx.currentSheetId;
    rangeArr === null || rangeArr === void 0 ? void 0 : rangeArr.forEach(function (ele) {
        var rangeText = getRangetxt(ctx, sheetId, {
            column: ele.column,
            row: ele.row,
        });
        result.push(rangeText);
    });
    return result;
}
export function getCellsByRange(ctx, range, options) {
    if (options === void 0) { options = {}; }
    var sheet = getSheet(ctx, options);
    if (!range || typeof range === "object") {
        return getdatabyselection(ctx, range, sheet.id);
    }
    throw INVALID_PARAMS;
}
export function getHtmlByRange(ctx, range, options) {
    if (options === void 0) { options = {}; }
    var sheet = getSheet(ctx, options);
    return rangeValueToHtml(ctx, sheet.id, range);
}
export function setSelection(ctx, range, options) {
    var sheet = getSheet(ctx, options);
    sheet.luckysheet_select_save = normalizeSelection(ctx, range);
    if (ctx.currentSheetId === sheet.id) {
        ctx.luckysheet_select_save = sheet.luckysheet_select_save;
    }
}
export function setCellValuesByRange(ctx, data, range, cellInput, options) {
    if (options === void 0) { options = {}; }
    if (data == null) {
        throw INVALID_PARAMS;
    }
    if (range instanceof Array) {
        throw new Error("setCellValuesByRange does not support multiple ranges");
    }
    if (!_.isPlainObject(range)) {
        throw INVALID_PARAMS;
    }
    var rowCount = range.row[1] - range.row[0] + 1;
    var columnCount = range.column[1] - range.column[0] + 1;
    if (data.length !== rowCount || data[0].length !== columnCount) {
        throw new Error("data size does not match range");
    }
    for (var i = 0; i < rowCount; i += 1) {
        for (var j = 0; j < columnCount; j += 1) {
            var row = range.row[0] + i;
            var column = range.column[0] + j;
            setCellValue(ctx, row, column, data[i][j], cellInput, options);
        }
    }
}
export function setCellFormatByRange(ctx, attr, value, range, options) {
    if (options === void 0) { options = {}; }
    if (_.isPlainObject(range)) {
        range = [range];
    }
    if (!_.isArray(range)) {
        throw INVALID_PARAMS;
    }
    range.forEach(function (singleRange) {
        for (var r = singleRange.row[0]; r <= singleRange.row[1]; r += 1) {
            for (var c = singleRange.column[0]; c <= singleRange.column[1]; c += 1) {
                setCellFormat(ctx, r, c, attr, value, options);
            }
        }
    });
}
