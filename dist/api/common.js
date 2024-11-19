var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
import _ from "lodash";
import { getSheetIndex } from "../utils";
import { SHEET_NOT_FOUND } from "./errors";
export var dataToCelldata = function (data) {
    var celldata = [];
    if (data == null) {
        return celldata;
    }
    for (var r = 0; r < data.length; r += 1) {
        for (var c = 0; c < data[r].length; c += 1) {
            var v = data[r][c];
            if (v != null) {
                celldata.push({ r: r, c: c, v: v });
            }
        }
    }
    return celldata;
};
export var celldataToData = function (celldata, rowCount, colCount) {
    var _a, _b;
    var lastRow = _.maxBy(celldata, "r");
    var lastCol = _.maxBy(celldata, "c");
    var lastRowNum = ((_a = lastRow === null || lastRow === void 0 ? void 0 : lastRow.r) !== null && _a !== void 0 ? _a : 0) + 1;
    var lastColNum = ((_b = lastCol === null || lastCol === void 0 ? void 0 : lastCol.c) !== null && _b !== void 0 ? _b : 0) + 1;
    if (rowCount != null && colCount != null && rowCount > 0 && colCount > 0) {
        lastRowNum = Math.max(lastRowNum, rowCount);
        lastColNum = Math.max(lastColNum, colCount);
    }
    if (lastRowNum && lastColNum) {
        var expandedData_1 = _.times(lastRowNum, function () {
            return _.times(lastColNum, function () { return null; });
        });
        celldata === null || celldata === void 0 ? void 0 : celldata.forEach(function (d) {
            expandedData_1[d.r][d.c] = d.v;
        });
        return expandedData_1;
    }
    return null;
};
export function getSheet(ctx, options) {
    if (options === void 0) { options = {}; }
    var _a = options.index, index = _a === void 0 ? getSheetIndex(ctx, options.id || ctx.currentSheetId) : _a;
    if (index == null) {
        throw SHEET_NOT_FOUND;
    }
    var sheet = ctx.luckysheetfile[index];
    if (sheet == null) {
        throw SHEET_NOT_FOUND;
    }
    return sheet;
}
export function getSheetWithLatestCelldata(ctx, options) {
    if (options === void 0) { options = {}; }
    var sheet = getSheet(ctx, options);
    return __assign(__assign({}, sheet), { celldata: dataToCelldata(sheet.data) });
}
