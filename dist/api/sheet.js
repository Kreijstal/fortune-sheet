import _ from "lodash";
import { v4 as uuidv4 } from "uuid";
import { dataToCelldata, getSheet } from "./common";
import { getSheetIndex } from "../utils";
import { api, execfunction, insertUpdateFunctionGroup, locale } from "..";
export function getAllSheets(ctx) {
    return ctx.luckysheetfile;
}
export { getSheet };
export function initSheetData(draftCtx, index, newData) {
    var _a, _b;
    var celldata = newData.celldata, row = newData.row, column = newData.column;
    var lastRow = _.maxBy(celldata, "r");
    var lastCol = _.maxBy(celldata, "c");
    var lastRowNum = ((_a = lastRow === null || lastRow === void 0 ? void 0 : lastRow.r) !== null && _a !== void 0 ? _a : 0) + 1;
    var lastColNum = ((_b = lastCol === null || lastCol === void 0 ? void 0 : lastCol.c) !== null && _b !== void 0 ? _b : 0) + 1;
    if (row != null && column != null && row > 0 && column > 0) {
        lastRowNum = Math.max(lastRowNum, row);
        lastColNum = Math.max(lastColNum, column);
    }
    else {
        lastRowNum = Math.max(lastRowNum, draftCtx.defaultrowNum);
        lastColNum = Math.max(lastColNum, draftCtx.defaultcolumnNum);
    }
    if (lastRowNum && lastColNum) {
        var expandedData_1 = _.times(lastRowNum, function () {
            return _.times(lastColNum, function () { return null; });
        });
        celldata === null || celldata === void 0 ? void 0 : celldata.forEach(function (d) {
            expandedData_1[d.r][d.c] = d.v;
        });
        if (draftCtx.luckysheetfile[index] == null) {
            newData.data = expandedData_1;
            delete newData.celldata;
            draftCtx.luckysheetfile.push(newData);
        }
        else {
            draftCtx.luckysheetfile[index].data = expandedData_1;
            delete draftCtx.luckysheetfile[index].celldata;
        }
        return expandedData_1;
    }
    return null;
}
export function hideSheet(ctx, sheetId) {
    var index = getSheetIndex(ctx, sheetId);
    ctx.luckysheetfile[index].hide = 1;
    ctx.luckysheetfile[index].status = 0;
    var shownSheets = ctx.luckysheetfile.filter(function (sheet) { return _.isUndefined(sheet.hide) || (sheet === null || sheet === void 0 ? void 0 : sheet.hide) !== 1; });
    ctx.currentSheetId = shownSheets[0].id;
}
export function showSheet(ctx, sheetId) {
    var index = getSheetIndex(ctx, sheetId);
    ctx.luckysheetfile[index].hide = undefined;
}
function generateCopySheetName(ctx, sheetId) {
    var info = locale(ctx).info;
    var copyWord = "(".concat(info.copy);
    var SheetIndex = getSheetIndex(ctx, sheetId);
    var sheetName = ctx.luckysheetfile[SheetIndex].name;
    var copy_i = sheetName.indexOf(copyWord);
    var index = 0;
    if (copy_i !== -1) {
        sheetName = sheetName.toString().substring(0, copy_i);
    }
    var nameCopy = sheetName + copyWord;
    var sheetNames = [];
    for (var i = 0; i < ctx.luckysheetfile.length; i += 1) {
        var fileName = ctx.luckysheetfile[i].name;
        sheetNames.push(fileName);
        var st_i = fileName.indexOf(nameCopy);
        if (st_i === 0) {
            index = index || 2;
            var ed_i = fileName.indexOf(")", st_i + nameCopy.length);
            var num = fileName.substring(st_i + nameCopy.length, ed_i);
            if (_.isNumber(num)) {
                if (Number.parseInt(num, 10) >= index) {
                    index = Number.parseInt(num, 10) + 1;
                }
            }
        }
    }
    var sheetCopyName;
    do {
        var postfix = "".concat(copyWord + (index || ""), ")");
        var lengthLimit = 31 - postfix.length;
        sheetCopyName = sheetName;
        if (sheetCopyName.length > lengthLimit) {
            sheetCopyName = "".concat(sheetCopyName.slice(0, lengthLimit - 1), "\u2026");
        }
        sheetCopyName += postfix;
        index = (index || 1) + 1;
    } while (sheetNames.indexOf(sheetCopyName) !== -1);
    return sheetCopyName;
}
export function copySheet(ctx, sheetId) {
    var index = getSheetIndex(ctx, sheetId);
    var order = ctx.luckysheetfile[index].order + 1;
    var sheetName = generateCopySheetName(ctx, sheetId);
    var sheetData = _.cloneDeep(ctx.luckysheetfile[index]);
    delete sheetData.id;
    delete sheetData.status;
    sheetData.celldata = dataToCelldata(sheetData.data);
    delete sheetData.data;
    api.addSheet(ctx, undefined, uuidv4(), ctx.luckysheetfile[index].isPivotTable, sheetName, sheetData);
    var sheetOrderList = {};
    sheetOrderList[ctx.luckysheetfile[ctx.luckysheetfile.length - 1].id] = order;
    api.setSheetOrder(ctx, sheetOrderList);
}
export function calculateSheetFromula(ctx, id) {
    var _a, _b;
    var index = getSheetIndex(ctx, id);
    if (!ctx.luckysheetfile[index].data)
        return;
    for (var r = 0; r < ctx.luckysheetfile[index].data.length; r += 1) {
        for (var c = 0; c < ctx.luckysheetfile[index].data[r].length; c += 1) {
            if (!((_a = ctx.luckysheetfile[index].data[r][c]) === null || _a === void 0 ? void 0 : _a.f)) {
                continue;
            }
            var result = execfunction(ctx, (_b = ctx.luckysheetfile[index].data[r][c]) === null || _b === void 0 ? void 0 : _b.f, r, c, id);
            api.setCellValue(ctx, r, c, result[1], null);
            insertUpdateFunctionGroup(ctx, r, c, id);
        }
    }
}
