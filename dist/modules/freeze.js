import _ from "lodash";
import { colLocationByIndex, getSheetIndex, rowLocationByIndex, } from "..";
function cutVolumn(arr, cutindex) {
    if (cutindex <= 0) {
        return arr;
    }
    var ret = arr.slice(cutindex);
    return ret;
}
function frozenTofreezen(ctx, cache, sheetId) {
    // get frozen type
    var file = ctx.luckysheetfile[getSheetIndex(ctx, sheetId)];
    var frozen = file.frozen;
    if (frozen == null) {
        delete cache.freezen;
        return;
    }
    var freezen = {};
    var range = frozen.range;
    if (!range) {
        range = {
            row_focus: 0,
            column_focus: 0,
        };
    }
    var type = frozen.type;
    if (type === "row") {
        type = "rangeRow";
    }
    else if (type === "column") {
        type = "rangeColumn";
    }
    else if (type === "both") {
        type = "rangeBoth";
    }
    // transform to freezen
    if (type === "rangeRow" || type === "rangeBoth") {
        var scrollTop = 0;
        var row_st = _.sortedIndex(ctx.visibledatarow, scrollTop);
        var row_focus = range.row_focus;
        if (row_focus > row_st) {
            row_st = row_focus;
        }
        if (row_st === -1) {
            row_st = 0;
        }
        var top_1 = ctx.visibledatarow[row_st] - 2 - scrollTop + ctx.columnHeaderHeight;
        var freezenhorizontaldata = [
            ctx.visibledatarow[row_st],
            row_st + 1,
            scrollTop,
            cutVolumn(ctx.visibledatarow, row_st + 1),
            top_1,
        ];
        freezen.horizontal = {
            freezenhorizontaldata: freezenhorizontaldata,
            top: top_1,
        };
    }
    if (type === "rangeColumn" || type === "rangeBoth") {
        var scrollLeft = 0;
        var col_st = _.sortedIndex(ctx.visibledatacolumn, scrollLeft);
        var column_focus = range.column_focus;
        if (column_focus > col_st) {
            col_st = column_focus;
        }
        if (col_st === -1) {
            col_st = 0;
        }
        var left = ctx.visibledatacolumn[col_st] - 2 - scrollLeft + ctx.rowHeaderWidth;
        var freezenverticaldata = [
            ctx.visibledatacolumn[col_st],
            col_st + 1,
            scrollLeft,
            cutVolumn(ctx.visibledatacolumn, col_st + 1),
            left,
        ];
        freezen.vertical = {
            freezenverticaldata: freezenverticaldata,
            left: left,
        };
    }
    cache.freezen || (cache.freezen = {});
    cache.freezen[ctx.currentSheetId] = freezen;
}
export function initFreeze(ctx, cache, sheetId) {
    frozenTofreezen(ctx, cache, sheetId);
}
export function scrollToFrozenRowCol(ctx, freeze) {
    var _a, _b;
    var _c, _d;
    var select_save = ctx.luckysheet_select_save;
    if (!select_save)
        return;
    var row;
    var row_focus = select_save[0].row_focus;
    if (row_focus === select_save[0].row[0]) {
        _a = select_save[0].row, row = _a[1];
    }
    else if (row_focus === select_save[0].row[1]) {
        row = select_save[0].row[0];
    }
    var column;
    var column_focus = select_save[0].column_focus;
    if (column_focus === select_save[0].column[0]) {
        _b = select_save[0].column, column = _b[1];
    }
    else if (column_focus === select_save[0].column[1]) {
        column = select_save[0].column[0];
    }
    var freezenverticaldata = (_c = freeze === null || freeze === void 0 ? void 0 : freeze.vertical) === null || _c === void 0 ? void 0 : _c.freezenverticaldata;
    var freezenhorizontaldata = (_d = freeze === null || freeze === void 0 ? void 0 : freeze.horizontal) === null || _d === void 0 ? void 0 : _d.freezenhorizontaldata;
    if (freezenverticaldata != null && column != null) {
        var freezen_colindex = freezenverticaldata[1];
        var offset = _.sortedIndex(freezenverticaldata[3], ctx.scrollLeft);
        var top_2 = freezenverticaldata[4];
        freezen_colindex += offset;
        if (column >= ctx.visibledatacolumn.length) {
            column = ctx.visibledatacolumn.length - 1;
        }
        if (freezen_colindex >= ctx.visibledatacolumn.length) {
            freezen_colindex = ctx.visibledatacolumn.length - 1;
        }
        var column_px = ctx.visibledatacolumn[column];
        var freezen_px = ctx.visibledatacolumn[freezen_colindex];
        if (column_px <= freezen_px + top_2) {
            ctx.scrollLeft = 0;
            // setTimeout(function () {
            //   $("#luckysheet-scrollbar-x").scrollLeft(0);
            // }, 100);
        }
    }
    if (freezenhorizontaldata != null && row != null) {
        var freezen_rowindex = freezenhorizontaldata[1];
        var offset = _.sortedIndex(freezenhorizontaldata[3], ctx.scrollTop);
        var left = freezenhorizontaldata[4];
        freezen_rowindex += offset;
        if (row >= ctx.visibledatarow.length) {
            row = ctx.visibledatarow.length - 1;
        }
        if (freezen_rowindex >= ctx.visibledatarow.length) {
            freezen_rowindex = ctx.visibledatarow.length - 1;
        }
        var row_px = ctx.visibledatarow[row];
        var freezen_px = ctx.visibledatarow[freezen_rowindex];
        if (row_px <= freezen_px + left) {
            ctx.scrollTop = 0;
            // setTimeout(function () {
            //   $("#luckysheet-scrollbar-y").scrollTop(0);
            // }, 100);
        }
    }
}
export function getFrozenHandleTop(ctx) {
    var _a, _b, _c, _d, _e, _f;
    var idx = getSheetIndex(ctx, ctx.currentSheetId);
    if (idx == null)
        return ctx.scrollTop;
    var sheet = ctx.luckysheetfile[idx];
    if (((_a = sheet === null || sheet === void 0 ? void 0 : sheet.frozen) === null || _a === void 0 ? void 0 : _a.type) === "row" ||
        ((_b = sheet === null || sheet === void 0 ? void 0 : sheet.frozen) === null || _b === void 0 ? void 0 : _b.type) === "rangeRow" ||
        ((_c = sheet === null || sheet === void 0 ? void 0 : sheet.frozen) === null || _c === void 0 ? void 0 : _c.type) === "rangeBoth" ||
        ((_d = sheet === null || sheet === void 0 ? void 0 : sheet.frozen) === null || _d === void 0 ? void 0 : _d.type) === "both") {
        return (rowLocationByIndex(((_f = (_e = sheet === null || sheet === void 0 ? void 0 : sheet.frozen) === null || _e === void 0 ? void 0 : _e.range) === null || _f === void 0 ? void 0 : _f.row_focus) || 0, ctx.visibledatarow)[1] + ctx.scrollTop);
    }
    return ctx.scrollTop;
}
export function getFrozenHandleLeft(ctx) {
    var _a, _b, _c, _d, _e, _f;
    var idx = getSheetIndex(ctx, ctx.currentSheetId);
    if (idx == null)
        return ctx.scrollLeft;
    var sheet = ctx.luckysheetfile[idx];
    if (((_a = sheet === null || sheet === void 0 ? void 0 : sheet.frozen) === null || _a === void 0 ? void 0 : _a.type) === "column" ||
        ((_b = sheet === null || sheet === void 0 ? void 0 : sheet.frozen) === null || _b === void 0 ? void 0 : _b.type) === "rangeColumn" ||
        ((_c = sheet === null || sheet === void 0 ? void 0 : sheet.frozen) === null || _c === void 0 ? void 0 : _c.type) === "rangeBoth" ||
        ((_d = sheet === null || sheet === void 0 ? void 0 : sheet.frozen) === null || _d === void 0 ? void 0 : _d.type) === "both") {
        return (colLocationByIndex(((_f = (_e = sheet === null || sheet === void 0 ? void 0 : sheet.frozen) === null || _e === void 0 ? void 0 : _e.range) === null || _f === void 0 ? void 0 : _f.column_focus) || 0, ctx.visibledatacolumn)[1] -
            2 +
            ctx.scrollLeft);
    }
    return ctx.scrollLeft;
}
