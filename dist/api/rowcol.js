import _ from "lodash";
import { deleteRowCol, insertRowCol } from "../modules";
import { getSheet } from "./common";
import { INVALID_PARAMS } from "./errors";
import { getSheetIndex } from "../utils";
export function freeze(ctx, type, range, options) {
    if (options === void 0) { options = {}; }
    var sheet = getSheet(ctx, options);
    var typeMap = {
        row: "rangeRow",
        column: "rangeColumn",
        both: "rangeBoth",
    };
    var innerType = typeMap[type];
    sheet.frozen = {
        // @ts-ignore
        type: innerType,
        range: {
            column_focus: range.column,
            row_focus: range.row,
        },
    };
}
export function insertRowOrColumn(ctx, type, index, count, direction, options) {
    if (options === void 0) { options = {}; }
    if (!["row", "column"].includes(type) ||
        !_.isNumber(index) ||
        !_.isNumber(count) ||
        !["lefttop", "rightbottom"].includes(direction)) {
        throw INVALID_PARAMS;
    }
    var sheet = getSheet(ctx, options);
    try {
        insertRowCol(ctx, {
            type: type,
            index: index,
            count: count,
            direction: direction,
            id: sheet.id,
        });
    }
    catch (e) {
        console.error(e);
    }
}
export function deleteRowOrColumn(ctx, type, start, end, options) {
    if (options === void 0) { options = {}; }
    if (!["row", "column"].includes(type) ||
        !_.isNumber(start) ||
        !_.isNumber(end)) {
        throw INVALID_PARAMS;
    }
    var sheet = getSheet(ctx, options);
    deleteRowCol(ctx, { type: type, start: start, end: end, id: sheet.id });
}
export function hideRowOrColumn(ctx, rowColInfo, type) {
    var _a, _b;
    if (!["row", "column"].includes(type)) {
        throw INVALID_PARAMS;
    }
    if (!ctx || !ctx.config)
        return;
    var index = getSheetIndex(ctx, ctx.currentSheetId);
    if (type === "row") {
        /* TODO: 工作表保护判断
        if (
          !checkProtectionAuthorityNormal(Store.currentSheetIndex, "formatRows")
        ) {
          return ;
        } */
        var rowhidden_1 = (_a = ctx.config.rowhidden) !== null && _a !== void 0 ? _a : {};
        rowColInfo.forEach(function (r) {
            rowhidden_1[r] = 0;
        });
        /* 保存撤销,luck中保存撤销用以下方式实现，而在本项目中不需要另外处理
          if(Store.clearjfundo){
            let redo = {};
            redo["type"] = "showHidRows";
            redo["sheetIndex"] = Store.currentSheetIndex;
            redo["config"] = $.extend(true, {}, Store.config);
            redo["curconfig"] = cfg;
    
            Store.jfundo.length  = 0;
            Store.jfredo.push(redo);
        } */
        ctx.config.rowhidden = rowhidden_1;
        // const rowLen = ctx.luckysheetfile[index].data!.length;
        /**
         * 计算要隐藏的行是否是最后一列
         * 符合最后一列的条件：要隐藏的index===表格的长度-1 或者
         * 记录隐藏数组里面的数-1===要隐藏的index
         */
    }
    else if (type === "column") {
        // 隐藏列
        var colhidden_1 = (_b = ctx.config.colhidden) !== null && _b !== void 0 ? _b : {};
        rowColInfo.forEach(function (r) {
            colhidden_1[r] = 0;
        });
        ctx.config.colhidden = colhidden_1;
        // const columnLen = ctx.luckysheetfile[index].data![0].length;
    }
    ctx.luckysheetfile[index].config = ctx.config;
}
export function showRowOrColumn(ctx, rowColInfo, type) {
    var _a, _b;
    if (!["row", "column"].includes(type)) {
        throw INVALID_PARAMS;
    }
    if (!ctx || !ctx.config)
        return;
    var index = getSheetIndex(ctx, ctx.currentSheetId);
    if (type === "row") {
        /* TODO: 工作表保护判断
        if (
          !checkProtectionAuthorityNormal(Store.currentSheetIndex, "formatRows")
        ) {
          return ;
        } */
        var rowhidden_2 = (_a = ctx.config.rowhidden) !== null && _a !== void 0 ? _a : {};
        rowColInfo.forEach(function (r) {
            delete rowhidden_2[r];
        });
        /* 保存撤销,luck中保存撤销用以下方式实现，而在本项目中不需要另外处理
          if(Store.clearjfundo){
            let redo = {};
            redo["type"] = "showHidRows";
            redo["sheetIndex"] = Store.currentSheetIndex;
            redo["config"] = $.extend(true, {}, Store.config);
            redo["curconfig"] = cfg;
    
            Store.jfundo.length  = 0;
            Store.jfredo.push(redo);
        } */
        ctx.config.rowhidden = rowhidden_2;
        // const rowLen = ctx.luckysheetfile[index].data!.length;
        /**
         * 计算要隐藏的行是否是最后一列
         * 符合最后一列的条件：要隐藏的index===表格的长度-1 或者
         * 记录隐藏数组里面的数-1===要隐藏的index
         */
    }
    else if (type === "column") {
        // 隐藏列
        var colhidden_2 = (_b = ctx.config.colhidden) !== null && _b !== void 0 ? _b : {};
        rowColInfo.forEach(function (r) {
            delete colhidden_2[r];
        });
        ctx.config.colhidden = colhidden_2;
        // const columnLen = ctx.luckysheetfile[index].data![0].length;
    }
    ctx.luckysheetfile[index].config = ctx.config;
}
export function setRowHeight(ctx, rowInfo, options, custom) {
    if (options === void 0) { options = {}; }
    if (custom === void 0) { custom = false; }
    if (!_.isPlainObject(rowInfo)) {
        throw INVALID_PARAMS;
    }
    var sheet = getSheet(ctx, options);
    var cfg = sheet.config || {};
    if (cfg.rowlen == null) {
        cfg.rowlen = {};
    }
    _.forEach(rowInfo, function (len, r) {
        var _a;
        if (Number(r) >= 0) {
            if (Number(len) >= 0) {
                cfg.rowlen[Number(r)] = Number(len);
                if (custom && _.isUndefined(cfg.customHeight)) {
                    cfg.customHeight = (_a = {}, _a[r] = 1, _a);
                }
                else if (custom) {
                    cfg.customHeight[r] = 1;
                }
            }
        }
    });
    sheet.config = cfg;
    if (ctx.currentSheetId === sheet.id) {
        ctx.config = cfg;
    }
    // server.saveParam("cg", file.id, cfg.rowlen, { k: "rowlen" });
}
export function setColumnWidth(ctx, columnInfo, options, custom) {
    if (options === void 0) { options = {}; }
    if (custom === void 0) { custom = false; }
    if (!_.isPlainObject(columnInfo)) {
        throw INVALID_PARAMS;
    }
    var sheet = getSheet(ctx, options);
    var cfg = sheet.config || {};
    if (cfg.columnlen == null) {
        cfg.columnlen = {};
    }
    _.forEach(columnInfo, function (len, c) {
        var _a;
        if (Number(c) >= 0) {
            if (Number(len) >= 0) {
                cfg.columnlen[Number(c)] = Number(len);
                if (custom && _.isUndefined(cfg.customWidth)) {
                    cfg.customWidth = (_a = {}, _a[c] = 1, _a);
                }
                else if (custom) {
                    cfg.customWidth[c] = 1;
                }
            }
        }
    });
    sheet.config = cfg;
    if (ctx.currentSheetId === sheet.id) {
        ctx.config = cfg;
    }
    // server.saveParam("cg", file.id, cfg.columnlen, { k: "columnlen" });
}
export function getRowHeight(ctx, rows, options) {
    if (options === void 0) { options = {}; }
    if (!_.isArray(rows) || rows.length === 0) {
        throw INVALID_PARAMS;
    }
    var sheet = getSheet(ctx, options);
    var cfg = sheet.config || {};
    var rowlen = cfg.rowlen || {};
    var rowlenObj = {};
    rows.forEach(function (item) {
        if (Number(item) >= 0) {
            var size = rowlen[Number(item)] || ctx.defaultrowlen;
            rowlenObj[Number(item)] = size;
        }
    });
    return rowlenObj;
}
export function getColumnWidth(ctx, columns, options) {
    if (options === void 0) { options = {}; }
    if (!_.isArray(columns) || columns.length === 0) {
        throw INVALID_PARAMS;
    }
    var sheet = getSheet(ctx, options);
    var cfg = sheet.config || {};
    var columnlen = cfg.columnlen || {};
    var columnlenObj = {};
    columns.forEach(function (item) {
        if (Number(item) >= 0) {
            var size = columnlen[Number(item)] || ctx.defaultcollen;
            columnlenObj[Number(item)] = size;
        }
    });
    return columnlenObj;
}
