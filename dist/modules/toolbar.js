import _ from "lodash";
import { mergeCells } from "./merge";
import { getFlowdata } from "../context";
import { getSheetIndex, isAllowEdit } from "../utils";
import { getRangetxt, isAllSelectedCellsInStatus, normalizedAttr, setCellValue, } from "./cell";
import { colors } from "./color";
import { genarate, is_date, update } from "./format";
import { execfunction, execFunctionGroup, israngeseleciton, rangeSetValue, setCaretPosition, createFormulaRangeSelect, } from "./formula";
import { inlineStyleAffectAttribute, updateInlineStringFormat, updateInlineStringFormatOutside, } from "./inline-string";
import { colLocationByIndex, rowLocationByIndex } from "./location";
import { normalizeSelection, selectionCopyShow, selectIsOverlap, } from "./selection";
import { sortSelection } from "./sort";
import { hasPartMC, isdatatypemulti, isRealNull, isRealNum, } from "./validation";
import { showLinkCard } from "./hyperlink";
import { cfSplitRange } from "./conditionalFormat";
import { getCellTextInfo } from "./text";
export function updateFormatCell(ctx, d, attr, foucsStatus, row_st, row_ed, col_st, col_ed, canvas) {
    var _a, _b;
    var _c;
    if (_.isNil(d) || _.isNil(attr)) {
        return;
    }
    if (attr === "ct") {
        for (var r = row_st; r <= row_ed; r += 1) {
            if (!_.isNil(ctx.config.rowhidden) && !_.isNil(ctx.config.rowhidden[r])) {
                continue;
            }
            for (var c = col_st; c <= col_ed; c += 1) {
                var cell = d[r][c];
                var value = void 0;
                if (_.isPlainObject(cell)) {
                    value = cell === null || cell === void 0 ? void 0 : cell.v;
                }
                else {
                    value = cell;
                }
                if (foucsStatus !== "@" && isRealNum(value)) {
                    value = Number(value);
                }
                var mask = update(foucsStatus, value);
                var type = "n";
                if (is_date(foucsStatus) ||
                    foucsStatus === 14 ||
                    foucsStatus === 15 ||
                    foucsStatus === 16 ||
                    foucsStatus === 17 ||
                    foucsStatus === 18 ||
                    foucsStatus === 19 ||
                    foucsStatus === 20 ||
                    foucsStatus === 21 ||
                    foucsStatus === 22 ||
                    foucsStatus === 45 ||
                    foucsStatus === 46 ||
                    foucsStatus === 47) {
                    type = "d";
                }
                else if (foucsStatus === "@" || foucsStatus === 49) {
                    type = "s";
                }
                else if (foucsStatus === "General" || foucsStatus === 0) {
                    // type = "g";
                    type = isRealNum(value) ? "n" : "g";
                }
                if (cell && _.isPlainObject(cell)) {
                    cell.m = "".concat(mask);
                    if (_.isNil(cell.ct)) {
                        cell.ct = {};
                    }
                    cell.ct.fa = foucsStatus;
                    cell.ct.t = type;
                }
                else {
                    d[r][c] = {
                        ct: { fa: foucsStatus, t: type },
                        v: value,
                        m: mask,
                    };
                }
            }
        }
    }
    else {
        if (attr === "ht") {
            if (foucsStatus === "left") {
                foucsStatus = "1";
            }
            else if (foucsStatus === "center") {
                foucsStatus = "0";
            }
            else if (foucsStatus === "right") {
                foucsStatus = "2";
            }
        }
        else if (attr === "vt") {
            if (foucsStatus === "top") {
                foucsStatus = "1";
            }
            else if (foucsStatus === "middle") {
                foucsStatus = "0";
            }
            else if (foucsStatus === "bottom") {
                foucsStatus = "2";
            }
        }
        else if (attr === "tb") {
            if (foucsStatus === "overflow") {
                foucsStatus = "1";
            }
            else if (foucsStatus === "clip") {
                foucsStatus = "0";
            }
            else if (foucsStatus === "wrap") {
                foucsStatus = "2";
            }
        }
        else if (attr === "tr") {
            if (foucsStatus === "none") {
                foucsStatus = "0";
            }
            else if (foucsStatus === "angleup") {
                foucsStatus = "1";
            }
            else if (foucsStatus === "angledown") {
                foucsStatus = "2";
            }
            else if (foucsStatus === "vertical") {
                foucsStatus = "3";
            }
            else if (foucsStatus === "rotation-up") {
                foucsStatus = "4";
            }
            else if (foucsStatus === "rotation-down") {
                foucsStatus = "5";
            }
        }
        var sheetIndex = getSheetIndex(ctx, ctx.currentSheetId);
        if (sheetIndex == null) {
            return;
        }
        for (var r = row_st; r <= row_ed; r += 1) {
            if (!_.isNil(ctx.config.rowhidden) && !_.isNil(ctx.config.rowhidden[r])) {
                continue;
            }
            for (var c = col_st; c <= col_ed; c += 1) {
                var value = d[r][c];
                if (value && _.isPlainObject(value)) {
                    // if(attr in inlineStyleAffectAttribute && isInlineStringCell(value)){
                    updateInlineStringFormatOutside(value, attr, foucsStatus);
                    // }
                    // else{
                    // @ts-ignore
                    value[attr] = foucsStatus;
                    // }
                    (_c = ctx.luckysheetfile[sheetIndex]).config || (_c.config = {});
                    var cfg = ctx.luckysheetfile[sheetIndex].config;
                    var cellWidth = ((_a = cfg.columnlen) === null || _a === void 0 ? void 0 : _a[c]) ||
                        ctx.luckysheetfile[sheetIndex].defaultColWidth;
                    if (attr === "fs" && canvas) {
                        var textInfo = getCellTextInfo(d[r][c], canvas, ctx, {
                            r: r,
                            c: c,
                            cellWidth: cellWidth,
                        });
                        if (!textInfo)
                            continue;
                        var rowHeight = _.round(textInfo.textHeightAll);
                        var currentRowHeight = ((_b = cfg.rowlen) === null || _b === void 0 ? void 0 : _b[r]) ||
                            ctx.luckysheetfile[sheetIndex].defaultRowHeight ||
                            19;
                        if (!_.isUndefined(rowHeight) &&
                            rowHeight > currentRowHeight &&
                            (!cfg.customHeight || cfg.customHeight[r] !== 1)) {
                            if (_.isUndefined(cfg.rowlen))
                                cfg.rowlen = {};
                            _.set(cfg, "rowlen.".concat(r), rowHeight);
                        }
                    }
                }
                else {
                    // @ts-ignore
                    d[r][c] = { v: value };
                    // @ts-ignore
                    d[r][c][attr] = foucsStatus;
                }
                // if(attr === "tr" && !_.isNil(d[r][c].tb)){
                //     d[r][c].tb = "0";
                // }
            }
        }
    }
}
export function updateFormat(ctx, $input, d, attr, foucsStatus, canvas) {
    //   if (!checkProtectionFormatCells(ctx.currentSheetId)) {
    //     return;
    //   }
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit)
        return;
    if (attr in inlineStyleAffectAttribute) {
        if (ctx.luckysheetCellUpdate.length > 0) {
            var value = $input.innerText;
            if (value.substring(0, 1) !== "=") {
                var cell = d[ctx.luckysheetCellUpdate[0]][ctx.luckysheetCellUpdate[1]];
                if (cell) {
                    updateInlineStringFormat(ctx, cell, attr, foucsStatus, $input);
                }
                return;
            }
        }
    }
    var cfg = _.cloneDeep(ctx.config);
    if (_.isNil(cfg.rowlen)) {
        cfg.rowlen = {};
    }
    _.forEach(ctx.luckysheet_select_save, function (selection) {
        var _a = selection.row, row_st = _a[0], row_ed = _a[1];
        var _b = selection.column, col_st = _b[0], col_ed = _b[1];
        updateFormatCell(ctx, d, attr, foucsStatus, row_st, row_ed, col_st, col_ed, canvas);
        // if (attr === "tb" || attr === "tr" || attr === "fs") {
        //   cfg = rowlenByRange(ctx, d, row_st, row_ed, cfg);
        // }
    });
    //   let allParam = {};
    //   if (attr === "tb" || attr === "tr" || attr === "fs") {
    //     allParam = {
    //       cfg,
    //       RowlChange: true,
    //     };
    //   }
    //   jfrefreshgrid(d, ctx.luckysheet_select_save, allParam, false);
}
function toggleAttr(ctx, cellInput, attr) {
    var flowdata = getFlowdata(ctx);
    if (!flowdata)
        return;
    var flag = isAllSelectedCellsInStatus(ctx, attr, 1);
    var foucsStatus = flag ? 0 : 1;
    updateFormat(ctx, cellInput, flowdata, attr, foucsStatus);
}
function setAttr(ctx, cellInput, attr, value, canvas) {
    var flowdata = getFlowdata(ctx);
    if (!flowdata)
        return;
    updateFormat(ctx, cellInput, flowdata, attr, value, canvas);
}
// @ts-ignore
function checkNoNullValue(cell) {
    var v = cell;
    if (_.isPlainObject(v)) {
        v = v.v;
    }
    if (!isRealNull(v) &&
        isdatatypemulti(v).num &&
        (cell.ct == null ||
            cell.ct.t == null ||
            cell.ct.t === "n" ||
            cell.ct.t === "g")) {
        return true;
    }
    return false;
}
// @ts-ignore
function checkNoNullValueAll(cell) {
    var v = cell;
    if (_.isPlainObject(v)) {
        v = v.v;
    }
    if (!isRealNull(v)) {
        return true;
    }
    return false;
}
function getNoNullValue(d, st_x, ed, type) {
    // let hasValueSum = 0;
    var hasValueStart = null;
    var nullNum = 0;
    var nullTime = 0;
    for (var r = ed - 1; r >= 0; r -= 1) {
        var cell = void 0;
        if (type === "c") {
            cell = d[st_x][r];
        }
        else {
            cell = d[r][st_x];
        }
        if (checkNoNullValue(cell)) {
            // hasValueSum += 1;
            hasValueStart = r;
        }
        else if (cell == null || cell.v == null || cell.v === "") {
            nullNum += 1;
            if (nullNum >= 40) {
                if (nullTime <= 0) {
                    nullTime = 1;
                }
                else {
                    break;
                }
            }
        }
        else {
            break;
        }
    }
    return hasValueStart;
}
function activeFormulaInput(cellInput, fxInput, ctx, row_index, col_index, rowh, columnh, formula, cache, isnull) {
    if (isnull == null) {
        isnull = false;
    }
    ctx.luckysheetCellUpdate = [row_index, col_index];
    cache.doNotUpdateCell = true;
    if (isnull) {
        var formulaTxt_1 = "<span dir=\"auto\" class=\"luckysheet-formula-text-color\">=</span><span dir=\"auto\" class=\"luckysheet-formula-text-color\">".concat(formula.toUpperCase(), "</span><span dir=\"auto\" class=\"luckysheet-formula-text-color\">(</span><span dir=\"auto\" class=\"luckysheet-formula-text-color\">)</span>");
        cellInput.innerHTML = formulaTxt_1;
        var spanList = cellInput.querySelectorAll("span");
        setCaretPosition(ctx, spanList[spanList.length - 2], 0, 1);
        return;
    }
    var row_pre = rowLocationByIndex(rowh[0], ctx.visibledatarow)[0];
    var row = rowLocationByIndex(rowh[1], ctx.visibledatarow)[1];
    var col_pre = colLocationByIndex(columnh[0], ctx.visibledatacolumn)[0];
    var col = colLocationByIndex(columnh[1], ctx.visibledatacolumn)[1];
    var formulaTxt = "<span dir=\"auto\" class=\"luckysheet-formula-text-color\">=</span><span dir=\"auto\" class=\"luckysheet-formula-text-color\">".concat(formula.toUpperCase(), "</span><span dir=\"auto\" class=\"luckysheet-formula-text-color\">(</span><span class=\"fortune-formula-functionrange-cell\" rangeindex=\"0\" dir=\"auto\" style=\"color:").concat(colors[0], ";\">").concat(getRangetxt(ctx, ctx.currentSheetId, { row: rowh, column: columnh }, ctx.currentSheetId), "</span><span dir=\"auto\" class=\"luckysheet-formula-text-color\">)</span>");
    cellInput.innerHTML = formulaTxt;
    israngeseleciton(ctx);
    ctx.formulaCache.rangestart = true;
    ctx.formulaCache.rangedrag_column_start = false;
    ctx.formulaCache.rangedrag_row_start = false;
    ctx.formulaCache.rangechangeindex = 0;
    rangeSetValue(ctx, cellInput, { row: rowh, column: columnh }, fxInput);
    ctx.formulaCache.func_selectedrange = {
        left: col_pre,
        width: col - col_pre - 1,
        top: row_pre,
        height: row - row_pre - 1,
        left_move: col_pre,
        width_move: col - col_pre - 1,
        top_move: row_pre,
        height_move: row - row_pre - 1,
        row: [row_index, row_index],
        column: [col_index, col_index],
    };
    createFormulaRangeSelect(ctx, {
        rangeIndex: ctx.formulaCache.rangeIndex || 0,
        left: col_pre,
        width: col - col_pre - 1,
        top: row_pre,
        height: row - row_pre - 1,
    });
    // $("#fortune-formula-functionrange-select")
    //   .css({
    //     left: col_pre,
    //     width: col - col_pre - 1,
    //     top: row_pre,
    //     height: row - row_pre - 1,
    //   })
    //   .show(); TODO！！！
    // $("#luckysheet-formula-help-c").hide();
}
function backFormulaInput(d, r, c, rowh, columnh, formula, ctx) {
    var _a;
    var f = "=".concat(formula.toUpperCase(), "(").concat(getRangetxt(ctx, ctx.currentSheetId, { row: rowh, column: columnh }, ctx.currentSheetId), ")");
    var v = execfunction(ctx, f, r, c);
    var value = { v: v[1], f: v[2] };
    setCellValue(ctx, r, c, d, value);
    (_a = ctx.formulaCache).execFunctionExist || (_a.execFunctionExist = []);
    ctx.formulaCache.execFunctionExist.push({
        r: r,
        c: c,
        i: ctx.currentSheetId,
    });
    // server.historyParam(d, ctx.currentSheetId, {
    //   row: [r, r],
    //   column: [c, c],
    // }); 目前没有server
}
function singleFormulaInput(cellInput, fxInput, ctx, d, _index, fix, st_m, ed_m, formula, type, cache, noNum, noNull) {
    if (type == null) {
        type = "r";
    }
    if (noNum == null) {
        noNum = true;
    }
    if (noNull == null) {
        noNull = true;
    }
    var isNull = true;
    var isNum = false;
    for (var c = st_m; c <= ed_m; c += 1) {
        var cell = null;
        if (type === "c") {
            cell = d[c][fix];
        }
        else {
            cell = d[fix][c];
        }
        if (checkNoNullValue(cell)) {
            isNull = false;
            isNum = true;
        }
        else if (checkNoNullValueAll(cell)) {
            isNull = false;
        }
    }
    if (isNull && noNull) {
        var st_r_r = getNoNullValue(d, _index, fix, type);
        if (st_r_r == null) {
            if (type === "c") {
                activeFormulaInput(cellInput, fxInput, ctx, _index, fix, null, null, formula, cache, true);
            }
            else {
                activeFormulaInput(cellInput, fxInput, ctx, fix, _index, null, null, formula, cache, true);
            }
        }
        else {
            if (_index === st_m) {
                for (var c = st_m; c <= ed_m; c += 1) {
                    st_r_r = getNoNullValue(d, c, fix, type);
                    if (st_r_r == null) {
                        break;
                    }
                    if (type === "c") {
                        backFormulaInput(d, c, fix, [c, c], [st_r_r, fix - 1], formula, ctx);
                    }
                    else {
                        backFormulaInput(d, fix, c, [st_r_r, fix - 1], [c, c], formula, ctx);
                    }
                }
            }
            else {
                for (var c = ed_m; c >= st_m; c -= 1) {
                    st_r_r = getNoNullValue(d, c, fix, type);
                    if (st_r_r == null) {
                        break;
                    }
                    if (type === "c") {
                        backFormulaInput(d, c, fix, [c, c], [st_r_r, fix - 1], formula, ctx);
                    }
                    else {
                        backFormulaInput(d, fix, c, [st_r_r, fix - 1], [c, c], formula, ctx);
                    }
                }
            }
        }
        return false;
    }
    if (isNum && noNum) {
        var cell = null;
        if (type === "c") {
            cell = d[ed_m + 1][fix];
        }
        else {
            cell = d[fix][ed_m + 1];
        }
        /* 备注：在搜寻的时候排除自己以解决单元格函数引用自己的问题 */
        if (cell != null && cell.v != null && cell.v.toString().length > 0) {
            var c = ed_m + 1;
            if (type === "c") {
                cell = d[ed_m + 1][fix];
            }
            else {
                cell = d[fix][ed_m + 1];
            }
            while (cell != null && cell.v != null && cell.v.toString().length > 0) {
                c += 1;
                var len = null;
                if (type === "c") {
                    len = d.length;
                }
                else {
                    len = d[0].length;
                }
                if (c >= len) {
                    return false;
                }
                if (type === "c") {
                    cell = d[c][fix];
                }
                else {
                    cell = d[fix][c];
                }
            }
            if (type === "c") {
                backFormulaInput(d, c, fix, [st_m, ed_m], [fix, fix], formula, ctx);
            }
            else {
                backFormulaInput(d, fix, c, [fix, fix], [st_m, ed_m], formula, ctx);
            }
        }
        else {
            if (type === "c") {
                backFormulaInput(d, ed_m + 1, fix, [st_m, ed_m], [fix, fix], formula, ctx);
            }
            else {
                backFormulaInput(d, fix, ed_m + 1, [fix, fix], [st_m, ed_m], formula, ctx);
            }
        }
        return false;
    }
    return true;
}
export function autoSelectionFormula(ctx, cellInput, fxInput, formula, cache) {
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit)
        return;
    var flowdata = getFlowdata(ctx);
    if (flowdata == null)
        return;
    // const nullfindnum = 40;
    var isfalse = true;
    ctx.formulaCache.execFunctionExist = [];
    function execFormulaInput_c(d, st_r, ed_r, st_c, ed_c, _formula) {
        var st_c_c = getNoNullValue(d, st_r, ed_c, "c");
        if (st_c_c == null) {
            activeFormulaInput(cellInput, fxInput, ctx, st_r, st_c, null, null, _formula, cache, true);
        }
        else {
            activeFormulaInput(cellInput, fxInput, ctx, st_r, st_c, [st_r, ed_r], [st_c_c, ed_c - 1], _formula, cache);
        }
    }
    function execFormulaInput(d, st_r, ed_r, st_c, ed_c, _formula) {
        var st_r_c = getNoNullValue(d, st_c, ed_r, "r");
        if (st_r_c == null) {
            execFormulaInput_c(d, st_r, ed_r, st_c, ed_c, _formula);
        }
        else {
            activeFormulaInput(cellInput, fxInput, ctx, st_r, st_c, [st_r_c, ed_r - 1], [st_c, ed_c], _formula, cache);
        }
    }
    if (!ctx.luckysheet_select_save)
        return;
    _.forEach(ctx.luckysheet_select_save, function (selection) {
        var _a = selection.row, st_r = _a[0], ed_r = _a[1];
        var _b = selection.column, st_c = _b[0], ed_c = _b[1];
        var row_index = selection.row_focus;
        var col_index = selection.column_focus;
        if (st_r === ed_r && st_c === ed_c) {
            if (ed_r - 1 < 0 && ed_c - 1 < 0) {
                activeFormulaInput(cellInput, fxInput, ctx, st_r, st_c, null, null, formula, cache, true);
                return;
            }
            if (ed_r - 1 >= 0 && checkNoNullValue(flowdata[ed_r - 1][st_c])) {
                execFormulaInput(flowdata, st_r, ed_r, st_c, ed_c, formula);
            }
            else if (ed_c - 1 >= 0 && checkNoNullValue(flowdata[st_r][ed_c - 1])) {
                execFormulaInput_c(flowdata, st_r, ed_r, st_c, ed_c, formula);
            }
            else {
                execFormulaInput(flowdata, st_r, ed_r, st_c, ed_c, formula);
            }
        }
        else if (st_r === ed_r) {
            isfalse = singleFormulaInput(cellInput, fxInput, ctx, flowdata, col_index, st_r, st_c, ed_c, formula, "r", cache);
        }
        else if (st_c === ed_c) {
            isfalse = singleFormulaInput(cellInput, fxInput, ctx, flowdata, row_index, st_c, st_r, ed_r, formula, "c", cache);
        }
        else {
            var r_false = true;
            for (var r = st_r; r <= ed_r; r += 1) {
                r_false =
                    singleFormulaInput(cellInput, fxInput, ctx, flowdata, col_index, r, st_c, ed_c, formula, "r", cache, true, false) && r_false;
            }
            var c_false = true;
            for (var c = st_c; c <= ed_c; c += 1) {
                c_false =
                    singleFormulaInput(cellInput, fxInput, ctx, flowdata, row_index, c, st_r, ed_r, formula, "c", cache, true, false) && c_false;
            }
            isfalse = !!r_false && !!c_false;
        }
        isfalse = isfalse && isfalse;
    });
    if (!isfalse) {
        ctx.formulaCache.execFunctionExist.reverse();
        // @ts-ignore
        execFunctionGroup(ctx, null, null, null, null, flowdata);
        ctx.formulaCache.execFunctionGlobalData = null;
    }
}
export function cancelPaintModel(ctx) {
    var _a;
    // $("#luckysheet-sheettable_0").removeClass("luckysheetPaintCursor");
    if (ctx.luckysheet_copy_save === null)
        return;
    if (((_a = ctx.luckysheet_copy_save) === null || _a === void 0 ? void 0 : _a.dataSheetId) === ctx.currentSheetId) {
        ctx.luckysheet_selection_range = [];
        selectionCopyShow(ctx.luckysheet_selection_range, ctx);
    }
    else {
        if (!ctx.luckysheet_copy_save)
            return;
        var index = getSheetIndex(ctx, ctx.luckysheet_copy_save.dataSheetId);
        if (!index)
            return;
        // ctx.luckysheetfile[getSheetIndex(ctx.luckysheet_copy_save["dataSheetIndex"])].luckysheet_selection_range = [];
        ctx.luckysheetfile[index].luckysheet_selection_range = [];
    }
    ctx.luckysheet_copy_save = {
        dataSheetId: "",
        copyRange: [{ row: [0], column: [0] }],
        RowlChange: false,
        HasMC: false,
    };
    ctx.luckysheetPaintModelOn = false;
    // $("#luckysheetpopover").fadeOut(200,function(){
    //     $("#luckysheetpopover").remove();
}
export function handleCurrencyFormat(ctx, cellInput) {
    var flowdata = getFlowdata(ctx);
    if (!flowdata)
        return;
    var currency = ctx.currency || "¥";
    updateFormat(ctx, cellInput, flowdata, "ct", "".concat(currency, " #.00"));
}
export function handlePercentageFormat(ctx, cellInput) {
    var flowdata = getFlowdata(ctx);
    if (!flowdata)
        return;
    updateFormat(ctx, cellInput, flowdata, "ct", "0.00%");
}
export function handleNumberDecrease(ctx, cellInput) {
    var flowdata = getFlowdata(ctx);
    if (!flowdata || !ctx.luckysheet_select_save)
        return;
    var row_index = ctx.luckysheet_select_save[0].row_focus;
    var col_index = ctx.luckysheet_select_save[0].column_focus;
    if (row_index === undefined || col_index === undefined)
        return;
    var foucsStatus = normalizedAttr(flowdata, row_index, col_index, "ct");
    var cell = flowdata[row_index][col_index];
    if (foucsStatus == null || foucsStatus.t !== "n") {
        return;
    }
    if (foucsStatus.fa === "General") {
        if (!cell || !cell.v)
            return;
        var mask = genarate(cell.v);
        if (!mask || mask.length < 2)
            return;
        foucsStatus = mask[1];
    }
    // 万亿格式
    var reg = /^(w|W)((0?)|(0\.0+))$/;
    if (reg.test(foucsStatus.fa)) {
        if (foucsStatus.fa.indexOf(".") > -1) {
            if (foucsStatus.fa.substr(-2) === ".0") {
                updateFormat(ctx, cellInput, flowdata, "ct", foucsStatus.fa.split(".")[0]);
            }
            else {
                updateFormat(ctx, cellInput, flowdata, "ct", foucsStatus.fa.substr(0, foucsStatus.fa.length - 1));
            }
        }
        else {
            updateFormat(ctx, cellInput, flowdata, "ct", foucsStatus.fa);
        }
        return;
    }
    // Uncaught ReferenceError: Cannot access 'fa' before initialization
    var prefix = "";
    var main = "";
    var fa = [];
    if (foucsStatus.fa.indexOf(".") > -1) {
        fa = foucsStatus.fa.split(".");
        prefix = fa[0], main = fa[1];
    }
    else {
        return;
    }
    fa = main.split("");
    var tail = "";
    for (var i = fa.length - 1; i >= 0; i -= 1) {
        var c = fa[i];
        if (c !== "#" && c !== "0" && c !== "," && Number.isNaN(parseInt(c, 10))) {
            tail = c + tail;
        }
        else {
            break;
        }
    }
    var fmt = "";
    if (foucsStatus.fa.indexOf(".") > -1) {
        var suffix = main;
        if (tail.length > 0) {
            suffix = main.replace(tail, "");
        }
        var pos = suffix.replace(/#/g, "0");
        pos = pos.substr(0, pos.length - 1);
        if (pos === "") {
            fmt = prefix + tail;
        }
        else {
            fmt = "".concat(prefix, ".").concat(pos).concat(tail);
        }
    }
    updateFormat(ctx, cellInput, flowdata, "ct", fmt);
}
export function handleNumberIncrease(ctx, cellInput) {
    var flowdata = getFlowdata(ctx);
    if (!flowdata)
        return;
    if (!ctx.luckysheet_select_save)
        return;
    var row_index = ctx.luckysheet_select_save[0].row_focus;
    var col_index = ctx.luckysheet_select_save[0].column_focus;
    if (row_index === undefined || col_index === undefined)
        return;
    var foucsStatus = normalizedAttr(flowdata, row_index, col_index, "ct");
    var cell = flowdata[row_index][col_index];
    if (foucsStatus == null || foucsStatus.t !== "n") {
        return;
    }
    if (foucsStatus.fa === "General") {
        if (!cell || !cell.v)
            return;
        var mask = genarate(cell.v);
        if (!mask || mask.length < 2)
            return;
        foucsStatus = mask[1];
    }
    if (foucsStatus.fa === "General") {
        updateFormat(ctx, cellInput, flowdata, "ct", "#.0");
        return;
    }
    // 万亿格式
    var reg = /^(w|W)((0?)|(0\.0+))$/;
    if (reg.test(foucsStatus.fa)) {
        if (foucsStatus.fa.indexOf(".") > -1) {
            updateFormat(ctx, cellInput, flowdata, "ct", "".concat(foucsStatus.fa, "0"));
        }
        else {
            if (foucsStatus.fa.substr(-1) === "0") {
                updateFormat(ctx, cellInput, flowdata, "ct", "".concat(foucsStatus.fa, ".0"));
            }
            else {
                updateFormat(ctx, cellInput, flowdata, "ct", "".concat(foucsStatus.fa, "0.0"));
            }
        }
        return;
    }
    // Uncaught ReferenceError: Cannot access 'fa' before initialization
    var prefix = "";
    var main = "";
    var fa = [];
    if (foucsStatus.fa.indexOf(".") > -1) {
        fa = foucsStatus.fa.split(".");
        prefix = fa[0], main = fa[1];
    }
    else {
        main = foucsStatus.fa;
    }
    fa = main.split("");
    var tail = "";
    for (var i = fa.length - 1; i >= 0; i -= 1) {
        var c = fa[i];
        if (c !== "#" && c !== "0" && c !== "," && Number.isNaN(parseInt(c, 10))) {
            tail = c + tail;
        }
        else {
            break;
        }
    }
    var fmt = "";
    if (foucsStatus.fa.indexOf(".") > -1) {
        var suffix = main;
        if (tail.length > 0) {
            suffix = main.replace(tail, "");
        }
        var pos = suffix.replace(/#/g, "0");
        pos += "0";
        fmt = "".concat(prefix, ".").concat(pos).concat(tail);
    }
    else {
        if (tail.length > 0) {
            fmt = "".concat(main.replace(tail, ""), ".0").concat(tail);
        }
        else {
            fmt = "".concat(main, ".0").concat(tail);
        }
    }
    updateFormat(ctx, cellInput, flowdata, "ct", fmt);
}
export function handleBold(ctx, cellInput) {
    toggleAttr(ctx, cellInput, "bl");
}
export function handleItalic(ctx, cellInput) {
    toggleAttr(ctx, cellInput, "it");
}
export function handleStrikeThrough(ctx, cellInput) {
    toggleAttr(ctx, cellInput, "cl");
}
export function handleUnderline(ctx, cellInput) {
    toggleAttr(ctx, cellInput, "un");
}
export function handleHorizontalAlign(ctx, cellInput, value) {
    setAttr(ctx, cellInput, "ht", value);
}
export function handleVerticalAlign(ctx, cellInput, value) {
    setAttr(ctx, cellInput, "vt", value);
}
export function handleFormatPainter(ctx) {
    //   if (!checkIsAllowEdit()) {
    //     tooltip.info("", locale().pivotTable.errorNotAllowEdit);
    //     return
    // }
    // e.stopPropagation();
    // let _locale = locale();
    // let locale_paint = _locale.paint;
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit)
        return;
    if (ctx.luckysheet_select_save == null ||
        ctx.luckysheet_select_save.length === 0) {
        // if(isEditMode()){
        //     alert(locale_paint.tipSelectRange);
        // }
        // else{
        //     tooltip.info("",locale_paint.tipSelectRange);
        // }
        return;
    }
    if (ctx.luckysheet_select_save.length > 1) {
        // if(isEditMode()){
        //     alert(locale_paint.tipNotMulti);
        // }
        // else{
        //     tooltip.info("",locale_paint.tipNotMulti);
        // }
        return;
    }
    // *增加了对选区范围是否为部分合并单元格的校验，如果为部分合并单元格，就阻止格式刷的下一步
    // TODO 这里也可以改为：判断到是合并单元格的一部分后，格式刷执行黏贴格式后删除范围单元格的 mc 值
    var has_PartMC = false;
    var r1 = ctx.luckysheet_select_save[0].row[0];
    var r2 = ctx.luckysheet_select_save[0].row[1];
    var c1 = ctx.luckysheet_select_save[0].column[0];
    var c2 = ctx.luckysheet_select_save[0].column[1];
    has_PartMC = hasPartMC(ctx, ctx.config, r1, r2, c1, c2);
    if (has_PartMC) {
        // *提示后中止下一步
        // tooltip.info('无法对部分合并单元格执行此操作', '');
        return;
    }
    // tooltip.popover("<i class='fa fa-paint-brush'></i> "+locale_paint.start+"", "topCenter", true, null, locale_paint.end,function(){
    cancelPaintModel(ctx);
    // });
    // $("#luckysheet-sheettable_0").addClass("luckysheetPaintCursor");
    ctx.luckysheet_selection_range = [
        {
            row: ctx.luckysheet_select_save[0].row,
            column: ctx.luckysheet_select_save[0].column,
        },
    ];
    selectionCopyShow(ctx.luckysheet_selection_range, ctx);
    var RowlChange = false;
    var HasMC = false;
    for (var r = ctx.luckysheet_select_save[0].row[0]; r <= ctx.luckysheet_select_save[0].row[1]; r += 1) {
        if (ctx.config.rowhidden != null && ctx.config.rowhidden[r] != null) {
            continue;
        }
        if (ctx.config.rowlen != null && r in ctx.config.rowlen) {
            RowlChange = true;
        }
        for (var c = ctx.luckysheet_select_save[0].column[0]; c <= ctx.luckysheet_select_save[0].column[1]; c += 1) {
            var flowdata = getFlowdata(ctx);
            if (!flowdata)
                return;
            var cell = flowdata[r][c];
            if (cell != null && cell.mc != null && cell.mc.rs != null) {
                HasMC = true;
            }
        }
    }
    ctx.luckysheet_copy_save = {
        dataSheetId: ctx.currentSheetId,
        copyRange: [
            {
                row: ctx.luckysheet_select_save[0].row,
                column: ctx.luckysheet_select_save[0].column,
            },
        ],
        RowlChange: RowlChange,
        HasMC: HasMC,
    };
    ctx.luckysheetPaintModelOn = true;
    ctx.luckysheetPaintSingle = true;
}
// 2022-10-10 废弃了handleClearFormat中的foreach写法，改为可跳出的every写法，以防止选区多次覆盖
export function handleClearFormat(ctx) {
    var _a;
    if (ctx.allowEdit === false)
        return;
    var flowdata = getFlowdata(ctx);
    if (!flowdata)
        return;
    (_a = ctx.luckysheet_select_save) === null || _a === void 0 ? void 0 : _a.every(function (selection) {
        var _a = selection.row, rowSt = _a[0], rowEd = _a[1];
        var _b = selection.column, colSt = _b[0], colEd = _b[1];
        for (var r = rowSt; r <= rowEd; r += 1) {
            if (!_.isNil(ctx.config.rowhidden) && !_.isNil(ctx.config.rowhidden[r])) {
                continue;
            }
            for (var c = colSt; c <= colEd; c += 1) {
                var cell = flowdata[r][c];
                if (!cell)
                    continue;
                flowdata[r][c] = _.pick(cell, "v", "m", "mc", "f", "ct");
            }
        }
        // 清空表格样式时，清除边框样式
        var index = getSheetIndex(ctx, ctx.currentSheetId);
        if (index == null)
            return false;
        // 表格边框为空时，不对表格进行操作
        if (ctx.config.borderInfo == null)
            return false;
        var cfg = ctx.config || {};
        if (cfg.borderInfo && cfg.borderInfo.length > 0) {
            var source_borderInfo = [];
            for (var i = 0; i < cfg.borderInfo.length; i += 1) {
                var bd_rangeType = cfg.borderInfo[i].rangeType;
                if (bd_rangeType === "range" &&
                    cfg.borderInfo[i].borderType !== "border-slash") {
                    var bd_range = cfg.borderInfo[i].range;
                    var bd_emptyRange = [];
                    for (var j = 0; j < bd_range.length; j += 1) {
                        bd_emptyRange = bd_emptyRange.concat(cfSplitRange(bd_range[j], { row: [rowSt, rowEd], column: [colSt, colEd] }, { row: [rowSt, rowEd], column: [colSt, colEd] }, "restPart"));
                    }
                    cfg.borderInfo[i].range = bd_emptyRange;
                    source_borderInfo.push(cfg.borderInfo[i]);
                }
                else if (bd_rangeType === "cell") {
                    var bd_r = cfg.borderInfo[i].value.row_index;
                    var bd_c = cfg.borderInfo[i].value.col_index;
                    if (!(bd_r >= rowSt && bd_r <= rowEd && bd_c >= colSt && bd_c <= colEd)) {
                        source_borderInfo.push(cfg.borderInfo[i]);
                    }
                }
                else if (bd_rangeType === "range" &&
                    cfg.borderInfo[i].borderType === "border-slash" &&
                    !(cfg.borderInfo[i].range[0].row[0] >= rowSt &&
                        cfg.borderInfo[i].range[0].row[0] <= rowEd &&
                        cfg.borderInfo[i].range[0].column[0] >= colSt &&
                        cfg.borderInfo[i].range[0].column[0] <= colEd)) {
                    source_borderInfo.push(cfg.borderInfo[i]);
                }
            }
            ctx.luckysheetfile[index].config.borderInfo = source_borderInfo;
        }
        return true;
    });
}
export function handleTextColor(ctx, cellInput, color) {
    setAttr(ctx, cellInput, "fc", color);
}
export function handleTextBackground(ctx, cellInput, color) {
    setAttr(ctx, cellInput, "bg", color);
}
export function handleBorder(ctx, type, borderColor, borderStyle) {
    // *如果禁止前台编辑，则中止下一步操作
    // if (!checkIsAllowEdit()) {
    //   tooltip.info("", locale().pivotTable.errorNotAllowEdit);
    //   return;
    // }
    // if (!checkProtectionFormatCells(Store.currentSheetId)) {
    //   return;
    // }
    // const d = editor.deepCopyFlowData(Store.flowdata);
    // let type = $(this).attr("type");
    // let type = "border-all";
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit)
        return;
    if (type == null) {
        type = "border-all";
    }
    // const subcolormenuid = "luckysheet-icon-borderColor-menuButton";
    // let color = $(`#${subcolormenuid}`).find(".luckysheet-color-selected").val();
    // let style = $("#luckysheetborderSizepreview").attr("itemvalue");
    // let color = "#000000";
    var color = borderColor;
    var style = borderStyle;
    if (color == null || color === "") {
        color = "#000";
    }
    if (style == null || style === "") {
        style = "1";
    }
    var cfg = ctx.config;
    if (cfg.borderInfo == null) {
        cfg.borderInfo = [];
    }
    if (type !== "border-slash") {
        var borderInfo = {
            rangeType: "range",
            borderType: type,
            color: color,
            style: style,
            range: _.cloneDeep(ctx.luckysheet_select_save) || [],
        };
        cfg.borderInfo.push(borderInfo);
    }
    else {
        var rangeList_1 = [];
        _.forEach(ctx.luckysheet_select_save, function (selection) {
            for (var r = selection.row[0]; r <= selection.row[1]; r += 1) {
                for (var c = selection.column[0]; c <= selection.column[1]; c += 1) {
                    var range = "".concat(r, "_").concat(c);
                    if (_.includes(rangeList_1, range))
                        continue;
                    var borderInfo = {
                        rangeType: "range",
                        borderType: type,
                        color: color,
                        style: style,
                        range: normalizeSelection(ctx, [{ row: [r, r], column: [c, c] }]),
                    };
                    cfg.borderInfo.push(borderInfo);
                    rangeList_1.push(range);
                }
            }
        });
    }
    // server.saveParam("cg", ctx.currentSheetId, cfg.borderInfo, {
    //   k: "borderInfo",
    // });
    var index = getSheetIndex(ctx, ctx.currentSheetId);
    if (index == null)
        return;
    ctx.luckysheetfile[index].config = ctx.config;
    // setTimeout(function () {
    //   luckysheetrefreshgrid();
    // }, 1);
}
export function handleMerge(ctx, type) {
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit)
        return;
    // if (!checkProtectionNotEnable(ctx.currentSheetId)) {
    //   return;
    // }
    if (selectIsOverlap(ctx)) {
        //   if (isEditMode()) {
        //     alert("不能合并重叠区域");
        //   } else {
        //     tooltip.info("不能合并重叠区域", "");
        //   }
        return;
    }
    if (ctx.config.merge != null) {
        var has_PartMC = false;
        if (!ctx.luckysheet_select_save)
            return;
        for (var s = 0; s < ctx.luckysheet_select_save.length; s += 1) {
            var r1 = ctx.luckysheet_select_save[s].row[0];
            var r2 = ctx.luckysheet_select_save[s].row[1];
            var c1 = ctx.luckysheet_select_save[s].column[0];
            var c2 = ctx.luckysheet_select_save[s].column[1];
            has_PartMC = hasPartMC(ctx, ctx.config, r1, r2, c1, c2);
            if (has_PartMC) {
                break;
            }
        }
        if (has_PartMC) {
            // if (isEditMode()) {
            //   alert("无法对部分合并单元格执行此操作");
            // } else {
            //   tooltip.info("无法对部分合并单元格执行此操作", "");
            // }
            return;
        }
    }
    var flowdata = getFlowdata(ctx);
    if (!flowdata)
        return;
    if (!ctx.luckysheet_select_save)
        return;
    mergeCells(ctx, ctx.currentSheetId, ctx.luckysheet_select_save, type);
}
export function handleSort(ctx, isAsc) {
    sortSelection(ctx, isAsc);
}
export function handleFreeze(ctx, type) {
    var _a, _b;
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit)
        return;
    var file = ctx.luckysheetfile[getSheetIndex(ctx, ctx.currentSheetId)];
    if (!file)
        return;
    if (type === "freeze-cancel") {
        delete file.frozen;
        return;
    }
    var firstSelection = (_a = ctx.luckysheet_select_save) === null || _a === void 0 ? void 0 : _a[0];
    if (!firstSelection)
        return;
    var row_focus = firstSelection.row_focus, column_focus = firstSelection.column_focus;
    if (row_focus == null || column_focus == null)
        return;
    var m = (_b = ctx.config.merge) === null || _b === void 0 ? void 0 : _b["".concat(row_focus, "_").concat(column_focus)];
    if (m) {
        row_focus = m.r + m.rs - 1;
        column_focus = m.c + m.cs - 1;
    }
    file.frozen = { type: "both", range: { row_focus: row_focus, column_focus: column_focus } };
    if (type === "freeze-row") {
        file.frozen.type = "rangeRow";
    }
    else if (type === "freeze-col") {
        file.frozen.type = "rangeColumn";
    }
}
export function handleTextSize(ctx, cellInput, size, canvas) {
    setAttr(ctx, cellInput, "fs", size, canvas);
}
export function handleSum(ctx, cellInput, fxInput, cache) {
    autoSelectionFormula(ctx, cellInput, fxInput, "SUM", cache);
}
export function handleLink(ctx) {
    var _a;
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit)
        return;
    var selection = (_a = ctx.luckysheet_select_save) === null || _a === void 0 ? void 0 : _a[0];
    var flowdata = getFlowdata(ctx);
    if (flowdata != null && selection != null) {
        showLinkCard(ctx, selection.row[0], selection.column[0], true);
    }
}
var handlerMap = {
    "currency-format": handleCurrencyFormat,
    "percentage-format": handlePercentageFormat,
    "number-decrease": handleNumberDecrease,
    "number-increase": handleNumberIncrease,
    "sort-cell": function (ctx) { return handleSort(ctx, true); },
    "merge-all": function (ctx) { return handleMerge(ctx, "mergeAll"); },
    "border-all": function (ctx) { return handleBorder(ctx, "border-all"); },
    bold: handleBold,
    italic: handleItalic,
    "strike-through": handleStrikeThrough,
    underline: handleUnderline,
    "clear-format": handleClearFormat,
    "format-painter": handleFormatPainter,
    search: function (ctx) {
        ctx.showSearch = true;
    },
    link: handleLink,
};
var selectedMap = {
    bold: function (cell) { return (cell === null || cell === void 0 ? void 0 : cell.bl) === 1; },
    italic: function (cell) { return (cell === null || cell === void 0 ? void 0 : cell.it) === 1; },
    "strike-through": function (cell) { return (cell === null || cell === void 0 ? void 0 : cell.cl) === 1; },
    underline: function (cell) { return (cell === null || cell === void 0 ? void 0 : cell.un) === 1; },
};
export function toolbarItemClickHandler(name) {
    return handlerMap[name];
}
export function toolbarItemSelectedFunc(name) {
    return selectedMap[name];
}
