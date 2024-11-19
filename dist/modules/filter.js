import _ from "lodash";
import { locale } from "../locale";
import { getFlowdata } from "../context";
import { getSheetIndex, isAllowEdit, rgbToHex } from "../utils";
import { update } from "./format";
import { normalizeSelection } from "./selection";
import { isRealNull } from "./validation";
import { normalizedAttr } from "./cell";
import { sortDataRange } from "./sort";
import { checkCF, getComputeMap } from "./ConditionFormat";
// 筛选配置状态
export function labelFilterOptionState(ctx, optionstate, rowhidden, caljs, str, edr, cindex, stc, edc, saveData) {
    var param = {
        caljs: caljs,
        rowhidden: rowhidden,
        optionstate: optionstate,
        str: str,
        edr: edr,
        cindex: cindex,
        stc: stc,
        edc: edc,
    };
    if (optionstate) {
        ctx.filter[cindex - stc] = param;
        // 条件格式参数
        if (caljs != null) {
        }
    }
    else {
        delete ctx.filter[cindex - stc];
    }
    if (saveData) {
        var sheetIndex = getSheetIndex(ctx, ctx.currentSheetId);
        if (sheetIndex == null)
            return;
        var file = ctx.luckysheetfile[sheetIndex];
        if (file.filter == null) {
            file.filter = {};
        }
        if (optionstate) {
            file.filter[cindex - stc] = param;
        }
        else {
            delete file.filter[cindex - stc];
        }
        // server.saveParam("all", Store.currentSheetIndex, file.filter, {
        //   k: "filter",
        // });
    }
}
// 筛选排序
export function orderbydatafiler(ctx, str, stc, edr, edc, curr, asc) {
    var _a;
    var d = getFlowdata(ctx);
    if (d == null) {
        return null;
    }
    str += 1;
    var hasMc = false; // 排序选区是否有合并单元格
    var data = [];
    for (var r = str; r <= edr; r += 1) {
        var data_row = [];
        for (var c = stc; c <= edc; c += 1) {
            if (d[r][c] != null && ((_a = d[r][c]) === null || _a === void 0 ? void 0 : _a.mc) != null) {
                hasMc = true;
                break;
            }
            data_row.push(d[r][c]);
        }
        data.push(data_row);
    }
    if (hasMc) {
        var filter = locale(ctx).filter;
        // if (isEditMode()) {
        //   alert(locale_filter.mergeError);
        // } else {
        return filter.mergeError;
        // }
    }
    sortDataRange(ctx, d, data, curr - stc, asc, str, edr, stc, edc);
    return null;
}
// 创建筛选配置
export function createFilterOptions(ctx, luckysheet_filter_save, sheetId, filterObj, saveData) {
    var _a, _b, _c, _d;
    // $(`#luckysheet-filter-selected-sheet${ctx.currentSheetIndex}`).remove();
    // $(`#luckysheet-filter-options-sheet${ctx.currentSheetIndex}`).remove();
    // eslint-disable-next-line no-undef
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit)
        return;
    if (sheetId != null && sheetId !== ctx.currentSheetId)
        return;
    var sheetIndex = getSheetIndex(ctx, ctx.currentSheetId);
    if (sheetIndex == null)
        return;
    if (luckysheet_filter_save == null || _.size(luckysheet_filter_save) === 0) {
        delete ctx.filterOptions;
        return;
    }
    var r1 = luckysheet_filter_save.row[0];
    var r2 = luckysheet_filter_save.row[1];
    var c1 = luckysheet_filter_save.column[0];
    var c2 = luckysheet_filter_save.column[1];
    var row = (_a = ctx.visibledatarow[r2]) !== null && _a !== void 0 ? _a : 0;
    var row_pre = r1 - 1 === -1 ? 0 : (_b = ctx.visibledatarow[r1 - 1]) !== null && _b !== void 0 ? _b : 0;
    var col = (_c = ctx.visibledatacolumn[c2]) !== null && _c !== void 0 ? _c : 0;
    var col_pre = c1 - 1 === -1 ? 0 : (_d = ctx.visibledatacolumn[c1 - 1]) !== null && _d !== void 0 ? _d : 0;
    var options = {
        startRow: r1,
        endRow: r2,
        startCol: c1,
        endCol: c2,
        left: col_pre,
        top: row_pre,
        width: col - col_pre - 1,
        height: row - row_pre - 1,
        items: [],
    };
    for (var c = c1; c <= c2; c += 1) {
        // TODO: filterObj
        if (filterObj == null || (filterObj === null || filterObj === void 0 ? void 0 : filterObj[c - c1]) == null) {
        }
        else {
        }
        var left = 0;
        if (ctx.visibledatacolumn[c]) {
            left = ctx.visibledatacolumn[c] - 20;
        }
        options.items.push({
            col: c,
            left: left,
            top: row_pre,
        });
    }
    if (saveData) {
        var file = ctx.luckysheetfile[sheetIndex];
        file.filter_select = luckysheet_filter_save;
    }
    ctx.filterOptions = options;
}
export function clearFilter(ctx) {
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit)
        return;
    var sheetIndex = getSheetIndex(ctx, ctx.currentSheetId);
    var hiddenRows = _.reduce(ctx.filter, function (pre, curr) { return _.assign(pre, (curr === null || curr === void 0 ? void 0 : curr.rowhidden) || {}); }, {});
    ctx.config.rowhidden = _.omit(ctx.config.rowhidden, _.keys(hiddenRows));
    ctx.luckysheet_filter_save = undefined;
    ctx.filterOptions = undefined;
    ctx.filterContextMenu = undefined;
    ctx.filter = {};
    if (sheetIndex != null) {
        ctx.luckysheetfile[sheetIndex].filter = undefined;
        ctx.luckysheetfile[sheetIndex].filter_select = undefined;
        ctx.luckysheetfile[sheetIndex].config = _.assign({}, ctx.config);
    }
}
export function createFilter(ctx) {
    // if (!checkProtectionAuthorityNormal(ctx.currentSheetIndex, "filter")) {
    //   return;
    // }
    var _a, _b;
    if (_.size(ctx.luckysheet_select_save) > 1) {
        // const locale_splitText = locale().splitText;
        // if (isEditMode()) {
        //   alert(locale_splitText.tipNoMulti);
        // } else {
        //   tooltip.info(locale_splitText.tipNoMulti, "");
        // }
        return;
    }
    if (_.size(ctx.luckysheet_filter_save) > 0) {
        clearFilter(ctx);
        return;
    }
    var sheetIndex = getSheetIndex(ctx, ctx.currentSheetId);
    if (sheetIndex == null || ctx.luckysheetfile[sheetIndex].isPivotTable) {
        return;
    }
    // $(
    //   `#luckysheet-filter-selected-sheet${sheetId}, #luckysheet-filter-options-sheet${ctx.currentSheetId}`
    // ).remove();
    var last = (_a = ctx.luckysheet_select_save) === null || _a === void 0 ? void 0 : _a[0];
    var flowdata = getFlowdata(ctx);
    var filterSave;
    if (last == null || flowdata == null)
        return;
    if (last.row[0] === last.row[1] && last.column[0] === last.column[1]) {
        var st_c = void 0;
        var ed_c = void 0;
        var curR = last.row[1];
        for (var c = 0; c < flowdata[curR].length; c += 1) {
            var cell = flowdata[curR][c];
            if (cell != null && !isRealNull(cell.v)) {
                if (st_c == null) {
                    st_c = c;
                }
            }
            else if (st_c != null) {
                ed_c = c - 1;
                break;
            }
        }
        if (ed_c == null) {
            ed_c = flowdata[curR].length - 1;
        }
        filterSave = normalizeSelection(ctx, [
            { row: [curR, curR], column: [st_c || 0, ed_c] }, // st_c default 0 ?
        ]);
        ctx.luckysheet_select_save = filterSave;
        ctx.luckysheet_shiftpositon = _.assign({}, last);
        // luckysheetMoveEndCell("down", "range");
    }
    else if (last.row[1] - last.row[0] < 2) {
        ctx.luckysheet_shiftpositon = _.assign({}, last);
        // luckysheetMoveEndCell("down", "range");
    }
    ctx.luckysheet_filter_save = _.assign({}, (filterSave === null || filterSave === void 0 ? void 0 : filterSave[0]) || ((_b = ctx.luckysheet_select_save) === null || _b === void 0 ? void 0 : _b[0]));
    createFilterOptions(ctx, ctx.luckysheet_filter_save, undefined, {}, true);
    // server.saveParam("all", ctx.currentSheetIndex, ctx.luckysheet_filter_save, {
    //   k: "filter_select",
    // });
    // if (ctx.filterchage) {
    //   ctx.jfredo.push({
    //     type: "filtershow",
    //     data: [],
    //     curdata: [],
    //     sheetIndex: ctx.currentSheetIndex,
    //     filter_save: ctx.luckysheet_filter_save,
    //   });
    // }
}
function getFilterHiddenRows(ctx, col, startCol) {
    var _a;
    var otherHiddenRows = _.reduce(ctx.filter, function (pre, curr) {
        return _.assign(pre, ((curr === null || curr === void 0 ? void 0 : curr.cindex) !== col && (curr === null || curr === void 0 ? void 0 : curr.rowhidden)) || {});
    }, {});
    var hiddenRows = ((_a = ctx.filter[col - startCol]) === null || _a === void 0 ? void 0 : _a.rowhidden) || {};
    return { otherHiddenRows: otherHiddenRows, hiddenRows: hiddenRows };
}
export function getFilterColumnValues(ctx, col, startRow, endRow, startCol) {
    var _a = getFilterHiddenRows(ctx, col, startCol), otherHiddenRows = _a.otherHiddenRows, hiddenRows = _a.hiddenRows;
    var visibleRows = [];
    var flattenValues = [];
    // 日期值
    var dates = [];
    var datesUncheck = [];
    var dateRowMap = {};
    // 除日期以外的值
    var valuesMap = new Map();
    var valuesUncheck = [];
    var valueRowMap = {};
    var flowdata = getFlowdata(ctx);
    if (flowdata == null)
        return {
            dates: dates,
            datesUncheck: datesUncheck,
            dateRowMap: dateRowMap,
            values: [],
            valuesUncheck: valuesUncheck,
            valueRowMap: valueRowMap,
            visibleRows: visibleRows,
            flattenValues: flattenValues,
        };
    var cell;
    var filter = locale(ctx).filter;
    var _loop_1 = function (r) {
        if (r in otherHiddenRows) {
            return "continue";
        }
        visibleRows.push(r);
        cell = flowdata[r][col];
        if (cell != null &&
            !isRealNull(cell.v) &&
            cell.ct != null &&
            cell.ct.t === "d") {
            // 单元格是日期
            var dateStr = update("YYYY-MM-DD", cell.v);
            var y_1 = dateStr.split("-")[0];
            var m_1 = dateStr.split("-")[1];
            var d_1 = dateStr.split("-")[2];
            var yearValue = _.find(dates, function (v) { return v.value === y_1; });
            if (yearValue == null) {
                yearValue = {
                    key: y_1,
                    type: "year",
                    value: y_1,
                    text: y_1 + filter.filiterYearText,
                    children: [],
                    rows: [],
                    dateValues: [],
                };
                dates.push(yearValue);
                flattenValues.push(dateStr);
            }
            var monthValue = _.find(yearValue.children, function (v) { return v.value === m_1; });
            if (monthValue == null) {
                monthValue = {
                    key: "".concat(y_1, "-").concat(m_1),
                    type: "month",
                    value: m_1,
                    text: m_1 + filter.filiterMonthText,
                    children: [],
                    rows: [],
                    dateValues: [],
                };
                yearValue.children.push(monthValue);
            }
            var dayValue = _.find(monthValue.children, function (v) { return v.value === d_1; });
            if (dayValue == null) {
                dayValue = {
                    key: dateStr,
                    type: "day",
                    value: d_1,
                    text: d_1,
                    children: [],
                    rows: [],
                    dateValues: [],
                };
                monthValue.children.push(dayValue);
            }
            yearValue.rows.push(r);
            yearValue.dateValues.push(dateStr);
            monthValue.rows.push(r);
            monthValue.dateValues.push(dateStr);
            dayValue.rows.push(r);
            dayValue.dateValues.push(dateStr);
            dateRowMap[dateStr] = (dateRowMap[dateStr] || []).concat(r);
            if (r in hiddenRows) {
                datesUncheck = _.union(datesUncheck, [dateStr]);
            }
        }
        else {
            var v = void 0;
            var m_2;
            if (cell == null || isRealNull(cell.v)) {
                v = null;
                m_2 = null;
            }
            else {
                v = cell.v;
                m_2 = cell.m;
            }
            var data = valuesMap.get("".concat(v));
            var text = m_2 == null ? filter.valueBlank : "".concat(m_2);
            var key = "".concat(v, "#$$$#").concat(m_2);
            if (data != null) {
                var maskValue = _.find(data, function (value) { return value.mask === m_2; });
                if (maskValue == null) {
                    maskValue = {
                        key: key,
                        value: v,
                        text: text,
                        mask: m_2,
                        rows: [],
                    };
                    data.push(maskValue);
                    flattenValues.push(text);
                }
                maskValue.rows.push(r);
            }
            else {
                valuesMap.set("".concat(v), [{ key: key, value: v, text: text, mask: m_2, rows: [r] }]);
                flattenValues.push(text);
            }
            if (r in hiddenRows) {
                valuesUncheck = _.union(valuesUncheck, [key]);
            }
            valueRowMap[key] = (valueRowMap[key] || []).concat(r);
        }
    };
    for (var r = startRow + 1; r <= endRow; r += 1) {
        _loop_1(r);
    }
    return {
        dates: dates,
        datesUncheck: datesUncheck,
        dateRowMap: dateRowMap,
        values: _.flatten(Array.from(valuesMap.values())),
        valuesUncheck: valuesUncheck,
        valueRowMap: valueRowMap,
        visibleRows: visibleRows,
        flattenValues: flattenValues,
    };
}
export function getFilterColumnColors(ctx, col, startRow, endRow) {
    var _a;
    // 遍历筛选列颜色
    var bgMap = new Map(); // 单元格颜色
    var fcMap = new Map(); // 字体颜色
    // const af_compute = alternateformat.getComputeMap();
    var cf_compute = getComputeMap(ctx);
    var flowdata = getFlowdata(ctx);
    if (flowdata == null)
        return { bgColors: [], fcColors: [] };
    for (var r = startRow + 1; r <= endRow; r += 1) {
        var cell = flowdata[r][col];
        // 单元格颜色
        var bg = normalizedAttr(flowdata, r, col, "bg");
        if (bg == null) {
            bg = "#ffffff";
        }
        // const checksAF = alternateformat.checksAF(r, col, af_compute);
        var checksAF = [];
        if (checksAF.length > 1) {
            // 若单元格有交替颜色
            bg = checksAF[1];
        }
        var checksCF = checkCF(r, col, cf_compute);
        if (checksCF != null && checksCF.cellColor != null) {
            // 若单元格有条件格式
            bg = checksCF.cellColor;
        }
        if (bg.indexOf("rgb") > -1) {
            bg = rgbToHex(bg);
        }
        if (bg.length === 4) {
            bg =
                bg.substr(0, 1) +
                    bg.substr(1, 1).repeat(2) +
                    bg.substr(2, 1).repeat(2) +
                    bg.substr(3, 1).repeat(2);
        }
        // 字体颜色
        var fc = normalizedAttr(flowdata, r, col, "fc");
        if (checksAF.length > 0) {
            // 若单元格有交替颜色
            fc = checksAF[0];
        }
        if (checksCF != null && checksCF.textColor != null) {
            // 若单元格有条件格式
            fc = checksCF.textColor;
        }
        if (fc != null) {
            if (fc.indexOf("rgb") > -1) {
                fc = rgbToHex(fc);
            }
            if (fc.length === 4) {
                fc =
                    fc.substr(0, 1) +
                        fc.substr(1, 1).repeat(2) +
                        fc.substr(2, 1).repeat(2) +
                        fc.substr(3, 1).repeat(2);
            }
        }
        var isRowHidden = r in (((_a = ctx.config) === null || _a === void 0 ? void 0 : _a.rowhidden) || {});
        var bgData = bgMap.get(bg);
        if (bgData != null) {
            bgData.rows.push(r);
            if (isRowHidden)
                bgData.checked = false;
        }
        else {
            bgMap.set(bg, { color: bg, checked: !isRowHidden, rows: [r] });
        }
        if (fc != null) {
            var fcData = fcMap.get(fc);
            if (fcData != null && cell != null && !isRealNull(cell.v)) {
                fcData.rows.push(r);
                if (isRowHidden)
                    fcData.checked = false;
            }
            else if (cell != null && !isRealNull(cell.v)) {
                fcMap.set(fc, { color: fc, checked: !isRowHidden, rows: [r] });
            }
        }
    }
    var bgColors = _.flatten(Array.from(bgMap.values()));
    var fcColors = _.flatten(Array.from(fcMap.values()));
    return {
        bgColors: bgColors.length < 2 ? [] : bgColors,
        fcColors: fcColors.length < 2 ? [] : fcColors,
    };
}
export function saveFilter(ctx, optionState, hiddenRows, caljs, st_r, ed_r, cindex, st_c, ed_c) {
    var otherHiddenRows = getFilterHiddenRows(ctx, cindex, st_c).otherHiddenRows;
    var rowHiddenAll = _.assign(otherHiddenRows, hiddenRows);
    labelFilterOptionState(ctx, optionState, hiddenRows, caljs, st_r, ed_r, cindex, st_c, ed_c, true);
    var cfg = _.assign({}, ctx.config);
    cfg.rowhidden = rowHiddenAll;
    // config
    ctx.config = cfg;
    var sheetIndex = getSheetIndex(ctx, ctx.currentSheetId);
    if (sheetIndex == null) {
        return;
    }
    ctx.luckysheetfile[sheetIndex].config = cfg;
    // server.saveParam("cg", Store.currentSheetIndex, cfg.rowhidden, {
    //   k: "rowhidden",
    // });
}
