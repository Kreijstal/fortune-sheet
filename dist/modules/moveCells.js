import _ from "lodash";
import { getdatabyselection } from "./cell";
import { getFlowdata } from "../context";
import { colLocation, colLocationByIndex, mousePosition, rowLocation, rowLocationByIndex, } from "./location";
import { hasPartMC } from "./validation";
import { locale } from "../locale";
import { getBorderInfoCompute } from "./border";
import { normalizeSelection } from "./selection";
import { getSheetIndex, isAllowEdit } from "../utils";
import { cfSplitRange } from "./conditionalFormat";
import { jfrefreshgrid } from "./refresh";
import { CFSplitRange } from "./ConditionFormat";
var dragCellThreshold = 8;
function getCellLocationByMouse(ctx, e, scrollbarX, scrollbarY, container) {
    var rect = container.getBoundingClientRect();
    var x = e.pageX - rect.left - ctx.rowHeaderWidth + scrollbarX.scrollLeft;
    var y = e.pageY - rect.top - ctx.columnHeaderHeight + scrollbarY.scrollTop;
    return {
        row: rowLocation(y, ctx.visibledatarow),
        column: colLocation(x, ctx.visibledatacolumn),
    };
}
export function onCellsMoveStart(ctx, globalCache, e, scrollbarX, scrollbarY, container) {
    var _a, _b, _c, _d;
    // if (isEditMode() || ctx.allowEdit === false) {
    var allowEdit = isAllowEdit(ctx);
    if (allowEdit === false) {
        // 此模式下禁用选区拖动
        return;
    }
    globalCache.dragCellStartPos = { x: e.pageX, y: e.pageY };
    ctx.luckysheet_cell_selected_move = true;
    ctx.luckysheet_scroll_status = true;
    var _e = getCellLocationByMouse(ctx, e, scrollbarX, scrollbarY, container), _f = _e.row, row_pre = _f[0], row = _f[1], row_index = _f[2], _g = _e.column, col_pre = _g[0], col = _g[1], col_index = _g[2];
    var range = _.last(ctx.luckysheet_select_save);
    if (range == null)
        return;
    if (row_index < range.row[0]) {
        row_index = range.row[0];
    }
    else if (row_index > range.row[1])
        _a = range.row, row_index = _a[1];
    if (col_index < range.column[0]) {
        col_index = range.column[0];
    }
    else if (col_index > range.column[1])
        _b = range.column, col_index = _b[1];
    _c = rowLocationByIndex(row_index, ctx.visibledatarow), row_pre = _c[0], row = _c[1];
    _d = colLocationByIndex(col_index, ctx.visibledatacolumn), col_pre = _d[0], col = _d[1];
    ctx.luckysheet_cell_selected_move_index = [row_index, col_index];
    var ele = document.getElementById("fortune-cell-selected-move");
    if (ele == null)
        return;
    ele.style.left = "".concat(col_pre, "px");
    ele.style.top = "".concat(row_pre, "px");
    ele.style.width = "".concat(col - col_pre - 1, "px");
    ele.style.height = "".concat(row - row_pre - 1, "px");
    ele.style.display = "block";
    e.stopPropagation();
}
export function onCellsMove(ctx, globalCache, e, scrollbarX, scrollbarY, container) {
    if (!ctx.luckysheet_cell_selected_move)
        return;
    if (globalCache.dragCellStartPos != null) {
        var deltaX = Math.abs(globalCache.dragCellStartPos.x - e.pageX);
        var deltaY = Math.abs(globalCache.dragCellStartPos.y - e.pageY);
        if (deltaX < dragCellThreshold && deltaY < dragCellThreshold) {
            return;
        }
        globalCache.dragCellStartPos = undefined;
    }
    var _a = mousePosition(e.pageX, e.pageY, ctx), x = _a[0], y = _a[1];
    var rect = container.getBoundingClientRect();
    var winH = rect.height - 20 * ctx.zoomRatio;
    var winW = rect.width - 60 * ctx.zoomRatio;
    var _b = getCellLocationByMouse(ctx, e, scrollbarX, scrollbarY, container), rowL = _b.row, column = _b.column;
    var row_pre = rowL[0], row = rowL[1];
    var col_pre = column[0], col = column[1];
    var row_index = rowL[2];
    var col_index = column[2];
    var row_index_original = ctx.luckysheet_cell_selected_move_index[0];
    var col_index_original = ctx.luckysheet_cell_selected_move_index[1];
    if (ctx.luckysheet_select_save == null)
        return;
    var row_s = ctx.luckysheet_select_save[0].row[0] - row_index_original + row_index;
    var row_e = ctx.luckysheet_select_save[0].row[1] - row_index_original + row_index;
    var col_s = ctx.luckysheet_select_save[0].column[0] - col_index_original + col_index;
    var col_e = ctx.luckysheet_select_save[0].column[1] - col_index_original + col_index;
    if (row_s < 0 || y < 0) {
        row_s = 0;
        row_e =
            ctx.luckysheet_select_save[0].row[1] -
                ctx.luckysheet_select_save[0].row[0];
    }
    if (col_s < 0 || x < 0) {
        col_s = 0;
        col_e =
            ctx.luckysheet_select_save[0].column[1] -
                ctx.luckysheet_select_save[0].column[0];
    }
    if (row_e >= ctx.visibledatarow.length - 1 || y > winH) {
        row_s =
            ctx.visibledatarow.length -
                1 -
                ctx.luckysheet_select_save[0].row[1] +
                ctx.luckysheet_select_save[0].row[0];
        row_e = ctx.visibledatarow.length - 1;
    }
    if (col_e >= ctx.visibledatacolumn.length - 1 || x > winW) {
        col_s =
            ctx.visibledatacolumn.length -
                1 -
                ctx.luckysheet_select_save[0].column[1] +
                ctx.luckysheet_select_save[0].column[0];
        col_e = ctx.visibledatacolumn.length - 1;
    }
    col_pre = col_s - 1 === -1 ? 0 : ctx.visibledatacolumn[col_s - 1];
    col = ctx.visibledatacolumn[col_e];
    row_pre = row_s - 1 === -1 ? 0 : ctx.visibledatarow[row_s - 1];
    row = ctx.visibledatarow[row_e];
    var ele = document.getElementById("fortune-cell-selected-move");
    if (ele == null)
        return;
    ele.style.left = "".concat(col_pre, "px");
    ele.style.top = "".concat(row_pre, "px");
    ele.style.width = "".concat(col - col_pre - 2, "px");
    ele.style.height = "".concat(row - row_pre - 2, "px");
    ele.style.display = "block";
}
export function onCellsMoveEnd(ctx, globalCache, e, scrollbarX, scrollbarY, container) {
    var _a, _b, _c, _d;
    // 改变选择框的位置并替换目标单元格
    if (!ctx.luckysheet_cell_selected_move)
        return;
    ctx.luckysheet_cell_selected_move = false;
    var ele = document.getElementById("fortune-cell-selected-move");
    if (ele != null)
        ele.style.display = "none";
    if (globalCache.dragCellStartPos != null) {
        globalCache.dragCellStartPos = undefined;
        return;
    }
    var _e = mousePosition(e.pageX, e.pageY, ctx), x = _e[0], y = _e[1];
    // if (
    //   !checkProtectionLockedRangeList(
    //     ctx.luckysheet_select_save,
    //     ctx.currentSheetIndex
    //   )
    // ) {
    //   return;
    // }
    var rect = container.getBoundingClientRect();
    var winH = rect.height - 20 * ctx.zoomRatio;
    var winW = rect.width - 60 * ctx.zoomRatio;
    var _f = getCellLocationByMouse(ctx, e, scrollbarX, scrollbarY, container), _g = _f.row, row_index = _g[2], _h = _f.column, col_index = _h[2];
    var allowEdit = isAllowEdit(ctx, [
        {
            row: [row_index, row_index],
            column: [col_index, col_index],
        },
    ]);
    if (!allowEdit)
        return;
    var row_index_original = ctx.luckysheet_cell_selected_move_index[0];
    var col_index_original = ctx.luckysheet_cell_selected_move_index[1];
    if (row_index === row_index_original && col_index === col_index_original) {
        return;
    }
    var d = getFlowdata(ctx);
    if (d == null || ctx.luckysheet_select_save == null)
        return;
    var last = ctx.luckysheet_select_save[ctx.luckysheet_select_save.length - 1];
    var data = _.cloneDeep(getdatabyselection(ctx, last, ctx.currentSheetId));
    var cfg = ctx.config;
    if (cfg.merge == null) {
        cfg.merge = {};
    }
    if (cfg.rowlen == null) {
        cfg.rowlen = {};
    }
    var locale_drag = locale(ctx).drag;
    // 选区包含部分单元格
    if (hasPartMC(ctx, cfg, last.row[0], last.row[1], last.column[0], last.column[1])) {
        // if (isEditMode()) {
        //   alert(locale_drag.noMerge);
        // } else {
        // drag.info(
        //   '<i class="fa fa-exclamation-triangle"></i>',
        throw new Error(locale_drag.noMerge);
        // );
        // }
        // return;
    }
    var row_s = last.row[0] - row_index_original + row_index;
    var row_e = last.row[1] - row_index_original + row_index;
    var col_s = last.column[0] - col_index_original + col_index;
    var col_e = last.column[1] - col_index_original + col_index;
    // if (
    //   !checkProtectionLockedRangeList(
    //     [{ row: [row_s, row_e], column: [col_s, col_e] }],
    //     ctx.currentSheetIndex
    //   )
    // ) {
    //   return;
    // }
    if (row_s < 0 || y < 0) {
        row_s = 0;
        row_e = last.row[1] - last.row[0];
    }
    if (col_s < 0 || x < 0) {
        col_s = 0;
        col_e = last.column[1] - last.column[0];
    }
    if (row_e >= ctx.visibledatarow.length - 1 || y > winH) {
        row_s = ctx.visibledatarow.length - 1 - last.row[1] + last.row[0];
        row_e = ctx.visibledatarow.length - 1;
    }
    if (col_e >= ctx.visibledatacolumn.length - 1 || x > winW) {
        col_s = ctx.visibledatacolumn.length - 1 - last.column[1] + last.column[0];
        col_e = ctx.visibledatacolumn.length - 1;
    }
    // 替换的位置包含部分单元格
    if (hasPartMC(ctx, cfg, row_s, row_e, col_s, col_e)) {
        // if (isEditMode()) {
        //   alert(locale_drag.noMerge);
        // } else {
        // tooltip.info(
        //   '<i class="fa fa-exclamation-triangle"></i>',
        throw new Error(locale_drag.noMerge);
        // );
        // }
        // return;
    }
    var borderInfoCompute = getBorderInfoCompute(ctx, ctx.currentSheetId);
    var hyperLinkList = {};
    // 删除原本位置的数据
    // const RowlChange = null;
    var index = getSheetIndex(ctx, ctx.currentSheetId);
    for (var r = last.row[0]; r <= last.row[1]; r += 1) {
        // if (r in cfg.rowlen) {
        //   RowlChange = true;
        // }
        for (var c = last.column[0]; c <= last.column[1]; c += 1) {
            var cellData = d[r][c];
            if ((cellData === null || cellData === void 0 ? void 0 : cellData.mc) != null) {
                var mergeKey = "".concat(cellData.mc.r, "_").concat(c);
                if (cfg.merge[mergeKey] != null) {
                    delete cfg.merge[mergeKey];
                }
            }
            d[r][c] = null;
            if ((_a = ctx.luckysheetfile[index].hyperlink) === null || _a === void 0 ? void 0 : _a["".concat(r, "_").concat(c)]) {
                hyperLinkList["".concat(r, "_").concat(c)] =
                    (_b = ctx.luckysheetfile[index].hyperlink) === null || _b === void 0 ? void 0 : _b["".concat(r, "_").concat(c)];
                (_c = ctx.luckysheetfile[getSheetIndex(ctx, ctx.currentSheetId)].hyperlink) === null || _c === void 0 ? true : delete _c["".concat(r, "_").concat(c)];
            }
        }
    }
    // 边框
    if (cfg.borderInfo && cfg.borderInfo.length > 0) {
        var borderInfo = [];
        for (var i = 0; i < cfg.borderInfo.length; i += 1) {
            var bd_rangeType = cfg.borderInfo[i].rangeType;
            if (bd_rangeType === "range" &&
                cfg.borderInfo[i].borderType !== "border-slash") {
                var bd_range = cfg.borderInfo[i].range;
                var bd_emptyRange = [];
                for (var j = 0; j < bd_range.length; j += 1) {
                    bd_emptyRange = bd_emptyRange.concat(cfSplitRange(bd_range[j], { row: last.row, column: last.column }, { row: [row_s, row_e], column: [col_s, col_e] }, "restPart"));
                }
                cfg.borderInfo[i].range = bd_emptyRange;
                borderInfo.push(cfg.borderInfo[i]);
            }
            else if (bd_rangeType === "cell") {
                var bd_r = cfg.borderInfo[i].value.row_index;
                var bd_c = cfg.borderInfo[i].value.col_index;
                if (!(bd_r >= last.row[0] &&
                    bd_r <= last.row[1] &&
                    bd_c >= last.column[0] &&
                    bd_c <= last.column[1])) {
                    borderInfo.push(cfg.borderInfo[i]);
                }
            }
            else if (bd_rangeType === "range" &&
                cfg.borderInfo[i].borderType === "border-slash" &&
                !(cfg.borderInfo[i].range[0].row[0] >= last.row[0] &&
                    cfg.borderInfo[i].range[0].row[0] <= last.row[1] &&
                    cfg.borderInfo[i].range[0].column[0] >= last.column[0] &&
                    cfg.borderInfo[i].range[0].column[0] <= last.column[1])) {
                borderInfo.push(cfg.borderInfo[i]);
            }
        }
        cfg.borderInfo = borderInfo;
    }
    // 替换位置数据更新
    var offsetMC = {};
    for (var r = 0; r < data.length; r += 1) {
        for (var c = 0; c < data[0].length; c += 1) {
            if (borderInfoCompute["".concat(r + last.row[0], "_").concat(c + last.column[0])] &&
                !borderInfoCompute["".concat(r + last.row[0], "_").concat(c + last.column[0])].s) {
                var bd_obj = {
                    rangeType: "cell",
                    value: {
                        row_index: r + row_s,
                        col_index: c + col_s,
                        l: borderInfoCompute["".concat(r + last.row[0], "_").concat(c + last.column[0])].l,
                        r: borderInfoCompute["".concat(r + last.row[0], "_").concat(c + last.column[0])].r,
                        t: borderInfoCompute["".concat(r + last.row[0], "_").concat(c + last.column[0])].t,
                        b: borderInfoCompute["".concat(r + last.row[0], "_").concat(c + last.column[0])].b,
                    },
                };
                if (cfg.borderInfo == null) {
                    cfg.borderInfo = [];
                }
                cfg.borderInfo.push(bd_obj);
            }
            else if (borderInfoCompute["".concat(r + last.row[0], "_").concat(c + last.column[0])]) {
                var bd_obj = {
                    rangeType: "range",
                    borderType: "border-slash",
                    color: borderInfoCompute["".concat(r + last.row[0], "_").concat(c + last.column[0])].s
                        .color,
                    style: borderInfoCompute["".concat(r + last.row[0], "_").concat(c + last.column[0])].s
                        .style,
                    range: normalizeSelection(ctx, [
                        { row: [r + row_s, r + row_s], column: [c + col_s, c + col_s] },
                    ]),
                };
                if (cfg.borderInfo == null) {
                    cfg.borderInfo = [];
                }
                cfg.borderInfo.push(bd_obj);
            }
            var value = null;
            if (data[r] != null && data[r][c] != null) {
                value = data[r][c];
            }
            if ((value === null || value === void 0 ? void 0 : value.mc) != null) {
                var mc = _.assign({}, value.mc);
                if ("rs" in value.mc) {
                    _.set(offsetMC, "".concat(mc.r, "_").concat(mc.c), [r + row_s, c + col_s]);
                    value.mc.r = r + row_s;
                    value.mc.c = c + col_s;
                    _.set(cfg.merge, "".concat(r + row_s, "_").concat(c + col_s), value.mc);
                }
                else {
                    _.set(value.mc, "r", offsetMC["".concat(mc.r, "_").concat(mc.c)][0]);
                    _.set(value.mc, "c", offsetMC["".concat(mc.r, "_").concat(mc.c)][1]);
                }
            }
            d[r + row_s][c + col_s] = value;
            if (hyperLinkList === null || hyperLinkList === void 0 ? void 0 : hyperLinkList["".concat(r + last.row[0], "_").concat(c + last.column[0])]) {
                ctx.luckysheetfile[index].hyperlink["".concat(r + row_s, "_").concat(c + col_s)] =
                    hyperLinkList === null || hyperLinkList === void 0 ? void 0 : hyperLinkList["".concat(r + last.row[0], "_").concat(c + last.column[0])];
            }
        }
    }
    // if (RowlChange) {
    //   cfg = rowlenByRange(d, last.row[0], last.row[1], cfg);
    //   cfg = rowlenByRange(d, row_s, row_e, cfg);
    // }
    // 条件格式
    var cdformat = (_d = ctx.luckysheetfile[getSheetIndex(ctx, ctx.currentSheetId)]
        .luckysheet_conditionformat_save) !== null && _d !== void 0 ? _d : [];
    if (cdformat != null && cdformat.length > 0) {
        for (var i = 0; i < cdformat.length; i += 1) {
            var cdformat_cellrange = cdformat[i].cellrange;
            var emptyRange = [];
            for (var j = 0; j < cdformat_cellrange.length; j += 1) {
                var range_1 = CFSplitRange(cdformat_cellrange[j], { row: last.row, column: last.column }, { row: [row_s, row_e], column: [col_s, col_e] }, "allPart");
                emptyRange = emptyRange.concat(range_1);
            }
            cdformat[i].cellrange = emptyRange;
        }
    }
    var rf;
    if (ctx.luckysheet_select_save[0].row_focus ===
        ctx.luckysheet_select_save[0].row[0]) {
        rf = row_s;
    }
    else {
        rf = row_e;
    }
    var cf;
    if (ctx.luckysheet_select_save[0].column_focus ===
        ctx.luckysheet_select_save[0].column[0]) {
        cf = col_s;
    }
    else {
        cf = col_e;
    }
    var range = [];
    range.push({ row: last.row, column: last.column });
    range.push({ row: [row_s, row_e], column: [col_s, col_e] });
    last.row = [row_s, row_e];
    last.column = [col_s, col_e];
    last.row_focus = rf;
    last.column_focus = cf;
    ctx.luckysheet_select_save = normalizeSelection(ctx, [last]);
    var sheetIndex = getSheetIndex(ctx, ctx.currentSheetId);
    if (sheetIndex != null) {
        ctx.luckysheetfile[sheetIndex].config = _.assign({}, cfg);
    }
    // const allParam = {
    //   cfg,
    //   RowlChange,
    //   cdformat,
    // };
    jfrefreshgrid(ctx, d, range);
    // selectHightlightShow();
    // $("#luckysheet-sheettable").css("cursor", "default");
    // clearTimeout(ctx.countfuncTimeout);
    // ctx.countfuncTimeout = setTimeout(function () {
    //   countfunc();
    // }, 500);
}
