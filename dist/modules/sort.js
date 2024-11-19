import numeral from "numeral";
import _ from "lodash";
import { execfunction, functionCopy, update } from ".";
import { diff, getFlowdata, isdatetime, isRealNull, isRealNum, } from "..";
import { jfrefreshgrid } from "./refresh";
export function orderbydata(isAsc, index, data) {
    if (isAsc == null) {
        isAsc = true;
    }
    var a = function (x, y) {
        var x1 = x[index];
        var y1 = y[index];
        if (x[index] != null) {
            x1 = x[index].v;
        }
        if (y[index] != null) {
            y1 = y[index].v;
        }
        if (isRealNull(x1)) {
            return isAsc ? 1 : -1;
        }
        if (isRealNull(y1)) {
            return isAsc ? -1 : 1;
        }
        if (isdatetime(x1) && isdatetime(y1)) {
            return diff(x1, y1);
        }
        if (isRealNum(x1) && isRealNum(y1)) {
            var y1Value = numeral(y1).value();
            var x1Value = numeral(x1).value();
            if (y1Value == null || x1Value == null)
                return null;
            return x1Value - y1Value;
        }
        if (!isRealNum(x1) && !isRealNum(y1)) {
            return x1.localeCompare(y1, "zh");
        }
        if (!isRealNum(x1)) {
            return 1;
        }
        if (!isRealNum(y1)) {
            return -1;
        }
        return 0;
    };
    var d = function (x, y) { return a(y, x); };
    var sortedData = _.clone(data);
    sortedData.sort(isAsc ? a : d);
    // calc row offsets
    var rowOffsets = sortedData.map(function (r, i) {
        var origIndex = _.findIndex(data, function (origR) { return origR === r; });
        return i - origIndex;
    });
    return { sortedData: sortedData, rowOffsets: rowOffsets };
}
export function sortDataRange(ctx, sheetData, dataRange, index, isAsc, str, edr, stc, edc) {
    var _a;
    var _b = orderbydata(isAsc, index, dataRange), sortedData = _b.sortedData, rowOffsets = _b.rowOffsets;
    for (var r = str; r <= edr; r += 1) {
        for (var c = stc; c <= edc; c += 1) {
            var cell = sortedData[r - str][c - stc];
            if (cell === null || cell === void 0 ? void 0 : cell.f) {
                var moveOffset = rowOffsets[r - str];
                var func = cell === null || cell === void 0 ? void 0 : cell.f;
                if (moveOffset > 0) {
                    func = "=".concat(functionCopy(ctx, func, "down", moveOffset));
                }
                else if (moveOffset < 0) {
                    func = "=".concat(functionCopy(ctx, func, "up", -moveOffset));
                }
                var funcV = execfunction(ctx, func, r, c, undefined, undefined, true);
                cell.v = funcV[1], cell.f = funcV[2];
                cell.m = update(((_a = cell.ct) === null || _a === void 0 ? void 0 : _a.fa) || "General", cell.v);
            }
            sheetData[r][c] = cell;
        }
    }
    // let allParam = {};
    // if (ctx.config.rowlen != null) {
    //   let cfg = _.assign({}, ctx.config);
    //   cfg = rowlenByRange(d, str, edr, cfg);
    //   allParam = {
    //     cfg,
    //     RowlChange: true,
    //   };
    // }
    jfrefreshgrid(ctx, sheetData, [{ row: [str, edr], column: [stc, edc] }]);
}
export function sortSelection(ctx, isAsc, colIndex) {
    var _a;
    if (colIndex === void 0) { colIndex = 0; }
    // if (!checkProtectionAuthorityNormal(ctx.currentSheetIndex, "sort")) {
    //   return;
    // }
    if (ctx.allowEdit === false)
        return;
    if (ctx.luckysheet_select_save == null)
        return;
    if (ctx.luckysheet_select_save.length > 1) {
        // if (isEditMode()) {
        //   alert("不能对多重选择区域执行此操作，请选择单个区域，然后再试");
        // } else {
        //   tooltip.info(
        //     "不能对多重选择区域执行此操作，请选择单个区域，然后再试",
        //     ""
        //   );
        // }
        return;
    }
    if (isAsc == null) {
        isAsc = true;
    }
    // const d = editor.deepCopyFlowData(Store.flowdata);
    var flowdata = getFlowdata(ctx);
    var d = flowdata;
    if (d == null)
        return;
    var r1 = ctx.luckysheet_select_save[0].row[0];
    var r2 = ctx.luckysheet_select_save[0].row[1];
    var c1 = ctx.luckysheet_select_save[0].column[0];
    var c2 = ctx.luckysheet_select_save[0].column[1];
    var str = null;
    var edr;
    for (var r = r1; r <= r2; r += 1) {
        if (d[r] != null && d[r][c1] != null) {
            var cell = d[r][c1];
            if (cell == null)
                return; //
            if (cell.mc != null || isRealNull(cell.v)) {
                continue;
            }
            if (str == null && /[\u4e00-\u9fa5]+/g.test("".concat(cell.v))) {
                str = r + 1;
                edr = r + 1;
                continue;
            }
            if (str == null) {
                str = r;
            }
            edr = r;
        }
    }
    if (str == null || str > r2) {
        return;
    }
    var hasMc = false; // 排序选区是否有合并单元格
    var data = [];
    if (edr == null)
        return;
    for (var r = str; r <= edr; r += 1) {
        var data_row = [];
        for (var c = c1; c <= c2; c += 1) {
            if (d[r][c] != null && ((_a = d[r][c]) === null || _a === void 0 ? void 0 : _a.mc) != null) {
                hasMc = true;
                break;
            }
            data_row.push(d[r][c]);
        }
        data.push(data_row);
    }
    if (hasMc) {
        // if (isEditMode()) {
        //   alert("选区有合并单元格，无法执行此操作！");
        // } else {
        //   tooltip.info("选区有合并单元格，无法执行此操作！", "");
        // }
        return;
    }
    sortDataRange(ctx, d, data, colIndex, isAsc, str, edr, c1, c2);
}
