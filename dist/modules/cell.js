import _ from "lodash";
import { getFlowdata } from "../context";
import { getSheetIndex, indexToColumnChar, rgbToHex } from "../utils";
import { checkCF, getComputeMap } from "./ConditionFormat";
import { getFailureText, validateCellData } from "./dataVerification";
import { genarate, update } from "./format";
import { delFunctionGroup, execfunction, execFunctionGroup, functionHTMLGenerate, getcellrange, iscelldata, } from "./formula";
import { attrToCssName, convertSpanToShareString, isInlineStringCell, isInlineStringCT, } from "./inline-string";
import { isRealNull, isRealNum, valueIsError } from "./validation";
import { getCellTextInfo } from "./text";
// TODO put these in context ref
// let rangestart = false;
// let rangedrag_column_start = false;
// let rangedrag_row_start = false;
export function normalizedCellAttr(cell, attr, defaultFontSize) {
    if (defaultFontSize === void 0) { defaultFontSize = 10; }
    var tf = { bl: 1, it: 1, ff: 1, cl: 1, un: 1 };
    var value = cell === null || cell === void 0 ? void 0 : cell[attr];
    if (attr in tf || (attr === "fs" && isInlineStringCell(cell))) {
        value || (value = "0");
    }
    else if (["fc", "bg", "bc"].includes(attr)) {
        if (["fc", "bc"].includes(attr)) {
            value || (value = "#000000");
        }
        if ((value === null || value === void 0 ? void 0 : value.indexOf("rgba")) > -1) {
            value = rgbToHex(value);
        }
    }
    else if (attr.substring(0, 2) === "bs") {
        value || (value = "none");
    }
    else if (attr === "ht" || attr === "vt") {
        var defaultValue = attr === "ht" ? "1" : "0";
        value = !_.isNil(value) ? value.toString() : defaultValue;
        if (["0", "1", "2"].indexOf(value.toString()) === -1) {
            value = defaultValue;
        }
    }
    else if (attr === "fs") {
        value || (value = defaultFontSize.toString());
    }
    else if (attr === "tb" || attr === "tr") {
        value || (value = "0");
    }
    return value;
}
export function normalizedAttr(data, r, c, attr) {
    if (!data || !data[r]) {
        console.warn("cell (%d, %d) is null", r, c);
        return null;
    }
    var cell = data[r][c];
    if (!cell)
        return undefined;
    return normalizedCellAttr(cell, attr);
}
export function getCellValue(r, c, data, attr) {
    if (!attr) {
        attr = "v";
    }
    var d_value;
    if (!_.isNil(r) && !_.isNil(c)) {
        d_value = data[r][c];
    }
    else if (!_.isNil(r)) {
        d_value = data[r];
    }
    else if (!_.isNil(c)) {
        var newData = data[0].map(function (col, i) {
            return data.map(function (row) {
                return row[i];
            });
        });
        d_value = newData[c];
    }
    else {
        return data;
    }
    var retv = d_value;
    if (_.isPlainObject(d_value)) {
        var d = d_value;
        retv = d[attr];
        if (attr === "f" && !_.isNil(retv)) {
            retv = functionHTMLGenerate(retv);
        }
        else if (attr === "f") {
            retv = d.v;
        }
        else if (d && d.ct && d.ct.t === "d") {
            retv = d.m;
        }
    }
    if (retv === undefined) {
        retv = null;
    }
    return retv;
}
export function setCellValue(ctx, r, c, d, v) {
    var _a, _b, _c, _d, _e, _f;
    if (_.isNil(d)) {
        d = getFlowdata(ctx);
    }
    if (!d)
        return;
    // 若采用深拷贝，初始化时的单元格属性丢失
    // let cell = $.extend(true, {}, d[r][c]);
    var cell = d[r][c];
    var vupdate;
    if (_.isPlainObject(v)) {
        if (_.isNil(cell)) {
            cell = v;
        }
        else {
            if (!_.isNil(v.f)) {
                cell.f = v.f;
            }
            else if ("f" in cell) {
                delete cell.f;
            }
            // if (!_.isNil(v.spl)) {
            //   cell.spl = v.spl;
            // }
            if (!_.isNil(v.ct)) {
                cell.ct = v.ct;
            }
        }
        if (_.isPlainObject(v.v)) {
            vupdate = v.v.v;
        }
        else {
            vupdate = v.v;
        }
    }
    else {
        vupdate = v;
    }
    if (isRealNull(vupdate)) {
        if (_.isPlainObject(cell)) {
            delete cell.m;
            // @ts-ignore
            delete cell.v;
        }
        else {
            cell = null;
        }
        d[r][c] = cell;
        return;
    }
    // 1.为null
    // 2.数据透视表的数据，flowdata的每个数据可能为字符串，结果就是cell === v === 一个字符串或者数字数据
    if (isRealNull(cell) ||
        ((_.isString(cell) || _.isNumber(cell)) && cell === v)) {
        cell = {};
    }
    if (!cell)
        return;
    var vupdateStr = vupdate.toString();
    if (vupdateStr.substr(0, 1) === "'") {
        cell.m = vupdateStr.substr(1);
        cell.ct = { fa: "@", t: "s" };
        cell.v = vupdateStr.substr(1);
        cell.qp = 1;
    }
    else if (cell.qp === 1) {
        cell.m = vupdateStr;
        cell.ct = { fa: "@", t: "s" };
        cell.v = vupdateStr;
    }
    else if (vupdateStr.toUpperCase() === "TRUE" &&
        (_.isNil((_a = cell.ct) === null || _a === void 0 ? void 0 : _a.fa) || ((_b = cell.ct) === null || _b === void 0 ? void 0 : _b.fa) !== "@")) {
        cell.m = "TRUE";
        cell.ct = { fa: "General", t: "b" };
        cell.v = true;
    }
    else if (vupdateStr.toUpperCase() === "FALSE" &&
        (_.isNil((_c = cell.ct) === null || _c === void 0 ? void 0 : _c.fa) || ((_d = cell.ct) === null || _d === void 0 ? void 0 : _d.fa) !== "@")) {
        cell.m = "FALSE";
        cell.ct = { fa: "General", t: "b" };
        cell.v = false;
    }
    else if (vupdateStr.substr(-1) === "%" &&
        isRealNum(vupdateStr.substring(0, vupdateStr.length - 1)) &&
        (_.isNil((_e = cell.ct) === null || _e === void 0 ? void 0 : _e.fa) || ((_f = cell.ct) === null || _f === void 0 ? void 0 : _f.fa) !== "@")) {
        cell.ct = { fa: "0%", t: "n" };
        cell.v = vupdateStr.substring(0, vupdateStr.length - 1) / 100;
        cell.m = vupdate;
    }
    else if (valueIsError(vupdate)) {
        cell.m = vupdateStr;
        // cell.ct = { "fa": "General", "t": "e" };
        if (!_.isNil(cell.ct)) {
            cell.ct.t = "e";
        }
        else {
            cell.ct = { fa: "General", t: "e" };
        }
        cell.v = vupdate;
    }
    else {
        if (!_.isNil(cell.f) &&
            isRealNum(vupdate) &&
            !/^\d{6}(18|19|20)?\d{2}(0[1-9]|1[12])(0[1-9]|[12]\d|3[01])\d{3}(\d|X)$/i.test(vupdate)) {
            cell.v = parseFloat(vupdate);
            if (_.isNil(cell.ct)) {
                cell.ct = { fa: "General", t: "n" };
            }
            if (cell.v === Infinity || cell.v === -Infinity) {
                cell.m = cell.v.toString();
            }
            else {
                if (cell.v.toString().indexOf("e") > -1) {
                    var len = void 0;
                    if (cell.v.toString().split(".").length === 1) {
                        len = 0;
                    }
                    else {
                        len = cell.v.toString().split(".")[1].split("e")[0].length;
                    }
                    if (len > 5) {
                        len = 5;
                    }
                    cell.m = cell.v.toExponential(len).toString();
                }
                else {
                    var v_p = Math.round(cell.v * 1000000000) / 1000000000;
                    if (_.isNil(cell.ct)) {
                        var mask = genarate(v_p);
                        if (mask != null) {
                            cell.m = mask[0].toString();
                        }
                    }
                    else {
                        var mask = update(cell.ct.fa, v_p);
                        cell.m = mask.toString();
                    }
                    // cell.m = mask[0].toString();
                }
            }
        }
        else if (!_.isNil(cell.ct) && cell.ct.fa === "@") {
            cell.m = vupdateStr;
            cell.v = vupdate;
        }
        else if (cell.ct != null && cell.ct.t === "d" && _.isString(vupdate)) {
            var mask = genarate(vupdate);
            if (mask[1].t !== "d" || mask[1].fa === cell.ct.fa) {
                cell.m = mask[0], cell.ct = mask[1], cell.v = mask[2];
            }
            else {
                cell.v = mask[2];
                cell.m = update(cell.ct.fa, cell.v);
            }
        }
        else if (!_.isNil(cell.ct) &&
            !_.isNil(cell.ct.fa) &&
            cell.ct.fa !== "General") {
            if (isRealNum(vupdate)) {
                vupdate = parseFloat(vupdate);
            }
            var mask = update(cell.ct.fa, vupdate);
            if (mask === vupdate) {
                // 若原来单元格格式 应用不了 要更新的值，则获取更新值的 格式
                mask = genarate(vupdate);
                cell.m = mask[0].toString();
                cell.ct = mask[1], cell.v = mask[2];
            }
            else {
                cell.m = mask.toString();
                cell.v = vupdate;
            }
        }
        else {
            if (isRealNum(vupdate) &&
                !/^\d{6}(18|19|20)?\d{2}(0[1-9]|1[12])(0[1-9]|[12]\d|3[01])\d{3}(\d|X)$/i.test(vupdate)) {
                if (typeof vupdate === "string") {
                    var flag = vupdate
                        .split("")
                        .every(function (ele) { return ele === "0" || ele === "."; });
                    if (flag) {
                        vupdate = parseFloat(vupdate);
                    }
                }
                cell.v =
                    vupdate; /* 备注：如果使用parseFloat，1.1111111111111111会转换为1.1111111111111112 ? */
                cell.ct = { fa: "General", t: "n" };
                if (cell.v === Infinity || cell.v === -Infinity) {
                    cell.m = cell.v.toString();
                }
                else if (cell.v != null) {
                    var mask = genarate(cell.v);
                    if (mask) {
                        cell.m = mask[0].toString();
                    }
                }
            }
            else {
                var mask = genarate(vupdate);
                if (mask) {
                    cell.m = mask[0].toString();
                    cell.ct = mask[1], cell.v = mask[2];
                }
            }
        }
    }
    // if (!server.allowUpdate && !luckysheetConfigsetting.pointEdit) {
    //   if (
    //     !_.isNil(cell.ct) &&
    //     /^(w|W)((0?)|(0\.0+))$/.test(cell.ct.fa) === false &&
    //     cell.ct.t === "n" &&
    //     !_.isNil(cell.v) &&
    //     parseInt(cell.v, 10).toString().length > 4
    //   ) {
    //     const autoFormatw = luckysheetConfigsetting.autoFormatw
    //       .toString()
    //       .toUpperCase();
    //     const { accuracy } = luckysheetConfigsetting;
    //     const sfmt = setAccuracy(autoFormatw, accuracy);
    //     if (sfmt !== "General") {
    //       cell.ct.fa = sfmt;
    //       cell.m = update(sfmt, cell.v);
    //     }
    //   }
    // }
    d[r][c] = cell;
}
export function getRealCellValue(r, c, data, attr) {
    var value = getCellValue(r, c, data, "m");
    if (_.isNil(value)) {
        value = getCellValue(r, c, data, attr);
        if (_.isNil(value)) {
            var ct = getCellValue(r, c, data, "ct");
            if (isInlineStringCT(ct)) {
                value = ct.s;
            }
        }
    }
    return value;
}
export function mergeBorder(ctx, d, row_index, col_index) {
    var _a, _b, _c, _d, _e, _f, _g, _h, _j;
    if (!d || !d[row_index]) {
        console.warn("Merge info is null", row_index, col_index);
        return null;
    }
    var value = d[row_index][col_index];
    if (!value)
        return null;
    if (value === null || value === void 0 ? void 0 : value.mc) {
        var margeMaindata = value.mc;
        if (!margeMaindata) {
            console.warn("Merge info is null", row_index, col_index);
            return null;
        }
        col_index = margeMaindata.c;
        row_index = margeMaindata.r;
        if (_.isNil((_a = d === null || d === void 0 ? void 0 : d[row_index]) === null || _a === void 0 ? void 0 : _a[col_index])) {
            console.warn("Main merge Cell info is null", row_index, col_index);
            return null;
        }
        var col_rs = (_d = (_c = (_b = d[row_index]) === null || _b === void 0 ? void 0 : _b[col_index]) === null || _c === void 0 ? void 0 : _c.mc) === null || _d === void 0 ? void 0 : _d.cs;
        var row_rs = (_g = (_f = (_e = d[row_index]) === null || _e === void 0 ? void 0 : _e[col_index]) === null || _f === void 0 ? void 0 : _f.mc) === null || _g === void 0 ? void 0 : _g.rs;
        var mergeMain = (_j = (_h = d[row_index]) === null || _h === void 0 ? void 0 : _h[col_index]) === null || _j === void 0 ? void 0 : _j.mc;
        if (!mergeMain ||
            _.isNil(mergeMain === null || mergeMain === void 0 ? void 0 : mergeMain.rs) ||
            _.isNil(mergeMain === null || mergeMain === void 0 ? void 0 : mergeMain.cs) ||
            _.isNil(col_rs) ||
            _.isNil(row_rs)) {
            console.warn("Main merge info is null", mergeMain);
            return null;
        }
        var start_r = void 0;
        var end_r = void 0;
        var row = void 0;
        var row_pre = void 0;
        for (var r = row_index; r < mergeMain.rs + row_index; r += 1) {
            if (r === 0) {
                start_r = -1;
            }
            else {
                start_r = ctx.visibledatarow[r - 1] - 1;
            }
            end_r = ctx.visibledatarow[r];
            if (row_pre === undefined) {
                row_pre = start_r;
                row = end_r;
            }
            else if (row !== undefined) {
                row += end_r - start_r - 1;
            }
        }
        var start_c = void 0;
        var end_c = void 0;
        var col = void 0;
        var col_pre = void 0;
        for (var c = col_index; c < mergeMain.cs + col_index; c += 1) {
            if (c === 0) {
                start_c = 0;
            }
            else {
                start_c = ctx.visibledatacolumn[c - 1];
            }
            end_c = ctx.visibledatacolumn[c];
            if (col_pre === undefined) {
                col_pre = start_c;
                col = end_c;
            }
            else if (col !== undefined) {
                col += end_c - start_c;
            }
        }
        if (_.isNil(row_pre) || _.isNil(col_pre) || _.isNil(row) || _.isNil(col)) {
            console.warn("Main merge info row_pre or col_pre or row or col is null", mergeMain);
            return null;
        }
        return {
            row: [row_pre, row, row_index, row_index + row_rs - 1],
            column: [col_pre, col, col_index, col_index + col_rs - 1],
        };
    }
    return null;
}
function mergeMove(ctx, mc, columnseleted, rowseleted, s, top, height, left, width) {
    var row_st = mc.r;
    var row_ed = mc.r + mc.rs - 1;
    var col_st = mc.c;
    var col_ed = mc.c + mc.cs - 1;
    var ismatch = false;
    columnseleted[0] = Math.min(columnseleted[0], columnseleted[1]);
    rowseleted[0] = Math.min(rowseleted[0], rowseleted[1]);
    if ((columnseleted[0] <= col_st &&
        columnseleted[1] >= col_ed &&
        rowseleted[0] <= row_st &&
        rowseleted[1] >= row_ed) ||
        (!(columnseleted[1] < col_st || columnseleted[0] > col_ed) &&
            !(rowseleted[1] < row_st || rowseleted[0] > row_ed))) {
        var flowdata = getFlowdata(ctx);
        if (!flowdata)
            return null;
        var margeset = mergeBorder(ctx, flowdata, mc.r, mc.c);
        if (margeset) {
            var row = margeset.row[1];
            var row_pre = margeset.row[0];
            var col = margeset.column[1];
            var col_pre = margeset.column[0];
            if (!(columnseleted[1] < col_st || columnseleted[0] > col_ed)) {
                // 向上滑动
                if (rowseleted[0] <= row_ed && rowseleted[0] >= row_st) {
                    height += top - row_pre;
                    top = row_pre;
                    rowseleted[0] = row_st;
                }
                // 向下滑动或者居中时往上滑动的向下补齐
                if (rowseleted[1] >= row_st && rowseleted[1] <= row_ed) {
                    if (s.row_focus >= row_st && s.row_focus <= row_ed) {
                        height = row - top;
                    }
                    else {
                        height = row - top;
                    }
                    rowseleted[1] = row_ed;
                }
            }
            if (!(rowseleted[1] < row_st || rowseleted[0] > row_ed)) {
                if (columnseleted[0] <= col_ed && columnseleted[0] >= col_st) {
                    width += left - col_pre;
                    left = col_pre;
                    columnseleted[0] = col_st;
                }
                // 向右滑动或者居中时往左滑动的向下补齐
                if (columnseleted[1] >= col_st && columnseleted[1] <= col_ed) {
                    if (s.column_focus >= col_st && s.column_focus <= col_ed) {
                        width = col - left;
                    }
                    else {
                        width = col - left;
                    }
                    columnseleted[1] = col_ed;
                }
            }
            ismatch = true;
        }
    }
    if (ismatch) {
        return [columnseleted, rowseleted, top, height, left, width];
    }
    return null;
}
export function mergeMoveMain(ctx, columnseleted, rowseleted, s, top, height, left, width) {
    var mergesetting = ctx.config.merge;
    if (!mergesetting) {
        return null;
    }
    var mcset = Object.keys(mergesetting);
    rowseleted[1] = Math.max(rowseleted[0], rowseleted[1]);
    columnseleted[1] = Math.max(columnseleted[0], columnseleted[1]);
    var offloop = true;
    var mergeMoveData = {};
    while (offloop) {
        offloop = false;
        for (var i = 0; i < mcset.length; i += 1) {
            var key = mcset[i];
            var mc = mergesetting[key];
            if (key in mergeMoveData) {
                continue;
            }
            var changeparam = mergeMove(ctx, mc, columnseleted, rowseleted, s, top, height, left, width);
            if (changeparam != null) {
                mergeMoveData[key] = mc;
                // @ts-ignore
                columnseleted = changeparam[0], rowseleted = changeparam[1], top = changeparam[2], height = changeparam[3], left = changeparam[4], width = changeparam[5];
                offloop = true;
            }
            else {
                delete mergeMoveData[key];
            }
        }
    }
    return [columnseleted, rowseleted, top, height, left, width];
}
// eslint-disable-next-line @typescript-eslint/no-unused-vars
export function cancelFunctionrangeSelected(ctx) {
    if (ctx.formulaCache.selectingRangeIndex === -1) {
        ctx.formulaRangeSelect = undefined;
    }
    // $("#luckysheet-row-count-show, #luckysheet-column-count-show").hide();
    // // $("#luckysheet-cols-h-selected, #luckysheet-rows-h-selected").hide();
    // $("#luckysheet-formula-search-c, #luckysheet-formula-help-c").hide();
}
export function cancelNormalSelected(ctx) {
    cancelFunctionrangeSelected(ctx);
    ctx.luckysheetCellUpdate = [];
    ctx.formulaRangeHighlight = [];
    ctx.functionHint = null;
    // $("#fortune-formula-functionrange .fortune-formula-functionrange-highlight").remove();
    // $("#luckysheet-input-box").removeAttr("style");
    // $("#luckysheet-input-box-index").hide();
    // $("#luckysheet-wa-functionbox-cancel, #luckysheet-wa-functionbox-confirm").removeClass("luckysheet-wa-calculate-active");
    ctx.formulaCache.rangestart = false;
    ctx.formulaCache.rangedrag_column_start = false;
    ctx.formulaCache.rangedrag_row_start = false;
}
// formula.updatecell
export function updateCell(ctx, r, c, $input, value, canvas) {
    var _a, _b, _c, _d, _e, _f, _g, _h, _j;
    var inputText = $input === null || $input === void 0 ? void 0 : $input.innerText;
    var inputHtml = $input === null || $input === void 0 ? void 0 : $input.innerHTML;
    var flowdata = getFlowdata(ctx);
    if (!flowdata)
        return;
    // if (!_.isNil(rangetosheet) && rangetosheet !== ctx.currentSheetId) {
    //   sheetmanage.changeSheetExec(rangetosheet);
    // }
    // if (!checkProtectionLocked(r, c, ctx.currentSheetId)) {
    //   return;
    // }
    // 数据验证 输入数据无效时禁止输入
    var index = getSheetIndex(ctx, ctx.currentSheetId);
    var dataVerification = ctx.luckysheetfile[index].dataVerification;
    if (!_.isNil(dataVerification)) {
        var dvItem = dataVerification["".concat(r, "_").concat(c)];
        if (!_.isNil(dvItem) &&
            dvItem.prohibitInput &&
            !validateCellData(ctx, dvItem, inputText)) {
            var failureText = getFailureText(ctx, dvItem);
            cancelNormalSelected(ctx);
            ctx.warnDialog = failureText;
            return;
        }
    }
    var curv = flowdata[r][c];
    // ctx.old value for hook function
    var oldValue = _.cloneDeep(curv);
    var isPrevInline = isInlineStringCell(curv);
    var isCurInline = (inputText === null || inputText === void 0 ? void 0 : inputText.slice(0, 1)) !== "=" && (inputHtml === null || inputHtml === void 0 ? void 0 : inputHtml.substring(0, 5)) === "<span";
    var isCopyVal = false;
    if (!isCurInline && inputText && inputText.length > 0) {
        var splitArr = inputText
            .replace(/\r\n/g, "_x000D_")
            .replace(/&#13;&#10;/g, "_x000D_")
            .replace(/\r/g, "_x000D_")
            .replace(/\n/g, "_x000D_")
            .split("_x000D_");
        if (splitArr.length > 1 && inputHtml !== "<br>") {
            isCopyVal = true;
            isCurInline = true;
            inputText = splitArr.join("\r\n");
        }
    }
    if ((curv === null || curv === void 0 ? void 0 : curv.ct) && !value && !isCurInline && isPrevInline) {
        delete curv.ct.s;
        curv.ct.t = "g";
        curv.ct.fa = "General";
        value = "";
    }
    else if (isCurInline) {
        if (!_.isPlainObject(curv)) {
            curv = {};
        }
        curv || (curv = {});
        var fontSize = curv.fs || 10;
        if (!curv.ct) {
            curv.ct = {};
            curv.ct.fa = "General";
        }
        curv.ct.t = "inlineStr";
        curv.ct.s = convertSpanToShareString($input.querySelectorAll("span"), curv);
        delete curv.fs;
        delete curv.f;
        delete curv.v;
        delete curv.m;
        curv.fs = fontSize;
        if (isCopyVal) {
            curv.ct.s = [
                {
                    v: inputText,
                    fs: fontSize,
                },
            ];
        }
    }
    // API, we get value from user
    value = value || ($input === null || $input === void 0 ? void 0 : $input.innerText);
    // Hook function
    if (((_b = (_a = ctx.hooks).beforeUpdateCell) === null || _b === void 0 ? void 0 : _b.call(_a, r, c, value)) === false) {
        cancelNormalSelected(ctx);
        return;
    }
    if (!isCurInline) {
        if (isRealNull(value) && !isPrevInline) {
            if (!curv || (isRealNull(curv.v) && !curv.spl && !curv.f)) {
                cancelNormalSelected(ctx);
                return;
            }
        }
        else if (curv && curv.qp !== 1) {
            if (_.isPlainObject(curv) &&
                (value === curv.f || value === curv.v || value === curv.m)) {
                cancelNormalSelected(ctx);
                return;
            }
            if (value === curv) {
                cancelNormalSelected(ctx);
                return;
            }
        }
        if (_.isString(value) && value.slice(0, 1) === "=" && value.length > 1) {
        }
        else if (_.isPlainObject(curv) &&
            curv &&
            curv.ct &&
            curv.ct.fa &&
            curv.ct.fa !== "@" &&
            !isRealNull(value)) {
            delete curv.m; // 更新时间m处理 ， 会实际删除单元格数据的参数（flowdata时已删除）
            if (curv.f) {
                // 如果原来是公式，而更新的数据不是公式，则把公式删除
                delete curv.f;
                delete curv.spl; // 删除单元格的sparklines的配置串
            }
        }
    }
    // TODO window.luckysheet_getcelldata_cache = null;
    var isRunExecFunction = true;
    var d = flowdata; // TODO const d = editor.deepCopyFlowData(flowdata);
    var dynamicArrayItem = null; // 动态数组
    if (_.isPlainObject(curv)) {
        if (!isCurInline) {
            if (_.isString(value) && value.slice(0, 1) === "=" && value.length > 1) {
                var v = execfunction(ctx, value, r, c, undefined, undefined, true);
                isRunExecFunction = false;
                curv = _.cloneDeep(((_c = d === null || d === void 0 ? void 0 : d[r]) === null || _c === void 0 ? void 0 : _c[c]) || {});
                curv.v = v[1], curv.f = v[2];
                // 打进单元格的sparklines的配置串， 报错需要单独处理。
                if (v.length === 4 && v[3].type === "sparklines") {
                    delete curv.m;
                    delete curv.v;
                    var curCalv = v[3].data;
                    if (_.isArray(curCalv) && !_.isPlainObject(curCalv[0])) {
                        curv.v = curCalv[0];
                    }
                    else {
                        curv.spl = v[3].data;
                    }
                }
                else if (v.length === 4 && v[3].type === "dynamicArrayItem") {
                    dynamicArrayItem = v[3].data;
                }
            }
            // from API setCellValue,luckysheet.setCellValue(0, 0, {f: "=sum(D1)", bg:"#0188fb"}),value is an object, so get attribute f as value
            else if (_.isPlainObject(value)) {
                var valueFunction = value.f;
                if (_.isString(valueFunction) &&
                    valueFunction.slice(0, 1) === "=" &&
                    valueFunction.length > 1) {
                    var v = execfunction(ctx, valueFunction, r, c, undefined, undefined, true);
                    isRunExecFunction = false;
                    // get v/m/ct
                    curv = _.cloneDeep(((_d = d === null || d === void 0 ? void 0 : d[r]) === null || _d === void 0 ? void 0 : _d[c]) || {});
                    curv.v = v[1], curv.f = v[2];
                    // 打进单元格的sparklines的配置串， 报错需要单独处理。
                    if (v.length === 4 && v[3].type === "sparklines") {
                        delete curv.m;
                        delete curv.v;
                        var curCalv = v[3].data;
                        if (_.isArray(curCalv) && !_.isPlainObject(curCalv[0])) {
                            curv.v = curCalv[0];
                        }
                        else {
                            curv.spl = v[3].data;
                        }
                    }
                    else if (v.length === 4 && v[3].type === "dynamicArrayItem") {
                        dynamicArrayItem = v[3].data;
                    }
                }
                // from API setCellValue,luckysheet.setCellValue(0, 0, {f: "=sum(D1)", bg:"#0188fb"}),value is an object, so get attribute f as value
                else {
                    Object.keys(value).forEach(function (attr) {
                        curv[attr] = value[attr];
                    });
                }
            }
            else {
                delFunctionGroup(ctx, r, c);
                execFunctionGroup(ctx, r, c, value);
                isRunExecFunction = false;
                curv = _.cloneDeep(((_e = d === null || d === void 0 ? void 0 : d[r]) === null || _e === void 0 ? void 0 : _e[c]) || {});
                curv.v = value;
                delete curv.f;
                delete curv.spl;
                if (curv.qp === 1 && "".concat(value).substring(0, 1) !== "'") {
                    // if quotePrefix is 1, cell is force string, cell clear quotePrefix when it is updated
                    curv.qp = 0;
                    if (curv.ct) {
                        curv.ct.fa = "General";
                        curv.ct.t = "n";
                    }
                }
            }
        }
        value = curv;
    }
    else {
        if (_.isString(value) && value.slice(0, 1) === "=" && value.length > 1) {
            var v = execfunction(ctx, value, r, c, undefined, undefined, true);
            isRunExecFunction = false;
            value = {
                v: v[1],
                f: v[2],
            };
            // 打进单元格的sparklines的配置串， 报错需要单独处理。
            if (v.length === 4 && v[3].type === "sparklines") {
                var curCalv = v[3].data;
                if (_.isArray(curCalv) && !_.isPlainObject(curCalv[0])) {
                    value.v = curCalv[0];
                }
                else {
                    value.spl = v[3].data;
                }
            }
            else if (v.length === 4 && v[3].type === "dynamicArrayItem") {
                dynamicArrayItem = v[3].data;
            }
        }
        // from API setCellValue,luckysheet.setCellValue(0, 0, {f: "=sum(D1)", bg:"#0188fb"}),value is an object, so get attribute f as value
        else if (_.isPlainObject(value)) {
            var valueFunction = value.f;
            if (_.isString(valueFunction) &&
                valueFunction.slice(0, 1) === "=" &&
                valueFunction.length > 1) {
                var v = execfunction(ctx, valueFunction, r, c, undefined, undefined, true);
                isRunExecFunction = false;
                // value = {
                //     "v": v[1],
                //     "f": v[2]
                // };
                // update attribute v
                value.v = v[1], value.f = v[2];
                // 打进单元格的sparklines的配置串， 报错需要单独处理。
                if (v.length === 4 && v[3].type === "sparklines") {
                    var curCalv = v[3].data;
                    if (_.isArray(curCalv) && !_.isPlainObject(curCalv[0])) {
                        value.v = curCalv[0];
                    }
                    else {
                        value.spl = v[3].data;
                    }
                }
                else if (v.length === 4 && v[3].type === "dynamicArrayItem") {
                    // eslint-disable-next-line @typescript-eslint/no-unused-vars
                    dynamicArrayItem = v[3].data;
                }
            }
            else {
                var v = curv;
                if (_.isNil(value.v)) {
                    value.v = v;
                }
            }
        }
        else {
            delFunctionGroup(ctx, r, c);
            execFunctionGroup(ctx, r, c, value);
            // eslint-disable-next-line @typescript-eslint/no-unused-vars
            isRunExecFunction = false;
        }
    }
    // value maybe an object
    setCellValue(ctx, r, c, d, value);
    cancelNormalSelected(ctx);
    /*
    let RowlChange = false;
    const cfg =
      ctx.luckysheetfile?.[getSheetIndex(ctx, ctx.currentSheetId)]?.config ||
      {};
    if (!cfg.rowlen) {
      cfg.rowlen = {};
    }
    */
    if (((curv === null || curv === void 0 ? void 0 : curv.tb) === "2" && curv.v) || isInlineStringCell(d[r][c])) {
        // 自动换行
        var defaultrowlen = ctx.defaultrowlen;
        // const canvas = $("#luckysheetTableContent").get(0).getContext("2d");
        // offlinecanvas.textBaseline = 'top'; //textBaseline以top计算
        // let fontset = luckysheetfontformat(d[r][c]);
        // offlinecanvas.font = fontset;
        var cfg = ctx.luckysheetfile[getSheetIndex(ctx, ctx.currentSheetId)].config || {};
        if (!(((_f = cfg.columnlen) === null || _f === void 0 ? void 0 : _f[c]) && ((_g = cfg.rowlen) === null || _g === void 0 ? void 0 : _g[r]))) {
            // let currentRowLen = defaultrowlen;
            // if(!_.isNil(cfg["rowlen"][r])){
            //     currentRowLen = cfg["rowlen"][r];
            // }
            var cellWidth = ((_h = cfg.columnlen) === null || _h === void 0 ? void 0 : _h[c]) || ctx.defaultcollen;
            var textInfo = canvas
                ? getCellTextInfo(d[r][c], canvas, ctx, {
                    r: r,
                    c: c,
                    cellWidth: cellWidth,
                })
                : null;
            var currentRowLen = defaultrowlen;
            // console.log("rowlen", textInfo);
            if (textInfo) {
                currentRowLen = textInfo.textHeightAll + 2;
            }
            if (currentRowLen > defaultrowlen && !((_j = cfg.customHeight) === null || _j === void 0 ? void 0 : _j[r])) {
                if (_.isNil(cfg.rowlen))
                    cfg.rowlen = {};
                cfg.rowlen[r] = currentRowLen;
            }
        }
    }
    // 动态数组
    /*
    let dynamicArray = null;
    if (dynamicArrayItem) {
      // let file = ctx.luckysheetfile[getSheetIndex(ctx.currentSheetId)];
      dynamicArray = $.extend(
        true,
        [],
        this.insertUpdateDynamicArray(dynamicArrayItem)
      );
      // dynamicArray.push(dynamicArrayItem);
    }
  
    let allParam = {
      dynamicArray,
    };
  
    if (RowlChange) {
      allParam = {
        cfg,
        dynamicArray,
        RowlChange,
      };
    }
    */
    if (ctx.hooks.afterUpdateCell) {
        var newValue_1 = _.cloneDeep(flowdata[r][c]);
        var afterUpdateCell_1 = ctx.hooks.afterUpdateCell;
        setTimeout(function () {
            afterUpdateCell_1 === null || afterUpdateCell_1 === void 0 ? void 0 : afterUpdateCell_1(r, c, oldValue, newValue_1);
        });
    }
    ctx.formulaCache.execFunctionGlobalData = null;
}
export function getOrigincell(ctx, r, c, i) {
    var data = getFlowdata(ctx, i);
    if (_.isNil(r) || _.isNil(c)) {
        return null;
    }
    if (!data || !data[r] || !data[r][c]) {
        return null;
    }
    return data[r][c];
}
export function getcellFormula(ctx, r, c, i, data) {
    var cell;
    if (_.isNil(data)) {
        cell = getOrigincell(ctx, r, c, i);
    }
    else {
        cell = data[r][c];
    }
    if (_.isNil(cell)) {
        return null;
    }
    return cell.f;
}
export function getRange(ctx) {
    var rangeArr = _.cloneDeep(ctx.luckysheet_select_save);
    var result = [];
    if (!rangeArr)
        return result;
    for (var i = 0; i < rangeArr.length; i += 1) {
        var rangeItem = rangeArr[i];
        var temp = {
            row: rangeItem.row,
            column: rangeItem.column,
        };
        result.push(temp);
    }
    return result;
}
export function getFlattenedRange(ctx, range) {
    range = range || getRange(ctx);
    var result = [];
    range.forEach(function (ele) {
        // 这个data可能是个范围或者是单个cell
        var rs = ele.row;
        var cs = ele.column;
        for (var r = rs[0]; r <= rs[1]; r += 1) {
            for (var c = cs[0]; c <= cs[1]; c += 1) {
                // r c 当前的r和当前的c
                result.push({ r: r, c: c });
            }
        }
    });
    return result;
}
// 把选区范围数组转为string A1:A2
export function getRangetxt(ctx, sheetId, range, currentId) {
    var sheettxt = "";
    if (currentId == null) {
        currentId = ctx.currentSheetId;
    }
    if (sheetId !== currentId) {
        // sheet名字包含'的，引用时应该替换为''
        var index = getSheetIndex(ctx, sheetId);
        if (index == null)
            return "";
        sheettxt = ctx.luckysheetfile[index].name.replace(/'/g, "''");
        // 如果包含除a-z、A-Z、0-9、下划线等以外的字符那么就用单引号包起来
        if (
        // eslint-disable-next-line no-misleading-character-class
        /^[:A-Z_a-z\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u02FF\u0370-\u037D\u037F-\u1FFF\u200C-\u200D\u2070-\u218F\u2C00-\u2FEF\u3001-\uD7FF\uF900-\uFDCF\uFDF0-\uFFFD][:A-Z_a-z\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u02FF\u0370-\u037D\u037F-\u1FFF\u200C-\u200D\u2070-\u218F\u2C00-\u2FEF\u3001-\uD7FF\uF900-\uFDCF\uFDF0-\uFFFD\-.0-9\u00B7\u0300-\u036F\u203F-\u2040]*$/.test(sheettxt)) {
            sheettxt += "!";
        }
        else {
            sheettxt = "'".concat(sheettxt, "'!");
        }
    }
    var row0 = range.row[0];
    var row1 = range.row[1];
    var column0 = range.column[0];
    var column1 = range.column[1];
    if (row0 == null && row1 == null) {
        return "".concat(sheettxt + indexToColumnChar(column0), ":").concat(indexToColumnChar(column1));
    }
    if (column0 == null && column1 == null) {
        return "".concat(sheettxt + (row0 + 1), ":").concat(row1 + 1);
    }
    if (column0 === column1 && row0 === row1) {
        return sheettxt + indexToColumnChar(column0) + (row0 + 1);
    }
    return "".concat(sheettxt + indexToColumnChar(column0) + (row0 + 1), ":").concat(indexToColumnChar(column1)).concat(row1 + 1);
}
// 把string A1:A2转为选区数组
export function getRangeByTxt(ctx, txt) {
    var range = [];
    if (txt.indexOf(",") !== -1) {
        var arr = txt.split(",");
        for (var i = 0; i < arr.length; i += 1) {
            if (iscelldata(arr[i])) {
                range.push(getcellrange(ctx, arr[i]));
            }
            else {
                range = [];
                break;
            }
        }
    }
    else {
        if (iscelldata(txt)) {
            range.push(getcellrange(ctx, txt));
        }
    }
    return range;
}
export function isAllSelectedCellsInStatus(ctx, attr, status) {
    var _a, _b, _c, _d;
    // editing mode
    if (!_.isEmpty(ctx.luckysheetCellUpdate)) {
        var w = window.getSelection();
        if (!w)
            return false;
        if (w.rangeCount === 0)
            return false;
        var range = w.getRangeAt(0);
        if (range.collapsed === true) {
            return false;
        }
        var endContainer = range.endContainer;
        var startContainer = range.startContainer;
        // @ts-ignore
        var cssField_1 = _.camelCase(attrToCssName[attr]);
        if (startContainer === endContainer) {
            return !_.isEmpty(
            // @ts-ignore
            (_a = startContainer.parentElement) === null || _a === void 0 ? void 0 : _a.style[cssField_1]);
        }
        if (((_b = startContainer.parentElement) === null || _b === void 0 ? void 0 : _b.tagName) === "SPAN" &&
            ((_c = endContainer.parentElement) === null || _c === void 0 ? void 0 : _c.tagName) === "SPAN") {
            var startSpan = startContainer.parentNode;
            var endSpan = endContainer.parentNode;
            var allSpans = (_d = startSpan === null || startSpan === void 0 ? void 0 : startSpan.parentNode) === null || _d === void 0 ? void 0 : _d.querySelectorAll("span");
            if (allSpans) {
                var startSpanIndex = _.indexOf(allSpans, startSpan);
                var endSpanIndex = _.indexOf(allSpans, endSpan);
                var rangeSpans = [];
                for (var i = startSpanIndex; i <= endSpanIndex; i += 1) {
                    rangeSpans.push(allSpans[i]);
                }
                // @ts-ignore
                return _.every(rangeSpans, function (s) { return !_.isEmpty(s.style[cssField_1]); });
            }
        }
    }
    /* 获取选区内所有的单元格-扁平后的处理 */
    var cells = getFlattenedRange(ctx);
    var flowdata = getFlowdata(ctx);
    return cells.every(function (_a) {
        var _b;
        var r = _a.r, c = _a.c;
        var cell = (_b = flowdata === null || flowdata === void 0 ? void 0 : flowdata[r]) === null || _b === void 0 ? void 0 : _b[c];
        if (_.isNil(cell)) {
            return false;
        }
        return cell[attr] === status;
    });
}
export function getFontStyleByCell(cell, checksAF, checksCF, isCheck) {
    if (isCheck === void 0) { isCheck = true; }
    var style = {};
    if (!cell) {
        return style;
    }
    // @ts-ignore
    _.forEach(cell, function (v, key) {
        var _a, _b, _c, _d;
        var value = cell[key];
        if (isCheck) {
            value = normalizedCellAttr(cell, key);
        }
        var valueNum = Number(value);
        if (key === "bl" && valueNum !== 0) {
            style.fontWeight = "bold";
        }
        if (key === "it" && valueNum !== 0) {
            style.fontStyle = "italic";
        }
        // if (key === "ff") {
        //   let f = value;
        //   if (!Number.isNaN(valueNum)) {
        //     f = locale_fontarray[parseInt(value)];
        //   } else {
        //     f = value;
        //   }
        //   style += "font-family: " + f + ";";
        // }
        if (key === "fs" && valueNum !== 10) {
            style.fontSize = "".concat(valueNum, "pt");
        }
        if ((key === "fc" && value !== "#000000") ||
            ((_a = checksAF === null || checksAF === void 0 ? void 0 : checksAF.length) !== null && _a !== void 0 ? _a : 0) > 0 ||
            (checksCF === null || checksCF === void 0 ? void 0 : checksCF.textColor)) {
            if (checksCF === null || checksCF === void 0 ? void 0 : checksCF.textColor) {
                style.color = checksCF.textColor;
            }
            else if (((_b = checksAF === null || checksAF === void 0 ? void 0 : checksAF.length) !== null && _b !== void 0 ? _b : 0) > 0) {
                style.color = checksAF[0];
            }
            else {
                style.color = value;
            }
        }
        if (key === "cl" && valueNum !== 0) {
            style.textDecoration = "line-through";
        }
        if (key === "un" && (valueNum === 1 || valueNum === 3)) {
            // @ts-ignore
            var color = (_c = cell._color) !== null && _c !== void 0 ? _c : cell.fc;
            // @ts-ignore
            var fs = (_d = cell._fontSize) !== null && _d !== void 0 ? _d : cell.fs;
            style.borderBottom = "".concat(Math.floor(fs / 9), "px solid ").concat(color);
        }
    });
    return style;
}
export function getStyleByCell(ctx, d, r, c) {
    var _a;
    var style = {};
    // 交替颜色
    //   const af_compute = alternateformat.getComputeMap();
    //   const checksAF = alternateformat.checksAF(r, c, af_compute);
    var checksAF = [];
    // 条件格式
    var cf_compute = getComputeMap(ctx);
    var checksCF = checkCF(r, c, cf_compute);
    var cell = (_a = d === null || d === void 0 ? void 0 : d[r]) === null || _a === void 0 ? void 0 : _a[c];
    if (!cell)
        return {};
    var isInline = isInlineStringCell(cell);
    if ("bg" in cell) {
        var value = normalizedCellAttr(cell, "bg");
        if (checksCF === null || checksCF === void 0 ? void 0 : checksCF.cellColor) {
            if (checksCF === null || checksCF === void 0 ? void 0 : checksCF.cellColor) {
                style.background = "".concat(checksCF.cellColor);
            }
            else if (checksAF.length > 1) {
                style.background = "".concat(checksAF[1]);
            }
            else {
                style.background = "".concat(value);
            }
        }
    }
    if ("ht" in cell) {
        var value = normalizedCellAttr(cell, "ht");
        if (Number(value) === 0) {
            style.textAlign = "center";
        }
        else if (Number(value) === 2) {
            style.textAlign = "right";
        }
    }
    if ("vt" in cell) {
        var value = normalizedCellAttr(cell, "vt");
        if (Number(value) === 0) {
            style.alignItems = "center";
        }
        else if (Number(value) === 2) {
            style.alignItems = "flex-end";
        }
    }
    if (!isInline) {
        style = _.assign(style, getFontStyleByCell(cell, checksAF, checksCF));
    }
    return style;
}
export function getInlineStringHTML(r, c, data) {
    var ct = getCellValue(r, c, data, "ct");
    if (isInlineStringCT(ct)) {
        var strings = ct.s;
        var value = "";
        for (var i = 0; i < strings.length; i += 1) {
            var strObj = strings[i];
            if (strObj.v) {
                var style = getFontStyleByCell(strObj);
                var styleStr = _.map(style, function (v, key) {
                    return "".concat(_.kebabCase(key), ":").concat(_.isNumber(v) ? "".concat(v, "px") : v, ";");
                }).join("");
                value += "<span class=\"luckysheet-input-span\" index='".concat(i, "' style='").concat(styleStr, "'>").concat(strObj.v, "</span>");
            }
        }
        return value;
    }
    return "";
}
export function getQKBorder(width, type, color) {
    var bordertype = "";
    if (width.toString().indexOf("pt") > -1) {
        var nWidth = parseFloat(width);
        if (nWidth < 1) {
        }
        else if (nWidth < 1.5) {
            bordertype = "Medium";
        }
        else {
            bordertype = "Thick";
        }
    }
    else {
        var nWidth = parseFloat(width);
        if (nWidth < 2) {
        }
        else if (nWidth < 3) {
            bordertype = "Medium";
        }
        else {
            bordertype = "Thick";
        }
    }
    var style = 0;
    type = type.toLowerCase();
    if (type === "double") {
        style = 2;
    }
    else if (type === "dotted") {
        if (bordertype === "Medium" || bordertype === "Thick") {
            style = 3;
        }
        else {
            style = 10;
        }
    }
    else if (type === "dashed") {
        if (bordertype === "Medium" || bordertype === "Thick") {
            style = 4;
        }
        else {
            style = 9;
        }
    }
    else if (type === "solid") {
        if (bordertype === "Medium") {
            style = 8;
        }
        else if (bordertype === "Thick") {
            style = 13;
        }
        else {
            style = 1;
        }
    }
    return [style, color];
}
/**
 * 计算范围行高
 *
 * @param d 原始数据
 * @param r1 起始行
 * @param r2 截至行
 * @param cfg 配置
 * @returns 计算后的配置
 */
/*
export function rowlenByRange(
  ctx: Context,
  d: CellMatrix,
  r1: number,
  r2: number,
  cfg: any
) {
  const cfg_clone = _.cloneDeep(cfg);
  if (cfg_clone.rowlen == null) {
    cfg_clone.rowlen = {};
  }

  if (cfg_clone.customHeight == null) {
    cfg_clone.customHeight = {};
  }

  const canvas = $("#luckysheetTableContent").get(0).getContext("2d");
  canvas.textBaseline = "top"; // textBaseline以top计算

  for (let r = r1; r <= r2; r += 1) {
    if (cfg_clone.rowhidden != null && cfg_clone.rowhidden[r] != null) {
      continue;
    }

    let currentRowLen = ctx.defaultrowlen;

    if (cfg_clone.customHeight[r] === 1) {
      continue;
    }

    delete cfg_clone.rowlen[r];

    for (let c = 0; c < d[r].length; c += 1) {
      const cell = d[r][c];

      if (cell == null) {
        continue;
      }

      if (cell != null && (cell.v != null || isInlineStringCell(cell))) {
        let cellWidth;
        if (cell.mc) {
          if (c === cell.mc.c) {
            const st_cellWidth = colLocationByIndex(
              c,
              ctx.visibledatacolumn
            )[0];
            const ed_cellWidth = colLocationByIndex(
              cell.mc.c + cell.mc.cs - 1,
              ctx.visibledatacolumn
            )[1];
            cellWidth = ed_cellWidth - st_cellWidth - 2;
          } else {
            continue;
          }
        } else {
          cellWidth =
            colLocationByIndex(c, ctx.visibledatacolumn)[1] -
            colLocationByIndex(c, ctx.visibledatacolumn)[0] -
            2;
        }

       const textInfo = getCellTextInfo(cell, canvas, {
          r,
          c,
          cellWidth,
        });

        let computeRowlen = 0;

        if (textInfo != null) {
          computeRowlen = textInfo.textHeightAll + 2;
        }

        // 比较计算高度和当前高度取最大高度
        if (computeRowlen > currentRowLen) {
          currentRowLen = computeRowlen;
        }
      }
    }

    currentRowLen /= ctx.zoomRatio;

    if (currentRowLen !== ctx.defaultrowlen) {
      cfg_clone.rowlen[r] = currentRowLen;
    } else {
      if (cfg.rowlen?.[r]) {
        cfg_clone.rowlen[r] = cfg.rowlen[r];
      }
    }
  }

  return cfg_clone;
}
*/
export function getdatabyselection(ctx, range, sheetId) {
    if (range == null && ctx.luckysheet_select_save) {
        range = ctx.luckysheet_select_save[0];
    }
    if (!range)
        return [];
    if (range.row == null || range.row.length === 0) {
        return [];
    }
    // 取数据
    var d;
    var cfg;
    if (sheetId != null && sheetId !== ctx.currentSheetId) {
        d = ctx.luckysheetfile[getSheetIndex(ctx, sheetId)].data;
        cfg = ctx.luckysheetfile[getSheetIndex(ctx, sheetId)].config;
    }
    else {
        d = getFlowdata(ctx);
        cfg = ctx.config;
    }
    var data = [];
    for (var r = range.row[0]; r <= range.row[1]; r += 1) {
        if ((d === null || d === void 0 ? void 0 : d[r]) == null) {
            continue;
        }
        if ((cfg === null || cfg === void 0 ? void 0 : cfg.rowhidden) != null && cfg.rowhidden[r] != null) {
            continue;
        }
        var row = [];
        for (var c = range.column[0]; c <= range.column[1]; c += 1) {
            if ((cfg === null || cfg === void 0 ? void 0 : cfg.colhidden) != null && cfg.colhidden[c] != null) {
                continue;
            }
            row.push(d[r][c]);
        }
        data.push(row);
    }
    return data;
}
export function luckysheetUpdateCell(ctx, row_index, col_index) {
    ctx.luckysheetCellUpdate = [row_index, col_index];
}
export function getDataBySelectionNoCopy(ctx, range) {
    if (!range || !range.row || range.row.length === 0)
        return [];
    var data = [];
    var flowData = getFlowdata(ctx);
    if (!flowData)
        return [];
    for (var r = range.row[0]; r <= range.row[1]; r += 1) {
        var row = [];
        if (ctx.config.rowhidden != null && ctx.config.rowhidden[r] != null) {
            continue;
        }
        for (var c = range.column[0]; c <= range.column[1]; c += 1) {
            var value = null;
            if (ctx.config.colhidden != null && ctx.config.colhidden[c] != null) {
                continue;
            }
            if (flowData[r] != null && flowData[r][c] != null) {
                value = flowData[r][c];
            }
            row.push(value);
        }
        data.push(row);
    }
    return data;
}
