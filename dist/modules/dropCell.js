import _ from "lodash";
import dayjs from "dayjs";
import { getFlowdata } from "../context";
import { colLocation, rowLocation } from "./location";
import { getSheetIndex, isAllowEdit } from "../utils";
import { getBorderInfoCompute } from "./border";
import { genarate, update } from "./format";
import * as formula from "./formula";
import { isRealNum } from "./validation";
import { CFSplitRange } from "./ConditionFormat";
import { normalizeSelection } from "./selection";
import { jfrefreshgrid } from "./refresh";
function toPx(v) {
    return "".concat(v, "px");
}
export var dropCellCache = {
    copyRange: {},
    applyRange: {},
    applyType: null,
    direction: null,
    chnNumChar: {
        零: 0,
        一: 1,
        二: 2,
        三: 3,
        四: 4,
        五: 5,
        六: 6,
        七: 7,
        八: 8,
        九: 9,
    },
    chnNameValue: {
        十: { value: 10, secUnit: false },
        百: { value: 100, secUnit: false },
        千: { value: 1000, secUnit: false },
        万: { value: 10000, secUnit: true },
        亿: { value: 100000000, secUnit: true },
    },
    chnNumChar2: ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"],
    chnUnitSection: ["", "万", "亿", "万亿", "亿亿"],
    chnUnitChar: ["", "十", "百", "千"],
};
function chineseToNumber(chnStr) {
    var rtn = 0;
    var section = 0;
    var number = 0;
    var secUnit = false;
    var str = chnStr.split("");
    for (var i = 0; i < str.length; i += 1) {
        var num = dropCellCache.chnNumChar[str[i]];
        if (typeof num !== "undefined") {
            number = num;
            if (i === str.length - 1) {
                section += number;
            }
        }
        else {
            var unit = dropCellCache.chnNameValue[str[i]].value;
            secUnit = dropCellCache.chnNameValue[str[i]].secUnit;
            if (secUnit) {
                section = (section + number) * unit;
                rtn += section;
                section = 0;
            }
            else {
                section += number * unit;
            }
            number = 0;
        }
    }
    return rtn + section;
}
function sectionToChinese(section) {
    var strIns = "";
    var chnStr = "";
    var unitPos = 0;
    var zero = true;
    while (section > 0) {
        var v = section % 10;
        if (v === 0) {
            if (!zero) {
                zero = true;
                chnStr = dropCellCache.chnNumChar2[v] + chnStr;
            }
        }
        else {
            zero = false;
            strIns = dropCellCache.chnNumChar2[v];
            strIns += dropCellCache.chnUnitChar[unitPos];
            chnStr = strIns + chnStr;
        }
        unitPos += 1;
        section = Math.floor(section / 10);
    }
    return chnStr;
}
function numberToChinese(num) {
    var strIns = "";
    var chnStr = "";
    var unitPos = 0;
    var needZero = false;
    if (num === 0) {
        return dropCellCache.chnNumChar2[0];
    }
    while (num > 0) {
        var section = num % 10000;
        if (needZero) {
            chnStr = dropCellCache.chnNumChar2[0] + chnStr;
        }
        strIns = sectionToChinese(section);
        strIns +=
            section !== 0
                ? dropCellCache.chnUnitSection[unitPos]
                : dropCellCache.chnUnitSection[0];
        chnStr = strIns + chnStr;
        needZero = section < 1000 && section > 0;
        num = Math.floor(num / 10000);
        unitPos += 1;
    }
    return chnStr;
}
function isChnNumber(txt) {
    if (typeof txt === "number") {
        txt = "".concat(txt);
    }
    var result = true;
    if (txt == null) {
        result = false;
    }
    else if (txt.length === 1) {
        if (txt === "日" || txt in dropCellCache.chnNumChar) {
            result = true;
        }
        else {
            result = false;
        }
    }
    else {
        var str = txt.split("");
        for (var i = 0; i < str.length; i += 1) {
            if (!(str[i] in dropCellCache.chnNumChar ||
                str[i] in dropCellCache.chnNameValue)) {
                result = false;
                break;
            }
        }
    }
    return result;
}
function isExtendNumber(txt) {
    if (txt == null)
        return [false];
    if (typeof txt === "number") {
        txt = "".concat(txt);
    }
    var reg = /0|([1-9]+[0-9]*)/g;
    var result = reg.test(txt);
    if (result) {
        var match = txt.match(reg);
        if (match) {
            var matchTxt = match[match.length - 1];
            var matchIndex = txt.lastIndexOf(matchTxt);
            var beforeTxt = txt.slice(0, matchIndex);
            var afterTxt = txt.slice(matchIndex + matchTxt.length);
            return [result, Number(matchTxt), beforeTxt, afterTxt];
        }
    }
    return [result];
}
// function isChnWeek1(txt: string | number) {
//   if (typeof txt === "number") {
//     txt = `${txt}`;
//   }
//   let result = false;
//   if (txt.length === 1 && (txt === "日" || chineseToNumber(txt) < 7)) {
//     result = true;
//   }
//   return result;
// }
function isChnWeek2(txt) {
    var result = false;
    if (typeof txt === "number") {
        txt = "".concat(txt);
    }
    if (txt !== undefined && txt.length === 2) {
        if (txt === "周一" ||
            txt === "周二" ||
            txt === "周三" ||
            txt === "周四" ||
            txt === "周五" ||
            txt === "周六" ||
            txt === "周日") {
            result = true;
        }
    }
    return result;
}
function isChnWeek3(txt) {
    if (typeof txt === "number") {
        txt = "".concat(txt);
    }
    var result = false;
    if (txt !== undefined && txt.length === 3) {
        if (txt === "星期一" ||
            txt === "星期二" ||
            txt === "星期三" ||
            txt === "星期四" ||
            txt === "星期五" ||
            txt === "星期六" ||
            txt === "星期日") {
            result = true;
        }
    }
    return result;
}
function isEqualDiff(arr) {
    var diff = true;
    var step = arr[1] - arr[0];
    for (var i = 1; i < arr.length; i += 1) {
        if (arr[i] - arr[i - 1] !== step) {
            diff = false;
            break;
        }
    }
    return diff;
}
function isEqualRatio(arr) {
    var ratio = true;
    var step = arr[1] / arr[0];
    for (var i = 1; i < arr.length; i += 1) {
        if (arr[i] / arr[i - 1] !== step) {
            ratio = false;
            break;
        }
    }
    return ratio;
}
function getXArr(len) {
    var xArr = [];
    for (var i = 1; i <= len; i += 1) {
        xArr.push(i);
    }
    return xArr;
}
function forecast(x, yArr, xArr) {
    function getAverage(arr) {
        var sum = 0;
        for (var i = 0; i < arr.length; i += 1) {
            sum += arr[i];
        }
        return sum / arr.length;
    }
    var ax = getAverage(xArr); // x数组 平均值
    var ay = getAverage(yArr); // y数组 平均值
    var sum_d = 0;
    var sum_n = 0;
    for (var j = 0; j < xArr.length; j += 1) {
        // 分母和
        sum_d += (xArr[j] - ax) * (yArr[j] - ay);
        // 分子和
        sum_n += (xArr[j] - ax) * (xArr[j] - ax);
    }
    var b;
    if (sum_n === 0) {
        b = 1;
    }
    else {
        b = sum_d / sum_n;
    }
    var a = ay - b * ax;
    return Math.round((a + b * x) * 100000) / 100000;
}
function judgeDate(data) {
    var _a, _b, _c, _d, _e, _f, _g, _h;
    var isSameDay = true;
    var isSameMonth = true;
    var isEqualDiffDays = true;
    var isEqualDiffMonths = true;
    var isEqualDiffYears = true;
    if (data[0] == null || data[1] == null)
        return [false, false, false, false, false];
    var sameDay = dayjs(data[0].m).date();
    var sameMonth = dayjs(data[0].m).month();
    var equalDiffDays = dayjs(data[1].m).diff(dayjs(data[0].m), "days");
    var equalDiffMonths = dayjs(data[1].m).diff(dayjs(data[0].m), "months");
    var equalDiffYears = dayjs(data[1].m).diff(dayjs(data[0].m), "years");
    for (var i = 1; i < data.length; i += 1) {
        // 日是否一样
        if (dayjs((_a = data[i]) === null || _a === void 0 ? void 0 : _a.m).date() !== sameDay) {
            isSameDay = false;
        }
        // 月是否一样
        if (dayjs((_b = data[i]) === null || _b === void 0 ? void 0 : _b.m).month() !== sameMonth) {
            isSameMonth = false;
        }
        // 日差是否是 等差数列
        if (dayjs((_c = data[i]) === null || _c === void 0 ? void 0 : _c.m).diff(dayjs((_d = data[i - 1]) === null || _d === void 0 ? void 0 : _d.m), "days") !== equalDiffDays) {
            isEqualDiffDays = false;
        }
        // 月差是否是 等差数列
        if (dayjs((_e = data[i]) === null || _e === void 0 ? void 0 : _e.m).diff(dayjs((_f = data[i - 1]) === null || _f === void 0 ? void 0 : _f.m), "months") !==
            equalDiffMonths) {
            isEqualDiffMonths = false;
        }
        // 年差是否是 等差数列
        if (dayjs((_g = data[i]) === null || _g === void 0 ? void 0 : _g.m).diff(dayjs((_h = data[i - 1]) === null || _h === void 0 ? void 0 : _h.m), "years") !== equalDiffYears) {
            isEqualDiffYears = false;
        }
    }
    if (equalDiffDays === 0) {
        isEqualDiffDays = false;
    }
    if (equalDiffMonths === 0) {
        isEqualDiffMonths = false;
    }
    if (equalDiffYears === 0) {
        isEqualDiffYears = false;
    }
    return [
        isSameDay,
        isSameMonth,
        isEqualDiffDays,
        isEqualDiffMonths,
        isEqualDiffYears,
    ];
}
export function showDropCellSelection(_a, container) {
    var width = _a.width, height = _a.height, top = _a.top, left = _a.left;
    var selectedExtend = container.querySelector(".fortune-cell-selected-extend");
    if (selectedExtend) {
        selectedExtend.style.left = toPx(left);
        selectedExtend.style.width = toPx(width);
        selectedExtend.style.top = toPx(top);
        selectedExtend.style.height = toPx(height);
        selectedExtend.style.display = "block";
    }
}
export function hideDropCellSelection(container) {
    var selectedExtend = container.querySelector(".fortune-cell-selected-extend");
    if (selectedExtend) {
        selectedExtend.style.display = "none";
    }
}
export function createDropCellRange(ctx, e, container) {
    ctx.luckysheet_cell_selected_extend = true;
    ctx.luckysheet_scroll_status = true;
    var scrollLeft = ctx.scrollLeft, scrollTop = ctx.scrollTop;
    var rect = container.getBoundingClientRect();
    var x = e.pageX - rect.left - ctx.rowHeaderWidth + scrollLeft;
    var y = e.pageY - rect.top - ctx.columnHeaderHeight + scrollTop;
    var row_location = rowLocation(y, ctx.visibledatarow);
    var row_pre = row_location[0];
    var row = row_location[1];
    var row_index = row_location[2];
    var col_location = colLocation(x, ctx.visibledatacolumn);
    var col_pre = col_location[0];
    var col = col_location[1];
    var col_index = col_location[2];
    ctx.luckysheet_cell_selected_extend_index = [row_index, col_index];
    showDropCellSelection({
        left: col_pre,
        width: col - col_pre - 1,
        top: row_pre,
        height: row - row_pre - 1,
    }, container);
}
export function onDropCellSelect(ctx, e, scrollX, scrollY, container) {
    var scrollLeft = scrollX.scrollLeft;
    var scrollTop = scrollY.scrollTop;
    var rect = container.getBoundingClientRect();
    var x = e.pageX - rect.left - ctx.rowHeaderWidth + scrollLeft;
    var y = e.pageY - rect.top - ctx.columnHeaderHeight + scrollTop;
    var row_location = rowLocation(y, ctx.visibledatarow);
    var row = row_location[1];
    var row_pre = row_location[0];
    var row_index = row_location[2];
    var col_location = colLocation(x, ctx.visibledatacolumn);
    var col = col_location[1];
    var col_pre = col_location[0];
    var col_index = col_location[2];
    var row_index_original = ctx.luckysheet_cell_selected_extend_index[0];
    var col_index_original = ctx.luckysheet_cell_selected_extend_index[1];
    if (!ctx.luckysheet_select_save)
        return;
    var row_s = ctx.luckysheet_select_save[0].row[0];
    var row_e = ctx.luckysheet_select_save[0].row[1];
    var col_s = ctx.luckysheet_select_save[0].column[0];
    var col_e = ctx.luckysheet_select_save[0].column[1];
    var top = ctx.luckysheet_select_save[0].top_move;
    var height = ctx.luckysheet_select_save[0].height_move;
    var left = ctx.luckysheet_select_save[0].left_move;
    var width = ctx.luckysheet_select_save[0].width_move;
    if (top == null || height == null || left == null || width == null)
        return;
    if (Math.abs(row_index_original - row_index) >
        Math.abs(col_index_original - col_index)) {
        if (!(row_index >= row_s && row_index <= row_e)) {
            if (top >= row_pre) {
                height += top - row_pre;
                top = row_pre;
            }
            else {
                height = row - top - 1;
            }
        }
    }
    else {
        if (!(col_index >= col_s && col_index <= col_e)) {
            if (left >= col_pre) {
                width += left - col_pre;
                left = col_pre;
            }
            else {
                width = col - left - 1;
            }
        }
    }
    if (y < 0) {
        row_s = 0;
        row_e = ctx.luckysheet_select_save[0].row[0];
    }
    if (x < 0) {
        col_s = 0;
        col_e = ctx.luckysheet_select_save[0].column[0];
    }
    showDropCellSelection({ left: left, width: width, top: top, height: height }, container);
}
function fillCopy(data, len) {
    var applyData = [];
    for (var i = 1; i <= len; i += 1) {
        var index = (i - 1) % data.length;
        var d = _.cloneDeep(data[index]);
        if (!_.isUndefined(d)) {
            applyData.push(d);
        }
    }
    return applyData;
}
function fillSeries(data, len, direction) {
    var applyData = [];
    var dataNumArr = [];
    for (var j = 0; j < data.length; j += 1) {
        var d = _.cloneDeep(data[j]);
        if (d != null) {
            dataNumArr.push(Number(d.v));
        }
    }
    if (data.length > 2 &&
        isEqualRatio(dataNumArr) &&
        data[0] != null &&
        data[1] != null) {
        // 等比数列
        for (var i = 1; i <= len; i += 1) {
            var index = (i - 1) % data.length;
            var d = _.cloneDeep(data[index]);
            if (d != null) {
                var num = void 0;
                if (direction === "down" || direction === "right") {
                    num =
                        Number(data[data.length - 1].v) *
                            Math.pow((Number(data[1].v) / Number(data[0].v)), i);
                }
                else {
                    //  direction == "up" || direction == "left"
                    num =
                        Number(data[0].v) / Math.pow((Number(data[1].v) / Number(data[0].v)), i);
                }
                d.v = num;
                if (d.ct != null && d.ct.fa != null) {
                    d.m = update(d.ct.fa, num);
                }
                applyData.push(d);
            }
        }
    }
    else {
        // 线性数列
        var xArr = getXArr(data.length);
        for (var i = 1; i <= len; i += 1) {
            var index = (i - 1) % data.length;
            var d = _.cloneDeep(data[index]);
            if (d != null) {
                var y = void 0;
                if (direction === "down" || direction === "right") {
                    y = forecast(data.length + i, dataNumArr, xArr);
                }
                else if (direction === "up" || direction === "left") {
                    y = forecast(1 - i, dataNumArr, xArr);
                }
                d.v = y;
                if (d.ct != null && d.ct.fa != null) {
                    d.m = update(d.ct.fa, y);
                }
                applyData.push(d);
            }
        }
    }
    return applyData;
}
function fillExtendNumber(data, len, step) {
    var _a;
    var applyData = [];
    var reg = /0|([1-9]+[0-9]*)/g;
    for (var i = 1; i <= len; i += 1) {
        var index = (i - 1) % data.length;
        var d = _.cloneDeep(data[index]);
        var last = (_a = data[data.length - 1]) === null || _a === void 0 ? void 0 : _a.m;
        if (d != null && last != null) {
            last = "".concat(last);
            var match = last.match(reg) || "";
            var lastTxt = match[match.length - 1];
            var num = Math.abs(Number(lastTxt) + step * i);
            var lastIndex = last.lastIndexOf(lastTxt);
            var valueTxt = last.slice(0, lastIndex) +
                num.toString() +
                last.slice(lastIndex + lastTxt.length);
            d.v = valueTxt;
            d.m = valueTxt;
            applyData.push(d);
        }
    }
    return applyData;
}
function fillDays(data, len, step) {
    var _a;
    var applyData = [];
    for (var i = 1; i <= len; i += 1) {
        var d = _.cloneDeep(data[data.length - 1]);
        if (d != null) {
            var date = update("yyyy-MM-dd", d.v);
            date = dayjs(date)
                .add(step * i, "days")
                .format("YYYY-MM-DD");
            // TODO generate的处理是否合适
            d.v = (_a = genarate(date)) === null || _a === void 0 ? void 0 : _a[2];
            if (d.ct != null && d.ct.fa != null) {
                d.m = update(d.ct.fa, d.v);
            }
            applyData.push(d);
        }
    }
    return applyData;
}
function fillMonths(data, len, step) {
    var _a;
    var applyData = [];
    for (var i = 1; i <= len; i += 1) {
        var d = _.cloneDeep(data[data.length - 1]);
        if (d != null) {
            var date = update("yyyy-MM-dd", d.v);
            date = dayjs(date)
                .add(step * i, "months")
                .format("YYYY-MM-DD");
            d.v = (_a = genarate(date)) === null || _a === void 0 ? void 0 : _a[2];
            if (d.ct != null && d.ct.fa != null) {
                d.m = update(d.ct.fa, d.v);
            }
            applyData.push(d);
        }
    }
    return applyData;
}
function fillYears(data, len, step) {
    var _a;
    var applyData = [];
    for (var i = 1; i <= len; i += 1) {
        var d = _.cloneDeep(data[data.length - 1]);
        if (d != null) {
            var date = update("yyyy-MM-dd", d.v);
            date = dayjs(date)
                .add(step * i, "years")
                .format("YYYY-MM-DD");
            d.v = (_a = genarate(date)) === null || _a === void 0 ? void 0 : _a[2];
            if (d.ct != null && d.ct.fa != null) {
                d.m = update(d.ct.fa, d.v);
            }
        }
        applyData.push(d);
    }
    return applyData;
}
function fillChnWeek(data, len, step) {
    var _a;
    var applyData = [];
    for (var i = 1; i <= len; i += 1) {
        var index = (i - 1) % data.length;
        var d = _.cloneDeep(data[index]);
        var num = void 0;
        var m = (_a = data[data.length - 1]) === null || _a === void 0 ? void 0 : _a.m;
        if (m != null && d != null) {
            if (m === "日") {
                num = 7 + step * i;
            }
            else {
                num = chineseToNumber("".concat(m)) + step * i;
            }
            if (num < 0) {
                num = Math.ceil(Math.abs(num) / 7) * 7 + num;
            }
            var rsd = num % 7;
            if (rsd === 0) {
                d.m = "日";
                d.v = "日";
            }
            else if (rsd === 1) {
                d.m = "一";
                d.v = "一";
            }
            else if (rsd === 2) {
                d.m = "二";
                d.v = "二";
            }
            else if (rsd === 3) {
                d.m = "三";
                d.v = "三";
            }
            else if (rsd === 4) {
                d.m = "四";
                d.v = "四";
            }
            else if (rsd === 5) {
                d.m = "五";
                d.v = "五";
            }
            else if (rsd === 6) {
                d.m = "六";
                d.v = "六";
            }
            applyData.push(d);
        }
    }
    return applyData;
}
function fillChnWeek2(data, len, step) {
    var _a;
    var applyData = [];
    for (var i = 1; i <= len; i += 1) {
        var index = (i - 1) % data.length;
        var d = _.cloneDeep(data[index]);
        var num = void 0;
        var m = (_a = data[data.length - 1]) === null || _a === void 0 ? void 0 : _a.m;
        if (m != null && d != null) {
            if (m === "周日") {
                num = 7 + step * i;
            }
            else {
                var last = "".concat(m);
                var txt = last.slice(last.length - 1, 1);
                num = chineseToNumber(txt) + step * i;
            }
            if (num < 0) {
                num = Math.ceil(Math.abs(num) / 7) * 7 + num;
            }
            var rsd = num % 7;
            if (rsd === 0) {
                d.m = "周日";
                d.v = "周日";
            }
            else if (rsd === 1) {
                d.m = "周一";
                d.v = "周一";
            }
            else if (rsd === 2) {
                d.m = "周二";
                d.v = "周二";
            }
            else if (rsd === 3) {
                d.m = "周三";
                d.v = "周三";
            }
            else if (rsd === 4) {
                d.m = "周四";
                d.v = "周四";
            }
            else if (rsd === 5) {
                d.m = "周五";
                d.v = "周五";
            }
            else if (rsd === 6) {
                d.m = "周六";
                d.v = "周六";
            }
        }
        applyData.push(d);
    }
    return applyData;
}
function fillChnWeek3(data, len, step) {
    var _a;
    var applyData = [];
    for (var i = 1; i <= len; i += 1) {
        var index = (i - 1) % data.length;
        var d = _.cloneDeep(data[index]);
        var num = void 0;
        var m = (_a = data[data.length - 1]) === null || _a === void 0 ? void 0 : _a.m;
        if (m != null && d != null) {
            if (m === "星期日") {
                num = 7 + step * i;
            }
            else {
                var last = "".concat(m);
                var txt = last.slice(last.length - 1, 1);
                num = chineseToNumber(txt) + step * i;
            }
            if (num < 0) {
                num = Math.ceil(Math.abs(num) / 7) * 7 + num;
            }
            var rsd = num % 7;
            if (rsd === 0) {
                d.m = "星期日";
                d.v = "星期日";
            }
            else if (rsd === 1) {
                d.m = "星期一";
                d.v = "星期一";
            }
            else if (rsd === 2) {
                d.m = "星期二";
                d.v = "星期二";
            }
            else if (rsd === 3) {
                d.m = "星期三";
                d.v = "星期三";
            }
            else if (rsd === 4) {
                d.m = "星期四";
                d.v = "星期四";
            }
            else if (rsd === 5) {
                d.m = "星期五";
                d.v = "星期五";
            }
            else if (rsd === 6) {
                d.m = "星期六";
                d.v = "星期六";
            }
        }
        applyData.push(d);
    }
    return applyData;
}
function fillChnNumber(data, len, step) {
    var _a;
    var applyData = [];
    for (var i = 1; i <= len; i += 1) {
        var index = (i - 1) % data.length;
        var d = _.cloneDeep(data[index]);
        var m = (_a = data[data.length - 1]) === null || _a === void 0 ? void 0 : _a.m;
        if (m != null && d != null) {
            var num = chineseToNumber("".concat(m)) + step * i;
            var txt = void 0;
            if (num <= 0) {
                txt = "零";
            }
            else {
                txt = numberToChinese(num);
            }
            d.v = txt;
            d.m = txt.toString();
            applyData.push(d);
        }
    }
    return applyData;
}
export function getTypeItemHide(ctx) {
    var copyRange = dropCellCache.copyRange;
    var str_r = copyRange.row[0];
    var end_r = copyRange.row[1];
    var str_c = copyRange.column[0];
    var end_c = copyRange.column[1];
    var hasNumber = false;
    var hasExtendNumber = false;
    var hasDate = false;
    var hasChn = false;
    var hasChnWeek1 = false;
    var hasChnWeek2 = false;
    var hasChnWeek3 = false;
    var flowdata = getFlowdata(ctx);
    if (flowdata == null)
        return [];
    for (var r = str_r; r <= end_r; r += 1) {
        for (var c = str_c; c <= end_c; c += 1) {
            if (flowdata[r][c]) {
                var cell = flowdata[r][c];
                if (cell !== null && cell.v != null && cell.f == null) {
                    if (cell.ct != null && cell.ct.t === "n") {
                        hasNumber = true;
                    }
                    else if (cell.ct != null && cell.ct.t === "d") {
                        hasDate = true;
                    }
                    else if (isExtendNumber(cell.m)[0]) {
                        hasExtendNumber = true;
                    }
                    else if (isChnNumber(cell.m) && cell.m !== "日") {
                        hasChn = true;
                    }
                    else if (cell.m != null && cell.m === "日") {
                        hasChnWeek1 = true;
                    }
                    else if (isChnWeek2(cell.m)) {
                        hasChnWeek2 = true;
                    }
                    else if (isChnWeek3(cell.m)) {
                        hasChnWeek3 = true;
                    }
                }
            }
        }
    }
    return [
        hasNumber,
        hasExtendNumber,
        hasDate,
        hasChn,
        hasChnWeek1,
        hasChnWeek2,
        hasChnWeek3,
    ];
}
function getLenS(indexArr, rsd) {
    var s = 0;
    for (var j = 0; j < indexArr.length; j += 1) {
        if (indexArr[j] <= rsd) {
            s += 1;
        }
        else {
            break;
        }
    }
    return s;
}
function getDataIndex(csLen, asLen, indexArr) {
    var obj = {};
    var num = Math.floor(asLen / csLen);
    var rsd = asLen % csLen;
    var sum = 0;
    if (num > 0) {
        for (var i = 1; i <= num; i += 1) {
            for (var j = 0; j < indexArr.length; j += 1) {
                obj[indexArr[j] + (i - 1) * csLen] = sum;
                sum += 1;
            }
        }
        for (var a = 0; a < indexArr.length; a += 1) {
            if (indexArr[a] <= rsd) {
                obj[indexArr[a] + csLen * num] = sum;
                sum += 1;
            }
            else {
                break;
            }
        }
    }
    else {
        for (var a = 0; a < indexArr.length; a += 1) {
            if (indexArr[a] <= rsd) {
                obj[indexArr[a]] = sum;
                sum += 1;
            }
            else {
                break;
            }
        }
    }
    return obj;
}
function getDataByType(data, len, direction, type, dataType) {
    var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o, _p, _q, _r, _s, _t, _u, _v, _w, _x, _y, _z, _0, _1, _2, _3, _4, _5, _6, _7, _8, _9, _10, _11, _12, _13, _14, _15, _16, _17, _18, _19, _20, _21, _22, _23, _24, _25, _26, _27, _28, _29, _30, _31, _32, _33, _34;
    data = _.cloneDeep(data);
    var applyData = [];
    if (type === "0" || data.length === 1) {
        // 复制单元格
        if (direction === "up" || direction === "left") {
            data.reverse();
        }
        applyData = fillCopy(data, len);
    }
    else if (type === "1") {
        // 填充序列
        if (dataType === "number") {
            // 数据类型是 数字
            applyData = fillSeries(data, len, direction);
        }
        else if (dataType === "extendNumber") {
            // 扩展数字
            var dataNumArr = [];
            for (var i = 0; i < data.length; i += 1) {
                var txt = (_a = data[i]) === null || _a === void 0 ? void 0 : _a.m;
                var _isExtendNumber = isExtendNumber(txt);
                if (_isExtendNumber[0]) {
                    dataNumArr.push(_isExtendNumber[1]);
                }
            }
            if (direction === "up" || direction === "left") {
                data.reverse();
                dataNumArr.reverse();
            }
            if (isEqualDiff(dataNumArr)) {
                // 等差数列，以等差为step
                var step = dataNumArr[1] - dataNumArr[0];
                applyData = fillExtendNumber(data, len, step);
            }
            else {
                // 不是等差数列，复制数据
                applyData = fillCopy(data, len);
            }
        }
        else if (dataType === "date") {
            // 数据类型是 日期
            if (direction === "up" || direction === "left") {
                data.reverse();
            }
            var _judgeDate = judgeDate(data);
            if (_judgeDate[0] && _judgeDate[3]) {
                // 日一样，月差为等差数列，以月差为step
                var step = dayjs((_b = data[1]) === null || _b === void 0 ? void 0 : _b.m).diff(dayjs((_c = data[0]) === null || _c === void 0 ? void 0 : _c.m), "months");
                applyData = fillMonths(data, len, step);
            }
            else if (!_judgeDate[0] && _judgeDate[2]) {
                // 日不一样，日差为等差数列，以日差为step
                var step = dayjs((_d = data[1]) === null || _d === void 0 ? void 0 : _d.m).diff(dayjs((_e = data[0]) === null || _e === void 0 ? void 0 : _e.m), "days");
                applyData = fillDays(data, len, step);
            }
            else {
                // 其它，复制数据
                applyData = fillCopy(data, len);
            }
        }
        else if (dataType === "chnNumber" && ((_f = data[0]) === null || _f === void 0 ? void 0 : _f.m) != null) {
            // 数据类型是 中文小写数字
            var hasweek = false;
            for (var i = 0; i < data.length; i += 1) {
                if (((_g = data[i]) === null || _g === void 0 ? void 0 : _g.m) === "日") {
                    hasweek = true;
                    break;
                }
            }
            var dataNumArr = [];
            var weekIndex = 0;
            for (var i = 0; i < data.length; i += 1) {
                var m = (_h = data[i]) === null || _h === void 0 ? void 0 : _h.m;
                if (m != null) {
                    m = "".concat(m);
                    if (m === "日") {
                        if (i === 0) {
                            dataNumArr.push(0);
                        }
                        else {
                            weekIndex += 1;
                            dataNumArr.push(weekIndex * 7);
                        }
                    }
                    else if (hasweek &&
                        chineseToNumber(m) > 0 &&
                        chineseToNumber(m) < 7) {
                        dataNumArr.push(chineseToNumber(m) + weekIndex * 7);
                    }
                    else {
                        dataNumArr.push(chineseToNumber(m));
                    }
                }
            }
            if (direction === "up" || direction === "left") {
                data.reverse();
                dataNumArr.reverse();
            }
            if (isEqualDiff(dataNumArr)) {
                if (hasweek ||
                    (dataNumArr[dataNumArr.length - 1] < 6 && dataNumArr[0] > 0) ||
                    (dataNumArr[0] < 6 && dataNumArr[dataNumArr.length - 1] > 0)) {
                    // 以周一~周日序列填充
                    var step = dataNumArr[1] - dataNumArr[0];
                    applyData = fillChnWeek(data, len, step);
                }
                else {
                    // 以中文小写数字序列填充
                    var step = dataNumArr[1] - dataNumArr[0];
                    applyData = fillChnNumber(data, len, step);
                }
            }
            else {
                // 不是等差数列，复制数据
                applyData = fillCopy(data, len);
            }
        }
        else if (dataType === "chnWeek2") {
            // 周一~周日
            var dataNumArr = [];
            var weekIndex = 0;
            for (var i = 0; i < data.length; i += 1) {
                var m = (_j = data[i]) === null || _j === void 0 ? void 0 : _j.m;
                if (m != null) {
                    m = "".concat(m);
                    var lastTxt = m.slice(m.length - 1, 1);
                    if (m === "周日") {
                        if (i === 0) {
                            dataNumArr.push(0);
                        }
                        else {
                            weekIndex += 1;
                            dataNumArr.push(weekIndex * 7);
                        }
                    }
                    else {
                        dataNumArr.push(chineseToNumber(lastTxt) + weekIndex * 7);
                    }
                }
            }
            if (direction === "up" || direction === "left") {
                data.reverse();
                dataNumArr.reverse();
            }
            if (isEqualDiff(dataNumArr)) {
                // 等差数列，以等差为step
                var step = dataNumArr[1] - dataNumArr[0];
                applyData = fillChnWeek2(data, len, step);
            }
            else {
                // 不是等差数列，复制数据
                applyData = fillCopy(data, len);
            }
        }
        else if (dataType === "chnWeek3") {
            // 星期一~星期日
            var dataNumArr = [];
            var weekIndex = 0;
            for (var i = 0; i < data.length; i += 1) {
                var m = (_k = data[i]) === null || _k === void 0 ? void 0 : _k.m;
                if (m != null) {
                    m = "".concat(m);
                    var lastTxt = m.slice(m.length - 1, 1);
                    if (m === "星期日") {
                        if (i === 0) {
                            dataNumArr.push(0);
                        }
                        else {
                            weekIndex += 1;
                            dataNumArr.push(weekIndex * 7);
                        }
                    }
                    else {
                        dataNumArr.push(chineseToNumber(lastTxt) + weekIndex * 7);
                    }
                }
            }
            if (direction === "up" || direction === "left") {
                data.reverse();
                dataNumArr.reverse();
            }
            if (isEqualDiff(dataNumArr)) {
                // 等差数列，以等差为step
                var step = dataNumArr[1] - dataNumArr[0];
                applyData = fillChnWeek3(data, len, step);
            }
            else {
                // 不是等差数列，复制数据
                applyData = fillCopy(data, len);
            }
        }
        else {
            // 数据类型是 其它
            if (direction === "up" || direction === "left") {
                data.reverse();
            }
            applyData = fillCopy(data, len);
        }
        // } else if (type === "2") {
        //   // 仅填充格式
        //   if (direction === "up" || direction === "left") {
        //     data.reverse();
        //   }
        //   applyData = fillOnlyFormat(data, len);
        // } else if (type === "3") {
        //   // 不带格式填充
        //   const dataArr = getDataByType(data, len, direction, "1", dataType);
        //   applyData = fillWithoutFormat(dataArr);
    }
    else if (type === "4") {
        // 以天数填充
        if (data.length === 2) {
            // 以日差为step
            if (direction === "up" || direction === "left") {
                data.reverse();
            }
            var step = dayjs((_l = data[1]) === null || _l === void 0 ? void 0 : _l.m).diff(dayjs((_m = data[0]) === null || _m === void 0 ? void 0 : _m.m), "days");
            applyData = fillDays(data, len, step);
        }
        else {
            if (direction === "up" || direction === "left") {
                data.reverse();
            }
            var _judgeDate = judgeDate(data);
            if (_judgeDate[0] && _judgeDate[3]) {
                // 日一样，且月差为等差数列，以月差为step
                var step = dayjs((_o = data[1]) === null || _o === void 0 ? void 0 : _o.m).diff(dayjs((_p = data[0]) === null || _p === void 0 ? void 0 : _p.m), "months");
                applyData = fillMonths(data, len, step);
            }
            else if (!_judgeDate[0] && _judgeDate[2]) {
                // 日不一样，且日差为等差数列，以日差为step
                var step = dayjs((_q = data[1]) === null || _q === void 0 ? void 0 : _q.m).diff(dayjs((_r = data[0]) === null || _r === void 0 ? void 0 : _r.m), "days");
                applyData = fillDays(data, len, step);
            }
            else {
                // 日差不是等差数列，复制数据
                applyData = fillCopy(data, len);
            }
        }
    }
    else if (type === "5") {
        // 以工作日填充
        if (data.length === 2) {
            if (dayjs((_s = data[1]) === null || _s === void 0 ? void 0 : _s.m).date() === dayjs((_t = data[0]) === null || _t === void 0 ? void 0 : _t.m).date() &&
                dayjs((_u = data[1]) === null || _u === void 0 ? void 0 : _u.m).diff(dayjs((_v = data[0]) === null || _v === void 0 ? void 0 : _v.m), "months") !== 0) {
                // 日一样，且月差大于一月，以月差为step（若那天为休息日，则向前取最近的工作日）
                if (direction === "up" || direction === "left") {
                    data.reverse();
                }
                var step = dayjs((_w = data[1]) === null || _w === void 0 ? void 0 : _w.m).diff(dayjs((_x = data[0]) === null || _x === void 0 ? void 0 : _x.m), "months");
                for (var i = 1; i <= len; i += 1) {
                    var index = (i - 1) % data.length;
                    var d = _.cloneDeep(data[index]);
                    var last = (_y = data[data.length - 1]) === null || _y === void 0 ? void 0 : _y.m;
                    if (d != null && last != null) {
                        var day = dayjs(last)
                            .add(step * i, "months")
                            .day();
                        var date = void 0;
                        if (day === 0) {
                            date = dayjs(last)
                                .add(step * i, "months")
                                .subtract(2, "days")
                                .format("YYYY-MM-DD");
                        }
                        else if (day === 6) {
                            date = dayjs(last)
                                .add(step * i, "months")
                                .subtract(1, "days")
                                .format("YYYY-MM-DD");
                        }
                        else {
                            date = dayjs(last)
                                .add(step * i, "months")
                                .format("YYYY-MM-DD");
                        }
                        d.m = date;
                        d.v = (_z = genarate(date)) === null || _z === void 0 ? void 0 : _z[2];
                        applyData.push(d);
                    }
                }
            }
            else {
                // 日不一样
                if (Math.abs(dayjs((_0 = data[1]) === null || _0 === void 0 ? void 0 : _0.m).diff(dayjs((_1 = data[0]) === null || _1 === void 0 ? void 0 : _1.m))) > 7) {
                    // 若日差大于7天，以一月为step（若那天是休息日，则向前取最近的工作日）
                    var step_month = void 0;
                    if (direction === "down" || direction === "right") {
                        step_month = 1;
                    }
                    else {
                        step_month = -1;
                        data.reverse();
                    }
                    var step = // 以数组第一个为对比
                     void 0; // 以数组第一个为对比
                    for (var i = 1; i <= len; i += 1) {
                        var index = (i - 1) % data.length;
                        var d = _.cloneDeep(data[index]);
                        if (d != null) {
                            var num = Math.ceil(i / data.length);
                            if (index === 0) {
                                step = dayjs(d.m)
                                    .add(step_month * num, "months")
                                    .diff(dayjs(d.m), "days");
                            }
                            var day = dayjs(d.m).add(step, "days").day();
                            var date = void 0;
                            if (day === 0) {
                                date = dayjs(d.m)
                                    .add(step, "days")
                                    .subtract(2, "days")
                                    .format("YYYY-MM-DD");
                            }
                            else if (day === 6) {
                                date = dayjs(d.m)
                                    .add(step, "days")
                                    .subtract(1, "days")
                                    .format("YYYY-MM-DD");
                            }
                            else {
                                date = dayjs(d.m).add(step, "days").format("YYYY-MM-DD");
                            }
                            d.m = date;
                            d.v = (_2 = genarate(date)) === null || _2 === void 0 ? void 0 : _2[2];
                            applyData.push(d);
                        }
                    }
                }
                else {
                    // 若日差小于等于7天，以7天为step（若那天是休息日，则向前取最近的工作日）
                    var step_day = void 0;
                    if (direction === "down" || direction === "right") {
                        step_day = 7;
                    }
                    else {
                        step_day = -7;
                        data.reverse();
                    }
                    var step = // 以数组第一个为对比
                     void 0; // 以数组第一个为对比
                    for (var i = 1; i <= len; i += 1) {
                        var index = (i - 1) % data.length;
                        var d = _.cloneDeep(data[index]);
                        if (d != null) {
                            var num = Math.ceil(i / data.length);
                            if (index === 0) {
                                step = dayjs(d.m)
                                    .add(step_day * num, "days")
                                    .diff(dayjs(d.m), "days");
                            }
                            var day = dayjs(d.m).add(step, "days").day();
                            var date = void 0;
                            if (day === 0) {
                                date = dayjs(d.m)
                                    .add(step, "days")
                                    .subtract(2, "days")
                                    .format("YYYY-MM-DD");
                            }
                            else if (day === 6) {
                                date = dayjs(d.m)
                                    .add(step, "days")
                                    .subtract(1, "days")
                                    .format("YYYY-MM-DD");
                            }
                            else {
                                date = dayjs(d.m).add(step, "days").format("YYYY-MM-DD");
                            }
                            d.m = date;
                            d.v = (_3 = genarate(date)) === null || _3 === void 0 ? void 0 : _3[2];
                            applyData.push(d);
                        }
                    }
                }
            }
        }
        else {
            var _judgeDate = judgeDate(data);
            if (_judgeDate[0] && _judgeDate[3]) {
                // 日一样，且月差为等差数列，以月差为step（若那天为休息日，则向前取最近的工作日）
                if (direction === "up" || direction === "left") {
                    data.reverse();
                }
                var step = dayjs((_4 = data[1]) === null || _4 === void 0 ? void 0 : _4.m).diff(dayjs((_5 = data[0]) === null || _5 === void 0 ? void 0 : _5.m), "months");
                for (var i = 1; i <= len; i += 1) {
                    var index = (i - 1) % data.length;
                    var d = _.cloneDeep(data[index]);
                    var last = (_6 = data[data.length - 1]) === null || _6 === void 0 ? void 0 : _6.m;
                    if (d != null) {
                        var day = dayjs(last)
                            .add(step * i, "months")
                            .day();
                        var date = void 0;
                        if (day === 0) {
                            date = dayjs(last)
                                .add(step * i, "months")
                                .subtract(2, "days")
                                .format("YYYY-MM-DD");
                        }
                        else if (day === 6) {
                            date = dayjs(last)
                                .add(step * i, "months")
                                .subtract(1, "days")
                                .format("YYYY-MM-DD");
                        }
                        else {
                            date = dayjs(last)
                                .add(step * i, "months")
                                .format("YYYY-MM-DD");
                        }
                        d.m = date;
                        d.v = (_7 = genarate(date)) === null || _7 === void 0 ? void 0 : _7[2];
                        applyData.push(d);
                    }
                }
            }
            else if (!_judgeDate[0] && _judgeDate[2]) {
                // 日不一样，且日差为等差数列
                if (Math.abs(dayjs((_8 = data[1]) === null || _8 === void 0 ? void 0 : _8.m).diff(dayjs((_9 = data[0]) === null || _9 === void 0 ? void 0 : _9.m))) > 7) {
                    // 若日差大于7天，以一月为step（若那天是休息日，则向前取最近的工作日）
                    var step_month = void 0;
                    if (direction === "down" || direction === "right") {
                        step_month = 1;
                    }
                    else {
                        step_month = -1;
                        data.reverse();
                    }
                    var step = // 以数组第一个为对比
                     void 0; // 以数组第一个为对比
                    for (var i = 1; i <= len; i += 1) {
                        var index = (i - 1) % data.length;
                        var d = _.cloneDeep(data[index]);
                        if (d != null) {
                            var num = Math.ceil(i / data.length);
                            if (index === 0) {
                                step = dayjs(d.m)
                                    .add(step_month * num, "months")
                                    .diff(dayjs(d.m), "days");
                            }
                            var day = dayjs(d.m).add(step, "days").day();
                            var date = void 0;
                            if (day === 0) {
                                date = dayjs(d.m)
                                    .add(step, "days")
                                    .subtract(2, "days")
                                    .format("YYYY-MM-DD");
                            }
                            else if (day === 6) {
                                date = dayjs(d.m)
                                    .add(step, "days")
                                    .subtract(1, "days")
                                    .format("YYYY-MM-DD");
                            }
                            else {
                                date = dayjs(d.m).add(step, "days").format("YYYY-MM-DD");
                            }
                            d.m = date;
                            d.v = (_10 = genarate(date)) === null || _10 === void 0 ? void 0 : _10[2];
                            applyData.push(d);
                        }
                    }
                }
                else {
                    // 若日差小于等于7天，以7天为step（若那天是休息日，则向前取最近的工作日）
                    var step_day = void 0;
                    if (direction === "down" || direction === "right") {
                        step_day = 7;
                    }
                    else {
                        step_day = -7;
                        data.reverse();
                    }
                    var step = // 以数组第一个为对比
                     void 0; // 以数组第一个为对比
                    for (var i = 1; i <= len; i += 1) {
                        var index = (i - 1) % data.length;
                        var d = _.cloneDeep(data[index]);
                        if (d != null) {
                            var num = Math.ceil(i / data.length);
                            if (index === 0) {
                                step = dayjs(d.m)
                                    .add(step_day * num, "days")
                                    .diff(dayjs(d.m), "days");
                            }
                            var day = dayjs(d.m).add(step, "days").day();
                            var date = void 0;
                            if (day === 0) {
                                date = dayjs(d.m)
                                    .add(step, "days")
                                    .subtract(2, "days")
                                    .format("YYYY-MM-DD");
                            }
                            else if (day === 6) {
                                date = dayjs(d.m)
                                    .add(step, "days")
                                    .subtract(1, "days")
                                    .format("YYYY-MM-DD");
                            }
                            else {
                                date = dayjs(d.m).add(step, "days").format("YYYY-MM-DD");
                            }
                            d.m = date;
                            d.v = (_11 = genarate(date)) === null || _11 === void 0 ? void 0 : _11[2];
                            applyData.push(d);
                        }
                    }
                }
            }
            else {
                // 日差不是等差数列，复制数据
                if (direction === "up" || direction === "left") {
                    data.reverse();
                }
                applyData = fillCopy(data, len);
            }
        }
    }
    else if (type === "6") {
        // 以月填充
        if (data.length === 2) {
            if (dayjs((_12 = data[1]) === null || _12 === void 0 ? void 0 : _12.m).date() === dayjs((_13 = data[0]) === null || _13 === void 0 ? void 0 : _13.m).date() &&
                dayjs((_14 = data[1]) === null || _14 === void 0 ? void 0 : _14.m).diff(dayjs((_15 = data[0]) === null || _15 === void 0 ? void 0 : _15.m), "months") !== 0) {
                // 日一样，且月差大于一月，以月差为step
                if (direction === "up" || direction === "left") {
                    data.reverse();
                }
                var step = dayjs((_16 = data[1]) === null || _16 === void 0 ? void 0 : _16.m).diff(dayjs((_17 = data[0]) === null || _17 === void 0 ? void 0 : _17.m), "months");
                applyData = fillMonths(data, len, step);
            }
            else {
                // 以一月为step
                var step_month = void 0;
                if (direction === "down" || direction === "right") {
                    step_month = 1;
                }
                else {
                    step_month = -1;
                    data.reverse();
                }
                var step = // 以数组第一个为对比
                 void 0; // 以数组第一个为对比
                for (var i = 1; i <= len; i += 1) {
                    var index = (i - 1) % data.length;
                    var d = _.cloneDeep(data[index]);
                    if (d != null) {
                        var num = Math.ceil(i / data.length);
                        if (index === 0) {
                            step = dayjs(d.m)
                                .add(step_month * num, "months")
                                .diff(dayjs(d.m), "days");
                        }
                        var date = dayjs(d.m).add(step, "days").format("YYYY-MM-DD");
                        d.m = date;
                        d.v = (_18 = genarate(date)) === null || _18 === void 0 ? void 0 : _18[2];
                        applyData.push(d);
                    }
                }
            }
        }
        else {
            var _judgeDate = judgeDate(data);
            if (_judgeDate[0] && _judgeDate[3]) {
                // 日一样，且月差为等差数列，以月差为step
                if (direction === "up" || direction === "left") {
                    data.reverse();
                }
                var step = dayjs((_19 = data[1]) === null || _19 === void 0 ? void 0 : _19.m).diff(dayjs((_20 = data[0]) === null || _20 === void 0 ? void 0 : _20.m), "months");
                applyData = fillMonths(data, len, step);
            }
            else if (!_judgeDate[0] && _judgeDate[2]) {
                // 日不一样，且日差为等差数列，以一月为step
                var step_month = void 0;
                if (direction === "down" || direction === "right") {
                    step_month = 1;
                }
                else {
                    step_month = -1;
                    data.reverse();
                }
                var step = // 以数组第一个为对比
                 void 0; // 以数组第一个为对比
                for (var i = 1; i <= len; i += 1) {
                    var index = (i - 1) % data.length;
                    var d = _.cloneDeep(data[index]);
                    if (d != null) {
                        var num = Math.ceil(i / data.length);
                        if (index === 0) {
                            step = dayjs(d.m)
                                .add(step_month * num, "months")
                                .diff(dayjs(d.m), "days");
                        }
                        var date = dayjs(d.m).add(step, "days").format("YYYY-MM-DD");
                        d.m = date;
                        d.v = (_21 = genarate(date)) === null || _21 === void 0 ? void 0 : _21[2];
                        applyData.push(d);
                    }
                }
            }
            else {
                // 日差不是等差数列，复制数据
                if (direction === "up" || direction === "left") {
                    data.reverse();
                }
                applyData = fillCopy(data, len);
            }
        }
    }
    else if (type === "7") {
        // 以年填充
        if (data.length === 2) {
            if (dayjs((_22 = data[1]) === null || _22 === void 0 ? void 0 : _22.m).date() === dayjs((_23 = data[0]) === null || _23 === void 0 ? void 0 : _23.m).date() &&
                dayjs((_24 = data[1]) === null || _24 === void 0 ? void 0 : _24.m).month() === dayjs((_25 = data[0]) === null || _25 === void 0 ? void 0 : _25.m).month() &&
                dayjs((_26 = data[1]) === null || _26 === void 0 ? void 0 : _26.m).diff(dayjs((_27 = data[0]) === null || _27 === void 0 ? void 0 : _27.m), "years") !== 0) {
                // 日月一样，且年差大于一年，以年差为step
                if (direction === "up" || direction === "left") {
                    data.reverse();
                }
                var step = dayjs((_28 = data[1]) === null || _28 === void 0 ? void 0 : _28.m).diff(dayjs((_29 = data[0]) === null || _29 === void 0 ? void 0 : _29.m), "years");
                applyData = fillYears(data, len, step);
            }
            else {
                // 以一年为step
                var step_year = void 0;
                if (direction === "down" || direction === "right") {
                    step_year = 1;
                }
                else {
                    step_year = -1;
                    data.reverse();
                }
                var step = // 以数组第一个为对比
                 void 0; // 以数组第一个为对比
                for (var i = 1; i <= len; i += 1) {
                    var index = (i - 1) % data.length;
                    var d = _.cloneDeep(data[index]);
                    if (d != null) {
                        var num = Math.ceil(i / data.length);
                        if (index === 0) {
                            step = dayjs(d.m)
                                .add(step_year * num, "years")
                                .diff(dayjs(d.m), "days");
                        }
                        var date = dayjs(d.m).add(step, "days").format("YYYY-MM-DD");
                        d.m = date;
                        d.v = (_30 = genarate(date)) === null || _30 === void 0 ? void 0 : _30[2];
                        applyData.push(d);
                    }
                }
            }
        }
        else {
            var _judgeDate = judgeDate(data);
            if (_judgeDate[0] && _judgeDate[1] && _judgeDate[4]) {
                // 日月一样，且年差为等差数列，以年差为step
                if (direction === "up" || direction === "left") {
                    data.reverse();
                }
                var step = dayjs((_31 = data[1]) === null || _31 === void 0 ? void 0 : _31.m).diff(dayjs((_32 = data[0]) === null || _32 === void 0 ? void 0 : _32.m), "years");
                applyData = fillYears(data, len, step);
            }
            else if ((_judgeDate[0] && _judgeDate[3]) || _judgeDate[2]) {
                // 日一样且月差为等差数列，或天差为等差数列，以一年为step
                var step_year = void 0;
                if (direction === "down" || direction === "right") {
                    step_year = 1;
                }
                else {
                    step_year = -1;
                    data.reverse();
                }
                var step = // 以数组第一个为对比
                 void 0; // 以数组第一个为对比
                for (var i = 1; i <= len; i += 1) {
                    var index = (i - 1) % data.length;
                    var d = _.cloneDeep(data[index]);
                    var num = Math.ceil(i / data.length);
                    if (d != null) {
                        if (index === 0) {
                            step = dayjs(d.m)
                                .add(step_year * num, "years")
                                .diff(dayjs(d.m), "days");
                        }
                        var date = dayjs(d.m).add(step, "days").format("YYYY-MM-DD");
                        d.m = date;
                        d.v = (_33 = genarate(date)) === null || _33 === void 0 ? void 0 : _33[2];
                        applyData.push(d);
                    }
                }
            }
            else {
                // 日差不是等差数列，复制数据
                if (direction === "up" || direction === "left") {
                    data.reverse();
                }
                applyData = fillCopy(data, len);
            }
        }
    }
    else if (type === "8") {
        // 以中文小写数字序列填充
        var dataNumArr = [];
        for (var i = 0; i < data.length; i += 1) {
            var m = (_34 = data[i]) === null || _34 === void 0 ? void 0 : _34.m;
            if (m != null) {
                m = "".concat(m);
                dataNumArr.push(chineseToNumber(m));
            }
        }
        if (direction === "up" || direction === "left") {
            data.reverse();
            dataNumArr.reverse();
        }
        if (isEqualDiff(dataNumArr)) {
            var step = dataNumArr[1] - dataNumArr[0];
            applyData = fillChnNumber(data, len, step);
        }
        else {
            // 不是等差数列，复制数据
            applyData = fillCopy(data, len);
        }
    }
    return applyData;
}
function getCopyData(d, r1, r2, c1, c2, direction) {
    var copyData = [];
    var a1;
    var a2;
    var b1;
    var b2;
    if (direction === "down" || direction === "up") {
        a1 = c1;
        a2 = c2;
        b1 = r1;
        b2 = r2;
    }
    else {
        a1 = r1;
        a2 = r2;
        b1 = c1;
        b2 = c2;
    }
    for (var a = a1; a <= a2; a += 1) {
        var obj = {};
        var arrData = [];
        var arrIndex = [];
        var text = "";
        var extendNumberBeforeStr = null;
        var extendNumberAfterStr = null;
        var isSameStr = true;
        for (var b = b1; b <= b2; b += 1) {
            // 单元格
            var data = void 0;
            if (direction === "down" || direction === "up") {
                data = d[b][a];
            }
            else if (direction === "right" || direction === "left") {
                data = d[a][b];
            }
            // 单元格值类型
            var str = void 0;
            if ((data === null || data === void 0 ? void 0 : data.v) != null && data.f == null) {
                if (!!data.ct && data.ct.t === "n") {
                    str = "number";
                    extendNumberBeforeStr = null;
                    extendNumberAfterStr = null;
                }
                else if (!!data.ct && data.ct.t === "d") {
                    str = "date";
                    extendNumberBeforeStr = null;
                    extendNumberAfterStr = null;
                }
                else if (isExtendNumber(data.m)[0]) {
                    str = "extendNumber";
                    var _isExtendNumber = isExtendNumber(data.m);
                    if (extendNumberBeforeStr == null || extendNumberAfterStr == null) {
                        isSameStr = true;
                        extendNumberBeforeStr = _isExtendNumber[2], extendNumberAfterStr = _isExtendNumber[3];
                    }
                    else {
                        if (_isExtendNumber[2] !== extendNumberBeforeStr ||
                            _isExtendNumber[3] !== extendNumberAfterStr) {
                            isSameStr = false;
                            extendNumberBeforeStr = _isExtendNumber[2], extendNumberAfterStr = _isExtendNumber[3];
                        }
                        else {
                            isSameStr = true;
                        }
                    }
                }
                else if (isChnNumber(data.m)) {
                    str = "chnNumber";
                    extendNumberBeforeStr = null;
                    extendNumberAfterStr = null;
                }
                else if (isChnWeek2(data.m)) {
                    str = "chnWeek2";
                    extendNumberBeforeStr = null;
                    extendNumberAfterStr = null;
                }
                else if (isChnWeek3(data.m)) {
                    str = "chnWeek3";
                    extendNumberBeforeStr = null;
                    extendNumberAfterStr = null;
                }
                else {
                    str = "other";
                    extendNumberBeforeStr = null;
                    extendNumberAfterStr = null;
                }
            }
            else {
                str = "other";
                extendNumberBeforeStr = null;
                extendNumberAfterStr = null;
            }
            if (str === "extendNumber") {
                if (b === b1) {
                    if (b1 === b2) {
                        text = str;
                        arrData.push(data);
                        arrIndex.push(b - b1 + 1);
                        obj[text] = [];
                        obj[text].push({ data: arrData, index: arrIndex });
                    }
                    else {
                        text = str;
                        arrData.push(data);
                        arrIndex.push(b - b1 + 1);
                    }
                }
                else if (b === b2) {
                    if (text === str && isSameStr) {
                        arrData.push(data);
                        arrIndex.push(b - b1 + 1);
                        if (text in obj) {
                            obj[text].push({ data: arrData, index: arrIndex });
                        }
                        else {
                            obj[text] = [];
                            obj[text].push({ data: arrData, index: arrIndex });
                        }
                    }
                    else {
                        if (text in obj) {
                            obj[text].push({ data: arrData, index: arrIndex });
                        }
                        else {
                            obj[text] = [];
                            obj[text].push({ data: arrData, index: arrIndex });
                        }
                        text = str;
                        arrData = [];
                        arrData.push(data);
                        arrIndex = [];
                        arrIndex.push(b - b1 + 1);
                        if (text in obj) {
                            obj[text].push({ data: arrData, index: arrIndex });
                        }
                        else {
                            obj[text] = [];
                            obj[text].push({ data: arrData, index: arrIndex });
                        }
                    }
                }
                else {
                    if (text === str && isSameStr) {
                        arrData.push(data);
                        arrIndex.push(b - b1 + 1);
                    }
                    else {
                        if (text in obj) {
                            obj[text].push({ data: arrData, index: arrIndex });
                        }
                        else {
                            obj[text] = [];
                            obj[text].push({ data: arrData, index: arrIndex });
                        }
                        text = str;
                        arrData = [];
                        arrData.push(data);
                        arrIndex = [];
                        arrIndex.push(b - b1 + 1);
                    }
                }
            }
            else {
                if (b === b1) {
                    if (b1 === b2) {
                        text = str;
                        arrData.push(data);
                        arrIndex.push(b - b1 + 1);
                        obj[text] = [];
                        obj[text].push({ data: arrData, index: arrIndex });
                    }
                    else {
                        text = str;
                        arrData.push(data);
                        arrIndex.push(b - b1 + 1);
                    }
                }
                else if (b === b2) {
                    if (text === str) {
                        arrData.push(data);
                        arrIndex.push(b - b1 + 1);
                        if (text in obj) {
                            obj[text].push({ data: arrData, index: arrIndex });
                        }
                        else {
                            obj[text] = [];
                            obj[text].push({ data: arrData, index: arrIndex });
                        }
                    }
                    else {
                        if (text in obj) {
                            obj[text].push({ data: arrData, index: arrIndex });
                        }
                        else {
                            obj[text] = [];
                            obj[text].push({ data: arrData, index: arrIndex });
                        }
                        text = str;
                        arrData = [];
                        arrData.push(data);
                        arrIndex = [];
                        arrIndex.push(b - b1 + 1);
                        if (text in obj) {
                            obj[text].push({ data: arrData, index: arrIndex });
                        }
                        else {
                            obj[text] = [];
                            obj[text].push({ data: arrData, index: arrIndex });
                        }
                    }
                }
                else {
                    if (text === str) {
                        arrData.push(data);
                        arrIndex.push(b - b1 + 1);
                    }
                    else {
                        if (text in obj) {
                            obj[text].push({ data: arrData, index: arrIndex });
                        }
                        else {
                            obj[text] = [];
                            obj[text].push({ data: arrData, index: arrIndex });
                        }
                        text = str;
                        arrData = [];
                        arrData.push(data);
                        arrIndex = [];
                        arrIndex.push(b - b1 + 1);
                    }
                }
            }
        }
        copyData.push(obj);
    }
    return copyData;
}
function getApplyData(copyD, csLen, asLen) {
    var applyData = [];
    var direction = dropCellCache.direction;
    var type = dropCellCache.applyType;
    var num = Math.floor(asLen / csLen);
    var rsd = asLen % csLen;
    // 纯数字类型
    var copyD_number = copyD.number;
    var applyD_number = [];
    if (copyD_number) {
        for (var i = 0; i < copyD_number.length; i += 1) {
            var s = getLenS(copyD_number[i].index, rsd);
            var len = copyD_number[i].index.length * num + s;
            var arrData = void 0;
            if (type === "1" || type === "3") {
                arrData = getDataByType(copyD_number[i].data, len, direction, type, "number");
            }
            else if (type === "2") {
                arrData = getDataByType(copyD_number[i].data, len, direction, type);
            }
            else {
                arrData = getDataByType(copyD_number[i].data, len, direction, "0");
            }
            var arrIndex = getDataIndex(csLen, asLen, copyD_number[i].index);
            applyD_number.push({ data: arrData, index: arrIndex });
        }
    }
    // 扩展数字型（即一串字符最后面的是数字）
    var copyD_extendNumber = copyD.extendNumber;
    var applyD_extendNumber = [];
    if (copyD_extendNumber) {
        for (var i = 0; i < copyD_extendNumber.length; i += 1) {
            var s = getLenS(copyD_extendNumber[i].index, rsd);
            var len = copyD_extendNumber[i].index.length * num + s;
            var arrData = void 0;
            if (type === "1" || type === "3") {
                arrData = getDataByType(copyD_extendNumber[i].data, len, direction, type, "extendNumber");
            }
            else if (type === "2") {
                arrData = getDataByType(copyD_extendNumber[i].data, len, direction, type);
            }
            else {
                arrData = getDataByType(copyD_extendNumber[i].data, len, direction, "0");
            }
            var arrIndex = getDataIndex(csLen, asLen, copyD_extendNumber[i].index);
            applyD_extendNumber.push({ data: arrData, index: arrIndex });
        }
    }
    // 日期类型
    var copyD_date = copyD.date;
    var applyD_date = [];
    if (copyD_date) {
        for (var i = 0; i < copyD_date.length; i += 1) {
            var s = getLenS(copyD_date[i].index, rsd);
            var len = copyD_date[i].index.length * num + s;
            var arrData = void 0;
            if (type === "1" || type === "3") {
                arrData = getDataByType(copyD_date[i].data, len, direction, type, "date");
            }
            else if (type === "8") {
                arrData = getDataByType(copyD_date[i].data, len, direction, "0");
            }
            else {
                arrData = getDataByType(copyD_date[i].data, len, direction, type);
            }
            var arrIndex = getDataIndex(csLen, asLen, copyD_date[i].index);
            applyD_date.push({ data: arrData, index: arrIndex });
        }
    }
    // 中文小写数字 或 一~日
    var copyD_chnNumber = copyD.chnNumber;
    var applyD_chnNumber = [];
    if (copyD_chnNumber) {
        for (var i = 0; i < copyD_chnNumber.length; i += 1) {
            var s = getLenS(copyD_chnNumber[i].index, rsd);
            var len = copyD_chnNumber[i].index.length * num + s;
            var arrData = void 0;
            if (type === "1" || type === "3") {
                arrData = getDataByType(copyD_chnNumber[i].data, len, direction, type, "chnNumber");
            }
            else if (type === "2" || type === "8") {
                arrData = getDataByType(copyD_chnNumber[i].data, len, direction, type);
            }
            else {
                arrData = getDataByType(copyD_chnNumber[i].data, len, direction, "0");
            }
            var arrIndex = getDataIndex(csLen, asLen, copyD_chnNumber[i].index);
            applyD_chnNumber.push({ data: arrData, index: arrIndex });
        }
    }
    // 周一~周日
    var copyD_chnWeek2 = copyD.chnWeek2;
    var applyD_chnWeek2 = [];
    if (copyD_chnWeek2) {
        for (var i = 0; i < copyD_chnWeek2.length; i += 1) {
            var s = getLenS(copyD_chnWeek2[i].index, rsd);
            var len = copyD_chnWeek2[i].index.length * num + s;
            var arrData = void 0;
            if (type === "1" || type === "3") {
                arrData = getDataByType(copyD_chnWeek2[i].data, len, direction, type, "chnWeek2");
            }
            else if (type === "2") {
                arrData = getDataByType(copyD_chnWeek2[i].data, len, direction, type);
            }
            else {
                arrData = getDataByType(copyD_chnWeek2[i].data, len, direction, "0");
            }
            var arrIndex = getDataIndex(csLen, asLen, copyD_chnWeek2[i].index);
            applyD_chnWeek2.push({ data: arrData, index: arrIndex });
        }
    }
    // 星期一~星期日
    var copyD_chnWeek3 = copyD.chnWeek3;
    var applyD_chnWeek3 = [];
    if (copyD_chnWeek3) {
        for (var i = 0; i < copyD_chnWeek3.length; i += 1) {
            var s = getLenS(copyD_chnWeek3[i].index, rsd);
            var len = copyD_chnWeek3[i].index.length * num + s;
            var arrData = void 0;
            if (type === "1" || type === "3") {
                arrData = getDataByType(copyD_chnWeek3[i].data, len, direction, type, "chnWeek3");
            }
            else if (type === "2") {
                arrData = getDataByType(copyD_chnWeek3[i].data, len, direction, type);
            }
            else {
                arrData = getDataByType(copyD_chnWeek3[i].data, len, direction, "0");
            }
            var arrIndex = getDataIndex(csLen, asLen, copyD_chnWeek3[i].index);
            applyD_chnWeek3.push({ data: arrData, index: arrIndex });
        }
    }
    // 其它
    var copyD_other = copyD.other;
    var applyD_other = [];
    if (copyD_other) {
        for (var i = 0; i < copyD_other.length; i += 1) {
            var s = getLenS(copyD_other[i].index, rsd);
            var len = copyD_other[i].index.length * num + s;
            var arrData = void 0;
            if (type === "2" || type === "3") {
                arrData = getDataByType(copyD_other[i].data, len, direction, type);
            }
            else {
                arrData = getDataByType(copyD_other[i].data, len, direction, "0");
            }
            var arrIndex = getDataIndex(csLen, asLen, copyD_other[i].index);
            applyD_other.push({ data: arrData, index: arrIndex });
        }
    }
    for (var x = 1; x <= asLen; x += 1) {
        if (applyD_number.length > 0) {
            for (var y = 0; y < applyD_number.length; y += 1) {
                if (x in applyD_number[y].index) {
                    applyData.push(applyD_number[y].data[applyD_number[y].index[x]]);
                }
            }
        }
        if (applyD_extendNumber.length > 0) {
            for (var y = 0; y < applyD_extendNumber.length; y += 1) {
                if (x in applyD_extendNumber[y].index) {
                    applyData.push(applyD_extendNumber[y].data[applyD_extendNumber[y].index[x]]);
                }
            }
        }
        if (applyD_date.length > 0) {
            for (var y = 0; y < applyD_date.length; y += 1) {
                if (x in applyD_date[y].index) {
                    applyData.push(applyD_date[y].data[applyD_date[y].index[x]]);
                }
            }
        }
        if (applyD_chnNumber.length > 0) {
            for (var y = 0; y < applyD_chnNumber.length; y += 1) {
                if (x in applyD_chnNumber[y].index) {
                    applyData.push(applyD_chnNumber[y].data[applyD_chnNumber[y].index[x]]);
                }
            }
        }
        if (applyD_chnWeek2.length > 0) {
            for (var y = 0; y < applyD_chnWeek2.length; y += 1) {
                if (x in applyD_chnWeek2[y].index) {
                    applyData.push(applyD_chnWeek2[y].data[applyD_chnWeek2[y].index[x]]);
                }
            }
        }
        if (applyD_chnWeek3.length > 0) {
            for (var y = 0; y < applyD_chnWeek3.length; y += 1) {
                if (x in applyD_chnWeek3[y].index) {
                    applyData.push(applyD_chnWeek3[y].data[applyD_chnWeek3[y].index[x]]);
                }
            }
        }
        if (applyD_other.length > 0) {
            for (var y = 0; y < applyD_other.length; y += 1) {
                if (x in applyD_other[y].index) {
                    applyData.push(applyD_other[y].data[applyD_other[y].index[x]]);
                }
            }
        }
    }
    return applyData;
}
export function updateDropCell(ctx) {
    // if (
    //   !checkProtectionLockedRangeList([_this.applyRange], ctx.currentSheetId)
    // ) {
    //   return;
    // }
    var _a, _b, _c, _d;
    var _e, _f, _g;
    var d = getFlowdata(ctx);
    var allowEdit = isAllowEdit(ctx);
    if (allowEdit === false || d == null) {
        return;
    }
    var index = getSheetIndex(ctx, ctx.currentSheetId);
    if (index == null)
        return;
    var file = ctx.luckysheetfile[index];
    var hiddenRows = new Set(Object.keys(((_e = file.config) === null || _e === void 0 ? void 0 : _e.rowhidden) || {}));
    var hiddenCols = new Set(Object.keys(((_f = file.config) === null || _f === void 0 ? void 0 : _f.colhidden) || {}));
    var cfg = _.cloneDeep(ctx.config);
    if (cfg.borderInfo == null) {
        cfg.borderInfo = [];
    }
    var borderInfoCompute = getBorderInfoCompute(ctx, ctx.currentSheetId);
    var dataVerification = _.cloneDeep(file.dataVerification);
    var direction = dropCellCache.direction;
    // const type = dropCellCache.applyType;
    // 复制范围
    var copyRange = dropCellCache.copyRange;
    var copy_str_r = copyRange.row[0];
    var copy_end_r = copyRange.row[1];
    var copy_str_c = copyRange.column[0];
    var copy_end_c = copyRange.column[1];
    var copyData = getCopyData(d, copy_str_r, copy_end_r, copy_str_c, copy_end_c, direction);
    var csLen;
    if (direction === "down" || direction === "up") {
        csLen = copy_end_r - copy_str_r + 1;
    }
    else {
        // direction === "right" || direction === "left"
        csLen = copy_end_c - copy_str_c + 1;
    }
    // 应用范围
    var applyRange = dropCellCache.applyRange;
    var apply_str_r = applyRange.row[0];
    var apply_end_r = applyRange.row[1];
    var apply_str_c = applyRange.column[0];
    var apply_end_c = applyRange.column[1];
    if (direction === "down" || direction === "up") {
        var asLen = apply_end_r - apply_str_r + 1;
        for (var i = apply_str_c; i <= apply_end_c; i += 1) {
            if (hiddenCols.has("".concat(i)))
                continue;
            var copyD = copyData[i - apply_str_c];
            var applyData = getApplyData(copyD, csLen, asLen);
            if (direction === "down") {
                for (var j = apply_str_r; j <= apply_end_r; j += 1) {
                    if (hiddenRows.has("".concat(j)))
                        continue;
                    var cell = applyData[j - apply_str_r];
                    if ((cell === null || cell === void 0 ? void 0 : cell.f) != null) {
                        var f = "=".concat(formula.functionCopy(ctx, cell.f, "down", j - apply_str_r + 1));
                        var v = formula.execfunction(ctx, f, j, i);
                        formula.execFunctionGroup(ctx, j, i, v[1], undefined, d);
                        cell.v = v[1], cell.f = v[2];
                        if (cell.spl != null) {
                            cell.spl = v[3].data;
                        }
                        else if (cell.v != null) {
                            if (isRealNum(cell.v) &&
                                !/^\d{6}(18|19|20)?\d{2}(0[1-9]|1[12])(0[1-9]|[12]\d|3[01])\d{3}(\d|X)$/i.test("".concat(cell.v))) {
                                if (cell.v === Infinity || cell.v === -Infinity) {
                                    cell.m = cell.v.toString();
                                }
                                else {
                                    if (cell.v.toString().indexOf("e") > -1) {
                                        var len = cell.v
                                            .toString()
                                            .split(".")[1]
                                            .split("e")[0].length;
                                        if (len > 5) {
                                            len = 5;
                                        }
                                        cell.m = cell.v.toExponential(len).toString();
                                    }
                                    else {
                                        var mask = void 0;
                                        if (((_g = cell.ct) === null || _g === void 0 ? void 0 : _g.fa) === "##0.00") {
                                            /* 如果是数字类型 */
                                            mask = genarate("".concat(Math.round(cell.v * 1000000000) /
                                                1000000000, ".00"));
                                            cell.m = mask[0].toString();
                                        }
                                        else {
                                            mask = genarate(Math.round(cell.v * 1000000000) / 1000000000);
                                            cell.m = mask[0].toString();
                                        }
                                    }
                                }
                                cell.ct = cell.ct || { fa: "General", t: "n" };
                            }
                            else {
                                var mask = genarate(cell.v);
                                cell.m = mask[0].toString();
                                _a = mask, cell.ct = _a[1];
                            }
                        }
                    }
                    d[j][i] = cell || null;
                    // 边框
                    var bd_r = copy_str_r + ((j - apply_str_r) % csLen);
                    var bd_c = i;
                    if (borderInfoCompute["".concat(bd_r, "_").concat(bd_c)]) {
                        var bd_obj = {
                            rangeType: "cell",
                            value: {
                                row_index: j,
                                col_index: i,
                                l: borderInfoCompute["".concat(bd_r, "_").concat(bd_c)].l,
                                r: borderInfoCompute["".concat(bd_r, "_").concat(bd_c)].r,
                                t: borderInfoCompute["".concat(bd_r, "_").concat(bd_c)].t,
                                b: borderInfoCompute["".concat(bd_r, "_").concat(bd_c)].b,
                            },
                        };
                        cfg.borderInfo.push(bd_obj);
                    }
                    else if (borderInfoCompute["".concat(j, "_").concat(i)]) {
                        var bd_obj = {
                            rangeType: "cell",
                            value: {
                                row_index: j,
                                col_index: i,
                                l: null,
                                r: null,
                                t: null,
                                b: null,
                            },
                        };
                        cfg.borderInfo.push(bd_obj);
                    }
                    // 数据验证
                    // Bug
                    if (dataVerification != null && dataVerification["".concat(bd_r, "_").concat(bd_c)]) {
                        dataVerification["".concat(j, "_").concat(i)] = dataVerification["".concat(bd_r, "_").concat(bd_c)];
                    }
                }
            }
            if (direction === "up") {
                for (var j = apply_end_r; j >= apply_str_r; j -= 1) {
                    if (hiddenRows.has("".concat(j)))
                        continue;
                    var cell = applyData[apply_end_r - j];
                    if ((cell === null || cell === void 0 ? void 0 : cell.f) != null) {
                        var f = "=".concat(formula.functionCopy(ctx, cell.f, "up", apply_end_r - j + 1));
                        var v = formula.execfunction(ctx, f, j, i);
                        formula.execFunctionGroup(ctx, j, i, v[1], undefined, d);
                        cell.v = v[1], cell.f = v[2];
                        if (cell.spl != null) {
                            cell.spl = v[3].data;
                        }
                        else if (cell.v != null) {
                            if (isRealNum(cell.v) &&
                                !/^\d{6}(18|19|20)?\d{2}(0[1-9]|1[12])(0[1-9]|[12]\d|3[01])\d{3}(\d|X)$/i.test("".concat(cell.v))) {
                                if (cell.v === Infinity || cell.v === -Infinity) {
                                    cell.m = cell.v.toString();
                                }
                                else {
                                    if (cell.v.toString().indexOf("e") > -1) {
                                        var len = cell.v
                                            .toString()
                                            .split(".")[1]
                                            .split("e")[0].length;
                                        if (len > 5) {
                                            len = 5;
                                        }
                                        cell.m = cell.v.toExponential(len).toString();
                                    }
                                    else {
                                        var mask = genarate(Math.round(cell.v * 1000000000) / 1000000000);
                                        cell.m = mask[0].toString();
                                    }
                                }
                                cell.ct = { fa: "General", t: "n" };
                            }
                            else {
                                var mask = genarate(cell.v);
                                cell.m = mask[0].toString();
                                _b = mask, cell.ct = _b[1];
                            }
                        }
                    }
                    d[j][i] = cell || null;
                    // 边框
                    var bd_r = copy_end_r - ((apply_end_r - j) % csLen);
                    var bd_c = i;
                    if (borderInfoCompute["".concat(bd_r, "_").concat(bd_c)]) {
                        var bd_obj = {
                            rangeType: "cell",
                            value: {
                                row_index: j,
                                col_index: i,
                                l: borderInfoCompute["".concat(bd_r, "_").concat(bd_c)].l,
                                r: borderInfoCompute["".concat(bd_r, "_").concat(bd_c)].r,
                                t: borderInfoCompute["".concat(bd_r, "_").concat(bd_c)].t,
                                b: borderInfoCompute["".concat(bd_r, "_").concat(bd_c)].b,
                            },
                        };
                        cfg.borderInfo.push(bd_obj);
                    }
                    else if (borderInfoCompute["".concat(j, "_").concat(i)]) {
                        var bd_obj = {
                            rangeType: "cell",
                            value: {
                                row_index: j,
                                col_index: i,
                                l: null,
                                r: null,
                                t: null,
                                b: null,
                            },
                        };
                        cfg.borderInfo.push(bd_obj);
                    }
                    // 数据验证
                    if (dataVerification != null && dataVerification["".concat(bd_r, "_").concat(bd_c)]) {
                        dataVerification["".concat(j, "_").concat(i)] = dataVerification["".concat(bd_r, "_").concat(bd_c)];
                    }
                }
            }
        }
    }
    else if (direction === "right" || direction === "left") {
        var asLen = apply_end_c - apply_str_c + 1;
        for (var i = apply_str_r; i <= apply_end_r; i += 1) {
            if (hiddenRows.has("".concat(i)))
                continue;
            var copyD = copyData[i - apply_str_r];
            var applyData = getApplyData(copyD, csLen, asLen);
            if (direction === "right") {
                for (var j = apply_str_c; j <= apply_end_c; j += 1) {
                    if (hiddenCols.has("".concat(j)))
                        continue;
                    var cell = applyData[j - apply_str_c];
                    if ((cell === null || cell === void 0 ? void 0 : cell.f) != null) {
                        var f = "=".concat(formula.functionCopy(ctx, cell.f, "right", j - apply_str_c + 1));
                        var v = formula.execfunction(ctx, f, i, j);
                        formula.execFunctionGroup(ctx, j, i, v[1], undefined, d);
                        cell.v = v[1], cell.f = v[2];
                        if (cell.spl != null) {
                            cell.spl = v[3].data;
                        }
                        else if (cell.v != null) {
                            if (isRealNum(cell.v) &&
                                !/^\d{6}(18|19|20)?\d{2}(0[1-9]|1[12])(0[1-9]|[12]\d|3[01])\d{3}(\d|X)$/i.test("".concat(cell.v))) {
                                if (cell.v === Infinity || cell.v === -Infinity) {
                                    cell.m = cell.v.toString();
                                }
                                else {
                                    if (cell.v.toString().indexOf("e") > -1) {
                                        var len = cell.v
                                            .toString()
                                            .split(".")[1]
                                            .split("e")[0].length;
                                        if (len > 5) {
                                            len = 5;
                                        }
                                        cell.m = cell.v.toExponential(len).toString();
                                    }
                                    else {
                                        var mask = genarate(Math.round(cell.v * 1000000000) / 1000000000);
                                        cell.m = mask[0].toString();
                                    }
                                }
                                cell.ct = { fa: "General", t: "n" };
                            }
                            else {
                                var mask = genarate(cell.v);
                                cell.m = mask[0].toString();
                                _c = mask, cell.ct = _c[1];
                            }
                        }
                    }
                    d[i][j] = cell || null;
                    // 边框
                    var bd_r = i;
                    var bd_c = copy_str_c + ((j - apply_str_c) % csLen);
                    if (borderInfoCompute["".concat(bd_r, "_").concat(bd_c)]) {
                        var bd_obj = {
                            rangeType: "cell",
                            value: {
                                row_index: i,
                                col_index: j,
                                l: borderInfoCompute["".concat(bd_r, "_").concat(bd_c)].l,
                                r: borderInfoCompute["".concat(bd_r, "_").concat(bd_c)].r,
                                t: borderInfoCompute["".concat(bd_r, "_").concat(bd_c)].t,
                                b: borderInfoCompute["".concat(bd_r, "_").concat(bd_c)].b,
                            },
                        };
                        cfg.borderInfo.push(bd_obj);
                    }
                    else if (borderInfoCompute["".concat(i, "_").concat(j)]) {
                        var bd_obj = {
                            rangeType: "cell",
                            value: {
                                row_index: i,
                                col_index: j,
                                l: null,
                                r: null,
                                t: null,
                                b: null,
                            },
                        };
                        cfg.borderInfo.push(bd_obj);
                    }
                    // 数据验证
                    if (dataVerification != null && dataVerification["".concat(bd_r, "_").concat(bd_c)]) {
                        dataVerification["".concat(i, "_").concat(j)] = dataVerification["".concat(bd_r, "_").concat(bd_c)];
                    }
                }
            }
            if (direction === "left") {
                for (var j = apply_end_c; j >= apply_str_c; j -= 1) {
                    if (hiddenCols.has("".concat(j)))
                        continue;
                    var cell = applyData[apply_end_c - j];
                    if ((cell === null || cell === void 0 ? void 0 : cell.f) != null) {
                        var f = "=".concat(formula.functionCopy(ctx, cell.f, "left", apply_end_c - j + 1));
                        var v = formula.execfunction(ctx, f, i, j);
                        formula.execFunctionGroup(ctx, j, i, v[1], undefined, d);
                        cell.v = v[1], cell.f = v[2];
                        if (cell.spl != null) {
                            cell.spl = v[3].data;
                        }
                        else if (cell.v != null) {
                            if (isRealNum(cell.v) &&
                                !/^\d{6}(18|19|20)?\d{2}(0[1-9]|1[12])(0[1-9]|[12]\d|3[01])\d{3}(\d|X)$/i.test("".concat(cell.v))) {
                                if (cell.v === Infinity || cell.v === -Infinity) {
                                    cell.m = cell.v.toString();
                                }
                                else {
                                    if (cell.v.toString().indexOf("e") > -1) {
                                        var len = cell.v
                                            .toString()
                                            .split(".")[1]
                                            .split("e")[0].length;
                                        if (len > 5) {
                                            len = 5;
                                        }
                                        cell.m = cell.v.toExponential(len).toString();
                                    }
                                    else {
                                        var mask = genarate(Math.round(cell.v * 1000000000) / 1000000000);
                                        cell.m = mask[0].toString();
                                    }
                                }
                                cell.ct = { fa: "General", t: "n" };
                            }
                            else {
                                var mask = genarate(cell.v);
                                cell.m = mask[0].toString();
                                _d = mask, cell.ct = _d[1];
                            }
                        }
                    }
                    d[i][j] = cell || null;
                    // 边框
                    var bd_r = i;
                    var bd_c = copy_end_c - ((apply_end_c - j) % csLen);
                    if (borderInfoCompute["".concat(bd_r, "_").concat(bd_c)]) {
                        var bd_obj = {
                            rangeType: "cell",
                            value: {
                                row_index: i,
                                col_index: j,
                                l: borderInfoCompute["".concat(bd_r, "_").concat(bd_c)].l,
                                r: borderInfoCompute["".concat(bd_r, "_").concat(bd_c)].r,
                                t: borderInfoCompute["".concat(bd_r, "_").concat(bd_c)].t,
                                b: borderInfoCompute["".concat(bd_r, "_").concat(bd_c)].b,
                            },
                        };
                        cfg.borderInfo.push(bd_obj);
                    }
                    else if (borderInfoCompute["".concat(i, "_").concat(j)]) {
                        var bd_obj = {
                            rangeType: "cell",
                            value: {
                                row_index: i,
                                col_index: j,
                                l: null,
                                r: null,
                                t: null,
                                b: null,
                            },
                        };
                        cfg.borderInfo.push(bd_obj);
                    }
                    // 数据验证
                    if (dataVerification != null && dataVerification["".concat(bd_r, "_").concat(bd_c)]) {
                        dataVerification["".concat(i, "_").concat(j)] = dataVerification["".concat(bd_r, "_").concat(bd_c)];
                    }
                }
            }
        }
    }
    // 条件格式
    var cdformat = file.luckysheet_conditionformat_save;
    if (cdformat != null && cdformat.length > 0) {
        for (var i = 0; i < cdformat.length; i += 1) {
            var cdformat_cellrange = cdformat[i].cellrange;
            var emptyRange = [];
            for (var j = 0; j < cdformat_cellrange.length; j += 1) {
                var range = CFSplitRange(cdformat_cellrange[j], { row: copyRange.row, column: copyRange.column }, { row: applyRange.row, column: applyRange.column }, "operatePart");
                if (range.length > 0) {
                    emptyRange = emptyRange.concat(range);
                }
            }
            if (emptyRange.length > 0) {
                cdformat[i].cellrange.push(applyRange);
            }
        }
    }
    // 刷新一次表格
    // const allParam = {
    //   cfg,
    //   cdformat,
    //   dataVerification,
    // };
    jfrefreshgrid(ctx, d, ctx.luckysheet_select_save);
    // selectHightlightShow();
}
export function onDropCellSelectEnd(ctx, e, container) {
    var _a, _b, _c;
    ctx.luckysheet_cell_selected_extend = false;
    hideDropCellSelection(container);
    // if (
    //   !checkProtectionLockedRangeList(
    //     ctx.luckysheet_select_save,
    //     ctx.currentSheetId
    //   )
    // ) {
    //   return;
    // }
    var scrollLeft = ctx.scrollLeft, scrollTop = ctx.scrollTop;
    var rect = container.getBoundingClientRect();
    var x = e.pageX - rect.left - ctx.rowHeaderWidth + scrollLeft;
    var y = e.pageY - rect.top - ctx.columnHeaderHeight + scrollTop;
    var row_location = rowLocation(y, ctx.visibledatarow);
    // const row = row_location[1];
    var row_pre = row_location[0];
    var row_index = row_location[2];
    var col_location = colLocation(x, ctx.visibledatacolumn);
    // const col = col_location[1];
    var col_pre = col_location[0];
    var col_index = col_location[2];
    var row_index_original = ctx.luckysheet_cell_selected_extend_index[0];
    var col_index_original = ctx.luckysheet_cell_selected_extend_index[1];
    var last = (_a = ctx.luckysheet_select_save) === null || _a === void 0 ? void 0 : _a[ctx.luckysheet_select_save.length - 1];
    if (last &&
        last.top != null &&
        last.left != null &&
        last.height != null &&
        last.width != null &&
        last.row_focus != null &&
        last.column_focus != null) {
        var row_s = last.row[0];
        var row_e = last.row[1];
        var col_s = last.column[0];
        var col_e = last.column[1];
        // 复制范围
        dropCellCache.copyRange = _.cloneDeep(_.pick(last, ["row", "column"]));
        // applyType
        var typeItemHide = getTypeItemHide(ctx);
        if (!typeItemHide[0] &&
            !typeItemHide[1] &&
            !typeItemHide[2] &&
            !typeItemHide[3] &&
            !typeItemHide[4] &&
            !typeItemHide[5] &&
            !typeItemHide[6]) {
            dropCellCache.applyType = "0";
        }
        else {
            dropCellCache.applyType = "1";
        }
        if (ctx.luckysheet_select_save == null)
            return;
        var _d = ctx.luckysheet_select_save[0], top_move = _d.top_move, left_move = _d.left_move;
        if (Math.abs(row_index_original - row_index) >
            Math.abs(col_index_original - col_index)) {
            if (!(row_index >= row_s && row_index <= row_e)) {
                if (top_move != null && top_move >= row_pre) {
                    // 当往上拖拽时
                    dropCellCache.applyRange = {
                        row: [row_index, last.row[0] - 1],
                        column: last.column,
                    };
                    dropCellCache.direction = "up";
                    row_s -= last.row[0] - row_index;
                    // 是否有数据透视表范围
                    // if (pivotTable.isPivotRange(row_s, col_e)) {
                    //   tooltip.info(locale_drag.affectPivot, "");
                    //   return;
                    // }
                }
                else {
                    // 当往下拖拽时
                    dropCellCache.applyRange = {
                        row: [last.row[1] + 1, row_index],
                        column: last.column,
                    };
                    dropCellCache.direction = "down";
                    row_e += row_index - last.row[1];
                    // 是否有数据透视表范围
                    // if (pivotTable.isPivotRange(row_e, col_e)) {
                    //   tooltip.info(locale_drag.affectPivot, "");
                    //   return;
                    // }
                }
            }
            else {
                return;
            }
        }
        else {
            if (!(col_index >= col_s && col_index <= col_e)) {
                if (left_move != null && left_move >= col_pre) {
                    // 当往左拖拽时
                    dropCellCache.applyRange = {
                        row: last.row,
                        column: [col_index, last.column[0] - 1],
                    };
                    dropCellCache.direction = "left";
                    col_s -= last.column[0] - col_index;
                    // 是否有数据透视表范围
                    // if (pivotTable.isPivotRange(row_e, col_s)) {
                    //   tooltip.info(locale_drag.affectPivot, "");
                    //   return;
                    // }
                }
                else {
                    // 当往右拖拽时
                    dropCellCache.applyRange = {
                        row: last.row,
                        column: [last.column[1] + 1, col_index],
                    };
                    dropCellCache.direction = "right";
                    col_e += col_index - last.column[1];
                    // 是否有数据透视表范围
                    // if (pivotTable.isPivotRange(row_e, col_e)) {
                    //   tooltip.info(locale_drag.affectPivot, "");
                    //   return;
                    // }
                }
            }
            else {
                return;
            }
        }
        if (y < 0) {
            row_s = 0;
            row_e = last.row[0];
        }
        if (x < 0) {
            col_s = 0;
            col_e = last.column[0];
        }
        var flowdata = getFlowdata(ctx);
        if (flowdata == null)
            return;
        if (ctx.config.merge != null) {
            var HasMC = false;
            for (var r = last.row[0]; r <= last.row[1]; r += 1) {
                for (var c = last.column[0]; c <= last.column[1]; c += 1) {
                    var cell = (_b = flowdata[r]) === null || _b === void 0 ? void 0 : _b[c];
                    if (cell != null && cell.mc != null) {
                        HasMC = true;
                        break;
                    }
                }
            }
            if (HasMC) {
                // if (isEditMode()) {
                //   alert(locale_drag.noMerge);
                // } else {
                //   tooltip.info(locale_drag.noMerge, "");
                // }
                return;
            }
            for (var r = row_s; r <= row_e; r += 1) {
                for (var c = col_s; c <= col_e; c += 1) {
                    var cell = (_c = flowdata[r]) === null || _c === void 0 ? void 0 : _c[c];
                    if (cell != null && cell.mc != null) {
                        HasMC = true;
                        break;
                    }
                }
            }
            if (HasMC) {
                // if (isEditMode()) {
                //   alert(locale_drag.noMerge);
                // } else {
                //   tooltip.info(locale_drag.noMerge, "");
                // }
                return;
            }
        }
        last.row = [row_s, row_e];
        last.column = [col_s, col_e];
        ctx.luckysheet_select_save = normalizeSelection(ctx, [
            {
                row: [row_s, row_e],
                column: [col_s, col_e],
            },
        ]);
        try {
            updateDropCell(ctx);
        }
        catch (err) {
            console.error(err);
        }
        // createIcon();
        var selectedMoveEle = container.querySelector(".fortune-cell-selected-move");
        if (selectedMoveEle) {
            selectedMoveEle.style.display = "none";
        }
        // clearTimeout(ctx.countfuncTimeout);
        // ctx.countfuncTimeout = setTimeout(() => {
        // countfunc();
        // }, 500);
    }
}
