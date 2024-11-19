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
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
import _ from "lodash";
// @ts-ignore
import { Parser, ERROR_REF } from "@fortune-sheet/formula-parser";
import { getFlowdata } from "../context";
import { columnCharToIndex, escapeScriptTag, getSheetIndex, indexToColumnChar, getSheetIdByName, escapeHTMLTag, } from "../utils";
import { getcellFormula, getRangetxt, mergeMoveMain, setCellValue, } from "./cell";
import { error } from "./validation";
import { moveToEnd } from "./cursor";
import { locale } from "../locale";
import { colors } from "./color";
import { colLocation, mousePosition, rowLocation } from "./location";
import { cancelFunctionrangeSelected, seletedHighlistByindex } from ".";
var functionHTMLIndex = 0;
var rangeIndexes = [];
var operatorPriority = {
    "^": 0,
    "%": 1,
    "*": 1,
    "/": 1,
    "+": 2,
    "-": 2,
};
var operatorArr = "==|!=|<>|<=|>=|=|+|-|>|<|/|*|%|&|^".split("|");
var operatorjson = {};
for (var i = 0; i < operatorArr.length; i += 1) {
    operatorjson[operatorArr[i].toString()] = 1;
}
var simpleSheetName = "[A-Za-z0-9_\u00C0-\u02AF]+";
var quotedSheetName = "'(?:(?!').|'')*'";
var sheetNameRegexp = "(".concat(simpleSheetName, "|").concat(quotedSheetName, ")!");
var rowColumnRegexp = "[$]?[A-Za-z]+[$]?[0-9]+";
var rowColumnWithSheetName = "(?:".concat(sheetNameRegexp, ")?(").concat(rowColumnRegexp, ")");
var LABEL_EXTRACT_REGEXP = new RegExp("^".concat(rowColumnWithSheetName, "(?:[:]").concat(rowColumnWithSheetName, ")?$"));
// FormulaCache is defined as class to avoid being frozen by immer
var FormulaCache = /** @class */ (function () {
    function FormulaCache() {
        var that = this;
        this.data_parm_index = 0;
        this.selectingRangeIndex = -1;
        this.functionlistMap = {};
        this.execFunctionGlobalData = {};
        this.cellTextToIndexList = {};
        this.parser = new Parser();
        this.parser.on("callCellValue", function (cellCoord, options, done) {
            var _a, _b;
            var context = that.parser.context;
            var id = cellCoord.sheetName == null
                ? options.sheetId
                : getSheetIdByName(context, cellCoord.sheetName);
            if (id == null)
                throw Error(ERROR_REF);
            var flowdata = getFlowdata(context, id);
            var cell = ((_a = context === null || context === void 0 ? void 0 : context.formulaCache.execFunctionGlobalData) === null || _a === void 0 ? void 0 : _a["".concat(cellCoord.row.index, "_").concat(cellCoord.column.index, "_").concat(id)]) || ((_b = flowdata === null || flowdata === void 0 ? void 0 : flowdata[cellCoord.row.index]) === null || _b === void 0 ? void 0 : _b[cellCoord.column.index]);
            var v = that.tryGetCellAsNumber(cell);
            done(v);
        });
        this.parser.on("callRangeValue", function (startCellCoord, endCellCoord, options, done) {
            var _a, _b, _c, _d;
            var context = that.parser.context;
            var id = startCellCoord.sheetName == null
                ? options.sheetId
                : getSheetIdByName(context, startCellCoord.sheetName);
            if (id == null)
                throw Error(ERROR_REF);
            var flowdata = getFlowdata(context, id);
            var fragment = [];
            var startRow = startCellCoord.row.index;
            var endRow = endCellCoord.row.index;
            var startCol = startCellCoord.column.index;
            var endCol = endCellCoord.column.index;
            var emptyRow = startRow === -1 || endRow === -1;
            var emptyCol = startCol === -1 || endCol === -1;
            if (emptyRow) {
                startRow = 0;
                endRow = (_a = flowdata === null || flowdata === void 0 ? void 0 : flowdata.length) !== null && _a !== void 0 ? _a : 0;
            }
            if (emptyCol) {
                startCol = 0;
                endCol = (_b = flowdata === null || flowdata === void 0 ? void 0 : flowdata[0].length) !== null && _b !== void 0 ? _b : 0;
            }
            if (emptyRow && emptyCol)
                throw Error(ERROR_REF);
            for (var row = startRow; row <= endRow; row += 1) {
                var colFragment = [];
                for (var col = startCol; col <= endCol; col += 1) {
                    var cell = ((_c = context === null || context === void 0 ? void 0 : context.formulaCache.execFunctionGlobalData) === null || _c === void 0 ? void 0 : _c["".concat(row, "_").concat(col, "_").concat(id)]) || ((_d = flowdata === null || flowdata === void 0 ? void 0 : flowdata[row]) === null || _d === void 0 ? void 0 : _d[col]);
                    var v = that.tryGetCellAsNumber(cell);
                    colFragment.push(v);
                }
                fragment.push(colFragment);
            }
            if (fragment) {
                done(fragment);
            }
        });
    }
    FormulaCache.prototype.tryGetCellAsNumber = function (cell) {
        var _a;
        if (((_a = cell === null || cell === void 0 ? void 0 : cell.ct) === null || _a === void 0 ? void 0 : _a.t) === "n") {
            var n = Number(cell === null || cell === void 0 ? void 0 : cell.v);
            return Number.isNaN(n) ? cell.v : n;
        }
        return cell === null || cell === void 0 ? void 0 : cell.v;
    };
    return FormulaCache;
}());
export { FormulaCache };
function parseElement(eleString) {
    return new DOMParser().parseFromString(eleString, "text/html").body
        .childNodes[0];
}
export function iscelldata(txt) {
    // 判断是否为单元格格式
    var val = txt.split("!");
    var rangetxt;
    if (val.length > 1) {
        rangetxt = val[1];
    }
    else {
        rangetxt = val[0];
    }
    var reg_cell = /^(([a-zA-Z]+)|([$][a-zA-Z]+))(([0-9]+)|([$][0-9]+))$/g; // 增加正则判断单元格为字母+数字的格式：如 A1:B3
    var reg_cellRange = /^(((([a-zA-Z]+)|([$][a-zA-Z]+))(([0-9]+)|([$][0-9]+)))|((([a-zA-Z]+)|([$][a-zA-Z]+))))$/g; // 增加正则判断单元格为字母+数字或字母的格式：如 A1:B3，A:A
    if (rangetxt.indexOf(":") === -1) {
        var row_1 = parseInt(rangetxt.replace(/[^0-9]/g, ""), 10) - 1;
        var col_1 = columnCharToIndex(rangetxt.replace(/[^A-Za-z]/g, ""));
        if (!Number.isNaN(row_1) &&
            !Number.isNaN(col_1) &&
            rangetxt.toString().match(reg_cell)) {
            return true;
        }
        if (!Number.isNaN(row_1)) {
            return false;
        }
        if (!Number.isNaN(col_1)) {
            return false;
        }
        return false;
    }
    reg_cellRange =
        /^(((([a-zA-Z]+)|([$][a-zA-Z]+))(([0-9]+)|([$][0-9]+)))|((([a-zA-Z]+)|([$][a-zA-Z]+)))|((([0-9]+)|([$][0-9]+s))))$/g;
    var rangetxtArr = rangetxt.split(":");
    var row = [];
    var col = [];
    row[0] = parseInt(rangetxtArr[0].replace(/[^0-9]/g, ""), 10) - 1;
    row[1] = parseInt(rangetxtArr[1].replace(/[^0-9]/g, ""), 10) - 1;
    if (row[0] > row[1]) {
        return false;
    }
    col[0] = columnCharToIndex(rangetxtArr[0].replace(/[^A-Za-z]/g, ""));
    col[1] = columnCharToIndex(rangetxtArr[1].replace(/[^A-Za-z]/g, ""));
    if (col[0] > col[1]) {
        return false;
    }
    if (rangetxtArr[0].toString().match(reg_cellRange) &&
        rangetxtArr[1].toString().match(reg_cellRange)) {
        return true;
    }
    return false;
}
function addToCellIndexList(ctx, txt, infoObj) {
    if (_.isNil(txt) || txt.length === 0 || _.isNil(infoObj)) {
        return;
    }
    if (_.isNil(ctx.formulaCache.cellTextToIndexList)) {
        ctx.formulaCache.cellTextToIndexList = {};
    }
    if (txt.indexOf("!") > -1) {
        txt = txt.replace(/\\'/g, "'").replace(/''/g, "'");
        ctx.formulaCache.cellTextToIndexList[txt] = infoObj;
    }
    else {
        ctx.formulaCache.cellTextToIndexList["".concat(txt, "_").concat(infoObj.sheetId)] = infoObj;
    }
}
export function getcellrange(ctx, txt, formulaId) {
    if (_.isNil(txt) || txt.length === 0) {
        return null;
    }
    var flowdata = getFlowdata(ctx, formulaId);
    var sheettxt = "";
    var rangetxt = "";
    var sheetId = null;
    var sheetdata = null;
    var luckysheetfile = ctx.luckysheetfile;
    if (txt.indexOf("!") > -1) {
        if (txt in ctx.formulaCache.cellTextToIndexList) {
            return ctx.formulaCache.cellTextToIndexList[txt];
        }
        var matchRes = txt.match(LABEL_EXTRACT_REGEXP);
        if (matchRes == null) {
            return null;
        }
        var sheettxt1 = matchRes[1], starttxt1 = matchRes[2], sheettxt2 = matchRes[3], starttxt2 = matchRes[4];
        if (sheettxt2 != null && sheettxt1 !== sheettxt2) {
            return null;
        }
        rangetxt = starttxt2 ? "".concat(starttxt1, ":").concat(starttxt2) : starttxt1;
        sheettxt = sheettxt1
            .replace(/^'|'$/g, "")
            .replace(/\\'/g, "'")
            .replace(/''/g, "'");
        _.forEach(luckysheetfile, function (f) {
            if (sheettxt === f.name) {
                sheetId = f.id;
                sheetdata = f.data;
                return false;
            }
            return true;
        });
    }
    else {
        var i = formulaId;
        if (_.isNil(i)) {
            i = ctx.currentSheetId;
        }
        if ("".concat(txt, "_").concat(i) in ctx.formulaCache.cellTextToIndexList) {
            return ctx.formulaCache.cellTextToIndexList["".concat(txt, "_").concat(i)];
        }
        var index = getSheetIndex(ctx, i);
        if (_.isNil(index)) {
            return null;
        }
        sheettxt = luckysheetfile[index].name;
        sheetId = luckysheetfile[index].id;
        sheetdata = flowdata;
        rangetxt = txt;
    }
    if (_.isNil(sheetdata)) {
        return null;
    }
    if (rangetxt.indexOf(":") === -1) {
        var row_2 = parseInt(rangetxt.replace(/[^0-9]/g, ""), 10) - 1;
        var col_2 = columnCharToIndex(rangetxt.replace(/[^A-Za-z]/g, ""));
        if (!Number.isNaN(row_2) && !Number.isNaN(col_2)) {
            var item_1 = {
                row: [row_2, row_2],
                column: [col_2, col_2],
                sheetId: sheetId,
            };
            addToCellIndexList(ctx, txt, item_1);
            return item_1;
        }
        return null;
    }
    var rangetxtArr = rangetxt.split(":");
    var row = [];
    var col = [];
    row[0] = parseInt(rangetxtArr[0].replace(/[^0-9]/g, ""), 10) - 1;
    row[1] = parseInt(rangetxtArr[1].replace(/[^0-9]/g, ""), 10) - 1;
    if (Number.isNaN(row[0])) {
        row[0] = 0;
    }
    if (Number.isNaN(row[1])) {
        row[1] = sheetdata.length - 1;
    }
    if (row[0] > row[1]) {
        return null;
    }
    col[0] = columnCharToIndex(rangetxtArr[0].replace(/[^A-Za-z]/g, ""));
    col[1] = columnCharToIndex(rangetxtArr[1].replace(/[^A-Za-z]/g, ""));
    if (Number.isNaN(col[0])) {
        col[0] = 0;
    }
    if (Number.isNaN(col[1])) {
        col[1] = sheetdata[0].length - 1;
    }
    if (col[0] > col[1]) {
        return null;
    }
    var item = {
        row: row,
        column: col,
        sheetId: sheetId,
    };
    addToCellIndexList(ctx, txt, item);
    return item;
}
function calPostfixExpression(cal) {
    if (cal.length === 0) {
        return "";
    }
    var stack = [];
    for (var i = cal.length - 1; i >= 0; i -= 1) {
        var c = cal[i];
        if (c in operatorjson) {
            var s2 = stack.pop();
            var s1 = stack.pop();
            var str = "luckysheet_compareWith(".concat(s1, ",'").concat(c, "', ").concat(s2, ")");
            stack.push(str);
        }
        else {
            stack.push(c);
        }
    }
    if (stack.length > 0) {
        return stack[0];
    }
    return "";
}
function checkSpecialFunctionRange(ctx, function_str, r, c, id, dynamicArray_compute, cellRangeFunction) {
    if (function_str.substring(0, 30) === "luckysheet_getSpecialReference" ||
        function_str.substring(0, 20) === "luckysheet_function.") {
        if (function_str.substring(0, 20) === "luckysheet_function.") {
            var funcName = function_str.split(".")[1];
            if (!_.isNil(funcName)) {
                funcName = funcName.toUpperCase();
                if (funcName !== "INDIRECT" &&
                    funcName !== "OFFSET" &&
                    funcName !== "INDEX") {
                    return;
                }
            }
        }
        try {
            ctx.calculateSheetId = id;
            var str = function_str
                .split(",")[function_str.split(",").length - 1].split("'")[1]
                .split("'")[0];
            var str_nb = _.trim(str);
            // console.log(function_str, tempFunc,str, this.iscelldata(str_nb),this.isFunctionRangeSave,r,c);
            if (iscelldata(str_nb)) {
                if (typeof cellRangeFunction === "function") {
                    cellRangeFunction(str_nb);
                }
                // this.isFunctionRangeSaveChange(str, r, c, index, dynamicArray_compute);
                // console.log(function_str, str, this.isFunctionRangeSave,r,c);
            }
        }
        catch (_a) { }
    }
}
function isFunctionRange(ctx, txt, r, c, id, dynamicArray_compute, cellRangeFunction) {
    var _a;
    if (txt.substring(0, 1) === "=") {
        txt = txt.substring(1);
    }
    var funcstack = txt.split("");
    var i = 0;
    var str = "";
    var function_str = "";
    var matchConfig = {
        bracket: 0,
        comma: 0,
        squote: 0,
        dquote: 0,
        compare: 0,
        braces: 0,
    };
    // let luckysheetfile = getluckysheetfile();
    // let dynamicArray_compute = luckysheetfile[getSheetIndex(Store.currentSheetId)_.isNil(]["dynamicArray_compute"]) ? {} : luckysheetfile[getSheetIndex(Store.currentSheetId)]["dynamicArray_compute"];
    // bracket 0为运算符括号、1为函数括号
    var cal1 = [];
    var cal2 = [];
    var bracket = [];
    var firstSQ = -1;
    while (i < funcstack.length) {
        var s = funcstack[i];
        if (s === "(" &&
            matchConfig.squote === 0 &&
            matchConfig.dquote === 0 &&
            matchConfig.braces === 0) {
            if (str.length > 0 && bracket.length === 0) {
                str = str.toUpperCase();
                if (str.indexOf(":") > -1) {
                    var funcArray = str.split(":");
                    function_str += "luckysheet_getSpecialReference(true,'".concat(_.trim(funcArray[0]).replace(/'/g, "\\'"), "', luckysheet_function.").concat(funcArray[1], ".f(#lucky#");
                }
                else {
                    function_str += "luckysheet_function.".concat(str, ".f(");
                }
                bracket.push(1);
                str = "";
            }
            else if (bracket.length === 0) {
                function_str += "(";
                bracket.push(0);
                str = "";
            }
            else {
                bracket.push(0);
                str += s;
            }
        }
        else if (s === ")" &&
            matchConfig.squote === 0 &&
            matchConfig.dquote === 0 &&
            matchConfig.braces === 0) {
            bracket.pop();
            if (bracket.length === 0) {
                // function_str += _this.isFunctionRange(str,r,c, index,dynamicArray_compute,cellRangeFunction) + ")";
                // str = "";
                var functionS = isFunctionRange(ctx, str, r, c, id, dynamicArray_compute, cellRangeFunction);
                if (functionS.indexOf("#lucky#") > -1) {
                    functionS = "".concat(functionS.replace(/#lucky#/g, ""), ")");
                }
                function_str += "".concat(functionS, ")");
                str = "";
            }
            else {
                str += s;
            }
        }
        else if (s === "{" &&
            matchConfig.squote === 0 &&
            matchConfig.dquote === 0) {
            str += "{";
            matchConfig.braces += 1;
        }
        else if (s === "}" &&
            matchConfig.squote === 0 &&
            matchConfig.dquote === 0) {
            str += "}";
            matchConfig.braces -= 1;
        }
        else if (s === '"' && matchConfig.squote === 0) {
            if (matchConfig.dquote > 0) {
                // 如果是""代表着输出"
                if (i < funcstack.length - 1 && funcstack[i + 1] === '"') {
                    i += 1;
                    str += "\x7F"; // 用DEL替换一下""
                }
                else {
                    matchConfig.dquote -= 1;
                    str += '"';
                }
            }
            else {
                matchConfig.dquote += 1;
                str += '"';
            }
        }
        else if (s === "'" && matchConfig.dquote === 0) {
            str += "'";
            if (matchConfig.squote > 0) {
                // if (firstSQ === i - 1)//配对的单引号后第一个字符不能是单引号
                // {
                //    代码到了此处应该是公式错误
                // }
                // 如果是''代表着输出'
                if (i < funcstack.length - 1 && funcstack[i + 1] === "'") {
                    i += 1;
                    str += "'";
                }
                else {
                    // 如果下一个字符不是'代表单引号结束
                    // if (funcstack[i - 1] === "'") {//配对的单引号后最后一个字符不能是单引号
                    //    代码到了此处应该是公式错误
                    // } else {
                    matchConfig.squote -= 1;
                    // }
                }
            }
            else {
                matchConfig.squote += 1;
                // eslint-disable-next-line @typescript-eslint/no-unused-vars
                firstSQ = i;
            }
        }
        else if (s === "," &&
            matchConfig.squote === 0 &&
            matchConfig.dquote === 0 &&
            matchConfig.braces === 0) {
            if (bracket.length <= 1) {
                // function_str += _this.isFunctionRange(str, r, c, index,dynamicArray_compute,cellRangeFunction) + ",";
                // str = "";
                var functionS = isFunctionRange(ctx, str, r, c, id, dynamicArray_compute, cellRangeFunction);
                if (functionS.indexOf("#lucky#") > -1) {
                    functionS = "".concat(functionS.replace(/#lucky#/g, ""), ")");
                }
                function_str += "".concat(functionS, ",");
                str = "";
            }
            else {
                str += ",";
            }
        }
        else if (s in operatorjson &&
            matchConfig.squote === 0 &&
            matchConfig.dquote === 0 &&
            matchConfig.braces === 0) {
            var s_next = "";
            var op = operatorPriority;
            if (i + 1 < funcstack.length) {
                s_next = funcstack[i + 1];
            }
            if (s + s_next in operatorjson) {
                if (bracket.length === 0) {
                    if (_.trim(str).length > 0) {
                        cal2.unshift(isFunctionRange(ctx, _.trim(str), r, c, id, dynamicArray_compute, cellRangeFunction));
                    }
                    else if (_.trim(function_str).length > 0) {
                        cal2.unshift(_.trim(function_str));
                    }
                    if (cal1[0] in operatorjson) {
                        var stackCeilPri = op[cal1[0]];
                        while (cal1.length > 0 && !_.isNil(stackCeilPri)) {
                            cal2.unshift(cal1.shift());
                            stackCeilPri = op[cal1[0]];
                        }
                    }
                    cal1.unshift(s + s_next);
                    function_str = "";
                    str = "";
                }
                else {
                    str += s + s_next;
                }
                i += 1;
            }
            else {
                if (bracket.length === 0) {
                    if (_.trim(str).length > 0) {
                        cal2.unshift(isFunctionRange(ctx, _.trim(str), r, c, id, dynamicArray_compute, cellRangeFunction));
                    }
                    else if (_.trim(function_str).length > 0) {
                        cal2.unshift(_.trim(function_str));
                    }
                    if (cal1[0] in operatorjson) {
                        var stackCeilPri = op[cal1[0]];
                        stackCeilPri = _.isNil(stackCeilPri) ? 1000 : stackCeilPri;
                        var sPri = op[s];
                        sPri = _.isNil(sPri) ? 1000 : sPri;
                        while (cal1.length > 0 && sPri >= stackCeilPri) {
                            cal2.unshift(cal1.shift());
                            stackCeilPri = op[cal1[0]];
                            stackCeilPri = _.isNil(stackCeilPri) ? 1000 : stackCeilPri;
                        }
                    }
                    cal1.unshift(s);
                    function_str = "";
                    str = "";
                }
                else {
                    str += s;
                }
            }
        }
        else {
            if (matchConfig.dquote === 0 && matchConfig.squote === 0) {
                str += _.trim(s);
            }
            else {
                str += s;
            }
        }
        if (i === funcstack.length - 1) {
            var endstr = "";
            var str_nb = _.trim(str).replace(/'/g, "\\'");
            if (iscelldata(str_nb) && str_nb.substring(0, 1) !== ":") {
                // endstr = "luckysheet_getcelldata('" + _.trim(str) + "')";
                endstr = "luckysheet_getcelldata('".concat(str_nb, "')");
            }
            else if (str_nb.substring(0, 1) === ":") {
                str_nb = str_nb.substring(1);
                if (iscelldata(str_nb)) {
                    endstr = "luckysheet_getSpecialReference(false,".concat(function_str, ",'").concat(str_nb, "')");
                }
            }
            else {
                str = _.trim(str);
                var regx = /{.*?}/;
                if (regx.test(str) &&
                    str.substring(0, 1) !== '"' &&
                    str.substring(str.length - 1, 1) !== '"') {
                    var arraytxt = (_a = regx.exec(str)) === null || _a === void 0 ? void 0 : _a[0];
                    var arraystart = str.search(regx);
                    if (arraystart > 0) {
                        endstr += str.substring(0, arraystart);
                    }
                    endstr += "luckysheet_getarraydata('".concat(arraytxt, "')");
                    if (arraystart + arraytxt.length < str.length) {
                        endstr += str.substring(arraystart + arraytxt.length, str.length);
                    }
                }
                else {
                    endstr = str;
                }
            }
            if (endstr.length > 0) {
                cal2.unshift(endstr);
            }
            if (cal1.length > 0) {
                if (function_str.length > 0) {
                    cal2.unshift(function_str);
                    function_str = "";
                }
                while (cal1.length > 0) {
                    cal2.unshift(cal1.shift());
                }
            }
            if (cal2.length > 0) {
                function_str = calPostfixExpression(cal2);
            }
            else {
                function_str += endstr;
            }
        }
        i += 1;
    }
    checkSpecialFunctionRange(ctx, function_str, r, c, id, dynamicArray_compute, cellRangeFunction);
    return function_str;
}
export function getAllFunctionGroup(ctx) {
    var luckysheetfile = ctx.luckysheetfile;
    var ret = [];
    for (var i = 0; i < luckysheetfile.length; i += 1) {
        var file = luckysheetfile[i];
        var calcChain = file.calcChain;
        /* 备注：再次加载表格获取的数据可能是JSON字符串格式(需要进行发序列化处理) */
        // if (calcChain) {
        //   const tempCalcChain: any[] = [];
        //   calcChain.forEach((item) => {
        //     if (typeof item === "string") {
        //       tempCalcChain.push(JSON.parse(item));
        //     } else {
        //       tempCalcChain.push(item);
        //     }
        //   });
        //   calcChain = tempCalcChain;
        //   file.calcChain = tempCalcChain;
        // }
        var dynamicArray_compute = file.dynamicArray_compute;
        if (_.isNil(calcChain)) {
            calcChain = [];
        }
        if (_.isNil(dynamicArray_compute)) {
            dynamicArray_compute = [];
        }
        ret = ret.concat(calcChain);
        for (var j = 0; j < dynamicArray_compute.length; j += 1) {
            var d = dynamicArray_compute[0];
            ret.push({
                r: d.r,
                c: d.c,
                id: d.id,
            });
        }
    }
    return ret;
}
export function delFunctionGroup(ctx, r, c, id) {
    if (_.isNil(id)) {
        id = ctx.currentSheetId;
    }
    var file = ctx.luckysheetfile[getSheetIndex(ctx, id)];
    var calcChain = file.calcChain;
    if (!_.isNil(calcChain)) {
        var modified = false;
        var calcChainClone = _.cloneDeep(calcChain);
        for (var i = 0; i < calcChainClone.length; i += 1) {
            var calc = calcChainClone[i];
            if (calc.r === r && calc.c === c && calc.id === id) {
                calcChainClone.splice(i, 1);
                modified = true;
                // server.saveParam("fc", index, calc, {
                //   op: "del",
                //   pos: i,
                // });
                break;
            }
        }
        if (modified) {
            file.calcChain = calcChainClone;
        }
    }
    var dynamicArray = file.dynamicArray;
    if (!_.isNil(dynamicArray)) {
        var modified = false;
        var dynamicArrayClone = _.cloneDeep(dynamicArray);
        for (var i = 0; i < dynamicArrayClone.length; i += 1) {
            var calc = dynamicArrayClone[i];
            if (calc.r === r &&
                calc.c === c &&
                (_.isNil(calc.id) || calc.id === id)) {
                dynamicArrayClone.splice(i, 1);
                modified = true;
                // server.saveParam("ac", index, null, {
                //   op: "del",
                //   pos: i,
                // });
                break;
            }
        }
        if (modified) {
            file.dynamicArray = dynamicArrayClone;
        }
    }
}
function checkBracketNum(fp) {
    var bra_l = fp.match(/\(/g);
    var bra_r = fp.match(/\)/g);
    var bra_tl_txt = fp.match(/(['"])(?:(?!\1).)*?\1/g);
    var bra_tr_txt = fp.match(/(['"])(?:(?!\1).)*?\1/g);
    var bra_l_len = 0;
    var bra_r_len = 0;
    if (!_.isNil(bra_l)) {
        bra_l_len += bra_l.length;
    }
    if (!_.isNil(bra_r)) {
        bra_r_len += bra_r.length;
    }
    var bra_tl_len = 0;
    var bra_tr_len = 0;
    if (!_.isNil(bra_tl_txt)) {
        for (var i = 0; i < bra_tl_txt.length; i += 1) {
            var bra_tl = bra_tl_txt[i].match(/\(/g);
            if (!_.isNil(bra_tl)) {
                bra_tl_len += bra_tl.length;
            }
        }
    }
    if (!_.isNil(bra_tr_txt)) {
        for (var i = 0; i < bra_tr_txt.length; i += 1) {
            var bra_tr = bra_tr_txt[i].match(/\)/g);
            if (!_.isNil(bra_tr)) {
                bra_tr_len += bra_tr.length;
            }
        }
    }
    bra_l_len -= bra_tl_len;
    bra_r_len -= bra_tr_len;
    if (bra_l_len !== bra_r_len) {
        return false;
    }
    return true;
}
export function insertUpdateFunctionGroup(ctx, r, c, id, calcChainSet) {
    if (_.isNil(id)) {
        id = ctx.currentSheetId;
    }
    // let func = getcellFormula(r, c, index);
    // if (_.isNil(func) || func.length==0) {
    //     this.delFunctionGroup(r, c, index);
    //     return;
    // }
    var luckysheetfile = ctx.luckysheetfile;
    var idx = getSheetIndex(ctx, id);
    if (_.isNil(idx)) {
        return;
    }
    var file = luckysheetfile[idx];
    var calcChain = file.calcChain;
    if (_.isNil(calcChain)) {
        calcChain = [];
    }
    if (calcChainSet) {
        if (calcChainSet.has("".concat(r, "_").concat(c, "_").concat(id)))
            return;
    }
    else {
        for (var i = 0; i < calcChain.length; i += 1) {
            var calc = calcChain[i];
            if (calc.r === r && calc.c === c && calc.id === id) {
                // server.saveParam("fc", index, calc, {
                //   op: "update",
                //   pos: i,
                // });
                return;
            }
        }
    }
    var cc = {
        r: r,
        c: c,
        id: id,
    };
    calcChain.push(cc);
    file.calcChain = calcChain;
    // server.saveParam("fc", index, cc, {
    //   op: "add",
    //   pos: file.calcChain.length - 1,
    // });
    ctx.luckysheetfile = luckysheetfile;
}
export function execfunction(ctx, txt, r, c, id, calcChainSet, isrefresh, notInsertFunc) {
    if (txt.indexOf(error.r) > -1) {
        return [false, error.r, txt];
    }
    if (!checkBracketNum(txt)) {
        txt += ")";
    }
    if (_.isNil(id)) {
        id = ctx.currentSheetId;
    }
    ctx.calculateSheetId = id;
    /*
    const fp = _.trim(functionParserExe(txt));
    if (
      fp.substring(0, 20) === "luckysheet_function." ||
      fp.substring(0, 22) === "luckysheet_compareWith"
    ) {
      functionHTMLIndex = 0;
    }
  
    if (!testFunction(txt) || fp === "") {
      // TODO tooltip.info("", locale_formulaMore.execfunctionError);
      return [false, error.n, txt];
    }
  
    let result = null;
    window.luckysheetCurrentRow = r;
    window.luckysheetCurrentColumn = c;
    window.luckysheetCurrentIndex = index;
    window.luckysheetCurrentFunction = txt;
  
    let sparklines = null;
  
    try {
      if (fp.indexOf("luckysheet_getcelldata") > -1) {
        const funcg = fp.split("luckysheet_getcelldata('");
  
        for (let i = 1; i < funcg.length; i += 1) {
          const funcgStr = funcg[i].split("')")[0];
          const funcgRange = getcellrange(ctx, funcgStr);
  
          if (funcgRange.row[0] < 0 || funcgRange.column[0] < 0) {
            return [true, error.r, txt];
          }
  
          if (
            funcgRange.sheetId === ctx.calculateSheetId &&
            r >= funcgRange.row[0] &&
            r <= funcgRange.row[1] &&
            c >= funcgRange.column[0] &&
            c <= funcgRange.column[1]
          ) {
            // TODO if (isEditMode()) {
            //   alert(locale_formulaMore.execfunctionSelfError);
            // } else {
            //   tooltip.info("", locale_formulaMore.execfunctionSelfErrorResult);
            // }
  
            return [false, 0, txt];
          }
        }
      }
  
      result = new Function(`return ${fp}`)();
      if (typeof result === "string") {
        // 把之前的非打印控制字符DEL替换回一个双引号。
        result = result.replace(/\x7F/g, '"');
      }
  
      // 加入sparklines的参数项目
      if (fp.indexOf("SPLINES") > -1) {
        sparklines = result;
        result = "";
      }
    } catch (e) {
      const err = e;
      // err错误提示处理
      console.log(e, fp);
      result = [error.n, err];
    }
  
    // 公式结果是对象，则表示只是选区。如果是单个单元格，则返回其值；如果是多个单元格，则返回 #VALUE!。
    if (_.isPlainObject(result) && !_.isNil(result.startCell)) {
      if (_.isArray(result.data)) {
        result = error.v;
      } else {
        if (_.isPlainObject(result.data) && !_.isEmpty(result.data.v)) {
          result = result.data.v;
        } else if (!_.isEmpty(result.data)) {
          // 只有data长或宽大于1才可能是选区
          if (result.cell > 1 || result.rowl > 1) {
            result = result.data;
          } // 否则就是单个不为null的没有值v的单元格
          else {
            result = 0;
          }
        } else {
          result = 0;
        }
      }
    }
  
    // 公式结果是数组，分错误值 和 动态数组 两种情况
    let dynamicArrayItem = null;
  
    if (_.isArray(result)) {
      let isErr = false;
  
      if (!_.isArray(result[0]) && result.length === 2) {
        isErr = valueIsError(result[0]);
      }
  
      if (!isErr) {
        if (
          _.isArray(result[0]) &&
          result.length === 1 &&
          result[0].length === 1
        ) {
          result = result[0][0];
        } else {
          dynamicArrayItem = { r, c, f: txt, id, data: result };
          result = "";
        }
      } else {
        result = result[0];
      }
    }
  
    window.luckysheetCurrentRow = null;
    window.luckysheetCurrentColumn = null;
    window.luckysheetCurrentIndex = null;
    window.luckysheetCurrentFunction = null;
    */
    ctx.formulaCache.parser.context = ctx;
    var parsedResponse = ctx.formulaCache.parser.parse(txt.substring(1), {
        sheetId: id || ctx.currentSheetId,
    });
    var formulaError = parsedResponse.error;
    var result = parsedResponse.result;
    // https://stackoverflow.com/a/643827/8200626
    // https://github.com/ruilisi/fortune-sheet/issues/551
    if (Object.prototype.toString.call(result) === "[object Date]" &&
        !_.isNil(result)) {
        result = result.toString();
    }
    if (!_.isNil(r) && !_.isNil(c)) {
        if (isrefresh) {
            // eslint-disable-next-line no-use-before-define
            execFunctionGroup(ctx, r, c, _.isNil(formulaError) ? result : formulaError, id);
        }
        if (!notInsertFunc) {
            insertUpdateFunctionGroup(ctx, r, c, id, calcChainSet);
        }
    }
    /*
    if (sparklines) {
      return [true, result, txt, { type: "sparklines", data: sparklines }];
    }
  
    if (dynamicArrayItem) {
      return [
        true,
        result,
        txt,
        { type: "dynamicArrayItem", data: dynamicArrayItem },
      ];
    }
    */
    // console.log(result, txt);
    return [true, _.isNil(formulaError) ? result : formulaError, txt];
}
function insertUpdateDynamicArray(ctx, dynamicArrayItem) {
    var r = dynamicArrayItem.r, c = dynamicArrayItem.c;
    var id = dynamicArrayItem.id;
    if (_.isNil(id)) {
        id = ctx.currentSheetId;
    }
    var luckysheetfile = ctx.luckysheetfile;
    var idx = getSheetIndex(ctx, id);
    if (idx == null)
        return [];
    var file = luckysheetfile[idx];
    var dynamicArray = file.dynamicArray;
    if (_.isNil(dynamicArray)) {
        dynamicArray = [];
    }
    for (var i = 0; i < dynamicArray.length; i += 1) {
        var calc = dynamicArray[i];
        if (calc.r === r && calc.c === c && calc.id === id) {
            calc.data = dynamicArrayItem.data;
            calc.f = dynamicArrayItem.f;
            return dynamicArray;
        }
    }
    dynamicArray.push(dynamicArrayItem);
    return dynamicArray;
}
export function groupValuesRefresh(ctx) {
    var luckysheetfile = ctx.luckysheetfile;
    if (ctx.groupValuesRefreshData.length > 0) {
        for (var i = 0; i < ctx.groupValuesRefreshData.length; i += 1) {
            var item = ctx.groupValuesRefreshData[i];
            // if(item.i !== ctx.currentSheetId){
            //     continue;
            // }
            var idx = getSheetIndex(ctx, item.id);
            if (idx == null)
                continue;
            var file = luckysheetfile[idx];
            var data = file.data;
            if (_.isNil(data)) {
                continue;
            }
            var updateValue = {};
            if (!_.isNil(item.spe)) {
                if (item.spe.type === "sparklines") {
                    updateValue.spl = item.spe.data;
                }
                else if (item.spe.type === "dynamicArrayItem") {
                    file.dynamicArray = insertUpdateDynamicArray(ctx, item.spe.data);
                }
            }
            updateValue.v = item.v;
            updateValue.f = item.f;
            setCellValue(ctx, item.r, item.c, data, updateValue);
            // server.saveParam("v", item.id, data[item.r][item.c], {
            //     "r": item.r,
            //     "c": item.c
            // });
        }
        // editor.webWorkerFlowDataCache(Store.flowdata); // worker存数据
        ctx.groupValuesRefreshData = [];
    }
}
export function execFunctionGroup(ctx, origin_r, origin_c, value, id, data, isForce) {
    if (isForce === void 0) { isForce = false; }
    if (_.isNil(data)) {
        data = getFlowdata(ctx);
    }
    // if (!window.luckysheet_compareWith) {
    //   window.luckysheet_compareWith = luckysheet_compareWith;
    //   window.luckysheet_getarraydata = luckysheet_getarraydata;
    //   window.luckysheet_getcelldata = luckysheet_getcelldata;
    //   window.luckysheet_parseData = luckysheet_parseData;
    //   window.luckysheet_getValue = luckysheet_getValue;
    //   window.luckysheet_indirect_check = luckysheet_indirect_check;
    //   window.luckysheet_indirect_check_return = luckysheet_indirect_check_return;
    //   window.luckysheet_offset_check = luckysheet_offset_check;
    //   window.luckysheet_calcADPMM = luckysheet_calcADPMM;
    //   window.luckysheet_getSpecialReference = luckysheet_getSpecialReference;
    // }
    if (_.isNil(ctx.formulaCache.execFunctionGlobalData)) {
        ctx.formulaCache.execFunctionGlobalData = {};
    }
    // let luckysheetfile = getluckysheetfile();
    // let dynamicArray_compute = luckysheetfile[getSheetIndex(ctx.currentSheetId)_.isNil(]["dynamicArray_compute"]) ? {} : luckysheetfile[getSheetIndex(ctx.currentSheetId)]["dynamicArray_compute"];
    if (_.isNil(id)) {
        id = ctx.currentSheetId;
    }
    if (!_.isNil(value)) {
        // 此处setcellvalue 中this.execFunctionGroupData会保存想要更新的值，本函数结尾不要设为null,以备后续函数使用
        // setcellvalue(origin_r, origin_c, _this.execFunctionGroupData, value);
        var cellCache = [[{ v: undefined }]];
        setCellValue(ctx, 0, 0, cellCache, value);
        ctx.formulaCache.execFunctionGlobalData["".concat(origin_r, "_").concat(origin_c, "_").concat(id)] = cellCache[0][0];
    }
    // { "r": r, "c": c, "id": id, "func": func}
    var calcChains = getAllFunctionGroup(ctx);
    var formulaObjects = {};
    var sheets = ctx.luckysheetfile;
    var sheetData = {};
    for (var i = 0; i < sheets.length; i += 1) {
        var sheet = sheets[i];
        sheetData[sheet.id] = sheet.data;
    }
    // 把修改涉及的单元格存储为对象
    var updateValueOjects = {};
    var updateValueArray = [];
    if (_.isNil(ctx.formulaCache.execFunctionExist)) {
        var key = "r".concat(origin_r, "c").concat(origin_c, "i").concat(id);
        updateValueOjects[key] = 1;
    }
    else {
        for (var x = 0; x < ctx.formulaCache.execFunctionExist.length; x += 1) {
            var cell = ctx.formulaCache.execFunctionExist[x];
            var key = "r".concat(cell.r, "c").concat(cell.c, "i").concat(cell.i);
            updateValueOjects[key] = 1;
        }
    }
    var arrayMatchCache = {};
    var arrayMatch = function (formulaArray, _formulaObjects, _updateValueOjects, func) {
        for (var a = 0; a < formulaArray.length; a += 1) {
            var range = formulaArray[a];
            var cacheKey = "r".concat(range.row[0]).concat(range.row[1], "c").concat(range.column[0]).concat(range.column[1], "id").concat(range.sheetId);
            if (cacheKey in arrayMatchCache) {
                var amc = arrayMatchCache[cacheKey];
                // console.log(amc);
                amc.forEach(function (item) {
                    func(item.key, item.r, item.c, item.sheetId);
                });
            }
            else {
                var functionArr = [];
                for (var r = range.row[0]; r <= range.row[1]; r += 1) {
                    for (var c = range.column[0]; c <= range.column[1]; c += 1) {
                        var key = "r".concat(r, "c").concat(c, "i").concat(range.sheetId);
                        func(key, r, c, range.sheetId);
                        if ((_formulaObjects && key in _formulaObjects) ||
                            (_updateValueOjects && key in _updateValueOjects)) {
                            functionArr.push({
                                key: key,
                                r: r,
                                c: c,
                                sheetId: range.sheetId,
                            });
                        }
                    }
                }
                if (_formulaObjects || _updateValueOjects) {
                    arrayMatchCache[cacheKey] = functionArr;
                }
            }
        }
    };
    var _loop_1 = function (i) {
        var formulaCell = calcChains[i];
        var key = "r".concat(formulaCell.r, "c").concat(formulaCell.c, "i").concat(formulaCell.id);
        var calc_funcStr = getcellFormula(ctx, formulaCell.r, formulaCell.c, formulaCell.id);
        if (_.isNil(calc_funcStr)) {
            return "continue";
        }
        var txt1 = calc_funcStr.toUpperCase();
        var isOffsetFunc = txt1.indexOf("INDIRECT(") > -1 ||
            txt1.indexOf("OFFSET(") > -1 ||
            txt1.indexOf("INDEX(") > -1;
        var formulaArray = [];
        if (isOffsetFunc) {
            isFunctionRange(ctx, calc_funcStr, null, null, formulaCell.id, null, function (str_nb) {
                var range = getcellrange(ctx, _.trim(str_nb), formulaCell.id);
                if (!_.isNil(range)) {
                    formulaArray.push(range);
                }
            });
        }
        else if (!(calc_funcStr.substring(0, 2) === '="' &&
            calc_funcStr.substring(calc_funcStr.length - 1, 1) === '"')) {
            // let formulaTextArray = calc_funcStr.split(/==|!=|<>|<=|>=|[,()=+-\/*%&^><]/g);//无法正确分割单引号或双引号之间有==、!=、-等运算符的情况。导致如='1-2'!A1公式中表名1-2的A1单元格内容更新后，公式的值不更新的bug
            // 解决='1-2'!A1+5会被calc_funcStr.split(/==|!=|<>|<=|>=|[,()=+-\/*%&^><]/g)分割成["","'1","2'!A1",5]的错误情况
            var point = 0; // 指针
            var squote = -1; // 双引号
            var dquote = -1; // 单引号
            var formulaTextArray = [];
            var sq_end_array = []; // 保存了配对的单引号在formulaTextArray的index索引。
            var calc_funcStr_length = calc_funcStr.length;
            for (var j = 0; j < calc_funcStr_length; j += 1) {
                var char = calc_funcStr.charAt(j);
                if (char === "'" && dquote === -1) {
                    // 如果是单引号开始
                    if (squote === -1) {
                        if (point !== j) {
                            formulaTextArray.push.apply(formulaTextArray, calc_funcStr
                                .substring(point, j)
                                .split(/==|!=|<>|<=|>=|[,()=+-/*%&^><]/));
                        }
                        squote = j;
                        point = j;
                    } // 单引号结束
                    else {
                        // if (squote === i - 1)//配对的单引号后第一个字符不能是单引号
                        // {
                        //    ;//到此处说明公式错误
                        // }
                        // 如果是''代表着输出'
                        if (j < calc_funcStr_length - 1 &&
                            calc_funcStr.charAt(j + 1) === "'") {
                            j += 1;
                        }
                        else {
                            // 如果下一个字符不是'代表单引号结束
                            // if (calc_funcStr.charAt(i - 1) === "'") {//配对的单引号后最后一个字符不能是单引号
                            //    ;//到此处说明公式错误
                            point = j + 1;
                            formulaTextArray.push(calc_funcStr.substring(squote, point));
                            sq_end_array.push(formulaTextArray.length - 1);
                            squote = -1;
                            // } else {
                            //    point = i + 1;
                            //    formulaTextArray.push(calc_funcStr.substring(squote, point));
                            //    sq_end_array.push(formulaTextArray.length - 1);
                            //    squote = -1;
                            // }
                        }
                    }
                }
                if (char === '"' && squote === -1) {
                    // 如果是双引号开始
                    if (dquote === -1) {
                        if (point !== j) {
                            formulaTextArray.push.apply(formulaTextArray, calc_funcStr
                                .substring(point, j)
                                .split(/==|!=|<>|<=|>=|[,()=+-/*%&^><]/));
                        }
                        dquote = j;
                        point = j;
                    }
                    else {
                        // 如果是""代表着输出"
                        if (j < calc_funcStr_length - 1 &&
                            calc_funcStr.charAt(j + 1) === '"') {
                            j += 1;
                        }
                        else {
                            // 双引号结束
                            point = j + 1;
                            formulaTextArray.push(calc_funcStr.substring(dquote, point));
                            dquote = -1;
                        }
                    }
                }
            }
            if (point !== calc_funcStr_length) {
                formulaTextArray.push.apply(formulaTextArray, calc_funcStr
                    .substring(point, calc_funcStr_length)
                    .split(/==|!=|<>|<=|>=|[,()=+-/*%&^><]/));
            }
            // 拼接所有配对单引号及之后一个单元格内容，例如["'1-2'","!A1"]拼接为["'1-2'!A1"]
            for (var j = sq_end_array.length - 1; j >= 0; j -= 1) {
                if (sq_end_array[j] !== formulaTextArray.length - 1) {
                    formulaTextArray[sq_end_array[j]] +=
                        formulaTextArray[sq_end_array[j] + 1];
                    formulaTextArray.splice(sq_end_array[j] + 1, 1);
                }
            }
            // 至此=SUM('1-2'!A1:A2&"'1-2'!A2")由原来的["","SUM","'1","2'!A1:A2","",""'1","2'!A2""]更正为["","SUM","","'1-2'!A1:A2","","",""'1-2'!A2""]
            for (var j = 0; j < formulaTextArray.length; j += 1) {
                var t = formulaTextArray[j];
                if (t.length <= 1) {
                    continue;
                }
                if ((t.substring(0, 1) === '"' && t.substring(t.length - 1, 1) === '"') ||
                    !iscelldata(t)) {
                    continue;
                }
                var range = getcellrange(ctx, _.trim(t), formulaCell.id);
                if (_.isNil(range)) {
                    continue;
                }
                formulaArray.push(range);
            }
        }
        var item = {
            formulaArray: formulaArray,
            calc_funcStr: calc_funcStr,
            key: key,
            r: formulaCell.r,
            c: formulaCell.c,
            id: formulaCell.id,
            parents: {},
            chidren: {},
            color: "w",
        };
        formulaObjects[key] = item;
    };
    // 创建公式缓存及其范围的缓存
    // console.time("1");
    for (var i = 0; i < calcChains.length; i += 1) {
        _loop_1(i);
    }
    // console.timeEnd("1");
    // console.time("2");
    // 形成一个公式之间引用的图结构
    Object.keys(formulaObjects).forEach(function (key) {
        var formulaObject = formulaObjects[key];
        arrayMatch(formulaObject.formulaArray, formulaObjects, updateValueOjects, function (childKey) {
            if (childKey in formulaObjects) {
                var childFormulaObject = formulaObjects[childKey];
                formulaObject.chidren[childKey] = 1;
                childFormulaObject.parents[key] = 1;
            }
            // console.log(childKey,formulaObject.formulaArray);
            if (!isForce && childKey in updateValueOjects) {
                updateValueArray.push(formulaObject);
            }
        });
        if (isForce) {
            updateValueArray.push(formulaObject);
        }
    });
    // console.log(formulaObjects)
    // console.timeEnd("2");
    // console.time("3");
    var formulaRunList = [];
    // 计算，采用深度优先遍历公式形成的图结构
    // updateValueArray.forEach((key)=>{
    //     let formulaObject = formulaObjects[key];
    // });
    var stack = updateValueArray;
    var existsFormulaRunList = {};
    var _loop_2 = function () {
        var formulaObject = stack.pop();
        if (_.isNil(formulaObject) || formulaObject.key in existsFormulaRunList) {
            return "continue";
        }
        if (formulaObject.color === "b") {
            formulaRunList.push(formulaObject);
            existsFormulaRunList[formulaObject.key] = 1;
            return "continue";
        }
        var cacheStack = [];
        Object.keys(formulaObject.parents).forEach(function (parentKey) {
            var parentFormulaObject = formulaObjects[parentKey];
            if (!_.isNil(parentFormulaObject)) {
                cacheStack.push(parentFormulaObject);
            }
        });
        if (cacheStack.length === 0) {
            formulaRunList.push(formulaObject);
            existsFormulaRunList[formulaObject.key] = 1;
        }
        else {
            formulaObject.color = "b";
            stack.push(formulaObject);
            stack = stack.concat(cacheStack);
        }
    };
    while (stack.length > 0) {
        _loop_2();
    }
    formulaRunList.reverse();
    var calcChainSet = new Set();
    calcChains.forEach(function (item) {
        calcChainSet.add("".concat(item.r, "_").concat(item.c, "_").concat(item.id));
    });
    // console.log(formulaObjects, ii)
    // console.timeEnd("3");
    // console.time("4");
    for (var i = 0; i < formulaRunList.length; i += 1) {
        var formulaCell = formulaRunList[i];
        if (formulaCell.level === Math.max) {
            continue;
        }
        var calc_funcStr = formulaCell.calc_funcStr;
        var v = execfunction(ctx, calc_funcStr, formulaCell.r, formulaCell.c, formulaCell.id, calcChainSet);
        ctx.groupValuesRefreshData.push({
            r: formulaCell.r,
            c: formulaCell.c,
            v: v[1],
            f: v[2],
            spe: v[3],
            id: formulaCell.id,
        });
        // _this.execFunctionGroupData[u.r][u.c] = value;
        ctx.formulaCache.execFunctionGlobalData["".concat(formulaCell.r, "_").concat(formulaCell.c, "_").concat(formulaCell.id)] = {
            v: v[1],
            f: v[2],
        };
    }
    // console.log(formulaRunList);
    // console.timeEnd("4");
    ctx.formulaCache.execFunctionExist = undefined;
}
function findrangeindex(ctx, v, vp) {
    var re = /<span.*?>/g;
    var v_a = v.replace(re, "").split("</span>");
    var vp_a = vp.replace(re, "").split("</span>");
    v_a.pop();
    if (vp_a[vp_a.length - 1] === "")
        vp_a.pop();
    var pfri = ctx.formulaCache.functionRangeIndex;
    if (pfri == null)
        return [];
    var vplen = vp_a.length;
    var vlen = v_a.length;
    // 不增加元素输入
    if (vplen === vlen) {
        var i = pfri[0];
        var p = vp_a[i];
        var n = v_a[i];
        if (_.isNil(p)) {
            if (vp_a.length <= i) {
                pfri = [vp_a.length - 1, vp_a.length - 1];
            }
            else if (v_a.length <= i) {
                pfri = [v_a.length - 1, v_a.length - 1];
            }
            return pfri;
        }
        if (p.length === n.length) {
            if (!_.isNil(vp_a[i + 1]) &&
                !_.isNil(v_a[i + 1]) &&
                vp_a[i + 1].length < v_a[i + 1].length) {
                pfri[0] += 1;
                pfri[1] = 1;
            }
            return pfri;
        }
        if (p.length > n.length) {
            if (!_.isNil(p) &&
                !_.isNil(v_a[i + 1]) &&
                v_a[i + 1].substring(0, 1) === '"' &&
                (p.indexOf("{") > -1 || p.indexOf("}") > -1)) {
                pfri[0] += 1;
                pfri[1] = 1;
            }
            return pfri;
        }
        if (p.length < n.length) {
            if (pfri[1] > n.length) {
                pfri[1] = n.length;
            }
            return pfri;
        }
    }
    // 减少元素输入
    else if (vplen > vlen) {
        var i = pfri[0];
        var p = vp_a[i];
        var n = v_a[i];
        if (_.isNil(n)) {
            if (v_a[i - 1].indexOf("{") > -1) {
                pfri[0] -= 1;
                var start = v_a[i - 1].search("{");
                pfri[1] += start;
            }
            else {
                pfri[0] = 0;
                pfri[1] = 0;
            }
        }
        else if (p.length === n.length) {
            if (!_.isNil(v_a[i + 1]) &&
                (v_a[i + 1].substring(0, 1) === '"' ||
                    v_a[i + 1].substring(0, 1) === "{" ||
                    v_a[i + 1].substring(0, 1) === "}")) {
                pfri[0] += 1;
                pfri[1] = 1;
            }
            else if (!_.isNil(p) &&
                p.length > 2 &&
                p.substring(0, 1) === '"' &&
                p.substring(p.length - 1, 1) === '"') {
                // pfri[1] = n.length-1;
            }
            else if (!_.isNil(v_a[i]) && v_a[i] === '")') {
                pfri[1] = 1;
            }
            else if (!_.isNil(v_a[i]) && v_a[i] === '"}') {
                pfri[1] = 1;
            }
            else if (!_.isNil(v_a[i]) && v_a[i] === "{)") {
                pfri[1] = 1;
            }
            else {
                pfri[1] = n.length;
            }
            return pfri;
        }
        else if (p.length > n.length) {
            if (!_.isNil(v_a[i + 1]) &&
                (v_a[i + 1].substring(0, 1) === '"' ||
                    v_a[i + 1].substring(0, 1) === "{" ||
                    v_a[i + 1].substring(0, 1) === "}")) {
                pfri[0] += 1;
                pfri[1] = 1;
            }
            return pfri;
        }
        else if (p.length < n.length) {
            return pfri;
        }
        return pfri;
    }
    // 增加元素输入
    else if (vplen < vlen) {
        var i = pfri[0];
        var p = vp_a[i];
        var n = v_a[i];
        if (_.isNil(p)) {
            pfri[0] = v_a.length - 1;
            if (!_.isNil(n)) {
                pfri[1] = n.length;
            }
            else {
                pfri[1] = 1;
            }
        }
        else if (p.length === n.length) {
            if (vp_a[i + 1] != null &&
                (vp_a[i + 1].substring(0, 1) === '"' ||
                    vp_a[i + 1].substring(0, 1) === "{" ||
                    vp_a[i + 1].substring(0, 1) === "}")) {
                pfri[1] = n.length;
            }
            else if (!_.isNil(v_a[i + 1]) &&
                v_a[i + 1].substring(0, 1) === '"' &&
                (v_a[i + 1].substring(0, 1) === "{" ||
                    v_a[i + 1].substring(0, 1) === "}")) {
                pfri[0] += 1;
                pfri[1] = 1;
            }
            else if (!_.isNil(n) &&
                n.substring(0, 1) === '"' &&
                n.substring(n.length - 1, 1) === '"' &&
                p.substring(0, 1) === '"' &&
                p.substring(p.length - 1, 1) === ")") {
                pfri[1] = n.length;
            }
            else if (!_.isNil(n) &&
                n.substring(0, 1) === "{" &&
                n.substring(n.length - 1, 1) === "}" &&
                p.substring(0, 1) === "{" &&
                p.substring(p.length - 1, 1) === ")") {
                pfri[1] = n.length;
            }
            else {
                pfri[0] = pfri[0] + vlen - vplen;
                if (v_a.length > vp_a.length) {
                    pfri[1] = v_a[i + 1].length;
                }
                else {
                    pfri[1] = 1;
                }
            }
            return pfri;
        }
        else if (p.length > n.length) {
            if (!_.isNil(p) && p.substring(0, 1) === '"') {
                pfri[1] = n.length;
            }
            else if (_.isNil(v_a[i + 1]) && /{.*?}/.test(v_a[i + 1])) {
                pfri[0] += 1;
                pfri[1] = v_a[i + 1].length;
            }
            else if (!_.isNil(p) &&
                v_a[i + 1].substring(0, 1) === '"' &&
                (p.indexOf("{") > -1 || p.indexOf("}") > -1)) {
                pfri[0] += 1;
                pfri[1] = 1;
            }
            else if (!_.isNil(p) && (p.indexOf("{") > -1 || p.indexOf("}") > -1)) {
            }
            else if (!_.isNil(p) &&
                !_.startsWith(p[0], "=") &&
                _.startsWith(n, "=")) {
                return [vlen - 1, v_a[vlen - 1].length];
            }
            else {
                pfri[0] = pfri[0] + vlen - vplen - 1;
                pfri[1] = v_a[(i || 1) - 1].length;
            }
            return pfri;
        }
        else if (p.length < n.length) {
            return pfri;
        }
        return pfri;
    }
    return null;
}
export function createFormulaRangeSelect(ctx, select) {
    ctx.formulaRangeSelect = select;
}
export function createRangeHightlight(ctx, inputInnerHtmlStr, ignoreRangeIndex) {
    if (ignoreRangeIndex === void 0) { ignoreRangeIndex = -1; }
    var $span = parseElement("<div>".concat(inputInnerHtmlStr, "</div>"));
    var formulaRanges = [];
    $span
        .querySelectorAll("span.fortune-formula-functionrange-cell")
        .forEach(function (ele) {
        var rangeIndex = parseInt(ele.getAttribute("rangeindex") || "0", 10);
        if (rangeIndex === ignoreRangeIndex)
            return;
        var cellrange = getcellrange(ctx, ele.textContent || "");
        if (rangeIndex === ctx.formulaCache.selectingRangeIndex ||
            cellrange == null)
            return;
        if (cellrange.sheetId === ctx.currentSheetId ||
            (cellrange.sheetId === -1 &&
                ctx.formulaCache.rangetosheet === ctx.currentSheetId)) {
            var rect = seletedHighlistByindex(ctx, cellrange.row[0], cellrange.row[1], cellrange.column[0], cellrange.column[1]);
            if (rect) {
                formulaRanges.push(__assign(__assign({ rangeIndex: rangeIndex }, rect), { backgroundColor: colors[rangeIndex] }));
            }
        }
    });
    ctx.formulaRangeHighlight = formulaRanges;
}
export function setCaretPosition(ctx, textDom, children, pos) {
    try {
        var el = textDom;
        var range = document.createRange();
        var sel = window.getSelection();
        range.setStart(el.childNodes[children], pos);
        range.collapse(true);
        sel === null || sel === void 0 ? void 0 : sel.removeAllRanges();
        sel === null || sel === void 0 ? void 0 : sel.addRange(range);
        el.focus();
    }
    catch (err) {
        console.error(err);
        moveToEnd(ctx.formulaCache.rangeResizeTo[0]);
    }
}
function functionRange(ctx, obj, v, vp) {
    if (window.getSelection) {
        // ie11 10 9 ff safari
        var currSelection = window.getSelection();
        if (!currSelection)
            return;
        var fri = findrangeindex(ctx, v, vp);
        if (_.isNil(fri)) {
            currSelection.selectAllChildren(obj);
            currSelection.collapseToEnd();
        }
        else {
            setCaretPosition(ctx, obj.querySelectorAll("span")[fri[0]], 0, fri[1]);
        }
        // @ts-ignore
    }
    else if (document.selection) {
        // ie10 9 8 7 6 5
        // @ts-ignore
        ctx.formulaCache.functionRangeIndex.moveToElementText(obj); // range定位到obj
        // @ts-ignore
        ctx.formulaCache.functionRangeIndex.collapse(false); // 光标移至最后
        // @ts-ignore
        ctx.formulaCache.functionRangeIndex.select();
    }
}
function searchFunction(ctx, searchtxt) {
    var functionlist = locale(ctx).functionlist;
    // // 这里的逻辑在原项目上做了修改
    // if (_.isNil($editer)) {
    //   return;
    // }
    // const inputContent = $editer.innerText.toUpperCase();
    // const reg = /^=([a-zA-Z_]+)\(?/;
    // const match = inputContent.match(reg);
    // if (!match) {
    //   ctx.functionCandidates = [];
    //   return;
    // }
    // const searchtxt = match[1];
    var f = [];
    var s = [];
    var t = [];
    var result_i = 0;
    for (var i = 0; i < functionlist.length; i += 1) {
        var item = functionlist[i];
        var n = item.n;
        if (n === searchtxt) {
            f.unshift(item);
            result_i += 1;
        }
        else if (_.startsWith(n, searchtxt)) {
            s.unshift(item);
            result_i += 1;
        }
        else if (n.indexOf(searchtxt) > -1) {
            t.unshift(item);
            result_i += 1;
        }
        if (result_i >= 10) {
            break;
        }
    }
    var list = __spreadArray(__spreadArray(__spreadArray([], f, true), s, true), t, true);
    if (list.length <= 0) {
        return;
    }
    ctx.functionCandidates = list;
    // const listHTML = _this.searchFunctionHTML(list);
    // $("#luckysheet-formula-search-c").html(listHTML).show();
    // $("#luckysheet-formula-help-c").hide();
    // const $c = $editer.parent();
    // const offset = $c.offset();
    // _this.searchFunctionPosition(
    //   $("#luckysheet-formula-search-c"),
    //   $c,
    //   offset.left,
    //   offset.top
    // );
}
export function getrangeseleciton() {
    var _a, _b, _c, _d, _e;
    var currSelection = window.getSelection();
    if (!currSelection)
        return null;
    var anchorNode = currSelection.anchorNode, anchorOffset = currSelection.anchorOffset;
    if (!anchorNode)
        return null;
    if (((_b = (_a = anchorNode.parentNode) === null || _a === void 0 ? void 0 : _a.nodeName) === null || _b === void 0 ? void 0 : _b.toLowerCase()) === "span" &&
        anchorOffset !== 0) {
        var txt = _.trim(anchorNode.textContent || "");
        if (txt.length === 0 && anchorNode.parentNode.previousSibling) {
            var ahr = anchorNode.parentNode.previousSibling;
            txt = _.trim(ahr.textContent || "");
            return ahr;
        }
        return anchorNode.parentNode;
    }
    var anchorElement = anchorNode;
    if (anchorElement.id === "luckysheet-rich-text-editor" ||
        anchorElement.id === "luckysheet-functionbox-cell") {
        var txt = _.trim((_c = _.last(anchorElement.querySelectorAll("span"))) === null || _c === void 0 ? void 0 : _c.innerText);
        if (txt.length === 0 && anchorElement.querySelectorAll("span").length > 1) {
            var ahr = anchorElement.querySelectorAll("span");
            txt = _.trim(ahr[ahr.length - 2].innerText);
            return ahr === null || ahr === void 0 ? void 0 : ahr[0];
        }
        return _.last(anchorElement.querySelectorAll("span"));
    }
    if (((_d = anchorNode === null || anchorNode === void 0 ? void 0 : anchorNode.parentElement) === null || _d === void 0 ? void 0 : _d.id) === "luckysheet-rich-text-editor" ||
        ((_e = anchorNode === null || anchorNode === void 0 ? void 0 : anchorNode.parentElement) === null || _e === void 0 ? void 0 : _e.id) === "luckysheet-functionbox-cell" ||
        anchorOffset === 0) {
        var newAnchorNode = anchorOffset === 0 ? anchorNode === null || anchorNode === void 0 ? void 0 : anchorNode.parentNode : anchorNode;
        if (newAnchorNode === null || newAnchorNode === void 0 ? void 0 : newAnchorNode.previousSibling) {
            return newAnchorNode === null || newAnchorNode === void 0 ? void 0 : newAnchorNode.previousSibling;
        }
    }
    return null;
}
function helpFunctionExe($editer, currSelection, ctx) {
    var _a;
    var functionlist = locale(ctx).functionlist;
    // let _locale = locale();
    // let locale_formulaMore = _locale.formulaMore;
    // if ($("#luckysheet-formula-help-c").length === 0) {
    //   $("body").after(
    //     replaceHtml(_this.helpHTML, {
    //       helpClose: locale_formulaMore.helpClose,
    //       helpCollapse: locale_formulaMore.helpCollapse,
    //       helpExample: locale_formulaMore.helpExample,
    //       helpAbstract: locale_formulaMore.helpAbstract,
    //     })
    //   );
    //   $("#luckysheet-formula-help-c .luckysheet-formula-help-close").click(
    //     function () {
    //       $("#luckysheet-formula-help-c").hide();
    //     }
    //   );
    //   $("#luckysheet-formula-help-c .luckysheet-formula-help-collapse").click(
    //     function () {
    //       let $content = $(
    //         "#luckysheet-formula-help-c .luckysheet-formula-help-content"
    //       );
    //       $content.slideToggle(100, function () {
    //         let $c = _this.rangeResizeTo.parent(),
    //           offset = $c.offset();
    //         _this.searchFunctionPosition(
    //           $("#luckysheet-formula-help-c"),
    //           $c,
    //           offset.left,
    //           offset.top,
    //           true
    //         );
    //       });
    //       if ($content.is(":hidden")) {
    //         $(this).html('<i class="fa fa-angle-up" aria-hidden="true"></i>');
    //       } else {
    //         $(this).html('<i class="fa fa-angle-down" aria-hidden="true"></i>');
    //       }
    //     }
    //   );
    //   for (let i = 0; i < functionlist.length; i++) {
    //     functionlistPosition[functionlist[i].n] = i;
    //   }
    // }
    if (_.isEmpty(ctx.formulaCache.functionlistMap)) {
        for (var i_1 = 0; i_1 < functionlist.length; i_1 += 1) {
            ctx.formulaCache.functionlistMap[functionlist[i_1].n] = functionlist[i_1];
        }
    }
    if (!currSelection) {
        return null;
    }
    var $prev = currSelection;
    var $span = $editer.querySelectorAll("span");
    var currentIndex = _.indexOf((_a = currSelection.parentNode) === null || _a === void 0 ? void 0 : _a.childNodes, currSelection);
    var i = currentIndex;
    if ($prev == null) {
        return null;
    }
    var funcName = null;
    var paramindex = null;
    if ($span[i].classList.contains("luckysheet-formula-text-func")) {
        funcName = $span[i].textContent;
    }
    else {
        var $cur = null;
        var exceptIndex = [-1, -1];
        // eslint-disable-next-line no-plusplus
        while (--i > 0) {
            $cur = $span[i];
            if ($cur.classList.contains("luckysheet-formula-text-func") ||
                _.trim($cur.textContent || "").toUpperCase() in
                    ctx.formulaCache.functionlistMap) {
                funcName = $cur.textContent;
                paramindex = null;
                var endstate = true;
                for (var a = i; a <= currentIndex; a += 1) {
                    if (!paramindex) {
                        paramindex = 0;
                    }
                    if (a >= exceptIndex[0] && a <= exceptIndex[1]) {
                        continue;
                    }
                    $cur = $span[a];
                    if ($cur.classList.contains("luckysheet-formula-text-rpar")) {
                        exceptIndex = [i, a];
                        funcName = null;
                        endstate = false;
                        break;
                    }
                    if ($cur.classList.contains("luckysheet-formula-text-comma")) {
                        paramindex += 1;
                    }
                }
                if (endstate) {
                    break;
                }
            }
        }
    }
    return funcName;
}
export function rangeHightlightselected(ctx, $editor) {
    var currSelection = getrangeseleciton();
    // $("#luckysheet-formula-search-c, #luckysheet-formula-help-c").hide();
    // $(
    //   "#fortune-formula-functionrange .fortune-formula-functionrange-highlight .fortune-selection-copy-hc"
    // ).css("opacity", "0.03");
    // $("#luckysheet-formula-search-c, #luckysheet-formula-help-c").hide();
    // if (
    //   $(currSelection).closest(".fortune-formula-functionrange-cell").length ==
    //   0
    // ) {
    if (!currSelection)
        return;
    var currText = _.trim(currSelection.textContent || "");
    if (currText === null || currText === void 0 ? void 0 : currText.match(/^[a-zA-Z_]+$/)) {
        searchFunction(ctx, currText.toUpperCase());
        ctx.functionHint = null;
    }
    else {
        var funcName = helpFunctionExe($editor, currSelection, ctx);
        ctx.functionHint = funcName === null || funcName === void 0 ? void 0 : funcName.toUpperCase();
        ctx.functionCandidates = [];
    }
    // return;
    // }
    // const $anchorOffset = $(currSelection).closest(
    //   ".fortune-formula-functionrange-cell"
    // );
    // const rangeindex = $anchorOffset.attr("rangeindex");
    // const rangeid = `fortune-formula-functionrange-highlight-${rangeindex}`;
    // $(`#${rangeid}`).find(".fortune-selection-copy-hc").css({
    //   opacity: "0.13",
    // });
}
function functionHTML(txt) {
    if (txt.substr(0, 1) === "=") {
        txt = txt.substr(1);
    }
    var funcstack = txt.split("");
    var i = 0;
    var str = "";
    var function_str = "";
    var matchConfig = {
        bracket: 0,
        comma: 0,
        squote: 0,
        dquote: 0,
        braces: 0,
    };
    while (i < funcstack.length) {
        var s = funcstack[i];
        if (s === "(" &&
            matchConfig.squote === 0 &&
            matchConfig.dquote === 0 &&
            matchConfig.braces === 0) {
            matchConfig.bracket += 1;
            if (str.length > 0) {
                function_str += "<span dir=\"auto\" class=\"luckysheet-formula-text-func\">".concat(str, "</span><span dir=\"auto\" class=\"luckysheet-formula-text-lpar\">(</span>");
            }
            else {
                function_str +=
                    '<span dir="auto" class="luckysheet-formula-text-lpar">(</span>';
            }
            str = "";
        }
        else if (s === ")" &&
            matchConfig.squote === 0 &&
            matchConfig.dquote === 0 &&
            matchConfig.braces === 0) {
            matchConfig.bracket -= 1;
            function_str += "".concat(functionHTML(str), "<span dir=\"auto\" class=\"luckysheet-formula-text-rpar\">)</span>");
            str = "";
        }
        else if (s === "{" &&
            matchConfig.squote === 0 &&
            matchConfig.dquote === 0) {
            str += "{";
            matchConfig.braces += 1;
        }
        else if (s === "}" &&
            matchConfig.squote === 0 &&
            matchConfig.dquote === 0) {
            str += "}";
            matchConfig.braces -= 1;
        }
        else if (s === '"' && matchConfig.squote === 0) {
            if (matchConfig.dquote > 0) {
                if (str.length > 0) {
                    function_str += "".concat(str, "\"</span>");
                }
                else {
                    function_str += '"</span>';
                }
                matchConfig.dquote -= 1;
                str = "";
            }
            else {
                matchConfig.dquote += 1;
                if (str.length > 0) {
                    function_str += "".concat(functionHTML(str), "<span dir=\"auto\" class=\"luckysheet-formula-text-string\">\"");
                }
                else {
                    function_str +=
                        '<span dir="auto" class="luckysheet-formula-text-string">"';
                }
                str = "";
            }
        }
        // 修正例如输入公式='1-2'!A1时，只有2'!A1是fortune-formula-functionrange-cell色，'1-是黑色的问题。
        else if (s === "'" && matchConfig.dquote === 0) {
            str += "'";
            matchConfig.squote = matchConfig.squote === 0 ? 1 : 0;
        }
        else if (s === "," &&
            matchConfig.squote === 0 &&
            matchConfig.dquote === 0 &&
            matchConfig.braces === 0) {
            // matchConfig.comma += 1;
            function_str += "".concat(functionHTML(str), "<span dir=\"auto\" class=\"luckysheet-formula-text-comma\">,</span>");
            str = "";
        }
        else if (s === "&" &&
            matchConfig.squote === 0 &&
            matchConfig.dquote === 0 &&
            matchConfig.braces === 0) {
            if (str.length > 0) {
                function_str +=
                    "".concat(functionHTML(str), "<span dir=\"auto\" class=\"luckysheet-formula-text-calc\">") +
                        "&" +
                        "</span>";
                str = "";
            }
            else {
                function_str +=
                    '<span dir="auto" class="luckysheet-formula-text-calc">' +
                        "&" +
                        "</span>";
            }
        }
        else if (s in operatorjson &&
            matchConfig.squote === 0 &&
            matchConfig.dquote === 0 &&
            matchConfig.braces === 0) {
            var s_next = "";
            if (i + 1 < funcstack.length) {
                s_next = funcstack[i + 1];
            }
            var p = i - 1;
            var s_pre = null;
            if (p >= 0) {
                do {
                    s_pre = funcstack[p];
                    p -= 1;
                } while (p >= 0 && s_pre === " ");
            }
            if (s + s_next in operatorjson) {
                if (str.length > 0) {
                    function_str += "".concat(functionHTML(str), "<span dir=\"auto\" class=\"luckysheet-formula-text-calc\">").concat(s).concat(s_next, "</span>");
                    str = "";
                }
                else {
                    function_str += "<span dir=\"auto\" class=\"luckysheet-formula-text-calc\">".concat(s).concat(s_next, "</span>");
                }
                i += 1;
            }
            else if (!/[^0-9]/.test(s_next) &&
                s === "-" &&
                (s_pre === "(" ||
                    _.isNil(s_pre) ||
                    s_pre === "," ||
                    s_pre === " " ||
                    s_pre in operatorjson)) {
                str += s;
            }
            else {
                if (str.length > 0) {
                    function_str += "".concat(functionHTML(str), "<span dir=\"auto\" class=\"luckysheet-formula-text-calc\">").concat(s, "</span>");
                    str = "";
                }
                else {
                    function_str += "<span dir=\"auto\" class=\"luckysheet-formula-text-calc\">".concat(s, "</span>");
                }
            }
        }
        else {
            str += s;
        }
        if (i === funcstack.length - 1) {
            // function_str += str;
            if (iscelldata(_.trim(str))) {
                var rangeIndex = rangeIndexes.length > functionHTMLIndex
                    ? rangeIndexes[functionHTMLIndex]
                    : functionHTMLIndex;
                function_str += "<span class=\"fortune-formula-functionrange-cell\" rangeindex=\"".concat(rangeIndex, "\" dir=\"auto\" style=\"color:").concat(colors[rangeIndex], ";\">").concat(str, "</span>");
                functionHTMLIndex += 1;
            }
            else if (matchConfig.dquote > 0) {
                function_str += "".concat(str, "</span>");
            }
            else if (str.indexOf("</span>") === -1 && str.length > 0) {
                var regx = /{.*?}/;
                if (regx.test(_.trim(str))) {
                    var arraytxt = regx.exec(str)[0];
                    var arraystart = str.search(regx);
                    var alltxt = "";
                    if (arraystart > 0) {
                        alltxt += "<span dir=\"auto\" class=\"luckysheet-formula-text-color\">".concat(str.substr(0, arraystart), "</span>");
                    }
                    alltxt += "<span dir=\"auto\" style=\"color:#959a05\" class=\"luckysheet-formula-text-array\">".concat(arraytxt, "</span>");
                    if (arraystart + arraytxt.length < str.length) {
                        alltxt += "<span dir=\"auto\" class=\"luckysheet-formula-text-color\">".concat(str.substr(arraystart + arraytxt.length, str.length), "</span>");
                    }
                    function_str += alltxt;
                }
                else {
                    function_str += "<span dir=\"auto\" class=\"luckysheet-formula-text-color\">".concat(str, "</span>");
                }
            }
        }
        i += 1;
    }
    return function_str;
}
export function functionHTMLGenerate(txt) {
    if (txt.length === 0 || txt.substring(0, 1) !== "=") {
        return txt;
    }
    functionHTMLIndex = 0;
    return "<span dir=\"auto\" class=\"luckysheet-formula-text-color\">=</span>".concat(functionHTML(txt));
}
function getRangeIndexes($editor) {
    var res = [];
    $editor
        .querySelectorAll("span.fortune-formula-functionrange-cell")
        .forEach(function (ele) {
        var indexStr = ele.getAttribute("rangeindex");
        if (indexStr) {
            var rangeIndex = parseInt(indexStr, 10);
            res.push(rangeIndex);
        }
    });
    return res;
}
export function handleFormulaInput(ctx, $copyTo, $editor, kcode, preText, refreshRangeSelect) {
    var _a, _b, _c, _d, _e, _f;
    if (refreshRangeSelect === void 0) { refreshRangeSelect = true; }
    // if (isEditMode()) {
    //   // 此模式下禁用公式栏
    //   return;
    // }
    var value1;
    var value1txt = preText !== null && preText !== void 0 ? preText : $editor.innerText;
    var value = $editor.innerText;
    value = escapeScriptTag(value);
    if (value.length > 0 &&
        value.substring(0, 1) === "=" &&
        (kcode !== 229 || value.length === 1)) {
        if (!refreshRangeSelect)
            rangeIndexes = getRangeIndexes($editor);
        value = functionHTMLGenerate(value);
        if (!refreshRangeSelect && functionHTMLIndex < rangeIndexes.length)
            refreshRangeSelect = true;
        value1 = functionHTMLGenerate(value1txt);
        rangeIndexes = [];
        if (window.getSelection) {
            // all browsers, except IE before version 9
            var currSelection = window.getSelection();
            if (!currSelection)
                return;
            if (((_a = currSelection.anchorNode) === null || _a === void 0 ? void 0 : _a.nodeName.toLowerCase()) === "div") {
                var editorlen = $editor.querySelectorAll("span").length;
                if (editorlen > 0)
                    ctx.formulaCache.functionRangeIndex = [
                        editorlen - 1,
                        (_b = $editor.querySelectorAll("span").item(editorlen - 1).textContent) === null || _b === void 0 ? void 0 : _b.length,
                    ];
            }
            else {
                ctx.formulaCache.functionRangeIndex = [
                    _.indexOf((_e = (_d = (_c = currSelection.anchorNode) === null || _c === void 0 ? void 0 : _c.parentNode) === null || _d === void 0 ? void 0 : _d.parentNode) === null || _e === void 0 ? void 0 : _e.childNodes, 
                    // @ts-ignore
                    (_f = currSelection.anchorNode) === null || _f === void 0 ? void 0 : _f.parentNode),
                    currSelection.anchorOffset,
                ];
            }
        }
        else {
            // Internet Explorer before version 9
            // @ts-ignore
            var textRange = document.selection.createRange();
            ctx.formulaCache.functionRangeIndex = textRange;
        }
        $editor.innerHTML = value;
        if ($copyTo)
            $copyTo.innerHTML = value;
        // the cursor will be set to the beginning of input box after set innerHTML,
        // restoring it to the correct position
        functionRange(ctx, $editor, value, value1);
        if (refreshRangeSelect) {
            cancelFunctionrangeSelected(ctx);
            if (kcode !== 46) {
                // delete不执行此函数
                createRangeHightlight(ctx, value);
            }
            ctx.formulaCache.rangestart = false;
            ctx.formulaCache.rangedrag_column_start = false;
            ctx.formulaCache.rangedrag_row_start = false;
            rangeHightlightselected(ctx, $editor);
        }
    }
    else if (_.startsWith(value1txt, "=") && !_.startsWith(value, "=")) {
        if ($copyTo)
            $copyTo.innerHTML = value;
        $editor.innerHTML = escapeHTMLTag(value);
    }
    else if (!_.startsWith(value1txt, "=")) {
        if (!$copyTo)
            return;
        if ($copyTo.id === "luckysheet-rich-text-editor") {
            if (!_.startsWith($copyTo.innerHTML, "<span")) {
                $copyTo.innerHTML = escapeHTMLTag(value);
            }
        }
        else {
            $copyTo.innerHTML = escapeHTMLTag(value);
        }
    }
}
function isfreezonFuc(txt) {
    var row = txt.replace(/[^0-9]/g, "");
    var col = txt.replace(/[^A-Za-z]/g, "");
    var row$ = txt.substr(txt.indexOf(row) - 1, 1);
    var col$ = txt.substr(txt.indexOf(col) - 1, 1);
    var ret = [false, false];
    if (row$ === "$") {
        ret[0] = true;
    }
    if (col$ === "$") {
        ret[1] = true;
    }
    return ret;
}
function functionStrChange_range(txt, type, rc, orient, stindex, step) {
    var val = txt.split("!");
    var rangetxt;
    var prefix = "";
    if (val.length > 1) {
        rangetxt = val[1];
        prefix = "".concat(val[0], "!");
    }
    else {
        rangetxt = val[0];
    }
    var r1;
    var r2;
    var c1;
    var c2;
    var $row0;
    var $row1;
    var $col0;
    var $col1;
    if (rangetxt.indexOf(":") === -1) {
        r1 = parseInt(rangetxt.replace(/[^0-9]/g, ""), 10) - 1;
        r2 = r1;
        c1 = columnCharToIndex(rangetxt.replace(/[^A-Za-z]/g, ""));
        c2 = c1;
        var freezonFuc = isfreezonFuc(rangetxt);
        $row0 = freezonFuc[0] ? "$" : "";
        $row1 = $row0;
        $col0 = freezonFuc[1] ? "$" : "";
        $col1 = $col0;
    }
    else {
        rangetxt = rangetxt.split(":");
        r1 = parseInt(rangetxt[0].replace(/[^0-9]/g, ""), 10) - 1;
        r2 = parseInt(rangetxt[1].replace(/[^0-9]/g, ""), 10) - 1;
        if (r1 > r2) {
            return txt;
        }
        c1 = columnCharToIndex(rangetxt[0].replace(/[^A-Za-z]/g, ""));
        c2 = columnCharToIndex(rangetxt[1].replace(/[^A-Za-z]/g, ""));
        if (c1 > c2) {
            return txt;
        }
        var freezonFuc0 = isfreezonFuc(rangetxt[0]);
        $row0 = freezonFuc0[0] ? "$" : "";
        $col0 = freezonFuc0[1] ? "$" : "";
        var freezonFuc1 = isfreezonFuc(rangetxt[1]);
        $row1 = freezonFuc1[0] ? "$" : "";
        $col1 = freezonFuc1[1] ? "$" : "";
    }
    if (type === "del") {
        if (rc === "row") {
            if (r1 >= stindex && r2 <= stindex + step - 1) {
                return error.r;
            }
            if (r1 > stindex + step - 1) {
                r1 -= step;
            }
            else if (r1 >= stindex) {
                r1 = stindex;
            }
            if (r2 > stindex + step - 1) {
                r2 -= step;
            }
            else if (r2 >= stindex) {
                r2 = stindex - 1;
            }
            if (r1 < 0) {
                r1 = 0;
            }
            if (r2 < r1) {
                r2 = r1;
            }
        }
        else if (rc === "col") {
            if (c1 >= stindex && c2 <= stindex + step - 1) {
                return error.r;
            }
            if (c1 > stindex + step - 1) {
                c1 -= step;
            }
            else if (c1 >= stindex) {
                c1 = stindex;
            }
            if (c2 > stindex + step - 1) {
                c2 -= step;
            }
            else if (c2 >= stindex) {
                c2 = stindex - 1;
            }
            if (c1 < 0) {
                c1 = 0;
            }
            if (c2 < c1) {
                c2 = c1;
            }
        }
        if (r1 === r2 && c1 === c2) {
            if (!Number.isNaN(r1) && !Number.isNaN(c1)) {
                return prefix + $col0 + indexToColumnChar(c1) + $row0 + (r1 + 1);
            }
            if (!Number.isNaN(r1)) {
                return prefix + $row0 + (r1 + 1);
            }
            if (!Number.isNaN(c1)) {
                return prefix + $col0 + indexToColumnChar(c1);
            }
            return txt;
        }
        if (Number.isNaN(c1) && Number.isNaN(c2)) {
            return "".concat(prefix + $row0 + (r1 + 1), ":").concat($row1).concat(r2 + 1);
        }
        if (Number.isNaN(r1) && Number.isNaN(r2)) {
            return "".concat(prefix + $col0 + indexToColumnChar(c1), ":").concat($col1).concat(indexToColumnChar(c2));
        }
        return "".concat(prefix + $col0 + indexToColumnChar(c1) + $row0 + (r1 + 1), ":").concat($col1).concat(indexToColumnChar(c2)).concat($row1).concat(r2 + 1);
    }
    if (type === "add") {
        if (rc === "row") {
            if (orient === "lefttop") {
                if (r1 >= stindex) {
                    r1 += step;
                }
                if (r2 >= stindex) {
                    r2 += step;
                }
            }
            else if (orient === "rightbottom") {
                if (r1 > stindex) {
                    r1 += step;
                }
                if (r2 > stindex) {
                    r2 += step;
                }
            }
        }
        else if (rc === "col") {
            if (orient === "lefttop") {
                if (c1 >= stindex) {
                    c1 += step;
                }
                if (c2 >= stindex) {
                    c2 += step;
                }
            }
            else if (orient === "rightbottom") {
                if (c1 > stindex) {
                    c1 += step;
                }
                if (c2 > stindex) {
                    c2 += step;
                }
            }
        }
        if (r1 === r2 && c1 === c2) {
            if (!Number.isNaN(r1) && !Number.isNaN(c1)) {
                return prefix + $col0 + indexToColumnChar(c1) + $row0 + (r1 + 1);
            }
            if (!Number.isNaN(r1)) {
                return prefix + $row0 + (r1 + 1);
            }
            if (!Number.isNaN(c1)) {
                return prefix + $col0 + indexToColumnChar(c1);
            }
            return txt;
        }
        if (Number.isNaN(c1) && Number.isNaN(c2)) {
            return "".concat(prefix + $row0 + (r1 + 1), ":").concat($row1).concat(r2 + 1);
        }
        if (Number.isNaN(r1) && Number.isNaN(r2)) {
            return "".concat(prefix + $col0 + indexToColumnChar(c1), ":").concat($col1).concat(indexToColumnChar(c2));
        }
        return "".concat(prefix + $col0 + indexToColumnChar(c1) + $row0 + (r1 + 1), ":").concat($col1).concat(indexToColumnChar(c2)).concat($row1).concat(r2 + 1);
    }
    return "";
}
export function israngeseleciton(ctx, istooltip) {
    var _a, _b, _c;
    if (istooltip == null) {
        istooltip = false;
    }
    var currSelection = window.getSelection();
    if (currSelection == null)
        return false;
    var anchor = currSelection.anchorNode;
    if (!(anchor === null || anchor === void 0 ? void 0 : anchor.textContent))
        return false;
    var anchorOffset = currSelection.anchorOffset;
    var anchorElement = anchor;
    var parentElement = anchor.parentNode;
    if (((_a = anchor === null || anchor === void 0 ? void 0 : anchor.parentNode) === null || _a === void 0 ? void 0 : _a.nodeName.toLowerCase()) === "span" &&
        anchorOffset !== 0) {
        var txt = _.trim(anchor.textContent);
        var lasttxt = "";
        if (txt.length === 0 && anchor.parentNode.previousSibling) {
            var ahr = anchor.parentNode.previousSibling;
            txt = _.trim(ahr.textContent || "");
            lasttxt = txt.substring(txt.length - 1, 1);
            ctx.formulaCache.rangeSetValueTo = anchor.parentNode;
        }
        else {
            lasttxt = txt.substring(anchorOffset - 1, 1);
            ctx.formulaCache.rangeSetValueTo = anchor.parentNode;
        }
        if ((istooltip && (lasttxt === "(" || lasttxt === ",")) ||
            (!istooltip &&
                (lasttxt === "(" ||
                    lasttxt === "," ||
                    lasttxt === "=" ||
                    lasttxt in operatorjson ||
                    lasttxt === "&"))) {
            return true;
        }
    }
    else if (anchorElement.id === "luckysheet-rich-text-editor" ||
        anchorElement.id === "luckysheet-functionbox-cell") {
        var txt = _.trim((_b = _.last(anchorElement.querySelectorAll("span"))) === null || _b === void 0 ? void 0 : _b.innerText);
        ctx.formulaCache.rangeSetValueTo = _.last(anchorElement.querySelectorAll("span"));
        if (txt.length === 0 && anchorElement.querySelectorAll("span").length > 1) {
            var ahr = anchorElement.querySelectorAll("span");
            txt = _.trim(ahr[ahr.length - 2].innerText);
            txt = _.trim(ahr[ahr.length - 2].innerText);
            ctx.formulaCache.rangeSetValueTo = ahr;
        }
        var lasttxt = txt.substring(txt.length - 1, 1);
        if ((istooltip && (lasttxt === "(" || lasttxt === ",")) ||
            (!istooltip &&
                (lasttxt === "(" ||
                    lasttxt === "," ||
                    lasttxt === "=" ||
                    lasttxt in operatorjson ||
                    lasttxt === "&"))) {
            return true;
        }
    }
    else if (parentElement.id === "luckysheet-rich-text-editor" ||
        parentElement.id === "luckysheet-functionbox-cell" ||
        anchorOffset === 0) {
        if (anchorOffset === 0) {
            anchor = anchor.parentNode;
        }
        if (!anchor)
            return false;
        if (((_c = anchor.previousSibling) === null || _c === void 0 ? void 0 : _c.textContent) == null)
            return false;
        if (anchor.previousSibling) {
            var txt = _.trim(anchor.previousSibling.textContent);
            var lasttxt = txt.substring(txt.length - 1, 1);
            ctx.formulaCache.rangeSetValueTo = anchor.previousSibling;
            if ((istooltip && (lasttxt === "(" || lasttxt === ",")) ||
                (!istooltip &&
                    (lasttxt === "(" ||
                        lasttxt === "," ||
                        lasttxt === "=" ||
                        lasttxt in operatorjson ||
                        lasttxt === "&"))) {
                return true;
            }
        }
    }
    return false;
}
export function functionStrChange(txt, type, rc, orient, stindex, step) {
    if (!txt) {
        return "";
    }
    if (txt.substring(0, 1) === "=") {
        txt = txt.substring(1);
    }
    var funcstack = txt.split("");
    var i = 0;
    var str = "";
    var function_str = "";
    var matchConfig = {
        bracket: 0,
        comma: 0,
        squote: 0,
        dquote: 0, // 双引号
    };
    while (i < funcstack.length) {
        var s = funcstack[i];
        if (s === "(" && matchConfig.dquote === 0) {
            matchConfig.bracket += 1;
            if (str.length > 0) {
                function_str += "".concat(str, "(");
            }
            else {
                function_str += "(";
            }
            str = "";
        }
        else if (s === ")" && matchConfig.dquote === 0) {
            matchConfig.bracket -= 1;
            function_str += "".concat(functionStrChange(str, type, rc, orient, stindex, step), ")");
            str = "";
        }
        else if (s === '"' && matchConfig.squote === 0) {
            if (matchConfig.dquote > 0) {
                function_str += "".concat(str, "\"");
                matchConfig.dquote -= 1;
                str = "";
            }
            else {
                matchConfig.dquote += 1;
                str += '"';
            }
        }
        else if (s === "," && matchConfig.dquote === 0) {
            function_str += "".concat(functionStrChange(str, type, rc, orient, stindex, step), ",");
            str = "";
        }
        else if (s === "&" && matchConfig.dquote === 0) {
            if (str.length > 0) {
                function_str += "".concat(functionStrChange(str, type, rc, orient, stindex, step), "&");
                str = "";
            }
            else {
                function_str += "&";
            }
        }
        else if (s in operatorjson && matchConfig.dquote === 0) {
            var s_next = "";
            if (i + 1 < funcstack.length) {
                s_next = funcstack[i + 1];
            }
            var p = i - 1;
            var s_pre = null;
            if (p >= 0) {
                do {
                    s_pre = funcstack[(p -= 1)];
                } while (p >= 0 && s_pre === " ");
            }
            if (s + s_next in operatorjson) {
                if (str.length > 0) {
                    function_str +=
                        functionStrChange(str, type, rc, orient, stindex, step) +
                            s +
                            s_next;
                    str = "";
                }
                else {
                    function_str += s + s_next;
                }
                i += 1;
            }
            else if (!/[^0-9]/.test(s_next) &&
                s === "-" &&
                (s_pre === "(" ||
                    s_pre == null ||
                    s_pre === "," ||
                    s_pre === " " ||
                    s_pre in operatorjson)) {
                str += s;
            }
            else {
                if (str.length > 0) {
                    function_str +=
                        functionStrChange(str, type, rc, orient, stindex, step) + s;
                    str = "";
                }
                else {
                    function_str += s;
                }
            }
        }
        else {
            str += s;
        }
        if (i === funcstack.length - 1) {
            if (iscelldata(_.trim(str))) {
                function_str += functionStrChange_range(_.trim(str), type, rc, orient, stindex, step);
            }
            else {
                function_str += _.trim(str);
            }
        }
        i += 1;
    }
    return function_str;
}
export function rangeSetValue(ctx, cellInput, selected, fxInput) {
    var _a, _b, _c, _d, _e;
    var $editor = cellInput;
    var $copyTo = fxInput;
    if (((_a = document.activeElement) === null || _a === void 0 ? void 0 : _a.id) === "luckysheet-functionbox-cell") {
        $editor = fxInput;
        $copyTo = cellInput;
    }
    var range = "";
    var rf = selected.row[0];
    var cf = selected.column[0];
    if (ctx.config.merge != null && "".concat(rf, "_").concat(cf) in ctx.config.merge) {
        range = getRangetxt(ctx, ctx.currentSheetId, {
            column: [cf, cf],
            row: [rf, rf],
        }, ctx.formulaCache.rangetosheet);
    }
    else {
        range = getRangetxt(ctx, ctx.currentSheetId, selected, ctx.formulaCache.rangetosheet);
    }
    // let $editor;
    if (!israngeseleciton(ctx) &&
        (ctx.formulaCache.rangestart ||
            ctx.formulaCache.rangedrag_column_start ||
            ctx.formulaCache.rangedrag_row_start)) {
        //   if (
        //     $("#luckysheet-search-formula-parm").is(":visible") ||
        //     $("#luckysheet-search-formula-parm-select").is(":visible")
        //   ) {
        //     // 公式参数框选取范围
        //     $editor = $("#luckysheet-rich-text-editor");
        //     $("#luckysheet-search-formula-parm-select-input").val(range);
        //     $("#luckysheet-search-formula-parm .parmBox")
        //       .eq(formulaCache.data_parm_index)
        //       .find(".txt input")
        //       .val(range);
        //     // 参数对应值显示
        //     const txtdata = luckysheet_getcelldata(range).data;
        //     if (txtdata instanceof Array) {
        //       // 参数为多个单元格选区
        //       const txtArr = [];
        //       for (let i = 0; i < txtdata.length; i += 1) {
        //         for (let j = 0; j < txtdata[i].length; j += 1) {
        //           if (txtdata[i][j] == null) {
        //             txtArr.push(null);
        //           } else {
        //             txtArr.push(txtdata[i][j].v);
        //           }
        //         }
        //       }
        //       $("#luckysheet-search-formula-parm .parmBox")
        //         .eq(formulaCache.data_parm_index)
        //         .find(".val")
        //         .text(` = {${txtArr.join(",")}}`);
        //     } else {
        //       // 参数为单个单元格选区
        //       $("#luckysheet-search-formula-parm .parmBox")
        //         .eq(formulaCache.data_parm_index)
        //         .find(".val")
        //         .text(` = {${txtdata.v}}`);
        //     }
        //     // 计算结果显示
        //     let isVal = true; // 参数不为空
        //     const parmValArr = []; // 参数值集合
        //     let lvi = -1; // 最后一个有值的参数索引
        //     $("#luckysheet-search-formula-parm .parmBox").each(function (i, e) {
        //       const parmtxt = $(e).find(".txt input").val();
        //       if (
        //         parmtxt === "" &&
        //         $(e).find(".txt input").attr("data_parm_require") === "m"
        //       ) {
        //         isVal = false;
        //       }
        //       if (parmtxt !== "") {
        //         lvi = i;
        //       }
        //     });
        // 单元格显示
        //     let functionHtmlTxt;
        //     if (lvi === -1) {
        //       functionHtmlTxt = `=${$(
        //         "#luckysheet-search-formula-parm .luckysheet-modal-dialog-title-text"
        //       ).text()}()`;
        //     } else if (lvi === 0) {
        //       functionHtmlTxt = `=${$(
        //         "#luckysheet-search-formula-parm .luckysheet-modal-dialog-title-text"
        //       ).text()}(${$("#luckysheet-search-formula-parm .parmBox")
        //         .eq(0)
        //         .find(".txt input")
        //         .val()})`;
        //     } else {
        //       for (let j = 0; j <= lvi; j += 1) {
        //         parmValArr.push(
        //           $("#luckysheet-search-formula-parm .parmBox")
        //             .eq(j)
        //             .find(".txt input")
        //             .val()
        //         );
        //       }
        //       functionHtmlTxt = `=${$(
        //         "#luckysheet-search-formula-parm .luckysheet-modal-dialog-title-text"
        //       ).text()}(${parmValArr.join(",")})`;
        //     }
        //     const function_str = functionHTMLGenerate(functionHtmlTxt);
        //     $("#luckysheet-rich-text-editor").html(function_str);
        //     $("#luckysheet-functionbox-cell").html(
        //       $("#luckysheet-rich-text-editor").html()
        //     );
        //     if (isVal) {
        //       // 公式计算
        //       const fp = _.trim(
        //         functionParserExe($("#luckysheet-rich-text-editor").text())
        //       );
        //       const result = new Function(`return ${fp}`)();
        //       $("#luckysheet-search-formula-parm .result span").text(result);
        //     }
        //   } else {
        // const currSelection = window.getSelection();
        // const anchorOffset = currSelection!.anchorNode;
        // $editor = $(anchorOffset).closest("div");
        // const $span = $editor
        //   .find(`span[rangeindex='${formulaCache.rangechangeindex}']`)
        //   .html(range);
        var span = $editor.querySelector("span[rangeindex='".concat(ctx.formulaCache.rangechangeindex, "']"));
        if (span) {
            span.innerHTML = range;
            setCaretPosition(ctx, span, 0, range.length);
        }
        //   }
    }
    else {
        var function_str = "<span class=\"fortune-formula-functionrange-cell\" rangeindex=\"".concat(functionHTMLIndex, "\" dir=\"auto\" style=\"color:").concat(colors[functionHTMLIndex], ";\">").concat(range, "</span>");
        var newEle = parseElement(function_str);
        var refEle = ctx.formulaCache.rangeSetValueTo;
        if (refEle && refEle.parentNode) {
            var leftPar = (_b = document.getElementsByClassName("luckysheet-formula-text-lpar")) === null || _b === void 0 ? void 0 : _b[0];
            // handle case when user autocompletes the formula
            if ((_c = leftPar === null || leftPar === void 0 ? void 0 : leftPar.parentElement) === null || _c === void 0 ? void 0 : _c.classList.contains("luckysheet-formula-text-color")) {
                (_e = (_d = document
                    .getElementsByClassName("luckysheet-formula-text-lpar")) === null || _d === void 0 ? void 0 : _d[0].parentNode) === null || _e === void 0 ? void 0 : _e.appendChild(newEle);
            }
            else {
                refEle.parentNode.insertBefore(newEle, refEle.nextSibling);
            }
        }
        else {
            $editor.appendChild(newEle);
        }
        ctx.formulaCache.rangechangeindex = functionHTMLIndex;
        var span = $editor.querySelector("span[rangeindex='".concat(ctx.formulaCache.rangechangeindex, "']"));
        setCaretPosition(ctx, span, 0, range.length);
        functionHTMLIndex += 1;
    }
    if ($copyTo)
        $copyTo.innerHTML = $editor.innerHTML;
}
export function onFormulaRangeDragEnd(ctx) {
    if (ctx.formulaCache.func_selectedrange) {
        var _a = ctx.formulaCache.func_selectedrange, left = _a.left_move, top_1 = _a.top_move, width = _a.width_move, height = _a.height_move;
        if (left != null &&
            top_1 != null &&
            width != null &&
            height != null &&
            (ctx.formulaCache.rangestart ||
                ctx.formulaCache.rangedrag_column_start ||
                ctx.formulaCache.rangedrag_row_start))
            ctx.formulaRangeSelect = {
                rangeIndex: ctx.formulaCache.rangeIndex || 0,
                left: left,
                top: top_1,
                width: width,
                height: height,
            };
    }
    ctx.formulaCache.selectingRangeIndex = -1;
}
function setRangeSelect(container, left, top, height, width) {
    var rangeElement = container.querySelector(".fortune-formula-functionrange-select");
    if (rangeElement == null)
        return;
    rangeElement.style.left = "".concat(left, "px");
    rangeElement.style.top = "".concat(top, "px");
    rangeElement.style.height = "".concat(height, "px");
    rangeElement.style.width = "".concat(width, "px");
}
export function rangeDrag(ctx, e, cellInput, scrollLeft, scrollTop, container, fxInput) {
    var func_selectedrange = ctx.formulaCache.func_selectedrange;
    if (!func_selectedrange ||
        func_selectedrange.left == null ||
        func_selectedrange.height == null ||
        func_selectedrange.top == null ||
        func_selectedrange.width == null)
        return;
    var rect = container.getBoundingClientRect();
    var x = e.pageX - rect.left - ctx.rowHeaderWidth + scrollLeft;
    var y = e.pageY - rect.top - ctx.columnHeaderHeight + scrollTop;
    var _a = rowLocation(y, ctx.visibledatarow), row_pre = _a[0], row = _a[1], row_index = _a[2];
    var _b = colLocation(x, ctx.visibledatacolumn), col_pre = _b[0], col = _b[1], col_index = _b[2];
    var top = 0;
    var height = 0;
    var rowseleted = [];
    if (func_selectedrange.top > row_pre) {
        top = row_pre;
        height = func_selectedrange.top + func_selectedrange.height - row_pre;
        rowseleted = [row_index, func_selectedrange.row[1]];
    }
    else if (func_selectedrange.top === row_pre) {
        top = row_pre;
        height = func_selectedrange.top + func_selectedrange.height - row_pre;
        rowseleted = [row_index, func_selectedrange.row[0]];
    }
    else {
        top = func_selectedrange.top;
        height = row - func_selectedrange.top - 1;
        rowseleted = [func_selectedrange.row[0], row_index];
    }
    var left = 0;
    var width = 0;
    var columnseleted = [];
    if (func_selectedrange.left > col_pre) {
        left = col_pre;
        width = func_selectedrange.left + func_selectedrange.width - col_pre;
        columnseleted = [col_index, func_selectedrange.column[1]];
    }
    else if (func_selectedrange.left === col_pre) {
        left = col_pre;
        width = func_selectedrange.left + func_selectedrange.width - col_pre;
        columnseleted = [col_index, func_selectedrange.column[0]];
    }
    else {
        left = func_selectedrange.left;
        width = col - func_selectedrange.left - 1;
        columnseleted = [func_selectedrange.column[0], col_index];
    }
    // rowseleted[0] = luckysheetFreezen.changeFreezenIndex(rowseleted[0], "h");
    // rowseleted[1] = luckysheetFreezen.changeFreezenIndex(rowseleted[1], "h");
    // columnseleted[0] = luckysheetFreezen.changeFreezenIndex(
    //   columnseleted[0],
    //   "v"
    // );
    // columnseleted[1] = luckysheetFreezen.changeFreezenIndex(
    //   columnseleted[1],
    //   "v"
    // );
    var changeparam = mergeMoveMain(ctx, columnseleted, rowseleted, func_selectedrange, top, height, left, width);
    if (changeparam != null) {
        // @ts-ignore
        columnseleted = changeparam[0], rowseleted = changeparam[1], top = changeparam[2], height = changeparam[3], left = changeparam[4], width = changeparam[5];
    }
    func_selectedrange.row = rowseleted;
    func_selectedrange.column = columnseleted;
    func_selectedrange.left_move = left;
    func_selectedrange.width_move = width;
    func_selectedrange.top_move = top;
    func_selectedrange.height_move = height;
    // luckysheet_count_show(left, top, width, height, rowseleted, columnseleted);
    // if ($("#luckysheet-ifFormulaGenerator-multiRange-dialog").is(":visible")) {
    //   // if公式生成器 选择范围
    //   const range = getRangetxt(
    //     ctx,
    //     ctx.currentSheetId,
    //     { row: rowseleted, column: columnseleted },
    //     ctx.currentSheetId
    //   );
    //   $("#luckysheet-ifFormulaGenerator-multiRange-dialog input").val(range);
    // } else {
    rangeSetValue(ctx, cellInput, {
        row: rowseleted,
        column: columnseleted,
    }, fxInput);
    setRangeSelect(container, left, top, height, width);
    // }
    // luckysheetFreezen.scrollFreezen(rowseleted, columnseleted);
    e.preventDefault();
}
export function rangeDragColumn(ctx, e, cellInput, scrollLeft, scrollTop, container, fxInput) {
    var func_selectedrange = ctx.formulaCache.func_selectedrange;
    if (!func_selectedrange ||
        func_selectedrange.left == null ||
        func_selectedrange.height == null ||
        func_selectedrange.top == null ||
        func_selectedrange.width == null)
        return;
    var mouse = mousePosition(e.pageX, e.pageY, ctx);
    var x = mouse[0] + scrollLeft;
    var visibledatarow = ctx.visibledatarow;
    var row_index = visibledatarow.length - 1;
    var row = visibledatarow[row_index];
    var row_pre = 0;
    var _a = colLocation(x, ctx.visibledatacolumn), col_pre = _a[0], col = _a[1], col_index = _a[2];
    var left = 0;
    var width = 0;
    var columnseleted = [];
    if (func_selectedrange.left > col_pre) {
        left = col_pre;
        width = func_selectedrange.left + func_selectedrange.width - col_pre;
        columnseleted = [col_index, func_selectedrange.column[1]];
    }
    else if (func_selectedrange.left === col_pre) {
        left = col_pre;
        width = func_selectedrange.left + func_selectedrange.width - col_pre;
        columnseleted = [col_index, func_selectedrange.column[0]];
    }
    else {
        left = func_selectedrange.left;
        width = col - func_selectedrange.left - 1;
        columnseleted = [func_selectedrange.column[0], col_index];
    }
    // // rowseleted[0] = luckysheetFreezen.changeFreezenIndex(rowseleted[0], "h");
    // // rowseleted[1] = luckysheetFreezen.changeFreezenIndex(rowseleted[1], "h");
    // columnseleted[0] = luckysheetFreezen.changeFreezenIndex(
    //   columnseleted[0],
    //   "v"
    // );
    // columnseleted[1] = luckysheetFreezen.changeFreezenIndex(
    //   columnseleted[1],
    //   "v"
    // );
    var changeparam = mergeMoveMain(ctx, columnseleted, [0, row_index], func_selectedrange, row_pre, row - row_pre - 1, left, width);
    if (changeparam != null) {
        // @ts-ignore
        columnseleted = changeparam[0], left = changeparam[4], width = changeparam[5];
        // rowseleted= changeparam[1];
        // top = changeparam[2];
        // height = changeparam[3];
        // left = changeparam[4];
        // width = changeparam[5];
    }
    func_selectedrange.column = columnseleted;
    func_selectedrange.left_move = left;
    func_selectedrange.width_move = width;
    // luckysheet_count_show(
    //   left,
    //   row_pre,
    //   width,
    //   row - row_pre - 1,
    //   [0, row_index],
    //   columnseleted
    // );
    rangeSetValue(ctx, cellInput, {
        row: [null, null],
        column: columnseleted,
    }, fxInput);
    setRangeSelect(container, left, row_pre, row - row_pre - 1, width);
    // luckysheetFreezen.scrollFreezen([0, row_index], columnseleted);
}
export function rangeDragRow(ctx, e, cellInput, scrollLeft, scrollTop, container, fxInput) {
    var func_selectedrange = ctx.formulaCache.func_selectedrange;
    if (!func_selectedrange ||
        func_selectedrange.left == null ||
        func_selectedrange.height == null ||
        func_selectedrange.top == null ||
        func_selectedrange.width == null)
        return;
    var mouse = mousePosition(e.pageX, e.pageY, ctx);
    var y = mouse[1] + scrollTop;
    var _a = rowLocation(y, ctx.visibledatarow), row_pre = _a[0], row = _a[1], row_index = _a[2];
    var visibledatacolumn = ctx.visibledatacolumn;
    var col_index = visibledatacolumn.length - 1;
    var col = visibledatacolumn[col_index];
    var col_pre = 0;
    var top = 0;
    var height = 0;
    var rowseleted = [];
    if (func_selectedrange.top > row_pre) {
        top = row_pre;
        height = func_selectedrange.top + func_selectedrange.height - row_pre;
        rowseleted = [row_index, func_selectedrange.row[1]];
    }
    else if (func_selectedrange.top === row_pre) {
        top = row_pre;
        height = func_selectedrange.top + func_selectedrange.height - row_pre;
        rowseleted = [row_index, func_selectedrange.row[0]];
    }
    else {
        top = func_selectedrange.top;
        height = row - func_selectedrange.top - 1;
        rowseleted = [func_selectedrange.row[0], row_index];
    }
    // rowseleted[0] = luckysheetFreezen.changeFreezenIndex(rowseleted[0], "h");
    // rowseleted[1] = luckysheetFreezen.changeFreezenIndex(rowseleted[1], "h");
    // // columnseleted[0] = luckysheetFreezen.changeFreezenIndex(columnseleted[0], "v");
    // // columnseleted[1] = luckysheetFreezen.changeFreezenIndex(columnseleted[1], "v");
    var changeparam = mergeMoveMain(ctx, [0, col_index], rowseleted, func_selectedrange, top, height, col_pre, col - col_pre - 1);
    if (changeparam != null) {
        // @ts-ignore
        rowseleted = changeparam[1], top = changeparam[2], height = changeparam[3];
    }
    func_selectedrange.row = rowseleted;
    func_selectedrange.top_move = top;
    func_selectedrange.height_move = height;
    // luckysheet_count_show(col_pre, top, col - col_pre - 1, height, rowseleted, [
    //   0,
    //   col_index,
    // ]);
    rangeSetValue(ctx, cellInput, {
        row: rowseleted,
        column: [null, null],
    }, fxInput);
    setRangeSelect(container, col_pre, top, height, col - col_pre - 1);
    // luckysheetFreezen.scrollFreezen(rowseleted, [0, col_index]);
}
function updateparam(orient, txt, step) {
    var val = txt.split("!");
    var rangetxt;
    var prefix = "";
    if (val.length > 1) {
        rangetxt = val[1];
        prefix = "".concat(val[0], "!");
    }
    else {
        rangetxt = val[0];
    }
    if (rangetxt.indexOf(":") === -1) {
        var row_3 = parseInt(rangetxt.replace(/[^0-9]/g, ""), 10);
        var col_3 = columnCharToIndex(rangetxt.replace(/[^A-Za-z]/g, ""));
        var freezonFuc = isfreezonFuc(rangetxt);
        var $row = freezonFuc[0] ? "$" : "";
        var $col = freezonFuc[1] ? "$" : "";
        if (orient === "u" && !freezonFuc[0]) {
            row_3 -= step;
        }
        else if (orient === "r" && !freezonFuc[1]) {
            col_3 += step;
        }
        else if (orient === "l" && !freezonFuc[1]) {
            col_3 -= step;
        }
        else if (orient === "d" && !freezonFuc[0]) {
            row_3 += step;
        }
        if (!Number.isNaN(row_3) && !Number.isNaN(col_3)) {
            return prefix + $col + indexToColumnChar(col_3) + $row + row_3;
        }
        if (!Number.isNaN(row_3)) {
            return prefix + $row + row_3;
        }
        if (!Number.isNaN(col_3)) {
            return prefix + $col + indexToColumnChar(col_3);
        }
        return txt;
    }
    rangetxt = rangetxt.split(":");
    var row = [];
    var col = [];
    row[0] = parseInt(rangetxt[0].replace(/[^0-9]/g, ""), 10);
    row[1] = parseInt(rangetxt[1].replace(/[^0-9]/g, ""), 10);
    if (row[0] > row[1]) {
        return txt;
    }
    col[0] = columnCharToIndex(rangetxt[0].replace(/[^A-Za-z]/g, ""));
    col[1] = columnCharToIndex(rangetxt[1].replace(/[^A-Za-z]/g, ""));
    if (col[0] > col[1]) {
        return txt;
    }
    var freezonFuc0 = isfreezonFuc(rangetxt[0]);
    var freezonFuc1 = isfreezonFuc(rangetxt[1]);
    var $row0 = freezonFuc0[0] ? "$" : "";
    var $col0 = freezonFuc0[1] ? "$" : "";
    var $row1 = freezonFuc1[0] ? "$" : "";
    var $col1 = freezonFuc1[1] ? "$" : "";
    if (orient === "u") {
        if (!freezonFuc0[0]) {
            row[0] -= step;
        }
        if (!freezonFuc1[0]) {
            row[1] -= step;
        }
    }
    else if (orient === "r") {
        if (!freezonFuc0[1]) {
            col[0] += step;
        }
        if (!freezonFuc1[1]) {
            col[1] += step;
        }
    }
    else if (orient === "l") {
        if (!freezonFuc0[1]) {
            col[0] -= step;
        }
        if (!freezonFuc1[1]) {
            col[1] -= step;
        }
    }
    else if (orient === "d") {
        if (!freezonFuc0[0]) {
            row[0] += step;
        }
        if (!freezonFuc1[0]) {
            row[1] += step;
        }
    }
    if (row[0] < 0 || col[0] < 0) {
        return error.r;
    }
    if (Number.isNaN(col[0]) && Number.isNaN(col[1])) {
        return "".concat(prefix + $row0 + row[0], ":").concat($row1).concat(row[1]);
    }
    if (Number.isNaN(row[0]) && Number.isNaN(row[1])) {
        return "".concat(prefix + $col0 + indexToColumnChar(col[0]), ":").concat($col1).concat(indexToColumnChar(col[1]));
    }
    return "".concat(prefix + $col0 + indexToColumnChar(col[0]) + $row0 + row[0], ":").concat($col1).concat(indexToColumnChar(col[1])).concat($row1).concat(row[1]);
}
function downparam(txt, step) {
    return updateparam("d", txt, step);
}
function upparam(txt, step) {
    return updateparam("u", txt, step);
}
function leftparam(txt, step) {
    return updateparam("l", txt, step);
}
function rightparam(txt, step) {
    return updateparam("r", txt, step);
}
export function functionCopy(ctx, txt, mode, step) {
    if (mode == null) {
        mode = "down";
    }
    if (step == null) {
        step = 1;
    }
    if (txt.substring(0, 1) === "=") {
        txt = txt.substring(1);
    }
    var funcstack = txt.split("");
    var i = 0;
    var str = "";
    var function_str = "";
    var matchConfig = {
        bracket: 0,
        comma: 0,
        squote: 0,
        dquote: 0,
    };
    while (i < funcstack.length) {
        var s = funcstack[i];
        if (s === "(" && matchConfig.dquote === 0) {
            matchConfig.bracket += 1;
            if (str.length > 0) {
                function_str += "".concat(str, "(");
            }
            else {
                function_str += "(";
            }
            str = "";
        }
        else if (s === ")" && matchConfig.dquote === 0) {
            matchConfig.bracket -= 1;
            function_str += "".concat(functionCopy(ctx, str, mode, step), ")");
            str = "";
        }
        else if (s === '"' && matchConfig.squote === 0) {
            if (matchConfig.dquote > 0) {
                function_str += "".concat(str, "\"");
                matchConfig.dquote -= 1;
                str = "";
            }
            else {
                matchConfig.dquote += 1;
                str += '"';
            }
        }
        else if (s === "," && matchConfig.dquote === 0) {
            function_str += "".concat(functionCopy(ctx, str, mode, step), ",");
            str = "";
        }
        else if (s === "&" && matchConfig.dquote === 0) {
            if (str.length > 0) {
                function_str += "".concat(functionCopy(ctx, str, mode, step), "&");
                str = "";
            }
            else {
                function_str += "&";
            }
        }
        else if (s in operatorjson && matchConfig.dquote === 0) {
            var s_next = "";
            if (i + 1 < funcstack.length) {
                s_next = funcstack[i + 1];
            }
            var p = i - 1;
            var s_pre = null;
            if (p >= 0) {
                do {
                    s_pre = funcstack[p];
                    p -= 1;
                } while (p >= 0 && s_pre === " ");
            }
            if (s + s_next in operatorjson) {
                if (str.length > 0) {
                    function_str += functionCopy(ctx, str, mode, step) + s + s_next;
                    str = "";
                }
                else {
                    function_str += s + s_next;
                }
                i += 1;
            }
            else if (!/[^0-9]/.test(s_next) &&
                s === "-" &&
                (s_pre === "(" ||
                    s_pre == null ||
                    s_pre === "," ||
                    s_pre === " " ||
                    s_pre in operatorjson)) {
                str += s;
            }
            else {
                if (str.length > 0) {
                    function_str += functionCopy(ctx, str, mode, step) + s;
                    str = "";
                }
                else {
                    function_str += s;
                }
            }
        }
        else {
            str += s;
        }
        if (i === funcstack.length - 1) {
            if (iscelldata(_.trim(str))) {
                if (mode === "down") {
                    function_str += downparam(_.trim(str), step);
                }
                else if (mode === "up") {
                    function_str += upparam(_.trim(str), step);
                }
                else if (mode === "left") {
                    function_str += leftparam(_.trim(str), step);
                }
                else if (mode === "right") {
                    function_str += rightparam(_.trim(str), step);
                }
            }
            else {
                function_str += _.trim(str);
            }
        }
        i += 1;
    }
    return function_str;
}
