import _ from "lodash";
import { getFlowdata } from "../context";
import { getCellValue, setCellValue } from "./cell";
// 生成二维数组
export function getNullData(rlen, clen) {
    var arr = [];
    for (var r = 0; r < rlen; r += 1) {
        var rowArr = [];
        for (var c = 0; c < clen; c += 1) {
            rowArr.push("");
        }
        arr.push(rowArr);
    }
    return arr;
}
// 批量更新数据到表格
export function updateMoreCell(r, c, dataMatrix, ctx) {
    if (ctx.allowEdit === false)
        return;
    var flowdata = getFlowdata(ctx);
    dataMatrix.forEach(function (datas, i) {
        datas.forEach(function (data, j) {
            var v = dataMatrix[i][j];
            setCellValue(ctx, r + i, c + j, flowdata, v);
        });
    });
    // jfrefreshgrid(d, range);
    // selectHightlightShow();
}
// 处理分隔符
export function getRegStr(regStr, splitSymbols) {
    regStr = "";
    var mark = 0;
    for (var i = 0; i < splitSymbols.length; i += 1) {
        var split = splitSymbols[i];
        var inputNode = split.childNodes[0];
        if (inputNode.checked) {
            var id = inputNode.id;
            if (id === "Tab") {
                // Tab键
                regStr += "\\t";
                mark += 1;
            }
            else if (id === "semicolon") {
                // 分号
                if (mark > 0) {
                    regStr += "|";
                }
                regStr += ";";
                mark = 1;
            }
            else if (id === "comma") {
                // 逗号
                if (mark > 0) {
                    regStr += "|";
                }
                regStr += ",";
                mark += 1;
            }
            else if (id === "space") {
                // 空格
                if (mark > 0) {
                    regStr += "|";
                }
                regStr += "\\s";
                mark += 1;
            }
            else if (id === "splitsimple") {
                // 连续分隔符号视为单个处理
                regStr = "[".concat(regStr, "]+");
            }
            else if (id === "other") {
                // 其他
                var txt = split.childNodes[2].value;
                if (txt !== "") {
                    if (mark > 0) {
                        regStr += "|";
                    }
                    regStr += txt;
                }
            }
        }
    }
    return regStr;
}
// 获得分割数据
export function getDataArr(regStr, ctx) {
    var arr = [];
    var r1 = ctx.luckysheet_select_save[0].row[0];
    var r2 = ctx.luckysheet_select_save[0].row[1];
    var c = ctx.luckysheet_select_save[0].column[0];
    var data = getFlowdata(ctx);
    if (!_.isNull(regStr) && regStr !== "") {
        var reg = new RegExp(regStr, "g");
        var dataArr = [];
        for (var r = r1; r <= r2; r += 1) {
            var rowArr = [];
            var cell = data[r][c];
            var value = void 0;
            if (!_.isNull(cell) && !_.isNull(cell.m)) {
                value = cell.m;
            }
            else {
                value = getCellValue(r, c, data);
            }
            if (_.isNull(value) || _.isUndefined(value)) {
                value = "";
            }
            rowArr = value.toString().split(reg);
            dataArr.push(rowArr);
        }
        var rlen = dataArr.length;
        var clen = 0;
        for (var i = 0; i < rlen; i += 1) {
            if (dataArr[i].length > clen) {
                clen = dataArr[i].length;
            }
        }
        arr = getNullData(rlen, clen);
        for (var i = 0; i < arr.length; i += 1) {
            for (var j = 0; j < arr[0].length; j += 1) {
                if (dataArr[i][j] != null) {
                    arr[i][j] = dataArr[i][j];
                }
            }
        }
    }
    else {
        for (var r = r1; r <= r2; r += 1) {
            var rowArr = [];
            var cell = data[r][c];
            var value = void 0;
            if (!_.isNull(cell) && !_.isNull(cell.m)) {
                value = cell.m;
            }
            else {
                value = getCellValue(r, c, data);
            }
            if (_.isNull(value)) {
                value = "";
            }
            rowArr.push(value);
            arr.push(rowArr);
        }
    }
    return arr;
}
