import _ from "lodash";
import { getFlowdata } from "../context";
import { locale } from "../locale";
import { chatatABC, getRegExpStr, getSheetIndex, isAllowEdit, replaceHtml, } from "../utils";
import { setCellValue } from "./cell";
import { valueShowEs } from "./format";
import { normalizeSelection, scrollToHighlightCell } from "./selection";
export function getSearchIndexArr(searchText, range, flowdata, _a) {
    var _b = _a === void 0 ? {
        regCheck: false,
        wordCheck: false,
        caseCheck: false,
    } : _a, regCheck = _b.regCheck, wordCheck = _b.wordCheck, caseCheck = _b.caseCheck;
    var arr = [];
    var obj = {};
    for (var s = 0; s < range.length; s += 1) {
        var r1 = range[s].row[0];
        var r2 = range[s].row[1];
        var c1 = range[s].column[0];
        var c2 = range[s].column[1];
        for (var r = r1; r <= r2; r += 1) {
            for (var c = c1; c <= c2; c += 1) {
                var cell = flowdata[r][c];
                if (cell != null) {
                    var value = valueShowEs(r, c, flowdata);
                    if (value === 0) {
                        value = value.toString();
                    }
                    if (value != null && value !== "") {
                        value = value.toString();
                        // 1. 勾选整词 直接匹配
                        // 2. 勾选了正则 结合是否勾选 构造正则
                        // 3. 什么都没选 用字符串 indexOf 匹配
                        if (wordCheck) {
                            // 整词
                            if (caseCheck) {
                                if (searchText === value) {
                                    if (!("".concat(r, "_").concat(c) in obj)) {
                                        _.set(obj, "".concat(r, "_").concat(c), 0);
                                        arr.push({ r: r, c: c });
                                    }
                                }
                            }
                            else {
                                var txt = searchText.toLowerCase();
                                if (txt === value.toLowerCase()) {
                                    if (!("".concat(r, "_").concat(c) in obj)) {
                                        _.set(obj, "".concat(r, "_").concat(c), 0);
                                        arr.push({ r: r, c: c });
                                    }
                                }
                            }
                        }
                        else if (regCheck) {
                            // 正则表达式
                            var reg = void 0;
                            // 是否区分大小写
                            if (caseCheck) {
                                reg = new RegExp(getRegExpStr(searchText), "g");
                            }
                            else {
                                reg = new RegExp(getRegExpStr(searchText), "ig");
                            }
                            if (reg.test(value)) {
                                if (!("".concat(r, "_").concat(c) in obj)) {
                                    _.set(obj, "".concat(r, "_").concat(c), 0);
                                    arr.push({ r: r, c: c });
                                }
                            }
                        }
                        else {
                            if (caseCheck) {
                                if (~value.indexOf(searchText)) {
                                    if (!("".concat(r, "_").concat(c) in obj)) {
                                        _.set(obj, "".concat(r, "_").concat(c), 0);
                                        arr.push({ r: r, c: c });
                                    }
                                }
                            }
                            else {
                                if (~value.toLowerCase().indexOf(searchText.toLowerCase())) {
                                    if (!("".concat(r, "_").concat(c) in obj)) {
                                        _.set(obj, "".concat(r, "_").concat(c), 0);
                                        arr.push({ r: r, c: c });
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    return arr;
}
export function searchNext(ctx, searchText, checkModes) {
    var _a, _b;
    var findAndReplace = locale(ctx).findAndReplace;
    var flowdata = getFlowdata(ctx);
    if (searchText === "" || searchText == null || flowdata == null) {
        return findAndReplace.searchInputTip;
    }
    var range;
    if (_.size(ctx.luckysheet_select_save) === 0 ||
        (((_a = ctx.luckysheet_select_save) === null || _a === void 0 ? void 0 : _a.length) === 1 &&
            ctx.luckysheet_select_save[0].row[0] ===
                ctx.luckysheet_select_save[0].row[1] &&
            ctx.luckysheet_select_save[0].column[0] ===
                ctx.luckysheet_select_save[0].column[1])) {
        range = [
            {
                row: [0, flowdata.length - 1],
                column: [0, flowdata[0].length - 1],
                row_focus: 0,
                column_focus: 0,
            },
        ];
    }
    else {
        range = _.assign([], ctx.luckysheet_select_save);
    }
    var searchIndexArr = getSearchIndexArr(searchText, range, flowdata, checkModes);
    if (searchIndexArr.length === 0) {
        return findAndReplace.noFindTip;
    }
    var count = 0;
    if (_.size(ctx.luckysheet_select_save) === 0 ||
        (((_b = ctx.luckysheet_select_save) === null || _b === void 0 ? void 0 : _b.length) === 1 &&
            ctx.luckysheet_select_save[0].row[0] ===
                ctx.luckysheet_select_save[0].row[1] &&
            ctx.luckysheet_select_save[0].column[0] ===
                ctx.luckysheet_select_save[0].column[1])) {
        if (_.size(ctx.luckysheet_select_save) === 0) {
            count = 0;
        }
        else {
            for (var i = 0; i < searchIndexArr.length; i += 1) {
                if (searchIndexArr[i].r === ctx.luckysheet_select_save[0].row[0] &&
                    searchIndexArr[i].c === ctx.luckysheet_select_save[0].column[0]) {
                    if (i === searchIndexArr.length - 1) {
                        count = 0;
                    }
                    else {
                        count = i + 1;
                    }
                    break;
                }
            }
        }
        ctx.luckysheet_select_save = normalizeSelection(ctx, [
            {
                row: [searchIndexArr[count].r, searchIndexArr[count].r],
                column: [searchIndexArr[count].c, searchIndexArr[count].c],
            },
        ]);
    }
    else {
        var rf = range[range.length - 1].row_focus;
        var cf = range[range.length - 1].column_focus;
        for (var i = 0; i < searchIndexArr.length; i += 1) {
            if (searchIndexArr[i].r === rf && searchIndexArr[i].c === cf) {
                if (i === searchIndexArr.length - 1) {
                    count = 0;
                }
                else {
                    count = i + 1;
                }
                break;
            }
        }
        for (var s = 0; s < range.length; s += 1) {
            var r1 = range[s].row[0];
            var r2 = range[s].row[1];
            var c1 = range[s].column[0];
            var c2 = range[s].column[1];
            if (searchIndexArr[count].r >= r1 &&
                searchIndexArr[count].r <= r2 &&
                searchIndexArr[count].c >= c1 &&
                searchIndexArr[count].c <= c2) {
                var obj = range[s];
                obj.row_focus = searchIndexArr[count].r;
                obj.column_focus = searchIndexArr[count].c;
                range.splice(s, 1);
                range.push(obj);
                break;
            }
        }
        ctx.luckysheet_select_save = range;
    }
    // selectHightlightShow();
    scrollToHighlightCell(ctx, searchIndexArr[count].r, searchIndexArr[count].c);
    return null;
}
export function searchAll(ctx, searchText, checkModes) {
    var _a, _b;
    var flowdata = getFlowdata(ctx);
    var searchResult = [];
    if (searchText === "" || searchText == null || flowdata == null) {
        return searchResult;
    }
    var range;
    if (_.size(ctx.luckysheet_select_save) === 0 ||
        (((_a = ctx.luckysheet_select_save) === null || _a === void 0 ? void 0 : _a.length) === 1 &&
            ctx.luckysheet_select_save[0].row[0] ===
                ctx.luckysheet_select_save[0].row[1] &&
            ctx.luckysheet_select_save[0].column[0] ===
                ctx.luckysheet_select_save[0].column[1])) {
        range = [
            {
                row: [0, flowdata.length - 1],
                column: [0, flowdata[0].length - 1],
            },
        ];
    }
    else {
        range = _.assign([], ctx.luckysheet_select_save);
    }
    var searchIndexArr = getSearchIndexArr(searchText, range, flowdata, checkModes);
    if (searchIndexArr.length === 0) {
        // if (isEditMode()) {
        //   alert(locale_findAndReplace.noFindTip);
        // } else {
        //   tooltip.info(locale_findAndReplace.noFindTip, "");
        // }
        return searchResult;
    }
    for (var i = 0; i < searchIndexArr.length; i += 1) {
        var value_ShowEs = valueShowEs(searchIndexArr[i].r, searchIndexArr[i].c, flowdata).toString();
        // if (value_ShowEs.indexOf("</") > -1 && value_ShowEs.indexOf(">") > -1) {
        searchResult.push({
            r: searchIndexArr[i].r,
            c: searchIndexArr[i].c,
            sheetName: (_b = ctx.luckysheetfile[getSheetIndex(ctx, ctx.currentSheetId) || 0]) === null || _b === void 0 ? void 0 : _b.name,
            sheetId: ctx.currentSheetId,
            cellPosition: "".concat(chatatABC(searchIndexArr[i].c)).concat(searchIndexArr[i].r + 1),
            value: value_ShowEs,
        });
        // } else {
        // searchAllHtml +=
        //   `<div class="boxItem" data-row="${searchIndexArr[i].r}" data-col="${searchIndexArr[i].c}" data-sheetIndex="${ctx.currentSheetIndex}">` +
        //   `<span>${
        //     ctx.luckysheetfile[getSheetIndex(ctx.currentSheetIndex)].name
        //   }</span>` +
        //   `<span>${chatatABC(searchIndexArr[i].c)}${
        //     searchIndexArr[i].r + 1
        //   }</span>` +
        //   `<span title="${value_ShowEs}">${value_ShowEs}</span>` +
        //   `</div>`;
        // }
    }
    ctx.luckysheet_select_save = normalizeSelection(ctx, [
        {
            row: [searchIndexArr[0].r, searchIndexArr[0].r],
            column: [searchIndexArr[0].c, searchIndexArr[0].c],
        },
    ]);
    return searchResult;
    // selectHightlightShow();
}
export function onSearchDialogMoveStart(globalCache, e, container) {
    var box = document.getElementById("fortune-search-replace");
    if (!box)
        return;
    // eslint-disable-next-line prefer-const
    var _a = box.getBoundingClientRect(), top = _a.top, left = _a.left, width = _a.width, height = _a.height;
    var rect = container.getBoundingClientRect();
    left -= rect.left;
    top -= rect.top;
    var initialPosition = { left: left, top: top, width: width, height: height };
    _.set(globalCache, "searchDialog.moveProps", {
        cursorMoveStartPosition: {
            x: e.pageX,
            y: e.pageY,
        },
        initialPosition: initialPosition,
    });
}
export function onSearchDialogMove(globalCache, e) {
    var searchDialog = globalCache === null || globalCache === void 0 ? void 0 : globalCache.searchDialog;
    var moveProps = searchDialog === null || searchDialog === void 0 ? void 0 : searchDialog.moveProps;
    if (moveProps == null)
        return;
    var dialog = document.getElementById("fortune-search-replace");
    var _a = moveProps.cursorMoveStartPosition, startX = _a.x, startY = _a.y;
    var _b = moveProps.initialPosition, top = _b.top, left = _b.left;
    left += e.pageX - startX;
    top += e.pageY - startY;
    if (top < 0)
        top = 0;
    dialog.style.left = "".concat(left, "px");
    dialog.style.top = "".concat(top, "px");
}
export function onSearchDialogMoveEnd(globalCache) {
    _.set(globalCache, "searchDialog.moveProps", undefined);
}
export function replace(ctx, searchText, replaceText, checkModes) {
    var _a, _b;
    var findAndReplace = locale(ctx).findAndReplace;
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit) {
        return findAndReplace.modeTip;
    }
    var flowdata = getFlowdata(ctx);
    if (searchText === "" || searchText == null || flowdata == null) {
        return findAndReplace.searchInputTip;
    }
    var range;
    if (_.size(ctx.luckysheet_select_save) === 0 ||
        (((_a = ctx.luckysheet_select_save) === null || _a === void 0 ? void 0 : _a.length) === 1 &&
            ctx.luckysheet_select_save[0].row[0] ===
                ctx.luckysheet_select_save[0].row[1] &&
            ctx.luckysheet_select_save[0].column[0] ===
                ctx.luckysheet_select_save[0].column[1])) {
        range = [
            {
                row: [0, flowdata.length - 1],
                column: [0, flowdata[0].length - 1],
            },
        ];
    }
    else {
        range = _.assign([], ctx.luckysheet_select_save);
    }
    var searchIndexArr = getSearchIndexArr(searchText, range, flowdata, checkModes);
    if (searchIndexArr.length === 0) {
        return findAndReplace.noReplceTip;
    }
    var count = null;
    var last = (_b = ctx.luckysheet_select_save) === null || _b === void 0 ? void 0 : _b[ctx.luckysheet_select_save.length - 1];
    var rf = last === null || last === void 0 ? void 0 : last.row_focus;
    var cf = last === null || last === void 0 ? void 0 : last.column_focus;
    for (var i = 0; i < searchIndexArr.length; i += 1) {
        if (searchIndexArr[i].r === rf && searchIndexArr[i].c === cf) {
            count = i;
            break;
        }
    }
    if (count == null) {
        if (searchIndexArr.length === 0) {
            return findAndReplace.noMatchTip;
        }
        count = 0;
    }
    var d = flowdata;
    var r;
    var c;
    if (checkModes.wordCheck) {
        r = searchIndexArr[count].r;
        c = searchIndexArr[count].c;
        var v = replaceText;
        // if (!checkProtectionLocked(r, c, ctx.currentSheetId)) {
        //   return;
        // }
        setCellValue(ctx, r, c, d, v);
    }
    else {
        var reg = void 0;
        if (checkModes.caseCheck) {
            reg = new RegExp(getRegExpStr(searchText), "g");
        }
        else {
            reg = new RegExp(getRegExpStr(searchText), "ig");
        }
        r = searchIndexArr[count].r;
        c = searchIndexArr[count].c;
        // if (!checkProtectionLocked(r, c, ctx.currentSheetId)) {
        //   return;
        // }
        var v = valueShowEs(r, c, d).toString().replace(reg, replaceText);
        setCellValue(ctx, r, c, d, v);
    }
    ctx.luckysheet_select_save = normalizeSelection(ctx, [
        { row: [r, r], column: [c, c] },
    ]);
    // jfrefreshgrid(d, ctx.luckysheet_select_save);
    // selectHightlightShow();
    scrollToHighlightCell(ctx, r, c);
    return null;
}
export function replaceAll(ctx, searchText, replaceText, checkModes) {
    var _a;
    var findAndReplace = locale(ctx).findAndReplace;
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit) {
        return findAndReplace.modeTip;
    }
    var flowdata = getFlowdata(ctx);
    if (searchText === "" || searchText == null || flowdata == null) {
        return findAndReplace.searchInputTip;
    }
    var range;
    if (_.size(ctx.luckysheet_select_save) === 0 ||
        (((_a = ctx.luckysheet_select_save) === null || _a === void 0 ? void 0 : _a.length) === 1 &&
            ctx.luckysheet_select_save[0].row[0] ===
                ctx.luckysheet_select_save[0].row[1] &&
            ctx.luckysheet_select_save[0].column[0] ===
                ctx.luckysheet_select_save[0].column[1])) {
        range = [
            {
                row: [0, flowdata.length - 1],
                column: [0, flowdata[0].length - 1],
            },
        ];
    }
    else {
        range = _.assign([], ctx.luckysheet_select_save);
    }
    var searchIndexArr = getSearchIndexArr(searchText, range, flowdata, checkModes);
    if (searchIndexArr.length === 0) {
        return findAndReplace.noReplceTip;
    }
    var d = flowdata;
    var replaceCount = 0;
    if (checkModes.wordCheck) {
        for (var i = 0; i < searchIndexArr.length; i += 1) {
            var r = searchIndexArr[i].r;
            var c = searchIndexArr[i].c;
            // if (!checkProtectionLocked(r, c, ctx.currentSheetIndex, false)) {
            //   continue;
            // }
            var v = replaceText;
            setCellValue(ctx, r, c, d, v);
            range.push({ row: [r, r], column: [c, c] });
            replaceCount += 1;
        }
    }
    else {
        var reg = void 0;
        if (checkModes.caseCheck) {
            reg = new RegExp(getRegExpStr(searchText), "g");
        }
        else {
            reg = new RegExp(getRegExpStr(searchText), "ig");
        }
        for (var i = 0; i < searchIndexArr.length; i += 1) {
            var r = searchIndexArr[i].r;
            var c = searchIndexArr[i].c;
            // if (!checkProtectionLocked(r, c, ctx.currentSheetIndex, false)) {
            //   continue;
            // }
            var v = valueShowEs(r, c, d).toString().replace(reg, replaceText);
            setCellValue(ctx, r, c, d, v);
            range.push({ row: [r, r], column: [c, c] });
            replaceCount += 1;
        }
    }
    // jfrefreshgrid(d, range);
    ctx.luckysheet_select_save = normalizeSelection(ctx, range);
    var succeedInfo = replaceHtml(findAndReplace.successTip, {
        xlength: replaceCount,
    });
    // if (isEditMode()) {
    //   alert(succeedInfo);
    // } else {
    //   tooltip.info(succeedInfo, "");
    // }
    return succeedInfo;
}
