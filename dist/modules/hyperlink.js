import _ from "lodash";
import { getFlowdata } from "../context";
import { getSheetIndex, isAllowEdit } from "../utils";
import { mergeBorder } from "./cell";
import { getcellrange, iscelldata } from "./formula";
import { colLocation, rowLocation } from "./location";
import { normalizeSelection } from "./selection";
import { changeSheet } from "./sheet";
import { locale } from "../locale";
export function getCellRowColumn(ctx, e, container, scrollX, scrollY) {
    var _a, _b;
    var flowdata = getFlowdata(ctx);
    if (flowdata == null)
        return undefined;
    var scrollLeft = scrollX.scrollLeft;
    var scrollTop = scrollY.scrollTop;
    var rect = container.getBoundingClientRect();
    var x = e.pageX - rect.left - ctx.rowHeaderWidth;
    var y = e.pageY - rect.top - ctx.columnHeaderHeight;
    x += scrollLeft;
    y += scrollTop;
    var r = rowLocation(y, ctx.visibledatarow)[2];
    var c = colLocation(x, ctx.visibledatacolumn)[2];
    var margeset = mergeBorder(ctx, flowdata, r, c);
    if (margeset) {
        _a = margeset.row, r = _a[2];
        _b = margeset.column, c = _b[2];
    }
    return { r: r, c: c };
}
export function getCellHyperlink(ctx, r, c) {
    var _a;
    var sheetIndex = getSheetIndex(ctx, ctx.currentSheetId);
    if (sheetIndex != null) {
        return (_a = ctx.luckysheetfile[sheetIndex].hyperlink) === null || _a === void 0 ? void 0 : _a["".concat(r, "_").concat(c)];
    }
    return undefined;
}
export function saveHyperlink(ctx, r, c, linkText, linkType, linkAddress) {
    var sheetIndex = getSheetIndex(ctx, ctx.currentSheetId);
    var flowdata = getFlowdata(ctx);
    if (sheetIndex != null && flowdata != null && linkType && linkAddress) {
        var cell = flowdata[r][c];
        if (cell == null)
            cell = {};
        _.set(ctx.luckysheetfile[sheetIndex], ["hyperlink", "".concat(r, "_").concat(c)], {
            linkType: linkType,
            linkAddress: linkAddress,
        });
        cell.fc = "rgb(0, 0, 255)";
        cell.un = 1;
        cell.v = linkText || linkAddress;
        cell.m = linkText || linkAddress;
        cell.hl = { r: r, c: c, id: ctx.currentSheetId };
        flowdata[r][c] = cell;
        ctx.linkCard = undefined;
    }
}
export function removeHyperlink(ctx, r, c) {
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit)
        return;
    var sheetIndex = getSheetIndex(ctx, ctx.currentSheetId);
    var flowdata = getFlowdata(ctx);
    if (flowdata != null && sheetIndex != null) {
        var hyperlink = _.omit(ctx.luckysheetfile[sheetIndex].hyperlink, "".concat(r, "_").concat(c));
        _.set(ctx.luckysheetfile[sheetIndex], "hyperlink", hyperlink);
        var cell = flowdata[r][c];
        if (cell != null) {
            flowdata[r][c] = { v: cell.v, m: cell.m };
        }
    }
    ctx.linkCard = undefined;
}
export function showLinkCard(ctx, r, c, isEditing, isMouseDown) {
    var _a, _b, _c, _d, _e, _f, _g;
    if (isEditing === void 0) { isEditing = false; }
    if (isMouseDown === void 0) { isMouseDown = false; }
    if ((_a = ctx.linkCard) === null || _a === void 0 ? void 0 : _a.selectingCellRange)
        return;
    if ("".concat(r, "_").concat(c) === ((_b = ctx.linkCard) === null || _b === void 0 ? void 0 : _b.rc))
        return;
    var link = getCellHyperlink(ctx, r, c);
    var cell = (_d = (_c = getFlowdata(ctx)) === null || _c === void 0 ? void 0 : _c[r]) === null || _d === void 0 ? void 0 : _d[c];
    if (!isEditing &&
        link == null &&
        (isMouseDown ||
            !((_e = ctx.linkCard) === null || _e === void 0 ? void 0 : _e.isEditing) ||
            ctx.linkCard.sheetId !== ctx.currentSheetId)) {
        ctx.linkCard = undefined;
        return;
    }
    if (isEditing ||
        (link != null && (!((_f = ctx.linkCard) === null || _f === void 0 ? void 0 : _f.isEditing) || isMouseDown)) ||
        ((_g = ctx.linkCard) === null || _g === void 0 ? void 0 : _g.sheetId) !== ctx.currentSheetId) {
        var col_pre = c - 1 === -1 ? 0 : ctx.visibledatacolumn[c - 1];
        var row = ctx.visibledatarow[r];
        ctx.linkCard = {
            sheetId: ctx.currentSheetId,
            r: r,
            c: c,
            rc: "".concat(r, "_").concat(c),
            originText: (cell === null || cell === void 0 ? void 0 : cell.v) == null ? "" : "".concat(cell.v),
            originType: (link === null || link === void 0 ? void 0 : link.linkType) || "webpage",
            originAddress: (link === null || link === void 0 ? void 0 : link.linkAddress) || "",
            position: {
                cellLeft: col_pre,
                cellBottom: row,
            },
            isEditing: isEditing,
        };
    }
}
export function goToLink(ctx, r, c, linkType, linkAddress, scrollbarX, scrollbarY) {
    var _a;
    var currSheetIndex = getSheetIndex(ctx, ctx.currentSheetId);
    if (currSheetIndex == null)
        return;
    if (((_a = ctx.luckysheetfile[currSheetIndex].hyperlink) === null || _a === void 0 ? void 0 : _a["".concat(r, "_").concat(c)]) == null) {
        return;
    }
    if (linkType === "webpage") {
        if (!/^http[s]?:\/\//.test(linkAddress)) {
            linkAddress = "https://".concat(linkAddress);
        }
        window.open(linkAddress);
    }
    else if (linkType === "sheet") {
        var sheetId_1;
        _.forEach(ctx.luckysheetfile, function (f) {
            if (linkAddress === f.name) {
                sheetId_1 = f.id;
            }
        });
        if (sheetId_1 != null)
            changeSheet(ctx, sheetId_1);
    }
    else {
        var range = _.cloneDeep(getcellrange(ctx, linkAddress));
        if (range == null)
            return;
        var row_pre = range.row[0] - 1 === -1 ? 0 : ctx.visibledatarow[range.row[0] - 1];
        var col_pre = range.column[0] - 1 === -1
            ? 0
            : ctx.visibledatacolumn[range.column[0] - 1];
        scrollbarX.scrollLeft = col_pre;
        scrollbarY.scrollLeft = row_pre;
        ctx.luckysheet_select_save = normalizeSelection(ctx, [range]);
        changeSheet(ctx, range.sheetId);
    }
    ctx.linkCard = undefined;
}
export function isLinkValid(ctx, linkType, linkAddress) {
    if (!linkAddress)
        return { isValid: false, tooltip: "" };
    var insertLink = locale(ctx).insertLink;
    if (linkType === "webpage") {
        if (!/^http[s]?:\/\//.test(linkAddress)) {
            linkAddress = "https://".concat(linkAddress);
        }
        if (
        // eslint-disable-next-line no-useless-escape
        !/^http[s]?:\/\/([\w\-\.]+)+[\w-]*([\w\-\.\/\?%&=]+)?$/gi.test(linkAddress))
            return { isValid: false, tooltip: insertLink.tooltipInfo1 };
    }
    if (linkType === "cellrange" && !iscelldata(linkAddress)) {
        return { isValid: false, tooltip: insertLink.invalidCellRangeTip };
    }
    return { isValid: true, tooltip: "" };
}
export function onRangeSelectionModalMoveStart(ctx, globalCache, e) {
    var box = document.querySelector("div.fortune-link-modify-modal.range-selection-modal");
    if (!box)
        return;
    var _a = box.getBoundingClientRect(), width = _a.width, height = _a.height;
    var left = box.offsetLeft;
    var top = box.offsetTop;
    var initialPosition = { left: left, top: top, width: width, height: height };
    _.set(globalCache, "linkCard.rangeSelectionModal", {
        cursorMoveStartPosition: {
            x: e.pageX,
            y: e.pageY,
        },
        initialPosition: initialPosition,
    });
}
export function onRangeSelectionModalMove(globalCache, e) {
    var _a;
    var moveProps = (_a = globalCache.linkCard) === null || _a === void 0 ? void 0 : _a.rangeSelectionModal;
    if (moveProps == null)
        return;
    var modal = document.querySelector("div.fortune-link-modify-modal.range-selection-modal");
    var _b = moveProps.cursorMoveStartPosition, startX = _b.x, startY = _b.y;
    var _c = moveProps.initialPosition, top = _c.top, left = _c.left;
    left += e.pageX - startX;
    top += e.pageY - startY;
    if (top < 0)
        top = 0;
    modal.style.left = "".concat(left, "px");
    modal.style.top = "".concat(top, "px");
}
export function onRangeSelectionModalMoveEnd(globalCache) {
    _.set(globalCache, "linkCard.rangeSelectionModal", undefined);
}
