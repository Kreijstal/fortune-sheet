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
import _ from "lodash";
import { mergeBorder } from "./cell";
import { getFlowdata } from "../context";
import { colLocation, rowLocation } from "./location";
import { isAllowEdit } from "../utils";
export function getArrowCanvasSize(fromX, fromY, toX, toY) {
    var left = toX - 5;
    if (fromX < toX) {
        left = fromX - 5;
    }
    var top = toY - 5;
    if (fromY < toY) {
        top = fromY - 5;
    }
    var width = Math.abs(fromX - toX) + 10;
    var height = Math.abs(fromY - toY) + 10;
    var x1 = width - 5;
    var x2 = 5;
    if (fromX < toX) {
        x1 = 5;
        x2 = width - 5;
    }
    var y1 = height - 5;
    var y2 = 5;
    if (fromY < toY) {
        y1 = 5;
        y2 = height - 5;
    }
    return { left: left, top: top, width: width, height: height, fromX: x1, fromY: y1, toX: x2, toY: y2 };
}
export function drawArrow(rc, _a, color, theta, headlen) {
    var left = _a.left, top = _a.top, width = _a.width, height = _a.height, fromX = _a.fromX, fromY = _a.fromY, toX = _a.toX, toY = _a.toY;
    var canvas = document.getElementById("arrowCanvas-".concat(rc));
    var ctx = canvas.getContext("2d");
    if (!canvas || !ctx)
        return;
    canvas.style.width = "".concat(width, "px");
    canvas.style.height = "".concat(height, "px");
    canvas.width = width;
    canvas.height = height;
    canvas.style.left = "".concat(left, "px");
    canvas.style.top = "".concat(top, "px");
    var _b = canvas.getBoundingClientRect(), canvasWidth = _b.width, canvasHeight = _b.height;
    ctx.clearRect(0, 0, canvasWidth, canvasHeight);
    theta = theta || 30;
    headlen = headlen || 6;
    // width = width || 1;
    var arrowWidth = 1;
    color = color || "#000";
    // 计算各角度和对应的P2,P3坐标
    var angle = (Math.atan2(fromY - toY, fromX - toX) * 180) / Math.PI;
    var angle1 = ((angle + theta) * Math.PI) / 180;
    var angle2 = ((angle - theta) * Math.PI) / 180;
    var topX = headlen * Math.cos(angle1);
    var topY = headlen * Math.sin(angle1);
    var botX = headlen * Math.cos(angle2);
    var botY = headlen * Math.sin(angle2);
    ctx.save();
    ctx.beginPath();
    var arrowX = fromX - topX;
    var arrowY = fromY - topY;
    ctx.moveTo(arrowX, arrowY);
    ctx.moveTo(fromX, fromY);
    ctx.lineTo(toX, toY);
    ctx.lineWidth = arrowWidth;
    ctx.strokeStyle = color;
    ctx.stroke();
    arrowX = toX + topX;
    arrowY = toY + topY;
    ctx.moveTo(arrowX, arrowY);
    ctx.lineTo(toX, toY);
    arrowX = toX + botX;
    arrowY = toY + botY;
    ctx.lineTo(arrowX, arrowY);
    ctx.fillStyle = color;
    ctx.fill();
    ctx.restore();
}
export var commentBoxProps = {
    defaultWidth: 144,
    defaultHeight: 84,
    currentObj: null,
    currentWinW: null,
    currentWinH: null,
    resize: null,
    resizeXY: null,
    move: false,
    moveXY: null,
    cursorStartPosition: null,
};
export function getCellTopRightPostion(ctx, flowdata, r, c) {
    var _a;
    // let row = ctx.visibledatarow[r];
    var row_pre = r - 1 === -1 ? 0 : ctx.visibledatarow[r - 1];
    var col = ctx.visibledatacolumn[c];
    //  let col_pre = c - 1 === -1 ? 0 : ctx.visibledatacolumn[c - 1];
    var margeset = mergeBorder(ctx, flowdata, r, c);
    if (margeset) {
        // row = margeset.row[1];
        row_pre = margeset.row[0];
        // col_pre = margeset.column[0];
        _a = margeset.column, col = _a[1];
    }
    var toX = col;
    var toY = row_pre;
    return { toX: toX, toY: toY };
}
export function getCommentBoxByRC(ctx, flowdata, r, c) {
    var _a;
    var comment = (_a = flowdata[r][c]) === null || _a === void 0 ? void 0 : _a.ps;
    var _b = getCellTopRightPostion(ctx, flowdata, r, c), toX = _b.toX, toY = _b.toY;
    // let scrollLeft = $("#luckysheet-cell-main").scrollLeft();
    // let scrollTop = $("#luckysheet-cell-main").scrollTop();
    // if(luckysheetFreezen.freezenverticaldata != null && toX < (luckysheetFreezen.freezenverticaldata[0] - luckysheetFreezen.freezenverticaldata[2])){
    //     toX += scrollLeft;
    // }
    // if(luckysheetFreezen.freezenhorizontaldata != null && toY < (luckysheetFreezen.freezenhorizontaldata[0] - luckysheetFreezen.freezenhorizontaldata[2])){
    //     toY += scrollTop;
    // }
    var left = (comment === null || comment === void 0 ? void 0 : comment.left) == null
        ? toX + 18 * ctx.zoomRatio
        : comment.left * ctx.zoomRatio;
    var top = (comment === null || comment === void 0 ? void 0 : comment.top) == null
        ? toY - 18 * ctx.zoomRatio
        : comment.top * ctx.zoomRatio;
    var width = (comment === null || comment === void 0 ? void 0 : comment.width) == null
        ? commentBoxProps.defaultWidth * ctx.zoomRatio
        : comment.width * ctx.zoomRatio;
    var height = (comment === null || comment === void 0 ? void 0 : comment.height) == null
        ? commentBoxProps.defaultHeight * ctx.zoomRatio
        : comment.height * ctx.zoomRatio;
    var value = (comment === null || comment === void 0 ? void 0 : comment.value) == null ? "" : comment.value;
    if (top < 0) {
        top = 2;
    }
    var size = getArrowCanvasSize(left, top, toX, toY);
    var rc = "".concat(r, "_").concat(c);
    return { r: r, c: c, rc: rc, left: left, top: top, width: width, height: height, value: value, size: size, autoFocus: false };
}
export function setEditingComment(ctx, flowdata, r, c) {
    ctx.editingCommentBox = getCommentBoxByRC(ctx, flowdata, r, c);
}
export function removeEditingComment(ctx, globalCache) {
    var _a, _b;
    var editingCommentBoxEle = globalCache.editingCommentBoxEle;
    ctx.editingCommentBox = undefined;
    var r = editingCommentBoxEle === null || editingCommentBoxEle === void 0 ? void 0 : editingCommentBoxEle.dataset.r;
    var c = editingCommentBoxEle === null || editingCommentBoxEle === void 0 ? void 0 : editingCommentBoxEle.dataset.c;
    if (!r || !c || !editingCommentBoxEle)
        return;
    r = parseInt(r, 10);
    c = parseInt(c, 10);
    var value = editingCommentBoxEle.innerHTML || "";
    var flowdata = getFlowdata(ctx);
    globalCache.editingCommentBoxEle = undefined;
    if (!flowdata)
        return;
    if (((_b = (_a = ctx.hooks).beforeUpdateComment) === null || _b === void 0 ? void 0 : _b.call(_a, r, c, value)) === false) {
        return;
    }
    //  const prevCell = _.cloneDeep(flowdata?.[r][c]) || {};
    var cell = flowdata === null || flowdata === void 0 ? void 0 : flowdata[r][c];
    if (!(cell === null || cell === void 0 ? void 0 : cell.ps))
        return;
    var oldValue = cell.ps.value;
    cell.ps.value = value;
    if (!cell.ps.isShow) {
        ctx.commentBoxes = _.filter(ctx.commentBoxes, function (v) { return v.rc !== "".concat(r, "_").concat(c); });
    }
    if (ctx.hooks.afterUpdateComment) {
        setTimeout(function () {
            var _a, _b;
            (_b = (_a = ctx.hooks).afterUpdateComment) === null || _b === void 0 ? void 0 : _b.call(_a, r, c, oldValue, value);
        });
    }
}
export function newComment(ctx, globalCache, r, c) {
    var _a, _b;
    // if(!checkProtectionAuthorityNormal(Store.currentSheetId, "editObjects")){
    //     return;
    // }
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit)
        return;
    if (((_b = (_a = ctx.hooks).beforeInsertComment) === null || _b === void 0 ? void 0 : _b.call(_a, r, c)) === false) {
        return;
    }
    removeEditingComment(ctx, globalCache);
    var flowdata = getFlowdata(ctx);
    if (!flowdata)
        return;
    var cell = flowdata[r][c];
    if (cell == null) {
        cell = {};
        flowdata[r][c] = cell;
    }
    cell.ps = {
        left: null,
        top: null,
        width: null,
        height: null,
        value: "",
        isShow: false,
    };
    ctx.editingCommentBox = __assign(__assign({}, getCommentBoxByRC(ctx, flowdata, r, c)), { autoFocus: true });
    if (ctx.hooks.afterInsertComment) {
        setTimeout(function () {
            var _a, _b;
            (_b = (_a = ctx.hooks).afterInsertComment) === null || _b === void 0 ? void 0 : _b.call(_a, r, c);
        });
    }
}
export function editComment(ctx, globalCache, r, c) {
    var _a;
    // if(!checkProtectionAuthorityNormal(Store.currentSheetId, "editObjects")){
    //     return;
    // }
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit)
        return;
    var flowdata = getFlowdata(ctx);
    removeEditingComment(ctx, globalCache);
    var comment = (_a = flowdata === null || flowdata === void 0 ? void 0 : flowdata[r][c]) === null || _a === void 0 ? void 0 : _a.ps;
    var commentBoxes = _.concat(ctx.commentBoxes, ctx.editingCommentBox);
    if (_.findIndex(commentBoxes, function (v) { return (v === null || v === void 0 ? void 0 : v.rc) === "".concat(r, "_").concat(c); }) !== -1) {
        var editCommentBox = document.getElementById("comment-editor-".concat(r, "_").concat(c));
        editCommentBox === null || editCommentBox === void 0 ? void 0 : editCommentBox.focus();
    }
    if (comment) {
        ctx.editingCommentBox = __assign(__assign({}, getCommentBoxByRC(ctx, flowdata, r, c)), { autoFocus: true });
    }
}
export function deleteComment(ctx, globalCache, r, c) {
    var _a, _b;
    // if(!checkProtectionAuthorityNormal(Store.currentSheetId, "editObjects")){
    //     return;
    // }
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit)
        return;
    if (((_b = (_a = ctx.hooks).beforeDeleteComment) === null || _b === void 0 ? void 0 : _b.call(_a, r, c)) === false) {
        return;
    }
    var flowdata = getFlowdata(ctx);
    if (!flowdata)
        return;
    var cell = flowdata[r][c];
    if (!cell)
        return;
    cell.ps = undefined;
    if (ctx.hooks.afterDeleteComment) {
        setTimeout(function () {
            var _a, _b;
            (_b = (_a = ctx.hooks).afterDeleteComment) === null || _b === void 0 ? void 0 : _b.call(_a, r, c);
        });
    }
}
export function showComments(ctx, commentShowCells) {
    var flowdata = getFlowdata(ctx);
    if (flowdata) {
        var commentBoxes = commentShowCells.map(function (_a) {
            var r = _a.r, c = _a.c;
            return getCommentBoxByRC(ctx, flowdata, r, c);
        });
        ctx.commentBoxes = commentBoxes;
    }
}
export function showHideComment(ctx, globalCache, r, c) {
    var _a;
    var flowdata = getFlowdata(ctx);
    var comment = (_a = flowdata === null || flowdata === void 0 ? void 0 : flowdata[r][c]) === null || _a === void 0 ? void 0 : _a.ps;
    if (!comment)
        return;
    var isShow = comment.isShow;
    var rc = "".concat(r, "_").concat(c);
    if (isShow) {
        comment.isShow = false;
        ctx.commentBoxes = _.filter(ctx.commentBoxes, function (v) { return v.rc !== rc; });
    }
    else {
        comment.isShow = true;
    }
}
export function showHideAllComments(ctx) {
    var _a, _b;
    var flowdata = getFlowdata(ctx);
    if (!flowdata)
        return;
    var isAllShow = true;
    var allComments = [];
    for (var r = 0; r < flowdata.length; r += 1) {
        for (var c = 0; c < flowdata[0].length; c += 1) {
            var cell = flowdata[r][c];
            if (cell === null || cell === void 0 ? void 0 : cell.ps) {
                allComments.push({ r: r, c: c });
                if (!cell.ps.isShow) {
                    isAllShow = false;
                }
            }
        }
    }
    var rcs = [];
    if (allComments.length > 0) {
        if (isAllShow) {
            // 全部显示，操作为隐藏所有批注
            for (var i = 0; i < allComments.length; i += 1) {
                var _c = allComments[i], r = _c.r, c = _c.c;
                var comment = (_a = flowdata[r][c]) === null || _a === void 0 ? void 0 : _a.ps;
                if (comment === null || comment === void 0 ? void 0 : comment.isShow) {
                    comment.isShow = false;
                    rcs.push("".concat(r, "_").concat(c));
                }
            }
            ctx.commentBoxes = [];
        }
        else {
            // 部分显示或全部隐藏，操作位显示所有批注
            for (var i = 0; i < allComments.length; i += 1) {
                var _d = allComments[i], r = _d.r, c = _d.c;
                var comment = (_b = flowdata[r][c]) === null || _b === void 0 ? void 0 : _b.ps;
                if (comment && !comment.isShow) {
                    comment.isShow = true;
                }
            }
        }
    }
}
// show comment when mouse is over cell with comment
export function overShowComment(ctx, e, scrollX, scrollY, container) {
    var _a, _b, _c;
    var _d, _e, _f, _g;
    var flowdata = getFlowdata(ctx);
    if (!flowdata)
        return;
    var scrollLeft = scrollX.scrollLeft;
    var scrollTop = scrollY.scrollTop;
    // $("#luckysheet-postil-overshow").remove();
    // if($(event.target).closest("#luckysheet-cell-main").length == 0){
    //     return;
    // }
    var rect = container.getBoundingClientRect();
    // const mouse = mousePosition(e.pageX, e.pageY, ctx);
    var x = e.pageX - rect.left - ctx.rowHeaderWidth;
    var y = e.pageY - rect.top - ctx.columnHeaderHeight;
    var offsetX = 0;
    var offsetY = 0;
    //   if (
    //     luckysheetFreezen.freezenverticaldata != null &&
    //     mouse[0] <
    //       luckysheetFreezen.freezenverticaldata[0] -
    //         luckysheetFreezen.freezenverticaldata[2]
    //   ) {
    //     offsetX = scrollLeft;
    //   } else {
    x += scrollLeft;
    //   }
    //   if (
    //     luckysheetFreezen.freezenhorizontaldata != null &&
    //     mouse[1] <
    //       luckysheetFreezen.freezenhorizontaldata[0] -
    //         luckysheetFreezen.freezenhorizontaldata[2]
    //   ) {
    //     offsetY = scrollTop;
    //   } else {
    y += scrollTop;
    //   }
    var r = rowLocation(y, ctx.visibledatarow)[2];
    var c = colLocation(x, ctx.visibledatacolumn)[2];
    var margeset = mergeBorder(ctx, flowdata, r, c);
    if (margeset) {
        _a = margeset.row, r = _a[2];
        _b = margeset.column, c = _b[2];
    }
    var rc = "".concat(r, "_").concat(c);
    var comment = (_e = (_d = flowdata[r]) === null || _d === void 0 ? void 0 : _d[c]) === null || _e === void 0 ? void 0 : _e.ps;
    if (comment == null ||
        comment.isShow ||
        _.findIndex(ctx.commentBoxes, function (v) { return v.rc === rc; }) !== -1 ||
        ((_f = ctx.editingCommentBox) === null || _f === void 0 ? void 0 : _f.rc) === rc) {
        ctx.hoveredCommentBox = undefined;
        return;
    }
    if (((_g = ctx.hoveredCommentBox) === null || _g === void 0 ? void 0 : _g.rc) === rc)
        return;
    // let row = ctx.visibledatarow[row_index];
    var row_pre = r - 1 === -1 ? 0 : ctx.visibledatarow[r - 1];
    var col = ctx.visibledatacolumn[c];
    // let col_pre = col_index - 1 === -1 ? 0 : ctx.visibledatacolumn[col_index - 1];
    if (margeset) {
        //  [, row] = margeset.row;
        row_pre = margeset.row[0];
        _c = margeset.column, col = _c[1];
        //  [col_pre] = margeset.column;
    }
    var toX = col + offsetX;
    var toY = row_pre + offsetY;
    var left = comment.left == null
        ? toX + 18 * ctx.zoomRatio
        : comment.left * ctx.zoomRatio;
    var top = comment.top == null
        ? toY - 18 * ctx.zoomRatio
        : comment.top * ctx.zoomRatio;
    if (top < 0) {
        top = 2;
    }
    var width = comment.width == null
        ? commentBoxProps.defaultWidth * ctx.zoomRatio
        : comment.width * ctx.zoomRatio;
    var height = comment.height == null
        ? commentBoxProps.defaultHeight * ctx.zoomRatio
        : comment.height * ctx.zoomRatio;
    var size = getArrowCanvasSize(left, top, toX, toY);
    var value = comment.value == null ? "" : comment.value;
    ctx.hoveredCommentBox = {
        r: r,
        c: c,
        rc: rc,
        left: left,
        top: top,
        width: width,
        height: height,
        size: size,
        value: value,
        autoFocus: false,
    };
}
export function getCommentBoxPosition(commentId) {
    var box = document.getElementById(commentId);
    if (!box)
        return undefined;
    var _a = box.getBoundingClientRect(), width = _a.width, height = _a.height;
    var left = box.offsetLeft;
    var top = box.offsetTop;
    return { left: left, top: top, width: width, height: height };
}
export function onCommentBoxResizeStart(ctx, globalCache, e, _a, resizingId, resizingSide) {
    var r = _a.r, c = _a.c, rc = _a.rc;
    var position = getCommentBoxPosition(resizingId);
    if (position) {
        _.set(globalCache, "commentBox", {
            cursorMoveStartPosition: {
                x: e.pageX,
                y: e.pageY,
            },
            resizingId: resizingId,
            resizingSide: resizingSide,
            commentRC: { r: r, c: c, rc: rc },
            boxInitialPosition: position,
        });
    }
}
export function onCommentBoxResize(ctx, globalCache, e) {
    if (ctx.allowEdit === false)
        return false;
    var commentBox = globalCache === null || globalCache === void 0 ? void 0 : globalCache.commentBox;
    if ((commentBox === null || commentBox === void 0 ? void 0 : commentBox.resizingId) && commentBox.resizingSide) {
        var box = document.getElementById(commentBox.resizingId);
        var _a = commentBox.cursorMoveStartPosition, startX = _a.x, startY = _a.y;
        var _b = commentBox.boxInitialPosition, top_1 = _b.top, left = _b.left, width = _b.width, height = _b.height;
        var dx = e.pageX - startX;
        var dy = e.pageY - startY;
        var minHeight = 60 * ctx.zoomRatio;
        var minWidth = 1.5 * 60 * ctx.zoomRatio;
        if (["lm", "lt", "lb"].includes(commentBox.resizingSide)) {
            if (width - dx < minWidth) {
                left += width - minWidth;
                width = minWidth;
            }
            else {
                left += dx;
                width -= dx;
            }
            if (left < 0)
                left = 0;
            box.style.left = "".concat(left, "px");
        }
        if (["rm", "rt", "rb"].includes(commentBox.resizingSide)) {
            width = width + dx < minWidth ? minWidth : width + dx;
        }
        if (["mt", "lt", "rt"].includes(commentBox.resizingSide)) {
            if (height - dy < minHeight) {
                top_1 += height - minHeight;
                height = minHeight;
            }
            else {
                top_1 += dy;
                height -= dy;
            }
            if (top_1 < 0)
                top_1 = 0;
            box.style.top = "".concat(top_1, "px");
        }
        if (["mb", "lb", "rb"].includes(commentBox.resizingSide)) {
            height = height + dy < minHeight ? minHeight : height + dy;
        }
        box.style.width = "".concat(width, "px");
        box.style.height = "".concat(height, "px");
        return true;
    }
    return false;
}
export function onCommentBoxResizeEnd(ctx, globalCache) {
    var _a;
    if ((_a = globalCache.commentBox) === null || _a === void 0 ? void 0 : _a.resizingId) {
        var _b = globalCache.commentBox, resizingId = _b.resizingId, _c = _b.commentRC, r = _c.r, c = _c.c;
        globalCache.commentBox.resizingId = undefined;
        var position = getCommentBoxPosition(resizingId);
        if (position) {
            var top_2 = position.top, left = position.left, width = position.width, height = position.height;
            var flowdata = getFlowdata(ctx);
            var cell = flowdata === null || flowdata === void 0 ? void 0 : flowdata[r][c];
            if (!flowdata || !(cell === null || cell === void 0 ? void 0 : cell.ps))
                return;
            cell.ps.left = left / ctx.zoomRatio;
            cell.ps.top = top_2 / ctx.zoomRatio;
            cell.ps.width = width / ctx.zoomRatio;
            cell.ps.height = height / ctx.zoomRatio;
            setEditingComment(ctx, flowdata, r, c);
        }
    }
}
export function onCommentBoxMoveStart(ctx, globalCache, e, _a, movingId) {
    var r = _a.r, c = _a.c, rc = _a.rc;
    var position = getCommentBoxPosition(movingId);
    if (position) {
        var top_3 = position.top, left = position.left;
        _.set(globalCache, "commentBox", {
            cursorMoveStartPosition: {
                x: e.pageX,
                y: e.pageY,
            },
            movingId: movingId,
            commentRC: { r: r, c: c, rc: rc },
            boxInitialPosition: { left: left, top: top_3 },
        });
    }
}
export function onCommentBoxMove(ctx, globalCache, e) {
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit)
        return false;
    var commentBox = globalCache === null || globalCache === void 0 ? void 0 : globalCache.commentBox;
    if (commentBox === null || commentBox === void 0 ? void 0 : commentBox.movingId) {
        var box = document.getElementById(commentBox.movingId);
        var _a = commentBox.cursorMoveStartPosition, startX = _a.x, startY = _a.y;
        var _b = commentBox.boxInitialPosition, top_4 = _b.top, left = _b.left;
        left += e.pageX - startX;
        top_4 += e.pageY - startY;
        if (top_4 < 0)
            top_4 = 0;
        box.style.left = "".concat(left, "px");
        box.style.top = "".concat(top_4, "px");
        return true;
    }
    return false;
}
export function onCommentBoxMoveEnd(ctx, globalCache) {
    var _a;
    if ((_a = globalCache.commentBox) === null || _a === void 0 ? void 0 : _a.movingId) {
        var _b = globalCache.commentBox, movingId = _b.movingId, _c = _b.commentRC, r = _c.r, c = _c.c;
        globalCache.commentBox.movingId = undefined;
        var position = getCommentBoxPosition(movingId);
        if (position) {
            var top_5 = position.top, left = position.left;
            var flowdata = getFlowdata(ctx);
            var cell = flowdata === null || flowdata === void 0 ? void 0 : flowdata[r][c];
            if (!flowdata || !(cell === null || cell === void 0 ? void 0 : cell.ps))
                return;
            cell.ps.left = left / ctx.zoomRatio;
            cell.ps.top = top_5 / ctx.zoomRatio;
            setEditingComment(ctx, flowdata, r, c);
        }
    }
}
/*
const luckysheetPostil = {
  getArrowCanvasSize(fromX, fromY, toX, toY) {
    let left = toX - 5;

    if (fromX < toX) {
      left = fromX - 5;
    }

    let top = toY - 5;

    if (fromY < toY) {
      top = fromY - 5;
    }

    const width = Math.abs(fromX - toX) + 10;
    const height = Math.abs(fromY - toY) + 10;

    let x1 = width - 5;
    let x2 = 5;

    if (fromX < toX) {
      x1 = 5;
      x2 = width - 5;
    }

    let y1 = height - 5;
    let y2 = 5;

    if (fromY < toY) {
      y1 = 5;
      y2 = height - 5;
    }

    return [left, top, width, height, x1, y1, x2, y2];
  },
  drawArrow(ctx, fromX, fromY, toX, toY, theta, headlen, width, color) {
    theta = getObjType(theta) == "undefined" ? 30 : theta;
    headlen = getObjType(headlen) == "undefined" ? 6 : headlen;
    width = getObjType(width) == "undefined" ? 1 : width;
    color = getObjType(color) == "undefined" ? "#000" : color;

    // 计算各角度和对应的P2,P3坐标
    const angle = (Math.atan2(fromY - toY, fromX - toX) * 180) / Math.PI;
    const angle1 = ((angle + theta) * Math.PI) / 180;
    const angle2 = ((angle - theta) * Math.PI) / 180;
    const topX = headlen * Math.cos(angle1);
    const topY = headlen * Math.sin(angle1);
    const botX = headlen * Math.cos(angle2);
    const botY = headlen * Math.sin(angle2);

    ctx.save();
    ctx.beginPath();

    let arrowX = fromX - topX;
    let arrowY = fromY - topY;

    ctx.moveTo(arrowX, arrowY);
    ctx.moveTo(fromX, fromY);
    ctx.lineTo(toX, toY);

    ctx.lineWidth = width;
    ctx.strokeStyle = color;
    ctx.stroke();

    arrowX = toX + topX;
    arrowY = toY + topY;
    ctx.moveTo(arrowX, arrowY);
    ctx.lineTo(toX, toY);
    arrowX = toX + botX;
    arrowY = toY + botY;
    ctx.lineTo(arrowX, arrowY);

    ctx.fillStyle = color;
    ctx.fill();
    ctx.restore();
  },
  buildAllPs(data) {
    const _this = this;

    $("#luckysheet-cell-main #luckysheet-postil-showBoxs").empty();

    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[0].length; c++) {
        if (data[r][c] != null && data[r][c].ps != null) {
          const postil = data[r][c].ps;
          _this.buildPs(r, c, postil);
        }
      }
    }

    _this.init();
  },
  buildPs(r, c, postil) {
    if ($(`#luckysheet-postil-show_${r}_${c}`).length > 0) {
      $(`#luckysheet-postil-show_${r}_${c}`).remove();
    }

    if (postil == null) {
      return;
    }

    const _this = this;
    const isShow = postil.isShow == null ? false : postil.isShow;

    if (isShow) {
      let row = Store.visibledatarow[r];
      let row_pre = r - 1 == -1 ? 0 : Store.visibledatarow[r - 1];
      let col = Store.visibledatacolumn[c];
      let col_pre = c - 1 == -1 ? 0 : Store.visibledatacolumn[c - 1];

      const margeset = menuButton.mergeborer(Store.flowdata, r, c);
      if (margeset) {
        row = margeset.row[1];
        row_pre = margeset.row[0];

        col = margeset.column[1];
        col_pre = margeset.column[0];
      }

      const toX = col;
      const toY = row_pre;

      const left =
        postil.left == null
          ? toX + 18 * Store.zoomRatio
          : postil.left * Store.zoomRatio;
      let top =
        postil.top == null
          ? toY - 18 * Store.zoomRatio
          : postil.top * Store.zoomRatio;
      const width =
        postil.width == null
          ? _this.defaultWidth * Store.zoomRatio
          : postil.width * Store.zoomRatio;
      const height =
        postil.height == null
          ? _this.defaultHeight * Store.zoomRatio
          : postil.height * Store.zoomRatio;
      const value = postil.value == null ? "" : postil.value;

      if (top < 0) {
        top = 2;
      }

      const size = _this.getArrowCanvasSize(left, top, toX, toY);

      let commentDivs = "";
      const valueLines = value.split("\n");
      for (const line of valueLines) {
        commentDivs += `<div>${_this.htmlEscape(line)}</div>`;
      }

      const html =
        `<div id="luckysheet-postil-show_${r}_${c}" class="luckysheet-postil-show">` +
        `<canvas class="arrowCanvas" width="${size[2]}" height="${size[3]}" style="position:absolute;left:${size[0]}px;top:${size[1]}px;z-index:100;pointer-events:none;"></canvas>` +
        `<div class="luckysheet-postil-show-main" style="width:${width}px;height:${height}px;color:#000;padding:5px;border:1px solid #000;background-color:rgb(255,255,225);position:absolute;left:${left}px;top:${top}px;box-sizing:border-box;z-index:100;">` +
        `<div class="luckysheet-postil-dialog-move">` +
        `<div class="luckysheet-postil-dialog-move-item luckysheet-postil-dialog-move-item-t" data-type="t"></div>` +
        `<div class="luckysheet-postil-dialog-move-item luckysheet-postil-dialog-move-item-r" data-type="r"></div>` +
        `<div class="luckysheet-postil-dialog-move-item luckysheet-postil-dialog-move-item-b" data-type="b"></div>` +
        `<div class="luckysheet-postil-dialog-move-item luckysheet-postil-dialog-move-item-l" data-type="l"></div>` +
        `</div>` +
        `<div class="luckysheet-postil-dialog-resize" style="display:none;">` +
        `<div class="luckysheet-postil-dialog-resize-item luckysheet-postil-dialog-resize-item-lt" data-type="lt"></div>` +
        `<div class="luckysheet-postil-dialog-resize-item luckysheet-postil-dialog-resize-item-mt" data-type="mt"></div>` +
        `<div class="luckysheet-postil-dialog-resize-item luckysheet-postil-dialog-resize-item-lm" data-type="lm"></div>` +
        `<div class="luckysheet-postil-dialog-resize-item luckysheet-postil-dialog-resize-item-rm" data-type="rm"></div>` +
        `<div class="luckysheet-postil-dialog-resize-item luckysheet-postil-dialog-resize-item-rt" data-type="rt"></div>` +
        `<div class="luckysheet-postil-dialog-resize-item luckysheet-postil-dialog-resize-item-lb" data-type="lb"></div>` +
        `<div class="luckysheet-postil-dialog-resize-item luckysheet-postil-dialog-resize-item-mb" data-type="mb"></div>` +
        `<div class="luckysheet-postil-dialog-resize-item luckysheet-postil-dialog-resize-item-rb" data-type="rb"></div>` +
        `</div>` +
        `<div style="width:100%;height:100%;overflow:hidden;">` +
        `<div class="formulaInputFocus" style="width:${width - 12}px;height:${
          height - 12
        }px;line-height:20px;box-sizing:border-box;text-align: center;;word-break:break-all;" spellcheck="false" contenteditable="true">${commentDivs}</div>` +
        `</div>` +
        `</div>` +
        `</div>`;

      $(html).appendTo($("#luckysheet-cell-main #luckysheet-postil-showBoxs"));

      const ctx = $(`#luckysheet-postil-show_${r}_${c} .arrowCanvas`)
        .get(0)
        .getContext("2d");

      _this.drawArrow(ctx, size[4], size[5], size[6], size[7]);
    }
  },
};

export default luckysheetPostil;
*/
