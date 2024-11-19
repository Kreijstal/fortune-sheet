import _ from "lodash";
import { mergeBorder } from ".";
import { getFlowdata } from "../context";
import { getSheetIndex } from "../utils";
export var imageProps = {
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
export function generateRandomId(prefix) {
    if (prefix == null) {
        prefix = "img";
    }
    var userAgent = window.navigator.userAgent
        .replace(/[^a-zA-Z0-9]/g, "")
        .split("");
    var mid = "";
    for (var i = 0; i < 12; i += 1) {
        mid += userAgent[Math.round(Math.random() * (userAgent.length - 1))];
    }
    var time = new Date().getTime();
    return "".concat(prefix, "_").concat(mid, "_").concat(time);
}
export function showImgChooser() {
    var chooser = document.getElementById("fortune-img-upload");
    if (chooser)
        chooser.click();
}
export function saveImage(ctx) {
    var index = getSheetIndex(ctx, ctx.currentSheetId);
    if (index == null)
        return;
    var file = ctx.luckysheetfile[index];
    file.images = ctx.insertedImgs;
}
export function removeActiveImage(ctx) {
    ctx.insertedImgs = _.filter(ctx.insertedImgs, function (image) { return image.id !== ctx.activeImg; });
    ctx.activeImg = undefined;
    saveImage(ctx);
}
export function insertImage(ctx, image) {
    var _a;
    try {
        var last = (_a = ctx.luckysheet_select_save) === null || _a === void 0 ? void 0 : _a[ctx.luckysheet_select_save.length - 1];
        var rowIndex = last === null || last === void 0 ? void 0 : last.row_focus;
        var colIndex = last === null || last === void 0 ? void 0 : last.column_focus;
        if (!last) {
            rowIndex = 0;
            colIndex = 0;
        }
        else {
            if (rowIndex == null) {
                rowIndex = last.row[0];
            }
            if (colIndex == null) {
                colIndex = last.column[0];
            }
        }
        var flowdata = getFlowdata(ctx);
        var left = colIndex === 0 ? 0 : ctx.visibledatacolumn[colIndex - 1];
        var top_1 = rowIndex === 0 ? 0 : ctx.visibledatarow[rowIndex - 1];
        if (flowdata) {
            var margeset = mergeBorder(ctx, flowdata, rowIndex, colIndex);
            if (margeset) {
                top_1 = margeset.row[0];
                left = margeset.column[0];
            }
        }
        var width = image.width;
        var height = image.height;
        var img = {
            id: generateRandomId("img"),
            src: image.src,
            left: left,
            top: top_1,
            width: width * 0.5,
            height: height * 0.5,
            originWidth: width,
            originHeight: height,
        };
        ctx.insertedImgs = (ctx.insertedImgs || []).concat(img);
        saveImage(ctx);
    }
    catch (err) {
        // eslint-disable-next-line no-console
        console.info(err);
    }
}
function getImagePosition() {
    var box = document.getElementById("luckysheet-modal-dialog-activeImage");
    if (!box)
        return undefined;
    var _a = box.getBoundingClientRect(), width = _a.width, height = _a.height;
    var left = box.offsetLeft;
    var top = box.offsetTop;
    return { left: left, top: top, width: width, height: height };
}
export function cancelActiveImgItem(ctx, globalCache) {
    ctx.activeImg = undefined;
    globalCache.image = undefined;
}
export function onImageMoveStart(ctx, globalCache, e
// { r, c, rc }: { r: number; c: number; rc: string },
) {
    var position = getImagePosition();
    if (position) {
        var top_2 = position.top, left = position.left;
        _.set(globalCache, "image", {
            cursorMoveStartPosition: {
                x: e.pageX,
                y: e.pageY,
            },
            // movingId,
            // imageRC: { r, c, rc },
            imgInitialPosition: { left: left, top: top_2 },
        });
    }
}
export function onImageMove(ctx, globalCache, e) {
    if (ctx.allowEdit === false)
        return false;
    var image = globalCache === null || globalCache === void 0 ? void 0 : globalCache.image;
    var img = document.getElementById("luckysheet-modal-dialog-activeImage");
    if (img && image && !image.resizingSide) {
        var _a = image.cursorMoveStartPosition, startX = _a.x, startY = _a.y;
        var _b = image.imgInitialPosition, top_3 = _b.top, left = _b.left;
        left += e.pageX - startX;
        top_3 += e.pageY - startY;
        if (top_3 < 0)
            top_3 = 0;
        img.style.left = "".concat(left, "px");
        img.style.top = "".concat(top_3, "px");
        return true;
    }
    return false;
}
export function onImageMoveEnd(ctx, globalCache) {
    var _a;
    var position = getImagePosition();
    if (!((_a = globalCache.image) === null || _a === void 0 ? void 0 : _a.resizingSide)) {
        globalCache.image = undefined;
        if (position) {
            var img = _.find(ctx.insertedImgs, function (v) { return v.id === ctx.activeImg; });
            if (img) {
                img.left = position.left / ctx.zoomRatio;
                img.top = position.top / ctx.zoomRatio;
                saveImage(ctx);
            }
        }
    }
}
export function onImageResizeStart(globalCache, e, resizingSide) {
    var position = getImagePosition();
    if (position) {
        _.set(globalCache, "image", {
            cursorMoveStartPosition: { x: e.pageX, y: e.pageY },
            resizingSide: resizingSide,
            imgInitialPosition: position,
        });
    }
}
export function onImageResize(ctx, globalCache, e) {
    if (ctx.allowEdit === false)
        return false;
    var image = globalCache === null || globalCache === void 0 ? void 0 : globalCache.image;
    if (image === null || image === void 0 ? void 0 : image.resizingSide) {
        var imgContainer = document.getElementById("luckysheet-modal-dialog-activeImage");
        var img = imgContainer === null || imgContainer === void 0 ? void 0 : imgContainer.querySelector(".luckysheet-modal-dialog-content");
        if (img == null)
            return false;
        var _a = image.cursorMoveStartPosition, startX = _a.x, startY = _a.y;
        var _b = image.imgInitialPosition, top_4 = _b.top, left = _b.left, width = _b.width, height = _b.height;
        var dx = e.pageX - startX;
        var dy = e.pageY - startY;
        var minHeight = 60 * ctx.zoomRatio;
        var minWidth = 1.5 * 60 * ctx.zoomRatio;
        if (["lm", "lt", "lb"].includes(image.resizingSide)) {
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
            img.style.left = "".concat(left, "px");
            imgContainer.style.left = "".concat(left, "px");
        }
        if (["rm", "rt", "rb"].includes(image.resizingSide)) {
            width = width + dx < minWidth ? minWidth : width + dx;
        }
        if (["mt", "lt", "rt"].includes(image.resizingSide)) {
            if (height - dy < minHeight) {
                top_4 += height - minHeight;
                height = minHeight;
            }
            else {
                top_4 += dy;
                height -= dy;
            }
            if (top_4 < 0)
                top_4 = 0;
            img.style.top = "".concat(top_4, "px");
            imgContainer.style.top = "".concat(top_4, "px");
        }
        if (["mb", "lb", "rb"].includes(image.resizingSide)) {
            height = height + dy < minHeight ? minHeight : height + dy;
        }
        img.style.width = "".concat(width, "px");
        imgContainer.style.width = "".concat(width, "px");
        img.style.height = "".concat(height, "px");
        imgContainer.style.height = "".concat(height, "px");
        img.style.backgroundSize = "".concat(width, "px ").concat(height, "px");
        return true;
    }
    return false;
}
export function onImageResizeEnd(ctx, globalCache) {
    var _a;
    if ((_a = globalCache.image) === null || _a === void 0 ? void 0 : _a.resizingSide) {
        globalCache.image = undefined;
        var position = getImagePosition();
        if (position) {
            var img = _.find(ctx.insertedImgs, function (v) { return v.id === ctx.activeImg; });
            if (img) {
                img.left = position.left / ctx.zoomRatio;
                img.top = position.top / ctx.zoomRatio;
                img.width = position.width / ctx.zoomRatio;
                img.height = position.height / ctx.zoomRatio;
                saveImage(ctx);
            }
        }
    }
}
