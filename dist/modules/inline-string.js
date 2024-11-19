import _ from "lodash";
import { getCellValue, getFontStyleByCell } from "./cell";
import { selectTextContent, selectTextContentCross } from "./cursor";
export var attrToCssName = {
    bl: "font-weight",
    it: "font-style",
    ff: "font-family",
    fs: "font-size",
    fc: "color",
    cl: "text-decoration",
    un: "border-bottom",
};
export var inlineStyleAffectAttribute = {
    bl: 1,
    it: 1,
    ff: 1,
    cl: 1,
    un: 1,
    fs: 1,
    fc: 1,
};
export var inlineStyleAffectCssName = {
    "font-weight": 1,
    "font-style": 1,
    "font-family": 1,
    "text-decoration": 1,
    "border-bottom": 1,
    "font-size": 1,
    color: 1,
};
export function isInlineStringCell(cell) {
    var _a, _b, _c, _d;
    return ((_a = cell === null || cell === void 0 ? void 0 : cell.ct) === null || _a === void 0 ? void 0 : _a.t) === "inlineStr" && ((_d = (_c = (_b = cell === null || cell === void 0 ? void 0 : cell.ct) === null || _b === void 0 ? void 0 : _b.s) === null || _c === void 0 ? void 0 : _c.length) !== null && _d !== void 0 ? _d : 0) > 0;
}
export function isInlineStringCT(ct) {
    var _a, _b;
    return (ct === null || ct === void 0 ? void 0 : ct.t) === "inlineStr" && ((_b = (_a = ct === null || ct === void 0 ? void 0 : ct.s) === null || _a === void 0 ? void 0 : _a.length) !== null && _b !== void 0 ? _b : 0) > 0;
}
export function getInlineStringNoStyle(r, c, data) {
    var ct = getCellValue(r, c, data, "ct");
    if (isInlineStringCT(ct)) {
        var strings = ct.s;
        var value = "";
        for (var i = 0; i < strings.length; i += 1) {
            var strObj = strings[i];
            if (strObj.v) {
                value += strObj.v;
            }
        }
        return value;
    }
    return "";
}
export function convertCssToStyleList(cssText, originCell) {
    if (_.isEmpty(cssText)) {
        return {};
    }
    var cssTextArray = cssText.split(";");
    var styleList = {
        // ff: locale_fontarray[0], // font family
        fc: originCell.fc || "#000000",
        fs: originCell.fs || 10,
        cl: originCell.cl || 0,
        un: originCell.un || 0,
        bl: originCell.bl || 0,
        it: originCell.it || 0,
        ff: originCell.ff || 0, // font family
    };
    cssTextArray.forEach(function (s) {
        s = s.toLowerCase();
        var key = _.trim(s.substring(0, s.indexOf(":")));
        var value = _.trim(s.substring(s.indexOf(":") + 1));
        if (key === "font-weight") {
            if (value === "bold") {
                styleList.bl = 1;
            }
        }
        if (key === "font-style") {
            if (value === "italic") {
                styleList.it = 1;
            }
        }
        // if (key === "font-family") {
        //   const ff = locale_fontjson[value];
        //   if (ff === null) {
        //     styleList.ff = value;
        //   } else {
        //     styleList.ff = ff;
        //   }
        // }
        if (key === "font-size") {
            styleList.fs = parseInt(value, 10);
        }
        if (key === "color") {
            styleList.fc = value;
        }
        if (key === "text-decoration") {
            styleList.cl = 1;
        }
        if (key === "border-bottom") {
            styleList.un = 1;
        }
        if (key === "lucky-strike") {
            styleList.cl = Number(value);
        }
        if (key === "lucky-underline") {
            styleList.un = Number(value);
        }
    });
    return styleList;
}
export function convertSpanToShareString(
// eslint-disable-next-line no-undef
$dom, originCell) {
    var styles = [];
    var preStyleList;
    var preStyleListString = null;
    for (var i = 0; i < $dom.length; i += 1) {
        var span = $dom[i];
        var styleList = convertCssToStyleList(span.style.cssText, originCell);
        var curStyleListString = JSON.stringify(styleList);
        // let v = span.innerHTML;
        var v = span.innerText;
        v = v.replace(/\n/g, "\r\n");
        if (i === $dom.length - 1) {
            if (v.endsWith("\r\n") && !v.endsWith("\r\n\r\n")) {
                v = v.slice(0, v.length - 2);
            }
        }
        if (curStyleListString === preStyleListString) {
            preStyleList.v += v;
        }
        else {
            styleList.v = v;
            styles.push(styleList);
            preStyleListString = curStyleListString;
            preStyleList = styleList;
        }
    }
    return styles;
}
export function updateInlineStringFormatOutside(cell, key, value) {
    if (_.isNil(cell.ct)) {
        return;
    }
    var s = cell.ct.s;
    if (_.isNil(s)) {
        return;
    }
    for (var i = 0; i < s.length; i += 1) {
        var item = s[i];
        item[key] = value;
    }
}
function getClassWithcss(cssText, ukey) {
    var cssTextArray = cssText.split(";");
    if (ukey == null || ukey.length === 0) {
        return cssText;
    }
    if (cssText.indexOf(ukey) > -1) {
        for (var i = 0; i < cssTextArray.length; i += 1) {
            var s = cssTextArray[i];
            s = s.toLowerCase();
            var key = _.trim(s.substring(0, s.indexOf(":")));
            var value = _.trim(s.substring(s.indexOf(":") + 1));
            if (key === ukey) {
                return value;
            }
        }
    }
    return "";
}
function upsetClassWithCss(cssText, ukey, uvalue) {
    var cssTextArray = cssText.split(";");
    var newCss = "";
    if (ukey == null || ukey.length === 0) {
        return cssText;
    }
    if (cssText.indexOf(ukey) > -1) {
        for (var i = 0; i < cssTextArray.length; i += 1) {
            var s = cssTextArray[i];
            s = s.toLowerCase();
            var key = _.trim(s.substring(0, s.indexOf(":")));
            var value = _.trim(s.substring(s.indexOf(":") + 1));
            if (key === ukey) {
                newCss += "".concat(key, ":").concat(uvalue, ";");
            }
            else if (key.length > 0) {
                newCss += "".concat(key, ":").concat(value, ";");
            }
        }
    }
    else if (ukey.length > 0) {
        cssText += "".concat(ukey, ":").concat(uvalue, ";");
        newCss = cssText;
    }
    return newCss;
}
function removeClassWidthCss(cssText, ukey) {
    var cssTextArray = cssText.split(";");
    var newCss = "";
    var oUkey = ukey;
    if (ukey == null || ukey.length === 0) {
        return cssText;
    }
    if (ukey in attrToCssName) {
        // @ts-ignore
        ukey = attrToCssName[ukey];
    }
    if (cssText.indexOf(ukey) > -1) {
        for (var i = 0; i < cssTextArray.length; i += 1) {
            var s = cssTextArray[i];
            s = s.toLowerCase();
            var key = _.trim(s.substring(0, s.indexOf(":")));
            var value = _.trim(s.substring(s.indexOf(":") + 1));
            if (key === ukey ||
                (oUkey === "cl" && key === "lucky-strike") ||
                (oUkey === "un" && key === "lucky-underline")) {
                continue;
            }
            else if (key.length > 0) {
                newCss += "".concat(key, ":").concat(value, ";");
            }
        }
    }
    else {
        newCss = cssText;
    }
    return newCss;
}
function getCssText(cssText, attr, value) {
    var styleObj = {};
    styleObj[attr] = value;
    if (attr === "un") {
        var fontColor = getClassWithcss(cssText, "color");
        if (fontColor === "") {
            fontColor = "#000000";
        }
        var fs = getClassWithcss(cssText, "font-size");
        if (fs === "") {
            fs = "11";
        }
        styleObj._fontSize = Number(fs);
        styleObj._color = fontColor;
    }
    var s = getFontStyleByCell(styleObj, undefined, undefined, false);
    var ukey = _.kebabCase(Object.keys(s)[0]);
    var uvalue = Object.values(s)[0];
    // let cssText = span.style.cssText;
    cssText = removeClassWidthCss(cssText, attr);
    cssText = upsetClassWithCss(cssText, ukey, uvalue);
    return cssText;
}
function extendCssText(origin, cover, isLimit) {
    if (isLimit === void 0) { isLimit = true; }
    var originArray = origin.split(";");
    var coverArray = cover.split(";");
    var newCss = "";
    var addKeyList = {};
    for (var i = 0; i < originArray.length; i += 1) {
        var so = originArray[i];
        var isAdd = true;
        so = so.toLowerCase();
        var okey = _.trim(so.substring(0, so.indexOf(":")));
        /* 不设置文字的大小，解决设置删除线等后字体变大的问题 */
        if (okey === "font-size") {
            continue;
        }
        var ovalue = _.trim(so.substring(so.indexOf(":") + 1));
        if (isLimit) {
            if (!(okey in inlineStyleAffectCssName)) {
                continue;
            }
        }
        for (var a = 0; a < coverArray.length; a += 1) {
            var sc = coverArray[a];
            sc = sc.toLowerCase();
            var ckey = _.trim(sc.substring(0, sc.indexOf(":")));
            var cvalue = _.trim(sc.substring(sc.indexOf(":") + 1));
            if (okey === ckey) {
                newCss += "".concat(ckey, ":").concat(cvalue, ";");
                isAdd = false;
                continue;
            }
        }
        if (isAdd) {
            newCss += "".concat(okey, ":").concat(ovalue, ";");
        }
        addKeyList[okey] = 1;
    }
    for (var a = 0; a < coverArray.length; a += 1) {
        var sc = coverArray[a];
        sc = sc.toLowerCase();
        var ckey = _.trim(sc.substring(0, sc.indexOf(":")));
        var cvalue = _.trim(sc.substring(sc.indexOf(":") + 1));
        if (isLimit) {
            if (!(ckey in inlineStyleAffectCssName)) {
                continue;
            }
        }
        if (!(ckey in addKeyList)) {
            newCss += "".concat(ckey, ":").concat(cvalue, ";");
        }
    }
    return newCss;
}
export function updateInlineStringFormat(ctx, cell, attr, value, cellInput) {
    var _a, _b, _c;
    // let s = ctx.inlineStringEditCache;
    var w = window.getSelection();
    if (!w)
        return;
    if (w.rangeCount === 0)
        return;
    var range = w.getRangeAt(0);
    var $textEditor = cellInput;
    if (range.collapsed === true) {
        return;
    }
    var endContainer = range.endContainer;
    var startContainer = range.startContainer;
    var endOffset = range.endOffset;
    var startOffset = range.startOffset;
    if ($textEditor) {
        if (startContainer === endContainer) {
            var span = startContainer.parentNode;
            var spanIndex = void 0;
            var inherit = false;
            var content = (span === null || span === void 0 ? void 0 : span.innerHTML) || "";
            var fullContent = $textEditor.innerHTML;
            if (fullContent.substring(0, 5) !== "<span") {
                inherit = true;
            }
            if (span) {
                var left = "";
                var mid = "";
                var right = "";
                var s1 = 0;
                var s2 = startOffset;
                var s3 = endOffset;
                var s4 = content.length;
                left = content.substring(s1, s2);
                mid = content.substring(s2, s3);
                right = content.substring(s3, s4);
                var cont = "";
                if (left !== "") {
                    var cssText = span.style.cssText;
                    if (inherit) {
                        var box = span.closest("#luckysheet-input-box");
                        if (box != null) {
                            cssText = extendCssText(box.style.cssText, cssText);
                        }
                    }
                    cont += "<span style='".concat(cssText, "'>").concat(left, "</span>");
                }
                if (mid !== "") {
                    var cssText = getCssText(span.style.cssText, attr, value);
                    if (inherit) {
                        var box = span.closest("#luckysheet-input-box");
                        if (box != null) {
                            cssText = extendCssText(box.style.cssText, cssText);
                        }
                    }
                    cont += "<span style='".concat(cssText, "'>").concat(mid, "</span>");
                }
                if (right !== "") {
                    var cssText = span.style.cssText;
                    if (inherit) {
                        var box = span.closest("#luckysheet-input-box");
                        if (box != null) {
                            cssText = extendCssText(box.style.cssText, cssText);
                        }
                    }
                    cont += "<span style='".concat(cssText, "'>").concat(right, "</span>");
                }
                if (((_a = startContainer.parentElement) === null || _a === void 0 ? void 0 : _a.tagName) === "SPAN") {
                    spanIndex = _.indexOf($textEditor.querySelectorAll("span"), span);
                    span.outerHTML = cont;
                }
                else {
                    spanIndex = 0;
                    span.innerHTML = cont;
                }
                var seletedNodeIndex = 0;
                if (s1 === s2) {
                    seletedNodeIndex = spanIndex;
                }
                else {
                    seletedNodeIndex = spanIndex + 1;
                }
                selectTextContent($textEditor.querySelectorAll("span")[seletedNodeIndex]);
            }
        }
        else {
            if (((_b = startContainer.parentElement) === null || _b === void 0 ? void 0 : _b.tagName) === "SPAN" &&
                ((_c = endContainer.parentElement) === null || _c === void 0 ? void 0 : _c.tagName) === "SPAN") {
                var startSpan = startContainer.parentNode;
                var endSpan = endContainer.parentNode;
                var allSpans = $textEditor.querySelectorAll("span");
                var startSpanIndex = _.indexOf(allSpans, startSpan);
                var endSpanIndex = _.indexOf(allSpans, endSpan);
                var startContent = (startSpan === null || startSpan === void 0 ? void 0 : startSpan.innerHTML) || "";
                var endContent = (endSpan === null || endSpan === void 0 ? void 0 : endSpan.innerHTML) || "";
                var sleft = "";
                var sright = "";
                var eleft = "";
                var eright = "";
                var s1 = 0;
                var s2 = startOffset;
                var s3 = endOffset;
                var s4 = endContent.length;
                sleft = startContent.substring(s1, s2);
                sright = startContent.substring(s2, startContent.length);
                eleft = endContent.substring(0, s3);
                eright = endContent.substring(s3, s4);
                var spans = $textEditor.querySelectorAll("span");
                // const replaceSpans = spans.slice(startSpanIndex, endSpanIndex + 1);
                var cont = "";
                for (var i = 0; i < startSpanIndex; i += 1) {
                    var span = spans[i];
                    var content = span.innerHTML;
                    cont += "<span style='".concat(span.style.cssText, "'>").concat(content, "</span>");
                }
                if (sleft !== "") {
                    cont += "<span style='".concat(startSpan.style.cssText, "'>").concat(sleft, "</span>");
                }
                if (sright !== "") {
                    var cssText = getCssText(startSpan.style.cssText, attr, value);
                    cont += "<span style='".concat(cssText, "'>").concat(sright, "</span>");
                }
                if (startSpanIndex < endSpanIndex) {
                    for (var i = startSpanIndex + 1; i < endSpanIndex; i += 1) {
                        var span = spans[i];
                        var content = span.innerHTML;
                        cont += "<span style='".concat(span.style.cssText, "'>").concat(content, "</span>");
                    }
                }
                if (eleft !== "") {
                    var cssText = getCssText(endSpan.style.cssText, attr, value);
                    cont += "<span style='".concat(cssText, "'>").concat(eleft, "</span>");
                }
                if (eright !== "") {
                    cont += "<span style='".concat(endSpan.style.cssText, "'>").concat(eright, "</span>");
                }
                for (var i = endSpanIndex + 1; i < spans.length; i += 1) {
                    var span = spans[i];
                    var content = span.innerHTML;
                    cont += "<span style='".concat(span.style.cssText, "'>").concat(content, "</span>");
                }
                $textEditor.innerHTML = cont;
                // console.log(replaceSpans, cont);
                // replaceSpans.replaceWith(cont);
                var startSeletedNodeIndex = void 0;
                var endSeletedNodeIndex = void 0;
                if (s1 === s2) {
                    startSeletedNodeIndex = startSpanIndex;
                    endSeletedNodeIndex = endSpanIndex;
                }
                else {
                    startSeletedNodeIndex = startSpanIndex + 1;
                    endSeletedNodeIndex = endSpanIndex + 1;
                }
                spans = $textEditor.querySelectorAll("span");
                selectTextContentCross(spans[startSeletedNodeIndex], spans[endSeletedNodeIndex]);
            }
        }
    }
}
