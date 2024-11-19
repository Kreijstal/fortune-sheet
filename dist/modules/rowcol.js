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
import { getSheetIndex } from "../utils";
import { getcellFormula } from "./cell";
import { functionStrChange } from "./formula";
var refreshLocalMergeData = function (merge_new, file) {
    Object.entries(merge_new).forEach(function (_a) {
        var _b, _c, _d, _e;
        var v = _a[1];
        var _f = v, r = _f.r, c = _f.c, rs = _f.rs, cs = _f.cs;
        for (var i = r; i < r + rs; i += 1) {
            for (var j = c; j < c + cs; j += 1) {
                if ((_c = (_b = file === null || file === void 0 ? void 0 : file.data) === null || _b === void 0 ? void 0 : _b[i]) === null || _c === void 0 ? void 0 : _c[j]) {
                    file.data[i][j] = __assign(__assign({}, file.data[i][j]), { mc: { r: r, c: c } });
                }
            }
        }
        if ((_e = (_d = file === null || file === void 0 ? void 0 : file.data) === null || _d === void 0 ? void 0 : _d[r]) === null || _e === void 0 ? void 0 : _e[c]) {
            file.data[r][c] = __assign(__assign({}, file.data[r][c]), { mc: { r: r, c: c, rs: rs, cs: cs } });
        }
    });
};
/**
 * 增加行列
 * @param {string} type 行或列 ['row', 'column'] 之一
 * @param {number} index 插入的位置 index
 * @param {number} count 插入 多少 行（列）
 * @param {string} direction 哪个方向插入 ['lefttop','rightbottom'] 之一
 * @param {string | number} id 操作的 sheet 的 id
 * @returns
 */
export function insertRowCol(ctx, op, changeSelection) {
    var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o, _p, _q, _r, _s, _t, _u, _v, _w;
    if (changeSelection === void 0) { changeSelection = true; }
    var count = op.count, id = op.id;
    var type = op.type, index = op.index, direction = op.direction;
    id = id || ctx.currentSheetId;
    // if (
    //   type === "row" &&
    //   !checkProtectionAuthorityNormal(sheetId, "insertRows")
    // ) {
    //   return;
    // } else if (
    //   type === "column" &&
    //   !checkProtectionAuthorityNormal(sheetId, "insertColumns")
    // ) {
    //   return;
    // }
    var curOrder = getSheetIndex(ctx, id);
    if (curOrder == null)
        return;
    var file = ctx.luckysheetfile[curOrder];
    if (!file)
        return;
    var d = file.data;
    if (!d)
        return;
    var cfg = file.config || {};
    if (changeSelection) {
        if (type === "row") {
            if ((_a = cfg.rowReadOnly) === null || _a === void 0 ? void 0 : _a[index]) {
                throw new Error("readOnly");
            }
        }
        else {
            if ((_b = cfg.colReadOnly) === null || _b === void 0 ? void 0 : _b[index]) {
                throw new Error("readOnly");
            }
        }
    }
    if (type === "row" && d.length + count >= 10000) {
        throw new Error("maxExceeded");
    }
    if (type === "column" && d[0] && d[0].length + count >= 1000) {
        throw new Error("maxExceeded");
    }
    count = Math.floor(count);
    // 合并单元格配置变动
    if (cfg.merge == null) {
        cfg.merge = {};
    }
    var merge_new = {};
    _.forEach(cfg.merge, function (mc) {
        var r = mc.r, c = mc.c, rs = mc.rs, cs = mc.cs;
        if (type === "row") {
            if (index < r) {
                merge_new["".concat(r + count, "_").concat(c)] = { r: r + count, c: c, rs: rs, cs: cs };
            }
            else if (index === r) {
                if (direction === "lefttop") {
                    merge_new["".concat(r + count, "_").concat(c)] = {
                        r: r + count,
                        c: c,
                        rs: rs,
                        cs: cs,
                    };
                }
                else {
                    merge_new["".concat(r, "_").concat(c)] = { r: r, c: c, rs: rs + count, cs: cs };
                }
            }
            else if (index < r + rs - 1) {
                merge_new["".concat(r, "_").concat(c)] = { r: r, c: c, rs: rs + count, cs: cs };
            }
            else if (index === r + rs - 1) {
                if (direction === "lefttop") {
                    merge_new["".concat(r, "_").concat(c)] = { r: r, c: c, rs: rs + count, cs: cs };
                }
                else {
                    merge_new["".concat(r, "_").concat(c)] = { r: r, c: c, rs: rs, cs: cs };
                }
            }
            else {
                merge_new["".concat(r, "_").concat(c)] = { r: r, c: c, rs: rs, cs: cs };
            }
        }
        else if (type === "column") {
            if (index < c) {
                merge_new["".concat(r, "_").concat(c + count)] = {
                    r: r,
                    c: c + count,
                    rs: rs,
                    cs: cs,
                };
            }
            else if (index === c) {
                if (direction === "lefttop") {
                    merge_new["".concat(r, "_").concat(c + count)] = {
                        r: r,
                        c: c + count,
                        rs: rs,
                        cs: cs,
                    };
                }
                else {
                    merge_new["".concat(r, "_").concat(c)] = { r: r, c: c, rs: rs, cs: cs + count };
                }
            }
            else if (index < c + cs - 1) {
                merge_new["".concat(r, "_").concat(c)] = { r: r, c: c, rs: rs, cs: cs + count };
            }
            else if (index === c + cs - 1) {
                if (direction === "lefttop") {
                    merge_new["".concat(r, "_").concat(c)] = { r: r, c: c, rs: rs, cs: cs + count };
                }
                else {
                    merge_new["".concat(r, "_").concat(c)] = { r: r, c: c, rs: rs, cs: cs };
                }
            }
            else {
                merge_new["".concat(r, "_").concat(c)] = { r: r, c: c, rs: rs, cs: cs };
            }
        }
    });
    cfg.merge = merge_new;
    // 公式配置变动
    var newCalcChain = [];
    for (var SheetIndex = 0; SheetIndex < ctx.luckysheetfile.length; SheetIndex += 1) {
        if (_.isNil(ctx.luckysheetfile[SheetIndex].calcChain) ||
            ctx.luckysheetfile.length === 0) {
            continue;
        }
        var calcChain = ctx.luckysheetfile[SheetIndex].calcChain;
        var data = ctx.luckysheetfile[SheetIndex].data;
        for (var i = 0; i < calcChain.length; i += 1) {
            var calc = _.cloneDeep(calcChain[i]);
            var calc_r = calc.r;
            var calc_c = calc.c;
            var calc_i = calc.id;
            var calc_funcStr = getcellFormula(ctx, calc_r, calc_c, calc_i);
            if (type === "row" && SheetIndex === curOrder) {
                var functionStr = "=".concat(functionStrChange(calc_funcStr, "add", "row", direction, index, count));
                if (((_d = (_c = d[calc_r]) === null || _c === void 0 ? void 0 : _c[calc_c]) === null || _d === void 0 ? void 0 : _d.f) === calc_funcStr) {
                    d[calc_r][calc_c].f = functionStr;
                }
                if (direction === "lefttop") {
                    if (calc_r >= index) {
                        calc.r += count;
                    }
                }
                else if (direction === "rightbottom") {
                    if (calc_r > index) {
                        calc.r += count;
                    }
                }
                newCalcChain.push(calc);
            }
            else if (type === "row") {
                var functionStr = "=".concat(functionStrChange(calc_funcStr, "add", "row", direction, index, count));
                if (((_f = (_e = data[calc_r]) === null || _e === void 0 ? void 0 : _e[calc_c]) === null || _f === void 0 ? void 0 : _f.f) === calc_funcStr) {
                    data[calc_r][calc_c].f = functionStr;
                }
            }
            else if (type === "column" && SheetIndex === curOrder) {
                var functionStr = "=".concat(functionStrChange(calc_funcStr, "add", "col", direction, index, count));
                if (((_h = (_g = d[calc_r]) === null || _g === void 0 ? void 0 : _g[calc_c]) === null || _h === void 0 ? void 0 : _h.f) === calc_funcStr) {
                    d[calc_r][calc_c].f = functionStr;
                }
                if (direction === "lefttop") {
                    if (calc_c >= index) {
                        calc.c += count;
                    }
                }
                else if (direction === "rightbottom") {
                    if (calc_c > index) {
                        calc.c += count;
                    }
                }
                newCalcChain.push(calc);
            }
            else if (type === "column") {
                var functionStr = "=".concat(functionStrChange(calc_funcStr, "add", "col", direction, index, count));
                if (((_k = (_j = data[calc_r]) === null || _j === void 0 ? void 0 : _j[calc_c]) === null || _k === void 0 ? void 0 : _k.f) === calc_funcStr) {
                    data[calc_r][calc_c].f = functionStr;
                }
            }
        }
    }
    // 筛选配置变动
    var filter_select = file.filter_select;
    var filter = file.filter;
    var newFilterObj = null;
    if (!_.isEmpty(filter_select) && filter_select != null) {
        newFilterObj = { filter_select: null, filter: null };
        var f_r1_1 = filter_select.row[0];
        var f_r2_1 = filter_select.row[1];
        var f_c1_1 = filter_select.column[0];
        var f_c2_1 = filter_select.column[1];
        if (type === "row") {
            if (f_r1_1 < index) {
                if (f_r2_1 === index && direction === "lefttop") {
                    f_r2_1 += count;
                }
                else if (f_r2_1 > index) {
                    f_r2_1 += count;
                }
            }
            else if (f_r1_1 === index) {
                if (direction === "lefttop") {
                    f_r1_1 += count;
                    f_r2_1 += count;
                }
                else if (direction === "rightbottom" && f_r2_1 > index) {
                    f_r2_1 += count;
                }
            }
            else {
                f_r1_1 += count;
                f_r2_1 += count;
            }
            if (filter != null) {
                newFilterObj.filter = {};
                _.forEach(filter, function (v, k) {
                    var f_rowhidden = filter[k].rowhidden;
                    var f_rowhidden_new = {};
                    _.forEach(f_rowhidden, function (v1, nstr) {
                        var n = parseFloat(nstr);
                        if (n < index) {
                            f_rowhidden_new[n] = 0;
                        }
                        else if (n === index) {
                            if (direction === "lefttop") {
                                f_rowhidden_new[n + count] = 0;
                            }
                            else if (direction === "rightbottom") {
                                f_rowhidden_new[n] = 0;
                            }
                        }
                        else {
                            f_rowhidden_new[n + count] = 0;
                        }
                    });
                    newFilterObj.filter[k] = _.cloneDeep(filter[k]);
                    newFilterObj.filter[k].rowhidden = f_rowhidden_new;
                    newFilterObj.filter[k].str = f_r1_1;
                    newFilterObj.filter[k].edr = f_r2_1;
                });
            }
        }
        else if (type === "column") {
            if (f_c1_1 < index) {
                if (f_c2_1 === index && direction === "lefttop") {
                    f_c2_1 += count;
                }
                else if (f_c2_1 > index) {
                    f_c2_1 += count;
                }
            }
            else if (f_c1_1 === index) {
                if (direction === "lefttop") {
                    f_c1_1 += count;
                    f_c2_1 += count;
                }
                else if (direction === "rightbottom" && f_c2_1 > index) {
                    f_c2_1 += count;
                }
            }
            else {
                f_c1_1 += count;
                f_c2_1 += count;
            }
            if (filter != null) {
                newFilterObj.filter = {};
                _.forEach(filter, function (v, k) {
                    var f_cindex = filter[k].cindex;
                    if (f_cindex === index && direction === "lefttop") {
                        f_cindex += count;
                    }
                    else if (f_cindex > index) {
                        f_cindex += count;
                    }
                    newFilterObj.filter[f_cindex - f_c1_1] = _.cloneDeep(filter[k]);
                    newFilterObj.filter[f_cindex - f_c1_1].cindex = f_cindex;
                    newFilterObj.filter[f_cindex - f_c1_1].stc = f_c1_1;
                    newFilterObj.filter[f_cindex - f_c1_1].edc = f_c2_1;
                });
            }
        }
        newFilterObj.filter_select = { row: [f_r1_1, f_r2_1], column: [f_c1_1, f_c2_1] };
    }
    if (newFilterObj != null && newFilterObj.filter != null) {
        if (cfg.rowhidden == null) {
            cfg.rowhidden = {};
        }
        _.forEach(newFilterObj.filter, function (v, k) {
            var f_rowhidden = newFilterObj.filter[k].rowhidden;
            _.forEach(f_rowhidden, function (v1, n) {
                cfg.rowhidden[n] = 0;
            });
        });
    }
    // 条件格式配置变动
    var CFarr = file.luckysheet_conditionformat_save;
    var newCFarr = [];
    if (CFarr != null && CFarr.length > 0) {
        for (var i = 0; i < CFarr.length; i += 1) {
            var cf_range = CFarr[i].cellrange;
            var cf_new_range = [];
            for (var j = 0; j < cf_range.length; j += 1) {
                var CFr1 = cf_range[j].row[0];
                var CFr2 = cf_range[j].row[1];
                var CFc1 = cf_range[j].column[0];
                var CFc2 = cf_range[j].column[1];
                if (type === "row") {
                    if (CFr1 < index) {
                        if (CFr2 === index && direction === "lefttop") {
                            CFr2 += count;
                        }
                        else if (CFr2 > index) {
                            CFr2 += count;
                        }
                    }
                    else if (CFr1 === index) {
                        if (direction === "lefttop") {
                            CFr1 += count;
                            CFr2 += count;
                        }
                        else if (direction === "rightbottom" && CFr2 > index) {
                            CFr2 += count;
                        }
                    }
                    else {
                        CFr1 += count;
                        CFr2 += count;
                    }
                }
                else if (type === "column") {
                    if (CFc1 < index) {
                        if (CFc2 === index && direction === "lefttop") {
                            CFc2 += count;
                        }
                        else if (CFc2 > index) {
                            CFc2 += count;
                        }
                    }
                    else if (CFc1 === index) {
                        if (direction === "lefttop") {
                            CFc1 += count;
                            CFc2 += count;
                        }
                        else if (direction === "rightbottom" && CFc2 > index) {
                            CFc2 += count;
                        }
                    }
                    else {
                        CFc1 += count;
                        CFc2 += count;
                    }
                }
                cf_new_range.push({ row: [CFr1, CFr2], column: [CFc1, CFc2] });
            }
            var cf = _.clone(CFarr[i]);
            cf.cellrange = cf_new_range;
            newCFarr.push(cf);
        }
    }
    // 交替颜色配置变动
    var AFarr = file.luckysheet_alternateformat_save;
    var newAFarr = [];
    if (AFarr != null && AFarr.length > 0) {
        for (var i = 0; i < AFarr.length; i += 1) {
            var AFr1 = AFarr[i].cellrange.row[0];
            var AFr2 = AFarr[i].cellrange.row[1];
            var AFc1 = AFarr[i].cellrange.column[0];
            var AFc2 = AFarr[i].cellrange.column[1];
            var af = _.clone(AFarr[i]);
            if (type === "row") {
                if (AFr1 < index) {
                    if (AFr2 === index && direction === "lefttop") {
                        AFr2 += count;
                    }
                    else if (AFr2 > index) {
                        AFr2 += count;
                    }
                }
                else if (AFr1 === index) {
                    if (direction === "lefttop") {
                        AFr1 += count;
                        AFr2 += count;
                    }
                    else if (direction === "rightbottom" && AFr2 > index) {
                        AFr2 += count;
                    }
                }
                else {
                    AFr1 += count;
                    AFr2 += count;
                }
            }
            else if (type === "column") {
                if (AFc1 < index) {
                    if (AFc2 === index && direction === "lefttop") {
                        AFc2 += count;
                    }
                    else if (AFc2 > index) {
                        AFc2 += count;
                    }
                }
                else if (AFc1 === index) {
                    if (direction === "lefttop") {
                        AFc1 += count;
                        AFc2 += count;
                    }
                    else if (direction === "rightbottom" && AFc2 > index) {
                        AFc2 += count;
                    }
                }
                else {
                    AFc1 += count;
                    AFc2 += count;
                }
            }
            af.cellrange = { row: [AFr1, AFr2], column: [AFc1, AFc2] };
            newAFarr.push(af);
        }
    }
    // 冻结配置变动
    var frozen = file.frozen;
    if (frozen) {
        var normalizedIndex = direction === "lefttop" ? index - 1 : index;
        if (type === "row" &&
            (frozen.type === "rangeRow" || frozen.type === "rangeBoth")) {
            if (((_m = (_l = frozen.range) === null || _l === void 0 ? void 0 : _l.row_focus) !== null && _m !== void 0 ? _m : -1) > normalizedIndex) {
                frozen.range.row_focus += count;
            }
        }
        if (type === "column" &&
            (frozen.type === "rangeColumn" || frozen.type === "rangeBoth")) {
            if (((_p = (_o = frozen.range) === null || _o === void 0 ? void 0 : _o.column_focus) !== null && _p !== void 0 ? _p : -1) > normalizedIndex) {
                frozen.range.column_focus += count;
            }
        }
    }
    // 数据验证配置变动
    var dataVerification = file.dataVerification;
    var newDataVerification = {};
    if (dataVerification != null) {
        _.forEach(dataVerification, function (v, key) {
            var r = Number(key.split("_")[0]);
            var c = Number(key.split("_")[1]);
            var item = dataVerification[key];
            if (type === "row") {
                if (index < r) {
                    newDataVerification["".concat(r + count, "_").concat(c)] = item;
                }
                else if (index === r) {
                    if (direction === "lefttop") {
                        newDataVerification["".concat(r + count, "_").concat(c)] = item;
                        for (var i = 0; i < count; i += 1) {
                            newDataVerification["".concat(r + i, "_").concat(c)] = item;
                        }
                    }
                    else {
                        newDataVerification["".concat(r, "_").concat(c)] = item;
                        for (var i = 0; i < count; i += 1) {
                            newDataVerification["".concat(r + i + 1, "_").concat(c)] = item;
                        }
                    }
                }
                else {
                    newDataVerification["".concat(r, "_").concat(c)] = item;
                }
            }
            else if (type === "column") {
                if (index < c) {
                    newDataVerification["".concat(r, "_").concat(c + count)] = item;
                }
                else if (index === c) {
                    if (direction === "lefttop") {
                        newDataVerification["".concat(r, "_").concat(c + count)] = item;
                        for (var i = 0; i < count; i += 1) {
                            newDataVerification["".concat(r, "_").concat(c + i)] = item;
                        }
                    }
                    else {
                        newDataVerification["".concat(r, "_").concat(c)] = item;
                        for (var i = 0; i < count; i += 1) {
                            newDataVerification["".concat(r, "_").concat(c + i + 1)] = item;
                        }
                    }
                }
                else {
                    newDataVerification["".concat(r, "_").concat(c)] = item;
                }
            }
        });
    }
    // 超链接配置变动
    var hyperlink = file.hyperlink;
    var newHyperlink = {};
    if (hyperlink != null) {
        _.forEach(hyperlink, function (v, key) {
            var r = Number(key.split("_")[0]);
            var c = Number(key.split("_")[1]);
            var item = hyperlink[key];
            if (type === "row") {
                if (index < r) {
                    newHyperlink["".concat(r + count, "_").concat(c)] = item;
                }
                else if (index === r) {
                    if (direction === "lefttop") {
                        newHyperlink["".concat(r + count, "_").concat(c)] = item;
                    }
                    else {
                        newHyperlink["".concat(r, "_").concat(c)] = item;
                    }
                }
                else {
                    newHyperlink["".concat(r, "_").concat(c)] = item;
                }
            }
            else if (type === "column") {
                if (index < c) {
                    newHyperlink["".concat(r, "_").concat(c + count)] = item;
                }
                else if (index === c) {
                    if (direction === "lefttop") {
                        newHyperlink["".concat(r, "_").concat(c + count)] = item;
                    }
                    else {
                        newHyperlink["".concat(r, "_").concat(c)] = item;
                    }
                }
                else {
                    newHyperlink["".concat(r, "_").concat(c)] = item;
                }
            }
        });
    }
    var type1;
    if (type === "row") {
        type1 = "r";
        // 行高配置变动
        if (cfg.rowlen != null) {
            var rowlen_new_1 = {};
            var rowReadOnly_new_1 = {};
            _.forEach(cfg.rowlen, function (v, rstr) {
                var r = parseFloat(rstr);
                if (r < index) {
                    rowlen_new_1[r] = cfg.rowlen[r];
                }
                else if (r === index) {
                    if (direction === "lefttop") {
                        rowlen_new_1[r + count] = cfg.rowlen[r];
                    }
                    else if (direction === "rightbottom") {
                        rowlen_new_1[r] = cfg.rowlen[r];
                    }
                }
                else {
                    rowlen_new_1[r + count] = cfg.rowlen[r];
                }
            });
            _.forEach(cfg.rowReadOnly, function (v, rstr) {
                var r = parseFloat(rstr);
                if (r < index) {
                    rowReadOnly_new_1[r] = cfg.rowReadOnly[r];
                }
                else if (r > index) {
                    rowReadOnly_new_1[r + count] = cfg.rowReadOnly[r];
                }
            });
            cfg.rowlen = rowlen_new_1;
            cfg.rowReadOnly = rowReadOnly_new_1;
        }
        // 自定义行高配置变动
        if (cfg.customHeight != null) {
            var customHeight_new_1 = {};
            _.forEach(cfg.customHeight, function (v, rstr) {
                var r = parseFloat(rstr);
                if (r < index) {
                    customHeight_new_1[r] = cfg.customHeight[r];
                }
                else if (r === index) {
                    if (direction === "lefttop") {
                        customHeight_new_1[r + count] = cfg.customHeight[r];
                    }
                    else if (direction === "rightbottom") {
                        customHeight_new_1[r] = cfg.customHeight[r];
                    }
                }
                else {
                    customHeight_new_1[r + count] = cfg.customHeight[r];
                }
            });
            cfg.customHeight = customHeight_new_1;
        }
        // 自定义行高配置变动
        if (cfg.customHeight != null) {
            var customHeight_new_2 = {};
            _.forEach(cfg.customHeight, function (v, rstr) {
                var r = parseFloat(rstr);
                if (r < index) {
                    customHeight_new_2[r] = cfg.customHeight[r];
                }
                else if (r === index) {
                    if (direction === "lefttop") {
                        customHeight_new_2[r + count] = cfg.customHeight[r];
                    }
                    else if (direction === "rightbottom") {
                        customHeight_new_2[r] = cfg.customHeight[r];
                    }
                }
                else {
                    customHeight_new_2[r + count] = cfg.customHeight[r];
                }
            });
            cfg.customHeight = customHeight_new_2;
        }
        // 隐藏行配置变动
        if (cfg.rowhidden != null) {
            var rowhidden_new_1 = {};
            _.forEach(cfg.rowhidden, function (v, rstr) {
                var r = parseFloat(rstr);
                if (r < index) {
                    rowhidden_new_1[r] = cfg.rowhidden[r];
                }
                else if (r === index) {
                    if (direction === "lefttop") {
                        rowhidden_new_1[r + count] = cfg.rowhidden[r];
                    }
                    else if (direction === "rightbottom") {
                        rowhidden_new_1[r] = cfg.rowhidden[r];
                    }
                }
                else {
                    rowhidden_new_1[r + count] = cfg.rowhidden[r];
                }
            });
            cfg.rowhidden = rowhidden_new_1;
        }
        // 空行模板
        var row = [];
        var curRow = __spreadArray([], d, true)[index];
        for (var c = 0; c < d[0].length; c += 1) {
            var cell = curRow[c];
            var templateCell = null;
            if ((cell === null || cell === void 0 ? void 0 : cell.mc) && (direction === "rightbottom" || index !== cell.mc.r)) {
                if (cell.mc.rs) {
                    cell.mc.rs += count;
                }
                templateCell = __assign({}, cell);
                if (!((_r = (_q = d === null || d === void 0 ? void 0 : d[index + 1]) === null || _q === void 0 ? void 0 : _q[c]) === null || _r === void 0 ? void 0 : _r.mc)) {
                    templateCell.mc = undefined;
                }
                delete templateCell.v;
                delete templateCell.m;
                delete templateCell.ps;
                delete templateCell.f;
            }
            row.push(templateCell);
        }
        var cellBorderConfig = [];
        // 边框
        if (cfg.borderInfo && cfg.borderInfo.length > 0) {
            var borderInfo = [];
            for (var i = 0; i < cfg.borderInfo.length; i += 1) {
                var rangeType = cfg.borderInfo[i].rangeType;
                if (rangeType === "range") {
                    var borderRange = cfg.borderInfo[i].range;
                    var emptyRange = [];
                    for (var j = 0; j < borderRange.length; j += 1) {
                        var bd_r1 = borderRange[j].row[0];
                        var bd_r2 = borderRange[j].row[1];
                        if (direction === "lefttop") {
                            if (index <= bd_r1) {
                                bd_r1 += count;
                                bd_r2 += count;
                            }
                            else if (index <= bd_r2) {
                                bd_r2 += count;
                            }
                        }
                        else {
                            if (index < bd_r1) {
                                bd_r1 += count;
                                bd_r2 += count;
                            }
                            else if (index < bd_r2) {
                                bd_r2 += count;
                            }
                        }
                        if (bd_r2 >= bd_r1) {
                            emptyRange.push({
                                row: [bd_r1, bd_r2],
                                column: borderRange[j].column,
                            });
                        }
                    }
                    if (emptyRange.length > 0) {
                        var bd_obj = {
                            rangeType: "range",
                            borderType: cfg.borderInfo[i].borderType,
                            style: cfg.borderInfo[i].style,
                            color: cfg.borderInfo[i].color,
                            range: emptyRange,
                        };
                        borderInfo.push(bd_obj);
                    }
                }
                else if (rangeType === "cell") {
                    var row_index = cfg.borderInfo[i].value.row_index;
                    // 位置相同标识边框相关 先缓存
                    if (row_index === index) {
                        cellBorderConfig.push(JSON.parse(JSON.stringify(cfg.borderInfo[i])));
                    }
                    if (direction === "lefttop") {
                        if (index <= row_index) {
                            row_index += count;
                        }
                    }
                    else {
                        if (index < row_index) {
                            row_index += count;
                        }
                    }
                    cfg.borderInfo[i].value.row_index = row_index;
                    borderInfo.push(cfg.borderInfo[i]);
                }
            }
            cfg.borderInfo = borderInfo;
        }
        var arr = [];
        var _loop_1 = function (r) {
            arr.push(JSON.stringify(row));
            // 同步拷贝 type 为 cell 类型的边框
            if (cellBorderConfig.length) {
                var cellBorderConfigCopy = _.cloneDeep(cellBorderConfig);
                cellBorderConfigCopy.forEach(function (item) {
                    if (direction === "rightbottom") {
                        // 向下插入时 基于模板行位置直接递增即可
                        item.value.row_index += r + 1;
                    }
                    else if (direction === "lefttop") {
                        // 向上插入时 目标行移动到后面 新增n行到前面 对于新增的行来说 也是递增，不过是从0开始
                        item.value.row_index += r;
                    }
                });
                (_s = cfg.borderInfo) === null || _s === void 0 ? void 0 : _s.push.apply(_s, cellBorderConfigCopy);
            }
        };
        for (var r = 0; r < count; r += 1) {
            _loop_1(r);
        }
        if (direction === "lefttop") {
            if (index === 0) {
                new Function("d", "return d.unshift(".concat(arr.join(","), ")"))(d);
            }
            else {
                new Function("d", "return d.splice(".concat(index, ", 0, ").concat(arr.join(","), ")"))(d);
            }
        }
        else {
            new Function("d", "return d.splice(".concat(index + 1, ", 0, ").concat(arr.join(","), ")"))(d);
        }
    }
    else {
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        type1 = "c";
        // 列宽配置变动
        if (cfg.columnlen != null) {
            var columnlen_new_1 = {};
            var columnReadOnly_new_1 = {};
            _.forEach(cfg.columnlen, function (v, cstr) {
                var c = parseFloat(cstr);
                if (c < index) {
                    columnlen_new_1[c] = cfg.columnlen[c];
                }
                else if (c === index) {
                    if (direction === "lefttop") {
                        columnlen_new_1[c + count] = cfg.columnlen[c];
                    }
                    else if (direction === "rightbottom") {
                        columnlen_new_1[c] = cfg.columnlen[c];
                    }
                }
                else {
                    columnlen_new_1[c + count] = cfg.columnlen[c];
                }
            });
            _.forEach(cfg.colReadOnly, function (v, cstr) {
                var c = parseFloat(cstr);
                if (c < index) {
                    columnReadOnly_new_1[c] = cfg.colReadOnly[c];
                }
                else if (c > index) {
                    columnReadOnly_new_1[c + count] = cfg.colReadOnly[c];
                }
            });
            cfg.columnlen = columnlen_new_1;
            cfg.colReadOnly = columnReadOnly_new_1;
        }
        // 自定义列宽配置变动
        if (cfg.customWidth != null) {
            var customWidth_new_1 = {};
            _.forEach(cfg.customWidth, function (v, cstr) {
                var c = parseFloat(cstr);
                if (c < index) {
                    customWidth_new_1[c] = cfg.customWidth[c];
                }
                else if (c === index) {
                    if (direction === "lefttop") {
                        customWidth_new_1[c + count] = cfg.customWidth[c];
                    }
                    else if (direction === "rightbottom") {
                        customWidth_new_1[c] = cfg.customWidth[c];
                    }
                }
                else {
                    customWidth_new_1[c + count] = cfg.customWidth[c];
                }
            });
            cfg.customWidth = customWidth_new_1;
        }
        // 自定义列宽配置变动
        if (cfg.customWidth != null) {
            var customWidth_new_2 = {};
            _.forEach(cfg.customWidth, function (v, cstr) {
                var c = parseFloat(cstr);
                if (c < index) {
                    customWidth_new_2[c] = cfg.customWidth[c];
                }
                else if (c === index) {
                    if (direction === "lefttop") {
                        customWidth_new_2[c + count] = cfg.customWidth[c];
                    }
                    else if (direction === "rightbottom") {
                        customWidth_new_2[c] = cfg.customWidth[c];
                    }
                }
                else {
                    customWidth_new_2[c + count] = cfg.customWidth[c];
                }
            });
            cfg.customWidth = customWidth_new_2;
        }
        // 隐藏列配置变动
        if (cfg.colhidden != null) {
            var colhidden_new_1 = {};
            _.forEach(cfg.colhidden, function (v, cstr) {
                var c = parseFloat(cstr);
                if (c < index) {
                    colhidden_new_1[c] = cfg.colhidden[c];
                }
                else if (c === index) {
                    if (direction === "lefttop") {
                        colhidden_new_1[c + count] = cfg.colhidden[c];
                    }
                    else if (direction === "rightbottom") {
                        colhidden_new_1[c] = cfg.colhidden[c];
                    }
                }
                else {
                    colhidden_new_1[c + count] = cfg.colhidden[c];
                }
            });
            cfg.colhidden = colhidden_new_1;
        }
        // 空列模板
        var col = [];
        var curd = __spreadArray([], d, true);
        for (var r = 0; r < d.length; r += 1) {
            var cell = curd[r][index];
            var templateCell = null;
            if ((cell === null || cell === void 0 ? void 0 : cell.mc) && (direction === "rightbottom" || index !== cell.mc.c)) {
                if (cell.mc.cs) {
                    cell.mc.cs += count;
                }
                templateCell = __assign({}, cell);
                if (!((_u = (_t = curd === null || curd === void 0 ? void 0 : curd[r]) === null || _t === void 0 ? void 0 : _t[index + 1]) === null || _u === void 0 ? void 0 : _u.mc)) {
                    templateCell.mc = undefined;
                }
                delete templateCell.v;
                delete templateCell.m;
                delete templateCell.ps;
                delete templateCell.f;
            }
            col.push(templateCell);
        }
        var cellBorderConfig = [];
        // 边框
        if (cfg.borderInfo && cfg.borderInfo.length > 0) {
            var borderInfo = [];
            for (var i = 0; i < cfg.borderInfo.length; i += 1) {
                var rangeType = cfg.borderInfo[i].rangeType;
                if (rangeType === "range") {
                    var borderRange = cfg.borderInfo[i].range;
                    var emptyRange = [];
                    for (var j = 0; j < borderRange.length; j += 1) {
                        var bd_c1 = borderRange[j].column[0];
                        var bd_c2 = borderRange[j].column[1];
                        if (direction === "lefttop") {
                            if (index <= bd_c1) {
                                bd_c1 += count;
                                bd_c2 += count;
                            }
                            else if (index <= bd_c2) {
                                bd_c2 += count;
                            }
                        }
                        else {
                            if (index < bd_c1) {
                                bd_c1 += count;
                                bd_c2 += count;
                            }
                            else if (index < bd_c2) {
                                bd_c2 += count;
                            }
                        }
                        if (bd_c2 >= bd_c1) {
                            emptyRange.push({
                                row: borderRange[j].row,
                                column: [bd_c1, bd_c2],
                            });
                        }
                    }
                    if (emptyRange.length > 0) {
                        var bd_obj = {
                            rangeType: "range",
                            borderType: cfg.borderInfo[i].borderType,
                            style: cfg.borderInfo[i].style,
                            color: cfg.borderInfo[i].color,
                            range: emptyRange,
                        };
                        borderInfo.push(bd_obj);
                    }
                }
                else if (rangeType === "cell") {
                    var col_index = cfg.borderInfo[i].value.col_index;
                    // 位置相同标识边框相关 先缓存
                    if (col_index === index) {
                        cellBorderConfig.push(JSON.parse(JSON.stringify(cfg.borderInfo[i])));
                    }
                    if (direction === "lefttop") {
                        if (index <= col_index) {
                            col_index += count;
                        }
                    }
                    else {
                        if (index < col_index) {
                            col_index += count;
                        }
                    }
                    cfg.borderInfo[i].value.col_index = col_index;
                    borderInfo.push(cfg.borderInfo[i]);
                }
            }
            cfg.borderInfo = borderInfo;
        }
        // 处理相关的 type 为 cell 类型的边框
        if (cellBorderConfig.length) {
            var _loop_2 = function (i) {
                var cellBorderConfigCopy = _.cloneDeep(cellBorderConfig);
                cellBorderConfigCopy.forEach(function (item) {
                    if (direction === "rightbottom") {
                        // 向右插入时 基于模板列位置直接递增即可
                        item.value.col_index += i + 1;
                    }
                    else if (direction === "lefttop") {
                        // 向左插入时 目标列移动到后面 新增n列到前面 对于新增的列来说 也是递增，不过是从0开始
                        item.value.col_index += i;
                    }
                });
                (_v = cfg.borderInfo) === null || _v === void 0 ? void 0 : _v.push.apply(_v, cellBorderConfigCopy);
            };
            for (var i = 0; i < count; i += 1) {
                _loop_2(i);
            }
        }
        for (var r = 0; r < d.length; r += 1) {
            var row = d[r];
            for (var i = 0; i < count; i += 1) {
                if (direction === "lefttop") {
                    if (index === 0) {
                        row.unshift(col[r]);
                    }
                    else {
                        row.splice(index, 0, col[r]);
                    }
                }
                else {
                    row.splice(index + 1, 0, col[r]);
                }
            }
        }
    }
    // 修改当前sheet页时刷新
    file.data = d;
    file.config = cfg;
    file.calcChain = newCalcChain;
    if (newFilterObj != null) {
        file.filter = newFilterObj.filter;
        file.filter_select = newFilterObj.filter_select;
    }
    file.luckysheet_conditionformat_save = newCFarr;
    file.luckysheet_alternateformat_save = newAFarr;
    file.dataVerification = newDataVerification;
    file.hyperlink = newHyperlink;
    if (file.id === ctx.currentSheetId) {
        ctx.config = cfg;
        // jfrefreshgrid_adRC(
        //   d,
        //   cfg,
        //   "addRC",
        //   {
        //     index,
        //     len: value,
        //     direction,
        //     rc: type1,
        //     restore: false,
        //   },
        //   newCalcChain,
        //   newFilterObj,
        //   newCFarr,
        //   newAFarr,
        //   newFreezen,
        //   newDataVerification,
        //   newHyperlink
        // );
    }
    var range = null;
    if (type === "row") {
        if (direction === "lefttop") {
            range = [
                { row: [index, index + count - 1], column: [0, d[0].length - 1] },
            ];
        }
        else {
            range = [
                { row: [index + 1, index + count], column: [0, d[0].length - 1] },
            ];
        }
        file.row = file.data.length;
    }
    else {
        if (direction === "lefttop") {
            range = [{ row: [0, d.length - 1], column: [index, index + count - 1] }];
        }
        else {
            range = [{ row: [0, d.length - 1], column: [index + 1, index + count] }];
        }
        file.column = (_w = file.data[0]) === null || _w === void 0 ? void 0 : _w.length;
    }
    if (changeSelection) {
        file.luckysheet_select_save = range;
        if (file.id === ctx.currentSheetId) {
            ctx.luckysheet_select_save = range;
            // selectHightlightShow();
        }
    }
    refreshLocalMergeData(merge_new, file);
    // if (type === "row") {
    //   const scrollLeft = $("#luckysheet-cell-main").scrollLeft();
    //   const scrollTop = $("#luckysheet-cell-main").scrollTop();
    //   const winH = $("#luckysheet-cell-main").height();
    //   const winW = $("#luckysheet-cell-main").width();
    //   const row = ctx.visibledatarow[range[0].row[1]];
    //   const row_pre =
    //     range[0].row[0] - 1 === -1 ? 0 : ctx.visibledatarow[range[0].row[0] - 1];
    //   if (row - scrollTop - winH + 20 > 0) {
    //     $("#luckysheet-scrollbar-y").scrollTop(row - winH + 20);
    //   } else if (row_pre - scrollTop - 20 < 0) {
    //     $("#luckysheet-scrollbar-y").scrollTop(row_pre - 20);
    //   }
    //   if (value > 30) {
    //     $("#luckysheet-row-count-show").hide();
    //   }
    // }
}
export function deleteRowCol(ctx, op) {
    var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o, _p, _q;
    var type = op.type;
    var start = op.start, end = op.end, id = op.id;
    id = id || ctx.currentSheetId;
    // if (
    //   type == "row" &&
    //   !checkProtectionAuthorityNormal(sheetId, "deleteRows")
    // ) {
    //   return;
    // }
    // if (
    //   type == "column" &&
    //   !checkProtectionAuthorityNormal(sheetId, "deleteColumns")
    // ) {
    //   return;
    // }
    var curOrder = getSheetIndex(ctx, id);
    if (curOrder == null)
        return;
    var file = ctx.luckysheetfile[curOrder];
    if (!file)
        return;
    var cfg = file.config || {};
    if (type === "row") {
        for (var r = start; r <= end; r += 1) {
            if ((_a = cfg.rowReadOnly) === null || _a === void 0 ? void 0 : _a[r]) {
                throw new Error("readOnly");
            }
        }
    }
    else {
        for (var c = start; c <= end; c += 1) {
            if ((_b = cfg.colReadOnly) === null || _b === void 0 ? void 0 : _b[c]) {
                throw new Error("readOnly");
            }
        }
    }
    var d = file.data;
    if (!d)
        return;
    if (start < 0) {
        start = 0;
    }
    if (end < 0) {
        end = 0;
    }
    if (type === "row") {
        if (start > d.length - 1) {
            start = d.length - 1;
        }
        if (end > d.length - 1) {
            end = d.length - 1;
        }
    }
    else {
        if (start > d[0].length - 1) {
            start = d[0].length - 1;
        }
        if (end > d[0].length - 1) {
            end = d[0].length - 1;
        }
    }
    if (start > end) {
        return;
    }
    var slen = end - start + 1;
    // 合并单元格配置变动
    if (cfg.merge == null) {
        cfg.merge = {};
    }
    var merge_new = {};
    _.forEach(cfg.merge, function (mc) {
        var r = mc.r;
        var c = mc.c;
        var rs = mc.rs;
        var cs = mc.cs;
        if (type === "row") {
            if (r < start) {
                if (r + rs - 1 < start) {
                    merge_new["".concat(r, "_").concat(c)] = { r: r, c: c, rs: rs, cs: cs };
                }
                else if (r + rs - 1 >= start && r + rs - 1 < end) {
                    merge_new["".concat(r, "_").concat(c)] = { r: r, c: c, rs: start - r, cs: cs };
                }
                else if (r + rs - 1 >= end) {
                    merge_new["".concat(r, "_").concat(c)] = { r: r, c: c, rs: rs - slen, cs: cs };
                }
            }
            else if (r >= start && r <= end) {
                if (r + rs - 1 > end) {
                    merge_new["".concat(start, "_").concat(c)] = {
                        r: start,
                        c: c,
                        rs: r + rs - 1 - end,
                        cs: cs,
                    };
                }
            }
            else if (r > end) {
                merge_new["".concat(r - slen, "_").concat(c)] = { r: r - slen, c: c, rs: rs, cs: cs };
            }
        }
        else if (type === "column") {
            if (c < start) {
                if (c + cs - 1 < start) {
                    merge_new["".concat(r, "_").concat(c)] = { r: r, c: c, rs: rs, cs: cs };
                }
                else if (c + cs - 1 >= start && c + cs - 1 < end) {
                    merge_new["".concat(r, "_").concat(c)] = { r: r, c: c, rs: rs, cs: start - c };
                }
                else if (c + cs - 1 >= end) {
                    merge_new["".concat(r, "_").concat(c)] = { r: r, c: c, rs: rs, cs: cs - slen };
                }
            }
            else if (c >= start && c <= end) {
                if (c + cs - 1 > end) {
                    merge_new["".concat(r, "_").concat(start)] = {
                        r: r,
                        c: start,
                        rs: rs,
                        cs: c + cs - 1 - end,
                    };
                }
            }
            else if (c > end) {
                merge_new["".concat(r, "_").concat(c - slen)] = { r: r, c: c - slen, rs: rs, cs: cs };
            }
        }
    });
    cfg.merge = merge_new;
    // 公式配置变动
    var newCalcChain = [];
    for (var SheetIndex = 0; SheetIndex < ctx.luckysheetfile.length; SheetIndex += 1) {
        if (_.isNil(ctx.luckysheetfile[SheetIndex].calcChain) ||
            ctx.luckysheetfile.length === 0) {
            continue;
        }
        var calcChain = ctx.luckysheetfile[SheetIndex].calcChain;
        var data = ctx.luckysheetfile[SheetIndex].data;
        for (var i = 0; i < calcChain.length; i += 1) {
            var calc = _.cloneDeep(calcChain[i]);
            var calc_r = calc.r;
            var calc_c = calc.c;
            var calc_i = calc.id;
            var calc_funcStr = getcellFormula(ctx, calc_r, calc_c, calc_i);
            if (type === "row" && SheetIndex === curOrder) {
                if (calc_r < start || calc_r > end) {
                    var functionStr = "=".concat(functionStrChange(calc_funcStr, "del", "row", null, start, slen));
                    if (((_d = (_c = data[calc_r]) === null || _c === void 0 ? void 0 : _c[calc_c]) === null || _d === void 0 ? void 0 : _d.f) === calc_funcStr) {
                        data[calc_r][calc_c].f = functionStr;
                    }
                    if (calc_r > end) {
                        calc.r = calc_r - slen;
                    }
                    newCalcChain.push(calc);
                }
            }
            else if (type === "row") {
                var functionStr = "=".concat(functionStrChange(calc_funcStr, "del", "row", null, start, slen));
                if (((_f = (_e = data[calc_r]) === null || _e === void 0 ? void 0 : _e[calc_c]) === null || _f === void 0 ? void 0 : _f.f) === calc_funcStr) {
                    data[calc_r][calc_c].f = functionStr;
                }
            }
            else if (type === "column" && SheetIndex === curOrder) {
                if (calc_c < start || calc_c > end) {
                    var functionStr = "=".concat(functionStrChange(calc_funcStr, "del", "col", null, start, slen));
                    if (((_h = (_g = data[calc_r]) === null || _g === void 0 ? void 0 : _g[calc_c]) === null || _h === void 0 ? void 0 : _h.f) === calc_funcStr) {
                        data[calc_r][calc_c].f = functionStr;
                    }
                    if (calc_c > end) {
                        calc.c = calc_c - slen;
                    }
                    newCalcChain.push(calc);
                }
            }
            else if (type === "column") {
                var functionStr = "=".concat(functionStrChange(calc_funcStr, "del", "col", null, start, slen));
                if (((_k = (_j = data[calc_r]) === null || _j === void 0 ? void 0 : _j[calc_c]) === null || _k === void 0 ? void 0 : _k.f) === calc_funcStr) {
                    data[calc_r][calc_c].f = functionStr;
                }
            }
        }
    }
    // 筛选配置变动
    var filter_select = file.filter_select;
    var filter = file.filter;
    var newFilterObj = null;
    if (!_.isEmpty(filter_select) && filter_select != null) {
        newFilterObj = { filter_select: null, filter: null };
        var f_r1_2 = filter_select.row[0];
        var f_r2_2 = filter_select.row[1];
        var f_c1_2 = filter_select.column[0];
        var f_c2_2 = filter_select.column[1];
        if (type === "row") {
            if (f_r1_2 > end) {
                f_r1_2 -= slen;
                f_r2_2 -= slen;
                newFilterObj.filter_select = {
                    row: [f_r1_2, f_r2_2],
                    column: [f_c1_2, f_c2_2],
                };
            }
            else if (f_r1_2 < start) {
                if (f_r2_2 < start) {
                }
                else if (f_r2_2 <= end) {
                    f_r2_2 = start - 1;
                }
                else {
                    f_r2_2 -= slen;
                }
                newFilterObj.filter_select = {
                    row: [f_r1_2, f_r2_2],
                    column: [f_c1_2, f_c2_2],
                };
            }
            if (newFilterObj.filter_select != null && filter != null) {
                _.forEach(filter, function (v, k) {
                    var f_rowhidden = filter[k].rowhidden;
                    var f_rowhidden_new = {};
                    _.forEach(f_rowhidden, function (v1, nstr) {
                        var n = parseFloat(nstr);
                        if (n < start) {
                            f_rowhidden_new[n] = 0;
                        }
                        else if (n > end) {
                            f_rowhidden_new[n - slen] = 0;
                        }
                    });
                    if (!_.isEmpty(f_rowhidden_new)) {
                        if (newFilterObj.filter == null) {
                            newFilterObj.filter = {};
                        }
                        newFilterObj.filter[k] = _.cloneDeep(filter[k]);
                        newFilterObj.filter[k].rowhidden = f_rowhidden_new;
                        newFilterObj.filter[k].str = f_r1_2;
                        newFilterObj.filter[k].edr = f_r2_2;
                    }
                });
            }
        }
        else if (type === "column") {
            if (f_c1_2 > end) {
                f_c1_2 -= slen;
                f_c2_2 -= slen;
                newFilterObj.filter_select = {
                    row: [f_r1_2, f_r2_2],
                    column: [f_c1_2, f_c2_2],
                };
            }
            else if (f_c1_2 < start) {
                if (f_c2_2 < start) {
                }
                else if (f_c2_2 <= end) {
                    f_c2_2 = start - 1;
                }
                else {
                    f_c2_2 -= slen;
                }
                newFilterObj.filter_select = {
                    row: [f_r1_2, f_r2_2],
                    column: [f_c1_2, f_c2_2],
                };
            }
            else {
                if (f_c2_2 > end) {
                    f_c1_2 = start;
                    f_c2_2 -= slen;
                    newFilterObj.filter_select = {
                        row: [f_r1_2, f_r2_2],
                        column: [f_c1_2, f_c2_2],
                    };
                }
            }
            if (newFilterObj.filter_select != null && filter != null) {
                _.forEach(filter, function (v, k) {
                    var f_cindex = filter[k].cindex;
                    if (f_cindex < start) {
                        if (newFilterObj.filter == null) {
                            newFilterObj.filter = {};
                        }
                        newFilterObj.filter[f_cindex - f_c1_2] = _.cloneDeep(filter[k]);
                        newFilterObj.filter[f_cindex - f_c1_2].edc = f_c2_2;
                    }
                    else if (f_cindex > end) {
                        f_cindex -= slen;
                        if (newFilterObj.filter == null) {
                            newFilterObj.filter = {};
                        }
                        newFilterObj.filter[f_cindex - f_c1_2] = _.cloneDeep(filter[k]);
                        newFilterObj.filter[f_cindex - f_c1_2].cindex = f_cindex;
                        newFilterObj.filter[f_cindex - f_c1_2].stc = f_c1_2;
                        newFilterObj.filter[f_cindex - f_c1_2].edc = f_c2_2;
                    }
                });
            }
        }
    }
    if (newFilterObj != null && newFilterObj.filter != null) {
        if (cfg.rowhidden == null) {
            cfg.rowhidden = {};
        }
        _.forEach(newFilterObj.filter, function (v, k) {
            var f_rowhidden = newFilterObj.filter[k].rowhidden;
            _.forEach(f_rowhidden, function (v1, n) {
                cfg.rowhidden[n] = 0;
            });
        });
    }
    // 条件格式配置变动
    var CFarr = file.luckysheet_conditionformat_save;
    var newCFarr = [];
    if (CFarr != null && CFarr.length > 0) {
        for (var i = 0; i < CFarr.length; i += 1) {
            var cf_range = CFarr[i].cellrange;
            var cf_new_range = [];
            for (var j = 0; j < cf_range.length; j += 1) {
                var CFr1 = cf_range[j].row[0];
                var CFr2 = cf_range[j].row[1];
                var CFc1 = cf_range[j].column[0];
                var CFc2 = cf_range[j].column[1];
                if (type === "row") {
                    if (!(CFr1 >= start && CFr2 <= end)) {
                        if (CFr1 > end) {
                            CFr1 -= slen;
                            CFr2 -= slen;
                        }
                        else if (CFr1 < start) {
                            if (CFr2 < start) {
                            }
                            else if (CFr2 <= end) {
                                CFr2 = start - 1;
                            }
                            else {
                                CFr2 -= slen;
                            }
                        }
                        else {
                            if (CFr2 > end) {
                                CFr1 = start;
                                CFr2 -= slen;
                            }
                        }
                        cf_new_range.push({ row: [CFr1, CFr2], column: [CFc1, CFc2] });
                    }
                }
                else if (type === "column") {
                    if (!(CFc1 >= start && CFc2 <= end)) {
                        if (CFc1 > end) {
                            CFc1 -= slen;
                            CFc2 -= slen;
                        }
                        else if (CFc1 < start) {
                            if (CFc2 < start) {
                            }
                            else if (CFc2 <= end) {
                                CFc2 = start - 1;
                            }
                            else {
                                CFc2 -= slen;
                            }
                        }
                        else {
                            if (CFc2 > end) {
                                CFc1 = start;
                                CFc2 -= slen;
                            }
                        }
                        cf_new_range.push({ row: [CFr1, CFr2], column: [CFc1, CFc2] });
                    }
                }
            }
            if (cf_new_range.length > 0) {
                var cf = _.clone(CFarr[i]);
                cf.cellrange = cf_new_range;
                newCFarr.push(cf);
            }
        }
    }
    // 交替颜色配置变动
    var AFarr = file.luckysheet_alternateformat_save;
    var newAFarr = [];
    if (AFarr != null && AFarr.length > 0) {
        for (var i = 0; i < AFarr.length; i += 1) {
            var AFr1 = AFarr[i].cellrange.row[0];
            var AFr2 = AFarr[i].cellrange.row[1];
            var AFc1 = AFarr[i].cellrange.column[0];
            var AFc2 = AFarr[i].cellrange.column[1];
            if (type === "row") {
                if (!(AFr1 >= start && AFr2 <= end)) {
                    var af = _.clone(AFarr[i]);
                    if (AFr1 > end) {
                        AFr1 -= slen;
                        AFr2 -= slen;
                    }
                    else if (AFr1 < start) {
                        if (AFr2 < start) {
                        }
                        else if (AFr2 <= end) {
                            AFr2 = start - 1;
                        }
                        else {
                            AFr2 -= slen;
                        }
                    }
                    else {
                        if (AFr2 > end) {
                            AFr1 = start;
                            AFr2 -= slen;
                        }
                    }
                    af.cellrange = { row: [AFr1, AFr2], column: [AFc1, AFc2] };
                    newAFarr.push(af);
                }
            }
            else if (type === "column") {
                if (!(AFc1 >= start && AFc2 <= end)) {
                    var af = _.clone(AFarr[i]);
                    if (AFc1 > end) {
                        AFc1 -= slen;
                        AFc2 -= slen;
                    }
                    else if (AFc1 < start) {
                        if (AFc2 < start) {
                        }
                        else if (AFc2 <= end) {
                            AFc2 = start - 1;
                        }
                        else {
                            AFc2 -= slen;
                        }
                    }
                    else {
                        if (AFc2 > end) {
                            AFc1 = start;
                            AFc2 -= slen;
                        }
                    }
                    af.cellrange = { row: [AFr1, AFr2], column: [AFc1, AFc2] };
                    newAFarr.push(af);
                }
            }
        }
    }
    // 冻结配置变动
    var frozen = file.frozen;
    if (frozen) {
        if (type === "row" &&
            (frozen.type === "rangeRow" || frozen.type === "rangeBoth")) {
            if (((_m = (_l = frozen.range) === null || _l === void 0 ? void 0 : _l.row_focus) !== null && _m !== void 0 ? _m : -1) >= start) {
                frozen.range.row_focus -=
                    Math.min(end, frozen.range.row_focus) - start + 1;
            }
        }
        if (type === "column" &&
            (frozen.type === "rangeColumn" || frozen.type === "rangeBoth")) {
            if (((_p = (_o = frozen.range) === null || _o === void 0 ? void 0 : _o.column_focus) !== null && _p !== void 0 ? _p : -1) >= start) {
                frozen.range.column_focus -=
                    Math.min(end, frozen.range.column_focus) - start + 1;
            }
        }
    }
    // 数据验证配置变动
    var dataVerification = file.dataVerification;
    var newDataVerification = {};
    if (dataVerification != null) {
        _.forEach(dataVerification, function (v, key) {
            var r = Number(key.split("_")[0]);
            var c = Number(key.split("_")[1]);
            var item = dataVerification[key];
            if (type === "row") {
                if (r < start) {
                    newDataVerification["".concat(r, "_").concat(c)] = item;
                }
                else if (r > end) {
                    newDataVerification["".concat(r - slen, "_").concat(c)] = item;
                }
            }
            else if (type === "column") {
                if (c < start) {
                    newDataVerification["".concat(r, "_").concat(c)] = item;
                }
                else if (c > end) {
                    newDataVerification["".concat(r, "_").concat(c - slen)] = item;
                }
            }
        });
    }
    // 超链接配置变动
    var hyperlink = file.hyperlink;
    var newHyperlink = {};
    if (hyperlink != null) {
        _.forEach(hyperlink, function (v, key) {
            var r = Number(key.split("_")[0]);
            var c = Number(key.split("_")[1]);
            var item = hyperlink[key];
            if (type === "row") {
                if (r < start) {
                    newHyperlink["".concat(r, "_").concat(c)] = item;
                }
                else if (r > end) {
                    newHyperlink["".concat(r - slen, "_").concat(c)] = item;
                }
            }
            else if (type === "column") {
                if (c < start) {
                    newHyperlink["".concat(r, "_").concat(c)] = item;
                }
                else if (c > end) {
                    newHyperlink["".concat(r, "_").concat(c - slen)] = item;
                }
            }
        });
    }
    // 主逻辑
    var type1;
    if (type === "row") {
        type1 = "r";
        // 行高配置变动
        if (cfg.rowlen == null) {
            cfg.rowlen = {};
        }
        var rowlen_new_2 = {};
        var rowReadOnly_new_2 = {};
        _.forEach(cfg.rowlen, function (v, rstr) {
            var r = parseFloat(rstr);
            if (r < start) {
                rowlen_new_2[r] = cfg.rowlen[r];
            }
            else if (r > end) {
                rowlen_new_2[r - slen] = cfg.rowlen[r];
            }
        });
        _.forEach(cfg.rowReadOnly, function (v, rstr) {
            var r = parseFloat(rstr);
            if (r < start) {
                rowReadOnly_new_2[r] = cfg.rowReadOnly[r];
            }
            else if (r > end) {
                rowReadOnly_new_2[r - slen] = cfg.rowReadOnly[r];
            }
        });
        cfg.rowlen = rowlen_new_2;
        cfg.rowReadOnly = rowReadOnly_new_2;
        // 隐藏行配置变动
        if (cfg.rowhidden == null) {
            cfg.rowhidden = {};
        }
        var rowhidden_new_2 = {};
        _.forEach(cfg.rowhidden, function (v, rstr) {
            var r = parseFloat(rstr);
            if (r < start) {
                rowhidden_new_2[r] = cfg.rowhidden[r];
            }
            else if (r > end) {
                rowhidden_new_2[r - slen] = cfg.rowhidden[r];
            }
        });
        // 自定义行高配置变动
        if (cfg.customHeight == null) {
            cfg.customHeight = {};
            var customHeight_new_3 = {};
            _.forEach(cfg.customHeight, function (v, rstr) {
                var r = parseFloat(rstr);
                if (r < start) {
                    customHeight_new_3[r] = cfg.customHeight[r];
                }
                else if (r > end) {
                    customHeight_new_3[r - slen] = cfg.customHeight[r];
                }
            });
            cfg.customHeight = customHeight_new_3;
        }
        // 自定义行高配置变动
        if (cfg.customHeight == null) {
            cfg.customHeight = {};
            var customHeight_new_4 = {};
            _.forEach(cfg.customHeight, function (v, rstr) {
                var r = parseFloat(rstr);
                if (r < start) {
                    customHeight_new_4[r] = cfg.customHeight[r];
                }
                else if (r > end) {
                    customHeight_new_4[r - slen] = cfg.customHeight[r];
                }
            });
            cfg.customHeight = customHeight_new_4;
        }
        cfg.rowhidden = rowhidden_new_2;
        // 边框配置变动
        if (cfg.borderInfo && cfg.borderInfo.length > 0) {
            var borderInfo = [];
            for (var i = 0; i < cfg.borderInfo.length; i += 1) {
                var rangeType = cfg.borderInfo[i].rangeType;
                if (rangeType === "range") {
                    var borderRange = cfg.borderInfo[i].range;
                    var emptyRange = [];
                    for (var j = 0; j < borderRange.length; j += 1) {
                        var bd_r1 = borderRange[j].row[0];
                        var bd_r2 = borderRange[j].row[1];
                        for (var r = start; r <= end; r += 1) {
                            if (r < borderRange[j].row[0]) {
                                bd_r1 -= 1;
                                bd_r2 -= 1;
                            }
                            else if (r <= borderRange[j].row[1]) {
                                bd_r2 -= 1;
                            }
                        }
                        if (bd_r2 >= bd_r1) {
                            emptyRange.push({
                                row: [bd_r1, bd_r2],
                                column: borderRange[j].column,
                            });
                        }
                    }
                    if (emptyRange.length > 0) {
                        var bd_obj = {
                            rangeType: "range",
                            borderType: cfg.borderInfo[i].borderType,
                            style: cfg.borderInfo[i].style,
                            color: cfg.borderInfo[i].color,
                            range: emptyRange,
                        };
                        borderInfo.push(bd_obj);
                    }
                }
                else if (rangeType === "cell") {
                    var row_index = cfg.borderInfo[i].value.row_index;
                    if (row_index < start) {
                        borderInfo.push(cfg.borderInfo[i]);
                    }
                    else if (row_index > end) {
                        cfg.borderInfo[i].value.row_index = row_index - (end - start + 1);
                        borderInfo.push(cfg.borderInfo[i]);
                    }
                }
            }
            cfg.borderInfo = borderInfo;
        }
        // 备注：该处理方式会在删除多行的时候会存在bug
        // 说明：删除多行后，会把同一个row空数组(引用类型)添加成为data多行的数据源，导致设置这些行数据时产生错误。
        // 空白行模板
        // let row = [];
        // for (let c = 0; c < d[0].length; c++) {
        //     row.push(null);
        // }
        // //删除选中行
        // d.splice(st, slen);
        // //删除多少行，增加多少行空白行
        // for (let r = 0; r < slen; r++) {
        //     d.push(row);
        // }
        // 删除选中行
        d.splice(start, slen);
        // 删除行后，调整行数
        file.row = d.length;
    }
    else {
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        type1 = "c";
        // 列宽配置变动
        if (cfg.columnlen == null) {
            cfg.columnlen = {};
        }
        var columnlen_new_2 = {};
        var columnReadOnly_new_2 = {};
        _.forEach(cfg.columnlen, function (v, cstr) {
            var c = parseFloat(cstr);
            if (c < start) {
                columnlen_new_2[c] = cfg.columnlen[c];
            }
            else if (c > end) {
                columnlen_new_2[c - slen] = cfg.columnlen[c];
            }
        });
        _.forEach(cfg.colReadOnly, function (v, cstr) {
            var c = parseFloat(cstr);
            if (c < start) {
                columnReadOnly_new_2[c] = cfg.colReadOnly[c];
            }
            else if (c > end) {
                columnReadOnly_new_2[c - slen] = cfg.colReadOnly[c];
            }
        });
        cfg.columnlen = columnlen_new_2;
        cfg.colReadOnly = columnReadOnly_new_2;
        // 自定义列宽配置变动
        if (cfg.customWidth == null) {
            cfg.customWidth = {};
            var customWidth_new_3 = {};
            _.forEach(cfg.customWidth, function (v, rstr) {
                var r = parseFloat(rstr);
                if (r < start) {
                    customWidth_new_3[r] = cfg.customWidth[r];
                }
                else if (r > end) {
                    customWidth_new_3[r - slen] = cfg.customWidth[r];
                }
            });
            cfg.customWidth = customWidth_new_3;
        }
        cfg.colReadOnly = columnReadOnly_new_2;
        // 隐藏列配置变动
        if (cfg.colhidden == null) {
            cfg.colhidden = {};
        }
        var colhidden_new_2 = {};
        _.forEach(cfg.colhidden, function (v, cstr) {
            var c = parseFloat(cstr);
            if (c < start) {
                colhidden_new_2[c] = cfg.colhidden[c];
            }
            else if (c > end) {
                colhidden_new_2[c - slen] = cfg.colhidden[c];
            }
        });
        cfg.colhidden = colhidden_new_2;
        // 边框配置变动
        if (cfg.borderInfo && cfg.borderInfo.length > 0) {
            var borderInfo = [];
            for (var i = 0; i < cfg.borderInfo.length; i += 1) {
                var rangeType = cfg.borderInfo[i].rangeType;
                if (rangeType === "range") {
                    var borderRange = cfg.borderInfo[i].range;
                    var emptyRange = [];
                    for (var j = 0; j < borderRange.length; j += 1) {
                        var bd_c1 = borderRange[j].column[0];
                        var bd_c2 = borderRange[j].column[1];
                        for (var c = start; c <= end; c += 1) {
                            if (c < borderRange[j].column[0]) {
                                bd_c1 -= 1;
                                bd_c2 -= 1;
                            }
                            else if (c <= borderRange[j].column[1]) {
                                bd_c2 -= 1;
                            }
                        }
                        if (bd_c2 >= bd_c1) {
                            emptyRange.push({
                                row: borderRange[j].row,
                                column: [bd_c1, bd_c2],
                            });
                        }
                    }
                    if (emptyRange.length > 0) {
                        var bd_obj = {
                            rangeType: "range",
                            borderType: cfg.borderInfo[i].borderType,
                            style: cfg.borderInfo[i].style,
                            color: cfg.borderInfo[i].color,
                            range: emptyRange,
                        };
                        borderInfo.push(bd_obj);
                    }
                }
                else if (rangeType === "cell") {
                    var col_index = cfg.borderInfo[i].value.col_index;
                    if (col_index < start) {
                        borderInfo.push(cfg.borderInfo[i]);
                    }
                    else if (col_index > end) {
                        cfg.borderInfo[i].value.col_index = col_index - (end - start + 1);
                        borderInfo.push(cfg.borderInfo[i]);
                    }
                }
            }
            cfg.borderInfo = borderInfo;
        }
        for (var r = 0; r < d.length; r += 1) {
            // 删除选中列
            d[r].splice(start, slen);
        }
        // 删除列后，调整列数
        file.column = (_q = d[0]) === null || _q === void 0 ? void 0 : _q.length;
    }
    // 选中元素被删取消选区
    ctx.luckysheet_select_save = undefined;
    // 修改当前sheet页时刷新
    file.data = d;
    file.config = cfg;
    file.calcChain = newCalcChain;
    if (newFilterObj != null) {
        file.filter = newFilterObj.filter;
        file.filter_select = newFilterObj.filter_select;
    }
    file.luckysheet_conditionformat_save = newCFarr;
    file.luckysheet_alternateformat_save = newAFarr;
    file.dataVerification = newDataVerification;
    file.hyperlink = newHyperlink;
    refreshLocalMergeData(merge_new, file);
    if (file.id === ctx.currentSheetId) {
        ctx.config = cfg;
        // jfrefreshgrid_adRC(
        //   d,
        //   cfg,
        //   "delRC",
        //   { index: st, len: ed - st + 1, rc: type1 },
        //   newCalcChain,
        //   newFilterObj,
        //   newCFarr,
        //   newAFarr,
        //   newFreezen,
        //   newDataVerification,
        //   newHyperlink
        // );
    }
    else {
    }
}
// 计算表格行高数组
export function computeRowlenArr(ctx, rowHeight, cfg) {
    var rowlenArr = [];
    var rh_height = 0;
    for (var i = 0; i < rowHeight; i += 1) {
        var rowlen = ctx.defaultrowlen;
        if (cfg.rowlen != null && cfg.rowlen[i] != null) {
            rowlen = cfg.rowlen[i];
        }
        if (cfg.rowhidden != null && cfg.rowhidden[i] != null) {
            rowlen = cfg.rowhidden[i];
            rowlenArr.push(rh_height);
            continue;
        }
        else {
            rh_height += rowlen + 1;
        }
        rowlenArr.push(rh_height); // 行的临时长度分布
    }
    return rowlenArr;
}
// 隐藏选中行列
export function hideSelected(ctx, type) {
    var _a, _b;
    if (!ctx.luckysheet_select_save || ctx.luckysheet_select_save.length > 1)
        return "noMulti";
    var index = getSheetIndex(ctx, ctx.currentSheetId);
    // 隐藏行
    if (type === "row") {
        /* TODO: 工作表保护判断
        if (
          !checkProtectionAuthorityNormal(Store.currentSheetIndex, "formatRows")
        ) {
          return ;
        } */
        var rowhidden = (_a = ctx.config.rowhidden) !== null && _a !== void 0 ? _a : {};
        var r1 = ctx.luckysheet_select_save[0].row[0];
        var r2 = ctx.luckysheet_select_save[0].row[1];
        var rowhiddenNumber_1 = r2;
        for (var r = r1; r <= r2; r += 1) {
            rowhidden[r] = 0;
        }
        /* 保存撤销,luck中保存撤销用以下方式实现，而在本项目中不需要另外处理
          if(Store.clearjfundo){
            let redo = {};
            redo["type"] = "showHidRows";
            redo["sheetIndex"] = Store.currentSheetIndex;
            redo["config"] = $.extend(true, {}, Store.config);
            redo["curconfig"] = cfg;
    
            Store.jfundo.length  = 0;
            Store.jfredo.push(redo);
        } */
        ctx.config.rowhidden = rowhidden;
        var rowLen = ctx.luckysheetfile[index].data.length;
        /**
         * 计算要隐藏的行是否是最后一列
         * 符合最后一列的条件：要隐藏的index===表格的长度-1 或者
         * 记录隐藏数组里面的数-1===要隐藏的index
         */
        var isEndRow = rowLen - 1 === rowhiddenNumber_1 ||
            Object.keys(rowhidden).findIndex(function (o) { return parseInt(o, 10) - 1 === rowhiddenNumber_1; }) >= 0;
        if (isEndRow) {
            ctx.luckysheet_select_save[0].row[0] -= 1;
            ctx.luckysheet_select_save[0].row[1] -= 1;
        }
        else {
            ctx.luckysheet_select_save[0].row[0] += 1;
            ctx.luckysheet_select_save[0].row[1] += 1;
        }
    }
    else if (type === "column") {
        // 隐藏列
        var colhidden = (_b = ctx.config.colhidden) !== null && _b !== void 0 ? _b : {};
        var c1 = ctx.luckysheet_select_save[0].column[0];
        var c2 = ctx.luckysheet_select_save[0].column[1];
        var colhiddenNumber_1 = c2;
        for (var c = c1; c <= c2; c += 1) {
            colhidden[c] = 0;
        }
        ctx.config.colhidden = colhidden;
        var columnLen = ctx.luckysheetfile[index].data[0].length;
        // 计算要隐藏的列是否是最后一列
        var isEndColumn = columnLen - 1 === colhiddenNumber_1 ||
            Object.keys(colhidden).findIndex(function (o) { return parseInt(o, 10) - 1 === colhiddenNumber_1; }) >= 0;
        if (isEndColumn) {
            ctx.luckysheet_select_save[0].column[0] -= 1;
            ctx.luckysheet_select_save[0].column[1] -= 1;
        }
        else {
            ctx.luckysheet_select_save[0].column[0] += 1;
            ctx.luckysheet_select_save[0].column[1] += 1;
        }
    }
    ctx.luckysheetfile[index].config = ctx.config;
    return "";
}
// 取消隐藏选中行列
export function showSelected(ctx, type) {
    var _a, _b;
    if (!ctx.luckysheet_select_save || ctx.luckysheet_select_save.length > 1)
        return "noMulti";
    var index = getSheetIndex(ctx, ctx.currentSheetId);
    // 取消隐藏行
    if (type === "row") {
        var rowhidden = (_a = ctx.config.rowhidden) !== null && _a !== void 0 ? _a : {};
        var r1 = ctx.luckysheet_select_save[0].row[0];
        var r2 = ctx.luckysheet_select_save[0].row[1];
        for (var r = r1; r <= r2; r += 1) {
            delete rowhidden[r];
        }
        ctx.config.rowhidden = rowhidden;
    }
    else if (type === "column") {
        // 取消隐藏列
        var colhidden = (_b = ctx.config.colhidden) !== null && _b !== void 0 ? _b : {};
        var c1 = ctx.luckysheet_select_save[0].column[0];
        var c2 = ctx.luckysheet_select_save[0].column[1];
        for (var c = c1; c <= c2; c += 1) {
            delete colhidden[c];
        }
        ctx.config.colhidden = colhidden;
    }
    ctx.luckysheetfile[index].config = ctx.config;
    return "";
}
// 判断当前选区是不是隐藏行列
export function isShowHidenCR(ctx) {
    var _a, _b, _c, _d;
    if (!ctx.luckysheet_select_save ||
        (!ctx.config.colhidden && !ctx.config.rowhidden))
        return false;
    // 如果当先选区处在隐藏行列的时候则不可编辑
    if (!!ctx.config.colhidden && _.size(ctx.config.colhidden) >= 1) {
        var ctxColumn_1 = (_b = (_a = ctx.luckysheet_select_save[0]) === null || _a === void 0 ? void 0 : _a.column) === null || _b === void 0 ? void 0 : _b[0];
        var isHidenColumn = Object.keys(ctx.config.colhidden).findIndex(function (o) {
            return ctxColumn_1 === parseInt(o, 10);
        }) >= 0;
        if (isHidenColumn) {
            return true;
        }
    }
    if (!!ctx.config.rowhidden && _.size(ctx.config.rowhidden) >= 1) {
        var ctxRow_1 = (_d = (_c = ctx.luckysheet_select_save[0]) === null || _c === void 0 ? void 0 : _c.row) === null || _d === void 0 ? void 0 : _d[0];
        var isHidenRow = Object.keys(ctx.config.rowhidden).findIndex(function (o) {
            return ctxRow_1 === parseInt(o, 10);
        }) >= 0;
        if (isHidenRow) {
            return true;
        }
    }
    return false;
}
// 计算键盘选区中要经过的隐藏行列的个数
export function hideCRCount(ctx, type) {
    var _a, _b;
    var count = 1;
    if (!ctx.luckysheet_select_save)
        return 0;
    var section = ctx.luckysheet_select_save[0];
    var rowhidden = (_a = ctx.config.rowhidden) !== null && _a !== void 0 ? _a : {};
    var colhidden = (_b = ctx.config.colhidden) !== null && _b !== void 0 ? _b : {};
    if (type === "ArrowUp" || type === "ArrowDown") {
        var rowArr = Object.keys(rowhidden);
        if (type === "ArrowUp") {
            var row = section.row[0] - 1;
            var rowIndex = rowArr.indexOf(row.toString());
            for (var i = rowIndex; i >= 0; i -= 1) {
                if (parseInt(rowArr[i], 10) === row) {
                    count += 1;
                    row -= 1;
                }
                else {
                    return count;
                }
            }
        }
        else {
            var row = section.row[0] + 1;
            var rowIndex = rowArr.indexOf("".concat(row));
            for (var i = rowIndex; i < rowArr.length; i += 1) {
                if (parseInt(rowArr[i], 10) === row) {
                    count += 1;
                    row += 1;
                }
                else {
                    return count;
                }
            }
        }
    }
    else if (type === "ArrowLeft" || type === "ArrowRight") {
        var columnArr = Object.keys(colhidden);
        if (type === "ArrowLeft") {
            var column = section.column[0] - 1;
            var columnIndex = columnArr.indexOf(column.toString());
            for (var i = columnIndex; i >= 0; i -= 1) {
                if (parseInt(columnArr[i], 10) === column) {
                    count += 1;
                    column -= 1;
                }
                else {
                    return count;
                }
            }
        }
        else {
            var column = section.column[0] + 1;
            var columnIndex = columnArr.indexOf("".concat(column));
            for (var i = columnIndex; i < columnArr.length; i += 1) {
                if (parseInt(columnArr[i], 10) === column) {
                    count += 1;
                    column += 1;
                }
                else {
                    return count;
                }
            }
        }
    }
    return count;
}
