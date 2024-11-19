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
import { getSheetIndex } from ".";
import { getFlowdata } from "../context";
var addtionalMergeOps = function (ops, id) {
    var merge_new = {};
    ops.some(function (op) {
        if (op.op === "replace" &&
            op.path[0] === "config" &&
            op.path[1] === "merge") {
            merge_new = op.value;
            return true;
        }
        return false;
    });
    var new_ops = [];
    Object.entries(merge_new).forEach(function (_a) {
        var v = _a[1];
        var _b = v, r = _b.r, c = _b.c, rs = _b.rs, cs = _b.cs;
        var headerOp = {
            op: "replace",
            path: ["data", r, c, "mc"],
            id: id,
            value: v,
        };
        for (var i = r; i < r + rs; i += 1) {
            for (var j = c; j < c + cs; j += 1) {
                new_ops.push({
                    op: "replace",
                    path: ["data", i, j, "mc"],
                    id: id,
                    value: { r: r, c: c },
                });
            }
        }
        new_ops.push(headerOp);
    });
    return new_ops;
};
function additionalCellOps(ctx, insertRowColOp) {
    var id = insertRowColOp.id, index = insertRowColOp.index, direction = insertRowColOp.direction, count = insertRowColOp.count, type = insertRowColOp.type;
    var d = getFlowdata(ctx, id);
    var startIndex = index + (direction === "rightbottom" ? 1 : 0);
    if (d == null) {
        return [];
    }
    var cellOps = [];
    if (type === "row") {
        for (var i = 0; i < d[startIndex].length; i += 1) {
            var cell = d[startIndex][i];
            if (cell != null) {
                for (var j = 0; j < count; j += 1) {
                    cellOps.push({
                        op: "replace",
                        id: id,
                        path: ["data", startIndex + j, i],
                        value: cell,
                    });
                }
            }
        }
    }
    else {
        for (var i = 0; i < d.length; i += 1) {
            var cell = d[i][startIndex];
            if (cell != null) {
                for (var j = 0; j < count; j += 1) {
                    cellOps.push({
                        op: "replace",
                        id: id,
                        path: ["data", i, startIndex + j],
                        value: cell,
                    });
                }
            }
        }
    }
    return cellOps;
}
export function filterPatch(patches) {
    return _.filter(patches, function (p) {
        return p.path[0] === "luckysheetfile" && p.path[2] !== "luckysheet_select_save";
    });
}
export function extractFormulaCellOps(ops) {
    // ops are ensured to be cell data ops
    var formulaOps = [];
    ops.forEach(function (op) {
        var _a, _b;
        if (op.op === "remove")
            return;
        if (op.path.length === 2 && Array.isArray(op.value)) {
            // entire row op
            for (var i = 0; i < op.value.length; i += 1) {
                if ((_a = op.value[i]) === null || _a === void 0 ? void 0 : _a.f) {
                    formulaOps.push({
                        op: "replace",
                        id: op.id,
                        path: __spreadArray(__spreadArray([], op.path, true), [i], false),
                        value: op.value[i],
                    });
                }
            }
        }
        else if (op.path.length === 3 && ((_b = op.value) === null || _b === void 0 ? void 0 : _b.f)) {
            formulaOps.push(op);
        }
        else if (op.path.length === 4 && op.path[3] === "f") {
            formulaOps.push(op);
        }
    });
    return formulaOps;
}
export function patchToOp(ctx, patches, options, undo) {
    var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o;
    if (undo === void 0) { undo = false; }
    var ops = patches.map(function (p) {
        var op = {
            op: p.op,
            value: p.value,
            path: p.path,
        };
        if (p.path[0] === "luckysheetfile" && _.isNumber(p.path[1])) {
            var id = ctx.luckysheetfile[p.path[1]].id;
            op.id = id;
            op.path = p.path.slice(2);
            if (_.isEqual(op.path, ["calcChain", "length"])) {
                op.path = ["calcChain"];
                op.value = ctx.luckysheetfile[p.path[1]].calcChain;
            }
        }
        return op;
    });
    _.every(ops, function (p) {
        var _a;
        if (p.op === "replace" &&
            !_.isNil((_a = p.value) === null || _a === void 0 ? void 0 : _a.hl) &&
            p.path.length === 3 &&
            p.path[0] === "data") {
            var index = getSheetIndex(ctx, p.id);
            ops.push({
                id: p.id,
                op: "replace",
                path: ["hyperlink", "".concat(p.path[1], "_").concat(p.path[2])],
                value: ctx.luckysheetfile[index].hyperlink["".concat(p.value.hl.r, "_").concat(p.value.hl.c)],
            });
        }
    });
    if (options === null || options === void 0 ? void 0 : options.insertRowColOp) {
        var _p = _.partition(ops, function (p) { return p.path[0] !== "data"; }), nonDataOps = _p[0], dataOps = _p[1];
        // find out formula cells as their formula range may be changed
        var formulaOps = extractFormulaCellOps(dataOps);
        ops = nonDataOps;
        ops.push({
            op: "insertRowCol",
            id: options.insertRowColOp.id,
            path: [],
            value: options.insertRowColOp,
        });
        ops = __spreadArray(__spreadArray([], ops, true), formulaOps, true);
        var mergeOps = addtionalMergeOps(ops, ctx.currentSheetId);
        ops = __spreadArray(__spreadArray([], ops, true), mergeOps, true);
        if (options === null || options === void 0 ? void 0 : options.restoreDeletedCells) {
            // undoing deleted row/col, find out cells to restore
            var restoreCellsOps = [];
            var flowdata = getFlowdata(ctx);
            if (flowdata) {
                var rowlen = flowdata.length;
                var collen = flowdata[0].length;
                for (var i = 0; i < rowlen; i += 1) {
                    for (var j = 0; j < collen; j += 1) {
                        var cell = flowdata[i][j];
                        if (!cell)
                            continue;
                        if ((options.insertRowColOp.type === "row" &&
                            i >= options.insertRowColOp.index &&
                            i <
                                options.insertRowColOp.index +
                                    options.insertRowColOp.count) ||
                            (options.insertRowColOp.type === "column" &&
                                j >= options.insertRowColOp.index &&
                                j < options.insertRowColOp.index + options.insertRowColOp.count)) {
                            restoreCellsOps.push({
                                op: "replace",
                                path: ["data", i, j],
                                id: ctx.currentSheetId,
                                value: cell,
                            });
                        }
                    }
                }
            }
            ops = __spreadArray(__spreadArray([], ops, true), restoreCellsOps, true);
        }
        else {
            var cellOps = additionalCellOps(ctx, options.insertRowColOp);
            ops = __spreadArray(__spreadArray([], ops, true), cellOps, true);
        }
    }
    else if (options === null || options === void 0 ? void 0 : options.deleteRowColOp) {
        var _q = _.partition(ops, function (p) { return p.path[0] !== "data"; }), nonDataOps = _q[0], dataOps = _q[1];
        // find out formula cells as their formula range may be changed
        var formulaOps = extractFormulaCellOps(dataOps);
        ops = nonDataOps;
        ops.push({
            op: "deleteRowCol",
            id: options.deleteRowColOp.id,
            path: [],
            value: options.deleteRowColOp,
        });
        ops = __spreadArray(__spreadArray([], ops, true), formulaOps, true);
        var mergeOps = addtionalMergeOps(ops, ctx.currentSheetId);
        ops = __spreadArray(__spreadArray([], ops, true), mergeOps, true);
    }
    else if (options === null || options === void 0 ? void 0 : options.addSheetOp) {
        var _r = _.partition(ops, function (op) { return op.path.length === 0 && op.op === "add"; }), addSheetOps = _r[0], otherOps = _r[1];
        options.id = options.addSheet.id;
        if (undo) {
            // 撤消增表
            var index = getSheetIndex(ctx, options.addSheet.id);
            var order_1 = (_b = (_a = options.addSheet) === null || _a === void 0 ? void 0 : _a.value) === null || _b === void 0 ? void 0 : _b.order;
            ops = otherOps;
            ops.push({
                op: "deleteSheet",
                id: (_c = options.addSheet) === null || _c === void 0 ? void 0 : _c.id,
                path: [],
                value: options.addSheet,
            });
            if (index !== ctx.luckysheetfile.length) {
                var sheetsRight = ctx.luckysheetfile.filter(function (sheet) { return (sheet === null || sheet === void 0 ? void 0 : sheet.order) >= order_1; });
                _.forEach(sheetsRight, function (sheet) {
                    ops.push({
                        id: sheet.id,
                        op: "replace",
                        path: ["order"],
                        value: (sheet === null || sheet === void 0 ? void 0 : sheet.order) - 1,
                    });
                });
            }
        }
        else {
            // 正常增表
            ops = otherOps;
            ops.push({
                op: "addSheet",
                id: (_d = options.addSheet) === null || _d === void 0 ? void 0 : _d.id,
                path: [],
                value: (_e = addSheetOps[0]) === null || _e === void 0 ? void 0 : _e.value,
            });
        }
    }
    else if (options === null || options === void 0 ? void 0 : options.deleteSheetOp) {
        options.id = options.deleteSheetOp.id;
        if (undo) {
            // 撤销删表
            ops = [
                {
                    op: "addSheet",
                    id: options.deleteSheetOp.id,
                    path: [],
                    value: (_f = options.deletedSheet) === null || _f === void 0 ? void 0 : _f.value,
                },
                {
                    id: options.deleteSheetOp.id,
                    op: "replace",
                    path: ["name"],
                    value: (_h = (_g = options.deletedSheet) === null || _g === void 0 ? void 0 : _g.value) === null || _h === void 0 ? void 0 : _h.name,
                },
            ];
            var order_2 = (_k = (_j = options.deletedSheet) === null || _j === void 0 ? void 0 : _j.value) === null || _k === void 0 ? void 0 : _k.order;
            var sheetsRight = ctx.luckysheetfile.filter(function (sheet) {
                var _a;
                return (sheet === null || sheet === void 0 ? void 0 : sheet.order) >= order_2 &&
                    sheet.id !== ((_a = options.deleteSheetOp) === null || _a === void 0 ? void 0 : _a.id);
            });
            _.forEach(sheetsRight, function (sheet) {
                ops.push({
                    id: sheet.id,
                    op: "replace",
                    path: ["order"],
                    value: sheet === null || sheet === void 0 ? void 0 : sheet.order,
                });
            });
        }
        else {
            // 正常删表
            ops = [
                {
                    op: "deleteSheet",
                    id: options.deleteSheetOp.id,
                    path: [],
                    value: options.deletedSheet,
                },
            ];
            var order_3 = (_m = (_l = options.deletedSheet) === null || _l === void 0 ? void 0 : _l.value) === null || _m === void 0 ? void 0 : _m.order;
            if (((_o = options.deletedSheet) === null || _o === void 0 ? void 0 : _o.order) !== ctx.luckysheetfile.length) {
                var sheetsRight = ctx.luckysheetfile.filter(function (sheet) { return (sheet === null || sheet === void 0 ? void 0 : sheet.order) >= order_3; });
                _.forEach(sheetsRight, function (sheet) {
                    ops.push({
                        id: sheet.id,
                        op: "replace",
                        path: ["order"],
                        value: sheet === null || sheet === void 0 ? void 0 : sheet.order,
                    });
                });
            }
        }
    }
    return ops;
}
export function opToPatch(ctx, ops) {
    var _a = _.partition(ops, function (op) { return op.op === "add" || op.op === "remove" || op.op === "replace"; }), normalOps = _a[0], specialOps = _a[1];
    var additionalPatches = [];
    var patches = normalOps.map(function (op) {
        var patch = {
            op: op.op,
            value: op.value,
            path: op.path,
        };
        if (op.id) {
            var i = getSheetIndex(ctx, op.id);
            if (i != null) {
                patch.path = __spreadArray(["luckysheetfile", i], op.path, true);
            }
            else {
                // throw new Error(`sheet id: ${op.id} not found`);
            }
            if (op.path[0] === "images" && op.id === ctx.currentSheetId) {
                additionalPatches.push(__assign(__assign({}, patch), { path: ["insertedImgs"] }));
            }
        }
        return patch;
    });
    return [patches.concat(additionalPatches), specialOps];
}
export function inverseRowColOptions(options) {
    if (!options)
        return options;
    if (options.insertRowColOp) {
        var index = options.insertRowColOp.index;
        if (options.insertRowColOp.direction === "rightbottom") {
            index += 1;
        }
        return {
            deleteRowColOp: {
                type: options.insertRowColOp.type,
                id: options.insertRowColOp.id,
                start: index,
                end: index + options.insertRowColOp.count - 1,
            },
        };
    }
    if (options.deleteRowColOp) {
        return {
            insertRowColOp: {
                type: options.deleteRowColOp.type,
                id: options.deleteRowColOp.id,
                index: options.deleteRowColOp.start,
                count: options.deleteRowColOp.end - options.deleteRowColOp.start + 1,
                direction: "lefttop",
            },
        };
    }
    return options;
}
