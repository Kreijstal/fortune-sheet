import _ from "lodash";
import { getSheetIndex } from "../utils";
import { isInlineStringCT } from "./inline-string";
export function mergeCells(ctx, sheetId, ranges, type) {
    // if (!checkIsAllowEdit()) {
    //   tooltip.info("", locale().pivotTable.errorNotAllowEdit);
    //   return;
    // }
    var idx = getSheetIndex(ctx, sheetId);
    if (idx == null)
        return;
    var sheet = ctx.luckysheetfile[idx];
    var cfg = sheet.config || {};
    if (cfg.merge == null) {
        cfg.merge = {};
    }
    var d = sheet.data;
    // if (!checkProtectionNotEnable(ctx.currentSheetId)) {
    //   return;
    // }
    if (type === "merge-cancel") {
        for (var i = 0; i < ranges.length; i += 1) {
            var range = ranges[i];
            var r1 = range.row[0];
            var r2 = range.row[1];
            var c1 = range.column[0];
            var c2 = range.column[1];
            if (r1 === r2 && c1 === c2) {
                continue;
            }
            var fv = {};
            for (var r = r1; r <= r2; r += 1) {
                for (var c = c1; c <= c2; c += 1) {
                    var cell = d[r][c];
                    if (cell != null && cell.mc != null) {
                        var mc_r = cell.mc.r;
                        var mc_c = cell.mc.c;
                        if ("rs" in cell.mc) {
                            delete cell.mc;
                            delete cfg.merge["".concat(mc_r, "_").concat(mc_c)];
                            fv["".concat(mc_r, "_").concat(mc_c)] = _.cloneDeep(cell) || {};
                        }
                        else {
                            // let cell_clone = fv[mc_r + "_" + mc_c];
                            var cell_clone = _.cloneDeep(fv["".concat(mc_r, "_").concat(mc_c)]);
                            delete cell_clone.v;
                            delete cell_clone.m;
                            delete cell_clone.ct;
                            delete cell_clone.f;
                            delete cell_clone.spl;
                            d[r][c] = cell_clone;
                        }
                    }
                }
            }
        }
    }
    else {
        var isHasMc = false; // 选区是否含有 合并的单元格
        for (var i = 0; i < ranges.length; i += 1) {
            var range = ranges[i];
            var r1 = range.row[0];
            var r2 = range.row[1];
            var c1 = range.column[0];
            var c2 = range.column[1];
            for (var r = r1; r <= r2; r += 1) {
                for (var c = c1; c <= c2; c += 1) {
                    var cell = d[r][c];
                    if (cell === null || cell === void 0 ? void 0 : cell.mc) {
                        isHasMc = true;
                        break;
                    }
                }
            }
        }
        if (isHasMc) {
            // 选区有合并单元格（选区都执行 取消合并）
            for (var i = 0; i < ranges.length; i += 1) {
                var range = ranges[i];
                var r1 = range.row[0];
                var r2 = range.row[1];
                var c1 = range.column[0];
                var c2 = range.column[1];
                if (r1 === r2 && c1 === c2) {
                    continue;
                }
                var fv = {};
                for (var r = r1; r <= r2; r += 1) {
                    for (var c = c1; c <= c2; c += 1) {
                        var cell = d[r][c];
                        if (cell != null && cell.mc != null) {
                            var mc_r = cell.mc.r;
                            var mc_c = cell.mc.c;
                            if ("rs" in cell.mc) {
                                delete cell.mc;
                                delete cfg.merge["".concat(mc_r, "_").concat(mc_c)];
                                fv["".concat(mc_r, "_").concat(mc_c)] = _.cloneDeep(cell) || {};
                            }
                            else {
                                // let cell_clone = fv[mc_r + "_" + mc_c];
                                var cell_clone = _.cloneDeep(fv["".concat(mc_r, "_").concat(mc_c)]);
                                delete cell_clone.v;
                                delete cell_clone.m;
                                delete cell_clone.ct;
                                delete cell_clone.f;
                                delete cell_clone.spl;
                                d[r][c] = cell_clone;
                            }
                        }
                    }
                }
            }
        }
        else {
            for (var i = 0; i < ranges.length; i += 1) {
                var range = ranges[i];
                var r1 = range.row[0];
                var r2 = range.row[1];
                var c1 = range.column[0];
                var c2 = range.column[1];
                if (r1 === r2 && c1 === c2) {
                    continue;
                }
                if (type === "merge-all") {
                    var fv = {};
                    var isfirst = false;
                    for (var r = r1; r <= r2; r += 1) {
                        for (var c = c1; c <= c2; c += 1) {
                            var cell = d[r][c];
                            if (cell != null &&
                                (isInlineStringCT(cell.ct) ||
                                    !_.isEmpty(cell.v) ||
                                    cell.f != null) &&
                                !isfirst) {
                                fv = _.cloneDeep(cell) || {};
                                isfirst = true;
                            }
                            d[r][c] = { mc: { r: r1, c: c1 } };
                        }
                    }
                    d[r1][c1] = fv;
                    var a = d[r1][c1];
                    if (!a)
                        return;
                    a.mc = { r: r1, c: c1, rs: r2 - r1 + 1, cs: c2 - c1 + 1 };
                    cfg.merge["".concat(r1, "_").concat(c1)] = {
                        r: r1,
                        c: c1,
                        rs: r2 - r1 + 1,
                        cs: c2 - c1 + 1,
                    };
                }
                else if (type === "merge-vertical") {
                    for (var c = c1; c <= c2; c += 1) {
                        var fv = {};
                        var isfirst = false;
                        for (var r = r1; r <= r2; r += 1) {
                            var cell = d[r][c];
                            if (cell != null &&
                                (!_.isEmpty(cell.v) || cell.f != null) &&
                                !isfirst) {
                                fv = _.cloneDeep(cell) || {};
                                isfirst = true;
                            }
                            d[r][c] = { mc: { r: r1, c: c } };
                        }
                        d[r1][c] = fv;
                        var a = d[r1][c];
                        if (!a)
                            return;
                        a.mc = { r: r1, c: c, rs: r2 - r1 + 1, cs: 1 };
                        cfg.merge["".concat(r1, "_").concat(c)] = {
                            r: r1,
                            c: c,
                            rs: r2 - r1 + 1,
                            cs: 1,
                        };
                    }
                }
                else if (type === "merge-horizontal") {
                    for (var r = r1; r <= r2; r += 1) {
                        var fv = {};
                        var isfirst = false;
                        for (var c = c1; c <= c2; c += 1) {
                            var cell = d[r][c];
                            if (cell != null &&
                                (!_.isEmpty(cell.v) || cell.f != null) &&
                                !isfirst) {
                                fv = _.cloneDeep(cell) || {};
                                isfirst = true;
                            }
                            d[r][c] = { mc: { r: r, c: c1 } };
                        }
                        d[r][c1] = fv;
                        var a = d[r][c1];
                        if (!a)
                            return;
                        a.mc = { r: r, c: c1, rs: 1, cs: c2 - c1 + 1 };
                        cfg.merge["".concat(r, "_").concat(c1)] = {
                            r: r,
                            c: c1,
                            rs: 1,
                            cs: c2 - c1 + 1,
                        };
                    }
                }
            }
        }
    }
    sheet.config = cfg;
    if (sheet.id === ctx.currentSheetId) {
        ctx.config = cfg;
    }
}
