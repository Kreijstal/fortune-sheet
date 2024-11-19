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
import { getFlowdata } from "../context";
import { locale } from "../locale";
import { delFunctionGroup, execfunction, execFunctionGroup, functionCopy, } from "../modules/formula";
import { getdatabyselection, getQKBorder } from "../modules/cell";
import { genarate, update } from "../modules/format";
import { normalizeSelection, selectionCache } from "../modules/selection";
import { getSheetIndex, isAllowEdit } from "../utils";
import { hasPartMC, isRealNum } from "../modules/validation";
import { getBorderInfoCompute } from "../modules/border";
import { expandRowsAndColumns, storeSheetParamALL } from "../modules/sheet";
import { jfrefreshgrid } from "../modules/refresh";
import { setRowHeight } from "../api";
import { CFSplitRange } from "../modules";
import clipboard from "../modules/clipboard";
function postPasteCut(ctx, source, target, RowlChange) {
    // 单元格数据更新联动
    var execF_rc = {};
    ctx.formulaCache.execFunctionExist = [];
    // clearTimeout(refreshCanvasTimeOut);
    for (var r = source.range.row[0]; r <= source.range.row[1]; r += 1) {
        for (var c = source.range.column[0]; c <= source.range.column[1]; c += 1) {
            if ("".concat(r, "_").concat(c, "_").concat(source.sheetId) in execF_rc) {
                continue;
            }
            execF_rc["".concat(r, "_").concat(c, "_").concat(source.sheetId)] = 0;
            ctx.formulaCache.execFunctionExist.push({ r: r, c: c, i: source.sheetId });
        }
    }
    for (var r = target.range.row[0]; r <= target.range.row[1]; r += 1) {
        for (var c = target.range.column[0]; c <= target.range.column[1]; c += 1) {
            if ("".concat(r, "_").concat(c, "_").concat(target.sheetId) in execF_rc) {
                continue;
            }
            execF_rc["".concat(r, "_").concat(c, "_").concat(target.sheetId)] = 0;
            ctx.formulaCache.execFunctionExist.push({ r: r, c: c, i: target.sheetId });
        }
    }
    // config
    var rowHeight;
    if (ctx.currentSheetId === source.sheetId) {
        ctx.config = source.curConfig;
        rowHeight = source.curData.length;
        ctx.luckysheetfile[getSheetIndex(ctx, target.sheetId)].config =
            target.curConfig;
    }
    else if (ctx.currentSheetId === target.sheetId) {
        ctx.config = target.curConfig;
        rowHeight = target.curData.length;
        ctx.luckysheetfile[getSheetIndex(ctx, source.sheetId)].config =
            source.curConfig;
    }
    if (RowlChange) {
        ctx.visibledatarow = [];
        ctx.rh_height = 0;
        for (var i = 0; i < rowHeight; i += 1) {
            var rowlen = ctx.defaultrowlen;
            if (ctx.config.rowlen != null && ctx.config.rowlen[i] != null) {
                rowlen = ctx.config.rowlen[i];
            }
            if (ctx.config.rowhidden != null && ctx.config.rowhidden[i] != null) {
                rowlen = ctx.config.rowhidden[i];
                ctx.visibledatarow.push(ctx.rh_height);
                continue;
            }
            else {
                ctx.rh_height += rowlen + 1;
            }
            ctx.visibledatarow.push(ctx.rh_height); // 行的临时长度分布
        }
        ctx.rh_height += 80;
        // sheetmanage.showSheet();
        if (ctx.currentSheetId === source.sheetId) {
            // const rowlenArr = computeRowlenArr(
            //   ctx,
            //   target.curData.length,
            //   target.curConfig
            // );
            // ctx.luckysheetfile[
            //   getSheetIndex(ctx, target.sheetId)!
            // ].visibledatarow = rowlenArr;
        }
        else if (ctx.currentSheetId === target.sheetId) {
            // const rowlenArr = computeRowlenArr(
            //   ctx,
            //   source.curData.length,
            //   source.curConfig
            // );
            // ctx.luckysheetfile[getSheetIndex(ctx, source.sheetId)].visibledatarow =
            //   rowlenArr;
        }
    }
    // ctx.flowdata
    if (ctx.currentSheetId === source.sheetId) {
        // ctx.flowdata = source.curData;
        ctx.luckysheetfile[getSheetIndex(ctx, target.sheetId)].data =
            target.curData;
    }
    else if (ctx.currentSheetId === target.sheetId) {
        // ctx.flowdata = target.curData;
        ctx.luckysheetfile[getSheetIndex(ctx, source.sheetId)].data =
            source.curData;
    }
    // editor.webWorkerFlowDataCache(ctx.flowdata); // worker存数据
    // ctx.luckysheetfile[getSheetIndex(ctx.currentSheetId)].data = ctx.flowdata;
    // luckysheet_select_save
    if (ctx.currentSheetId === target.sheetId) {
        ctx.luckysheet_select_save = [
            { row: target.range.row, column: target.range.column },
        ];
    }
    else {
        ctx.luckysheet_select_save = [
            { row: source.range.row, column: source.range.column },
        ];
    }
    if (ctx.luckysheet_select_save.length > 0) {
        // 有选区时，刷新一下选区
        // selectHightlightShow();
    }
    // 条件格式
    ctx.luckysheetfile[getSheetIndex(ctx, source.sheetId)].luckysheet_conditionformat_save = source.curCdformat;
    ctx.luckysheetfile[getSheetIndex(ctx, target.sheetId)].luckysheet_conditionformat_save = target.curCdformat;
    // 数据验证
    // if (ctx.currentSheetId === source.sheetId) {
    //   dataVerificationCtrl.dataVerification = source.curDataVerification;
    // } else if (ctx.currentSheetId === target.sheetId) {
    //   dataVerificationCtrl.dataVerification = target.curDataVerification;
    // }
    ctx.luckysheetfile[getSheetIndex(ctx, source.sheetId)].dataVerification =
        source.curDataVerification;
    ctx.luckysheetfile[getSheetIndex(ctx, target.sheetId)].dataVerification =
        target.curDataVerification;
    ctx.formulaCache.execFunctionExist.reverse();
    // @ts-ignore
    execFunctionGroup(ctx, null, null, null, null, target.curData);
    ctx.formulaCache.execFunctionGlobalData = null;
    // const index = getSheetIndex(ctx, ctx.currentSheetId);
    // const file = ctx.luckysheetfile[index];
    // file.scrollTop = $("#luckysheet-cell-main").scrollTop();
    // file.scrollLeft = $("#luckysheet-cell-main").scrollLeft();
    // showSheet();
    // refreshCanvasTimeOut = setTimeout(function () {
    //   luckysheetrefreshgrid();
    // }, 1);
    storeSheetParamALL(ctx);
    // saveparam
    // //来源表
    // server.saveParam("all", source["sheetId"], source["curConfig"], {
    //   k: "config",
    // });
    // //目的表
    // server.saveParam("all", target["sheetId"], target["curConfig"], {
    //   k: "config",
    // });
    // //来源表
    // server.historyParam(source["curData"], source["sheetId"], {
    //   row: source["range"]["row"],
    //   column: source["range"]["column"],
    // });
    // //目的表
    // server.historyParam(target["curData"], target["sheetId"], {
    //   row: target["range"]["row"],
    //   column: target["range"]["column"],
    // });
    // //来源表
    // server.saveParam("all", source["sheetId"], source["curCdformat"], {
    //   k: "luckysheet_conditionformat_save",
    // });
    // //目的表
    // server.saveParam("all", target["sheetId"], target["curCdformat"], {
    //   k: "luckysheet_conditionformat_save",
    // });
    // //来源表
    // server.saveParam("all", source["sheetId"], source["curDataVerification"], {
    //   k: "dataVerification",
    // });
    // //目的表
    // server.saveParam("all", target["sheetId"], target["curDataVerification"], {
    //   k: "dataVerification",
    // });
}
function pasteHandler(ctx, data, borderInfo) {
    var _a;
    var _b, _c, _d, _e, _f, _g;
    // if (
    //   !checkProtectionLockedRangeList(
    //     ctx.luckysheet_select_save,
    //     ctx.currentSheetId
    //   )
    // ) {
    //   return;
    // }
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit)
        return;
    if (((_c = (_b = ctx.luckysheet_select_save) === null || _b === void 0 ? void 0 : _b.length) !== null && _c !== void 0 ? _c : 0) !== 1) {
        // if (isEditMode()) {
        //   alert("不能对多重选择区域执行此操作，请选择单个区域，然后再试");
        // } else {
        //   tooltip.info(
        //     '<i class="fa fa-exclamation-triangle"></i>提示',
        //     "不能对多重选择区域执行此操作，请选择单个区域，然后再试"
        //   );
        // }
        return;
    }
    if (typeof data === "object") {
        if (data.length === 0) {
            return;
        }
        var cfg = ctx.config || {};
        if (cfg.merge == null) {
            cfg.merge = {};
        }
        if (JSON.stringify(borderInfo).length > 2 && cfg.borderInfo == null) {
            cfg.borderInfo = [];
        }
        var copyh = data.length;
        var copyc = data[0].length;
        var minh = ctx.luckysheet_select_save[0].row[0]; // 应用范围首尾行
        var maxh = minh + copyh - 1;
        var minc = ctx.luckysheet_select_save[0].column[0]; // 应用范围首尾列
        var maxc = minc + copyc - 1;
        // 应用范围包含部分合并单元格，则return提示
        var has_PartMC = false;
        if (cfg.merge != null) {
            has_PartMC = hasPartMC(ctx, cfg, minh, maxh, minc, maxc);
        }
        if (has_PartMC) {
            // if (isEditMode()) {
            //   alert("不能对合并单元格做部分更改");
            // } else {
            //   tooltip.info(
            //     '<i class="fa fa-exclamation-triangle"></i>提示',
            //     "不能对合并单元格做部分更改"
            //   );
            // }
            return;
        }
        var d = getFlowdata(ctx); // 取数据
        if (!d)
            return;
        var rowMaxLength = d.length;
        var cellMaxLength = d[0].length;
        // 若应用范围超过最大行或最大列，增加行列
        var addr = maxh - rowMaxLength + 1;
        var addc = maxc - cellMaxLength + 1;
        if (addr > 0 || addc > 0) {
            expandRowsAndColumns(d, addr, addc);
        }
        if (!d)
            return;
        if (cfg.rowlen == null) {
            cfg.rowlen = {};
        }
        var RowlChange = false;
        var offsetMC = {};
        for (var h = minh; h <= maxh; h += 1) {
            var x = d[h];
            var currentRowLen = ctx.defaultrowlen;
            if (cfg.rowlen[h] != null) {
                currentRowLen = cfg.rowlen[h];
            }
            for (var c = minc; c <= maxc; c += 1) {
                if ((_d = x === null || x === void 0 ? void 0 : x[c]) === null || _d === void 0 ? void 0 : _d.mc) {
                    if ("rs" in x[c].mc) {
                        delete cfg.merge["".concat(x[c].mc.r, "_").concat(x[c].mc.c)];
                    }
                    delete x[c].mc;
                }
                var value = null;
                if (data[h - minh] != null && data[h - minh][c - minc] != null) {
                    value = data[h - minh][c - minc];
                }
                x[c] = value;
                if (value != null && ((_e = x === null || x === void 0 ? void 0 : x[c]) === null || _e === void 0 ? void 0 : _e.mc)) {
                    if (x[c].mc.rs != null) {
                        x[c].mc.r = h;
                        x[c].mc.c = c;
                        // @ts-ignore
                        cfg.merge["".concat(x[c].mc.r, "_").concat(x[c].mc.c)] = x[c].mc;
                        offsetMC["".concat(value.mc.r, "_").concat(value.mc.c)] = [
                            x[c].mc.r,
                            x[c].mc.c,
                        ];
                    }
                    else {
                        x[c] = {
                            mc: {
                                r: offsetMC["".concat(value.mc.r, "_").concat(value.mc.c)][0],
                                c: offsetMC["".concat(value.mc.r, "_").concat(value.mc.c)][1],
                            },
                        };
                    }
                }
                if (borderInfo["".concat(h - minh, "_").concat(c - minc)]) {
                    var bd_obj = {
                        rangeType: "cell",
                        value: {
                            row_index: h,
                            col_index: c,
                            l: borderInfo["".concat(h - minh, "_").concat(c - minc)].l,
                            r: borderInfo["".concat(h - minh, "_").concat(c - minc)].r,
                            t: borderInfo["".concat(h - minh, "_").concat(c - minc)].t,
                            b: borderInfo["".concat(h - minh, "_").concat(c - minc)].b,
                        },
                    };
                    (_f = cfg.borderInfo) === null || _f === void 0 ? void 0 : _f.push(bd_obj);
                }
                // const fontset = luckysheetfontformat(x[c]);
                // const oneLineTextHeight = menuButton.getTextSize("田", fontset)[1];
                // // 比较计算高度和当前高度取最大高度
                // if (oneLineTextHeight > currentRowLen) {
                //   currentRowLen = oneLineTextHeight;
                //   RowlChange = true;
                // }
            }
            d[h] = x;
            if (currentRowLen !== ctx.defaultrowlen) {
                cfg.rowlen[h] = currentRowLen;
            }
        }
        ctx.luckysheet_select_save = [{ row: [minh, maxh], column: [minc, maxc] }];
        if (addr > 0 || addc > 0 || RowlChange) {
            // const allParam = {
            //   cfg,
            //   RowlChange: true,
            // };
            ctx.luckysheetfile[getSheetIndex(ctx, ctx.currentSheetId)].config = cfg;
            // jfrefreshgrid(d, ctx.luckysheet_select_save, allParam);
        }
        else {
            // const allParam = {
            //   cfg,
            // };
            ctx.luckysheetfile[getSheetIndex(ctx, ctx.currentSheetId)].config = cfg;
            // jfrefreshgrid(d, ctx.luckysheet_select_save, allParam);
            // selectHightlightShow();
        }
        jfrefreshgrid(ctx, null, undefined);
    }
    else {
        data = data.replace(/\r/g, "");
        var dataChe = [];
        var che = data.split("\n");
        var colchelen = che[0].split("\t").length;
        for (var i = 0; i < che.length; i += 1) {
            if (che[i].split("\t").length < colchelen) {
                continue;
            }
            dataChe.push(che[i].split("\t"));
        }
        var d = getFlowdata(ctx); // 取数据
        if (!d)
            return;
        var last = (_g = ctx.luckysheet_select_save) === null || _g === void 0 ? void 0 : _g[ctx.luckysheet_select_save.length - 1];
        if (!last)
            return;
        var curR = last.row == null ? 0 : last.row[0];
        var curC = last.column == null ? 0 : last.column[0];
        var rlen = dataChe.length;
        var clen = dataChe[0].length;
        // 应用范围包含部分合并单元格，则return提示
        var has_PartMC = false;
        if (ctx.config.merge != null) {
            has_PartMC = hasPartMC(ctx, ctx.config, curR, curR + rlen - 1, curC, curC + clen - 1);
        }
        if (has_PartMC) {
            // if (isEditMode()) {
            //   alert("不能对合并单元格做部分更改");
            // } else {
            //   tooltip.info(
            //     '<i class="fa fa-exclamation-triangle"></i>提示',
            //     "不能对合并单元格做部分更改"
            //   );
            // }
            return;
        }
        var addr = curR + rlen - d.length;
        var addc = curC + clen - d[0].length;
        if (addr > 0 || addc > 0) {
            expandRowsAndColumns(d, addr, addc);
        }
        if (!d)
            return;
        for (var r = 0; r < rlen; r += 1) {
            var x = d[r + curR];
            for (var c = 0; c < clen; c += 1) {
                var originCell = x[c + curC];
                var value = dataChe[r][c];
                if (isRealNum(value)) {
                    // 如果单元格设置了纯文本格式，那么就不要转成数值类型了，防止数值过大自动转成科学计数法
                    if (originCell && originCell.ct && originCell.ct.fa === "@") {
                        value = String(value);
                    }
                    else {
                        value = parseFloat(value);
                    }
                }
                if (originCell) {
                    originCell.v = value;
                    if (originCell.ct != null && originCell.ct.fa != null) {
                        originCell.m = update(originCell.ct.fa, value);
                    }
                    else {
                        originCell.m = value;
                    }
                    if (originCell.f != null && originCell.f.length > 0) {
                        originCell.f = "";
                        delFunctionGroup(ctx, r + curR, c + curC, ctx.currentSheetId);
                    }
                }
                else {
                    var cell = {};
                    var mask = genarate(value);
                    _a = mask, cell.m = _a[0], cell.ct = _a[1], cell.v = _a[2];
                    x[c + curC] = cell;
                }
            }
            d[r + curR] = x;
        }
        last.row = [curR, curR + rlen - 1];
        last.column = [curC, curC + clen - 1];
        // if (addr > 0 || addc > 0) {
        //   const allParam = {
        //     RowlChange: true,
        //   };
        //   jfrefreshgrid(d, ctx.luckysheet_select_save, allParam);
        // } else {
        //   jfrefreshgrid(d, ctx.luckysheet_select_save);
        //   selectHightlightShow();
        // }
        jfrefreshgrid(ctx, null, undefined);
    }
}
function setCellHyperlink(ctx, id, r, c, link) {
    var index = getSheetIndex(ctx, id);
    if (!ctx.luckysheetfile[index].hyperlink) {
        ctx.luckysheetfile[index].hyperlink = {};
    }
    ctx.luckysheetfile[index].hyperlink["".concat(r, "_").concat(c)] = link;
}
function pasteHandlerOfCutPaste(ctx, copyRange) {
    var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l;
    // if (
    //   !checkProtectionLockedRangeList(
    //     ctx.luckysheet_select_save,
    //     ctx.currentSheetId
    //   )
    // ) {
    //   return;
    // }
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit)
        return;
    if (!copyRange)
        return;
    var cfg = ctx.config || {};
    if (cfg.merge == null) {
        cfg.merge = {};
    }
    // 复制范围
    var copyHasMC = copyRange.HasMC;
    var copyRowlChange = copyRange.RowlChange;
    var copySheetId = copyRange.dataSheetId;
    var c_r1 = copyRange.copyRange[0].row[0];
    var c_r2 = copyRange.copyRange[0].row[1];
    var c_c1 = copyRange.copyRange[0].column[0];
    var c_c2 = copyRange.copyRange[0].column[1];
    var copyData = _.cloneDeep(getdatabyselection(ctx, { row: [c_r1, c_r2], column: [c_c1, c_c2] }, copySheetId));
    var copyh = copyData.length;
    var copyc = copyData[0].length;
    // 应用范围
    var last = (_a = ctx.luckysheet_select_save) === null || _a === void 0 ? void 0 : _a[ctx.luckysheet_select_save.length - 1];
    if (!last || last.row_focus == null || last.column_focus == null)
        return;
    var minh = last.row_focus;
    var maxh = minh + copyh - 1; // 应用范围首尾行
    var minc = last.column_focus;
    var maxc = minc + copyc - 1; // 应用范围首尾列
    // 应用范围包含部分合并单元格，则提示
    var has_PartMC = false;
    if (cfg.merge != null) {
        has_PartMC = hasPartMC(ctx, cfg, minh, maxh, minc, maxc);
    }
    if (has_PartMC) {
        // if (isEditMode()) {
        //   alert("不能对合并单元格做部分更改");
        // } else {
        //   tooltip.info(
        //     '<i class="fa fa-exclamation-triangle"></i>提示',
        //     "不能对合并单元格做部分更改"
        //   );
        // }
        return;
    }
    var d = getFlowdata(ctx); // 取数据
    if (!d)
        return;
    var rowMaxLength = d.length;
    var cellMaxLength = d[0].length;
    var addr = copyh + minh - rowMaxLength;
    var addc = copyc + minc - cellMaxLength;
    if (addr > 0 || addc > 0) {
        expandRowsAndColumns(d, addr, addc);
    }
    var borderInfoCompute = getBorderInfoCompute(ctx, copySheetId);
    var c_dataVerification = _.cloneDeep(ctx.luckysheetfile[getSheetIndex(ctx, copySheetId)].dataVerification) || {};
    var dataVerification = _.cloneDeep(ctx.luckysheetfile[getSheetIndex(ctx, ctx.currentSheetId)]
        .dataVerification) || {};
    // 若选区内包含超链接
    if (((_b = ctx.luckysheet_select_save) === null || _b === void 0 ? void 0 : _b.length) === 1 &&
        ((_c = ctx.luckysheet_copy_save) === null || _c === void 0 ? void 0 : _c.copyRange.length) === 1) {
        _.forEach((_d = ctx.luckysheet_copy_save) === null || _d === void 0 ? void 0 : _d.copyRange, function (range) {
            var _a, _b, _c;
            for (var r = 0; r <= range.row[1] - range.row[0]; r += 1) {
                for (var c = 0; c <= range.column[1] - range.column[0]; c += 1) {
                    var index = getSheetIndex(ctx, (_a = ctx.luckysheet_copy_save) === null || _a === void 0 ? void 0 : _a.dataSheetId);
                    if (((_b = ctx.luckysheetfile[index].data[r + range.row[0]][c + range.column[0]]) === null || _b === void 0 ? void 0 : _b.hl) &&
                        ctx.luckysheetfile[index].hyperlink["".concat(r, "_").concat(c)]) {
                        setCellHyperlink(ctx, (_c = ctx.luckysheet_copy_save) === null || _c === void 0 ? void 0 : _c.dataSheetId, r + ctx.luckysheet_select_save[0].row[0], c + ctx.luckysheet_select_save[0].column[0], ctx.luckysheetfile[index].hyperlink["".concat(r, "_").concat(c)]);
                    }
                }
            }
        });
    }
    // 剪切粘贴在当前表操作，删除剪切范围内数据、合并单元格、数据验证和超链接
    if (ctx.currentSheetId === copySheetId) {
        for (var i = c_r1; i <= c_r2; i += 1) {
            for (var j = c_c1; j <= c_c2; j += 1) {
                var cell = d[i][j];
                if (cell && _.isPlainObject(cell) && "mc" in cell) {
                    if (((_e = cell.mc) === null || _e === void 0 ? void 0 : _e.rs) != null) {
                        delete cfg.merge["".concat(cell.mc.r, "_").concat(cell.mc.c)];
                    }
                    delete cell.mc;
                }
                d[i][j] = null;
                delete dataVerification["".concat(i, "_").concat(j)];
                (_f = ctx.luckysheetfile[getSheetIndex(ctx, ctx.currentSheetId)].hyperlink) === null || _f === void 0 ? true : delete _f["".concat(i, "_").concat(j)];
            }
        }
        // 边框
        if (cfg.borderInfo && cfg.borderInfo.length > 0) {
            var source_borderInfo = [];
            for (var i = 0; i < cfg.borderInfo.length; i += 1) {
                var bd_rangeType = cfg.borderInfo[i].rangeType;
                if (bd_rangeType === "range") {
                    var bd_range = cfg.borderInfo[i].range;
                    var bd_emptyRange = [];
                    for (var j = 0; j < bd_range.length; j += 1) {
                        bd_emptyRange = bd_emptyRange.concat(CFSplitRange(bd_range[j], { row: [c_r1, c_r2], column: [c_c1, c_c2] }, { row: [minh, maxh], column: [minc, maxc] }, "restPart"));
                    }
                    cfg.borderInfo[i].range = bd_emptyRange;
                    source_borderInfo.push(cfg.borderInfo[i]);
                }
                else if (bd_rangeType === "cell") {
                    var bd_r = cfg.borderInfo[i].value.row_index;
                    var bd_c = cfg.borderInfo[i].value.col_index;
                    if (!(bd_r >= c_r1 && bd_r <= c_r2 && bd_c >= c_c1 && bd_c <= c_c2)) {
                        source_borderInfo.push(cfg.borderInfo[i]);
                    }
                }
            }
            cfg.borderInfo = source_borderInfo;
        }
    }
    var offsetMC = {};
    for (var h = minh; h <= maxh; h += 1) {
        var x = d[h];
        for (var c = minc; c <= maxc; c += 1) {
            if (borderInfoCompute["".concat(c_r1 + h - minh, "_").concat(c_c1 + c - minc)] &&
                !borderInfoCompute["".concat(c_r1 + h - minh, "_").concat(c_c1 + c - minc)].s) {
                var bd_obj = {
                    rangeType: "cell",
                    value: {
                        row_index: h,
                        col_index: c,
                        l: borderInfoCompute["".concat(c_r1 + h - minh, "_").concat(c_c1 + c - minc)].l,
                        r: borderInfoCompute["".concat(c_r1 + h - minh, "_").concat(c_c1 + c - minc)].r,
                        t: borderInfoCompute["".concat(c_r1 + h - minh, "_").concat(c_c1 + c - minc)].t,
                        b: borderInfoCompute["".concat(c_r1 + h - minh, "_").concat(c_c1 + c - minc)].b,
                    },
                };
                if (cfg.borderInfo == null) {
                    cfg.borderInfo = [];
                }
                cfg.borderInfo.push(bd_obj);
            }
            else if (borderInfoCompute["".concat(h, "_").concat(c)]) {
                var bd_obj = {
                    rangeType: "cell",
                    value: {
                        row_index: h,
                        col_index: c,
                        l: null,
                        r: null,
                        t: null,
                        b: null,
                    },
                };
                if (cfg.borderInfo == null) {
                    cfg.borderInfo = [];
                }
                cfg.borderInfo.push(bd_obj);
            }
            else if (borderInfoCompute["".concat(c_r1 + h - minh, "_").concat(c_c1 + c - minc)]) {
                var bd_obj = {
                    rangeType: "range",
                    borderType: "border-slash",
                    color: borderInfoCompute["".concat(c_r1 + h - minh, "_").concat(c_c1 + c - minc)].s.color,
                    style: borderInfoCompute["".concat(c_r1 + h - minh, "_").concat(c_c1 + c - minc)].s.style,
                    range: normalizeSelection(ctx, [{ row: [h, h], column: [c, c] }]),
                };
                if (cfg.borderInfo == null) {
                    cfg.borderInfo = [];
                }
                cfg.borderInfo.push(bd_obj);
            }
            // 数据验证 剪切
            if (c_dataVerification["".concat(c_r1 + h - minh, "_").concat(c_c1 + c - minc)]) {
                dataVerification["".concat(h, "_").concat(c)] =
                    c_dataVerification["".concat(c_r1 + h - minh, "_").concat(c_c1 + c - minc)];
            }
            if ((_g = x[c]) === null || _g === void 0 ? void 0 : _g.mc) {
                if (((_j = (_h = x[c]) === null || _h === void 0 ? void 0 : _h.mc) === null || _j === void 0 ? void 0 : _j.rs) != null) {
                    delete cfg.merge["".concat(x[c].mc.r, "_").concat(x[c].mc.c)];
                }
                delete x[c].mc;
            }
            var value = null;
            if (copyData[h - minh] != null && copyData[h - minh][c - minc] != null) {
                value = copyData[h - minh][c - minc];
            }
            x[c] = _.cloneDeep(value);
            if (value != null && copyHasMC && ((_k = x[c]) === null || _k === void 0 ? void 0 : _k.mc)) {
                if (x[c].mc.rs != null) {
                    x[c].mc.r = h;
                    x[c].mc.c = c;
                    // @ts-ignore
                    cfg.merge["".concat(x[c].mc.r, "_").concat(x[c].mc.c)] = x[c].mc;
                    offsetMC["".concat(value.mc.r, "_").concat(value.mc.c)] = [
                        x[c].mc.r,
                        x[c].mc.c,
                    ];
                }
                else {
                    x[c] = {
                        mc: {
                            r: offsetMC["".concat(value.mc.r, "_").concat(value.mc.c)][0],
                            c: offsetMC["".concat(value.mc.r, "_").concat(value.mc.c)][1],
                        },
                    };
                }
            }
        }
        d[h] = x;
    }
    last.row = [minh, maxh];
    last.column = [minc, maxc];
    // 若有行高改变，重新计算行高改变
    if (copyRowlChange) {
        // if (ctx.currentSheetId !== copySheetIndex) {
        //   cfg = rowlenByRange(d, minh, maxh, cfg);
        // } else {
        //   cfg = rowlenByRange(d, c_r1, c_r2, cfg);
        //   cfg = rowlenByRange(d, minh, maxh, cfg);
        // }
    }
    var source;
    var target;
    if (ctx.currentSheetId !== copySheetId) {
        // 跨表操作
        var sourceData = _.cloneDeep(ctx.luckysheetfile[getSheetIndex(ctx, copySheetId)].data);
        var sourceConfig = _.cloneDeep(ctx.luckysheetfile[getSheetIndex(ctx, copySheetId)].config);
        var sourceCurData = _.cloneDeep(sourceData);
        var sourceCurConfig = _.cloneDeep(sourceConfig) || {};
        if (sourceCurConfig.merge == null) {
            sourceCurConfig.merge = {};
        }
        for (var source_r = c_r1; source_r <= c_r2; source_r += 1) {
            for (var source_c = c_c1; source_c <= c_c2; source_c += 1) {
                var cell = sourceCurData[source_r][source_c];
                if (cell === null || cell === void 0 ? void 0 : cell.mc) {
                    if ("rs" in cell.mc) {
                        delete sourceCurConfig.merge["".concat(cell.mc.r, "_").concat(cell.mc.c)];
                    }
                    delete cell.mc;
                }
                sourceCurData[source_r][source_c] = null;
            }
        }
        if (copyRowlChange) {
            // sourceCurConfig = rowlenByRange(
            //   sourceCurData,
            //   c_r1,
            //   c_r2,
            //   sourceCurConfig
            // );
        }
        // 边框
        if (sourceCurConfig.borderInfo && sourceCurConfig.borderInfo.length > 0) {
            var source_borderInfo = [];
            for (var i = 0; i < sourceCurConfig.borderInfo.length; i += 1) {
                var bd_rangeType = sourceCurConfig.borderInfo[i].rangeType;
                if (bd_rangeType === "range") {
                    var bd_range = sourceCurConfig.borderInfo[i].range;
                    var bd_emptyRange = [];
                    for (var j = 0; j < bd_range.length; j += 1) {
                        bd_emptyRange = bd_emptyRange.concat(CFSplitRange(bd_range[j], { row: [c_r1, c_r2], column: [c_c1, c_c2] }, { row: [minh, maxh], column: [minc, maxc] }, "restPart"));
                    }
                    sourceCurConfig.borderInfo[i].range = bd_emptyRange;
                    source_borderInfo.push(sourceCurConfig.borderInfo[i]);
                }
                else if (bd_rangeType === "cell") {
                    var bd_r = sourceCurConfig.borderInfo[i].value.row_index;
                    var bd_c = sourceCurConfig.borderInfo[i].value.col_index;
                    if (!(bd_r >= c_r1 && bd_r <= c_r2 && bd_c >= c_c1 && bd_c <= c_c2)) {
                        source_borderInfo.push(sourceCurConfig.borderInfo[i]);
                    }
                }
            }
            sourceCurConfig.borderInfo = source_borderInfo;
        }
        // 条件格式
        var source_cdformat = _.cloneDeep(ctx.luckysheetfile[getSheetIndex(ctx, copySheetId)]
            .luckysheet_conditionformat_save);
        var source_curCdformat = _.cloneDeep(source_cdformat);
        var ruleArr = [];
        if (source_curCdformat != null && source_curCdformat.length > 0) {
            for (var i = 0; i < source_curCdformat.length; i += 1) {
                var source_curCdformat_cellrange = source_curCdformat[i].cellrange;
                var emptyRange = [];
                var emptyRange2 = [];
                for (var j = 0; j < source_curCdformat_cellrange.length; j += 1) {
                    var range = CFSplitRange(source_curCdformat_cellrange[j], { row: [c_r1, c_r2], column: [c_c1, c_c2] }, { row: [minh, maxh], column: [minc, maxc] }, "restPart");
                    emptyRange = emptyRange.concat(range);
                    var range2 = CFSplitRange(source_curCdformat_cellrange[j], { row: [c_r1, c_r2], column: [c_c1, c_c2] }, { row: [minh, maxh], column: [minc, maxc] }, "operatePart");
                    if (range2.length > 0) {
                        emptyRange2 = emptyRange2.concat(range2);
                    }
                }
                source_curCdformat[i].cellrange = emptyRange;
                if (emptyRange2.length > 0) {
                    var ruleObj = (_l = source_curCdformat[i]) !== null && _l !== void 0 ? _l : {};
                    ruleObj.cellrange = emptyRange2;
                    ruleArr.push(ruleObj);
                }
            }
        }
        var target_cdformat = _.cloneDeep(ctx.luckysheetfile[getSheetIndex(ctx, ctx.currentSheetId)]
            .luckysheet_conditionformat_save);
        var target_curCdformat = _.cloneDeep(target_cdformat);
        if (ruleArr.length > 0) {
            target_curCdformat = target_curCdformat === null || target_curCdformat === void 0 ? void 0 : target_curCdformat.concat(ruleArr);
        }
        // 数据验证
        for (var i = c_r1; i <= c_r2; i += 1) {
            for (var j = c_c1; j <= c_c2; j += 1) {
                delete c_dataVerification["".concat(i, "_").concat(j)];
            }
        }
        source = {
            sheetId: copySheetId,
            data: sourceData,
            curData: sourceCurData,
            config: sourceConfig,
            curConfig: sourceCurConfig,
            cdformat: source_cdformat,
            curCdformat: source_curCdformat,
            dataVerification: _.cloneDeep(ctx.luckysheetfile[getSheetIndex(ctx, copySheetId)].dataVerification),
            curDataVerification: c_dataVerification,
            range: {
                row: [c_r1, c_r2],
                column: [c_c1, c_c2],
            },
        };
        target = {
            sheetId: ctx.currentSheetId,
            data: getFlowdata(ctx),
            curData: d,
            config: _.cloneDeep(ctx.config),
            curConfig: cfg,
            cdformat: target_cdformat,
            curCdformat: target_curCdformat,
            dataVerification: _.cloneDeep(ctx.luckysheetfile[getSheetIndex(ctx, ctx.currentSheetId)]
                .dataVerification),
            curDataVerification: dataVerification,
            range: {
                row: [minh, maxh],
                column: [minc, maxc],
            },
        };
    }
    else {
        // 条件格式
        var cdformat = _.cloneDeep(ctx.luckysheetfile[getSheetIndex(ctx, ctx.currentSheetId)]
            .luckysheet_conditionformat_save);
        var curCdformat = _.cloneDeep(cdformat);
        if (curCdformat != null && curCdformat.length > 0) {
            for (var i = 0; i < curCdformat.length; i += 1) {
                var cellrange = curCdformat[i].cellrange;
                var emptyRange = [];
                for (var j = 0; j < cellrange.length; j += 1) {
                    var range = CFSplitRange(cellrange[j], { row: [c_r1, c_r2], column: [c_c1, c_c2] }, { row: [minh, maxh], column: [minc, maxc] }, "allPart");
                    emptyRange = emptyRange.concat(range);
                }
                curCdformat[i].cellrange = emptyRange;
            }
        }
        // 当前表操作
        source = {
            sheetId: ctx.currentSheetId,
            data: getFlowdata(ctx),
            curData: d,
            config: _.cloneDeep(ctx.config),
            curConfig: cfg,
            cdformat: cdformat,
            curCdformat: curCdformat,
            dataVerification: _.cloneDeep(ctx.luckysheetfile[getSheetIndex(ctx, ctx.currentSheetId)]
                .dataVerification),
            curDataVerification: dataVerification,
            range: {
                row: [c_r1, c_r2],
                column: [c_c1, c_c2],
            },
        };
        target = {
            sheetId: ctx.currentSheetId,
            data: getFlowdata(ctx),
            curData: d,
            config: _.cloneDeep(ctx.config),
            curConfig: cfg,
            cdformat: cdformat,
            curCdformat: curCdformat,
            dataVerification: _.cloneDeep(ctx.luckysheetfile[getSheetIndex(ctx, ctx.currentSheetId)]
                .dataVerification),
            curDataVerification: dataVerification,
            range: {
                row: [minh, maxh],
                column: [minc, maxc],
            },
        };
    }
    if (addr > 0 || addc > 0) {
        postPasteCut(ctx, source, target, true);
    }
    else {
        postPasteCut(ctx, source, target, copyRowlChange);
    }
}
function pasteHandlerOfCopyPaste(ctx, copyRange) {
    var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o;
    // if (
    //   !checkProtectionLockedRangeList(
    //     ctx.luckysheet_select_save,
    //     ctx.currentSheetId
    //   )
    // ) {
    //   return;
    // }
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit)
        return;
    if (!copyRange)
        return;
    var cfg = ctx.config;
    if (_.isNil(cfg.merge)) {
        cfg.merge = {};
    }
    // 复制范围
    var copyHasMC = copyRange.HasMC;
    var copyRowlChange = copyRange.RowlChange;
    var copySheetIndex = copyRange.dataSheetId;
    var c_r1 = copyRange.copyRange[0].row[0];
    var c_r2 = copyRange.copyRange[0].row[1];
    var c_c1 = copyRange.copyRange[0].column[0];
    var c_c2 = copyRange.copyRange[0].column[1];
    var arr = [];
    var isSameRow = false;
    var _loop_1 = function (i) {
        var arrData = getdatabyselection(ctx, {
            row: copyRange.copyRange[i].row,
            column: copyRange.copyRange[i].column,
        }, copySheetIndex);
        if (copyRange.copyRange.length > 1) {
            if (c_r1 === copyRange.copyRange[1].row[0] &&
                c_r2 === copyRange.copyRange[1].row[1]) {
                arrData = arrData[0].map(function (col, a) {
                    return arrData.map(function (row) {
                        return row[a];
                    });
                });
                arr = arr.concat(arrData);
                isSameRow = true;
            }
            else if (c_c1 === copyRange.copyRange[1].column[0] &&
                c_c2 === copyRange.copyRange[1].column[1]) {
                arr = arr.concat(arrData);
            }
        }
        else {
            arr = arrData;
        }
    };
    for (var i = 0; i < copyRange.copyRange.length; i += 1) {
        _loop_1(i);
    }
    if (isSameRow) {
        arr = arr[0].map(function (col, b) {
            return arr.map(function (row) {
                return row[b];
            });
        });
    }
    var copyData = _.cloneDeep(arr);
    // 多重选择选择区域 单元格如果有函数 则只取值 不取函数
    if (copyRange.copyRange.length > 1) {
        for (var i = 0; i < copyData.length; i += 1) {
            for (var j = 0; j < copyData[i].length; j += 1) {
                if (copyData[i][j] != null && copyData[i][j].f != null) {
                    delete copyData[i][j].f;
                    delete copyData[i][j].spl;
                }
            }
        }
    }
    var copyh = copyData.length;
    var copyc = copyData[0].length;
    // 应用范围
    var last = (_a = ctx.luckysheet_select_save) === null || _a === void 0 ? void 0 : _a[ctx.luckysheet_select_save.length - 1];
    if (!last)
        return;
    var minh = last.row[0];
    var maxh = last.row[1]; // 应用范围首尾行
    var minc = last.column[0];
    var maxc = last.column[1]; // 应用范围首尾列
    var mh = (maxh - minh + 1) % copyh;
    var mc = (maxc - minc + 1) % copyc;
    if (mh !== 0 || mc !== 0) {
        // 若应用范围不是copydata行列数的整数倍，则取copydata的行列数
        maxh = minh + copyh - 1;
        maxc = minc + copyc - 1;
    }
    // 应用范围包含部分合并单元格，则提示
    var has_PartMC = false;
    if (!_.isNil(cfg.merge)) {
        has_PartMC = hasPartMC(ctx, cfg, minh, maxh, minc, maxc);
    }
    if (has_PartMC) {
        // if (isEditMode()) {
        //   alert("不能对合并单元格做部分更改");
        // } else {
        //   tooltip.info(
        //     '<i class="fa fa-exclamation-triangle"></i>提示',
        //     "不能对合并单元格做部分更改"
        //   );
        // }
        return;
    }
    var timesH = (maxh - minh + 1) / copyh;
    var timesC = (maxc - minc + 1) / copyc;
    var d = getFlowdata(ctx); // 取数据
    if (!d)
        return;
    var rowMaxLength = d.length;
    var cellMaxLength = d[0].length;
    // 若应用范围超过最大行或最大列，增加行列
    var addr = copyh + minh - rowMaxLength;
    var addc = copyc + minc - cellMaxLength;
    if (addr > 0 || addc > 0) {
        expandRowsAndColumns(d, addr, addc);
    }
    var borderInfoCompute = getBorderInfoCompute(ctx, copySheetIndex);
    var c_dataVerification = _.cloneDeep(ctx.luckysheetfile[getSheetIndex(ctx, copySheetIndex)].dataVerification) || {};
    var dataVerification = null;
    var mth = 0;
    var mtc = 0;
    var maxcellCahe = 0;
    var maxrowCache = 0;
    var file = ctx.luckysheetfile[getSheetIndex(ctx, ctx.currentSheetId)];
    var hiddenRows = new Set(Object.keys(((_b = file.config) === null || _b === void 0 ? void 0 : _b.rowhidden) || {}));
    var hiddenCols = new Set(Object.keys(((_c = file.config) === null || _c === void 0 ? void 0 : _c.colhidden) || {}));
    for (var th = 1; th <= timesH; th += 1) {
        for (var tc = 1; tc <= timesC; tc += 1) {
            mth = minh + (th - 1) * copyh;
            mtc = minc + (tc - 1) * copyc;
            maxrowCache = minh + th * copyh;
            maxcellCahe = minc + tc * copyc;
            // 行列位移值 用于单元格有函数
            var offsetRow = mth - c_r1;
            var offsetCol = mtc - c_c1;
            var offsetMC = {};
            for (var h = mth; h < maxrowCache; h += 1) {
                // skip if row is hidden
                if (hiddenRows === null || hiddenRows === void 0 ? void 0 : hiddenRows.has(h.toString()))
                    continue;
                var x = d[h];
                for (var c = mtc; c < maxcellCahe; c += 1) {
                    if (hiddenCols === null || hiddenCols === void 0 ? void 0 : hiddenCols.has(c.toString()))
                        continue;
                    if (borderInfoCompute["".concat(c_r1 + h - mth, "_").concat(c_c1 + c - mtc)] &&
                        !borderInfoCompute["".concat(c_r1 + h - mth, "_").concat(c_c1 + c - mtc)].s) {
                        var bd_obj = {
                            rangeType: "cell",
                            value: {
                                row_index: h,
                                col_index: c,
                                l: borderInfoCompute["".concat(c_r1 + h - mth, "_").concat(c_c1 + c - mtc)].l,
                                r: borderInfoCompute["".concat(c_r1 + h - mth, "_").concat(c_c1 + c - mtc)].r,
                                t: borderInfoCompute["".concat(c_r1 + h - mth, "_").concat(c_c1 + c - mtc)].t,
                                b: borderInfoCompute["".concat(c_r1 + h - mth, "_").concat(c_c1 + c - mtc)].b,
                            },
                        };
                        if (_.isNil(cfg.borderInfo)) {
                            cfg.borderInfo = [];
                        }
                        cfg.borderInfo.push(bd_obj);
                    }
                    else if (borderInfoCompute["".concat(h, "_").concat(c)]) {
                        var bd_obj = {
                            rangeType: "cell",
                            value: {
                                row_index: h,
                                col_index: c,
                                l: null,
                                r: null,
                                t: null,
                                b: null,
                            },
                        };
                        if (_.isNil(cfg.borderInfo)) {
                            cfg.borderInfo = [];
                        }
                        cfg.borderInfo.push(bd_obj);
                    }
                    else if (borderInfoCompute["".concat(c_r1 + h - mth, "_").concat(c_c1 + c - mtc)]) {
                        var bd_obj = {
                            rangeType: "range",
                            borderType: "border-slash",
                            color: borderInfoCompute["".concat(c_r1 + h - minh, "_").concat(c_c1 + c - minc)].s
                                .color,
                            style: borderInfoCompute["".concat(c_r1 + h - minh, "_").concat(c_c1 + c - minc)].s
                                .style,
                            range: normalizeSelection(ctx, [{ row: [h, h], column: [c, c] }]),
                        };
                        if (cfg.borderInfo == null) {
                            cfg.borderInfo = [];
                        }
                        cfg.borderInfo.push(bd_obj);
                    }
                    // 数据验证 复制
                    if (c_dataVerification["".concat(c_r1 + h - mth, "_").concat(c_c1 + c - mtc)]) {
                        if (_.isNil(dataVerification)) {
                            dataVerification = _.cloneDeep(((_d = ctx.luckysheetfile[getSheetIndex(ctx, ctx.currentSheetId)]) === null || _d === void 0 ? void 0 : _d.dataVerification) || {});
                        }
                        dataVerification["".concat(h, "_").concat(c)] =
                            c_dataVerification["".concat(c_r1 + h - mth, "_").concat(c_c1 + c - mtc)];
                    }
                    if (((_e = x[c]) === null || _e === void 0 ? void 0 : _e.mc) != null) {
                        if ("rs" in x[c].mc) {
                            delete cfg.merge["".concat(x[c].mc.r, "_").concat(x[c].mc.c)];
                        }
                        delete x[c].mc;
                    }
                    var value = null;
                    if ((_f = copyData[h - mth]) === null || _f === void 0 ? void 0 : _f[c - mtc]) {
                        value = _.cloneDeep(copyData[h - mth][c - mtc]);
                    }
                    if (!_.isNil(value) && !_.isNil(value.f)) {
                        var func = value.f;
                        if (offsetRow > 0) {
                            func = "=".concat(functionCopy(ctx, func, "down", offsetRow));
                        }
                        if (offsetRow < 0) {
                            func = "=".concat(functionCopy(ctx, func, "up", Math.abs(offsetRow)));
                        }
                        if (offsetCol > 0) {
                            func = "=".concat(functionCopy(ctx, func, "right", offsetCol));
                        }
                        if (offsetCol < 0) {
                            func = "=".concat(functionCopy(ctx, func, "left", Math.abs(offsetCol)));
                        }
                        var funcV = execfunction(ctx, func, h, c, undefined, undefined, true);
                        if (!_.isNil(value.spl)) {
                            // value.f = funcV[2];
                            // value.v = funcV[1];
                            // value.spl = funcV[3].data;
                        }
                        else {
                            value.v = funcV[1], value.f = funcV[2];
                            if (!_.isNil(value.ct) && !_.isNil(value.ct.fa)) {
                                value.m = update(value.ct.fa, funcV[1]);
                            }
                            else {
                                value.m = update("General", funcV[1]);
                            }
                        }
                    }
                    x[c] = _.cloneDeep(value);
                    if (value != null && copyHasMC && ((_g = x === null || x === void 0 ? void 0 : x[c]) === null || _g === void 0 ? void 0 : _g.mc)) {
                        if (((_j = (_h = x === null || x === void 0 ? void 0 : x[c]) === null || _h === void 0 ? void 0 : _h.mc) === null || _j === void 0 ? void 0 : _j.rs) != null) {
                            x[c].mc.r = h;
                            x[c].mc.c = c;
                            // @ts-ignore
                            cfg.merge["".concat(h, "_").concat(c)] = x[c].mc;
                            offsetMC["".concat(value.mc.r, "_").concat(value.mc.c)] = [
                                x[c].mc.r,
                                x[c].mc.c,
                            ];
                        }
                        else {
                            x[c] = {
                                mc: {
                                    r: offsetMC["".concat(value.mc.r, "_").concat(value.mc.c)][0],
                                    c: offsetMC["".concat(value.mc.r, "_").concat(value.mc.c)][1],
                                },
                            };
                        }
                    }
                }
                d[h] = x;
            }
        }
    }
    // 复制范围 是否有 条件格式和数据验证
    var cdformat = null;
    if (copyRange.copyRange.length === 1) {
        var c_file = ctx.luckysheetfile[getSheetIndex(ctx, copySheetIndex)];
        var a_file = ctx.luckysheetfile[getSheetIndex(ctx, ctx.currentSheetId)];
        var ruleArr_cf = _.cloneDeep(c_file.luckysheet_conditionformat_save);
        if (!_.isNil(ruleArr_cf) && ruleArr_cf.length > 0) {
            cdformat = (_k = _.cloneDeep(a_file.luckysheet_conditionformat_save)) !== null && _k !== void 0 ? _k : [];
            for (var i = 0; i < ruleArr_cf.length; i += 1) {
                var cf_range = ruleArr_cf[i].cellrange;
                var emptyRange = [];
                for (var th = 1; th <= timesH; th += 1) {
                    for (var tc = 1; tc <= timesC; tc += 1) {
                        mth = minh + (th - 1) * copyh;
                        mtc = minc + (tc - 1) * copyc;
                        maxrowCache = minh + th * copyh;
                        maxcellCahe = minc + tc * copyc;
                        for (var j = 0; j < cf_range.length; j += 1) {
                            var range = CFSplitRange(cf_range[j], { row: [c_r1, c_r2], column: [c_c1, c_c2] }, { row: [mth, maxrowCache - 1], column: [mtc, maxcellCahe - 1] }, "operatePart");
                            if (range.length > 0) {
                                emptyRange = emptyRange.concat(range);
                            }
                        }
                    }
                }
                if (emptyRange.length > 0) {
                    ruleArr_cf[i].cellrange = emptyRange;
                    cdformat.push(ruleArr_cf[i]);
                }
            }
        }
    }
    last.row = [minh, maxh];
    last.column = [minc, maxc];
    file.config = cfg;
    file.luckysheet_conditionformat_save = cdformat;
    file.dataVerification = __assign(__assign({}, file.dataVerification), dataVerification);
    // 若选区内包含超链接
    if (((_l = ctx.luckysheet_select_save) === null || _l === void 0 ? void 0 : _l.length) === 1 &&
        ((_m = ctx.luckysheet_copy_save) === null || _m === void 0 ? void 0 : _m.copyRange.length) === 1) {
        _.forEach((_o = ctx.luckysheet_copy_save) === null || _o === void 0 ? void 0 : _o.copyRange, function (range) {
            var _a, _b, _c;
            for (var r = 0; r <= range.row[1] - range.row[0]; r += 1) {
                for (var c = 0; c <= range.column[1] - range.column[0]; c += 1) {
                    var index = getSheetIndex(ctx, (_a = ctx.luckysheet_copy_save) === null || _a === void 0 ? void 0 : _a.dataSheetId);
                    if (((_b = ctx.luckysheetfile[index].data[r + range.row[0]][c + range.column[0]]) === null || _b === void 0 ? void 0 : _b.hl) &&
                        ctx.luckysheetfile[index].hyperlink["".concat(r, "_").concat(c)]) {
                        setCellHyperlink(ctx, (_c = ctx.luckysheet_copy_save) === null || _c === void 0 ? void 0 : _c.dataSheetId, r + ctx.luckysheet_select_save[0].row[0], c + ctx.luckysheet_select_save[0].column[0], ctx.luckysheetfile[index].hyperlink["".concat(r, "_").concat(c)]);
                    }
                }
            }
        });
    }
    if (copyRowlChange || addr > 0 || addc > 0) {
        // cfg = rowlenByRange(d, minh, maxh, cfg);
        // const allParam = {
        //   cfg,
        //   RowlChange: true,
        //   cdformat,
        //   dataVerification,
        // };
        jfrefreshgrid(ctx, d, ctx.luckysheet_select_save);
    }
    else {
        // const allParam = {
        //   cfg,
        //   cdformat,
        //   dataVerification,
        // };
        jfrefreshgrid(ctx, d, ctx.luckysheet_select_save);
        // selectHightlightShow();
    }
}
function handleFormulaStringPaste(ctx, formulaStr) {
    // plaintext formula is applied only to one cell
    var r = ctx.luckysheet_select_save[0].row[0];
    var c = ctx.luckysheet_select_save[0].column[0];
    var funcV = execfunction(ctx, formulaStr, r, c, undefined, undefined, true);
    var val = funcV[1];
    var d = getFlowdata(ctx);
    if (!d)
        return;
    if (!d[r][c])
        d[r][c] = {};
    d[r][c].m = val.toString();
    d[r][c].v = val;
    d[r][c].f = formulaStr;
}
export function handlePaste(ctx, e) {
    var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o, _p, _q;
    // if (isEditMode()) {
    //   // 此模式下禁用粘贴
    //   return;
    // }
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit)
        return;
    if (selectionCache.isPasteAction) {
        ctx.luckysheetCellUpdate = [];
        // $("#luckysheet-rich-text-editor").blur();
        selectionCache.isPasteAction = false;
        var clipboardData = e.clipboardData;
        if (!clipboardData) {
            // @ts-ignore
            // for IE
            clipboardData = window.clipboardData;
        }
        if (!clipboardData)
            return;
        var txtdata = clipboardData.getData("text/html") || clipboardData.getData("text/plain");
        // 如果标示是qksheet复制的内容，判断剪贴板内容是否是当前页面复制的内容
        var isEqual = true;
        if (txtdata.indexOf("fortune-copy-action-table") > -1 &&
            ((_a = ctx.luckysheet_copy_save) === null || _a === void 0 ? void 0 : _a.copyRange) != null &&
            ctx.luckysheet_copy_save.copyRange.length > 0) {
            // 剪贴板内容解析
            var cpDataArr = [];
            var reg = /<tr.*?>(.*?)<\/tr>/g;
            var reg2 = /<td.*?>(.*?)<\/td>/g;
            var regArr = txtdata.match(reg) || [];
            for (var i = 0; i < regArr.length; i += 1) {
                var cpRowArr = [];
                var reg2Arr = regArr[i].match(reg2);
                if (!_.isNil(reg2Arr)) {
                    for (var j = 0; j < reg2Arr.length; j += 1) {
                        var cpValue = reg2Arr[j]
                            .replace(/<td.*?>/g, "")
                            .replace(/<\/td>/g, "");
                        cpRowArr.push(cpValue);
                    }
                }
                cpDataArr.push(cpRowArr);
            }
            // 当前页面复制区内容
            var copy_r1 = ctx.luckysheet_copy_save.copyRange[0].row[0];
            var copy_r2 = ctx.luckysheet_copy_save.copyRange[0].row[1];
            var copy_c1 = ctx.luckysheet_copy_save.copyRange[0].column[0];
            var copy_c2 = ctx.luckysheet_copy_save.copyRange[0].column[1];
            var copy_index = ctx.luckysheet_copy_save.dataSheetId;
            var d = void 0;
            if (copy_index === ctx.currentSheetId) {
                d = getFlowdata(ctx);
            }
            else {
                var index = getSheetIndex(ctx, copy_index);
                if (_.isNil(index))
                    return;
                d = ctx.luckysheetfile[index].data;
            }
            if (!d)
                return;
            for (var r = copy_r1; r <= copy_r2; r += 1) {
                if (r - copy_r1 > cpDataArr.length - 1) {
                    break;
                }
                for (var c = copy_c1; c <= copy_c2; c += 1) {
                    var cell = d[r][c];
                    var isInlineStr = false;
                    if (!_.isNil(cell) && !_.isNil(cell.mc) && _.isNil(cell.mc.rs)) {
                        continue;
                    }
                    var v = void 0;
                    if (!_.isNil(cell)) {
                        if (((_d = (_c = (_b = cell.ct) === null || _b === void 0 ? void 0 : _b.fa) === null || _c === void 0 ? void 0 : _c.indexOf("w")) !== null && _d !== void 0 ? _d : -1) > -1) {
                            v = (_f = (_e = d[r]) === null || _e === void 0 ? void 0 : _e[c]) === null || _f === void 0 ? void 0 : _f.v;
                        }
                        else {
                            v = (_h = (_g = d[r]) === null || _g === void 0 ? void 0 : _g[c]) === null || _h === void 0 ? void 0 : _h.m;
                        }
                    }
                    else {
                        v = "";
                    }
                    if (_.isNil(v) && ((_l = (_k = (_j = d[r]) === null || _j === void 0 ? void 0 : _j[c]) === null || _k === void 0 ? void 0 : _k.ct) === null || _l === void 0 ? void 0 : _l.t) === "inlineStr") {
                        v = d[r][c].ct.s.map(function (val) { return val.v; }).join("");
                        isInlineStr = true;
                    }
                    if (_.isNil(v)) {
                        v = "";
                    }
                    if (isInlineStr) {
                        // const cpData = $(cpDataArr[r - copy_r1][c - copy_c1])
                        //   .text()
                        //   .replace(/\s|\n/g, " ");
                        // const ctx.alue = v.replace(/\n/g, "").replace(/\s/g, " ");
                        // if (cpData !== ctx.alue) {
                        //   isEqual = false;
                        //   break;
                        // }
                    }
                    else {
                        if (_.trim(cpDataArr[r - copy_r1][c - copy_c1]) !== _.trim(v)) {
                            isEqual = false;
                            break;
                        }
                    }
                }
            }
        }
        var locale_fontjson_1 = locale(ctx).fontjson;
        if (((_o = (_m = ctx.hooks).beforePaste) === null || _o === void 0 ? void 0 : _o.call(_m, ctx.luckysheet_select_save, txtdata)) === false) {
            return;
        }
        if (txtdata.indexOf("fortune-copy-action-table") > -1 &&
            ((_p = ctx.luckysheet_copy_save) === null || _p === void 0 ? void 0 : _p.copyRange) != null &&
            ctx.luckysheet_copy_save.copyRange.length > 0 &&
            isEqual) {
            // 剪切板内容 和 luckysheet本身复制的内容 一致
            if (ctx.luckysheet_paste_iscut) {
                ctx.luckysheet_paste_iscut = false;
                pasteHandlerOfCutPaste(ctx, ctx.luckysheet_copy_save);
                ctx.luckysheet_selection_range = [];
                // selection.clearcopy(e);
            }
            else {
                pasteHandlerOfCopyPaste(ctx, ctx.luckysheet_copy_save);
            }
        }
        else if (txtdata.indexOf("fortune-copy-action-image") > -1) {
            // imageCtrl.pasteImgItem();
        }
        else {
            if (txtdata.indexOf("table") > -1) {
                var ele = document.createElement("div");
                ele.innerHTML = txtdata;
                var trList = ele.querySelectorAll("table tr");
                if (trList.length === 0) {
                    ele.remove();
                    return;
                }
                var data_1 = new Array(trList.length);
                var colLen_1 = 0;
                _.forEach(trList[0].querySelectorAll("td"), function (td) {
                    var colspan = td.colSpan;
                    if (Number.isNaN(colspan)) {
                        colspan = 1;
                    }
                    colLen_1 += colspan;
                });
                for (var i = 0; i < data_1.length; i += 1) {
                    data_1[i] = new Array(colLen_1);
                }
                var r_1 = 0;
                var borderInfo_1 = {};
                var styleInner = ((_q = ele.querySelectorAll("style")[0]) === null || _q === void 0 ? void 0 : _q.innerHTML) || "";
                var patternReg = /{([^}]*)}/g;
                var patternStyle = styleInner.match(patternReg);
                var nameReg = /^[^\t].*/gm;
                var patternName = _.initial(styleInner.match(nameReg));
                var allStyleList_1 = patternName.length === (patternStyle === null || patternStyle === void 0 ? void 0 : patternStyle.length) &&
                    typeof patternName === typeof patternStyle
                    ? _.fromPairs(_.zip(patternName, patternStyle))
                    : {};
                var index_1 = getSheetIndex(ctx, ctx.currentSheetId);
                if (!_.isNil(index_1)) {
                    if (_.isNil(ctx.luckysheetfile[index_1].config)) {
                        ctx.luckysheetfile[index_1].config = {};
                    }
                    if (_.isNil(ctx.luckysheetfile[index_1].config.rowlen)) {
                        ctx.luckysheetfile[index_1].config.rowlen = {};
                    }
                    var rowHeightList_1 = ctx.luckysheetfile[index_1].config.rowlen;
                    _.forEach(trList, function (tr) {
                        var c = 0;
                        var targetR = ctx.luckysheet_select_save[0].row[0] + r_1;
                        var targetRowHeight = !_.isNil(tr.getAttribute("height"))
                            ? parseInt(tr.getAttribute("height"), 10)
                            : null;
                        if ((_.has(ctx.luckysheetfile[index_1].config.rowlen, targetR) &&
                            ctx.luckysheetfile[index_1].config.rowlen[targetR] !==
                                targetRowHeight) ||
                            (!_.has(ctx.luckysheetfile[index_1].config.rowlen, targetR) &&
                                ctx.luckysheetfile[index_1].defaultRowHeight !== targetRowHeight)) {
                            rowHeightList_1[targetR] = targetRowHeight;
                        }
                        _.forEach(tr.querySelectorAll("td"), function (td) {
                            // build cell from td
                            var className = td.className;
                            var cell = {};
                            var txt = td.innerText || td.innerHTML;
                            if (_.trim(txt).length === 0) {
                                cell.v = undefined;
                                cell.m = "";
                            }
                            else {
                                var mask = genarate(txt);
                                // @ts-ignore
                                cell.m = mask[0], cell.ct = mask[1], cell.v = mask[2];
                            }
                            var styleString = typeof allStyleList_1[".".concat(className)] === "string"
                                ? allStyleList_1[".".concat(className)]
                                    .substring(1, allStyleList_1[".".concat(className)].length - 1)
                                    .split("\n\t")
                                : [];
                            var styles = {};
                            _.forEach(styleString, function (s) {
                                var styleList = s.split(":");
                                styles[styleList[0]] = styleList === null || styleList === void 0 ? void 0 : styleList[1].replace(";", "");
                            });
                            if (!_.isNil(styles.border))
                                td.style.border = styles.border;
                            var bg = td.style.backgroundColor || styles.background;
                            if (bg === "rgba(0, 0, 0, 0)" || _.isEmpty(bg)) {
                                bg = undefined;
                            }
                            cell.bg = bg;
                            var fontWight = td.style.fontWeight;
                            cell.bl =
                                (fontWight.toString() === "400" ||
                                    fontWight === "normal" ||
                                    _.isEmpty(fontWight)) &&
                                    !_.includes(styles["font-style"], "bold") &&
                                    (!styles["font-weight"] || styles["font-weight"] === "400")
                                    ? 0
                                    : 1;
                            cell.it =
                                (td.style.fontStyle === "normal" ||
                                    _.isEmpty(td.style.fontStyle)) &&
                                    !_.includes(styles["font-style"], "italic")
                                    ? 0
                                    : 1;
                            cell.un = !_.includes(styles["text-decoration"], "underline")
                                ? undefined
                                : 1;
                            cell.cl = !_.includes(td.innerHTML, "<s>") ? undefined : 1;
                            var ff = td.style.fontFamily || styles["font-family"] || "";
                            var ffs = ff.split(",");
                            for (var i = 0; i < ffs.length; i += 1) {
                                var fa = _.trim(ffs[i].toLowerCase());
                                // @ts-ignore
                                fa = locale_fontjson_1[fa];
                                if (_.isNil(fa)) {
                                    cell.ff = 0;
                                }
                                else {
                                    cell.ff = fa;
                                    break;
                                }
                            }
                            var fs = Math.round(styles["font-size"]
                                ? parseInt(styles["font-size"].replace("pt", ""), 10)
                                : (parseInt(td.style.fontSize || "13", 10) * 72) / 96);
                            cell.fs = fs;
                            cell.fc = td.style.color || styles.color;
                            var ht = td.style.textAlign || styles["text-align"] || "left";
                            if (ht === "center") {
                                cell.ht = 0;
                            }
                            else if (ht === "right") {
                                cell.ht = 2;
                            }
                            else {
                                cell.ht = 1;
                            }
                            var regex = /vertical-align:\s*(.*?);/;
                            var vt = td.style.verticalAlign ||
                                styles["vertical-align"] ||
                                (!_.isNil(allStyleList_1.td) &&
                                    allStyleList_1.td.match(regex).length > 0 &&
                                    allStyleList_1.td.match(regex)[1]) ||
                                "top";
                            if (vt === "middle") {
                                cell.vt = 0;
                            }
                            else if (vt === "top" || vt === "text-top") {
                                cell.vt = 1;
                            }
                            else {
                                cell.vt = 2;
                            }
                            if ("mso-rotate" in styles) {
                                var rt = styles["mso-rotate"];
                                cell.rt = parseFloat(rt);
                            }
                            while (c < colLen_1 && !_.isNil(data_1[r_1][c])) {
                                c += 1;
                            }
                            if (c === colLen_1) {
                                return true;
                            }
                            if (_.isNil(data_1[r_1][c])) {
                                data_1[r_1][c] = cell;
                                // @ts-ignore
                                var rowspan = parseInt(td.getAttribute("rowspan"), 10);
                                // @ts-ignore
                                var colspan = parseInt(td.getAttribute("colspan"), 10);
                                if (Number.isNaN(rowspan)) {
                                    rowspan = 1;
                                }
                                if (Number.isNaN(colspan)) {
                                    colspan = 1;
                                }
                                var r_ab = ctx.luckysheet_select_save[0].row[0] + r_1;
                                var c_ab = ctx.luckysheet_select_save[0].column[0] + c;
                                for (var rp = 0; rp < rowspan; rp += 1) {
                                    for (var cp = 0; cp < colspan; cp += 1) {
                                        if (rp === 0) {
                                            var bt = td.style.borderTop;
                                            if (!_.isEmpty(bt) &&
                                                bt.substring(0, 3).toLowerCase() !== "0px") {
                                                var width = td.style.borderTopWidth;
                                                var type = td.style.borderTopStyle;
                                                var color = td.style.borderTopColor;
                                                var borderconfig = getQKBorder(width, type, color);
                                                if (!borderInfo_1["".concat(r_1 + rp, "_").concat(c + cp)]) {
                                                    borderInfo_1["".concat(r_1 + rp, "_").concat(c + cp)] = {};
                                                }
                                                borderInfo_1["".concat(r_1 + rp, "_").concat(c + cp)].t = {
                                                    style: borderconfig[0],
                                                    color: borderconfig[1],
                                                };
                                            }
                                        }
                                        if (rp === rowspan - 1) {
                                            var bb = td.style.borderBottom;
                                            if (!_.isEmpty(bb) &&
                                                bb.substring(0, 3).toLowerCase() !== "0px") {
                                                var width = td.style.borderBottomWidth;
                                                var type = td.style.borderBottomStyle;
                                                var color = td.style.borderBottomColor;
                                                var borderconfig = getQKBorder(width, type, color);
                                                if (!borderInfo_1["".concat(r_1 + rp, "_").concat(c + cp)]) {
                                                    borderInfo_1["".concat(r_1 + rp, "_").concat(c + cp)] = {};
                                                }
                                                borderInfo_1["".concat(r_1 + rp, "_").concat(c + cp)].b = {
                                                    style: borderconfig[0],
                                                    color: borderconfig[1],
                                                };
                                            }
                                        }
                                        if (cp === 0) {
                                            var bl = td.style.borderLeft;
                                            if (!_.isEmpty(bl) &&
                                                bl.substring(0, 3).toLowerCase() !== "0px") {
                                                var width = td.style.borderLeftWidth;
                                                var type = td.style.borderLeftStyle;
                                                var color = td.style.borderLeftColor;
                                                var borderconfig = getQKBorder(width, type, color);
                                                if (!borderInfo_1["".concat(r_1 + rp, "_").concat(c + cp)]) {
                                                    borderInfo_1["".concat(r_1 + rp, "_").concat(c + cp)] = {};
                                                }
                                                borderInfo_1["".concat(r_1 + rp, "_").concat(c + cp)].l = {
                                                    style: borderconfig[0],
                                                    color: borderconfig[1],
                                                };
                                            }
                                        }
                                        if (cp === colspan - 1) {
                                            var br = td.style.borderLeft;
                                            if (!_.isEmpty(br) &&
                                                br.substring(0, 3).toLowerCase() !== "0px") {
                                                var width = td.style.borderRightWidth;
                                                var type = td.style.borderRightStyle;
                                                var color = td.style.borderRightColor;
                                                var borderconfig = getQKBorder(width, type, color);
                                                if (!borderInfo_1["".concat(r_1 + rp, "_").concat(c + cp)]) {
                                                    borderInfo_1["".concat(r_1 + rp, "_").concat(c + cp)] = {};
                                                }
                                                borderInfo_1["".concat(r_1 + rp, "_").concat(c + cp)].r = {
                                                    style: borderconfig[0],
                                                    color: borderconfig[1],
                                                };
                                            }
                                        }
                                        if (rp === 0 && cp === 0) {
                                            continue;
                                        }
                                        data_1[r_1 + rp][c + cp] = { mc: { r: r_ab, c: c_ab } };
                                    }
                                }
                                if (rowspan > 1 || colspan > 1) {
                                    var first = { rs: rowspan, cs: colspan, r: r_ab, c: c_ab };
                                    data_1[r_1][c].mc = first;
                                }
                            }
                            c += 1;
                            if (c === colLen_1) {
                                return true;
                            }
                            return true;
                        });
                        r_1 += 1;
                    });
                    setRowHeight(ctx, rowHeightList_1);
                }
                ctx.luckysheet_selection_range = [];
                pasteHandler(ctx, data_1, borderInfo_1);
                // $("#fortune-copy-content").empty();
                ele.remove();
            }
            // 复制的是图片
            else if (clipboardData.files.length === 1 &&
                clipboardData.files[0].type.indexOf("image") > -1) {
                //   imageCtrl.insertImg(clipboardData.files[0]);
            }
            else {
                txtdata = clipboardData.getData("text/plain");
                var isExcelFormula = txtdata.startsWith("=");
                if (isExcelFormula) {
                    handleFormulaStringPaste(ctx, txtdata);
                }
                else {
                    pasteHandler(ctx, txtdata);
                }
            }
        }
    }
    else if (ctx.luckysheetCellUpdate.length > 0) {
        // 阻止默认粘贴
        e.preventDefault();
        var clipboardData = e.clipboardData;
        if (!clipboardData) {
            // for IE
            // @ts-ignore
            clipboardData = window.clipboardData;
        }
        var text = clipboardData === null || clipboardData === void 0 ? void 0 : clipboardData.getData("text/plain");
        if (text) {
            document.execCommand("insertText", false, text);
        }
    }
}
export function handlePasteByClick(ctx, clipboardData, triggerType) {
    var _a, _b, _c;
    var allowEdit = isAllowEdit(ctx);
    if (!allowEdit)
        return;
    if (clipboardData)
        clipboard.writeHtml(clipboardData);
    var textarea = document.querySelector("#fortune-copy-content");
    // textarea.focus();
    // textarea.select();
    // 等50毫秒，keyPress事件发生了再去处理数据
    // setTimeout(function () {
    var data = (textarea === null || textarea === void 0 ? void 0 : textarea.innerHTML) || (textarea === null || textarea === void 0 ? void 0 : textarea.textContent);
    if (!data)
        return;
    if (((_b = (_a = ctx.hooks).beforePaste) === null || _b === void 0 ? void 0 : _b.call(_a, ctx.luckysheet_select_save, data)) === false) {
        return;
    }
    if (data.indexOf("fortune-copy-action-table") > -1 &&
        ((_c = ctx.luckysheet_copy_save) === null || _c === void 0 ? void 0 : _c.copyRange) != null &&
        ctx.luckysheet_copy_save.copyRange.length > 0) {
        if (ctx.luckysheet_paste_iscut) {
            ctx.luckysheet_paste_iscut = false;
            pasteHandlerOfCutPaste(ctx, ctx.luckysheet_copy_save);
            // clearcopy(e);
        }
        else {
            pasteHandlerOfCopyPaste(ctx, ctx.luckysheet_copy_save);
        }
    }
    else if (data.indexOf("fortune-copy-action-image") > -1) {
        // imageCtrl.pasteImgItem();
    }
    else if (triggerType !== "btn") {
        var isExcelFormula = clipboardData.startsWith("=");
        if (isExcelFormula) {
            handleFormulaStringPaste(ctx, clipboardData);
        }
        else {
            pasteHandler(ctx, clipboardData);
        }
    }
    else {
        // if (isEditMode()) {
        //   alert(local_drag.pasteMustKeybordAlert);
        // } else {
        //   tooltip.info(
        //     local_drag.pasteMustKeybordAlertHTMLTitle,
        //     local_drag.pasteMustKeybordAlertHTML
        //   );
        // }
    }
    // }, 10);
}
