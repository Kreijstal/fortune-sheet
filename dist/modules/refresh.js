import { getFlowdata } from "../context";
import { execFunctionGroup } from "./formula";
function runExecFunction(ctx, range, index, data) {
    ctx.formulaCache.execFunctionExist = [];
    for (var s = 0; s < range.length; s += 1) {
        for (var r = range[s].row[0]; r <= range[s].row[1]; r += 1) {
            for (var c = range[s].column[0]; c <= range[s].column[1]; c += 1) {
                ctx.formulaCache.execFunctionExist.push({ r: r, c: c, i: index });
            }
        }
    }
    ctx.formulaCache.execFunctionExist.reverse();
    // @ts-ignore
    execFunctionGroup(ctx, null, null, null, null, data);
    ctx.formulaCache.execFunctionGlobalData = null;
}
export function jfrefreshgrid(ctx, data, range, isRunExecFunction) {
    if (isRunExecFunction === void 0) { isRunExecFunction = true; }
    if (data == null) {
        data = getFlowdata(ctx);
    }
    if (range == null) {
        range = ctx.luckysheet_select_save;
        if (range == null)
            return;
    }
    // clearTimeout(refreshCanvasTimeOut);
    // 更新数据的范围
    // for (let s = 0; s < range.length; s += 1) {
    //   const r1 = range[s].row[0];
    //   const c1 = range[s].column[0];
    //   if (server.allowUpdate) {
    //     // 共享编辑模式
    //     server.historyParam(ctx.flowdata, ctx.currentSheetIndex, range[s]);
    //   }
    //   // 刷新图表
    //   if (typeof ctx.chartparam.jfrefreshchartall === "function") {
    //     ctx.chartparam.jfrefreshchartall(
    //       ctx.flowdata,
    //       range[s].row[0],
    //       range[s].row[1],
    //       range[s].column[0],
    //       range[s].column[1]
    //     );
    //   }
    // }
    // 单元格数据更新联动
    if (isRunExecFunction) {
        runExecFunction(ctx, range, ctx.currentSheetId, data);
    }
    /* 选区同步 */
    // selectHightlightShow();
}
