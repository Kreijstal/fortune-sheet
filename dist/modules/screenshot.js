import { Canvas, defaultStyle } from "../canvas";
// import { locale } from "../locale";
import { hasPartMC } from "./validation";
export function handleScreenShot(ctx) {
    // const { screenshot } = locale;
    if (ctx.luckysheet_select_save == null)
        return undefined;
    if (ctx.luckysheet_select_save.length === 0) {
        // if (isEditMode()) {
        //   alert(locale_screenshot.screenshotTipNoSelection);
        // } else {
        //   tooltip.info(
        //     locale_screenshot.screenshotTipTitle,
        //     locale_screenshot.screenshotTipNoSelection
        //   );
        // }
        return undefined;
    }
    if (ctx.luckysheet_select_save.length > 1) {
        // if (isEditMode()) {
        //   alert(locale_screenshot.screenshotTipHasMulti);
        // } else {
        //   tooltip.info(
        //     locale_screenshot.screenshotTipTitle,
        //     locale_screenshot.screenshotTipHasMulti
        //   );
        // }
        return undefined;
    }
    // 截图范围内包含部分合并单元格，提示
    if (ctx.config.merge != null) {
        var has_PartMC = false;
        for (var s = 0; s < ctx.luckysheet_select_save.length; s += 1) {
            var r1 = ctx.luckysheet_select_save[s].row[0];
            var r2 = ctx.luckysheet_select_save[s].row[1];
            var c1 = ctx.luckysheet_select_save[s].column[0];
            var c2 = ctx.luckysheet_select_save[s].column[1];
            has_PartMC = hasPartMC(ctx, ctx.config, r1, r2, c1, c2);
            if (has_PartMC) {
                break;
            }
        }
        if (has_PartMC) {
            // if (isEditMode()) {
            //   alert(locale_screenshot.screenshotTipHasMerge);
            // } else {
            //   tooltip.info(
            //     locale_screenshot.screenshotTipTitle,
            //     locale_screenshot.screenshotTipHasMerge
            //   );
            // }
            return undefined;
        }
    }
    var st_r = ctx.luckysheet_select_save[0].row[0];
    var ed_r = ctx.luckysheet_select_save[0].row[1];
    var st_c = ctx.luckysheet_select_save[0].column[0];
    var ed_c = ctx.luckysheet_select_save[0].column[1];
    var scrollHeight;
    var rh_height;
    if (st_r - 1 < 0) {
        scrollHeight = 0;
        rh_height = ctx.visibledatarow[ed_r];
    }
    else {
        scrollHeight = ctx.visibledatarow[st_r - 1];
        rh_height = ctx.visibledatarow[ed_r] - ctx.visibledatarow[st_r - 1];
    }
    var scrollWidth;
    var ch_width;
    if (st_c - 1 < 0) {
        scrollWidth = 0;
        ch_width = ctx.visibledatacolumn[ed_c];
    }
    else {
        scrollWidth = ctx.visibledatacolumn[st_c - 1];
        ch_width = ctx.visibledatacolumn[ed_c] - ctx.visibledatacolumn[st_c - 1];
    }
    var newCanvasElement = document.createElement("canvas");
    newCanvasElement.width = Math.ceil(ch_width * devicePixelRatio);
    newCanvasElement.height = Math.ceil(rh_height * devicePixelRatio);
    newCanvasElement.style.width = "".concat(ch_width, "px");
    newCanvasElement.style.height = "".concat(rh_height, "px");
    var newCanvas = new Canvas(newCanvasElement, ctx);
    newCanvas.drawMain({
        scrollWidth: scrollWidth,
        scrollHeight: scrollHeight,
        drawWidth: ch_width,
        drawHeight: rh_height,
        offsetLeft: 1,
        offsetTop: 1,
        clear: true,
    });
    var ctx_newCanvas = newCanvasElement.getContext("2d");
    if (ctx_newCanvas == null)
        return undefined;
    // 补上 左边框和上边框
    ctx_newCanvas.beginPath();
    ctx_newCanvas.moveTo(0, 0);
    ctx_newCanvas.lineTo(0, ctx.devicePixelRatio * rh_height);
    ctx_newCanvas.lineWidth = ctx.devicePixelRatio * 2;
    ctx_newCanvas.strokeStyle = defaultStyle.strokeStyle;
    ctx_newCanvas.stroke();
    ctx_newCanvas.closePath();
    ctx_newCanvas.beginPath();
    ctx_newCanvas.moveTo(0, 0);
    ctx_newCanvas.lineTo(ctx.devicePixelRatio * ch_width, 0);
    ctx_newCanvas.lineWidth = ctx.devicePixelRatio * 2;
    ctx_newCanvas.strokeStyle = defaultStyle.strokeStyle;
    ctx_newCanvas.stroke();
    ctx_newCanvas.closePath();
    var image = new Image();
    var url = newCanvasElement.toDataURL("image/png");
    image.src = url;
    if (ch_width > rh_height) {
        image.style.width = "100%";
    }
    else {
        image.style.height = "100%";
    }
    newCanvasElement.remove();
    return image.src;
}
