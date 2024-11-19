// export default function mobileinit(ctx: Context) {
//   ctx.cellMainSrollBarSize = 0;
//   // let luckysheet_touchmove_status = false;
//   // let luckysheet_touchmove_startPos = {};
//   // const luckysheet_touchhandle_status = false;
// }
export function handleOverlayTouchStart(ctx, e, globalCache) {
    globalCache.touchMoveStatus = true;
    var touch = e.targetTouches[0];
    globalCache.touchMoveStartPos = {
        x: touch.pageX,
        y: touch.pageY,
        vy: 0,
        moveType: "y",
    };
}
export function handleOverlayTouchMove(ctx, e, globalCache, scrollbarX, scrollbarY) {
    if (e.targetTouches.length > 1)
        return;
    var touch = e.targetTouches[0];
    if (globalCache.touchMoveStatus) {
        if (!globalCache.touchMoveStartPos)
            return;
        var slideX = touch.pageX - globalCache.touchMoveStartPos.x;
        var slideY = touch.pageY - globalCache.touchMoveStartPos.y;
        var scrollLeft = ctx.scrollLeft;
        var scrollTop = ctx.scrollTop;
        scrollLeft -= slideX;
        scrollTop -= slideY;
        scrollbarY.scrollTop = scrollTop;
        globalCache.touchMoveStartPos.vy_y = slideY;
        globalCache.touchMoveStartPos.scrollTop = scrollTop;
        scrollbarX.scrollLeft = scrollLeft;
        globalCache.touchMoveStartPos.vy_x = slideX;
        globalCache.touchMoveStartPos.scrollLeft = scrollLeft;
    }
}
export function handleOverlayTouchEnd(globalCache) {
    globalCache.touchMoveStatus = false;
    globalCache.touchHandleStatus = false;
}
