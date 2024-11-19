import _ from "lodash";
export function mousePosition(x, y, ctx) {
    var newX = x - ctx.rowHeaderWidth;
    var newY = y -
        // ctx.infobarHeight -
        ctx.toolbarHeight -
        ctx.calculatebarHeight -
        ctx.columnHeaderHeight;
    return [newX, newY];
}
export function rowLocationByIndex(row_index, visibleRow) {
    var row = 0;
    var row_pre = 0;
    row = visibleRow[row_index];
    if (row_index === 0) {
        row_pre = 0;
    }
    else {
        row_pre = visibleRow[row_index - 1];
    }
    return [row_pre, row, row_index];
}
export function rowLocation(y, visibleRow) {
    var row_index = _.sortedIndex(visibleRow, y);
    if (row_index >= visibleRow.length && y > 0) {
        row_index = visibleRow.length - 1;
    }
    else if (row_index === -1 && y <= 0) {
        row_index = 0;
    }
    return rowLocationByIndex(row_index, visibleRow);
}
export function colLocationByIndex(col_index, visibleCol) {
    var col = 0;
    var col_pre = 0;
    col = visibleCol[col_index];
    if (col_index === 0) {
        col_pre = 0;
    }
    else {
        col_pre = visibleCol[col_index - 1];
    }
    return [col_pre, col, col_index];
}
export function colLocation(x, visibleCol) {
    var col_index = _.sortedIndex(visibleCol, x);
    if (col_index >= visibleCol.length && x > 0) {
        col_index = visibleCol.length - 1;
    }
    else if (col_index === -1 && x <= 0) {
        col_index = 0;
    }
    return colLocationByIndex(col_index, visibleCol);
}
