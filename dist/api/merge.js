import { getSheet } from "./common";
import { mergeCells as mergeCellsInternal } from "../modules";
export function mergeCells(ctx, ranges, type, options = {}) {
    const sheet = getSheet(ctx, options);
    mergeCellsInternal(ctx, sheet.id, ranges, type);
}
export function cancelMerge(ctx, ranges, options = {}) {
    mergeCells(ctx, ranges, "merge-cancel", options);
}
