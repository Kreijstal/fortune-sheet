import { Context } from "../context";
/**
 * 增加行列
 * @param {string} type 行或列 ['row', 'column'] 之一
 * @param {number} index 插入的位置 index
 * @param {number} count 插入 多少 行（列）
 * @param {string} direction 哪个方向插入 ['lefttop','rightbottom'] 之一
 * @param {string | number} id 操作的 sheet 的 id
 * @returns
 */
export declare function insertRowCol(ctx: Context, op: {
    type: "row" | "column";
    index: number;
    count: number;
    direction: "lefttop" | "rightbottom";
    id: string;
}, changeSelection?: boolean): void;
export declare function deleteRowCol(ctx: Context, op: {
    type: "row" | "column";
    start: number;
    end: number;
    id?: string;
}): void;
export declare function computeRowlenArr(ctx: Context, rowHeight: number, cfg: any): number[];
export declare function hideSelected(ctx: Context, type: string): "" | "noMulti";
export declare function showSelected(ctx: Context, type: string): "" | "noMulti";
export declare function isShowHidenCR(ctx: Context): boolean;
export declare function hideCRCount(ctx: Context, type: string): number;
