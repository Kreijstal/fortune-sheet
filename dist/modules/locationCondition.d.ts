import { Context } from "../context";
export declare function deleteCellInSave(cellSave: Record<string, number>, range: {
    row: any[];
    column: any[];
}): Record<string, number>;
export declare function getRangeArr(minR: number, // 选区行起点
maxR: number, // 选区行终点
minC: number, // 选区列起点
maxC: number, // 选区列终点
cellSave: Record<string, number>, rangeArr: {
    row: (number | null)[];
    column: (number | null)[];
}[], ctx: Context): any;
export declare function getOptionValue(constants: Record<string, boolean>): string | undefined;
export declare function getSelectRange(ctx: Context): {
    row: number[];
    column: number[];
}[];
export declare function applyLocation(range: {
    row: any[];
    column: any[];
}[], type: string, value: string | undefined, ctx: Context): {
    column: any[];
    row: any[];
}[];
