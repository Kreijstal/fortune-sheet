import { Context } from "../context";
import { Cell, CellMatrix, Range, Selection, SingleRange } from "../types";
export declare function normalizedCellAttr(cell: Cell, attr: keyof Cell, defaultFontSize?: number): any;
export declare function normalizedAttr(data: CellMatrix, r: number, c: number, attr: keyof Cell): any;
export declare function getCellValue(r: number, c: number, data: CellMatrix, attr?: keyof Cell): any;
export declare function setCellValue(ctx: Context, r: number, c: number, d: CellMatrix | null | undefined, v: any): void;
export declare function getRealCellValue(r: number, c: number, data: CellMatrix, attr?: keyof Cell): any;
export declare function mergeBorder(ctx: Context, d: CellMatrix, row_index: number, col_index: number): {
    row: number[];
    column: number[];
} | null;
export declare function mergeMoveMain(ctx: Context, columnseleted: number[], rowseleted: number[], s: Partial<Selection>, top: number, height: number, left: number, width: number): (number | number[])[] | null;
export declare function cancelFunctionrangeSelected(ctx: Context): void;
export declare function cancelNormalSelected(ctx: Context): void;
export declare function updateCell(ctx: Context, r: number, c: number, $input?: HTMLDivElement | null, value?: any, canvas?: CanvasRenderingContext2D): void;
export declare function getOrigincell(ctx: Context, r: number, c: number, i: string): Cell | null;
export declare function getcellFormula(ctx: Context, r: number, c: number, i: string, data?: any): any;
export declare function getRange(ctx: Context): Range;
export declare function getFlattenedRange(ctx: Context, range?: Range): {
    r: number;
    c: number;
}[];
export declare function getRangetxt(ctx: Context, sheetId: string, range: SingleRange, currentId?: string): string;
export declare function getRangeByTxt(ctx: Context, txt: string): any[];
export declare function isAllSelectedCellsInStatus(ctx: Context, attr: keyof Cell, status: any): boolean;
export declare function getFontStyleByCell(cell: Cell | null | undefined, checksAF?: any[], checksCF?: any, isCheck?: boolean): any;
export declare function getStyleByCell(ctx: Context, d: CellMatrix, r: number, c: number): any;
export declare function getInlineStringHTML(r: number, c: number, data: CellMatrix): string;
export declare function getQKBorder(width: string, type: string, color: string): (string | number)[];
/**
 * 计算范围行高
 *
 * @param d 原始数据
 * @param r1 起始行
 * @param r2 截至行
 * @param cfg 配置
 * @returns 计算后的配置
 */
export declare function getdatabyselection(ctx: Context, range: Selection | undefined, sheetId: string): (Cell | null)[][];
export declare function luckysheetUpdateCell(ctx: Context, row_index: number, col_index: number): void;
export declare function getDataBySelectionNoCopy(ctx: Context, range: Selection): (Cell | null)[][];
