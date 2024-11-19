import { Context } from "../context";
import { Settings } from "../settings";
import { CellMatrix, Sheet } from "../types";
export declare function storeSheetParamALL(ctx: Context): void;
export declare function changeSheet(ctx: Context, id: string, isPivotInitial?: boolean, isNewSheet?: boolean, isCopySheet?: boolean): void;
export declare function addSheet(ctx: Context, settings?: Required<Settings>, newSheetID?: string | undefined, // if action is from websocket, there will be a new sheetID
isPivotTable?: boolean, sheetName?: string | undefined, sheetData?: Sheet | undefined): void;
export declare function deleteSheet(ctx: Context, id: string): void;
export declare function updateSheet(ctx: Context, newData: Sheet[]): void;
export declare function editSheetName(ctx: Context, editable: HTMLSpanElement): void;
export declare function expandRowsAndColumns(data: CellMatrix, rowsToAdd: number, columnsToAdd: number): CellMatrix;
