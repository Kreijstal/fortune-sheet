import React from "react";
import { Sheet, Selection, CellMatrix, Cell } from "./types";
export type Hooks = {
    beforeUpdateCell?: (r: number, c: number, value: any) => boolean;
    afterUpdateCell?: (row: number, column: number, oldValue: any, newValue: any) => void;
    afterSelectionChange?: (sheetId: string, selection: Selection) => void;
    beforeRenderRowHeaderCell?: (rowNumber: string, rowIndex: number, top: number, width: number, height: number, ctx: CanvasRenderingContext2D) => boolean;
    afterRenderRowHeaderCell?: (rowNumber: string, rowIndex: number, top: number, width: number, height: number, ctx: CanvasRenderingContext2D) => void;
    beforeRenderColumnHeaderCell?: (columnChar: string, columnIndex: number, left: number, width: number, height: number, ctx: CanvasRenderingContext2D) => boolean;
    afterRenderColumnHeaderCell?: (columnChar: string, columnIndex: number, left: number, width: number, height: number, ctx: CanvasRenderingContext2D) => void;
    beforeRenderCellArea?: (cells: CellMatrix, ctx: CanvasRenderingContext2D) => boolean;
    beforeRenderCell?: (cell: Cell | null, cellInfo: {
        row: number;
        column: number;
        startX: number;
        startY: number;
        endX: number;
        endY: number;
    }, ctx: CanvasRenderingContext2D) => boolean;
    afterRenderCell?: (cell: Cell | null, cellInfo: {
        row: number;
        column: number;
        startX: number;
        startY: number;
        endX: number;
        endY: number;
    }, ctx: CanvasRenderingContext2D) => void;
    beforeCellMouseDown?: (cell: Cell | null, cellInfo: {
        row: number;
        column: number;
        startRow: number;
        startColumn: number;
        endRow: number;
        endColumn: number;
    }) => boolean;
    afterCellMouseDown?: (cell: Cell | null, cellInfo: {
        row: number;
        column: number;
        startRow: number;
        startColumn: number;
        endRow: number;
        endColumn: number;
    }) => void;
    beforePaste?: (selection: Selection[] | undefined, content: string) => boolean;
    beforeUpdateComment?: (row: number, column: number, value: any) => boolean;
    afterUpdateComment?: (row: number, column: number, oldValue: any, value: any) => void;
    beforeInsertComment?: (row: number, column: number) => boolean;
    afterInsertComment?: (row: number, column: number) => void;
    beforeDeleteComment?: (row: number, column: number) => boolean;
    afterDeleteComment?: (row: number, column: number) => void;
    beforeAddSheet?: (sheet: Sheet) => boolean;
    afterAddSheet?: (sheet: Sheet) => void;
    beforeActivateSheet?: (id: string) => boolean;
    afterActivateSheet?: (id: string) => void;
    beforeDeleteSheet?: (id: string) => boolean;
    afterDeleteSheet?: (id: string) => void;
    beforeUpdateSheetName?: (id: string, oldName: string, newName: string) => boolean;
    afterUpdateSheetName?: (id: string, oldName: string, newName: string) => void;
};
export type Settings = {
    column?: number;
    row?: number;
    addRows?: number;
    allowEdit?: boolean;
    showToolbar?: boolean;
    showFormulaBar?: boolean;
    showSheetTabs?: boolean;
    data: Sheet[];
    config?: any;
    devicePixelRatio?: number;
    lang?: string | null;
    forceCalculation?: boolean;
    rowHeaderWidth?: number;
    columnHeaderHeight?: number;
    defaultColWidth?: number;
    defaultRowHeight?: number;
    defaultFontSize?: number;
    toolbarItems?: string[];
    cellContextMenu?: string[];
    headerContextMenu?: string[];
    sheetTabContextMenu?: string[];
    filterContextMenu?: string[];
    generateSheetId?: () => string;
    hooks?: Hooks;
    customToolbarItems?: {
        key: string;
        tooltip?: string;
        children?: React.ReactNode;
        iconName?: string;
        icon?: React.ReactNode;
        onClick?: (e: React.MouseEvent<HTMLDivElement, MouseEvent>) => void;
    }[];
    currency?: string;
};
export declare const defaultSettings: Required<Settings>;
