import _ from "lodash";
import { FormulaCache } from "./modules";
import { normalizeSelection } from "./modules/selection";
import { getSheetIndex } from "./utils";
export function defaultContext(refs) {
    return {
        luckysheetfile: [],
        defaultcolumnNum: 60,
        defaultrowNum: 84,
        addDefaultRows: 50,
        fullscreenmode: true,
        devicePixelRatio: (typeof globalThis !== 'undefined' ? globalThis.devicePixelRatio : (typeof window !== 'undefined' ? window.devicePixelRatio : 1)),
        contextMenu: {},
        sheetTabContextMenu: {},
        currentSheetId: "",
        calculateSheetId: "",
        config: {},
        // 提醒弹窗
        warnDialog: undefined,
        currency: "¥",
        rangeDialog: {
            show: false,
            rangeTxt: "",
            type: "",
            singleSelect: false,
        },
        dataVerification: {
            selectStatus: false,
            selectRange: [],
            optionLabel_en: {
                number: "numeric",
                number_integer: "integer",
                number_decimal: "decimal",
                between: "between",
                notBetween: "not between",
                equal: "equal to",
                notEqualTo: "not equal to",
                moreThanThe: "greater",
                lessThan: "less than",
                greaterOrEqualTo: "greater or equal to",
                lessThanOrEqualTo: "less than or equal to",
                include: "include",
                exclude: "not include",
                earlierThan: "earlier than",
                noEarlierThan: "not earlier than",
                laterThan: "later than",
                noLaterThan: "not later than",
                identificationNumber: "identification number",
                phoneNumber: "phone number",
            },
            optionLabel_hi: {
                number: "संख्यात्मक",
                number_integer: "पूर्णांक",
                number_decimal: "दशमलव",
                between: "के बीच",
                notBetween: "के बीच नहीं",
                equal: "के बराबर",
                notEqualTo: "के बराबर नहीं",
                moreThanThe: "से अधिक",
                lessThan: "से कम",
                greaterOrEqualTo: "के बराबर या अधिक",
                lessThanOrEqualTo: "के बराबर या कम",
                include: "शामिल",
                exclude: "शामिल नहीं",
                earlierThan: "से पहले",
                noEarlierThan: "से पहले नहीं",
                laterThan: "के बाद",
                noLaterThan: "के बाद नहीं",
                identificationNumber: "पहचान संख्या",
                phoneNumber: "फोन नंबर",
            },
            optionLabel_zh: {
                number: "数值",
                number_integer: "整数",
                number_decimal: "小数",
                between: "介于",
                notBetween: "不介于",
                equal: "等于",
                notEqualTo: "不等于",
                moreThanThe: "大于",
                lessThan: "小于",
                greaterOrEqualTo: "大于等于",
                lessThanOrEqualTo: "小于等于",
                include: "包括",
                exclude: "不包括",
                earlierThan: "早于",
                noEarlierThan: "不早于",
                laterThan: "晚于",
                noLaterThan: "不晚于",
                identificationNumber: "身份证号码",
                phoneNumber: "手机号",
            },
            optionLabel_zh_tw: {
                number: "數位",
                number_integer: "數位-整數",
                number_decimal: "數位-小數",
                between: "介於",
                notBetween: "不介於",
                equal: "等於",
                notEqualTo: "不等於",
                moreThanThe: "大於",
                lessThan: "小於",
                greaterOrEqualTo: "大於等於",
                lessThanOrEqualTo: "小於等於",
                include: "包括",
                exclude: "不包括",
                earlierThan: "早於",
                noEarlierThan: "不早於",
                laterThan: "晚於",
                noLaterThan: "不晚於",
                identificationNumber: "身份證號碼",
                phoneNumber: "手機號",
            },
            optionLabel_es: {
                number: "Número",
                number_integer: "Número entero",
                number_decimal: "Número decimal",
                between: "Entre",
                notBetween: "No entre",
                equal: "Iqual",
                notEqualTo: "No iqual a",
                moreThanThe: "Más que el",
                lessThan: "Menos que",
                greaterOrEqualTo: "Mayor o igual a",
                lessThanOrEqualTo: "Menor o igual a",
                include: "Incluir",
                exclude: "Excluir",
                earlierThan: "Antes de",
                noEarlierThan: "No antes de",
                laterThan: "Después de",
                noLaterThan: "No después de",
                identificationNumber: "Número de identificación",
                phoneNumber: "Número de teléfono",
            },
            dataRegulation: {
                type: "",
                type2: "",
                rangeTxt: "",
                value1: "",
                value2: "",
                validity: "",
                remote: false,
                prohibitInput: false,
                hintShow: false,
                hintValue: "",
            },
        },
        dataVerificationDropDownList: false,
        conditionRules: {
            rulesType: "",
            rulesValue: "",
            textColor: { check: true, color: "#000000" },
            cellColor: { check: true, color: "#000000" },
            betweenValue: { value1: "", value2: "" },
            dateValue: "",
            repeatValue: "0",
            projectValue: "10",
        },
        visibledatarow: [],
        visibledatacolumn: [],
        ch_width: 0,
        rh_height: 0,
        cellmainWidth: 0,
        cellmainHeight: 0,
        toolbarHeight: 41,
        infobarHeight: 57,
        calculatebarHeight: 29,
        rowHeaderWidth: 46,
        columnHeaderHeight: 20,
        cellMainSrollBarSize: 12,
        sheetBarHeight: 31,
        statisticBarHeight: 23,
        luckysheetTableContentHW: [0, 0],
        defaultcollen: 73,
        defaultrowlen: 19,
        scrollLeft: 0,
        scrollTop: 0,
        sheetScrollRecord: {},
        luckysheet_select_status: false,
        luckysheet_select_save: undefined,
        luckysheet_selection_range: [],
        formulaRangeHighlight: [],
        formulaRangeSelect: undefined,
        functionCandidates: [],
        functionHint: null,
        luckysheet_copy_save: undefined,
        luckysheet_paste_iscut: false,
        filterchage: true,
        filter: {},
        luckysheet_sheet_move_status: false,
        luckysheet_sheet_move_data: [],
        luckysheet_scroll_status: false,
        luckysheetcurrentisPivotTable: false,
        luckysheet_rows_selected_status: false,
        luckysheet_cols_selected_status: false,
        luckysheet_rows_change_size: false,
        luckysheet_rows_change_size_start: [],
        luckysheet_cols_change_size: false,
        luckysheet_cols_change_size_start: [],
        luckysheet_cols_freeze_drag: false,
        luckysheet_rows_freeze_drag: false,
        luckysheetCellUpdate: [],
        luckysheet_shiftkeydown: false,
        luckysheet_shiftpositon: undefined,
        iscopyself: true,
        activeImg: undefined,
        orderbyindex: 0,
        luckysheet_model_move_state: false,
        luckysheet_model_xy: [0, 0],
        luckysheet_model_move_obj: null,
        luckysheet_cell_selected_move: false,
        luckysheet_cell_selected_move_index: [],
        luckysheet_cell_selected_extend: false,
        luckysheet_cell_selected_extend_index: [],
        lang: null,
        chart_selection: {},
        zoomRatio: 1,
        showGridLines: true,
        allowEdit: true,
        fontList: [],
        defaultFontSize: 10,
        luckysheetPaintModelOn: false,
        luckysheetPaintSingle: false,
        // 默认单元格
        defaultCell: {
            bl: 0,
            ct: { fa: "General", t: "n" },
            fc: "rgb(51, 51, 51)",
            ff: 0,
            fs: 11,
            ht: 1,
            it: 0,
            vt: 1,
            m: "",
            v: "",
        },
        groupValuesRefreshData: [],
        formulaCache: new FormulaCache(),
        hooks: {},
        getRefs: function () { return refs; },
    };
}
export function getFlowdata(ctx, id) {
    var _a, _b;
    if (!ctx)
        return null;
    var i = getSheetIndex(ctx, id || ctx.currentSheetId);
    if (_.isNil(i)) {
        return null;
    }
    return (_b = (_a = ctx.luckysheetfile) === null || _a === void 0 ? void 0 : _a[i]) === null || _b === void 0 ? void 0 : _b.data;
}
function calcRowColSize(ctx, rowCount, colCount) {
    var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l;
    ctx.visibledatarow = [];
    ctx.rh_height = 0;
    for (var r = 0; r < rowCount; r += 1) {
        var rowlen = ctx.defaultrowlen;
        if ((_a = ctx.config.rowlen) === null || _a === void 0 ? void 0 : _a[r]) {
            rowlen = (_c = (_b = ctx.config) === null || _b === void 0 ? void 0 : _b.rowlen) === null || _c === void 0 ? void 0 : _c[r];
        }
        if (((_e = (_d = ctx.config) === null || _d === void 0 ? void 0 : _d.rowhidden) === null || _e === void 0 ? void 0 : _e[r]) != null) {
            ctx.visibledatarow.push(ctx.rh_height);
            continue;
        }
        // 自动行高计算
        // if (rowlen === "auto") {
        //   rowlen = computeRowlenByContent(ctx.flowdata, r);
        // }
        ctx.rh_height += Math.round((rowlen + 1) * ctx.zoomRatio);
        ctx.visibledatarow.push(ctx.rh_height); // 行的临时长度分布
    }
    // 如果增加行和回到顶部按钮隐藏，则减少底部空白区域，但是预留足够空间给单元格下拉按钮
    // if (
    //   !luckysheetConfigsetting.enableAddRow &&
    //   !luckysheetConfigsetting.enableAddBackTop
    // ) {
    //   ctx.rh_height += 29;
    // } else {
    // }
    ctx.rh_height += 80; // 最底部增加空白
    ctx.visibledatacolumn = [];
    ctx.ch_width = 0;
    var maxColumnlen = 120;
    var flowdata = getFlowdata(ctx);
    for (var c = 0; c < colCount; c += 1) {
        var firstcolumnlen = ctx.defaultcollen;
        if ((_g = (_f = ctx.config) === null || _f === void 0 ? void 0 : _f.columnlen) === null || _g === void 0 ? void 0 : _g[c]) {
            firstcolumnlen = ctx.config.columnlen[c];
        }
        else {
            if ((_h = flowdata === null || flowdata === void 0 ? void 0 : flowdata[0]) === null || _h === void 0 ? void 0 : _h[c]) {
                if (firstcolumnlen > 300) {
                    firstcolumnlen = 300;
                }
                else if (firstcolumnlen < ctx.defaultcollen) {
                    firstcolumnlen = ctx.defaultcollen;
                }
                if (firstcolumnlen !== ctx.defaultcollen) {
                    if (!((_j = ctx.config) === null || _j === void 0 ? void 0 : _j.columnlen)) {
                        ctx.config.columnlen = {};
                    }
                    ctx.config.columnlen[c] = firstcolumnlen;
                }
            }
        }
        if (((_l = (_k = ctx.config) === null || _k === void 0 ? void 0 : _k.colhidden) === null || _l === void 0 ? void 0 : _l[c]) != null) {
            ctx.visibledatacolumn.push(ctx.ch_width);
            continue;
        }
        // 自动行高计算
        // if (firstcolumnlen === "auto") {
        //   firstcolumnlen = computeColWidthByContent(
        //     ctx.flowdata,
        //     c,
        //     rowCount
        //   );
        // }
        ctx.ch_width += Math.round((firstcolumnlen + 1) * ctx.zoomRatio);
        ctx.visibledatacolumn.push(ctx.ch_width); // 列的临时长度分布
    }
    ctx.ch_width += maxColumnlen;
}
export function ensureSheetIndex(data, generateSheetId) {
    if ((data === null || data === void 0 ? void 0 : data.length) > 0) {
        var hasActive_1 = false;
        var indexs_1 = [];
        data.forEach(function (item) {
            if (item.id == null) {
                item.id = generateSheetId();
            }
            if (indexs_1.includes(item.id)) {
                item.id = generateSheetId();
            }
            else {
                indexs_1.push(item.id);
            }
            if (item.status == null) {
                item.status = 0;
            }
            if (item.status === 1) {
                if (hasActive_1) {
                    item.status = 0;
                }
                else {
                    hasActive_1 = true;
                }
            }
        });
        if (!hasActive_1) {
            data[0].status = 1;
        }
    }
}
export function initSheetIndex(ctx) {
    // get current sheet
    var shownSheets = ctx.luckysheetfile.filter(function (singleSheet) { return _.isUndefined(singleSheet.hide) || singleSheet.hide !== 1; });
    ctx.currentSheetId = _.sortBy(shownSheets, function (sheet) { return sheet.order; })[0]
        .id;
    for (var i = 0; i < ctx.luckysheetfile.length; i += 1) {
        if (ctx.luckysheetfile[i].status === 1 &&
            ctx.luckysheetfile[i].hide !== 1) {
            ctx.currentSheetId = ctx.luckysheetfile[i].id;
            break;
        }
    }
}
export function updateContextWithSheetData(ctx, data) {
    var rowCount = data.length;
    var colCount = rowCount === 0 ? 0 : data[0].length;
    calcRowColSize(ctx, rowCount, colCount);
    normalizeSelection(ctx, ctx.luckysheet_select_save);
}
export function updateContextWithCanvas(ctx, canvas, placeholder) {
    ctx.luckysheetTableContentHW = [
        placeholder.clientWidth,
        placeholder.clientHeight,
    ];
    ctx.cellmainHeight = placeholder.clientHeight - ctx.columnHeaderHeight;
    ctx.cellmainWidth = placeholder.clientWidth - ctx.rowHeaderWidth;
    canvas.style.width = "".concat(ctx.luckysheetTableContentHW[0], "px");
    canvas.style.height = "".concat(ctx.luckysheetTableContentHW[1], "px");
    canvas.width = Math.ceil(ctx.luckysheetTableContentHW[0] * ctx.devicePixelRatio);
    canvas.height = Math.ceil(ctx.luckysheetTableContentHW[1] * ctx.devicePixelRatio);
}
