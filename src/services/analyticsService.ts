// import mysql from "mysql";
import { getCategories, getCustomerInvoiceLinesReport, getDevoluciones, getPaymentMethod } from "../models/analyticsModel";
import { getExcelColumnLetter, CustomError } from "../tools";
import { ReportOrderFilters, RowReportOrder } from "../types";
import * as ExcelJS from "exceljs";

/* --- Estilos --- */
const HORIZONTAL_CENTER_ALIGNMENT: Partial<ExcelJS.Alignment> = {
    horizontal: 'center'
};
const HORIZONTAL_RIGHT_ALIGNMENT: Partial<ExcelJS.Alignment> = {
    horizontal: 'right'
};
const CURRENCY_FORMAT = '"$"#,##0.00;[Red]\-"$"#,##0.00';
const PERCENT_FORMAT = '0.00%;[Red]\-0.00%';
const STYLE_FILTER_TEXT_TITLE: Partial<ExcelJS.Style> = {
    font: {
        bold: true,
        color: { argb: '00676D' }
    },
    alignment: {
        horizontal: 'left',
        vertical: 'middle',
    },
    fill: {
        type: 'pattern',
        pattern:'solid',
        fgColor: { argb: 'FFE6FFF2' }
    }
};
const STYLE_FILTER_TEXT: Partial<ExcelJS.Style> = {
    font: {
        italic: true,
        color: { argb: '00676D' }
    },
    alignment: {
        horizontal: 'left',
        vertical: 'top',
        wrapText: true,
        indent: 1
    },
    fill: {
        type: 'pattern',
        pattern:'solid',
        fgColor: { argb: 'FFE6FFF2' }
    }
};
const STYLE_HEAD: Partial<ExcelJS.Style> = {
    font: {
        bold: true,
        color: { argb: 'FFFFFF' }
    },
    alignment: {
        horizontal: 'center',
        vertical: 'middle'
    },
    border: {
        bottom: {
            style: 'thick',
            color: { argb: 'FFD240' }
        }
    },
    fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF00676D' }
    }
};
const STYLE_ROW_STRIPED: Partial<ExcelJS.Style> = {
    fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF3F3F3' }
    }
};
const STYLE_FOOT: Partial<ExcelJS.Style> = {
    font: {
        bold: true
    },
    alignment: {
        vertical: 'middle'
    },
    border: {
        top: {
            style: 'thin',
            color: { argb: 'FF666666' }
        },
        bottom: {
            style: 'thin',
            color: { argb: 'FF666666' }
        }
    },
    fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFCCCCCC' }
    }
};
const STYLE_SOFT_VERSION: Partial<ExcelJS.Style> = {
    font: {
        italic: true,
        color: { argb: '888888' },
        size: 8
    }
};

const validateCustomerInvoiceFilters = (filters: ReportOrderFilters): [ReportOrderFilters, number, number, string] => {
    if (!filters.orderStatusId) {
        let invoiceStatusId = [3,5,6,7,8];
        if (filters.showCanceled) {
            invoiceStatusId.push(4);
        }

        filters.orderStatusId = invoiceStatusId.join(',');
    }

    if (!filters.reportType) {
        filters.reportType = 'customer';
    }

    let group = 0;
    let sort = 0;
    let order = '';
    switch(filters.reportType) {
        case 'customer':
            group = 0;
            sort = 0;
            break;
        case 'transaction_sequence':
            group = 2;
            sort = 1;
            break;
        case 'detail':
            group = 1;
            sort = 1;
            break;
        case 'branch':
            group = 3;
            sort = 1;
            break;
        case 'box':
            group = 4;
            sort = 2;
            break;
        case 'product':
            group = 5;
            sort = 1;
            break;
        case 'category':
            group = 6;
            sort = 1;
            break;
        default:
            group = 1;
            sort = 1;
            break;
    }
    
    if (filters.sort) {
        sort = Number(filters.sort);
    }

    if (!filters.order) {
        order = 'ASC';
    }

    return [filters, group, sort, order];
}

/**
 * Consulta las ordenes, y genera un archivo excel de los registros obtenidos.
 * - Nota: Para optimizar los tiempos de ejecución, se utilizan los métodos de iteración de 
 * {@link Array} ({@link Array.forEach}, {@link Array.map}, etc)
 * en lugar de los ciclos convencionales (for, while, etc).
 * @param filtersUn Filtros para la consulta {@link ReportOrderFilters}
 * @returns El objeto xlsx del libro de trabajo {@link ExcelJS.Xlsx}
 */
const generateReportOrderExcel = async (filtersUn: ReportOrderFilters): Promise<ExcelJS.Xlsx> => {
    return new Promise(async(resolve, reject) => {
        const [filters, group, sort, order] = validateCustomerInvoiceFilters(filtersUn);
        const orders: RowReportOrder[] = await getCustomerInvoiceLinesReport(filters, group, sort, order);

        if (orders.length === 0 && filters.reportType !== 'category') {
            reject(new CustomError(404, 'No se encontraron resultados'));
        }
        
        /* --- Se crea un nuevo libro de Excel --- */
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Sheet 1');

        /* --- Textos de Filtros --- */
        let filterText = '';
        filterText +=  filters.name != null ? `${filterText.length > 0 ? '\n' : ''} Número: ${filters.name}` : '';
        filterText +=  filters.dateFrom != null || filters.dateTo ? `${filterText.length > 0 ? '\n' : ''} Fecha: Desde ${filters.dateFrom != null ? filters.dateFrom : '--'} hasta ${filters.dateTo != null ? filters.dateTo : '--'}` : '';
        filterText +=  filters.origin != null ? `${filterText.length > 0 ? '\n' : ''} Doc. Origen: ${filters.origin}` : '';
        filterText +=  filters.customer != null ? `${filterText.length > 0 ? '\n' : ''} Cliente: ${filters.customer}` : '';
        filterText +=  filters.salesperson != null ? `${filterText.length > 0 ? '\n' : ''} Vendedor: ${filters.salesperson}` : '';
        filterText +=  filters.productCategoryId != null ? `${filterText.length > 0 ? '\n' : ''} Categoria de producto: ${filters.productCategoryId}` : '';
        filterText +=  filters.showCanceled != null ? `${filterText.length > 0 ? '\n' : ''} Mostrar cancelados: ${filters.showCanceled}` : '';
        filterText +=  `${filterText.length > 0 ? '\n' : ''} Fecha de creación: ${new Date().toLocaleDateString('es-mx')}`;

        const reportType = filters.reportType;
        /* Reporte agrupado por cliente */
        if (reportType === 'customer') {
            /* --- Inicio --- */
            let rowIni = 1;
            const colEndNumber = 7;
            const colEnd = getExcelColumnLetter(colEndNumber);
            let cell = worksheet.getCell(`A${rowIni}`);
            worksheet.mergeCells(`A${rowIni}:${colEnd}${rowIni}`);
            cell.value = 'Reporte pedidos de venta';
            cell.style = STYLE_FILTER_TEXT_TITLE;
    
            /* --- Encabezado --- */
            rowIni++;
            cell = worksheet.getCell(`A${rowIni}`);
            worksheet.mergeCells(`A${rowIni}:${colEnd}${rowIni}`);
            cell.value = filterText;
            cell.style = STYLE_FILTER_TEXT;
            worksheet.getRow(rowIni).height = (filterText.split('\n').length * 13) + 6;
    
            /* --- Ancho de columnas --- */
            worksheet.getColumn('A').width = 35;
            worksheet.getColumn('B').width = 20;
            worksheet.getColumn('C').width = 15;
            worksheet.getColumn('D').width = 15;
            worksheet.getColumn('E').width = 15;
            worksheet.getColumn('F').width = 15;
            worksheet.getColumn('G').width = 15;
    
            /* --- Títulos de las columnas --- */
            const titleRowData: string[] = [
                'Cliente',
                'Moneda',
                'Subtotal',
                'Impuestos',
                'Total',
                'Margen',
                '% Margen',
            ];
    
            const titleRow = worksheet.addRow(titleRowData);
            titleRow.height = 30;
            titleRow.eachCell((cell) => {
                cell.style = STYLE_HEAD;
            });
            rowIni++;
    
            /* --- Registros --- */
            let row = (rowIni);
            console.time('orders')
            const promises = orders.map(async(result) => {
                const configCurrency = filters.configCurrencyCode;
    
                /* --- Valores de Columnas --- */
                const bodyRowData: (string | number | Date)[] = [
                    result.customer,
                    configCurrency,
                    result.total_amount_untaxed,
                    result.total_amount_tax_total,
                    result.total_amount_total,
                    result.total_amount_margin,
                    result.total_percent_margin,
                ];
    
                return bodyRowData;
            });
            
            Promise.all(promises).then((rows) => {
                console.timeEnd('customerInvoices');
                const rowsU = rows.filter(row => row);
                if (rowsU.length > 0) {
                    const createdRows = worksheet.addRows(rowsU);
                    createdRows.forEach(row => {
                        /* Estilo de celdas */
                        row.eachCell((cell) => {
                            const col = +cell.col;
                            if ([2].includes(col)) {
                                cell.alignment = HORIZONTAL_CENTER_ALIGNMENT;
                            } else if (col > 2 && col < colEndNumber) {
                                cell.alignment = HORIZONTAL_RIGHT_ALIGNMENT;
                                cell.numFmt = CURRENCY_FORMAT;
                            } else if (col == colEndNumber) {
                                cell.numFmt = PERCENT_FORMAT;
                            }
    
                            if (col === 1) {
                                cell.border = {
                                    left: {
                                        style: 'thin',
                                        color: { argb: 'FF666666' }
                                    }
                                };
                            } else if (col === colEndNumber) {
                                cell.border = {
                                    right: {
                                        style: 'thin',
                                        color: { argb: 'FF666666' }
                                    }
                                };
                            }
    
                            if (+cell.row % 2) {
                                cell.style = Object.assign(cell.style, STYLE_ROW_STRIPED);
                            }
                        });
                    });
                }
                /* --- Totales --- */
                const rowCount = worksheet.rowCount;
                const totalRowData: (string | ExcelJS.CellValue)[] = [
                    '',
                    '',
                    {
                        formula: `=SUM(C${rowIni + 1}:C${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(D${rowIni + 1}:D${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(E${rowIni + 1}:E${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(F${rowIni + 1}:F${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=(F${rowIni + 1}/C${rowCount})`,
                        date1904: false,
                    },
                ];

                const totalRow = worksheet.addRow(totalRowData);
    
                /* Estilo de Totales */
                totalRow.eachCell((cell) => {
                    const col = getExcelColumnLetter(+cell.col);
                    cell.style = Object.assign(cell.style, STYLE_FOOT);
                    if (cell.type != ExcelJS.ValueType.Null) {
                        if ((+cell.col) === (colEndNumber)) {
                            cell.numFmt = PERCENT_FORMAT;
                        } else {
                            cell.numFmt = CURRENCY_FORMAT;
                        }
                    }
    
                    if (col === 'A') {
                        cell.border = Object.assign({
                            left: {
                                style: 'thin',
                                color: { argb: 'FF666666' }
                            }
                        }, STYLE_FOOT.border);
                    } else if (col === colEnd) {
                        cell.border = Object.assign({
                            right: {
                                style: 'thin',
                                color: { argb: 'FF666666' }
                            }
                        }, STYLE_FOOT.border);
                    }
                });

                row += rowCount;
                cell = worksheet.getCell('A' + row);
                cell.value = (SOFT_NAME + ' v.' + VERSION);
                cell.style = STYLE_SOFT_VERSION;
                
                resolve(workbook.xlsx);
            });
        } else if (reportType === 'transaction_sequence') {
            /* --- Inicio --- */
            let rowIni = 1;
            const colEndNumber = 16;
            const colEnd = getExcelColumnLetter(colEndNumber);
            let cell = worksheet.getCell(`A${rowIni}`);
            worksheet.mergeCells(`A${rowIni}:${colEnd}${rowIni}`);
            cell.value = 'Reporte pedidos de venta';
            cell.style = STYLE_FILTER_TEXT_TITLE;
    
            /* --- Encabezado --- */
            rowIni++;
            cell = worksheet.getCell(`A${rowIni}`);
            worksheet.mergeCells(`A${rowIni}:${colEnd}${rowIni}`);
            cell.value = filterText;
            cell.style = STYLE_FILTER_TEXT;
            worksheet.getRow(rowIni).height = (filterText.split('\n').length * 13) + 6;
    
            /* --- Ancho de columnas --- */
            worksheet.getColumn('A').width = 35;
            worksheet.getColumn('B').width = 13;
            worksheet.getColumn('C').width = 13;
            worksheet.getColumn('D').width = 20;
            worksheet.getColumn('E').width = 35;
            worksheet.getColumn('F').width = 25;
            worksheet.getColumn('G').width = 14;
            worksheet.getColumn('H').width = 14;
            worksheet.getColumn('I').width = 14;
            worksheet.getColumn('J').width = 14;
            worksheet.getColumn('K').width = 14;
            worksheet.getColumn('L').width = 15;
            worksheet.getColumn('M').width = 15;
            worksheet.getColumn('N').width = 15;
            worksheet.getColumn('O').width = 15;
            worksheet.getColumn('P').width = 25;
    
            /* --- Títulos de las columnas --- */
            const titleRowData: string[] = [
                'Fecha',
                'Caja',
                'Sucursal',
                'Número',
                'Doc. Origen',
                'Referencia',
                'Cliente',
                'Vendedor',
                'Moneda',
                'Subtotal',
                'Impuestos',
                'Total',
                'Margen',
                '% Margen',
                'Formato de pago',
                'Estatus',
            ];
    
            const titleRow = worksheet.addRow(titleRowData);
            titleRow.height = 30;
            titleRow.eachCell((cell) => {
                cell.style = STYLE_HEAD;
            });
            rowIni++;
    
            /* --- Registros --- */
            let row = (rowIni);
            console.time('orders')
            const promises = orders.map(async(result) => {
                const configCurrency = filters.configCurrencyCode;
    
                /* --- Valores de Columnas --- */
                const paymentMethods = await getPaymentMethod(result.payment_method);
                const bodyRowData: (string | number | Date)[] = [
                    result.quotation_date,
                    result.shift_description,
                    result.name_branch,
                    result.order_name,
                    result.origin,
                    result.customer_order_ref,
                    result.customer,
                    result.salesperson,
                    configCurrency,
                    result.total_amount_untaxed,
                    result.total_amount_tax_total,
                    result.total_amount_total,
                    result.total_amount_margin,
                    result.total_percent_margin,
                    paymentMethods ? paymentMethods : '',
                    result.order_status_id === '4' ? 'Cancelado' : 'Activo',
                ];
    
                return bodyRowData;
            });
            
            Promise.all(promises).then((rows) => {
                console.timeEnd('customerInvoices');
                const rowsU = rows.filter(row => row);
                if (rowsU.length > 0) {
                    const createdRows = worksheet.addRows(rowsU);
                    createdRows.forEach(row => {
                        /* Estilo de celdas */
                        row.eachCell((cell) => {
                            const col = +cell.col;
                            if ([1, 2, 3, 4, 5, 6, 9, 15, colEndNumber].includes(col)) {
                                cell.alignment = HORIZONTAL_CENTER_ALIGNMENT;
                            } else if (col >= 10 && col <= 13) {
                                cell.alignment = HORIZONTAL_RIGHT_ALIGNMENT;
                                cell.numFmt = CURRENCY_FORMAT;
                            } else if (col === 14) {
                                cell.numFmt = PERCENT_FORMAT;
                            }
    
                            if (col === 1) {
                                cell.border = {
                                    left: {
                                        style: 'thin',
                                        color: { argb: 'FF666666' }
                                    }
                                };
                            } else if (col === colEndNumber) {
                                cell.border = {
                                    right: {
                                        style: 'thin',
                                        color: { argb: 'FF666666' }
                                    }
                                };
                            }
    
                            if (+cell.row % 2) {
                                cell.style = Object.assign(cell.style, STYLE_ROW_STRIPED);
                            }
                        });
                    });
                }
                /* --- Totales --- */
                const rowCount = worksheet.rowCount;
                const totalRowData: (string | ExcelJS.CellValue)[] = [
                    '',
                    '',
                    '',
                    '',
                    '',
                    '',
                    '',
                    '',
                    '',
                    {
                        formula: `=SUM(H${rowIni + 1}:H${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(I${rowIni + 1}:I${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(J${rowIni + 1}:J${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(K${rowIni + 1}:K${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=(K${rowIni + 1}/H${rowCount})`,
                        date1904: false,
                    },
                    '',
                ];

                const totalRow = worksheet.addRow(totalRowData);
    
                /* Estilo de Totales */
                totalRow.eachCell((cell) => {
                    const col = getExcelColumnLetter(+cell.col);
                    cell.style = Object.assign(cell.style, STYLE_FOOT);
                    if (cell.type != ExcelJS.ValueType.Null) {
                        if ((+cell.col) === (14)) {
                            cell.numFmt = PERCENT_FORMAT;
                        } else {
                            cell.numFmt = CURRENCY_FORMAT;
                        }
                    }
    
                    if (col === 'A') {
                        cell.border = Object.assign({
                            left: {
                                style: 'thin',
                                color: { argb: 'FF666666' }
                            }
                        }, STYLE_FOOT.border);
                    } else if (col === colEnd) {
                        cell.border = Object.assign({
                            right: {
                                style: 'thin',
                                color: { argb: 'FF666666' }
                            }
                        }, STYLE_FOOT.border);
                    }
                });

                row += rowCount;
                cell = worksheet.getCell('A' + row);
                cell.value = (SOFT_NAME + ' v.' + VERSION);
                cell.style = STYLE_SOFT_VERSION;
                
                resolve(workbook.xlsx);
            });
        } else if (reportType === 'detail') {
            /* --- Inicio --- */
            let rowIni = 1;
            const colEndNumber = 17;
            const colEnd = getExcelColumnLetter(colEndNumber);
            let cell = worksheet.getCell(`A${rowIni}`);
            worksheet.mergeCells(`A${rowIni}:${colEnd}${rowIni}`);
            cell.value = 'Reporte pedidos de venta';
            cell.style = STYLE_FILTER_TEXT_TITLE;
    
            /* --- Encabezado --- */
            rowIni++;
            cell = worksheet.getCell(`A${rowIni}`);
            worksheet.mergeCells(`A${rowIni}:${colEnd}${rowIni}`);
            cell.value = filterText;
            cell.style = STYLE_FILTER_TEXT;
            worksheet.getRow(rowIni).height = (filterText.split('\n').length * 13) + 6;
    
            /* --- Ancho de columnas --- */
            worksheet.getColumn('A').width = 20;
            worksheet.getColumn('B').width = 35;
            worksheet.getColumn('C').width = 15;
            worksheet.getColumn('D').width = 35;
            worksheet.getColumn('E').width = 15;
            worksheet.getColumn('F').width = 14;
            worksheet.getColumn('G').width = 14;
            worksheet.getColumn('H').width = 14;
            worksheet.getColumn('I').width = 14;
            worksheet.getColumn('J').width = 14;
            worksheet.getColumn('K').width = 14;
            worksheet.getColumn('L').width = 14;
            worksheet.getColumn('M').width = 14;
            worksheet.getColumn('N').width = 14;
            worksheet.getColumn('O').width = 14;
            worksheet.getColumn('P').width = 15;
            worksheet.getColumn('Q').width = 15;
    
            /* --- Títulos de las columnas --- */
            const titleRowData: string[] = [
                'Fecha',
                'Caja',
                'Sucursal',
                'Número',
                'Cliente',
                'Código',
                'Producto',
                'Categoría',
                'Moneda',
                'Cantidad',
                'Subtotal',
                'Impuestos',
                'Total',
                'Costo',
                'Margen',
                '% Margen',
                'Estatus',
            ];
    
            const titleRow = worksheet.addRow(titleRowData);
            titleRow.height = 30;
            titleRow.eachCell((cell) => {
                cell.style = STYLE_HEAD;
            });
            rowIni++;
    
            /* --- Registros --- */
            let row = (rowIni);
            console.time('orders')
            const promises = orders.map(async(result) => {
                const configCurrency = filters.configCurrencyCode;
    
                /* --- Valores de Columnas --- */
                const bodyRowData: (string | number | Date)[] = [
                    result.quotation_date,
                    result.shift_description,
                    result.name_branch,
                    result.order_name,
                    result.customer,
                    result.default_code,
                    result.name,
                    result.product_category,
                    configCurrency,
                    result.total_quantity,
                    result.total_amount_untaxed,
                    result.total_amount_tax_total,
                    result.total_amount_total,
                    result.total_amount_cost,
                    result.total_amount_margin,
                    result.total_percent_margin,
                    result.order_status_id === '4' ? 'Cancelado' : 'Activo',
                ];
    
                return bodyRowData;
            });
            
            Promise.all(promises).then((rows) => {
                console.timeEnd('customerInvoices');
                const rowsU = rows.filter(row => row);
                if (rowsU.length > 0) {
                    const createdRows = worksheet.addRows(rowsU);
                    createdRows.forEach(row => {
                        /* Estilo de celdas */
                        row.eachCell((cell) => {
                            const col = +cell.col;
                            if ([1, 2, 3, 4, 6, 8, 9, colEndNumber].includes(col)) {
                                cell.alignment = HORIZONTAL_CENTER_ALIGNMENT;
                            } else if (col >= 10 && col <= 15) {
                                cell.alignment = HORIZONTAL_RIGHT_ALIGNMENT;
                                cell.numFmt = CURRENCY_FORMAT;
                            } else if (col === 16) {
                                cell.numFmt = PERCENT_FORMAT;
                            }
    
                            if (col === 1) {
                                cell.border = {
                                    left: {
                                        style: 'thin',
                                        color: { argb: 'FF666666' }
                                    }
                                };
                            } else if (col === colEndNumber) {
                                cell.border = {
                                    right: {
                                        style: 'thin',
                                        color: { argb: 'FF666666' }
                                    }
                                };
                            }
    
                            if (+cell.row % 2) {
                                cell.style = Object.assign(cell.style, STYLE_ROW_STRIPED);
                            }
                        });
                    });
                }
                /* --- Totales --- */
                const rowCount = worksheet.rowCount;
                const totalRowData: (string | ExcelJS.CellValue)[] = [
                    '',
                    '',
                    '',
                    '',
                    '',
                    '',
                    '',
                    '',
                    '',
                    {
                        formula: `=SUM(H${rowIni + 1}:H${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(I${rowIni + 1}:I${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(J${rowIni + 1}:J${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(K${rowIni + 1}:K${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(L${rowIni + 1}:L${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(M${rowIni + 1}:M${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=(M${rowIni + 1}/I${rowCount})`,
                        date1904: false,
                    },
                    '',
                ];

                const totalRow = worksheet.addRow(totalRowData);
    
                /* Estilo de Totales */
                totalRow.eachCell((cell) => {
                    const col = getExcelColumnLetter(+cell.col);
                    cell.style = Object.assign(cell.style, STYLE_FOOT);
                    if (cell.type != ExcelJS.ValueType.Null) {
                        if ((+cell.col) === (16)) {
                            cell.numFmt = PERCENT_FORMAT;
                        } else if ((+cell.col) !== 10) {
                            cell.numFmt = CURRENCY_FORMAT;
                        }
                    }
    
                    if (col === 'A') {
                        cell.border = Object.assign({
                            left: {
                                style: 'thin',
                                color: { argb: 'FF666666' }
                            }
                        }, STYLE_FOOT.border);
                    } else if (col === colEnd) {
                        cell.border = Object.assign({
                            right: {
                                style: 'thin',
                                color: { argb: 'FF666666' }
                            }
                        }, STYLE_FOOT.border);
                    }
                });

                row += rowCount;
                cell = worksheet.getCell('A' + row);
                cell.value = (SOFT_NAME + ' v.' + VERSION);
                cell.style = STYLE_SOFT_VERSION;
                
                resolve(workbook.xlsx);
            });
        } else if (reportType === 'branch') {
            /* --- Inicio --- */
            let rowIni = 1;
            const colEndNumber = 7;
            const colEnd = getExcelColumnLetter(colEndNumber);
            let cell = worksheet.getCell(`A${rowIni}`);
            worksheet.mergeCells(`A${rowIni}:${colEnd}${rowIni}`);
            cell.value = 'Reporte pedidos de venta';
            cell.style = STYLE_FILTER_TEXT_TITLE;
    
            /* --- Encabezado --- */
            rowIni++;
            cell = worksheet.getCell(`A${rowIni}`);
            worksheet.mergeCells(`A${rowIni}:${colEnd}${rowIni}`);
            cell.value = filterText;
            cell.style = STYLE_FILTER_TEXT;
            worksheet.getRow(rowIni).height = (filterText.split('\n').length * 13) + 6;
    
            /* --- Ancho de columnas --- */
            worksheet.getColumn('A').width = 35;
            worksheet.getColumn('B').width = 20;
            worksheet.getColumn('C').width = 15;
            worksheet.getColumn('D').width = 15;
            worksheet.getColumn('E').width = 15;
            worksheet.getColumn('F').width = 15;
            worksheet.getColumn('G').width = 15;
    
            /* --- Títulos de las columnas --- */
            const titleRowData: string[] = [
                'Sucursal',
                'No. Ventas',
                'Venta promedio',
                'Artículos',
                'Total',
                'Margen',
                '% Margen',
            ];
    
            const titleRow = worksheet.addRow(titleRowData);
            titleRow.height = 30;
            titleRow.eachCell((cell) => {
                cell.style = STYLE_HEAD;
            });
            rowIni++;
    
            /* --- Registros --- */
            let row = (rowIni);
            console.time('orders')
            const promises = orders.map(async(result) => {
    
                /* --- Valores de Columnas --- */
                const bodyRowData: (string | number | Date)[] = [
                    result.name_branch,
                    result.total_order_count,
                    result.average,
                    result.quantity,
                    result.total_amount_total,
                    result.total_amount_margin,
                    result.total_percent_margin,
                ];
    
                return bodyRowData;
            });
            
            Promise.all(promises).then((rows) => {
                console.timeEnd('customerInvoices');
                const rowsU = rows.filter(row => row);
                if (rowsU.length > 0) {
                    const createdRows = worksheet.addRows(rowsU);
                    createdRows.forEach(row => {
                        /* Estilo de celdas */
                        row.eachCell((cell) => {
                            const col = +cell.col;
                            if ([2, 4, 7].includes(col)) {
                                cell.alignment = HORIZONTAL_CENTER_ALIGNMENT;
                            } else if (col >= 5 && col <= 6) {
                                cell.alignment = HORIZONTAL_RIGHT_ALIGNMENT;
                                cell.numFmt = CURRENCY_FORMAT;
                            } else if (col === colEndNumber) {
                                cell.numFmt = PERCENT_FORMAT;
                            }
    
                            if (col === 1) {
                                cell.border = {
                                    left: {
                                        style: 'thin',
                                        color: { argb: 'FF666666' }
                                    }
                                };
                            } else if (col === colEndNumber) {
                                cell.border = {
                                    right: {
                                        style: 'thin',
                                        color: { argb: 'FF666666' }
                                    }
                                };
                            }
    
                            if (+cell.row % 2) {
                                cell.style = Object.assign(cell.style, STYLE_ROW_STRIPED);
                            }
                        });
                    });
                }
                /* --- Totales --- */
                const rowCount = worksheet.rowCount;
                const totalRowData: (string | ExcelJS.CellValue)[] = [
                    '',
                    {
                        formula: `=SUM(B${rowIni + 1}:B${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(C${rowIni + 1}:C${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(D${rowIni + 1}:D${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(E${rowIni + 1}:E${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(F${rowIni + 1}:F${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=(F${rowIni + 1}/C${rowCount})`,
                        date1904: false,
                    },
                ];

                const totalRow = worksheet.addRow(totalRowData);
    
                /* Estilo de Totales */
                totalRow.eachCell((cell) => {
                    const col = getExcelColumnLetter(+cell.col);
                    cell.style = Object.assign(cell.style, STYLE_FOOT);
                    if (cell.type != ExcelJS.ValueType.Null) {
                        if ((+cell.col) === colEndNumber) {
                            cell.numFmt = PERCENT_FORMAT;
                        } else if (![2, 3, 4].includes(+cell.col)) {
                            cell.numFmt = CURRENCY_FORMAT;
                        }
                    }
    
                    if (col === 'A') {
                        cell.border = Object.assign({
                            left: {
                                style: 'thin',
                                color: { argb: 'FF666666' }
                            }
                        }, STYLE_FOOT.border);
                    } else if (col === colEnd) {
                        cell.border = Object.assign({
                            right: {
                                style: 'thin',
                                color: { argb: 'FF666666' }
                            }
                        }, STYLE_FOOT.border);
                    }
                });

                row += rowCount;
                cell = worksheet.getCell('A' + row);
                cell.value = (SOFT_NAME + ' v.' + VERSION);
                cell.style = STYLE_SOFT_VERSION;
                
                resolve(workbook.xlsx);
            });
        } else if (reportType === 'product') {
            /* --- Inicio --- */
            let rowIni = 1;
            const colEndNumber = 4;
            const colEnd = getExcelColumnLetter(colEndNumber);
            let cell = worksheet.getCell(`A${rowIni}`);
            worksheet.mergeCells(`A${rowIni}:${colEnd}${rowIni}`);
            cell.value = 'Reporte pedidos de venta';
            cell.style = STYLE_FILTER_TEXT_TITLE;
    
            /* --- Encabezado --- */
            rowIni++;
            cell = worksheet.getCell(`A${rowIni}`);
            worksheet.mergeCells(`A${rowIni}:${colEnd}${rowIni}`);
            cell.value = filterText;
            cell.style = STYLE_FILTER_TEXT;
            worksheet.getRow(rowIni).height = (filterText.split('\n').length * 13) + 6;
    
            /* --- Ancho de columnas --- */
            worksheet.getColumn('A').width = 25;
            worksheet.getColumn('B').width = 25;
            worksheet.getColumn('C').width = 25;
            worksheet.getColumn('D').width = 25;
    
            /* --- Títulos de las columnas --- */
            const titleRowData: string[] = [
                'Código',
                'Producto',
                'Cantidad',
                'Importe',
            ];
    
            const titleRow = worksheet.addRow(titleRowData);
            titleRow.height = 30;
            titleRow.eachCell((cell) => {
                cell.style = STYLE_HEAD;
            });
            rowIni++;
    
            /* --- Registros --- */
            let row = (rowIni);
            console.time('orders')
            const promises = orders.map(async(result) => {
    
                /* --- Valores de Columnas --- */
                const bodyRowData: (string | number | Date)[] = [
                    result.default_code,
                    result.name,
                    result.total_quantity,
                    result.amount_total,
                ];
    
                return bodyRowData;
            });
            
            Promise.all(promises).then((rows) => {
                console.timeEnd('customerInvoices');
                const rowsU = rows.filter(row => row);
                if (rowsU.length > 0) {
                    const createdRows = worksheet.addRows(rowsU);
                    createdRows.forEach(row => {
                        /* Estilo de celdas */
                        row.eachCell((cell) => {
                            const col = +cell.col;
                            if ([1].includes(col)) {
                                cell.alignment = HORIZONTAL_CENTER_ALIGNMENT;
                            } else if (col === colEndNumber) {
                                cell.alignment = HORIZONTAL_RIGHT_ALIGNMENT;
                                cell.numFmt = CURRENCY_FORMAT;
                            }
    
                            if (col === 1) {
                                cell.border = {
                                    left: {
                                        style: 'thin',
                                        color: { argb: 'FF666666' }
                                    }
                                };
                            } else if (col === colEndNumber) {
                                cell.border = {
                                    right: {
                                        style: 'thin',
                                        color: { argb: 'FF666666' }
                                    }
                                };
                            }
    
                            if (+cell.row % 2) {
                                cell.style = Object.assign(cell.style, STYLE_ROW_STRIPED);
                            }
                        });
                    });
                }
                /* --- Totales --- */
                const rowCount = worksheet.rowCount;
                const totalRowData: (string | ExcelJS.CellValue)[] = [
                    '',
                    '',
                    {
                        formula: `=SUM(C${rowIni + 1}:C${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(D${rowIni + 1}:D${rowCount})`,
                        date1904: false,
                    },
                ];

                const totalRow = worksheet.addRow(totalRowData);
    
                /* Estilo de Totales */
                totalRow.eachCell((cell) => {
                    const col = getExcelColumnLetter(+cell.col);
                    cell.style = Object.assign(cell.style, STYLE_FOOT);
                    if (cell.type != ExcelJS.ValueType.Null) {
                        if ((+cell.col) === colEndNumber) {
                            cell.numFmt = CURRENCY_FORMAT;
                        }
                    }
    
                    if (col === 'A') {
                        cell.border = Object.assign({
                            left: {
                                style: 'thin',
                                color: { argb: 'FF666666' }
                            }
                        }, STYLE_FOOT.border);
                    } else if (col === colEnd) {
                        cell.border = Object.assign({
                            right: {
                                style: 'thin',
                                color: { argb: 'FF666666' }
                            }
                        }, STYLE_FOOT.border);
                    }
                });

                row += rowCount;
                cell = worksheet.getCell('A' + row);
                cell.value = (SOFT_NAME + ' v.' + VERSION);
                cell.style = STYLE_SOFT_VERSION;
                
                resolve(workbook.xlsx);
            });
        } else if (reportType === 'box') {
            /* --- Inicio --- */
            let rowIni = 1;
            const colEndNumber = 10;
            const colEnd = getExcelColumnLetter(colEndNumber);
            let cell = worksheet.getCell(`A${rowIni}`);
            worksheet.mergeCells(`A${rowIni}:${colEnd}${rowIni}`);
            cell.value = 'Reporte pedidos de venta';
            cell.style = STYLE_FILTER_TEXT_TITLE;
    
            /* --- Encabezado --- */
            rowIni++;
            cell = worksheet.getCell(`A${rowIni}`);
            worksheet.mergeCells(`A${rowIni}:${colEnd}${rowIni}`);
            cell.value = filterText;
            cell.style = STYLE_FILTER_TEXT;
            worksheet.getRow(rowIni).height = (filterText.split('\n').length * 13) + 6;
    
            /* --- Ancho de columnas --- */
            worksheet.getColumn('A').width = 25;
            worksheet.getColumn('B').width = 25;
            worksheet.getColumn('C').width = 25;
            worksheet.getColumn('D').width = 25;
            worksheet.getColumn('E').width = 15;
            worksheet.getColumn('F').width = 25;
            worksheet.getColumn('G').width = 25;
            worksheet.getColumn('H').width = 25;
            worksheet.getColumn('I').width = 25;
            worksheet.getColumn('J').width = 25;
    
            /* --- Títulos de las columnas --- */
            const titleRowData: string[] = [
                'Sucursal',
                'Fecha',
                'Caja',
                'Hora',
                'Cajero',
                'Subtotal',
                'Impuestos',
                'Total',
                'Ventas',
                'Devoluciones',
            ];
    
            const titleRow = worksheet.addRow(titleRowData);
            titleRow.height = 30;
            titleRow.eachCell((cell) => {
                cell.style = STYLE_HEAD;
            });
            rowIni++;
    
            /* --- Registros --- */
            let row = (rowIni);
            console.time('orders')
            const promises = orders.map(async(result) => {
                const devFilters: ReportOrderFilters = { ...filters, boxId: result.id_shift };
                const devoluciones = await getDevoluciones(devFilters);
    
                /* --- Valores de Columnas --- */
                const bodyRowData: (string | number | Date)[] = [
                    result.name_branch,
                    result.date_order,
                    result.shift_description,
                    result.create_date,
                    result.usr_order,
                    result.total_amount_untaxed,
                    result.total_amount_tax,
                    result.sum_total,
                    result.total_order_count,
                    devoluciones,
                ];
    
                return bodyRowData;
            });
            
            Promise.all(promises).then((rows) => {
                console.timeEnd('customerInvoices');
                const rowsU = rows.filter(row => row);
                if (rowsU.length > 0) {
                    const createdRows = worksheet.addRows(rowsU);
                    createdRows.forEach(row => {
                        /* Estilo de celdas */
                        row.eachCell((cell) => {
                            const col = +cell.col;
                            if ([1, 3, 4].includes(col)) {
                                cell.alignment = HORIZONTAL_CENTER_ALIGNMENT;
                            } else if (col >= 6 && col <= 7) {
                                cell.alignment = HORIZONTAL_RIGHT_ALIGNMENT;
                                cell.numFmt = CURRENCY_FORMAT;
                            } else if ([8, 9, 10].includes(col)) {
                                cell.alignment = HORIZONTAL_RIGHT_ALIGNMENT;
                            }
    
                            if (col === 1) {
                                cell.border = {
                                    left: {
                                        style: 'thin',
                                        color: { argb: 'FF666666' }
                                    }
                                };
                            } else if (col === colEndNumber) {
                                cell.border = {
                                    right: {
                                        style: 'thin',
                                        color: { argb: 'FF666666' }
                                    }
                                };
                            }
    
                            if (+cell.row % 2) {
                                cell.style = Object.assign(cell.style, STYLE_ROW_STRIPED);
                            }
                        });
                    });
                }
                /* --- Totales --- */
                const rowCount = worksheet.rowCount;
                const totalRowData: (string | ExcelJS.CellValue)[] = [
                    '',
                    '',
                    '',
                    '',
                    'Total',
                    {
                        formula: `=SUM(F${rowIni + 1}:F${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(G${rowIni + 1}:G${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(H${rowIni + 1}:H${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(I${rowIni + 1}:I${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(J${rowIni + 1}:J${rowCount})`,
                        date1904: false,
                    },
                ];

                const totalRow = worksheet.addRow(totalRowData);
    
                /* Estilo de Totales */
                totalRow.eachCell((cell) => {
                    const col = getExcelColumnLetter(+cell.col);
                    cell.style = Object.assign(cell.style, STYLE_FOOT);
                    if (cell.type != ExcelJS.ValueType.Null) {
                        if ([6, 7].includes(+cell.col)) {
                            cell.numFmt = CURRENCY_FORMAT;
                        }
                    }
    
                    if (col === 'A') {
                        cell.border = Object.assign({
                            left: {
                                style: 'thin',
                                color: { argb: 'FF666666' }
                            }
                        }, STYLE_FOOT.border);
                    } else if (col === colEnd) {
                        cell.border = Object.assign({
                            right: {
                                style: 'thin',
                                color: { argb: 'FF666666' }
                            }
                        }, STYLE_FOOT.border);
                    }
                });

                row += rowCount;
                cell = worksheet.getCell('A' + row);
                cell.value = (SOFT_NAME + ' v.' + VERSION);
                cell.style = STYLE_SOFT_VERSION;
                
                resolve(workbook.xlsx);
            });
        } else if (reportType === 'category') {
            const categories = await getCategories(filters);
            if (categories.length === 0) {
                reject(new CustomError(404, 'No se encontraron resultados'));
            }
            /* --- Inicio --- */
            let rowIni = 1;
            const colEndNumber = 9;
            const colEnd = getExcelColumnLetter(colEndNumber);
            let cell = worksheet.getCell(`A${rowIni}`);
            worksheet.mergeCells(`A${rowIni}:${colEnd}${rowIni}`);
            cell.value = 'Reporte pedidos de venta';
            cell.style = STYLE_FILTER_TEXT_TITLE;
    
            /* --- Encabezado --- */
            rowIni++;
            cell = worksheet.getCell(`A${rowIni}`);
            worksheet.mergeCells(`A${rowIni}:${colEnd}${rowIni}`);
            cell.value = filterText;
            cell.style = STYLE_FILTER_TEXT;
            worksheet.getRow(rowIni).height = (filterText.split('\n').length * 13) + 6;
    
            /* --- Ancho de columnas --- */
            worksheet.getColumn('A').width = 35;
            worksheet.getColumn('B').width = 25;
            worksheet.getColumn('C').width = 25;
            worksheet.getColumn('D').width = 25;
            worksheet.getColumn('E').width = 15;
            worksheet.getColumn('F').width = 15;
            worksheet.getColumn('G').width = 15;
            worksheet.getColumn('H').width = 15;
            worksheet.getColumn('I').width = 15;
    
            /* --- Títulos de las columnas --- */
            const titleRowData: string[] = [
                'Categoría',
                'Sub categoría',
                'Sub sub categoría',
                'Subtotal',
                'Impuestos',
                'Total',
                'Artículos',
                'Ventas',
                'Porcentaje',
            ];
    
            const titleRow = worksheet.addRow(titleRowData);
            titleRow.height = 30;
            titleRow.eachCell((cell) => {
                cell.style = STYLE_HEAD;
            });
            rowIni++;
    
            /* --- Registros --- */
            let row = (rowIni);
            console.time('orders')
            let total = 0;
            categories.forEach(result => total += result.amount_total);
            const promises = categories.map(async(result) => {
    
                /* --- Valores de Columnas --- */
                const bodyRowData: (string | number | Date)[] = [
                    result.uno,
                    result.dos,
                    result.tres,
                    result.amount_untaxed,
                    result.amount_tax,
                    result.amount_total,
                    result.articulos,
                    result.total_ventas,
                    ((result.amount_total / total) * 100).toFixed(2),
                ];
    
                return bodyRowData;
            });
            
            Promise.all(promises).then((rows) => {
                console.timeEnd('customerInvoices');
                const rowsU = rows.filter(row => row);
                if (rowsU.length > 0) {
                    const createdRows = worksheet.addRows(rowsU);
                    createdRows.forEach(row => {
                        /* Estilo de celdas */
                        row.eachCell((cell) => {
                            const col = +cell.col;
                            if ([1, 2, 3].includes(col)) {
                                cell.alignment = HORIZONTAL_CENTER_ALIGNMENT;
                            } else if (col >= 4 && col <= 6) {
                                cell.alignment = HORIZONTAL_RIGHT_ALIGNMENT;
                                cell.numFmt = CURRENCY_FORMAT;
                            } else if (col === colEndNumber) {
                                cell.numFmt = PERCENT_FORMAT;
                            }
    
                            if (col === 1) {
                                cell.border = {
                                    left: {
                                        style: 'thin',
                                        color: { argb: 'FF666666' }
                                    }
                                };
                            } else if (col === colEndNumber) {
                                cell.border = {
                                    right: {
                                        style: 'thin',
                                        color: { argb: 'FF666666' }
                                    }
                                };
                            }
    
                            if (+cell.row % 2) {
                                cell.style = Object.assign(cell.style, STYLE_ROW_STRIPED);
                            }
                        });
                    });
                }
                /* --- Totales --- */
                const rowCount = worksheet.rowCount;
                const totalRowData: (string | ExcelJS.CellValue)[] = [
                    '',
                    '',
                    'Total',
                    {
                        formula: `=SUM(D${rowIni + 1}:D${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(E${rowIni + 1}:E${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(F${rowIni + 1}:F${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(G${rowIni + 1}:G${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(H${rowIni + 1}:H${rowCount})`,
                        date1904: false,
                    },
                    {
                        formula: `=SUM(I${rowIni + 1}:I${rowCount})`,
                        date1904: false,
                    },
                ];

                const totalRow = worksheet.addRow(totalRowData);
    
                /* Estilo de Totales */
                totalRow.eachCell((cell) => {
                    const col = getExcelColumnLetter(+cell.col);
                    cell.style = Object.assign(cell.style, STYLE_FOOT);
                    if (cell.type != ExcelJS.ValueType.Null) {
                        if ((+cell.col) === colEndNumber) {
                            cell.numFmt = PERCENT_FORMAT;
                        } else if (![7, 8].includes(+cell.col)) {
                            cell.numFmt = CURRENCY_FORMAT;
                        }
                    }
    
                    if (col === 'A') {
                        cell.border = Object.assign({
                            left: {
                                style: 'thin',
                                color: { argb: 'FF666666' }
                            }
                        }, STYLE_FOOT.border);
                    } else if (col === colEnd) {
                        cell.border = Object.assign({
                            right: {
                                style: 'thin',
                                color: { argb: 'FF666666' }
                            }
                        }, STYLE_FOOT.border);
                    }
                });

                row += rowCount;
                cell = worksheet.getCell('A' + row);
                cell.value = (SOFT_NAME + ' v.' + VERSION);
                cell.style = STYLE_SOFT_VERSION;
                
                resolve(workbook.xlsx);
            });
        }
    });
};

export { generateReportOrderExcel };