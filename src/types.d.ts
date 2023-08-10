import { Request } from "express";

export type ReportType = 'customer' | 'transaction_sequence' | 'detail' | 'branch' | 'box' | 'product' | 'category';
export type AnalyticsRequest = Request<{}, any, any, AnalyticsFilters | ReportOrderFilters>;

export interface AnalyticsFilters {
    dateFrom?: string,
    dateTo: string,
    name?: string,
    origin?: string,
    customer?: string | number | Array,
    salesperson?: string | number,
    reportType: ReportType,
};

export interface ReportOrderFilters extends AnalyticsFilters {
    productId?: string | number,
    productCategoryId?: string | number,
    reference?: string,
    boxId?: string | number,
    branchId?: string | number,
    showCanceled?: string | number,
    orderStatusId?: string | Array,
    sort?: string | number,
    order?: string | number,
    configCurrencyCode: string,
    productCategory2Id?: string | number
    productCategory3Id?: string | number
}

export interface RowReportOrder {
    quotation_date: Date,
    date_order: Date,
    create_date: Date,
    shift_description: string,
    id_shift: string,
    name_branch: string,
    payment_method: string,
    total_quantity: number,
    total_amount_cost: number,
    total_amount_untaxed: number,
    total_amount_untaxed_withoutdiscount: number,
    total_amount_discount: number,
    total_amount_tax: number,
    total_amount_tax_ret: number,
    total_amount_tax_total: number,
    total_amount_total: number,
    total_amount_cost: number,
    total_amount_margin: number,
    total_percent_margin: number,
    order_name: string,
    origin: string,
    customer_order_ref: string,
    remission: string,
    salesperson: string,
    order_status_id: string,
    product: string,
    product_uom_id: string,
    decimal_place_qty: string,
    product_category: string,
    customer: string,
    name: string,
    default_code: string,
    total_order_count: number,
    average: number,
    quantity: number,
    amount_total: number,
    usr_order: string,
    sum_total: number
}

export interface RowCategories {
    total_ventas: number,
    name: string,
    product_category_id: string,
    amount_untaxed: number,
    amount_tax: number,
    amount_total: number,
    uno: string,
    dos: string,
    tres: string,
    articulos: number,
}

declare global {
    interface Date {
        addDays(days: number): Date,
        subDays(days: number): Date,
        formatSql(): string,
    };
    declare var SOFT_NAME: String;
    declare var VERSION: String;
}
