import connection from "../database";
import { dateStringFormatSql } from "../tools";
import { ReportOrderFilters, RowCategories } from "../types";
const PREFIX = process.env.DB_PREFIX;

const getCustomerInvoiceLinesReport = (filters: ReportOrderFilters, group?: number, sort?: number, order?: string): Promise<any[]> => {
    return new Promise((resolve, reject) => {
        const groupBy = [
            'ord.customer_id',
            'oli.order_line_id',
            'ord.order_id',
            'bra.branch_id',
            'shi.id_shift',
            'pro.product_id',
            'pro.product_category_id, pro.product_category2_id, pro.product_category3_id',
        ];
        const sortBy = [
            'pro.default_code',
            'ord.quotation_date',
            'shi.description',
        ];

        let sql = `SELECT SQL_CALC_FOUND_ROWS
        COUNT(oli.order_line_id) AS total_count,
        COUNT( DISTINCT ord.order_id ) AS total_order_count,
        SUM(oli.quantity) AS total_quantity,
        SUM(oli.price_subtotal/ord.currency_value) AS total_amount_untaxed,
        SUM(oli.price_tax/ord.currency_value) AS total_amount_tax,
        SUM(oli.price_tax_ret/ord.currency_value) AS total_amount_tax_ret,
        SUM((oli.price_tax/ord.currency_value) + (oli.price_tax_ret/ord.currency_value)) AS total_amount_tax_total,
        SUM(oli.price_total/ord.currency_value) AS total_amount_total,
        SUM(oli.cost/ord.currency_value) AS total_amount_cost,
        SUM(oli.margin/ord.currency_value) AS total_amount_margin,
        SUM( oli.price_total / ord.currency_value ) / COUNT( DISTINCT ord.order_id ) AS average,
        SUM(oli.margin/oli.price_subtotal/ord.currency_value * 100) AS total_percent_margin, oli.product_id, oli.\`name\`, oli.delivery_date,
        SUM(oli.quantity) AS quantity,
        SUM(oli.delivery_qty) AS delivery_qty,
        SUM(oli.invoice_qty) AS invoice_qty, ord.order_id, ord.\`name\` AS order_name, ord.date_order, ord.create_date, ord.quotation_date,
        ord.customer_order_ref,ord.location_id, ord.company_id, cus.\`name\` AS customer, sap.name AS salesperson,
        ost.\`name\` AS order_status,curr.\`name\` AS currency,pte.\`name\` AS payment_term,loc.\`name\` AS location, usr.name as usr_order,
        pro.\`name\` AS product,pro.product_uom_id,uom.decimal_place AS decimal_place_qty, pro.track, pro.subtract,
        pro.default_code, pro.description, pro.description_sale, ord.amount_untaxed, ord.amount_tax,ord.amount_total,
        SUM(ord.amount_total) AS sum_total,
        SUM(oli.quantity-oli.delivery_qty) AS per_delivery_qty,
        SUM(oli.delivery_qty-oli.invoice_qty) AS per_invoice_qty,prc.\`name\` AS product_category,ord.origin,
        ord.order_status_id,shi.description AS shift_description, shi.id_shift, bra.name AS name_branch, payment_method
        FROM ${PREFIX}order_line AS oli
            LEFT JOIN ${PREFIX}product_uom AS uom ON oli.product_uom_id=uom.product_uom_id
            LEFT JOIN ${PREFIX}product AS pro ON oli.product_id=pro.product_id
            LEFT JOIN ${PREFIX}product_category AS prc ON pro.product_category_id=prc.product_category_id
            LEFT JOIN ${PREFIX}stock_production_lot AS spl ON oli.production_lot_id=spl.production_lot_id
            INNER JOIN ${PREFIX}order AS ord ON oli.order_id=ord.order_id
            LEFT JOIN ${PREFIX}location AS loc ON ord.location_id=loc.location_id
            INNER JOIN ${PREFIX}customer AS cus ON ord.customer_id=cus.customer_id
            INNER JOIN ${PREFIX}user AS sap ON ord.salesperson_id=sap.id
            INNER JOIN ${PREFIX}user AS usr ON ord.create_uid=usr.id
            INNER JOIN (SELECT cash_open_id, order_id, GROUP_CONCAT(payment_method) AS payment_method 
                        FROM ${PREFIX}cash_move GROUP BY order_id) AS mov ON ord.order_id = mov.order_id
            INNER JOIN ${PREFIX}cash_open AS op ON mov.cash_open_id = op.cash_open_id
            INNER JOIN ${PREFIX}shift AS shi ON op.shift_id = shi.id_shift
            INNER JOIN ${PREFIX}branch AS bra ON bra.branch_id=shi.branch_id
            INNER JOIN ${PREFIX}order_status AS ost ON ord.order_status_id=ost.order_status_id
            INNER JOIN ${PREFIX}currency AS curr ON ord.currency_id=curr.currency_id
            LEFT JOIN ${PREFIX}payment_term AS pte ON ord.payment_term_id=pte.payment_term_id
        WHERE oli.order_line_id > 0`;

        /* Filtros */
        if (filters) {
            if (filters.name) {
                sql += ` AND ord.name LIKE '%${ filters.name.replace(/\s/gm, '%%') }'`;
            }
            if (filters.dateFrom) {
                const date: string = dateStringFormatSql(filters.dateFrom);
                sql += ` AND DATE(ord.date_order) >= DATE('${ date }')`;
            }
            if (filters.dateTo) {
                const date: string = dateStringFormatSql(filters.dateTo);
                sql += ` AND DATE(ord.date_order) <= DATE('${ date }')`;
            }
            if (filters.customer) {
                sql += ` AND cus.name LIKE '%${ filters.customer.replace(/\s/gm, '%%') }%'`;
            }
            if (filters.reference) {
                sql += ` AND ord.customer_order_ref LIKE '%${ filters.reference.replace(' ', '%%') }%'`;
            }
            if (filters.salesperson) {
                sql += ` AND sap.name LIKE '%${ String(filters.salesperson).replace(/\s/gm, '%%') }%'`;
            }
            if (filters.origin) {
                sql += ` AND ord.origin LIKE '%${ filters.origin.replace(/\s/gm, '%%') }'`;
            }
            if (filters.orderStatusId) {
                sql += ` AND ord.order_status_id IN (${ filters.orderStatusId })`;
            }
            if (filters.boxId) {
                sql += ` AND shi.id_shift = ${filters.boxId}`;
            }
            if (filters.branchId) {
                sql += ` AND bra.branch_id = ${filters.branchId}`;
            }
            if (filters.productId) {
                sql += ` AND pro.product_id = ${filters.productId}`;
            }
            if (filters.productCategoryId) {
                sql += ` AND (pro.product_category_id = ${filters.productCategoryId} OR pro.product_category2_id = ${filters.productCategoryId} OR pro.product_category3_id = ${filters.productCategoryId})`;
            }
            sql += ` AND ord.quotation = 0`;
        }

        if (group) {
            sql += ` ${groupBy[group] ? `GROUP BY ${groupBy[group]}` : ''}`;
        }

        if (sort) {
            sql += ` ORDER BY ${sortBy[sort] ? sortBy[sort] : 'cin.date_invoice,cin.\`name\`'}`;
            
            if (order) {
                sql += ` ${order}`;
            }
        }

        connection.query(sql, (err, rows) => {
            if (err)
                reject(err);

            resolve(rows);
        });
    });
};

const getPaymentMethod = (paymentMethodIds: string): Promise<string | null> => {
    return new Promise((resolve, reject) => {
        const sql = `SELECT GROUP_CONCAT(name SEPARATOR ', ') AS payment_methods
                    FROM ${PREFIX}payment_method WHERE payment_method IN (${paymentMethodIds})`;

        connection.query(sql, (err, rows) => {
            if (err)
                reject(err);

            resolve(rows[0] && rows[0].payment_methods ? rows[0].payment_methods : null);
        })
    });
}

const getDevoluciones = (filters: ReportOrderFilters): Promise<number> => {
    return new Promise((resolve, reject) => {
        let sql = `SELECT SQL_CALC_FOUND_ROWS SQL_NO_CACHE DISTINCT(cin.customer_invoice_id)
        FROM ya_customer_invoice AS cin
        	INNER JOIN ya_customer_remission_refund_invoice_rel AS rel On cin.customer_invoice_id = rel.customer_remission_refund_id
            INNER JOIN ya_customer_invoice AS rem ON rel.customer_invoice_id = rem.customer_invoice_id
            INNER JOIN ya_customer AS cus ON cin.customer_id=cus.customer_id
            INNER JOIN ya_order AS ord ON rem.origin = ord.name
            INNER JOIN ya_order_line AS oli ON oli.order_id = ord.order_id
            LEFT JOIN ya_product AS pro ON oli.product_id=pro.product_id
            INNER JOIN ya_user AS sap ON ord.salesperson_id=sap.id
            INNER JOIN ya_user AS usr ON ord.create_uid=usr.id
            LEFT JOIN (SELECT cash_open_id, order_id, GROUP_CONCAT(payment_method) AS payment_method
                        FROM ya_cash_move GROUP BY order_id) AS mov ON ord.order_id = mov.order_id
            LEFT JOIN ya_cash_open AS op ON mov.cash_open_id = op.cash_open_id
            LEFT JOIN ya_shift AS shi ON op.shift_id = shi.id_shift
            LEFT JOIN ya_branch AS bra ON bra.branch_id=shi.branch_id
            WHERE cin.customer_invoice_id>0 AND cin.transaction_sequence_id = 23
            AND (cin.invoice_status_id = '1' OR cin.invoice_status_id = '2' OR cin.invoice_status_id = '3'
            OR cin.invoice_status_id = '5')`;

        /* Filtros */
        if (filters) {
            if (filters.name) {
                sql += ` AND ord.name LIKE '%${ filters.name.replace(/\s/gm, '%%') }'`;
            }
            if (filters.dateFrom) {
                const date: string = dateStringFormatSql(filters.dateFrom);
                sql += ` AND DATE(cin.date_invoice) >= DATE('${ date }')`;
            }
            if (filters.dateTo) {
                const date: string = dateStringFormatSql(filters.dateTo);
                sql += ` AND DATE(cin.date_invoice) <= DATE('${ date }')`;
            }
            if (filters.customer) {
                sql += ` AND cus.name LIKE '%${ filters.customer.replace(/\s/gm, '%%') }%'`;
            }
            if (filters.reference) {
                sql += ` AND ord.customer_order_ref LIKE '%${ filters.reference.replace(' ', '%%') }%'`;
            }
            if (filters.salesperson) {
                sql += ` AND sap.name LIKE '%${ String(filters.salesperson).replace(/\s/gm, '%%') }%'`;
            }
            if (filters.origin) {
                sql += ` AND ord.origin LIKE '%${ filters.origin.replace(/\s/gm, '%%') }'`;
            }
            if (filters.orderStatusId) {
                sql += ` AND ord.order_status_id IN (${ filters.orderStatusId })`;
            }
            if (filters.boxId) {
                sql += ` AND shi.id_shift = ${filters.boxId}`;
            }
            if (filters.branchId) {
                sql += ` AND bra.branch_id = ${filters.branchId}`;
            }
            if (filters.productId) {
                sql += ` AND pro.product_id = ${filters.productId}`;
            }
            if (filters.productCategoryId) {
                sql += ` AND (pro.product_category_id = ${filters.productCategoryId} OR pro.product_category2_id = ${filters.productCategoryId} OR pro.product_category3_id = ${filters.productCategoryId})`;
            }
            sql += ` AND ord.quotation = 0`;
        }

        connection.query(sql, (err, rows) => {
            if (err)
                reject(err);

            resolve(rows ? rows.length : 0);
        });
    });
};

const getCategories = (filters: ReportOrderFilters): Promise<RowCategories[]> => {
    return new Promise((resolve, reject) => {
        let sql = `SELECT SQL_CALC_FOUND_ROWS COUNT(DISTINCT ord.order_id) as total_ventas,
        pro.\`name\`, cat.name, cat.product_category_id, ROUND(SUM(oli.price_subtotal),2) AS amount_untaxed,
        ROUND(SUM(oli.price_tax), 2) AS  amount_tax, ROUND(SUM(oli.price_total),2) AS amount_total,
        cat.\`name\` as uno, prc2.name as dos, prc.name as tres,
        COUNT(oli.order_line_id) AS articulos
        FROM ya_product AS pro
        LEFT JOIN ya_product_category AS cat ON pro.product_category_id = cat.product_category_id
        LEFT JOIN ya_product_category AS prc2 ON pro.product_category2_id=prc2.product_category_id
        LEFT JOIN ya_product_category AS prc ON pro.product_category3_id=prc.product_category_id
        INNER JOIN ya_order_line AS oli ON oli.product_id = pro.product_id
        INNER JOIN ya_order AS ord ON oli.order_id = ord.order_id
        LEFT JOIN ya_user AS sap ON ord.salesperson_id=sap.id
        LEFT JOIN ya_user AS usr ON ord.create_uid=usr.id
        LEFT JOIN ya_shift AS shi ON shi.salesperson_id=sap.id
        LEFT JOIN ya_branch AS bra ON bra.branch_id=shi.branch_id
        WHERE pro.product_id <> 1`;
        
        /* Filtros */
        if (filters) {
            if (filters.dateFrom) {
                const date: string = dateStringFormatSql(filters.dateFrom);
                sql += ` AND DATE(ord.date_order) >= DATE('${ date }')`;
            }
            if (filters.dateTo) {
                const date: string = dateStringFormatSql(filters.dateTo);
                sql += ` AND DATE(ord.date_order) <= DATE('${ date }')`;
            }
            if (filters.customer) {
                sql += ` AND cus.name LIKE '%${ filters.customer.replace(/\s/gm, '%%') }%'`;
            }
            if (filters.reference) {
                sql += ` AND ord.customer_order_ref LIKE '%${ filters.reference.replace(' ', '%%') }%'`;
            }
            if (filters.salesperson) {
                sql += ` AND sap.name LIKE '%${ String(filters.salesperson).replace(/\s/gm, '%%') }%'`;
            }
            if (filters.origin) {
                sql += ` AND ord.origin LIKE '%${ filters.origin.replace(/\s/gm, '%%') }'`;
            }
            if (filters.orderStatusId) {
                sql += ` AND ord.order_status_id IN (${ filters.orderStatusId })`;
            }
            if (filters.boxId) {
                sql += ` AND shi.id_shift = ${filters.boxId}`;
            }
            if (filters.branchId) {
                sql += ` AND bra.branch_id = ${filters.branchId}`;
            }
            if (filters.productCategoryId) {
                sql += ` AND (pro.product_category_id = ${filters.productCategoryId} OR pro.product_category2_id = ${filters.productCategoryId} OR pro.product_category3_id = ${filters.productCategoryId})`;
            }
            if (filters.productCategory2Id) {
                sql += ` AND pro.product_category2_id = ${filters.productCategory2Id}`;
            }
            if (filters.productCategory3Id) {
                sql += ` AND pro.product_category3_id = ${filters.productCategory3Id}`;
            }
            sql += ` AND ord.quotation = 0`;
        }

        sql += ` GROUP BY pro.product_category3_id`;

        connection.query(sql, (err, rows) => {
            if (err)
                reject(err);

            resolve(rows);
        });
    });
};

export { getCustomerInvoiceLinesReport, getPaymentMethod, getDevoluciones, getCategories };