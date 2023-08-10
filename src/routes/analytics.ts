import express from "express";
import 'dotenv/config';
import { AnalyticsRequest, ReportOrderFilters } from "../types";
import { generateReportOrderExcel } from "../services/analyticsService";
import { CustomError } from "../tools";
const analyticsRouter = express.Router();

analyticsRouter.get('/reporte_pedidos_de_venta.xlsx', async(req: AnalyticsRequest, res) => {
    const data = req.query;
    const fileName = 'reporte_pedidos_de_venta.xlsx';

    try {
        const xlsx = await generateReportOrderExcel(data as ReportOrderFilters);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader("Content-Disposition", "attachment; filename=" + fileName);
        await xlsx.write(res);
        res.end();
    } catch(err) {
        console.error(err);
        if (err instanceof CustomError) {
            res.status(err.code).send(err.message);
        }
    }
});

export default analyticsRouter;