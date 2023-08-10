import express from 'express';
import cors from 'cors';
import analyticsRouter from './routes/analytics';
import { configure } from './tools';
const app = express();
const PORT = process.env.PORT || 3000;

configure();
app.use((req, res, next) => {
    req.setTimeout(300000, () => {
        /* Esta función de devolución de llamada se ejecutará cuando la solicitud supere el  (5 min). */
        console.log('La solicitud ha superado el tiempo limite!');
        res.status(504).send('Gateway Timeout');
    });
    next();
})
app.use(cors());
app.use(express.json());

app.all("/ping", (_req, res) => {
    console.log("Someone pinged here!!");
    res.status(200)
        .json({ message: "Pong!" });
});

app.use("/api/analytics", analyticsRouter);

app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});