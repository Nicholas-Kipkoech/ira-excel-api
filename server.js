import express from "express";
import IRApremRouter from "./src/routes/IRAPrems-routes.js";
import cors from "cors";
const app = express();
app.use(cors());

app.listen(8000, () => console.log("server started"));

app.use("/ira", IRApremRouter);
