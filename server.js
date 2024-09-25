import express from "express";
import IRApremRouter from "./src/routes/IRAPrems-routes.js";

const app = express();

app.listen(8000, () => console.log("server started"));

app.use("/ira", IRApremRouter);
