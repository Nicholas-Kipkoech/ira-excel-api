import { Router } from "express";
import { IRAPremClass } from "../services/IRA-class-services.js";

const IRApremRouter = Router();

IRApremRouter.post("/ira-premiums", (req, res) => {
  IRAPremClass.getPremiums(req, res);
});

export default IRApremRouter;
