import { Router } from "express";
import { IRAPremClass } from "../services/IRA-class-services.js";
import { IRABusinessForce } from "../services/IRA-business-force.js";

const IRApremRouter = Router();

IRApremRouter.post("/ira-premiums", (req, res) => {
  IRAPremClass.getPremiums(req, res);
});
IRApremRouter.post("/business-force", (req, res) => {
  IRABusinessForce.getBusinessForcePrems(req, res);
});

export default IRApremRouter;
