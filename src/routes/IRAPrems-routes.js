import { Router } from "express";
import { IRAPremClass } from "../services/IRA-class-services.js";
import { IRABusinessForce } from "../services/IRA-business-force.js";
import { IRACommissionService } from "../services/IRA-commisions.js";

const IRApremRouter = Router();

IRApremRouter.post("/ira-premiums", (req, res) => {
  IRAPremClass.getPremiums(req, res);
});
IRApremRouter.post("/business-force", (req, res) => {
  IRABusinessForce.getBusinessForcePrems(req, res);
});

IRApremRouter.post("/ira-commisions", (req, res) => {
  IRACommissionService.getCommissions(req, res);
});

export default IRApremRouter;
