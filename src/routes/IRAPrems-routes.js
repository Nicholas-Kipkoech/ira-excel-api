import { Router } from "express";
import { IRAPremClass } from "../services/IRA-class-services.js";
import { IRABusinessForce } from "../services/IRA-business-force.js";
import { IRACommissionService } from "../services/IRA-commisions.js";
import { IRAIncurredClaimsService } from "../services/IRAIncurredClaims.js";
import { IRAReinsurancePremiumsService } from "../services/IRA-reinsurance-premiums.js";
import { IRAUnearnedPremiumsService } from "../services/IRA-unearned-premiums.js";

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
IRApremRouter.post("/ira-incurred-claims", (req, res) => {
  IRAIncurredClaimsService.getIncuredClaims(req, res);
});
IRApremRouter.post("/ira-reinsurance-premiums", (req, res) => {
  IRAReinsurancePremiumsService.getReinsurancePremiums(req, res);
});
IRApremRouter.post("/ira-unearned-premiums", (req, res) => {
  IRAUnearnedPremiumsService.getUnearnedPremiums(req, res);
});

export default IRApremRouter;
