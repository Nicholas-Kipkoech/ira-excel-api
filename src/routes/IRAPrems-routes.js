import { Router } from "express";
import { IRAPremClass } from "../services/IRA-class-services.js";
import { IRABusinessForce } from "../services/IRA-business-force.js";
import { IRACommissionService } from "../services/IRA-commisions.js";
import { IRAIncurredClaimsService } from "../services/IRAIncurredClaims.js";
import { IRAReinsurancePremiumsService } from "../services/IRA-reinsurance-premiums.js";
import { IRAUnearnedPremiumsService } from "../services/IRA-unearned-premiums.js";
import { IRAPremiumsByCounty } from "../services/IRA-premiums-county.js";
import { IRAPremiumRegisterService } from "../services/IRA-premiums-register.js";
import { IRAReinsuranceBalancesService } from "../services/IRA-reinsurance-balanace-.js";

const IRApremRouter = Router();

IRApremRouter.get("/ira-premiums", (req, res) => {
  IRAPremClass.getPremiums(req, res);
});
IRApremRouter.get("/business-force", (req, res) => {
  IRABusinessForce.getBusinessForcePrems(req, res);
});

IRApremRouter.get("/ira-commisions", (req, res) => {
  IRACommissionService.getCommissions(req, res);
});
IRApremRouter.get("/ira-incurred-claims", (req, res) => {
  IRAIncurredClaimsService.getIncuredClaims(req, res);
});
IRApremRouter.get("/ira-reinsurance-premiums", (req, res) => {
  IRAReinsurancePremiumsService.getReinsurancePremiums(req, res);
});
IRApremRouter.get("/ira-unearned-premiums", (req, res) => {
  IRAUnearnedPremiumsService.getUnearnedPremiums(req, res);
});
IRApremRouter.get("/ira-premiums-county", (req, res) => {
  IRAPremiumsByCounty.getPremiumsByCounty(req, res);
});

IRApremRouter.get("/ira-premiums-register", (req, res) => {
  IRAPremiumRegisterService.getPremiums(req, res);
});
IRApremRouter.get("/ira-reinsurance-balances", (req, res) => {
  IRAReinsuranceBalancesService.getBalanceReport(req, res);
});

export default IRApremRouter;
