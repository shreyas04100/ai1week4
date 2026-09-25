# KPI Framework Analysis
## NovaMart Week 4: Performance Evaluation & Impact Assessment
### Research Date: September 2026

---

## KPI Hierarchy

### Level 1: Strategic KPIs (Board / C-Suite)
- Total AI ROI
- Revenue attributable to AI
- Cost savings from AI
- Customer satisfaction (NPS)
- AI adoption rate across business units

### Level 2: Business KPIs (Business Unit Heads)
- Stockout rate reduction
- Inventory turnover improvement
- Conversion rate uplift
- Customer churn reduction
- Fraud loss reduction
- Support ticket deflection rate

### Level 3: Operational KPIs (Operations / Category Managers)
- Forecast accuracy (MAPE/WAPE)
- Recommendation acceptance rate
- Human override rate
- Replenishment cycle time
- Data pipeline success rate

### Level 4: Technical KPIs (Data Science / Engineering)
- Model precision / recall / F1
- Prediction latency
- Data drift PSI
- Model staleness (days since retraining)
- API uptime / error rate
- Feature store freshness

---

## Metric Ownership Matrix

| Metric Category | Primary Owner | Secondary Owner | Reporting Frequency |
|---|---|---|---|
| Model performance (MAPE, F1, AUC) | Data Science Lead | ML Engineer | Weekly |
| Data quality | Data Engineering Lead | Data Analyst | Daily |
| Business KPIs (stockout, conversion) | Business Unit Head | Category Manager | Weekly/Monthly |
| Financial KPIs (ROI, cost savings) | CFO / Finance | AI Product Owner | Monthly/Quarterly |
| Customer KPIs (NPS, CSAT, churn) | Customer Experience Head | CRM Team | Monthly |
| Adoption KPIs (override rate, usage) | AI Product Owner | Change Manager | Weekly |
| Risk/Governance KPIs | Chief Risk Officer | Security Team | Monthly |
| Strategic KPIs | CEO / Board | All | Quarterly |

---

## Baseline Definition Template

For each AI initiative, the following baseline must be established BEFORE deployment:

| Field | Description |
|---|---|
| Baseline period | Minimum 3 months of pre-AI data |
| Baseline metric values | Measured values for all KPIs during baseline period |
| Baseline data sources | Which systems provided baseline data |
| Seasonality adjustment | Whether baseline is adjusted for seasonal effects |
| Control group | Whether a control group (non-AI stores/users) is maintained |
| Baseline owner | Who is responsible for baseline data integrity |
| Baseline sign-off | Business owner approval of baseline values |

---

## A/B Testing Framework

### When to use A/B testing:
- Customer-facing AI (personalization, recommendations, conversational AI)
- When a clean control group can be maintained
- When sample size is sufficient for statistical significance

### When NOT to use A/B testing:
- Supply chain AI (cannot run two inventory systems simultaneously)
- Fraud detection (cannot allow fraud in control group)
- Use before/after comparison with seasonality adjustment instead

### Statistical significance requirements:
- Minimum sample size: calculated using power analysis (α=0.05, β=0.20, minimum detectable effect)
- Test duration: minimum 2 weeks; avoid festive season contamination
- Primary metric: pre-specified before test starts (no p-hacking)
- Guardrail metrics: metrics that must not degrade (e.g., page load time, error rate)

---

## ROI Calculation Framework

### Investment Components (Costs):
1. Initial development cost (one-time)
2. Infrastructure cost (ongoing, annual)
3. Licensing / API cost (ongoing, annual)
4. Team cost (ongoing, annual)
5. Training and change management (one-time)
6. Maintenance and monitoring (ongoing, annual)

### Benefit Components:
1. Direct revenue increase (measurable, attributable)
2. Cost savings (measurable, attributable)
3. Loss prevention (fraud, waste)
4. Productivity gains (labor hours saved)
5. Indirect benefits (customer satisfaction, brand value — harder to quantify)

### ROI Formula:
ROI = (Total Benefits − Total Costs) / Total Costs × 100%

### Payback Period:
Payback Period = Total Investment / Annual Net Benefit

### Net Present Value (NPV):
NPV = Σ [Net Benefit_t / (1 + r)^t] − Initial Investment
Where r = discount rate (typically 10–15% for Indian enterprises)

### Attribution Challenge:
Not all business improvement can be attributed to AI alone.
Use: Incremental attribution — compare AI group vs control group, or before vs after with external factor adjustment.
