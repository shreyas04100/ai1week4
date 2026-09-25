from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Inches
from docx.enum.table import WD_TABLE_ALIGNMENT


def add_h1(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(16)
    p.style = 'Heading 1'


def add_h2(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(13)
    p.style = 'Heading 2'


def add_para(doc, text):
    doc.add_paragraph(text)


def add_bold_para(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True


def add_table_from_rows(doc, rows, widths=None):
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    if widths:
        for idx, width in enumerate(widths):
            table.columns[idx].width = Inches(width)
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            table.cell(r, c).text = value
    return table


def add_diagram_table(doc, labels, title=None):
    if title:
        add_h2(doc, title)
    table = doc.add_table(rows=1, cols=len(labels))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, label in enumerate(labels):
        cell = table.cell(0, i)
        cell.text = label
        for p in cell.paragraphs:
            for r in p.runs:
                r.bold = True
                r.font.size = Pt(10)
    for cell in table.columns[0].cells:
        cell.width = Inches(1.4)
    for i in range(1, len(labels)):
        table.columns[i].width = Inches(1.2)
    return table


def add_diagram_flow(doc, title, steps):
    add_h2(doc, title)
    row = [str(step) for step in steps]
    table = doc.add_table(rows=1, cols=len(row))
    for idx, text in enumerate(row):
        cell = table.cell(0, idx)
        cell.text = text
        for p in cell.paragraphs:
            for r in p.runs:
                r.bold = True
                r.font.size = Pt(9)
        if idx < len(row) - 1:
            cell.text = text + '  →'
    return table


doc = Document()
section = doc.sections[0]
section.top_margin = Inches(0.7)
section.bottom_margin = Inches(0.7)
section.left_margin = Inches(0.7)
section.right_margin = Inches(0.7)

# 1. Executive Summary
add_h1(doc, '1. Executive Summary')
add_para(doc, 'AI solutions need continuous evaluation after deployment because model quality is only one input to business performance. In retail, a forecast model may be highly accurate on a holdout sample but still fail to create value if supply-chain workflows, inventory policies, or sales processes do not change in time. Continuous evaluation ensures that model quality, system reliability, operational execution, and business results remain aligned with strategic goals.')
add_para(doc, 'Technical model accuracy alone is not enough. A churn model may achieve strong precision and recall but deliver limited value if interventions are not available, customer teams cannot act on predictions, or the retention strategy is poorly timed. Similarly, a recommendation system may increase click-through but reduce average order value or create customer dissatisfaction if not governed carefully. The real question is not whether the model is good in isolation; it is whether the AI capability creates measurable improvements in revenue, costs, service quality, operational efficiency, and customer outcomes.')
add_para(doc, 'This framework distinguishes model performance from business impact. Model performance measures statistical quality (for example, precision, recall, MAPE, or ROC-AUC), while business impact measures the value created for the organization in terms of revenue uplift, loss reduction, productivity, customer retention, and risk reduction. The framework connects AI metrics to business KPIs through a structured hierarchy: technical metrics feed operational metrics, which feed customer and financial metrics, which in turn support management decisions and strategic planning.')
add_para(doc, 'The framework supports continuous improvement by establishing baselines, monitoring drift and governance signals, analyzing root causes, and strengthening interventions through retraining, process redesign, and workflow education. It creates a disciplined loop for decision-makers so that AI investments are reviewed not only on technical merit but on actual business value and accountable operational performance.')

doc.add_page_break()

# 2. Introduction
add_h1(doc, '2. Introduction')
add_para(doc, 'This document presents a practical evaluation and impact assessment framework for AI solutions in a business context, using NovaMart Retail & E-Commerce Private Limited as a hypothetical omnichannel retailer. The framework is designed to evaluate AI initiatives in a structured, objective, and actionable way across model performance, operational execution, customer value, financial return, and governance.')
add_para(doc, 'The evaluation model is especially relevant for retailers using AI in demand forecasting, inventory optimization, personalization, churn prevention, fraud detection, and conversational commerce. These use cases differ in business objectives and evaluation requirements. For example, demand forecasting is evaluated mainly through forecast accuracy and operational service metrics, while fraud detection emphasizes precision, recall, detection latency, and customer friction. Generative AI in customer support must be evaluated with relevance, groundedness, factuality, latency, and human escalation, not only with classification-style metrics.')
add_para(doc, 'The purpose of this framework is to prevent an overly narrow focus on model scores or demo-level performance. Real business value is created only when AI capabilities are accurate, trusted, operationally feasible, financially worthwhile, compliant, adopted by users, and sustainable over time. The framework therefore combines technical, operational, customer, financial, governance, and strategic lenses.')
add_para(doc, 'All NovaMart-specific assumptions, KPI targets, example results, costs, and financial projections in this document are clearly labelled as illustrative, hypothetical, or proposed because they are not based on actual NovaMart measured data.')

doc.add_page_break()

# 3. NovaMart Business Context
add_h1(doc, '3. NovaMart Business Context')
add_para(doc, 'NovaMart is treated as a hypothetical omnichannel retailer operating across physical stores, e-commerce, marketplace channels, fulfilment centres, and a digital customer engagement platform. The company has a broad and complex operating model that makes AI relevant across merchandising, catalog management, customer engagement, fraud prevention, and fulfillment planning.')
add_para(doc, 'The following hypothetical AI initiatives are considered as examples for the evaluation framework:')
for item in [
    'Demand forecasting and inventory optimization for category replenishment across stores and warehouses.',
    'Customer personalization and product recommendation engines for homepage, email, and app experiences.',
    'Fraud detection to identify suspicious transactions before fulfillment or payment settlement.',
    'Customer churn prediction to target retention offers and service interventions with high-risk customers.',
    'Generative AI conversational commerce for product discovery, order support, and customer assistant workflows.'
]:
    doc.add_paragraph(item, style='List Bullet')
add_para(doc, 'These initiatives are presented as proposed or illustrative use cases. They are not claimed to be already deployed or operating at NovaMart. This evaluation framework is designed so that each initiative can be assessed using an appropriate baseline, measurement method, and value model before or during implementation.')

doc.add_page_break()

# 4. Purpose of the Evaluation Framework
add_h1(doc, '4. Purpose of the Evaluation Framework')
add_para(doc, 'The purpose of the framework is to measure whether an AI solution produces value in a real operating environment. It is intended to answer five fundamental management questions:')
for item in [
    'Is the AI model performing correctly and reliably?',
    'Is the system delivering value to operations and customer experience?',
    'Is the solution improving financial outcomes or reducing risk?',
    'Is adoption occurring across the business, and are users trusting the system?',
    'What should be improved next to generate sustained value?'
]:
    doc.add_paragraph(item, style='List Bullet')
add_para(doc, 'The framework is built for business decision-making rather than technical model reporting alone. It ensures that AI performance is interpreted through the lens of strategy, accountability, and implementation readiness.')

doc.add_page_break()

# 5. Evaluation Principles
add_h1(doc, '5. Evaluation Principles')
for item in [
    'Create a baseline before AI deployment to establish the pre-AI operating state.',
    'Measure both model performance and business impact using consistent definitions and time periods.',
    'Use metrics that match the use case: classification, forecasting, recommendation, NLP, or anomaly detection require different evaluations.',
    'Separate technical efficacy from operational value and financial return.',
    'Assess both expected benefits and unintended harms, including customer friction, privacy risk, bias, and governance issues.',
    'Use experimental or quasi-experimental methods where possible, including A/B tests, control groups, and before/after comparisons.',
    'Tie AI metrics to owner accountability and decision triggers.',
    'Support continuous improvement through robust feedback loops, retraining, and process adaptation.'
]:
    doc.add_paragraph(item, style='List Bullet')
add_para(doc, 'Good model accuracy does not automatically prove business value. An AI system may show excellent test metrics but still fail in production because of poor data quality, weak operating workflows, bad adoption, insufficient actionability, or flawed incentives. The evaluation framework is therefore designed to avoid misleading attribution and to separate genuine value creation from noise, seasonality, and operational changes.')

doc.add_page_break()

# 6. AI Performance Evaluation Framework
add_h1(doc, '6. AI Performance Evaluation Framework')
add_para(doc, 'The following evaluation dimensions are used to assess AI initiatives at NovaMart. Each dimension is measured using business-relevant metrics, connected to owners, and linked to action thresholds when performance is weak.')
rows = [
    ['Metric', 'Definition', 'Formula/Method', 'Data Source', 'Frequency', 'Target/Threshold', 'Owner', 'Action if Poor'],
    ['Model Performance', 'Ability of the model to predict correctly or rank outcomes as intended.', 'Compare model outputs to labelled validation/test data or time-based holdouts.', 'Model logs, labelled datasets, feature stores', 'Daily/Weekly', 'Meets or exceeds business threshold (e.g., F1 or MAPE target)', 'Data Science Lead', 'Retrain, adjust threshold, review feature set'],
    ['Data Quality', 'Reliability and completeness of input and label data.', 'Missingness, valid range checks, freshness, duplicate rate, lineage review.', 'ERP, POS, CRM, WMS, API logs', 'Daily', 'Data quality SLA >= 99%', 'Data Engineering Lead', 'Fix ingestion, data contract, and monitoring'],
    ['System Performance', 'Operational responsiveness and reliability of the AI platform.', 'Latency, throughput, uptime, API error rate, queue depth.', 'Application analytics, API logs, infra metrics', 'Daily', 'P99 latency below defined SLA and uptime >= 99.9%', 'ML Engineer', 'Scale infra, optimize model serving'],
    ['Operational Efficiency', 'Impact on staffing, cycle time, and workflow productivity.', 'Time saved or task completion rate per workflow.', 'WMS, ERP, productivity logs', 'Weekly/Monthly', 'Measure improvement versus baseline', 'Operations Manager', 'Redesign workflow and retrain staff'],
    ['Customer Experience', 'How AI changes customer behavior or service quality.', 'Customer satisfaction, conversion, repeat purchase, support metrics.', 'CRM, e-commerce, surveys', 'Weekly/Monthly', 'CSAT/NPS above target range', 'Customer Experience Lead', 'Adjust UX, escalation, or recommendation logic'],
    ['Financial Impact', 'Economic benefit generated by AI adoption.', 'Incremental revenue, cost savings, profit contribution, ROI.', 'Finance systems, ERP, POS, BI dashboards', 'Monthly/Quarterly', 'ROI > 0 and payback within target period', 'CFO / Finance', 'Re-evaluate vendor cost or redesign business case'],
    ['AI Adoption', 'How often users and teams use AI-supported workflows.', 'Adoption rate, override rate, training completion, active users.', 'Application analytics, HR training data', 'Weekly/Monthly', 'Usage above critical threshold', 'AI Product Owner', 'Train users, improve UX, provide decision support'],
    ['Risk and Governance', 'Ability to manage bias, privacy, security, explainability, and compliance.', 'Incident counts, audit completion, explainability coverage, model drift reviews.', 'Audit logs, security tickets, model registry', 'Monthly', 'Zero critical incidents and policy compliance', 'Chief Risk Officer', 'Restrict access, escalate controls, suspend model'],
    ['Strategic Impact', 'Contribution to competitiveness and strategic goals.', 'Outcome of leader-approved strategic scorecards and business reviews.', 'Leadership roadmap, BI scorecards', 'Quarterly', 'Strategic KPI aligned to business growth and risk posture', 'CEO / Business Leaders', 'Re-prioritize AI investment portfolio']
]

table = add_table_from_rows(doc, rows, widths=[1.4, 1.5, 1.8, 1.6, 0.8, 1.1, 0.9, 1.7])
for row in table.rows:
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(8)

add_h2(doc, 'Overall AI Evaluation Framework Diagram')
add_diagram_flow(doc, '', ['Business Objective', 'AI Capability', 'Data & Model Quality', 'Operational Value', 'Customer Value', 'Financial Return', 'Governance & Risk'])
add_para(doc, 'The use of each metric is not generic. A forecasting model is evaluated on MAPE, WAPE, bias, and service-level outcomes; a fraud model is judged on precision, recall, false-positive rate, and fraud loss reduction; a recommendation model is evaluated on CTR, conversion rate, and recommendation acceptance; and a GenAI assistant is evaluated on groundedness, factuality, relevance, escalation rate, and human satisfaction. This discipline ensures evaluation matches context and decision value.')

doc.add_page_break()

# 7. KPI Framework
add_h1(doc, '7. KPI Framework')
add_para(doc, 'The KPI framework should align technical, operational, customer, adoption, governance, and financial metrics. NovaMart can use a four-level dashboard hierarchy: technical metrics, operational metrics, business metrics, and executive strategy metrics. The purpose is to connect all outcomes to decision-making without overwhelming executives with raw model scores alone.')
add_h2(doc, 'Model KPIs')
for item in ['Accuracy', 'Precision', 'Recall', 'F1-score', 'ROC-AUC where applicable', 'MAE', 'RMSE', 'MAPE', 'WAPE', 'Forecast Bias', 'Prediction Latency', 'Model Drift']:
    doc.add_paragraph(item, style='List Bullet')
add_h2(doc, 'Business KPIs')
for item in ['Revenue uplift', 'Cost reduction', 'Inventory turnover', 'Stockout reduction', 'Overstock reduction', 'Conversion rate', 'Customer retention', 'Customer lifetime value', 'Average order value', 'Return rate', 'Customer satisfaction', 'NPS', 'Service level', 'Productivity']:
    doc.add_paragraph(item, style='List Bullet')
add_h2(doc, 'Financial KPIs')
for item in ['ROI', 'Incremental revenue', 'Cost savings', 'AI implementation cost', 'Operating cost', 'Payback period', 'Net benefit', 'Benefit-cost ratio']:
    doc.add_paragraph(item, style='List Bullet')
add_h2(doc, 'Customer KPIs')
for item in ['Customer satisfaction', 'NPS', 'Conversion rate', 'Repeat purchase rate', 'Churn', 'Recommendation engagement', 'Complaint rate', 'Customer effort score']:
    doc.add_paragraph(item, style='List Bullet')
add_h2(doc, 'Adoption KPIs')
for item in ['AI recommendation acceptance rate', 'Human override rate', 'Active users', 'Workflow adoption', 'AI feature usage', 'Training completion']:
    doc.add_paragraph(item, style='List Bullet')
add_h2(doc, 'Governance KPIs')
for item in ['Bias incidents', 'Privacy incidents', 'Security incidents', 'Explainability coverage', 'Audit completion', 'Data quality failures', 'Model drift incidents', 'SLA violations']:
    doc.add_paragraph(item, style='List Bullet')
add_h2(doc, 'KPI Hierarchy Diagram')
add_diagram_flow(doc, '', ['Strategic KPIs', 'Business KPIs', 'Operational KPIs', 'Technical KPIs'])
add_para(doc, 'Each KPI should be defined in one sentence, include a data owner, specify the time period, and include a threshold or target. This avoids ambiguity and makes it easier to compare outcomes across initiatives.')

doc.add_page_break()

# 8. Data Collection Methodology
add_h1(doc, '8. Data Collection Methodology')
add_para(doc, 'NovaMart can collect evaluation data from multiple operational systems. These are proposed systems and hypothetical assumptions for the NovaMart evaluation environment; they are not presented as actual deployed systems.')
for item in [
    'POS: transaction amount, item SKU, channel, basket size, discount behavior, and return events.',
    'E-commerce: product views, clicks, cart events, conversion funnel activity, session duration, and recommendation engagement.',
    'CRM: customer profile, loyalty state, purchase history, support interactions, segmentation, and churn risk flags.',
    'ERP: purchase orders, supplier lead times, procurement cost, margin, and inventory adjustments.',
    'WMS: inventory positions, fulfillment SLA, stock movement, cycle time, warehouse allocation, and waste events.',
    'Customer support: ticket volume, resolution time, satisfaction scores, escalation paths, and chatbot interactions.',
    'Marketing platforms: campaign performance, customer targeting, response rate, and promotion effectiveness.',
    'Model logs: prediction outputs, confidence scores, feature values, threshold decisions, and model version metadata.',
    'API logs: latency, payload errors, hallucination flags, call volume, uptime, and usage rate.',
    'Application analytics: clickstream data, workflow usage, human override rate, approval actions, and conversion funnel efficiency.',
    'Finance systems: gross margin, incremental revenue, support costs, working capital, and realized savings.',
    'Customer surveys: CSAT, NPS, CES, complaint comments, and sentiment analysis.'
]:
    doc.add_paragraph(item, style='List Bullet')
add_para(doc, 'The full data process should be: Data Collection → Validation → Cleaning → Integration → Storage → Metric Calculation → Dashboard → Decision. This process ensures that metrics are driven by trusted data and not by unsupported or stale inputs. Validation should include completeness checks, uniqueness checks, business-rule validation, and freshness thresholds. Cleaning should eliminate duplicates, reconcile missing values, and define agreed treatment of outliers. Integration should unify events from POS, CRM, marketing, WMS, and finance into a common timeline and grain. Storage should retain raw and curated data for traceability and auditing. Metric calculation should occur in a controlled pipeline with reproducible logic. Dashboard outputs should be reviewed through a governance routine and not treated as one-off snapshots.')
add_h2(doc, 'Data Collection and Measurement Pipeline Diagram')
add_diagram_flow(doc, '', ['Data Collection', 'Validation', 'Cleaning', 'Integration', 'Storage', 'Metric Calculation', 'Dashboard', 'Decision'])

doc.add_page_break()

# 9. Measurement and Analysis Methodology
add_h1(doc, '9. Measurement and Analysis Methodology')
add_para(doc, 'The measurement process begins by defining a baseline before AI is introduced. Baseline metrics should represent the historical operating state across a meaningful period of time, often 3–6 months or more depending on seasonality and business cycles. A baseline should also include key controls such as store segments, customer segments, or region groups, so that results are not misinterpreted due to changes in demand patterns.')
add_para(doc, 'After the AI implementation, NovaMart should measure the same KPIs, ideally using a defined control group and a treatment group. A control group helps isolate the effect of the AI system while accounting for general business changes. A/B testing is ideal when the intervention is customer-facing and can be randomized safely. For supply-chain or fraud scenarios, it may be more appropriate to use before/after comparison, synthetic controls, or region-based comparison with causal adjustment methods.')
add_para(doc, 'The sequence is: Baseline ↓ AI Implementation ↓ Measurement ↓ Comparison ↓ Impact Assessment ↓ Optimization. Good model accuracy does not automatically create business value, so the measurement process must evaluate whether the AI system materially changed key business outcomes beyond the baseline.')
add_h2(doc, 'Baseline → AI Deployment → Measurement → Impact Assessment Diagram')
add_diagram_flow(doc, '', ['Baseline', 'AI Implementation', 'Measurement', 'Comparison', 'Impact Assessment', 'Optimization'])
add_para(doc, 'Statistical significance should be interpreted at a practical level. Business teams usually care less about a tiny p-value than about whether the measured uplift is large enough to matter operationally and financially. This means that practical significance, minimum detectable effect, and confidence intervals should be considered alongside statistical tests. Avoid misleading attribution: improvements may occur due to marketing campaigns, seasonality, operating changes, pricing actions, or external shocks. An AI solution should only receive business credit for the incremental improvement it creates, after controlling for confounding factors and comparing against a credible baseline or control group.')

doc.add_page_break()

# 10. AI Model Performance Metrics
add_h1(doc, '10. AI Model Performance Metrics')
add_para(doc, 'Different AI systems require different evaluation metrics because the business objective differs. The following metrics are used appropriately depending on the problem type.')
add_h2(doc, 'Classification')
add_para(doc, 'For classification use cases such as churn prediction or fraud detection, metrics include accuracy, precision, recall, F1-score, and ROC-AUC. Accuracy is useful for balanced datasets but can be misleading when rare events dominate. Precision measures the proportion of positive predictions that are actually correct. Recall measures the proportion of true positives identified. F1 balances precision and recall. ROC-AUC indicates discrimination power across thresholds. These metrics are appropriate when the cost of false positives and false negatives differs materially.')
add_para(doc, 'Precision = TP / (TP + FP)')
add_para(doc, 'Recall = TP / (TP + FN)')
add_para(doc, 'F1 = 2 × Precision × Recall / (Precision + Recall)')
add_h2(doc, 'Regression')
add_para(doc, 'For demand forecasting and price forecasting, regression metrics such as MAE, RMSE, and MAPE are often used. MAE is intuitive and interpretable in the unit of the target variable. RMSE penalizes larger errors more than MAE. MAPE measures percentage error but is weak when actual values are near zero or show highly variable demand. In retail scenarios with intermittent demand or zero-sale days, WAPE or bias may be more stable and informative than MAPE alone.')
add_para(doc, 'MAPE = Mean(|Actual - Forecast| / Actual) × 100')
add_h2(doc, 'Time-Series Forecasting')
add_para(doc, 'Time-series evaluation should include MAPE, WAPE, MAE, RMSE, and forecast bias. Forecast bias identifies systematic over- or under-forecasting, which can create chronic stockouts or overstock. WAPE is often preferred when comparing across SKUs with different demand scales because it mitigates distortions from low-volume items. These metrics are suitable for replenishment, promotion planning, and capacity planning.')
add_h2(doc, 'Recommendation Systems')
add_para(doc, 'Recommendation systems are evaluated on CTR, conversion rate, precision@K, recall@K, and NDCG when ranking quality matters. CTR and conversion are closer to business value than offline ranking metrics because they reflect customer response. Precision@K assesses whether top-ranked items are relevant; NDCG rewards ranking quality and item positioning. These metrics are useful when the business objective is engagement and incremental revenue.')
add_h2(doc, 'NLP / Generative AI')
add_para(doc, 'Generative AI systems should be assessed with relevance, groundedness, factuality, hallucination rate, response quality, and response latency. Human evaluation remains essential because automated scores may miss context, ambiguity, or unsafe outputs. For customer-service assistants, the most important signals may include resolution rate, escalation rate, groundedness, and trust. In regulated business settings, explainability, user override, and human review are critical controls.')
add_h2(doc, 'Anomaly Detection')
add_para(doc, 'For fraud and operational anomaly detection, precision, recall, false-positive rate, and detection latency matter most. A lower false-positive rate reduces customer friction, while higher recall reduces financial loss. The appropriate threshold should be set based on cost trade-offs: blocking a legitimate transaction may be inconvenient, but missing a true fraud event can create direct financial loss and reputational damage.')
add_para(doc, 'The important design principle is that no single metric is universally best. The business objective and cost structure determine which metrics matter most and which trade-offs are acceptable.')

doc.add_page_break()

# 11. Operational Performance Metrics
add_h1(doc, '11. Operational Performance Metrics')
add_para(doc, 'Operational performance measures whether AI-driven workflows are improving execution quality and business efficiency. This includes inventory efficiency, fulfillment speed, labor productivity, and service delivery reliability.')
rows = [
    ['Metric', 'Definition', 'Example Target (Illustrative)', 'Why It Matters'],
    ['Inventory Turnover', 'Number of times stock is sold and replaced over a period.', '> 12 turns annually (illustrative target)', 'Improves capital efficiency and reduces carrying cost.'],
    ['Stockout Rate', 'Percentage of demand not served due to inventory shortage.', '< 2–3% for high-priority SKUs (illustrative)', 'Directly affects revenue and customer satisfaction.'],
    ['Overstock Rate', 'Excess stock as a percentage of total inventory value.', '< 15% (illustrative)', 'Signals poor forecasting and working capital burden.'],
    ['Service Level', 'Percentage of demand fulfilled on time in full.', '> 95% (illustrative)', 'Tracks customer fulfilment performance.'],
    ['Productivity', 'Output per labor hour or workflow.', '+10–20% improvement vs baseline (illustrative)', 'Measures operational efficiency gains from automation.'],
    ['Replenishment Cycle Time', 'Time between demand signal and decision or purchase order creation.', 'Reduction of 20–30% (illustrative)', 'Supports faster response to demand variation.'],
    ['Waste Rate', 'Products lost due to obsolescence or spoilage.', 'Defined threshold by category and season', 'Critical for fresh goods and low-margin categories.']
]

table = add_table_from_rows(doc, rows, widths=[1.5, 2.2, 1.8, 2.5])
for row in table.rows:
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(8)
add_para(doc, 'The business chain is simple: Better Forecast → Better Replenishment → Fewer Stockouts → Lower Overstock → Better Inventory Efficiency → Financial Impact. The value of operational metrics is that they are closer to the actual business lever than raw model accuracy alone.')

doc.add_page_break()

# 12. Customer Experience Metrics
add_h1(doc, '12. Customer Experience Metrics')
add_para(doc, 'AI systems intersect with customer experience through personalization, service interactions, fraud decisions, and digital support. Customer experience should be evaluated from both operational and commercial perspectives.')
rows = [
    ['Metric', 'Definition', 'Illustrative Threshold', 'Business Question'],
    ['Customer Satisfaction (CSAT)', 'Customer rating of the experience.', '> 4.2/5 (illustrative)', 'Did the service meet expectations?'],
    ['NPS', 'Net promoter score from satisfaction surveys.', '> 35 (illustrative)', 'How likely are customers to recommend NovaMart?'],
    ['Conversion Rate', 'Percentage of visitors who purchase.', '+2–5 percentage points uplift (illustrative)', 'Did AI improve sales?'],
    ['Repeat Purchase Rate', 'Share of customers making multiple orders.', '+5–10% vs baseline (illustrative)', 'Did AI create retention?'],
    ['Recommendation CTR', 'Clicks on AI-suggested products or content.', '> 3–6% (illustrative)', 'Did the personalization engine attract attention?'],
    ['Complaint Rate', 'Share of customers raising issues after AI-assisted interactions.', 'Below defined threshold', 'Did the experience create friction?'],
    ['Customer Effort Score (CES)', 'How easy the customer experience was to navigate.', 'Lower score indicates better experience', 'Was the experience easy or frustrating?']
]
table = add_table_from_rows(doc, rows, widths=[1.7, 2.2, 1.4, 2.2])
for row in table.rows:
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(8)
add_para(doc, 'Customer experience metrics are not vanity metrics. If an AI recommendation engine raises CTR but lowers average order value or introduces irrelevant product suggestions, the business must treat the system as partially unsuccessful. The correct measure is customer value created, not only engagement created.')

doc.add_page_break()

# 13. Financial and ROI Metrics
add_h1(doc, '13. Financial and ROI Metrics')
add_para(doc, 'Financial evaluation is the formal test of whether AI adds real economic value. It should be built from a clear cost-benefit model with direct, indirect, tangible, and intangible benefits identified explicitly. This is necessary because some AI benefits are easy to measure in revenue and cost terms, while others affect customer sentiment or strategic differentiation and are harder to quantify.')
add_para(doc, 'AI Costs: initial implementation, infrastructure, model or API costs, data engineering, employee training, maintenance, monitoring, governance, and integration.')
add_para(doc, 'Benefits: revenue increase, cost savings, productivity gains, loss prevention, customer retention, lower service cost, and faster decision-making.')
add_para(doc, 'ROI = ((Financial Benefit - AI Investment Cost) / AI Investment Cost) × 100')
add_para(doc, 'Revenue Uplift % = ((Revenue After AI - Baseline Revenue) / Baseline Revenue) × 100')
add_para(doc, 'Examples in a hypothetical NovaMart business case may include:')
for item in [
    'Initial implementation cost: Hypothetical Example cost of INR 22 million for forecasting, personalization, and data integration.',
    'Annual operating cost: Hypothetical Example INR 8 million for cloud services, monitoring, and model maintenance.',
    'Annual direct benefit: Hypothetical Example INR 45 million in inventory cost savings and revenue uplift.',
    'Payback period: Hypothetical Example 10–14 months depending on adoption and seasonality.',
    'ROI: Hypothetical Example 120% over the measurement period, assuming the benefit is strong and attributable.'
]:
    doc.add_paragraph(item, style='List Bullet')
add_para(doc, 'The most important caution is that cost and benefit estimates must be clearly labelled as hypothetical and not treated as actual NovaMart results. Financial value should be estimated using incremental attribution, not just total business revenue, and should distinguish direct benefit from indirect or strategic value.')

doc.add_page_break()

# 14. Business Impact Assessment
add_h1(doc, '14. Business Impact Assessment')
add_para(doc, 'Impact assessment looks beyond model scores and asks whether the AI solution changes business performance in meaningful ways. The assessment should cover technical impact, operational impact, financial impact, customer impact, employee impact, strategic impact, and risk/compliance impact.')
rows = [
    ['Impact Area', 'Baseline', 'Target', 'Measured Result', 'Interpretation', 'Business Action'],
    ['Technical Impact', 'Current model quality or manual process baseline', 'Target metric thresholds', 'Model quality vs holdout and production performance', 'Determines if the technical system is stable and useful', 'Retrain, tune features, adjust deployment settings'],
    ['Operational Impact', 'Current workflow output and labor productivity', 'Improved cycle time or reduced effort', 'Time saved, task throughput, service level', 'Shows whether work is improved in practice', 'Redesign workflow, train users, standardize decisions'],
    ['Financial Impact', 'Revenue and cost baseline without AI', 'Positive ROI and reduced cost per unit', 'Incremental revenue, cost savings, payback period', 'Determines whether value outweighs cost', 'Scale, optimize, or discontinue solution'],
    ['Customer Impact', 'Historical customer service and conversion baseline', 'Improved conversion and reduced friction', 'CSAT, NPS, repeat purchase, complaint rate', 'Shows whether customers benefit in practice', 'Refine user experience, decision thresholds, and escalation'],
    ['Employee Impact', 'Manual or legacy decision process', 'Higher productivity and confidence', 'Time spent on high-value tasks', 'Captures human-factor implications', 'Train staff and align incentives'],
    ['Strategic Impact', 'Current strategic measures without AI', 'Longer-term competitive advantage', 'Portfolio or category performance', 'Assesses business direction and market relevance', 'Adjust portfolio prioritization'],
    ['Risk/Compliance Impact', 'Without governance and controls', 'No critical incidents; compliance maintained', 'Bias, privacy, SLA, and audit outcomes', 'Determines whether risk is acceptable', 'Restrict usage, improve governance, or halt rollout']
]
table = add_table_from_rows(doc, rows, widths=[1.3, 1.2, 1.0, 1.6, 1.5, 1.6])
for row in table.rows:
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(8)
add_para(doc, 'This format creates a structured business impact story: each metric is interpreted in context and tied to a business action. It ensures that management decisions are made from a balanced scorecard rather than a narrow technical lens.')

doc.add_page_break()

# 15. Risk and Governance Metrics
add_h1(doc, '15. Risk and Governance Metrics')
add_para(doc, 'Governance evaluation is essential because AI systems can create regulatory, fairness, privacy, and reputational risk even when they improve model performance. Evaluation should include controls for bias, privacy, security, explainability, deployment safety, and operational continuity.')
rows = [
    ['Metric', 'Definition', 'Why It Matters', 'Target/Threshold'],
    ['Bias Incidents', 'Observed unfair or discriminatory model behavior across protected groups or segments.', 'Protects fairness and trust.', 'Zero critical incidents'],
    ['Privacy Incidents', 'Unauthorized access, misuse, or leakage of data.', 'Protects compliant data handling and customer trust.', 'Zero material incidents'],
    ['Security Incidents', 'Cybersecurity or model-security breaches affecting the system.', 'Prevents large operational and reputational impact.', 'Zero critical breaches'],
    ['Explainability Coverage', 'Share of model decisions with interpretable explanations or rationale.', 'Supports auditability and governance.', '> 90% for high-risk decisions'],
    ['Audit Completion', 'Percentage of required governance review steps completed on time.', 'Ensures accountability and control vibrancy.', '100% of required audits'],
    ['Data Quality Failures', 'Missing, stale, or invalid data causing incorrect model behavior.', 'Reduces risk of poor operational decisions.', '< 1% critical failures'],
    ['Model Drift Incidents', 'Occurrence of major drift or performance degradation beyond thresholds.', 'Supports retraining and preventive maintenance.', 'No unreviewed major drift events'],
    ['SLA Violations', 'Missed service-level commitments for latency or uptime.', 'Protects operational reliability and trust.', 'Zero critical SLA violations']
]
table = add_table_from_rows(doc, rows, widths=[1.5, 2.3, 2.0, 1.2])
for row in table.rows:
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(8)
add_para(doc, 'Risk and governance decisions must be tied to operational safeguards, including human review for sensitive cases, model changes controlled through approvals, and access management for customer data.')

doc.add_page_break()

# 16. Continuous Improvement Framework
add_h1(doc, '16. Continuous Improvement Framework')
add_para(doc, 'Assessment does not end when a model is deployed. AI systems require a continuous feedback loop because data drift, feature changes, workflow changes, and user behavior can all shift performance over time.')
add_h2(doc, 'Continuous Improvement Loop Diagram')
add_diagram_flow(doc, '', ['Measure', 'Analyze', 'Identify Gap', 'Root Cause', 'Improve Model / Data / Process', 'Deploy Change', 'Measure Again'])
add_para(doc, 'The continuous improvement cycle should include retraining when drift or decline is detected, threshold changes when the business cost structure changes, feature improvements when more predictive information becomes available, data quality improvements to repair weak inputs, user training to improve adoption and trust, workflow changes to ensure operational processes can act on AI outputs, and A/B testing to validate new logic before broad deployment. The goal is to avoid static implementation and instead maintain a disciplined learning system.')

doc.add_page_break()

# 17. Case Study Scenarios
add_h1(doc, '17. Case Study Scenarios')
add_h2(doc, 'Case Study 1: AI Demand Forecasting and Inventory Optimization')
add_para(doc, 'Business Problem: NovaMart needs improved demand planning to avoid stockouts during peak periods while reducing slow-moving inventory and excess working capital.')
add_para(doc, 'AI Solution: A forecasting model combined with replenishment optimization across categories and stores using sales history, promotions, weather, price, and SKU-level signals.')
add_para(doc, 'Baseline: Legacy forecasting with manual review and simple moving average methods, resulting in inconsistent replenishment decisions.')
add_para(doc, 'KPIs: MAPE, WAPE, forecast bias, stockout rate, overstock rate, inventory turnover, service level, working capital, waste.')
add_para(doc, 'Data Sources: POS, ERP, WMS, supplier lead time data, promotional calendars, and historical demand files.')
add_para(doc, 'Evaluation Method: Before/after comparison across selected categories with seasonality adjustments and a control set of non-optimized categories or stores.')
add_para(doc, 'Measurement Period: 12 weeks post-rollout, with monthly review and seasonality adjustment.')
add_para(doc, 'Target/Threshold: MAPE improvement, service level > 95%, stockout reduction, and inventory turnover increase, all as illustrative targets.')
add_para(doc, 'Hypothetical Example Result: “Hypothetical Example” forecast MAPE improves by 18%, stockout rate falls by 22%, and working capital tied to slow-moving inventory reduces by 12% over a 12-week evaluation period. These figures are illustrative only and are not actual NovaMart measured results.')
add_para(doc, 'Business Impact: Improved forecast quality enables better replenishment decisions, fewer stockouts, and less waste. This also improves customer service and reduces tied-up working capital.')
add_para(doc, 'Risks: Data quality issues, anomalous promotions, demand shocks, supplier delays, and human override of system recommendations.')
add_para(doc, 'Decision Criteria: Continue if service level improves without major waste or excess overstock; otherwise refine replenishment logic and data quality.')
add_para(doc, 'Recommended Action: Expand AI forecasting to additional categories and create exception-based human review for outlier SKUs.')

add_h2(doc, 'Case Study 2: Real-Time Customer Personalization')
add_para(doc, 'Business Problem: NovaMart wants to improve customer discovery and conversion using personalized products and content across web and mobile channels.')
add_para(doc, 'AI Solution: Recommendation engine using collaborative filtering, session-based personalization, lifecycle stage, and product affinity signals.')
add_para(doc, 'Baseline: Generic homepage and email recommendations with limited personalization.')
add_para(doc, 'KPIs: CTR, conversion rate, average order value, revenue per user, repeat purchase, engagement rate.')
add_para(doc, 'Data Sources: E-commerce clickstream, CRM, order history, product metadata, campaign analytics, and session logs.')
add_para(doc, 'Evaluation Method: A/B test comparing control group vs AI personalization group for comparable traffic segments.')
add_para(doc, 'Measurement Period: 6–8 weeks, with daily monitoring and weekly review.')
add_para(doc, 'Target/Threshold: CTR uplift and conversion lift above baseline with no significant decline in margin or customer satisfaction.')
add_para(doc, 'Hypothetical Example Result: “Hypothetical Example” AI personalization group shows a 7% increase in conversion rate and a 9% increase in average order value relative to the control group. This is illustrative only and not a measured NovaMart output.')
add_para(doc, 'Business Impact: Higher conversion and revenue per user can outweigh the incremental model and infrastructure cost if recommendation quality remains stable.')
add_para(doc, 'Risks: Irrelevant recommendations, reduced trust, margin dilution, and poor experiences for new customers.')
add_para(doc, 'Decision Criteria: Maintain if incremental revenue exceeds cost and customer satisfaction remains stable or improves.')
add_para(doc, 'Recommended Action: Use layered personalization with guardrails on margin, freshness, and customer segment relevance.')

add_h2(doc, 'Case Study 3: AI Fraud Detection')
add_para(doc, 'Business Problem: NovaMart needs to reduce fraudulent transactions while minimizing false positives that block legitimate customers.')
add_para(doc, 'AI Solution: Real-time anomaly detection model using transaction patterns, shipping data, device data, customer behavior, and historical fraud markers.')
add_para(doc, 'Baseline: Rule-based fraud screening with manual review and delayed checks.')
add_para(doc, 'KPIs: Precision, recall, false positive rate, fraud loss prevented, detection latency, customer friction rate.')
add_para(doc, 'Data Sources: POS, e-commerce transactions, customer profiles, payment logs, fraud investigation records, and customer support tickets.')
add_para(doc, 'Evaluation Method: Backtesting and shadow deployment with threshold tuning based on cost matrix for false positives and false negatives.')
add_para(doc, 'Measurement Period: 30–90 days, using daily threshold review and monthly governance review.')
add_para(doc, 'Target/Threshold: High recall for high-value fraud patterns while keeping false positive rate below an agreed threshold.')
add_para(doc, 'Hypothetical Example Result: “Hypothetical Example” model detects 80% of fraud attempts with a false positive rate of 1.5%; this could be materially better than the rule-only baseline. The numbers are illustrative and not actual NovaMart results.')
add_para(doc, 'Business Impact: Fraud loss prevented and improved operational efficiency for fraud analysts, but the business must also consider customer friction and churn risk from false positives.')
add_para(doc, 'Risks: Over-blocking legitimate orders, case backlog, unfair treatment of certain segments, and poor explainability in disputed cases.')
add_para(doc, 'Decision Criteria: Adopt only if loss reduction outweighs customer friction and compliance risk.')
add_para(doc, 'Recommended Action: Use a risk-tiered response strategy and preserve human review for high-risk or sensitive cases.')

add_h2(doc, 'Case Study 4: AI Customer Churn Prediction')
add_para(doc, 'Business Problem: NovaMart wants to retain valuable customers and reduce churn in segments with high lifetime value.')
add_para(doc, 'AI Solution: Predictive retention model using purchase frequency, recency, customer service history, product affinity, and engagement signals.')
add_para(doc, 'Baseline: Reactive retention campaigns and loyalty outreach after a customer has already become inactive.')
add_para(doc, 'KPIs: Precision, recall, retention rate, intervention acceptance, customer lifetime value, incremental retained revenue.')
add_para(doc, 'Data Sources: CRM, POS, loyalty program data, email engagement, support records, and customer segmentation files.')
add_para(doc, 'Evaluation Method: Retention pilot comparing AI-identified customers against a matched historical baseline or control group.')
add_para(doc, 'Measurement Period: 8–12 weeks with monthly follow-up on the retention cohort.')
add_para(doc, 'Target/Threshold: Intervention must improve retention or incremental revenue enough to compensate for outreach costs.')
add_para(doc, 'Hypothetical Example Result: “Hypothetical Example” retention campaign for high-risk customers produces a 6% improvement in 90-day retention and an incremental retained revenue uplift considered financially worthwhile, but the exact values are illustrative only.')
add_para(doc, 'Business Impact: Better customer retention and improved profitability, but the system only creates value when interventions are executed well and customers respond to offers or service actions.')
add_para(doc, 'Risks: Over-targeting customers, customer annoyance, inappropriate incentives, poor messaging, and lack of actionability for frontline teams.')
add_para(doc, 'Decision Criteria: Expand only if the lift in retained revenue exceeds campaign cost and experimental evidence shows repeatable benefit.')
add_para(doc, 'Recommended Action: Pair the model with targeted incentives, customer service playbooks, and personalized recovery journeys.')

add_h2(doc, 'Case Study 5: GenAI Conversational Commerce / Customer Assistant')
add_para(doc, 'Business Problem: NovaMart wants to improve customer assistance for product discovery, order status, returns, and general shopping queries at scale.')
add_para(doc, 'AI Solution: GenAI-powered customer assistant connected to product catalog, order information, policy rules, and escalation workflows.')
add_para(doc, 'Baseline: Human support agents and static FAQ pages with slower resolution and inconsistent availability.')
add_para(doc, 'KPIs: Resolution rate, customer satisfaction, first-contact resolution, escalation rate, hallucination rate, groundedness, response latency, cost per interaction.')
add_para(doc, 'Data Sources: Customer support tickets, conversation logs, product catalog, order database, policy documents, and chatbot analytics.')
add_para(doc, 'Evaluation Method: Human review, red-team testing, business policy checks, and measured outcomes across live and simulated interactions.')
add_para(doc, 'Measurement Period: 30–60 days with continuous review of quality, latency, and escalation patterns.')
add_para(doc, 'Target/Threshold: High groundedness, low hallucination, acceptable latency, and strong customer satisfaction without excessive unsafe or policy-violating response patterns.')
add_para(doc, 'Hypothetical Example Result: “Hypothetical Example” GenAI assistant resolves 45% of routine support interactions without escalation and maintains acceptable customer satisfaction levels, while high-risk or sensitive cases are escalated to human agents. These values are illustrative and not a claim of actual NovaMart performance.')
add_para(doc, 'Business Impact: Faster responses, lower support cost, and more customer convenience, but only if groundedness and escalation quality are maintained.')
add_para(doc, 'Risks: Hallucinated product details, inaccurate order information, policy non-compliance, and poor handoff quality for sensitive cases.')
add_para(doc, 'Decision Criteria: Continue if response quality remains high and escalation remains low for routine cases while risky cases are escalated responsibly.')
add_para(doc, 'Recommended Action: Add human escalation hooks, grounding, and strict policy guardrails before scaling to broader customer traffic.')

add_h2(doc, 'Case Study Evaluation Flow Diagram')
add_diagram_flow(doc, '', ['Business Problem', 'AI Solution', 'Baseline', 'KPIs', 'Measurement', 'Impact', 'Decision'])

doc.add_page_break()

# 18. Scenario Evaluation
add_h1(doc, '18. Scenario Evaluation')
add_para(doc, 'The evaluation of each scenario should answer the same set of business questions: Is the model accurate enough? Is the workflow operationally practical? Is the solution adopted? Does the outcome improve customer value? Does the financial impact justify the investment? Does governance risk remain acceptable?')
rows = [
    ['Scenario', 'Primary Metric', 'Business Decision', 'Illustrative Evaluation Conclusion'],
    ['Demand Forecasting', 'MAPE / WAPE / Service level', 'Scale if stockout and waste improve with acceptable cost', 'Illustrative: improves planning efficiency but needs stronger SKU exceptions'],
    ['Personalization', 'CTR / Conversion / Revenue per user', 'Scale if conversion uplift exceeds cost and satisfies quality guardrails', 'Illustrative: promising if basket size and satisfaction remain positive'],
    ['Fraud Detection', 'Precision / Recall / False positive rate', 'Scale if losses prevented outweigh blocked-legitimate trade-offs', 'Illustrative: viable with layered review for edge cases'],
    ['Churn Prediction', 'Precision / Retention / Incremental retained revenue', 'Scale if customer interventions generate measurable retention value', 'Illustrative: positive when outreach is well calibrated'],
    ['GenAI Assistant', 'Resolution rate / Hallucination / Satisfaction', 'Scale if governance and latency remain within acceptable thresholds', 'Illustrative: use for routine queries while escalating complex cases']
]
table = add_table_from_rows(doc, rows, widths=[1.4, 1.7, 1.8, 2.5])
for row in table.rows:
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(8)
add_para(doc, 'Scenario evaluation requires not only business intuition but also evidence from controlled experiments, data quality checks, and an informed view of what good performance means in context.')

doc.add_page_break()

# 19. Evaluation Dashboard
add_h1(doc, '19. Evaluation Dashboard')
add_para(doc, 'A proposed dashboard for NovaMart should provide a concise, decision-oriented view of AI performance and business impact. The dashboard should be separated into executive summary, model health, business KPIs, financial impact, customer impact, AI adoption, risk/governance, and alerts. The following structure is illustrative and proposed.')
rows = [
    ['Dashboard Area', 'KPIs and Signals', 'Status', 'Owner'],
    ['Executive Summary', 'Portfolio ROI, strategic alignment, adoption, risk exposure', 'Green / Amber / Red', 'CEO / AI Office'],
    ['Model Health', 'Accuracy, precision, recall, MAPE, drift, latency', 'Stable / Monitor / Critical', 'Data Science Lead'],
    ['Business KPIs', 'Stockout rate, service level, conversion, retention', 'Trend vs baseline', 'Business Unit Head'],
    ['Financial Impact', 'Incremental revenue, cost savings, ROI, payback period', 'Quarterly review', 'CFO'],
    ['Customer Impact', 'CSAT, NPS, complaint rate, repeat purchase', 'Customer sentiment and trend', 'Customer Experience Lead'],
    ['AI Adoption', 'Usage, active users, overrides, training completion', 'Usage rate vs target', 'AI Product Owner'],
    ['Risk/Governance', 'Bias incidents, privacy incidents, SLA breaches, audit status', 'Policy compliance watchlist', 'Chief Risk Officer'],
    ['Alerts', 'Drift events, data quality failures, SLA breaches, spike in false positives', 'Immediate action thresholds', 'Operations / Engineering']
]
table = add_table_from_rows(doc, rows, widths=[1.3, 2.5, 1.1, 1.3])
for row in table.rows:
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(8)
add_h2(doc, 'AI Performance Dashboard Mockup')
add_diagram_flow(doc, '', ['Executive Summary', 'Model Health', 'Business KPIs', 'Financial Impact', 'Customer Impact', 'AI Adoption', 'Risk/Governance', 'Alerts'])
add_para(doc, 'The dashboard should be filtered by initiative, region, category, or customer segment. A dashboard is valuable only when it helps leaders decide whether to expand, pause, optimize, or restructure an AI system.')

doc.add_page_break()

# 20. Reporting and Decision-Making Framework
add_h1(doc, '20. Reporting and Decision-Making Framework')
add_para(doc, 'The reporting cadence should match both operational urgency and strategic importance. Daily reporting supports operational intervention; weekly reporting supports model quality and workflow review; monthly reporting supports value and adoption tracking; quarterly reporting supports portfolio-level return and strategy decisions.')
rows = [
    ['Metric Category', 'Owner', 'Frequency', 'Audience', 'Required Action'],
    ['Operational alerts and model health', 'ML Engineer / Operations Lead', 'Daily', 'Operations and engineering teams', 'Resolve outages, data breakages, or model quality degradation'],
    ['Model performance', 'Data Science Lead', 'Weekly', 'AI product and business owners', 'Review drift, threshold changes, and retraining triggers'],
    ['Business KPI review', 'Business Unit Head', 'Monthly', 'Leadership and category managers', 'Act on variances and improve operational execution'],
    ['ROI and strategic impact', 'CFO / CEO', 'Quarterly', 'Board and executive leadership', 'Decide on scaling, re-prioritization, or exit strategies']
]
table = add_table_from_rows(doc, rows, widths=[1.7, 1.1, 1.0, 1.2, 2.2])
for row in table.rows:
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(8)
add_para(doc, 'Decision-making should be tied to thresholds and actions. For example, if forecast MAPE exceeds a business threshold for two consecutive weeks, action is required. If model drift rises above policy threshold, retraining is triggered. If AI adoption remains low, targeted training and workflow redesign must be performed. This ensures reporting creates action, not merely reporting output.')

doc.add_page_break()

# 21. 30/60/90-Day Evaluation Plan
add_h1(doc, '21. 30/60/90-Day Evaluation Plan')
add_h2(doc, 'First 30 Days')
for item in [
    'Define baselines for each AI initiative and identify relevant pre-AI historical periods.',
    'Identify data sources, owners, schema mappings, data gaps, and quality constraints.',
    'Establish KPI definitions, target thresholds, and a common measurement taxonomy across functions.',
    'Build the measurement plan, including control-group design, experiment rules, and governance expectations.'
]:
    doc.add_paragraph(item, style='List Bullet')
add_h2(doc, 'Days 31–60')
for item in [
    'Build the metric pipeline and validation logic for model, operational, customer, and financial metrics.',
    'Create the dashboard and review cadence with business owners and technology partners.',
    'Start initial evaluations and monitor the system against agreed thresholds and guardrails.',
    'Validate data quality and investigate drift, missingness, or pipeline failures.'
]:
    doc.add_paragraph(item, style='List Bullet')
add_h2(doc, 'Days 61–90')
for item in [
    'Conduct the first formal business impact assessment across technical, operational, customer, and financial dimensions.',
    'Analyze gaps between expected and realized value and identify root causes.',
    'Optimize model logic, data quality, workflow process, or user training as needed.',
    'Present results to leadership with recommended actions for scale, redesign, or pause.'
]:
    doc.add_paragraph(item, style='List Bullet')
add_h2(doc, '30/60/90-Day Roadmap Diagram')
add_diagram_flow(doc, '', ['Day 1–30', 'Day 31–60', 'Day 61–90'],)

doc.add_page_break()

# 22. Long-Term Evaluation Roadmap
add_h1(doc, '22. Long-Term Evaluation Roadmap')
add_h2(doc, '0–3 Months')
add_para(doc, 'Establish KPI definitions, data lineage, baseline metrics, experiment design, and dashboard foundations. Focus on operational health and basic outcome measurement.')
add_h2(doc, '3–6 Months')
add_para(doc, 'Expand to end-to-end measurement, compare business impact against baselines, refine governance controls, and improve model retraining cadence.')
add_h2(doc, '6–12 Months')
add_para(doc, 'Integrate value-based performance management across business units, increase automation of exception handling, and improve adoption metrics and user workflows.')
add_h2(doc, '12–18 Months')
add_para(doc, 'Advance to continuous AI impact management: portfolio optimization, scenario planning, causal value estimation, enterprise governance, and strategic value review.')
add_para(doc, 'NovaMart can evolve from basic KPI monitoring to a mature AI impact management system where performance evaluation is treated as a continuous business capability, not a one-time project activity.')

doc.add_page_break()

# 23. Conclusion
add_h1(doc, '23. Conclusion')
add_para(doc, 'The success of AI in a business environment depends on disciplined evaluation, not on isolated model scores. For NovaMart, an AI initiative should be judged by whether it improves forecasting accuracy, inventory performance, customer experience, financial value, and governance outcomes in a measurable and sustainable way. Technical performance is necessary but not sufficient. Operational execution, user adoption, cost-benefit realism, and risk control are equally important.')
add_para(doc, 'A comprehensive performance and impact assessment framework provides the organization with a consistent language for decision-making. It aligns technical teams with business owners, supports evidence-based management, and makes continuous improvement possible. By using a structured evaluation process and by tying AI outcomes to business decisions, NovaMart can move from experimentation to accountable and scalable value creation.')

doc.add_page_break()

# 24. References
add_h1(doc, '24. References')
for entry in [
    'Géron, A. (2022). Hands-On Machine Learning with Scikit-Learn, Keras, and TensorFlow (3rd ed.). O\'Reilly Media.',
    'Hyndman, R.J. and Athanasopoulos, G. (2021). Forecasting: Principles and Practice (3rd ed.). OTexts.',
    'Herlocker, J.L. et al. (2004). “Evaluating Collaborative Filtering Recommender Systems.” ACM Transactions on Information Systems, 22(1), 5–53.',
    'Microsoft. (2024). Evaluation of Generative AI Applications. Azure AI documentation.',
    'Google Cloud. (2024). Vertex AI Model Evaluation. Google Cloud documentation.',
    'Brynjolfsson, E. and McElheran, K. (2016). “The Rapid Adoption of Data-Driven Decision-Making.” American Economic Review, 106(5), 133–139.',
    'Davenport, T.H. and Ronanki, R. (2018). “Artificial Intelligence for the Real World.” Harvard Business Review, Jan–Feb 2018.',
    'McKinsey Global Institute. (2023). The Economic Potential of Generative AI. McKinsey & Company.',
    'NIST. (2023). AI Risk Management Framework (AI RMF 1.0). National Institute of Standards and Technology.',
    'IBM. (2024). AI Fairness 360. IBM Research.',
    'European Commission. (2019). Ethics Guidelines for Trustworthy AI. High-Level Expert Group on AI.',
    'Kohavi, R., Tang, D., and Xu, Y. (2020). Trustworthy Online Controlled Experiments. Cambridge University Press.',
    'Kaplan, R.S. and Norton, D.P. (1996). The Balanced Scorecard. Harvard Business School Press.'
]:
    doc.add_paragraph(entry)

# Save
out = r'C:\Users\shrey\OneDrive\Desktop\Ai1\WEEK_4\documents\Task_4_Performance_Evaluation_Impact_Framework_NovaMart_FINAL.docx'
doc.save(out)
print('Saved:', out)
