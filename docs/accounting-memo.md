# Fictional corporate cash-flow case: classification memo

Case period: July 1, 2024–June 30, 2025. Unit: actual US dollars.
This memo explains the authored case; it is not an accounting opinion.

## Reporting boundary and assumptions

Harbor Tools Inc. owns 100% of Harbor Parts LLC for the entire period. Both are
fictional domestic nonfinancial entities using US GAAP and USD. There is no
foreign currency, acquisition or disposal of a business, equity-method income,
lease activity, discontinued operation, or noncontrolling interest. There are no
cash equivalents or restricted cash equivalents; the cash-total policy would
include them if present. Cash flows are shown using the indirect method.
The two entities' stated net income excludes duplicative equity-method pickup.

The only intercompany activity is a $20,000 loan advanced by the parent and
received by the subsidiary. It has no interest or repayment in this period.
Entity support classifies the lender's outflow as investing and the borrower's
inflow as financing. Consolidation eliminates both legs in their respective
sections, rather than leaving gross internal activity merely because the net
cash effect is zero. The due-from/due-to balances are excluded from the
**external** debt rollforward. There are no intercompany sales, profits,
dividends or other intragroup balances in this limited case.

## Operating, investing and financing activity

Consolidated net income of $105,000 plus depreciation of $30,000, less increases
in trade receivables of $16,000 and inventory of $7,000, plus an increase in trade
payables of $18,000, gives operating cash flow of $130,000. The stated working
capital changes exclude equipment purchases, note proceeds, restricted-cash
transfers, and intercompany lending. They are assumed to have no noncash or
acquired-balance component.

Cash equipment purchases are $80,000, classified as investing outflows. External
cash borrowing of $50,000 less principal repayments of $20,000 and external
dividends of $10,000 gives financing inflows of $20,000. Cash interest paid of
$5,000 and cash income taxes paid of $10,000 are operating amounts already
reflected in net income and the stated changes. They are disclosed
supplementally, with no second cash-flow adjustment.

The governing framework is [FASB ASC Topic 230, Statement of Cash Flows](https://asc.fasb.org/).
ASC 230-10-45 classifies operating, investing and financing cash flows; ASC
230-10-45-28 describes the indirect reconciliation. The original
[FASB Statement 95](https://storage.fasb.org/aop_fas95.pdf), paragraphs 15–23 and
28–32, gives publicly readable historical explanation and a consolidating
intercompany-loan illustration. It is a historical source, not the current
governing standard.

## Noncash equipment financed by seller note

The parent receives equipment by issuing a $30,000 seller note directly to the
seller. No cash passes through either entity. This is an asset purchase, not a
business combination, and carries no initial interest component in this case.
It is disclosed as noncash investing and financing activity under ASC
230-10-50-3–6. It is excluded from both cash equipment purchases and cash debt
proceeds. A total-cash bridge alone would not catch inclusion of equal amounts
in both sections, so the case separately checks those sections' cash-only amounts.

External debt reconciles $120,000 beginning + $50,000 cash borrowing − $20,000
principal paid + $30,000 noncash seller note = $180,000 ending. Net equipment
reconciles $300,000 beginning + $80,000 cash purchases + $30,000 noncash purchases
− $30,000 depreciation = $380,000 ending. These are support schedules, not a
complete balance sheet or acquisition-accounting model.

## Cash and restricted cash

Cash restricted as collateral under a fictional lending covenant is presented
separately from unrestricted cash in the assumed balance-sheet captions. The
report reconciles both components to the cash-flow total at each period end:

| Component | Beginning | Ending | Change |
|---|---:|---:|---:|
| Unrestricted cash | 135,000 | 185,000 | 50,000 |
| Restricted cash | 15,000 | 35,000 | 20,000 |
| Combined total | 150,000 | 220,000 | 70,000 |

The parent transfers $15,000 and the subsidiary transfers $5,000 from unrestricted
to restricted accounts during the year. These transfers have no effect on the
combined cash total and are excluded from all activity sections. The restriction
and beginning/ending components are disclosed. This follows ASC
230-10-45-4–5 and 230-10-50-7–8, whose added text is publicly available in
[FASB ASU 2016-18](https://storage.fasb.org/ASU_2016-18.pdf). That amendment's
text supports the cash-total/reconciliation rule; its historical transition
dates are not evidence that this is a new 2025 policy.

## Residuals, controls and limitations

Twelve exact arithmetic controls cover entity cash bridges, matched internal loan
legs, elimination of internal activity, cash-only equipment/debt proceeds,
combined cash reconciliation, external debt and equipment rollforwards. A
separate source reconciliation compares the ledger ending total with a
**fictional separate bank-control input**, not an independent bank confirmation.
Both inputs were authored for this case. A source difference is retained as a
numeric amount; the code neither plugs it nor treats a cent as rounding noise.

The hand-authored expected table is a regression oracle for this fixed example,
not an accounting standards engine. The system cannot establish completeness
of unknown cash movements, controls over actual banks, or the correctness of
new classification judgments. Source formula caches are consumed as supplied.
Accounting approval, source reconciliation, compiler arithmetic, generated
Office files and spreadsheet-engine recalculation are separate facts.
