# HR Assistant (Streamlit)

A branch-HR toolkit with four modules:

1. **Employee Database** (`Home.py`) — upload your master sheet (first row = headers), search/filter/browse.
2. **Long Absence Tracker** — upload an attendance muster, auto-counts every `A` per employee and buckets them:
   - 3–4 days → Informed Leave
   - 5–20 days → To Be Checked
   - 21+ days → Probable Exit Case
   (All thresholds are adjustable in the app.)
3. **Incentive Validator** — bulk-upload branch data using columns:
   `Branch Name, Bus Maid Count, Washroom Maid Count, Max Washroom Incentive, Max Bus Incentive`
   A downloadable sample template is built into the page.
4. **Payroll Calculator**
   - Full-time employees: Gross = Basic + HRA + Other Allowances; PF (12%/12%, wage capped at ₹15,000 by default); ESIC (0.75%/3.25%, applicable if Gross ≤ ₹21,000).
   - Gig workers: Monthly Billing Amount = Payment Amount + 1% TDS (adjustable %).

## ⚠️ Assumptions you should confirm before relying on this for real payroll/incentive decisions

- **Incentive formula**: since the specific calculation rule wasn't provided, the app uses
  `Incentive = Maid Count × Rate per Maid`, capped at each branch's Max Incentive. The
  per-maid rate is an editable input in the app (defaults to ₹500). If your actual policy
  is different (e.g. slab-based, fixed amounts per band), edit `compute_incentives()` in `utils.py`.
- **PF/ESIC rates & thresholds**: current commonly-used rates are pre-filled (PF 12%/12% on
  Basic capped at ₹15,000; ESIC 0.75%/3.25% below ₹21,000 gross) but all are editable
  in the sidebar of the Payroll page — please verify against the latest statutory rules
  before using this for actual compliance/filing.
- **Gig worker TDS**: implemented literally as stated — "1% TDS added to the billing amount"
  → `Billing = Amount + (Amount × TDS%)`. If your policy instead means the worker should
  *net* a fixed amount after the payer deducts TDS from the billing amount, the formula
  would need to be a gross-up (`Billing = Amount / (1 - TDS%)`) instead — flag this if that's the case.

## Running locally

```bash
pip install -r requirements.txt
streamlit run Home.py
```

## File structure

```
hr_assistant/
├── Home.py                          # Employee Database (landing page)
├── utils.py                         # All calculation logic lives here
├── requirements.txt
└── pages/
    ├── 1_Long_Absence_Tracker.py
    ├── 2_Incentive_Validator.py
    └── 3_Payroll_Calculator.py
```

Streamlit automatically turns the `pages/` folder into sidebar navigation —
no extra routing code needed.