# Notebook methodology revision implementation plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Repair the project notebook so its data audit, EDA, product analysis, text baseline, and evaluation support the conclusions in the report.

**Architecture:** Keep `Online_Retail_Repeat_Purchase_Project.ipynb` as the single analysis entry point. Add reproducible input handling and an EDA stage, separate merchandise purchases from transaction charges, replace raw sparse TF-IDF clustering with an LSA-based benchmark, and add customer-grouped uncertainty estimates. Enforce the resulting contracts with a small pytest suite and a clean top-to-bottom notebook run.

**Tech Stack:** Python 3.12, pandas, NumPy, scikit-learn, OpenPyXL, matplotlib, seaborn, sentence-transformers, Jupyter, pytest

**Spec:** `elec3544_proposal.pdf`, `group-project-assessment-standard.pdf`, `group-project-checkpoint.md`, `Online_Retail_Repeat_Purchase_Project.ipynb`, `Detailed_Project_Pipeline_Aligned.docx`, and `WORK_SPLIT.md`

## Global constraints

- Use `data/online+retail.zip` and the `Online Retail` sheet. The archive has 541,909 invoice lines from 1 December 2010 to 9 December 2011.
- Keep the eight snapshot dates on 31 March, 30 April, 31 May, 30 June, 31 July, 31 August, 30 September, and 31 October 2011.
- Features use the 90 days ending at the end of the snapshot day. Labels use the following 30 days.
- No feature may read a transaction from the target window. The target-window transactions may create `RepeatPurchase30D` only.
- `CustomerID` is an identifier and grouping key. It must never enter a predictor matrix.
- PR-AUC is the primary metric. ROC-AUC, F1, and lift at the top 20% are secondary.
- A difference in PR-AUC is not evidence of improvement until a paired customer-grouped uncertainty interval excludes zero.
- The main text comparison must use an adequately reduced TF-IDF representation and a matched-cluster-count sensitivity comparison.
- Postage, carriage, manual adjustments, fees, and gift vouchers count as transaction charges, not merchandise products. They may appear in the revenue audit but not in semantic product features.
- Preserve the current executed notebook and results as v1 before the clean v2 run. Do not overwrite them without a successful test run.
- Update the report and presentation only after the v2 metrics are available. Do not preserve the old claim that MiniLM adds almost no gain if the repaired benchmark changes the evidence.
- NEEDS CONFIRMATION: the final report and presentation deadlines remain unconfirmed in `WORK_SPLIT.md`.

## Review findings that drive this plan

These are facts from the current notebook and direct data inspection.

- The notebook is not reproducible in the current workspace. Its kernel metadata names a missing `.venv (3.12.14)`, and `DATA_PATH` expects `Online Retail 2.xlsx`, while the repository contains `data/online+retail.zip`.
- There is no dedicated EDA section. The notebook has audits and model-result plots, but no exploration that explains class drift, missing customers, revenue concentration, country concentration, or feature skew.
- The TF-IDF benchmark is too weak. The current raw sparse TF-IDF KMeans selects K=8 with cosine silhouette 0.015, and its largest cluster contains 2,270 of 3,336 training descriptions.
- An LSA-TF-IDF check with 100 singular-value components increases silhouette scores to approximately 0.076, 0.081, 0.110, and 0.132 for K=8, 12, 16, and 20. This is a more suitable conventional baseline.
- The model's monetary and semantic features include transaction codes such as `POST`, `DOT`, `C2`, `M`, and `BANK CHARGES`. In the deduplicated valid identified data, non-product codes account for 1,542 of 392,692 lines, or 0.39%, but 1.69% of revenue.
- During the July snapshot, these codes appear in the monetary history of 201 of 2,025 customers. Excluding them changes only 0 to 4 target labels in each monthly cohort, so the product definition can be corrected without materially changing prevalence.
- Duplicate removal is defensible. It removes 5,192 valid identified lines and 0.27% of valid revenue, and none of the eight month-end target labels changes.
- The no-leakage design is otherwise sound. Feature and target windows join without overlap, product clusters are fitted on descriptions observed before July, and rolling test models use only earlier snapshots.
- The small final PR-AUC differences have no uncertainty estimates. The current report calls Boosting + TF-IDF the July winner by roughly 0.0004 PR-AUC, which is too small to interpret without a paired interval.

## File structure

- Modify: `Online_Retail_Repeat_Purchase_Project.ipynb`. It remains the complete analysis and report evidence source.
- Create: `requirements.in`. It lists direct dependencies for rebuilding the notebook environment.
- Create: `requirements.lock.txt`. It records the exact resolved package versions after installation.
- Create: `.gitignore`. It excludes the local `.venv`, caches, and generated execution files while leaving selected notebook outputs available for review.
- Create: `tests/test_pipeline_contract.py`. It checks the archive, temporal windows, product taxonomy, cluster outputs, and uncertainty outputs.
- Modify after successful execution: `Detailed_Project_Pipeline_Aligned.docx`. Match the implemented methods and v2 results.
- Modify after successful execution: `Group_Project_Report.docx` and any checkpoint text that cites model results.

## Task 1: Make the notebook reproducible

**Files:**
- Modify: `Online_Retail_Repeat_Purchase_Project.ipynb`, environment cell and data-loading cells
- Create: `requirements.in`
- Create: `requirements.lock.txt`
- Create: `.gitignore`

**Interfaces:**
- Consumes: `data/online+retail.zip`
- Produces: `raw`, a DataFrame with 541,909 rows and columns `Invoice`, `StockCode`, `Description`, `Quantity`, `InvoiceDate`, `Price`, `CustomerID`, `Country`
- Produces: `DATA_ARCHIVE`, `ARCHIVE_MEMBER`, `PROJECT_DIR`, and `OUTPUT_DIR`

- [ ] **Step 1: Write the failing archive contract test**

Add this test to `tests/test_pipeline_contract.py`:

```python
import json
from pathlib import Path
from zipfile import ZipFile

import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
ARCHIVE = ROOT / "data" / "online+retail.zip"


def load_raw():
    with ZipFile(ARCHIVE) as archive, archive.open("Online Retail.xlsx") as source:
        return pd.read_excel(source, sheet_name="Online Retail", engine="openpyxl")


def test_raw_archive_contract():
    raw = load_raw()
    assert len(raw) == 541_909
    assert raw.columns.tolist() == [
        "InvoiceNo",
        "StockCode",
        "Description",
        "Quantity",
        "InvoiceDate",
        "UnitPrice",
        "CustomerID",
        "Country",
    ]
    assert raw["InvoiceDate"].min() == pd.Timestamp("2010-12-01 08:26:00")
    assert raw["InvoiceDate"].max() == pd.Timestamp("2011-12-09 12:50:00")
```

- [ ] **Step 2: Run the test before notebook changes**

Run:

```bash
pytest tests/test_pipeline_contract.py::test_raw_archive_contract -q
```

Expected: PASS. This confirms the archived source is the intended dataset and does not change.

- [ ] **Step 3: Create the environment files**

Create `requirements.in`:

```text
numpy==1.26.4
pandas==2.2.3
scikit-learn==1.8.0
openpyxl==3.1.5
matplotlib==3.10.0
seaborn==0.13.2
sentence-transformers
jupyterlab==4.2.6
ipykernel==6.29.5
pytest==9.1.1
```

Create `.gitignore`:

```text
.venv/
.cache/
__pycache__/
.pytest_cache/
notebook_outputs/execution.log
notebook_outputs/executed_notebook.ipynb
```

- [ ] **Step 4: Build and pin the environment**

Run:

```bash
/Users/devoob/miniconda3/bin/python3.12 -m venv .venv
.venv/bin/python -m pip install --upgrade pip
.venv/bin/pip install -r requirements.in
.venv/bin/pip freeze > requirements.lock.txt
```

Expected: the commands complete without dependency conflicts. The first model run may download `sentence-transformers/all-MiniLM-L6-v2` if it is not cached.

- [ ] **Step 5: Replace the broken configuration**

In the environment cell of `Online_Retail_Repeat_Purchase_Project.ipynb`, replace the current `DATA_PATH` block with:

```python
from pathlib import Path
from zipfile import ZipFile

PROJECT_DIR = Path.cwd()
DATA_ARCHIVE = PROJECT_DIR / "data" / "online+retail.zip"
ARCHIVE_MEMBER = "Online Retail.xlsx"
SHEET_NAME = "Online Retail"
OUTPUT_DIR = PROJECT_DIR / "notebook_outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

assert DATA_ARCHIVE.exists(), f"Missing source archive: {DATA_ARCHIVE}"
assert (PROJECT_DIR / "elec3544_proposal.pdf").exists(), "Run the notebook from the project root."
```

- [ ] **Step 6: Load the workbook directly from the archive**

Replace the current `pd.read_excel(DATA_PATH, ...)` line with:

```python
with ZipFile(DATA_ARCHIVE) as archive:
    with archive.open(ARCHIVE_MEMBER) as source:
        raw = pd.read_excel(source, sheet_name=SHEET_NAME, engine="openpyxl")
```

Keep the existing column rename and 541,909-row assertion.

- [ ] **Step 7: Verify the notebook configuration**

Run:

```bash
.venv/bin/python -m pytest tests/test_pipeline_contract.py::test_raw_archive_contract -q
```

Expected: PASS.

- [ ] **Step 8: Commit the reproducibility change**

Run:

```bash
git add Online_Retail_Repeat_Purchase_Project.ipynb requirements.in requirements.lock.txt .gitignore tests/test_pipeline_contract.py
git commit -m "fix: make notebook data input reproducible"
```

## Task 2: Add an EDA stage that drives project decisions

**Files:**
- Modify: `Online_Retail_Repeat_Purchase_Project.ipynb`, after the raw-data audit and again after snapshot construction
- Create at run time: `notebook_outputs/eda_data_quality.csv`
- Create at run time: `notebook_outputs/eda_monthly_activity.csv`
- Create at run time: `notebook_outputs/eda_customer_concentration.csv`
- Create at run time: `notebook_outputs/eda_country_activity.csv`
- Create at run time: `notebook_outputs/eda_snapshot_balance.csv`
- Create at run time: `notebook_outputs/eda_raw_activity.png`
- Create at run time: `notebook_outputs/eda_customer_concentration.png`
- Create at run time: `notebook_outputs/eda_snapshot_balance.png`

**Interfaces:**
- Consumes: `raw`, `data`, `positive_lines`, `invoice_summary`, and `base_snapshots`
- Produces: the five CSV audit tables and three PNG figures listed above
- Produces no new model features

- [ ] **Step 1: Add the EDA contract tests**

Append these tests to `tests/test_pipeline_contract.py`:

```python
import pytest


OUTPUT_DIR = ROOT / "notebook_outputs"


@pytest.mark.parametrize(
    "filename",
    [
        "eda_data_quality.csv",
        "eda_monthly_activity.csv",
        "eda_customer_concentration.csv",
        "eda_country_activity.csv",
        "eda_snapshot_balance.csv",
        "eda_raw_activity.png",
        "eda_customer_concentration.png",
        "eda_snapshot_balance.png",
    ],
)
def test_required_eda_output_exists(filename):
    assert (OUTPUT_DIR / filename).exists(), filename


def test_eda_monthly_activity_has_all_months():
    monthly = pd.read_csv(OUTPUT_DIR / "eda_monthly_activity.csv")
    assert len(monthly) == 13
    assert monthly["Month"].iloc[0] == "2010-12"
    assert monthly["Month"].iloc[-1] == "2011-12"
    assert monthly["EligibleCustomers"].gt(0).all()
```

- [ ] **Step 2: Run the EDA tests to verify they fail**

Run:

```bash
pytest tests/test_pipeline_contract.py -q -k eda
```

Expected: FAIL because the EDA outputs do not exist.

- [ ] **Step 3: Add the data-quality and monthly activity section**

Insert a markdown cell named `## Stage 1B. Exploratory data analysis` after the cleaning audit. Add a code cell that creates the audit tables:

```python
valid_purchase_lines = data.loc[data["IsValidPurchase"]].copy()

eda_data_quality = pd.DataFrame(
    [
        {"Check": "Raw invoice lines", "Rows": len(raw), "ShareOfRaw": 1.0},
        {"Check": "Exact duplicate rows", "Rows": int(raw.duplicated().sum()),
         "ShareOfRaw": raw.duplicated().mean()},
        {"Check": "Missing CustomerID", "Rows": int(raw["CustomerID"].isna().sum()),
         "ShareOfRaw": raw["CustomerID"].isna().mean()},
        {"Check": "Missing Description", "Rows": int(raw["Description"].isna().sum()),
         "ShareOfRaw": raw["Description"].isna().mean()},
        {"Check": "Cancellation or negative quantity",
         "Rows": int(data["IsCancellation"].sum()), "ShareOfRaw": data["IsCancellation"].mean()},
        {"Check": "Non-positive quantity",
         "Rows": int(data["Quantity"].le(0).sum()), "ShareOfRaw": data["Quantity"].le(0).mean()},
        {"Check": "Non-positive price",
         "Rows": int(data["Price"].le(0).sum()), "ShareOfRaw": data["Price"].le(0).mean()},
        {"Check": "Valid identified purchase lines",
         "Rows": len(valid_purchase_lines),
         "ShareOfRaw": len(valid_purchase_lines) / len(raw)},
    ]
)
display(eda_data_quality.style.format({"ShareOfRaw": "{:.2%}"}))
eda_data_quality.to_csv(OUTPUT_DIR / "eda_data_quality.csv", index=False)

monthly_eda = (
    valid_purchase_lines.assign(Month=valid_purchase_lines["InvoiceDate"].dt.to_period("M").astype(str))
    .groupby("Month", as_index=False)
    .agg(
        ValidLines=("Invoice", "size"),
        Revenue=("LineRevenue", "sum"),
        EligibleCustomers=("CustomerID", "nunique"),
        Invoices=("Invoice", "nunique"),
    )
)
monthly_eda.to_csv(OUTPUT_DIR / "eda_monthly_activity.csv", index=False)
display(monthly_eda.style.format({"Revenue": "{:,.0f}"}))
```

- [ ] **Step 4: Add the time-series figure**

Add a code cell:

```python
fig, axes = plt.subplots(2, 1, figsize=(12, 8), sharex=True)
monthly_plot = monthly_eda.assign(MonthDate=pd.to_datetime(monthly_eda["Month"] + "-01"))

axes[0].plot(monthly_plot["MonthDate"], monthly_plot["Revenue"], marker="o", color="#4C78A8")
axes[0].set_ylabel("Valid revenue")
axes[0].set_title("Monthly valid revenue")

axes[1].plot(monthly_plot["MonthDate"], monthly_plot["EligibleCustomers"],
             marker="o", color="#E15759")
axes[1].set_ylabel("Identified customers")
axes[1].set_xlabel("Month")
axes[1].set_title("Monthly identified purchasing customers")

plt.tight_layout()
plt.savefig(OUTPUT_DIR / "eda_raw_activity.png", dpi=180, bbox_inches="tight")
plt.show()
```

Add Markdown immediately after the figure:

```markdown
The monthly activity series shows whether the annual totals hide a material change in customer activity. Report the observed direction and magnitude, and explain how it affects the rolling test design. Do not describe the series as stable unless the figure supports that statement.
```

- [ ] **Step 5: Add customer concentration and country analysis**

Add a code cell:

```python
customer_spend = (
    valid_purchase_lines.groupby("CustomerID", as_index=False)["LineRevenue"]
    .sum()
    .sort_values("LineRevenue", ascending=False)
    .reset_index(drop=True)
)
customer_spend["CumulativeRevenueShare"] = (
    customer_spend["LineRevenue"].cumsum() / customer_spend["LineRevenue"].sum()
)
customer_spend["CustomerShare"] = (customer_spend.index + 1) / len(customer_spend)

concentration_rows = []
for share in [0.01, 0.05, 0.10, 0.20]:
    n_customers = max(1, int(np.ceil(len(customer_spend) * share)))
    concentration_rows.append(
        {
            "TopCustomerShare": share,
            "CustomerCount": n_customers,
            "RevenueShare": customer_spend.iloc[:n_customers]["LineRevenue"].sum()
            / customer_spend["LineRevenue"].sum(),
        }
    )
eda_customer_concentration = pd.DataFrame(concentration_rows)
display(eda_customer_concentration.style.format(
    {"TopCustomerShare": "{:.0%}", "RevenueShare": "{:.1%}"}
))
eda_customer_concentration.to_csv(
    OUTPUT_DIR / "eda_customer_concentration.csv", index=False
)

country_eda = (
    valid_purchase_lines.groupby("Country", as_index=False)
    .agg(ValidLines=("Invoice", "size"), Customers=("CustomerID", "nunique"),
         Revenue=("LineRevenue", "sum"))
    .sort_values("Revenue", ascending=False)
)
country_eda["RevenueShare"] = country_eda["Revenue"] / country_eda["Revenue"].sum()
display(country_eda.head(15).style.format({"Revenue": "{:,.0f}", "RevenueShare": "{:.1%}"}))
country_eda.to_csv(OUTPUT_DIR / "eda_country_activity.csv", index=False)
```

Add a second code cell for the Lorenz curve:

```python
plt.figure(figsize=(7, 6))
plt.plot([0, 1], [0, 1], linestyle="--", color="#999999", label="Equal spending")
plt.plot(customer_spend["CustomerShare"], customer_spend["CumulativeRevenueShare"],
         color="#4C78A8", label="Observed spending")
plt.xlabel("Cumulative share of customers")
plt.ylabel("Cumulative share of revenue")
plt.title("Customer revenue concentration")
plt.legend()
plt.tight_layout()
plt.savefig(OUTPUT_DIR / "eda_customer_concentration.png", dpi=180, bbox_inches="tight")
plt.show()
```

- [ ] **Step 6: Add snapshot class-balance EDA**

Immediately after `snapshot_audit` is created, add:

```python
eda_snapshot_balance = snapshot_audit.rename(
    columns={
        "CustomerSnapshots": "EligibleCustomers",
        "PositiveTargets": "RepeatCustomers",
        "PositiveRate": "RepeatPurchaseRate",
    }
).copy()
eda_snapshot_balance.to_csv(OUTPUT_DIR / "eda_snapshot_balance.csv", index=False)

plt.figure(figsize=(9, 5))
sns.lineplot(
    data=eda_snapshot_balance,
    x="SnapshotDate",
    y="RepeatPurchaseRate",
    marker="o",
    color="#E15759",
)
plt.axhline(
    base_snapshots["RepeatPurchase30D"].mean(),
    linestyle="--",
    color="#666666",
    label="Overall prevalence",
)
plt.ylabel("Repeat-purchase rate")
plt.xlabel("Snapshot date")
plt.title("Target prevalence by snapshot")
plt.legend()
plt.tight_layout()
plt.savefig(OUTPUT_DIR / "eda_snapshot_balance.png", dpi=180, bbox_inches="tight")
plt.show()
```

- [ ] **Step 7: Add EDA conclusions in prose**

Add a markdown cell containing exact observed values after execution. It must answer:

1. How much of the raw file cannot support customer-level modelling?
2. Which months have the largest revenue and customer counts?
3. How concentrated is revenue among the largest customers?
4. How dominant is the United Kingdom, and does `IsUK` have useful variation?
5. How much does target prevalence change across snapshots?
6. Which observed skew motivates the log transforms and PR-AUC metric?

Do not add generic statements that are not tied to the generated tables or figures.

- [ ] **Step 8: Run the EDA tests**

Run the notebook through the snapshot stage, then run:

```bash
pytest tests/test_pipeline_contract.py -q -k eda
```

Expected: PASS.

- [ ] **Step 9: Commit the EDA stage**

Run:

```bash
git add Online_Retail_Repeat_Purchase_Project.ipynb tests/test_pipeline_contract.py
git commit -m "feat: add methodology-driving EDA"
```

## Task 3: Separate merchandise purchases from transaction charges

**Files:**
- Modify: `Online_Retail_Repeat_Purchase_Project.ipynb`, cleaning and transaction-classification section
- Create at run time: `notebook_outputs/non_product_code_audit.csv`
- Modify: `tests/test_pipeline_contract.py`

**Interfaces:**
- Consumes: parsed `data`
- Produces: `NON_PRODUCT_STOCK_CODES`, `IsNonProductLine`, `valid_purchase_lines`, `charge_lines`, and `positive_lines`
- `positive_lines` remains the merchandise-only input used by invoice, snapshot, text, and model stages

- [ ] **Step 1: Add the failing product taxonomy test**

Append:

```python
NON_PRODUCT_CODES = {
    "POST",
    "DOT",
    "C2",
    "M",
    "m",
    "D",
    "S",
    "BANK CHARGES",
    "AMAZONFEE",
    "CRUK",
    "B",
}


def test_non_product_codes_are_reported():
    audit = pd.read_csv(OUTPUT_DIR / "non_product_code_audit.csv")
    assert set(audit.loc[audit["IsNonProduct"], "StockCode"]).issubset(
        NON_PRODUCT_CODES | {"gift_0001_10", "gift_0001_20", "gift_0001_30",
                              "gift_0001_40", "gift_0001_50"}
    )
    assert audit.loc[audit["IsNonProduct"], "IncludedInProductFeatures"].eq(False).all()
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```bash
pytest tests/test_pipeline_contract.py::test_non_product_codes_are_reported -q
```

Expected: FAIL because `non_product_code_audit.csv` does not exist.

- [ ] **Step 3: Define the transaction taxonomy**

After type parsing and before `data["IsValidPurchase"]`, add:

```python
NON_PRODUCT_STOCK_CODES = {
    "POST",
    "DOT",
    "C2",
    "M",
    "m",
    "D",
    "S",
    "BANK CHARGES",
    "AMAZONFEE",
    "CRUK",
    "B",
}

def is_non_product_stock(code):
    code_text = "" if pd.isna(code) else str(code).strip()
    return code_text in NON_PRODUCT_STOCK_CODES or code_text.lower().startswith("gift_0001_")

data["IsNonProductLine"] = data["StockCode"].map(is_non_product_stock)
```

- [ ] **Step 4: Split valid purchases into transaction charges and products**

Replace the current `positive_lines = data.loc[data["IsValidPurchase"]].copy()` block with:

```python
valid_purchase_lines = data.loc[data["IsValidPurchase"]].copy()
charge_lines = valid_purchase_lines.loc[valid_purchase_lines["IsNonProductLine"]].copy()
positive_lines = valid_purchase_lines.loc[
    ~valid_purchase_lines["IsNonProductLine"]
    & valid_purchase_lines["CleanDescription"].notna()
].copy()
return_lines = data.loc[data["IsValidPurchase"].eq(False)
                        & data["HasCustomerID"]
                        & data["IsCancellation"]
                        & ~data["IsNonProductLine"]].copy()
```

Keep the existing return-value calculations after `return_lines` is defined.

- [ ] **Step 5: Export the non-product audit**

Add after the split:

```python
non_product_code_audit = (
    data.assign(ExplodedRevenue=data["Quantity"] * data["Price"])
    .groupby(["StockCode", "IsNonProductLine"], dropna=False, as_index=False)
    .agg(
        Rows=("Invoice", "size"),
        Invoices=("Invoice", "nunique"),
        Customers=("CustomerID", "nunique"),
        Revenue=("ExplodedRevenue", "sum"),
        ExampleDescription=("CleanDescription", lambda s: s.dropna().iloc[0] if s.notna().any() else ""),
    )
    .rename(columns={"IsNonProductLine": "IsNonProduct"})
)
non_product_code_audit["IncludedInProductFeatures"] = (
    ~non_product_code_audit["IsNonProduct"]
)
display(non_product_code_audit.query("IsNonProduct").sort_values("Rows", ascending=False))
non_product_code_audit.to_csv(
    OUTPUT_DIR / "non_product_code_audit.csv", index=False
)
```

- [ ] **Step 6: Update the revenue audit**

Replace the invoice audit so it reports three totals:

```python
invoice_audit = pd.DataFrame({
    "Scope": ["Product lines", "Charge lines", "All valid identified lines"],
    "Lines": [len(positive_lines), len(charge_lines), len(valid_purchase_lines)],
    "Revenue": [
        positive_lines["LineRevenue"].sum(),
        charge_lines["LineRevenue"].sum(),
        valid_purchase_lines["LineRevenue"].sum(),
    ],
})
assert np.isclose(
    valid_purchase_lines["LineRevenue"].sum(),
    positive_lines["LineRevenue"].sum() + charge_lines["LineRevenue"].sum(),
)
display(invoice_audit.style.format({"Revenue": "{:,.2f}"}))
```

Keep the invoice reconciliation assertion against `positive_lines`.

- [ ] **Step 7: Update the definitions in Markdown**

State explicitly that:

- `positive_lines` contains merchandise products only.
- `charge_lines` contains postage, carriage, adjustments, fees, and vouchers.
- Repeat-purchase labels and semantic features use merchandise purchases.
- The report discloses the excluded charge revenue and the label changes from the sensitivity check.

- [ ] **Step 8: Run the taxonomy and EDA tests**

Run the notebook through snapshot construction, then run:

```bash
pytest tests/test_pipeline_contract.py -q -k "non_product or eda"
```

Expected: PASS.

- [ ] **Step 9: Commit the taxonomy correction**

Run:

```bash
git add Online_Retail_Repeat_Purchase_Project.ipynb tests/test_pipeline_contract.py
git commit -m "fix: separate product purchases from transaction charges"
```

## Task 4: Replace the weak TF-IDF baseline and test matched cluster counts

**Files:**
- Modify: `Online_Retail_Repeat_Purchase_Project.ipynb`, product-text and semantic-feature stages
- Create at run time: `notebook_outputs/cluster_quality.csv`
- Create at run time: `notebook_outputs/cluster_matched_k_quality.csv`
- Modify: `tests/test_pipeline_contract.py`

**Interfaces:**
- Consumes: `description_table`, `train_mask`, `all_descriptions`, and `minilm_embeddings`
- Produces: `tfidf_svd`, `X_tfidf_train`, `X_tfidf_all`, `tfidf_k_results`, `minilm_k_results`, `tfidf_clusterer`, and `minilm_clusterer`
- Produces matched models `tfidf_k16_clusterer` and `minilm_k16_clusterer`
- Produces cluster prefixes `TFIDFBest`, `MiniLMBest`, `TFIDFK16`, and `MiniLMK16`

- [ ] **Step 1: Add failing cluster-quality tests**

Append:

```python
def test_cluster_quality_has_reduced_tfidf_benchmark():
    quality = pd.read_csv(OUTPUT_DIR / "cluster_quality.csv")
    assert {"TF-IDF + LSA", "MiniLM"} <= set(quality["Representation"])
    assert quality["SilhouetteCosine"].notna().all()
    assert quality["LargestClusterShare"].between(0, 1).all()


def test_matched_cluster_count_is_reported_for_both_representations():
    matched = pd.read_csv(OUTPUT_DIR / "cluster_matched_k_quality.csv")
    assert set(matched["Representation"]) == {"TF-IDF + LSA", "MiniLM"}
    assert matched["K"].nunique() == 1
```

- [ ] **Step 2: Run the tests to verify they fail**

Run:

```bash
pytest tests/test_pipeline_contract.py -q -k cluster
```

Expected: FAIL because the cluster-quality files do not exist.

- [ ] **Step 3: Extend cluster diagnostics**

Modify `choose_cluster_count` to add largest-cluster share:

```python
cluster_counts = np.bincount(labels, minlength=k)
rows.append(
    {
        "Representation": representation_name,
        "K": k,
        "SilhouetteCosine": score,
        "SmallestCluster": int(cluster_counts.min()),
        "LargestCluster": int(cluster_counts.max()),
        "LargestClusterShare": float(cluster_counts.max() / len(labels)),
    }
)
```

- [ ] **Step 4: Add a fixed K=16 helper**

Add after `choose_cluster_count`:

```python
def fit_fixed_k(X_train, k, random_state=RANDOM_STATE):
    model = MiniBatchKMeans(
        n_clusters=k,
        random_state=random_state,
        n_init=20,
        batch_size=512,
        reassignment_ratio=0.01,
    )
    model.fit(X_train)
    return model
```

- [ ] **Step 5: Replace raw sparse TF-IDF KMeans with LSA**

Add `TruncatedSVD` to the scikit-learn imports and `normalize` to the preprocessing imports. Replace the old TF-IDF clustering block with:

```python
TFIDF_COMPONENTS = 100

tfidf_vectorizer = TfidfVectorizer(
    ngram_range=(1, 2),
    sublinear_tf=True,
    max_features=8000,
)
X_tfidf_train_sparse = tfidf_vectorizer.fit_transform(
    description_table.loc[train_mask, "CleanDescription"]
)
X_tfidf_all_sparse = tfidf_vectorizer.transform(description_table["CleanDescription"])

tfidf_svd = TruncatedSVD(
    n_components=min(TFIDF_COMPONENTS, X_tfidf_train_sparse.shape[0] - 1),
    random_state=RANDOM_STATE,
)
X_tfidf_train = normalize(tfidf_svd.fit_transform(X_tfidf_train_sparse))
X_tfidf_all = normalize(tfidf_svd.transform(X_tfidf_all_sparse))

tfidf_k_results, tfidf_clusterer = choose_cluster_count(
    X_tfidf_train,
    K_CANDIDATES,
    "TF-IDF + LSA",
)
description_table["TFIDFBestCluster"] = tfidf_clusterer.predict(X_tfidf_all)
display(tfidf_k_results)
```

- [ ] **Step 6: Rename the selected MiniLM cluster output**

After MiniLM clustering, use:

```python
description_table["MiniLMBestCluster"] = minilm_clusterer.predict(minilm_embeddings)
```

Do not fit MiniLM again. The fixed pre-trained encoder remains frozen.

Update the cluster-example cell so it uses the new labels:

```python
tfidf_cluster_examples = representative_descriptions(
    X_tfidf_all,
    tfidf_clusterer,
    description_table["TFIDFBestCluster"],
    "TF-IDF + LSA",
)
minilm_cluster_examples = representative_descriptions(
    minilm_embeddings,
    minilm_clusterer,
    description_table["MiniLMBestCluster"],
    "MiniLM",
)
```

- [ ] **Step 7: Fit the matched K=16 sensitivity models**

Add:

```python
K_MATCHED = 16
tfidf_k16_clusterer = fit_fixed_k(X_tfidf_train, K_MATCHED)
minilm_k16_clusterer = fit_fixed_k(minilm_embeddings[train_mask], K_MATCHED)

description_table["TFIDFK16Cluster"] = tfidf_k16_clusterer.predict(X_tfidf_all)
description_table["MiniLMK16Cluster"] = minilm_k16_clusterer.predict(minilm_embeddings)
```

- [ ] **Step 8: Combine and save cluster-quality evidence**

Add:

```python
cluster_quality = pd.concat(
    [tfidf_k_results.assign(Selection="BestWithinRepresentation"),
     minilm_k_results.assign(Selection="BestWithinRepresentation")],
    ignore_index=True,
)
cluster_quality.to_csv(OUTPUT_DIR / "cluster_quality.csv", index=False)

matched_quality_rows = []
rng = np.random.default_rng(RANDOM_STATE)
for name, matrix, model in [
    ("TF-IDF + LSA", X_tfidf_train, tfidf_k16_clusterer),
    ("MiniLM", minilm_embeddings[train_mask], minilm_k16_clusterer),
]:
    labels = model.predict(matrix)
    sample_size = min(2000, len(labels))
    sample_idx = rng.choice(len(labels), size=sample_size, replace=False)
    counts = np.bincount(labels, minlength=K_MATCHED)
    matched_quality_rows.append(
        {
            "Representation": name,
            "K": K_MATCHED,
            "SilhouetteCosine": silhouette_score(
                matrix[sample_idx], labels[sample_idx], metric="cosine"
            ),
            "SmallestCluster": int(counts.min()),
            "LargestCluster": int(counts.max()),
            "LargestClusterShare": float(counts.max() / len(labels)),
        }
    )
cluster_matched_k_quality = pd.DataFrame(matched_quality_rows)
display(cluster_matched_k_quality)
cluster_matched_k_quality.to_csv(
    OUTPUT_DIR / "cluster_matched_k_quality.csv", index=False
)
```

- [ ] **Step 9: Build the four semantic feature sets**

Update the description-to-cluster maps and feature-frame construction:

```python
cluster_maps = {
    "TFIDFBest": description_table.set_index("CleanDescription")["TFIDFBestCluster"].to_dict(),
    "MiniLMBest": description_table.set_index("CleanDescription")["MiniLMBestCluster"].to_dict(),
    "TFIDFK16": description_table.set_index("CleanDescription")["TFIDFK16Cluster"].to_dict(),
    "MiniLMK16": description_table.set_index("CleanDescription")["MiniLMK16Cluster"].to_dict(),
}

feature_frames = []
for prefix, mapping in cluster_maps.items():
    cluster_column = f"{prefix}ClusterCode"
    positive_lines[cluster_column] = positive_lines["CleanDescription"].map(mapping)
    assert positive_lines[cluster_column].notna().all()
    n_clusters = int(positive_lines[cluster_column].max()) + 1
    feature_frames.append(
        pd.concat(
            [
                build_cluster_features(
                    snapshot_date,
                    cluster_column,
                    prefix,
                    n_clusters,
                )
                for snapshot_date in SNAPSHOT_DATES
            ],
            ignore_index=True,
        )
    )

semantic_features = feature_frames[0]
for frame in feature_frames[1:]:
    semantic_features = semantic_features.merge(
        frame, on=["SnapshotDate", "CustomerID"], how="outer"
    )
```

Replace the two existing semantic merges with one merge against `semantic_features`. Keep the zero-fill assertion after checking that each eligible row has a merchandise history.

- [ ] **Step 10: Add matched-K model variants**

Add feature lists:

```python
TFIDF_BEST_FEATURES = CANCELLATION_FEATURES + [
    c for c in model_data if c.startswith("TFIDFBest")
]
TFIDF_K16_FEATURES = CANCELLATION_FEATURES + [
    c for c in model_data if c.startswith("TFIDFK16")
]
MINILM_BEST_FEATURES = CANCELLATION_FEATURES + [
    c for c in model_data if c.startswith("MiniLMBest")
]
MINILM_K16_FEATURES = CANCELLATION_FEATURES + [
    c for c in model_data if c.startswith("MiniLMK16")
]
```

Replace the old TF-IDF and MiniLM feature lists and model specifications with the following exact names:

```python
model_specs = {
    "Logistic RFM": (logistic_model(), RFM_FEATURES),
    "Logistic Behaviour": (logistic_model(), CANCELLATION_FEATURES),
    "Logistic + MiniLM Best": (logistic_model(), MINILM_BEST_FEATURES),
    "Boosting RFM": (boosting_model(BEST_BOOSTING_PARAMS), RFM_FEATURES),
    "Boosting Behaviour": (
        boosting_model(BEST_BOOSTING_PARAMS),
        CANCELLATION_FEATURES,
    ),
    "Boosting + TFIDF Best": (
        boosting_model(BEST_BOOSTING_PARAMS),
        TFIDF_BEST_FEATURES,
    ),
    "Boosting + MiniLM Best": (
        boosting_model(BEST_BOOSTING_PARAMS),
        MINILM_BEST_FEATURES,
    ),
    "Boosting + TFIDF K16": (
        boosting_model(BEST_BOOSTING_PARAMS),
        TFIDF_K16_FEATURES,
    ),
    "Boosting + MiniLM K16": (
        boosting_model(BEST_BOOSTING_PARAMS),
        MINILM_K16_FEATURES,
    ),
}
```

Keep the best-K models as the main comparison. Use matched K=16 as a fairness sensitivity check.

Update the ablation lookup to use the exact model names:

```python
ablation_pairs = [
    ("Logistic regression", "Logistic RFM", "Logistic Behaviour",
     "Logistic + MiniLM Best"),
    ("Gradient boosting", "Boosting RFM", "Boosting Behaviour",
     "Boosting + MiniLM Best"),
]

tfidf_gain = (
    result_lookup.loc["Boosting + TFIDF Best", "PR_AUC"]
    - result_lookup.loc["Boosting Behaviour", "PR_AUC"]
)
```

Replace the old tuple list and `tfidf_gain` expression in the conclusion cell with these exact names.

- [ ] **Step 11: Interpret cluster quality before predictive results**

Add Markdown that reports:

- Silhouette score for every representation and candidate K.
- Largest-cluster share.
- Representative descriptions reviewed by the text owner.
- Whether the TF-IDF baseline is separated enough to be considered a credible benchmark.
- Whether matched-K results agree with best-K results.

If clusters remain weak, say so. Do not claim that MiniLM adds no value when the conventional baseline is itself poorly clustered.

- [ ] **Step 12: Run the cluster tests**

Run:

```bash
pytest tests/test_pipeline_contract.py -q -k cluster
```

Expected: PASS.

- [ ] **Step 13: Commit the text-benchmark repair**

Run:

```bash
git add Online_Retail_Repeat_Purchase_Project.ipynb tests/test_pipeline_contract.py
git commit -m "fix: use a fair TF-IDF benchmark and matched cluster comparison"
```

## Task 5: Add uncertainty and stop overinterpreting small differences

**Files:**
- Modify: `Online_Retail_Repeat_Purchase_Project.ipynb`, evaluation and conclusions stages
- Create at run time: `notebook_outputs/metric_uncertainty.csv`
- Create at run time: `notebook_outputs/paired_model_differences.csv`
- Modify: `tests/test_pipeline_contract.py`

**Interfaces:**
- Consumes: `predictions`, `rolling_results`, and `pooled_results`
- Produces: `metric_uncertainty` with `PR_AUC`, `PR_AUC_Low`, and `PR_AUC_High`
- Produces: `paired_model_differences` with paired gain intervals
- Produces: `monthly_macro_pr_auc`

- [ ] **Step 1: Add failing uncertainty tests**

Append:

```python
def test_pr_auc_uncertainty_is_reported():
    metrics = pd.read_csv(OUTPUT_DIR / "metric_uncertainty.csv")
    required = {"Model", "PR_AUC", "PR_AUC_Low", "PR_AUC_High"}
    assert required <= set(metrics.columns)
    assert metrics["PR_AUC_Low"].le(metrics["PR_AUC"]).all()
    assert metrics["PR_AUC"].le(metrics["PR_AUC_High"]).all()


def test_paired_model_differences_have_intervals():
    differences = pd.read_csv(OUTPUT_DIR / "paired_model_differences.csv")
    required = {"ModelA", "ModelB", "GainAminusB", "GainLow", "GainHigh"}
    assert required <= set(differences.columns)
    assert differences["GainLow"].le(differences["GainHigh"]).all()
```

- [ ] **Step 2: Run the tests to verify they fail**

Run:

```bash
pytest tests/test_pipeline_contract.py -q -k "uncertainty or differences"
```

Expected: FAIL because the uncertainty outputs do not exist.

- [ ] **Step 3: Add the grouped bootstrap function**

After the metric helper functions, add:

```python
def bootstrap_ap_interval(y_true, scores, groups, n_boot=1000, random_state=RANDOM_STATE):
    y_true = np.asarray(y_true)
    scores = np.asarray(scores)
    groups = np.asarray(groups)
    unique_groups = np.unique(groups)
    index_by_group = {
        group: np.flatnonzero(groups == group)
        for group in unique_groups
    }
    rng = np.random.default_rng(random_state)
    estimates = []
    for _ in range(n_boot):
        sampled_groups = rng.choice(unique_groups, size=len(unique_groups), replace=True)
        row_idx = np.concatenate([index_by_group[group] for group in sampled_groups])
        if len(np.unique(y_true[row_idx])) < 2:
            continue
        estimates.append(average_precision_score(y_true[row_idx], scores[row_idx]))
    low, high = np.percentile(estimates, [2.5, 97.5])
    return float(low), float(high)
```

- [ ] **Step 4: Add pooled metric intervals**

Add after `pooled_results` is created:

```python
uncertainty_rows = []
for model_name, group in predictions.groupby("Model"):
    low, high = bootstrap_ap_interval(
        group["Actual"],
        group["Score"],
        group["CustomerID"],
    )
    uncertainty_rows.append(
        {
            "Model": model_name,
            "PR_AUC": average_precision_score(group["Actual"], group["Score"]),
            "PR_AUC_Low": low,
            "PR_AUC_High": high,
            "BootstrapReplicates": 1000,
            "ResamplingUnit": "CustomerID within pooled test",
        }
    )
metric_uncertainty = pd.DataFrame(uncertainty_rows).sort_values(
    "PR_AUC", ascending=False
)
display(metric_uncertainty.style.format(
    {"PR_AUC": "{:.3f}", "PR_AUC_Low": "{:.3f}", "PR_AUC_High": "{:.3f}"}
))
metric_uncertainty.to_csv(OUTPUT_DIR / "metric_uncertainty.csv", index=False)
```

- [ ] **Step 5: Add paired model differences**

Add:

```python
score_wide = predictions.pivot_table(
    index=["SnapshotDate", "CustomerID"],
    columns="Model",
    values="Score",
)
actual = (
    predictions.drop_duplicates(["SnapshotDate", "CustomerID"])
    .set_index(["SnapshotDate", "CustomerID"])["Actual"]
    .reindex(score_wide.index)
)
customer_groups = score_wide.index.get_level_values("CustomerID").to_numpy()

def paired_ap_interval(actual_values, scores_a, scores_b, groups,
                       n_boot=1000, random_state=RANDOM_STATE):
    actual_values = np.asarray(actual_values)
    groups = np.asarray(groups)
    unique_groups = np.unique(groups)
    index_by_group = {group: np.flatnonzero(groups == group) for group in unique_groups}
    rng = np.random.default_rng(random_state)
    estimates = []
    for _ in range(n_boot):
        sampled_groups = rng.choice(unique_groups, size=len(unique_groups), replace=True)
        row_idx = np.concatenate([index_by_group[group] for group in sampled_groups])
        if len(np.unique(actual_values[row_idx])) < 2:
            continue
        estimates.append(
            average_precision_score(actual_values[row_idx], scores_a[row_idx])
            - average_precision_score(actual_values[row_idx], scores_b[row_idx])
        )
    return float(np.mean(estimates)), *map(float, np.percentile(estimates, [2.5, 97.5]))

pair_rows = []
for model_a, model_b in [
    ("Boosting Behaviour", "Boosting RFM"),
    ("Boosting + TFIDF Best", "Boosting Behaviour"),
    ("Boosting + MiniLM Best", "Boosting Behaviour"),
    ("Boosting + TFIDF K16", "Boosting Behaviour"),
    ("Boosting + MiniLM K16", "Boosting Behaviour"),
]:
    gain, low, high = paired_ap_interval(
        actual,
        score_wide[model_a],
        score_wide[model_b],
        customer_groups,
    )
    pair_rows.append(
        {
            "ModelA": model_a,
            "ModelB": model_b,
            "GainAminusB": gain,
            "GainLow": low,
            "GainHigh": high,
            "ExcludesZero": bool(low > 0 or high < 0),
        }
    )
paired_model_differences = pd.DataFrame(pair_rows)
display(paired_model_differences.style.format(
    {"GainAminusB": "{:+.3f}", "GainLow": "{:+.3f}", "GainHigh": "{:+.3f}"}
))
paired_model_differences.to_csv(
    OUTPUT_DIR / "paired_model_differences.csv", index=False
)
```

- [ ] **Step 6: Report monthly and macro-average performance**

Add:

```python
monthly_macro_pr_auc = (
    rolling_results.groupby("Model", as_index=False)["PR_AUC"]
    .agg(["mean", "std", "min", "max"])
    .reset_index()
    .sort_values("mean", ascending=False)
)
display(monthly_macro_pr_auc.style.format(
    {"mean": "{:.3f}", "std": "{:.3f}", "min": "{:.3f}", "max": "{:.3f}"}
))
```

The report must show both the monthly results and the macro average. Do not rely on a pooled AP alone when class prevalence changes sharply by month.

- [ ] **Step 7: Rewrite the conclusion rules**

Replace the current absolute `0.005` decision rule with:

1. A model has a defensible gain only when the paired 95% interval for its PR-AUC difference is entirely above zero.
2. A model with overlapping intervals is described as indistinguishable from its comparator.
3. The simpler model is preferred when predictive performance is indistinguishable.
4. The July choice is reported as validation evidence. The August to October performance remains final evidence and never changes the July selection.
5. The report keeps unsuccessful iterations and weak cluster results.

- [ ] **Step 8: Run the uncertainty tests**

Run the complete model and error-analysis section, then run:

```bash
pytest tests/test_pipeline_contract.py -q -k "uncertainty or differences"
```

Expected: PASS.

- [ ] **Step 9: Commit the evaluation repair**

Run:

```bash
git add Online_Retail_Repeat_Purchase_Project.ipynb tests/test_pipeline_contract.py
git commit -m "feat: add paired uncertainty for model comparisons"
```

## Task 6: Enforce temporal and modelling contracts

**Files:**
- Modify: `tests/test_pipeline_contract.py`

**Interfaces:**
- Consumes: `notebook_outputs/customer_snapshot_features.csv`
- Consumes: `notebook_outputs/non_product_code_audit.csv`
- Consumes: `notebook_outputs/cluster_quality.csv`
- Produces: a single contract suite that fails when leakage or malformed snapshots reappear

- [ ] **Step 1: Add the temporal contract tests**

Append:

```python
def test_snapshot_windows_are_exactly_90_and_30_days():
    snapshots = pd.read_csv(OUTPUT_DIR / "customer_snapshot_features.csv")
    for column in [
        "SnapshotDate",
        "FeatureStart",
        "FeatureEndExclusive",
        "TargetStart",
        "TargetEndExclusive",
    ]:
        snapshots[column] = pd.to_datetime(snapshots[column])
    assert (snapshots["FeatureEndExclusive"] - snapshots["FeatureStart"]).dt.days.eq(90).all()
    assert (snapshots["TargetEndExclusive"] - snapshots["TargetStart"]).dt.days.eq(30).all()
    assert snapshots["FeatureEndExclusive"].eq(snapshots["TargetStart"]).all()
    assert not snapshots.duplicated(["SnapshotDate", "CustomerID"]).any()


def test_snapshot_dates_are_the_eight_pre_registered_months():
    snapshots = pd.read_csv(OUTPUT_DIR / "customer_snapshot_features.csv")
    observed = sorted(pd.to_datetime(snapshots["SnapshotDate"].unique()))
    expected = list(pd.to_datetime([
        "2011-03-31",
        "2011-04-30",
        "2011-05-31",
        "2011-06-30",
        "2011-07-31",
        "2011-08-31",
        "2011-09-30",
        "2011-10-31",
    ]))
    assert observed == expected
```

- [ ] **Step 2: Add semantic-share checks**

Append:

```python
def test_semantic_cluster_shares_cover_the_full_product_history():
    snapshots = pd.read_csv(OUTPUT_DIR / "customer_snapshot_features.csv")
    for prefix in ["TFIDFBest", "MiniLMBest", "TFIDFK16", "MiniLMK16"]:
        columns = [
            column for column in snapshots.columns
            if column.startswith(prefix) and column.endswith("Share")
        ]
        assert columns
        totals = snapshots[columns].sum(axis=1)
        assert np.allclose(totals, 1.0, atol=1e-8)
```

- [ ] **Step 3: Add predictor exclusion checks**

Append:

```python
def test_identifiers_are_not_predictors():
    notebook = json.loads(
        (ROOT / "Online_Retail_Repeat_Purchase_Project.ipynb").read_text()
    )
    source = "\n".join(
        "".join(cell.get("source", []))
        for cell in notebook["cells"]
        if cell["cell_type"] == "code"
    )
    feature_block = source.split("RFM_FEATURES =", 1)[1].split("feature_sets =", 1)[0]
    assert "CustomerID" not in feature_block
    assert "SnapshotDate" not in feature_block
```

- [ ] **Step 4: Run the full contract suite**

Run:

```bash
.venv/bin/python -m pytest tests/test_pipeline_contract.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit the contract suite**

Run:

```bash
git add tests/test_pipeline_contract.py
git commit -m "test: enforce temporal and modelling contracts"
```

## Task 7: Re-execute, compare v1 and v2, and update the report

**Files:**
- Modify: `Online_Retail_Repeat_Purchase_Project.ipynb` outputs after clean execution
- Create at run time: `notebook_outputs/executed_notebook.ipynb`
- Create at run time: `notebook_outputs/execution.log`
- Create: `notebook_outputs/v1_v2_comparison.csv`
- Modify: `Detailed_Project_Pipeline_Aligned.docx`
- Modify: `Group_Project_Report.docx`
- Modify: `WORK_SPLIT.md` only if task ownership or deadlines changed

**Interfaces:**
- Consumes: the notebook and contract test suite
- Produces: a clean executed notebook, final outputs, a v1/v2 decision log, and synchronized report text

- [ ] **Step 1: Record v1 metrics before rerunning**

Extract these values from the current executed notebook into `notebook_outputs/v1_v2_comparison.csv`:

```csv
Version,Stage,Metric,Value
v1,TF-IDF benchmark,Best K,8
v1,TF-IDF benchmark,Silhouette cosine,0.014682
v1,TF-IDF benchmark,Largest cluster share,0.680
v1,July validation,Best model,Boosting + TFIDF
v1,Final pooled,Best model,Boosting + MiniLM
v1,Final pooled,Boosting Behaviour PR-AUC,0.638
v1,Final pooled,Boosting + MiniLM PR-AUC,0.638
v1,Final pooled,Logistic Behaviour PR-AUC,0.634
```

Add v2 rows after execution. Do not delete v1 rows.

- [ ] **Step 2: Run one clean notebook execution**

Run:

```bash
mkdir -p notebook_outputs .cache/matplotlib
MPLCONFIGDIR=.cache/matplotlib .venv/bin/jupyter nbconvert \
  --to notebook \
  --execute Online_Retail_Repeat_Purchase_Project.ipynb \
  --output executed_notebook.ipynb \
  --output-dir notebook_outputs \
  --ExecutePreprocessor.timeout=1800 \
  > notebook_outputs/execution.log 2>&1
```

Expected: exit code 0 and `Notebook completed successfully.` in the output.

- [ ] **Step 3: Run the complete contract suite**

Run:

```bash
.venv/bin/python -m pytest tests/test_pipeline_contract.py -q
```

Expected: all tests PASS.

- [ ] **Step 4: Inspect the evidence**

Check these files directly:

```text
notebook_outputs/eda_data_quality.csv
notebook_outputs/eda_monthly_activity.csv
notebook_outputs/eda_customer_concentration.csv
notebook_outputs/non_product_code_audit.csv
notebook_outputs/cluster_quality.csv
notebook_outputs/cluster_matched_k_quality.csv
notebook_outputs/metric_uncertainty.csv
notebook_outputs/paired_model_differences.csv
notebook_outputs/rolling_monthly_results.csv
notebook_outputs/final_pooled_results.csv
notebook_outputs/model_comparison.png
notebook_outputs/error_analysis_summary.csv
```

Verify that every number cited in the notebook's conclusion appears in one of these outputs.

- [ ] **Step 5: Replace the notebook output only after success**

Run:

```bash
cp notebook_outputs/executed_notebook.ipynb Online_Retail_Repeat_Purchase_Project.ipynb
```

- [ ] **Step 6: Update the pipeline document**

Revise `Detailed_Project_Pipeline_Aligned.docx` so Stage 1B, the product taxonomy, the LSA-TF-IDF benchmark, matched K=16 sensitivity, and uncertainty intervals match the notebook.

Replace all copied result tables with v2 values. Keep v1 values only in an explicitly labeled iteration section.

- [ ] **Step 7: Update the group report**

Revise `Group_Project_Report.docx` so the report:

- Explains why online retail data fits the retention question.
- Reports missing identifiers, duplicate handling, charge exclusion, and the resulting modelling scope.
- Uses EDA evidence to justify rolling evaluation, log transforms, class weighting, and PR-AUC.
- States whether the LSA-TF-IDF and MiniLM benchmarks are credible.
- Reports paired intervals for all central comparisons.
- Describes indistinguishable models as indistinguishable.
- Preserves unsuccessful results and states what changed after review.

- [ ] **Step 8: Update the checkpoint evidence**

If a checkpoint PDF cites model results, regenerate it from v2 outputs. Each owner should reference their own artifact:

- Data owner: EDA tables, duplicate sensitivity, and charge audit.
- Text owner: LSA-TF-IDF comparison, matched-K comparison, cluster review, and error analysis.
- Pipeline owner: temporal contracts, uncertainty intervals, and baseline comparison.

- [ ] **Step 9: Review the final diff**

Run:

```bash
git status --short
git diff --stat
git diff -- tests/test_pipeline_contract.py
```

Expected: only the notebook, methodology documents, tests, and intended support files change. No unrelated edits are present.

- [ ] **Step 10: Commit the verified revision**

Run:

```bash
git add Online_Retail_Repeat_Purchase_Project.ipynb Detailed_Project_Pipeline_Aligned.docx Group_Project_Report.docx tests/test_pipeline_contract.py
git commit -m "docs: align project evidence with corrected methodology"
```

## Completion criteria

The revision is complete only when all of these statements are true:

- A clean checkout can rebuild the environment and read the zipped source dataset.
- The notebook runs from top to bottom without manual file renaming.
- EDA tables and figures exist before the modelling sections.
- All merchandise and non-product transaction codes are explicitly classified.
- Semantic features contain merchandise descriptions only.
- The TF-IDF benchmark uses LSA or another reduced representation with reported cluster quality.
- Best-K and matched-K text comparisons are both reported.
- All central model comparisons include paired customer-grouped uncertainty intervals.
- Snapshot and target durations are enforced by tests.
- The executed notebook, pipeline document, report, and checkpoint use the same v2 numbers.
- Any claim of improvement is supported by an interval that excludes zero.
