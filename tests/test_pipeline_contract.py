import json
from pathlib import Path
from zipfile import ZipFile

import numpy as np
import pandas as pd
import pytest


ROOT = Path(__file__).resolve().parents[1]
ARCHIVE = ROOT / "data" / "online+retail.zip"
OUTPUT_DIR = ROOT / "notebook_outputs"
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


def test_non_product_codes_are_reported():
    audit = pd.read_csv(OUTPUT_DIR / "non_product_code_audit.csv")
    non_product = audit.loc[audit["IsNonProduct"]]
    assert set(non_product["StockCode"]).issubset(
        NON_PRODUCT_CODES
        | {
            "gift_0001_10",
            "gift_0001_20",
            "gift_0001_30",
            "gift_0001_40",
            "gift_0001_50",
        }
    )
    assert non_product["IncludedInProductFeatures"].eq(False).all()


def test_cluster_quality_has_reduced_tfidf_benchmark():
    quality = pd.read_csv(OUTPUT_DIR / "cluster_quality.csv")
    assert {"TF-IDF + LSA", "MiniLM"} <= set(quality["Representation"])
    assert quality["SilhouetteCosine"].notna().all()
    assert quality["LargestClusterShare"].between(0, 1).all()


def test_matched_cluster_count_is_reported_for_both_representations():
    matched = pd.read_csv(OUTPUT_DIR / "cluster_matched_k_quality.csv")
    assert set(matched["Representation"]) == {"TF-IDF + LSA", "MiniLM"}
    assert matched["K"].nunique() == 1


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
    assert (snapshots["FeatureEndExclusive"] - snapshots["FeatureStart"]).dt.days.eq(
        90
    ).all()
    assert (snapshots["TargetEndExclusive"] - snapshots["TargetStart"]).dt.days.eq(
        30
    ).all()
    assert snapshots["FeatureEndExclusive"].eq(snapshots["TargetStart"]).all()
    assert not snapshots.duplicated(["SnapshotDate", "CustomerID"]).any()


def test_snapshot_dates_are_the_eight_pre_registered_months():
    snapshots = pd.read_csv(OUTPUT_DIR / "customer_snapshot_features.csv")
    observed = sorted(pd.to_datetime(snapshots["SnapshotDate"].unique()))
    expected = list(
        pd.to_datetime(
            [
                "2011-03-31",
                "2011-04-30",
                "2011-05-31",
                "2011-06-30",
                "2011-07-31",
                "2011-08-31",
                "2011-09-30",
                "2011-10-31",
            ]
        )
    )
    assert observed == expected


def test_semantic_cluster_shares_cover_the_full_product_history():
    snapshots = pd.read_csv(OUTPUT_DIR / "customer_snapshot_features.csv")
    for prefix in ["TFIDFBest", "MiniLMBest", "TFIDFK16", "MiniLMK16"]:
        columns = [
            column
            for column in snapshots.columns
            if column.startswith(prefix) and column.endswith("Share")
        ]
        assert columns
        totals = snapshots[columns].sum(axis=1)
        assert np.allclose(totals, 1.0, atol=1e-8)


def test_identifiers_are_not_predictors():
    notebook = json.loads(
        (ROOT / "Online_Retail_Repeat_Purchase_Project.ipynb").read_text()
    )
    source = "\n".join(
        "".join(cell.get("source", []))
        for cell in notebook["cells"]
        if cell["cell_type"] == "code"
    )
    feature_block = source.split("RFM_FEATURES =", 1)[1].split(
        "feature_sets =", 1
    )[0]
    assert "CustomerID" not in feature_block
    assert "SnapshotDate" not in feature_block
