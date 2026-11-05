# Work split and current progress

Last updated: 5 November 2026

Grading areas: **A** Workload, **B** Reasoning quality, **C** Presentation, **D** Report.

The notebook is now reproducible from the checked-in archive, has been executed end to end, and passes 19 contract tests. The work below separates shared pipeline outputs from each person's assigned review and reporting work. Shared code does not by itself prove that every member completed their individual rerun.

## Completed shared work

| Area | Completed output | Evidence |
|---|---|---|
| Data access | The notebook reads `Online Retail.xlsx` directly from `data/online+retail.zip` | Executed notebook, `notebook_outputs/raw_data_audit.csv` |
| EDA and cleaning | Missingness, monthly activity, concentration, country and target-balance tables and figures | `notebook_outputs/eda_*` |
| Product taxonomy | Merchandise products separated from postage, carriage, adjustments, fees and vouchers | `notebook_outputs/non_product_code_audit.csv`, `invoice_revenue_audit.csv` |
| Text benchmark | Raw sparse TF-IDF replaced by 100-component LSA before KMeans | `cluster_quality.csv`, `cluster_matched_k_quality.csv` |
| Evaluation | Customer-grouped bootstrap intervals and paired model differences | `metric_uncertainty.csv`, `paired_model_differences.csv` |
| Testing | Temporal windows, identifiers, product scope, semantic shares and uncertainty files checked | `tests/test_pipeline_contract.py`, 19 passing tests |
| Reporting | Report, pipeline document and Jonathan's two-page checkpoint updated to v2 | `.docx` and checkpoint `.pdf` outputs |

## Ownership and contribution

| Person | Assigned stages | Completed or current work | Evidence | Approx. hours | Remaining personal task |
|---|---|---|---|---|---|
| Stacia Benita | 1-3: data audit, cleaning, invoice construction | Review the shared EDA, product/charge split, revenue audit and duplicate sensitivity | `eda_data_quality.csv`, `eda_monthly_activity.csv`, `invoice_revenue_audit.csv` | 18 | Confirm the data section and independently rerun her stage |
| Devano Jonathan | 8-9: text features; 30-32: interpretation and errors | Reviewed the methodology, identified the weak sparse TF-IDF baseline, specified LSA and matched K=16, checked cluster and error outputs, prepared the checkpoint, and verified the final artifacts | `cluster_quality.csv`, `cluster_matched_k_quality.csv`, `product_cluster_examples.csv`, `error_analysis_summary.csv`, `tests/test_pipeline_contract.py` | 20 | Accept the descriptive cluster labels |
| Clara Natalie Surjadjajadi | 10-13 and integration | Review snapshot construction, model ablation, July selection, rolling evaluation and merged report | `customer_snapshot_features.csv`, `validation_results.csv`, `rolling_monthly_results.csv`, `paired_model_differences.csv` | 19 | Verify the pipeline section and merge final text |

## Jonathan's completed contribution

Jonathan's assigned text track covers the product representations and the error interpretation. His work in this version also includes the methodology review that changed the final experiment:

- Compared the local v1 pipeline with the newer GitHub notebook and rejected raw sparse TF-IDF KMeans as an unfair benchmark.
- Adopted LSA-TF-IDF with 100 components and a matched K=16 sensitivity comparison.
- Used cosine silhouette and largest-cluster share to test whether TF-IDF and MiniLM form credible groups.
- Interpreted October false positives and false negatives against recency, frequency, spend and product variety.
- Required paired customer-grouped intervals before claiming a model improvement.
- Directed the notebook repair, reviewed the executed outputs, and prepared the report and checkpoint evidence.

Codex generated code and document drafts under Jonathan's direction. The checks against evidence include a clean top-to-bottom notebook execution, 19 passing tests, rendered report pages, and a two-page checkpoint. Jonathan still needs to complete the human interpretation of cluster names.

## Review circle

- Stacia checks the data and cleaning claims.
- Clara checks the text and error-analysis claims.
- Jonathan checks the pipeline, uncertainty and report integration.

Each reviewer should record the file or table used for their check.

## Current schedule

| Date | Work | Dependency or output |
|---|---|---|
| 5-6 November | Each member reviews their assigned stage and confirms personal evidence | Independent checkpoint evidence |
| 7 November | Each member submits a two-page checkpoint PDF with actual hours | Name, student ID, evidence, hours and AI-check note |
| 9-14 November | Reconcile sensitivity, cluster and model results; finalise figures | Agreed final tables and iteration notes |
| 15-20 November | Merge and edit the final report | One consistent set of v2 results |
| 21-25 November | Build slides, rehearse and check numbers | Final presentation |

## Remaining risks

- Each member still needs to confirm their own approximate hours and independent review.
- Cluster descriptions require human judgement even when the silhouette improves.
- The final presentation must use the v2 values in the executed notebook, not the older v1 tables.
- Course deadlines for the final report and presentation are still unconfirmed.
