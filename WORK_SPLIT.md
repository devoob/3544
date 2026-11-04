# Work split (proposal, to be confirmed by the team)

Grading areas: **A** Workload, **B** Reasoning quality, **C** Presentation, **D** Report.
The pipeline notebook already exists (Codex-generated), so nobody waits at the start. Each person re-runs their own stage in parallel, adds one experiment of their own, writes their own report section, and presents it.

## Who owns what
| Person | Stages | Produces | Area |
|---|---|---|---|
| Stacia (data) | 1-3: raw data, cleaning, invoice table | Data audit table (5,268 duplicates, 135,080 missing IDs, cancellations); revenue reconciliation; "duplicates kept vs dropped" rerun; 2-3 EDA plots; dataset sourcing notes | A |
| Devano (text) | 8-9 clusters; 30-32 error analysis | Named clusters from `product_cluster_examples.csv`; K = 8/12/16/20 silhouette table; why MiniLM gives about no gain; error-analysis table; limitations | A, B |
| Clara (pipeline lead) | 10-13: snapshots, features, models, evaluation | Leakage check; baseline vs model comparison; July selection; different threshold rule experiment; iteration log; merges everyone's work | B |

## How everyone contributes
- Re-run your own stage in `.venv` and check it against `notebook_outputs/`.
- Run one experiment of your own (the notebook doesn't have these).
- Write your own report section and present it.
- Review in a circle: Stacia checks Clara, Clara checks Devano, Devano checks Stacia.
- Have something real done by 6 Nov so your checkpoint row has evidence (Stacia: audit table; Devano: cluster examples reviewed; Clara: pipeline reproduced).

## Timeline and dependencies
| Dates | What | Depends on |
|---|---|---|
| 3-4 Nov | All read `Detailed_Project_Pipeline_Aligned.docx` and get the notebook running. Clara keeps an untouched "v1" copy. | nothing |
| 4-9 Nov | Parallel tracks, no blocking: Stacia audit + duplicate rerun (in a copy of the notebook); Devano cluster review + error analysis (uses existing `final_test_predictions.csv`); Clara leakage check + baseline comparison | v1 copy |
| by 7 Nov | Each member submits their own checkpoint PDF | own evidence from the line above |
| 9-14 Nov | Iteration log and final figures (Clara) | **Stacia's sensitivity result and Devano's K result**. Stacia and Devano start writing their sections as soon as their own results are in. |
| 15-20 Nov | Report draft, merged into one story (Clara) | all three sections |
| 21-25 Nov | Slides, then rehearsal | report draft |

Risk: if Stacia's audit finds something that should change the main cleaning, every downstream number changes and Clara must re-run the whole pipeline, so tell Clara early. If it is only a sensitivity result, it stays a side experiment and blocks no one.

Final report and presentation deadlines are still to be confirmed.
