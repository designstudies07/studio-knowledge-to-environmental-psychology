# Public data dictionary

## `participant_flow.csv`

- `stage`: T0, T1, or T2.
- `raw_rows`: number of submitted rows before duplicate handling.
- `unique_students`: number of retained unique responses.
- `duplicates_removed`: duplicate rows removed under the primary latest-response rule.
- `alpha_all`: Cronbach's alpha in the available stage-specific sample.
- `matched_n`: complete longitudinal cases contributing at that stage.
- `alpha_matched`: Cronbach's alpha within the 34-person matched sample.

## `composite_trajectory.csv`

- `stage`: measurement stage.
- `n`: matched sample size.
- `mean`: eight-item composite mean.
- `sd`: standard deviation.

## `composite_pairwise.csv`

- `contrast`: paired stage comparison.
- `holm_p`: Holm-adjusted Wilcoxon p value.
- `rank_biserial_r`: paired rank-biserial correlation, oriented so positive values indicate a higher score at the later stage.

## `item_results.csv`

- `measure`: self-report item label.
- `t0_mean`, `t0_sd`, `t1_mean`, `t1_sd`, `t2_mean`, `t2_sd`: matched descriptive statistics.
- `friedman_q`, `friedman_p`: omnibus repeated-measures test.
- `significant_contrasts`: Holm-significant paired contrasts, or `None`.

## `assignment_dimensions.csv`

- `dimension`: assignment scoring dimension.
- `n`: number of scored assignments.
- `mean`, `median`, `sd`: score summary on the 0-3 scale.

## `open_ended_themes.csv`

- `theme`: concept-guided multi-label indicator.
- `t0_proportion`, `t1_proportion`, `t2_proportion`: proportion of responses receiving the indicator at each stage.

The public files contain aggregates only and cannot reproduce participant-level inferential tests without the restricted inputs.

## `exploratory_associations.csv`

- `analysis`: exploratory comparison or association.
- `n`: total number of matched cases contributing to the analysis.
- `estimate`: Spearman correlation when applicable.
- `p_value`: adjusted or unadjusted p value, as identified by `p_adjustment`.
- `p_adjustment`: multiplicity adjustment applied to the reported p value.
- `group_1`, `group_1_n`, `group_1_mean`: first comparison group and its aggregate summary.
- `group_2`, `group_2_n`, `group_2_mean`: second comparison group and its aggregate summary.
- `effect_size`: named aggregate effect-size estimate.
