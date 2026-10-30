# Learning Environmental Psychology From Intuitive Human Centred Understanding to Spatial Analysis

Reproducibility materials for a three-stage study of third-year interior architecture students. The study examined students' intuitive human-centred understanding before a concept-focused lecture (T0), their environmental psychology learning immediately after Lecture 1 (T1), and their spatial analysis after using the concepts to examine a personally experienced interior (T2).

## Repository scope

This repository contains:

- the Python analysis and manuscript-building pipeline;
- publication-ready, privacy-safe figures;
- aggregate, non-identifying result tables;
- a data dictionary and analysis notes; and
- a checklist for creating a public GitHub/Zenodo release.

It intentionally does **not** contain raw survey exports, student names, email addresses, verbatim responses, submitted assignments, linkage keys, hashing salts, or participant-level analytic files.

## Study framing

The three observations represent successive states of conceptualization, not a randomized comparison of teaching methods:

- **T0:** intuitive human-centred understanding before the concept-focused lecture;
- **T1:** environmental psychology learning immediately after Lecture 1;
- **T2:** spatial analysis after using the concepts to examine a real interior.

Lecture 1 was the first unit of a broader course and does not represent the full course content.

## Repository structure

```text
.
|-- data/
|   |-- derived/                 # Aggregate tables safe for public release
|   `-- private/                 # Local-only restricted inputs; ignored by Git
|-- docs/                        # Methods, data dictionary, release checklist
|-- figures/                     # Publication-ready PNG figures
|-- outputs/                     # Generated manuscript output; ignored by Git
|-- src/
|   `-- build_manuscript.py      # Analysis, figures, tables, and DOCX builder
|-- .gitignore
|-- requirements.txt
`-- README.md
```

## Reproducing the restricted analysis locally

1. Use Python 3.12.
2. Create and activate a virtual environment.
3. Install dependencies:

   ```bash
   python -m pip install -r requirements.txt
   ```

4. Place the restricted files described in `data/private/README.md` in that directory.
5. Provide the confidential linkage salt through the environment variable `STUDY_LINKAGE_SALT` or a local `data/private/hash_salt.txt` file.
6. Run:

   ```bash
   python src/build_manuscript.py
   ```

The script writes the manuscript to `outputs/` and refreshes the five figures in `figures/`.

The private inputs can also be supplied with `STUDY_ANALYSIS_WORKBOOK` and `STUDY_T0_WORKBOOK`. Do not expose the paths or values in public logs.

## Public data policy

The participant-level data are withheld because repeated survey responses and course assignments may remain re-identifiable even after direct identifiers are removed. Public materials are limited to aggregate tables, analysis code, and figures. Any controlled sharing of participant-level data should require institutional ethics approval, an appropriate data-use agreement, and secure transfer.

Suggested manuscript statement:

> Analysis code, aggregate result tables, and publication figures are available in the accompanying repository. Participant-level survey and assignment data are not publicly deposited because linked educational records and open-ended responses may create a residual risk of re-identification. Access requests will be considered subject to institutional ethics approval and data-protection requirements.

## Main analytic sample

- T0: 60 unique responses from 63 rows
- T1: 53 unique responses from 53 rows
- T2: 40 unique responses from 48 rows
- Complete longitudinal sample: 34 students
- Research-scored assignments: 44 submissions
- Complete cases with a matched assignment: 28 students

## Release guidance

Before creating a versioned release, complete `docs/github_release_checklist.md`, confirm the ethics and journal data-sharing requirements, and add the final author and citation information. A tagged GitHub release can be archived through Zenodo to obtain a DOI.

## License

No reuse license is assigned in this preparation package. Add a license only after the authors and institution confirm ownership and the intended terms of reuse.
