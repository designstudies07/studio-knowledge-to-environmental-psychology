# GitHub and Zenodo release checklist

- [ ] Confirm the final author list, affiliations, and preferred repository name.
- [ ] Confirm the journal's code and data availability policy.
- [ ] Confirm that the ethics approval permits the intended public materials.
- [ ] Keep `data/private/` empty except for its README.
- [ ] Search the repository for names, email addresses, absolute local paths, and linkage material.
- [ ] Verify that no raw spreadsheet, assignment, verbatim response, timestamp, salt, or alias map is tracked.
- [ ] Verify that public figures contain no individual trajectories, participant-level points, or small-cell distributions.
- [ ] Run `python src/verify_public_repository.py`.
- [ ] Re-run the analysis in a clean virtual environment using restricted local inputs.
- [ ] Check all figures and aggregate CSVs against the accepted manuscript.
- [ ] Add final citation metadata and an institutionally approved license.
- [ ] Create a tagged GitHub release.
- [ ] Archive the release with Zenodo if a DOI is needed.
- [ ] Replace the manuscript's provisional availability statement with the final repository URL and DOI.
