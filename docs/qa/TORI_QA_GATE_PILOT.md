# TORI QA evidence pilot — solar
This repository is a staged QA-enforcement prototype from Tori_Common_Project PR #5. No existing development branches were touched.

Every PR should have a specific feature evidence file in `qa/evidence/<feature>.json`, with the common contract fields. The QA job fails if there are no evidence JSON files or they fail validation. The validator verifies declared evidence structure, not the authenticity of referenced proof; independent source/log review remains mandatory.

**Caveat:** The current draft workflow checks all evidence files in the repository, rather than just the changed feature. It can block preexisting PRs without evidence. Do not merge or mark the check as required until an evidence migration and PR-level scoping are reviewed. Never invent PASS evidence to satisfy the gate.

Actual model training, browser tests, data lineage and security checks remain project-specific. Acceptance is NOT complete until they run where applicable.
