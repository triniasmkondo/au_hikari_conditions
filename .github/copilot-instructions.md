# Repository guidance

- Keep the P0 implementation dependency-light and use the existing `src/au_hikari_conditions` package.
- Treat the project specification as the source of truth for folder names, deliverable item IDs, and business rules.
- Do not invent deliverable IDs or production data; update the placeholder dictionary only from approved specifications.
- Keep operational files and audit records in their designated numbered directories.
- Run `python scripts/check_environment.py` and `python -m unittest discover -s tests` after changes.
