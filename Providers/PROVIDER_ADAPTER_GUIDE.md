# Adding a licensed provider

1. Obtain the provider's current API documentation and a written confirmation of the intended use rights.
2. Implement a class conforming to `RacingDataProvider` in `base.py`.
3. Normalize provider responses into the CSV canonical field names or write normalized rows directly into a repository/service layer.
4. Store provider-specific identifiers in a future `external_ids` table; do not use horse names as permanent identity keys.
5. Preserve source, update timestamp, provider record ID, and field-level provenance.
6. Add idempotent upserts, retry/backoff, rate-limit handling, schema validation, logging and tests.
7. Test non-runners, late changes, abandoned races, duplicate horse names, missing ratings, and amended official marks.
8. Keep API credentials in environment variables or a secret manager.

Do not assume a personal-use API plan allows publication or resale. Confirm storage, derivative-work, cache, display, and redistribution rights in writing before connecting the feed to a public site.
