# Change Log

## 2026-07-21

### Migrated `settings.ini` secrets into `.env`

Consolidated application secrets so that `.env` is the single canonical configuration
store, removing the dependency on `settings.ini` (`configparser`).

**`graph_api/secrets/.env`**
- Added all five `settings.ini` sections as environment variables using the existing
  `UPPER_SNAKE` + section-prefix naming convention:
  - `[DomoDB]` → `DOMODB_*`
  - `[DomoDB_Internal]` → `DOMODB_INTERNAL_*`
  - `[SMTP]` → `SMTP_*`
  - `[SMTP_Internal]` → `SMTP_INTERNAL_*`
  - `[Integration]` → `INTEGRATION_*`
- Per-section keys: `_USERNAME`, `_PASSWORD`, `_SERVER`, `_PORT`, plus `_DATABASE`/`_DRIVER`
  (database sections) or `_SENDER`/`_RECIPIENTS` (SMTP sections). Driver strings retain their
  exact original value (including the `driver=` / `?driver=` prefix).

**`graph_api/graph_connector_app/sqlalchemy_models/sql_models.py`**
- Removed `configparser` and the `Config.read('./settings.ini')` block.
- Switched to `django-environ` (the same library already used in `settings.py`); loads
  `secrets/.env` for local runs, and relies on docker-compose's `env_file` injection in
  containers.
- `DatabaseConnection.__init__` now resolves a section-to-env-prefix map and reads each
  value from the environment. The generated connection strings are byte-for-byte identical
  to the previous `configparser`-based output for all sections.

**`graph_api/secrets/settings.ini`**
- Left in place (now unused); no longer read by any code.

**Verification**
- Confirmed connection-string parity between the old `configparser` output and the new
  environment-based output for `DomoDB`, `DomoDB_Internal`, and `Integration`, plus value
  parity for the SMTP sections.
- `sql_models.py` compiles cleanly; no remaining references to `configparser`/`settings.ini`
  or the removed `Config`/`working_directory` symbols.

**Deployment note**
- The `lb_graph_django` container must be rebuilt and recreated
  (`docker compose up -d --build django`) to pick up the new code and `.env` values.
