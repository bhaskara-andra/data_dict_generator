# data_dict_generator
AI tool to create the Data Dictionary across systems

## Data Dictionary Generator v1

A Streamlit application that connects to a database, auto-discovers schemas/tables/columns via `information_schema`, generates a formatted Excel-based data dictionary, packages it as a downloadable zip, and (optionally) stores it in ChromaDB for RAG-powered chat/observability.

---

## Features

- **Encrypted connection vault** — save aliased DB connectors (Postgres, MySQL, MSSQL, Snowflake, Oracle) with credentials encrypted at rest using `cryptography.Fernet`.
- **Test Connection** button before you commit a connector.
- **Flexible input** — supply a list of databases / schemas / `db.schema.table` entries, or leave blank to auto-discover everything (system schemas are always excluded).
- **Automatic metadata extraction** — pulls tables, columns, data types, and nullability from `information_schema`.
- **Rule-based workbook layout**:
  - ≤ 13 tables in a schema → one sheet per table
  - \> 13 tables in a schema → one summary sheet per schema
  - \> 150 tables in the whole database → single workbook, one sheet per schema
- **Zip output** — all generated workbook(s) bundled into one downloadable `.zip`.
- **Cache-and-clear** — intermediate extraction results are cached (via `diskcache`, with a dict fallback) during generation and cleared immediately after the zip is built, so nothing lingers in memory/disk longer than needed.
- **Chatbot-ready** — optionally store the generated dictionary as embeddings in ChromaDB and query it from an in-app chat panel (RAG retrieval stub included, ready to wire into an LLM call).

---

## Project layout

```
data_dictionary_generator_v1.py   # single-file Streamlit app
vault.json                        # created at runtime — encrypted connector store
.vault.key                        # created at runtime — Fernet key (keep secret, do not commit)
.ddg_cache/                       # created at runtime — diskcache working directory
chroma_store/                     # created at runtime — persistent ChromaDB store
```

---

## Requirements

```bash
pip install streamlit sqlalchemy cryptography pandas openpyxl diskcache chromadb sentence-transformers redis
```

Depending on which databases you connect to, also install the matching driver:

| Connector  | Driver package        |
|------------|------------------------|
| PostgreSQL | `psycopg2-binary`      |
| MySQL      | `pymysql`               |
| MSSQL      | `pyodbc` + ODBC Driver 17 for SQL Server |
| Snowflake  | `snowflake-sqlalchemy` |
| Oracle     | `cx_Oracle`             |

---

## Running the app

```bash
streamlit run data_dictionary_generator_v1.py
```

Then open the URL Streamlit prints (default `http://localhost:8501`).

---

## Usage

### 1. Connections tab
1. Fill in alias, connector type, host, port, database, user, password.
2. Click **Test Connection** to validate.
3. Click **Save to Vault** to encrypt and persist it under the alias.

### 2. Data Dictionary Studio tab
- **Left / Input panel**: pick a saved connector alias (mandatory), optionally paste a list of databases/schemas/tables, then click **Generate Data Dictionary**.
- **Left / Output panel**: once generation finishes, download the zip. Check "Store dictionary for chatbot" beforehand if you want it embedded into ChromaDB.
- **Middle / Viewer**: browse the generated catalog schema-by-schema, table-by-table.
- **Right / Observability & Chat**: see quick counts (schemas/tables/columns) and ask questions about the generated dictionary via the chat box.

---

## Extending this v1

The following are intentionally left as lightweight stubs so the app stays runnable end-to-end; swap them out for production use:

- `generate_table_description()` / `generate_column_description()` — currently rule-based placeholder text. Point these at an LLM (e.g. the Anthropic API) for real, context-aware descriptions.
- `rag_query()` — retrieves top-k matching chunks from ChromaDB but does not yet synthesize a final answer via an LLM. Wire in a completion call using the retrieved context.
- Caching — uses `diskcache` locally. Swap the `CACHE` object for a `redis.Redis()` client if you need distributed/shared caching across app instances.
- Connection URIs — `DRIVER_MAP` covers common cases; adjust for edge cases like Snowflake account identifiers or Oracle service name formats specific to your environment.

---

## Security notes

- `.vault.key` is the encryption key for all stored credentials — treat it like a secret (exclude from version control, restrict file permissions).
- Consider rotating the Fernet key and re-encrypting `vault.json` periodically in production.
- The current vault is file-based for simplicity; for multi-user/production deployments, back it with a proper secrets manager (e.g. Azure Key Vault, AWS Secrets Manager, HashiCorp Vault).
