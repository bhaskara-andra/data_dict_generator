"""
Data Dictionary Generator v1
------------------------------------------------------------
pip install streamlit sqlalchemy cryptography pandas openpyxl diskcache chromadb sentence-transformers redis

Run:  streamlit run data_dictionary_generator_v1.py
------------------------------------------------------------
Structure:
  Tab 1 - Connections   : add/alias DB connectors, encrypted vault, test connection
  Tab 2 - Data Dictionary Studio (main page)
      Left   : Input panel + Output panel (download zip)
      Middle : Viewer canvas (table/column preview)
      Right  : Observability / Chat panel (RAG over generated dictionary)
"""

import os, io, json, zipfile, base64
import streamlit as st
import pandas as pd
from datetime import datetime
from cryptography.fernet import Fernet
from sqlalchemy import create_engine, text

try:
    import diskcache
    CACHE = diskcache.Cache("./.ddg_cache")
except Exception:
    CACHE = {}  # fallback dict-cache

# ---------------------------------------------------------------------------
# 1. VAULT (encrypted connection storage)
# ---------------------------------------------------------------------------
VAULT_FILE = "vault.json"
KEY_FILE = ".vault.key"

def get_fernet():
    if not os.path.exists(KEY_FILE):
        with open(KEY_FILE, "wb") as f:
            f.write(Fernet.generate_key())
    with open(KEY_FILE, "rb") as f:
        return Fernet(f.read())

def load_vault():
    if not os.path.exists(VAULT_FILE):
        return {}
    f = get_fernet()
    with open(VAULT_FILE, "rb") as fh:
        try:
            return json.loads(f.decrypt(fh.read()).decode())
        except Exception:
            return {}

def save_vault(vault: dict):
    f = get_fernet()
    with open(VAULT_FILE, "wb") as fh:
        fh.write(f.encrypt(json.dumps(vault).encode()))

DRIVER_MAP = {
    "PostgreSQL": "postgresql+psycopg2://{user}:{pwd}@{host}:{port}/{db}",
    "MySQL": "mysql+pymysql://{user}:{pwd}@{host}:{port}/{db}",
    "MSSQL": "mssql+pyodbc://{user}:{pwd}@{host}:{port}/{db}?driver=ODBC+Driver+17+for+SQL+Server",
    "Snowflake": "snowflake://{user}:{pwd}@{host}/{db}",
    "Oracle": "oracle+cx_oracle://{user}:{pwd}@{host}:{port}/?service_name={db}",
}

SYSTEM_SCHEMAS = {
    "PostgreSQL": {"pg_catalog", "information_schema", "pg_toast"},
    "MySQL": {"mysql", "information_schema", "performance_schema", "sys"},
    "MSSQL": {"sys", "information_schema", "guest", "db_owner"},
    "Snowflake": {"information_schema"},
    "Oracle": {"sys", "system"},
}

def build_uri(conn: dict) -> str:
    return DRIVER_MAP[conn["type"]].format(
        user=conn["user"], pwd=conn["password"], host=conn["host"],
        port=conn.get("port", ""), db=conn.get("database", "")
    )

def test_connection(conn: dict):
    try:
        engine = create_engine(build_uri(conn))
        with engine.connect() as c:
            c.execute(text("SELECT 1"))
        return True, "Connection successful"
    except Exception as e:
        return False, str(e)

# ---------------------------------------------------------------------------
# 2. METADATA EXTRACTION
# ---------------------------------------------------------------------------
def get_schemas(engine, db_type):
    q = "SELECT schema_name FROM information_schema.schemata"
    df = pd.read_sql(text(q), engine)
    col = df.columns[0]
    return [s for s in df[col] if s.lower() not in SYSTEM_SCHEMAS.get(db_type, set())]

def get_tables(engine, schema):
    q = f"""SELECT table_schema, table_name FROM information_schema.tables
            WHERE table_schema = :schema"""
    return pd.read_sql(text(q), engine, params={"schema": schema})

def get_columns(engine, schema, table):
    q = """SELECT column_name, data_type, is_nullable, character_maximum_length
           FROM information_schema.columns
           WHERE table_schema = :schema AND table_name = :table
           ORDER BY ordinal_position"""
    return pd.read_sql(text(q), engine, params={"schema": schema, "table": table})

# ---------------------------------------------------------------------------
# 3. DESCRIPTION GENERATION (pluggable - heuristic fallback / LLM hook)
# ---------------------------------------------------------------------------
def generate_table_description(schema, table, columns_df):
    return f"Table '{table}' in schema '{schema}' with {len(columns_df)} columns."

def generate_column_description(col_name, data_type):
    return f"Column '{col_name}' of type {data_type}."

# ---------------------------------------------------------------------------
# 4. EXCEL WORKBOOK BUILDER (per business rules)
# ---------------------------------------------------------------------------
def safe_sheet_name(name):
    return name[:31].replace("/", "_").replace("\\", "_")

def build_workbooks(catalog: dict):
    """
    catalog = {schema: {table: {desc, columns_df}}}
    Rule: if a schema has >13 tables -> one sheet per SCHEMA (not per table)
          if total tables across DB > 150 -> one sheet per SCHEMA for whole DB
          else -> one sheet per TABLE
    Returns dict {workbook_filename: BytesIO}
    """
    total_tables = sum(len(t) for t in catalog.values())
    files = {}

    def schema_summary_df(tables):
        rows = []
        for tname, meta in tables.items():
            for _, r in meta["columns"].iterrows():
                rows.append({
                    "table_name": tname, "table_description": meta["desc"],
                    "column_name": r["column_name"], "data_type": r["data_type"],
                    "nullable": r["is_nullable"],
                    "column_description": generate_column_description(r["column_name"], r["data_type"]),
                })
        return pd.DataFrame(rows)

    if total_tables > 150:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            for schema, tables in catalog.items():
                schema_summary_df(tables).to_excel(writer, sheet_name=safe_sheet_name(schema), index=False)
        buf.seek(0)
        files["data_dictionary_full_database.xlsx"] = buf
        return files

    for schema, tables in catalog.items():
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            if len(tables) > 13:
                schema_summary_df(tables).to_excel(writer, sheet_name=safe_sheet_name(schema), index=False)
            else:
                for tname, meta in tables.items():
                    df = meta["columns"].copy()
                    df["column_description"] = df.apply(
                        lambda r: generate_column_description(r["column_name"], r["data_type"]), axis=1)
                    df.insert(0, "table_description", meta["desc"])
                    df.to_excel(writer, sheet_name=safe_sheet_name(tname), index=False)
        buf.seek(0)
        files[f"data_dictionary_{schema}.xlsx"] = buf
    return files

def zip_workbooks(files: dict) -> io.BytesIO:
    zbuf = io.BytesIO()
    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, buf in files.items():
            z.writestr(name, buf.getvalue())
    zbuf.seek(0)
    return zbuf

# ---------------------------------------------------------------------------
# 5. RAG / CHATBOT STORAGE (ChromaDB)
# ---------------------------------------------------------------------------
def store_to_chroma(catalog: dict, collection_name="data_dictionary"):
    try:
        import chromadb
        client = chromadb.PersistentClient(path="./chroma_store")
        coll = client.get_or_create_collection(collection_name)
        docs, ids, metas = [], [], []
        for schema, tables in catalog.items():
            for tname, meta in tables.items():
                text_blob = meta["desc"] + " Columns: " + ", ".join(
                    f"{r['column_name']}({r['data_type']})" for _, r in meta["columns"].iterrows())
                docs.append(text_blob)
                ids.append(f"{schema}.{tname}")
                metas.append({"schema": schema, "table": tname})
        if docs:
            coll.upsert(documents=docs, ids=ids, metadatas=metas)
        return True, f"Stored {len(docs)} table docs to ChromaDB."
    except Exception as e:
        return False, str(e)

def rag_query(question, collection_name="data_dictionary", k=4):
    try:
        import chromadb
        client = chromadb.PersistentClient(path="./chroma_store")
        coll = client.get_or_create_collection(collection_name)
        res = coll.query(query_texts=[question], n_results=k)
        docs = res.get("documents", [[]])[0]
        context = "\n---\n".join(docs)
        return f"Relevant context:\n{context}\n\n(Hook this into your LLM call for a synthesized answer.)"
    except Exception as e:
        return f"RAG error: {e}"

# ---------------------------------------------------------------------------
# 6. STREAMLIT UI
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Data Dictionary Generator v1", layout="wide")
if "vault" not in st.session_state:
    st.session_state.vault = load_vault()
if "catalog" not in st.session_state:
    st.session_state.catalog = None
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

tab_conn, tab_main = st.tabs(["🔌 Connections", "📚 Data Dictionary Studio"])

# --- Connections tab ---
with tab_conn:
    st.subheader("Configure a Connector")
    c1, c2 = st.columns(2)
    with c1:
        alias = st.text_input("Connector alias (name you'll select later)")
        conn_type = st.selectbox("Connector type", list(DRIVER_MAP.keys()))
        host = st.text_input("Host")
        port = st.text_input("Port")
    with c2:
        database = st.text_input("Database")
        user = st.text_input("User")
        password = st.text_input("Password", type="password")

    bc1, bc2 = st.columns(2)
    candidate = {"type": conn_type, "host": host, "port": port,
                 "database": database, "user": user, "password": password}
    if bc1.button("Test Connection"):
        ok, msg = test_connection(candidate)
        (st.success if ok else st.error)(msg)
    if bc2.button("Save to Vault"):
        if not alias:
            st.warning("Alias is required.")
        else:
            st.session_state.vault[alias] = candidate
            save_vault(st.session_state.vault)
            st.success(f"Saved connector '{alias}' to encrypted vault.")

    st.divider()
    st.subheader("Saved Connectors")
    if st.session_state.vault:
        st.table(pd.DataFrame([
            {"alias": a, "type": v["type"], "host": v["host"], "database": v["database"]}
            for a, v in st.session_state.vault.items()
        ]))
    else:
        st.info("No connectors saved yet.")

# --- Main page ---
with tab_main:
    st.title("📚 Data Dictionary Studio")
    left, mid, right = st.columns([1.1, 1.6, 1.1])

    # ---------------- LEFT: Input + Output panels ----------------
    with left:
        st.markdown("### Input Panel")
        if not st.session_state.vault:
            st.warning("Add a connector first in the Connections tab.")
        alias_sel = st.selectbox("Connector (mandatory)", list(st.session_state.vault.keys()) or ["-"])
        raw_input = st.text_area(
            "Optional list: database / schema / db.schema.table (one per line). "
            "Leave blank to auto-discover everything.", height=140)
        run_btn = st.button("🚀 Generate Data Dictionary", use_container_width=True)

        st.markdown("### Output Panel")
        if st.session_state.get("zip_buf"):
            st.download_button("⬇️ Download Data Dictionary (.zip)",
                                data=st.session_state["zip_buf"],
                                file_name=f"data_dictionary_{datetime.now():%Y%m%d_%H%M%S}.zip",
                                mime="application/zip", use_container_width=True)
        rag_toggle = st.checkbox("Store dictionary for chatbot (ChromaDB RAG)")

    # ---------------- Generation pipeline ----------------
    if run_btn and alias_sel in st.session_state.vault:
        conn = st.session_state.vault[alias_sel]
        try:
            engine = create_engine(build_uri(conn))
            CACHE["last_conn_alias"] = alias_sel

            # Step 1: parse input list
            lines = [l.strip() for l in raw_input.splitlines() if l.strip()]

            # Step 2: connection already validated via engine creation
            # Step 3: build target schema list (skip system schemas)
            target_schemas = set()
            if not lines:
                target_schemas.update(get_schemas(engine, conn["type"]))
            else:
                for l in lines:
                    parts = l.split(".")
                    if len(parts) >= 2:
                        target_schemas.add(parts[-2] if len(parts) == 3 else parts[0])
                    else:
                        target_schemas.update(get_schemas(engine, conn["type"]))
            target_schemas -= SYSTEM_SCHEMAS.get(conn["type"], set())

            # Step 4 & 5: pull columns, build descriptions
            catalog = {}
            progress = st.progress(0.0, text="Extracting metadata...")
            for i, schema in enumerate(sorted(target_schemas)):
                tables_df = get_tables(engine, schema)
                catalog[schema] = {}
                for tname in tables_df["table_name"]:
                    cols = get_columns(engine, schema, tname)
                    desc = generate_table_description(schema, tname, cols)
                    catalog[schema][tname] = {"desc": desc, "columns": cols}
                    CACHE[f"cache:{schema}.{tname}"] = {"desc": desc, "columns": cols.to_dict()}
                progress.progress((i + 1) / max(len(target_schemas), 1))

            st.session_state.catalog = catalog

            # Step 6: build workbook(s) + zip
            files = build_workbooks(catalog)
            st.session_state["zip_buf"] = zip_workbooks(files).getvalue()

            # Step 7: optional RAG store
            if rag_toggle:
                ok, msg = store_to_chroma(catalog)
                (st.success if ok else st.error)(msg)

            # clear intermediate cache after zip prepared
            try:
                CACHE.clear()
            except Exception:
                pass

            st.success("Data dictionary generated. Download it from the Output panel.")
        except Exception as e:
            st.error(f"Generation failed: {e}")

    # ---------------- MIDDLE: Viewer canvas ----------------
    with mid:
        st.markdown("### Viewer")
        catalog = st.session_state.catalog
        if catalog:
            schema_pick = st.selectbox("Schema", list(catalog.keys()))
            table_pick = st.selectbox("Table", list(catalog[schema_pick].keys()))
            meta = catalog[schema_pick][table_pick]
            st.caption(meta["desc"])
            st.dataframe(meta["columns"], use_container_width=True, height=420)
        else:
            st.info("Run generation to preview tables and columns here.")

    # ---------------- RIGHT: Observability / Chat ----------------
    with right:
        st.markdown("### Data Observability & Chat")
        with st.expander("📈 Quick stats", expanded=True):
            if st.session_state.catalog:
                n_schemas = len(st.session_state.catalog)
                n_tables = sum(len(t) for t in st.session_state.catalog.values())
                n_cols = sum(len(m["columns"]) for t in st.session_state.catalog.values() for m in t.values())
                st.metric("Schemas", n_schemas)
                st.metric("Tables", n_tables)
                st.metric("Columns", n_cols)
            else:
                st.write("No metadata generated yet.")

        st.markdown("#### 💬 Ask about your data")
        q = st.chat_input("Ask a question about the generated dictionary...")
        for role, msg in st.session_state.chat_history:
            with st.chat_message(role):
                st.write(msg)
        if q:
            st.session_state.chat_history.append(("user", q))
            answer = rag_query(q)
            st.session_state.chat_history.append(("assistant", answer))
            st.rerun()
