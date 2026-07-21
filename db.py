import os
import json
from hdbcli import dbapi


def load_env_from_dotenv():
    path = os.path.join(os.getcwd(), ".env")
    if not os.path.exists(path):
        return
    try:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                s = line.strip()
                if not s or s.startswith("#"):
                    continue
                if "=" in s:
                    k, v = s.split("=", 1)
                    k = k.strip()
                    v = v.strip().strip('"').strip("'")
                    if not os.getenv(k):
                        os.environ[k] = v
    except Exception:
        pass


def get_hana_credentials():
    if not (
        os.getenv("HANA_HOST") and os.getenv("HANA_USER") and os.getenv("HANA_PASSWORD")
    ):
        vcap = os.getenv("VCAP_SERVICES")
        if vcap:
            try:
                data = json.loads(vcap)
                creds = None
                for _, services in data.items():
                    for s in services:
                        c = s.get("credentials", {})
                        if (
                            c.get("host")
                            and (c.get("user") or c.get("username"))
                            and c.get("password")
                        ):
                            creds = c
                            break
                    if creds:
                        break
                if creds:
                    os.environ.setdefault("HANA_HOST", str(creds.get("host")))
                    port_val = creds.get("port") or creds.get("port_tls")
                    if port_val is not None:
                        os.environ.setdefault("HANA_PORT", str(port_val))
                    os.environ.setdefault(
                        "HANA_USER", str(creds.get("user") or creds.get("username"))
                    )
                    os.environ.setdefault("HANA_PASSWORD", str(creds.get("password")))
                    if creds.get("schema"):
                        os.environ.setdefault("HANA_SCHEMA", str(creds.get("schema")))
            except Exception:
                pass
    return {
        "host": os.getenv("HANA_HOST"),
        "port": int(os.getenv("HANA_PORT")) if os.getenv("HANA_PORT") else None,
        "user": os.getenv("HANA_USER"),
        "password": os.getenv("HANA_PASSWORD"),
        "schema": os.getenv("HANA_SCHEMA"),
    }


def get_hana_connection():
    c = get_hana_credentials()
    missing = []
    for key in ["host", "user", "password", "schema"]:
        if not c.get(key):
            missing.append(key)
    if missing:
        raise ValueError(
            "Faltan variables de entorno para HANA: "
            + ", ".join(
                [
                    {
                        "host": "HANA_HOST",
                        "user": "HANA_USER",
                        "password": "HANA_PASSWORD",
                        "schema": "HANA_SCHEMA",
                    }[m]
                    for m in missing
                ]
            )
        )
    conn = dbapi.connect(
        address=c.get("host"),
        port=c.get("port") or 443,
        user=c.get("user"),
        password=c.get("password"),
        encrypt=True,
        sslValidateCertificate=False,
    )
    schema = c.get("schema")
    if schema:
        cur = conn.cursor()
        cur.execute(f'SET SCHEMA "{schema}"')
        cur.close()
    return conn


def _columns():
    return [
        "RPU",
        "HEADERPERIOD",
        "PERIOD",
        "MOVTYPE",
        "CHECKDIG",
        "TOTALAMOUNT",
        "ENERGYAMOUNT",
        "IVA",
        "REASON",
        "AMOUNT",
        "CONSUMPTION",
        "DEMAND",
        "REACTIVEPOWER",
        "CUSTOMERNAME",
        "ADDRESS",
        "CITY",
        "STATE",
        "RFC",
        "NEIGHBORHOOD",
        "STREET1",
        "STREET2",
        "COMPANYID",
        "IVATYPE",
        "METER",
        "ACCOUNT",
        "CONTRACTEDDEMAND",
    ]


def _table_fqn():
    schema = os.getenv("HANA_SCHEMA")
    if schema:
        return f'"{schema}"."TELCEL_EE_TEMPFLATFILECFE"'
    return '"TELCEL_EE_TEMPFLATFILECFE"'


def truncate_archivoplano(conn):
    cur = conn.cursor()
    try:
        print(f"Truncando tabla {_table_fqn()}")
        cur.execute(f"TRUNCATE TABLE {_table_fqn()}")
        conn.commit()
        print("Truncate exitoso")
    except Exception as e:
        print(f"TRUNCATE falló, usando DELETE: {e}")
        cur.execute(f"delete from {_table_fqn()}")
        conn.commit()
        print("Delete ejecutado")
    finally:
        cur.close()


def insert_archivoplano(conn, entities):
    """
    Inserta en batch usando executemany.
    Optimizado para volúmenes grandes: un solo viaje por batch al servidor HANA.
    Solo hace fallback fila-por-fila si el batch completo falla.
    """
    cols = _columns()
    qmarks = ",".join(["?"] * len(cols))
    sql = (
        "insert into "
        + _table_fqn()
        + " ("
        + ",".join([f'"{c}"' for c in cols])
        + ")"
        + f" values ({qmarks})"
    )
    params = [[e.get(c) for c in cols] for e in entities]
    cur = conn.cursor()
    inserted = 0
    failed = 0
    errors = []
    try:
        cur.executemany(sql, params)
        conn.commit()
        inserted = len(entities)
        return {"inserted": inserted, "failed": 0, "errors": []}
    except Exception as e:
        print(f"Error en batch, intentando fila por fila: {e}")
        conn.rollback()
        for idx, p in enumerate(params):
            try:
                cur.execute(sql, p)
                inserted += 1
            except Exception as ie:
                failed += 1
                rpu_val = p[0]  # Rpu siempre es el primer campo
                errors.append({"index": idx, "rpu": rpu_val, "message": str(ie)})
        conn.commit()
        print(f"Fallback completado: ok={inserted}, errores={failed}")
        return {"inserted": inserted, "failed": failed, "errors": errors}
    finally:
        cur.close()


def upsert_archivoplano(conn, entities):
    cols = _columns()
    key = "RPU"
    update_cols = [c for c in cols if c != key]
    update_sql = (
        "update "
        + _table_fqn()
        + " set "
        + ",".join([f'"{c}"=?' for c in update_cols])
        + f' where "{key}"=?'
    )
    insert_qmarks = ",".join(["?"] * len(cols))
    insert_sql = (
        "insert into "
        + _table_fqn()
        + " ("
        + ",".join([f'"{c}"' for c in cols])
        + f") values ({insert_qmarks})"
    )
    cur = conn.cursor()
    updated = 0
    inserted = 0
    failed = 0
    errors = []
    try:
        print(f"Upsert de {len(entities)} registros en {_table_fqn()}")
        to_insert = []
        idx_map = []
        for idx, e in enumerate(entities):
            try:
                upd_params = [e.get(c) for c in update_cols] + [e.get(key)]
                cur.execute(update_sql, upd_params)
                if cur.rowcount and cur.rowcount > 0:
                    updated += 1
                else:
                    to_insert.append([e.get(c) for c in cols])
                    idx_map.append(idx)
            except Exception as ue:
                failed += 1
                rpu_val = e.get(key)
                errors.append(
                    {
                        "index": idx,
                        "rpu": rpu_val,
                        "message": str(ue),
                        "operation": "update",
                    }
                )
        if to_insert:
            try:
                cur.executemany(insert_sql, to_insert)
                inserted += len(to_insert)
            except Exception as be:
                print(f"Error en inserción batch dentro de upsert: {be}")
                for pos, p in enumerate(to_insert):
                    idx = idx_map[pos]
                    try:
                        cur.execute(insert_sql, p)
                        inserted += 1
                    except Exception as ie:
                        failed += 1
                        rpu_idx = cols.index("RPU") if "RPU" in cols else None
                        rpu_val = p[rpu_idx] if rpu_idx is not None else None
                        errors.append(
                            {
                                "index": idx,
                                "rpu": rpu_val,
                                "message": str(ie),
                                "operation": "insert",
                            }
                        )
        conn.commit()
        print(
            f"Upsert resumen: updated={updated}, inserted={inserted}, errores={failed}"
        )
        return {
            "updated": updated,
            "inserted": inserted,
            "failed": failed,
            "errors": errors,
            "update_sql": update_sql,
            "insert_sql": insert_sql,
        }
    except Exception as e:
        print(f"Fallo general en upsert: {e}")
        return {
            "updated": updated,
            "inserted": inserted,
            "failed": failed,
            "errors": errors or [{"index": None, "rpu": None, "message": str(e)}],
        }
    finally:
        cur.close()
