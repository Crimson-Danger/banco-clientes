from __future__ import annotations

import os
import sys

import psycopg


def env(name: str, default: str = "") -> str:
    return os.getenv(name, default).strip()


def main() -> int:
    host = env("PGHOST")
    port = env("PGPORT", "5432")
    admin_user = env("PGADMINUSER", env("PGUSER", "postgres"))
    admin_password = env("PGADMINPASSWORD", env("PGPASSWORD"))
    database_name = env("PGDATABASE")

    if not host or not database_name:
        print("Defina PGHOST e PGDATABASE antes de rodar este script.")
        return 1

    admin_conninfo = f"host={host} port={port} user={admin_user} dbname=postgres"
    if admin_password:
        admin_conninfo += f" password={admin_password}"

    try:
        with psycopg.connect(admin_conninfo, autocommit=True) as connection:
            with connection.cursor() as cursor:
                cursor.execute("SELECT 1 FROM pg_database WHERE datname = %s", (database_name,))
                exists = cursor.fetchone() is not None
                if exists:
                    print(f"Banco '{database_name}' ja existe.")
                else:
                    cursor.execute(f'CREATE DATABASE "{database_name}"')
                    print(f"Banco '{database_name}' criado com sucesso.")
    except Exception as exc:
        print(f"Falha ao criar/verificar banco PostgreSQL: {exc}")
        return 1

    print("Pronto. Agora configure DATABASE_URL ou PGHOST/PGPORT/PGUSER/PGPASSWORD/PGDATABASE e suba o app.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
