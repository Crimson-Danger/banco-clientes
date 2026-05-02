from inss_db_app.config import settings
from inss_db_app.main import app

if __name__ == "__main__":
    import uvicorn

    print(
        "Startup INSS app:",
        f"port={settings.port}",
        f"cep_enrichment={settings.enable_cep_enrichment}",
        f"ddd_enrichment={settings.enable_ddd_enrichment}",
        f"startup_schema_maintenance={settings.enable_startup_schema_maintenance}",
        f"startup_address_backfill={settings.enable_startup_address_backfill}",
    )
    uvicorn.run(app, host=settings.host, port=settings.port)
