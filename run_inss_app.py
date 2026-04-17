from inss_db_app.config import settings
from inss_db_app.main import app

if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host=settings.host, port=settings.port)
