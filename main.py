import os
import pymysql
from urllib.parse import urlparse, unquote
from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates


app = FastAPI(
    title="橘貓代購後台"
)

app.mount(
    "/static",
    StaticFiles(directory="static"),
    name="static"
)

templates = Jinja2Templates(
    directory="templates"
)


@app.get("/")
async def home(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={}
    )


@app.get("/db-test")
def db_test():
    try:
        database_url = os.environ["DATABASE_URL"]
        url = urlparse(database_url)

        conn = pymysql.connect(
            host=url.hostname,
            port=url.port or 3306,
            user=unquote(url.username),
            password=unquote(url.password),
            database=url.path.lstrip("/"),
            charset="utf8mb4"
        )

        with conn.cursor() as cursor:
            cursor.execute("SELECT VERSION();")
            version = cursor.fetchone()[0]

        conn.close()

        return {
            "status": "success",
            "message": "MySQL 資料庫連線成功",
            "database": version
        }

    except Exception as e:
        return {
            "status": "error",
            "message": str(e)
        }
