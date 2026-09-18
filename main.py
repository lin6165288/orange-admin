import os
import pymysql
from urllib.parse import urlparse, unquote
from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates


def get_db():
    database_url = os.environ["DATABASE_URL"]
    url = urlparse(database_url)

    return pymysql.connect(
        host=url.hostname,
        port=url.port or 3306,
        user=unquote(url.username),
        password=unquote(url.password),
        database=url.path.lstrip("/"),
        charset="utf8mb4",
        cursorclass=pymysql.cursors.DictCursor
    )

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



@app.get("/db-tables")
def db_tables():
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
            cursor.execute("SHOW TABLES;")
            tables = [row[0] for row in cursor.fetchall()]

        conn.close()

        return {
            "status": "success",
            "tables": tables
        }

    except Exception as e:
        return {
            "status": "error",
            "message": str(e)
        }




@app.get("/db-orders-columns")
def db_orders_columns():
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
            cursor.execute("SHOW COLUMNS FROM orders;")
            columns = cursor.fetchall()

        conn.close()

        result = []

        for col in columns:
            result.append({
                "field": col[0],
                "type": col[1],
                "null": col[2],
                "key": col[3],
                "default": col[4],
                "extra": col[5]
            })

        return {
            "status": "success",
            "columns": result
        }

    except Exception as e:
        return {
            "status": "error",
            "message": str(e)
        }




@app.get("/orders")
async def orders_page(request: Request):

    return templates.TemplateResponse(
        request=request,
        name="orders.html",
        context={}
    )



@app.get("/orders/search")
async def search_orders(
    request: Request,
    q: str = ""
):

    q = q.strip()

    if not q:
        return templates.TemplateResponse(
            request=request,
            name="order_results.html",
            context={
                "orders": []
            }
        )

    conn = get_db()

    try:

        with conn.cursor() as cursor:

            sql = """
                SELECT
                    order_id,
                    order_time,
                    customer_name,
                    platform,
                    tracking_number,
                    amount_rmb,
                    weight_kg,
                    is_arrived,
                    is_returned,
                    order_status
                FROM orders
                WHERE
                    customer_name LIKE %s
                    OR tracking_number LIKE %s
                    OR CAST(order_id AS CHAR) LIKE %s
                ORDER BY order_id DESC
                LIMIT 200
            """

            keyword = f"%{q}%"

            cursor.execute(
                sql,
                (
                    keyword,
                    keyword,
                    keyword
                )
            )

            orders = cursor.fetchall()

    finally:
        conn.close()

    return templates.TemplateResponse(
        request=request,
        name="order_results.html",
        context={
            "orders": orders
        }
    )
