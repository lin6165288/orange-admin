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
    q: str = "",
    unreturned: str = ""
):

    q = q.strip()

    conditions = []
    params = []

    # 關鍵字搜尋
    if q:
        conditions.append("""
            (
                customer_name LIKE %s
                OR tracking_number LIKE %s
                OR CAST(order_id AS CHAR) LIKE %s
            )
        """)

        keyword = f"%{q}%"

        params.extend([
            keyword,
            keyword,
            keyword
        ])

    # 只看未運回
    if unreturned == "1":

        conditions.append(
            "COALESCE(is_returned, 0) = 0"
        )

        conditions.append(
            "COALESCE(order_status, '正常') <> '取消'"
        )

    # 沒輸入任何條件
    if not conditions:

        return templates.TemplateResponse(
            request=request,
            name="order_results.html",
            context={
                "orders": [],
                "total_count": 0,
                "total_weight": 0
            }
        )

    where_sql = " AND ".join(conditions)

    conn = get_db()

    try:

        with conn.cursor() as cursor:

            # 統計
            stats_sql = f"""
                SELECT
                    COUNT(*) AS total_count,
                    COALESCE(SUM(weight_kg), 0) AS total_weight
                FROM orders
                WHERE {where_sql}
            """

            cursor.execute(
                stats_sql,
                params
            )

            stats = cursor.fetchone()

            total_count = stats["total_count"] or 0
            total_weight = float(
                stats["total_weight"] or 0
            )

            # 訂單列表
            sql = f"""
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
                WHERE {where_sql}
                ORDER BY order_id DESC
                LIMIT 300
            """

            cursor.execute(
                sql,
                params
            )

            orders = cursor.fetchall()

    finally:
        conn.close()

    return templates.TemplateResponse(
        request=request,
        name="order_results.html",
        context={
            "orders": orders,
            "total_count": total_count,
            "total_weight": total_weight
        }
    )
