import os
import secrets
from typing import Optional
from urllib.parse import urlparse, unquote

import pymysql

from fastapi import (
    FastAPI,
    Request,
    Depends,
    Form,
    HTTPException
)

from fastapi.security import (
    HTTPBasic,
    HTTPBasicCredentials
)

from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates


# =========================================================
# FastAPI
# =========================================================

app = FastAPI(
    title="橘貓代購後台"
)


# =========================================================
# Static / Templates
# =========================================================

app.mount(
    "/static",
    StaticFiles(directory="static"),
    name="static"
)

templates = Jinja2Templates(
    directory="templates"
)


# =========================================================
# Admin Login
# =========================================================

security = HTTPBasic()


def verify_admin(
    credentials: HTTPBasicCredentials = Depends(security)
):

    admin_username = os.environ.get(
        "ADMIN_USERNAME",
        ""
    )

    admin_password = os.environ.get(
        "ADMIN_PASSWORD",
        ""
    )

    # Railway 尚未設定帳密
    if not admin_username or not admin_password:

        raise HTTPException(
            status_code=500,
            detail="後台帳號密碼尚未設定"
        )


    correct_username = secrets.compare_digest(
        credentials.username,
        admin_username
    )

    correct_password = secrets.compare_digest(
        credentials.password,
        admin_password
    )


    if not (
        correct_username
        and correct_password
    ):

        raise HTTPException(
            status_code=401,
            detail="帳號或密碼錯誤",
            headers={
                "WWW-Authenticate": "Basic"
            }
        )


    return credentials.username


# =========================================================
# Database
# =========================================================

def get_db():

    database_url = os.environ["DATABASE_URL"]

    url = urlparse(
        database_url
    )


    return pymysql.connect(

        host=url.hostname,

        port=url.port or 3306,

        user=unquote(
            url.username or ""
        ),

        password=unquote(
            url.password or ""
        ),

        database=url.path.lstrip("/"),

        charset="utf8mb4",

        cursorclass=pymysql.cursors.DictCursor,

        connect_timeout=10,

        read_timeout=15,

        write_timeout=15
    )


# =========================================================
# 共用：抓取單筆訂單
# =========================================================

def fetch_order(
    order_id: int
):

    conn = get_db()


    try:

        with conn.cursor() as cursor:

            cursor.execute(
                """
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
                    remarks,
                    service_fee,
                    early_return,
                    is_early_returned,
                    reconcile_enabled,
                    exchange_rate,
                    member_level_snapshot,
                    original_service_fee,
                    vip_discount_rate,
                    final_service_fee,
                    extra_discount,
                    order_status,
                    cancel_note
                FROM orders
                WHERE order_id = %s
                LIMIT 1
                """,
                (
                    order_id,
                )
            )


            order = cursor.fetchone()


        return order


    finally:

        conn.close()


# =========================================================
# Dashboard
# =========================================================

@app.get("/")
async def home(
    request: Request,
    admin: str = Depends(
        verify_admin
    )
):

    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={}
    )


# =========================================================
# Orders Page
# =========================================================

@app.get("/orders")
async def orders_page(
    request: Request,
    admin: str = Depends(
        verify_admin
    )
):

    return templates.TemplateResponse(
        request=request,
        name="orders.html",
        context={}
    )


# =========================================================
# Search Orders
# =========================================================

@app.get("/orders/search")
async def search_orders(

    request: Request,

    q: str = "",

    unreturned: str = "",

    admin: str = Depends(
        verify_admin
    )
):

    q = q.strip()


    conditions = []

    params = []


    # -----------------------------------------------------
    # 關鍵字
    # -----------------------------------------------------

    if q:

        conditions.append(
            """
            (
                customer_name LIKE %s
                OR tracking_number LIKE %s
                OR CAST(order_id AS CHAR) LIKE %s
            )
            """
        )


        keyword = f"%{q}%"


        params.extend(
            [
                keyword,
                keyword,
                keyword
            ]
        )


    # -----------------------------------------------------
    # 只看未運回
    # -----------------------------------------------------

    if unreturned == "1":

        conditions.append(
            "COALESCE(is_returned, 0) = 0"
        )

        conditions.append(
            "COALESCE(order_status, '正常') <> '取消'"
        )


    # -----------------------------------------------------
    # 沒有任何條件
    # -----------------------------------------------------

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


    where_sql = " AND ".join(
        conditions
    )


    conn = get_db()


    try:

        with conn.cursor() as cursor:


            # =============================================
            # 統計
            # =============================================

            stats_sql = f"""
                SELECT
                    COUNT(*) AS total_count,
                    COALESCE(
                        SUM(weight_kg),
                        0
                    ) AS total_weight

                FROM orders

                WHERE
                    {where_sql}
            """


            cursor.execute(
                stats_sql,
                params
            )


            stats = cursor.fetchone()


            total_count = (
                stats["total_count"]
                if stats
                else 0
            ) or 0


            total_weight = float(
                (
                    stats["total_weight"]
                    if stats
                    else 0
                )
                or 0
            )


            # =============================================
            # 訂單列表
            # =============================================

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

                WHERE
                    {where_sql}

                ORDER BY
                    order_id DESC

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


# =========================================================
# Order Detail
# =========================================================

@app.get(
    "/orders/{order_id}/detail"
)
async def order_detail(

    request: Request,

    order_id: int,

    admin: str = Depends(
        verify_admin
    )
):

    order = fetch_order(
        order_id
    )


    if not order:

        raise HTTPException(
            status_code=404,
            detail="找不到訂單"
        )


    return templates.TemplateResponse(
        request=request,
        name="order_detail.html",
        context={
            "order": order
        }
    )


# =========================================================
# Edit Order Page
# =========================================================

@app.get(
    "/orders/{order_id}/edit"
)
async def edit_order_page(

    request: Request,

    order_id: int,

    admin: str = Depends(
        verify_admin
    )
):

    order = fetch_order(
        order_id
    )


    if not order:

        raise HTTPException(
            status_code=404,
            detail="找不到訂單"
        )


    return templates.TemplateResponse(
        request=request,
        name="order_edit.html",
        context={
            "order": order
        }
    )


# =========================================================
# Save Order
# =========================================================

@app.post(
    "/orders/{order_id}/edit"
)
async def update_order(

    request: Request,

    order_id: int,

    weight_kg: str = Form(""),

    remarks: str = Form(""),

    is_arrived: Optional[str] = Form(
        None
    ),

    is_returned: Optional[str] = Form(
        None
    ),

    admin: str = Depends(
        verify_admin
    )
):

    # -----------------------------------------------------
    # 確認訂單存在
    # -----------------------------------------------------

    old_order = fetch_order(
        order_id
    )


    if not old_order:

        raise HTTPException(
            status_code=404,
            detail="找不到訂單"
        )


    # -----------------------------------------------------
    # 重量
    # -----------------------------------------------------

    weight_kg = weight_kg.strip()


    if weight_kg == "":

        weight_value = None


    else:

        try:

            weight_value = float(
                weight_kg
            )


        except ValueError:

            raise HTTPException(
                status_code=400,
                detail="重量格式錯誤"
            )


        if weight_value < 0:

            raise HTTPException(
                status_code=400,
                detail="重量不能小於 0"
            )


    # -----------------------------------------------------
    # Checkbox
    # -----------------------------------------------------

    arrived_value = (
        1
        if is_arrived == "1"
        else 0
    )


    returned_value = (
        1
        if is_returned == "1"
        else 0
    )


    # -----------------------------------------------------
    # Update
    # -----------------------------------------------------

    conn = get_db()


    try:

        with conn.cursor() as cursor:

            cursor.execute(
                """
                UPDATE orders

                SET
                    weight_kg = %s,
                    remarks = %s,
                    is_arrived = %s,
                    is_returned = %s

                WHERE
                    order_id = %s
                """,
                (
                    weight_value,
                    remarks.strip(),
                    arrived_value,
                    returned_value,
                    order_id
                )
            )


        conn.commit()


    except Exception:

        conn.rollback()

        raise


    finally:

        conn.close()


    # -----------------------------------------------------
    # 重新讀取最新資料
    # -----------------------------------------------------

    order = fetch_order(
        order_id
    )


    return templates.TemplateResponse(
        request=request,
        name="order_detail.html",
        context={
            "order": order
        }
    )
