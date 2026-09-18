import os
import secrets
from decimal import Decimal, InvalidOperation
from datetime import datetime
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
# 單筆訂單
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
                (order_id,)
            )

            return cursor.fetchone()

    finally:

        conn.close()


# =========================================================
# 客戶姓名建議
# =========================================================

def fetch_customer_names():

    conn = get_db()

    try:

        with conn.cursor() as cursor:

            cursor.execute(
                """
                SELECT DISTINCT customer_name

                FROM orders

                WHERE
                    customer_name IS NOT NULL
                    AND customer_name <> ''

                ORDER BY customer_name
                """
            )

            rows = cursor.fetchall()

            return [
                row["customer_name"]
                for row in rows
            ]

    finally:

        conn.close()


# =========================================================
# 平台建議
# =========================================================

def fetch_platforms():

    conn = get_db()

    try:

        with conn.cursor() as cursor:

            cursor.execute(
                """
                SELECT DISTINCT platform

                FROM orders

                WHERE
                    platform IS NOT NULL
                    AND platform <> ''

                ORDER BY platform
                """
            )

            rows = cursor.fetchall()

            return [
                row["platform"]
                for row in rows
            ]

    finally:

        conn.close()


# =========================================================
# 搜尋共用函式
# =========================================================

def query_orders(
    q: str = "",
    unreturned: str = ""
):

    q = q.strip()

    conditions = []

    params = []


    # 關鍵字
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


    # 未運回
    if unreturned == "1":

        conditions.append(
            "COALESCE(is_returned, 0) = 0"
        )

        conditions.append(
            "COALESCE(order_status, '正常') <> '取消'"
        )


    if not conditions:

        return {
            "orders": [],
            "total_count": 0,
            "total_weight": 0
        }


    where_sql = " AND ".join(
        conditions
    )

    conn = get_db()

    try:

        with conn.cursor() as cursor:

            # 統計
            cursor.execute(
                f"""
                SELECT
                    COUNT(*) AS total_count,

                    COALESCE(
                        SUM(weight_kg),
                        0
                    ) AS total_weight

                FROM orders

                WHERE {where_sql}
                """,
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
                ) or 0
            )


            # 列表
            cursor.execute(
                f"""
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
                """,
                params
            )

            orders = cursor.fetchall()


        return {
            "orders": orders,
            "total_count": total_count,
            "total_weight": total_weight
        }

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

    result = query_orders(
        q,
        unreturned
    )

    return templates.TemplateResponse(
        request=request,
        name="order_results.html",
        context={
            **result,
            "q": q,
            "unreturned": unreturned
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
    q: str = "",
    unreturned: str = "",
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
            "order": order,
            "q": q,
            "unreturned": unreturned
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
    q: str = "",
    unreturned: str = "",
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

    customer_names = fetch_customer_names()

    platforms = fetch_platforms()

    return templates.TemplateResponse(
        request=request,
        name="order_edit.html",
        context={
            "order": order,
            "customer_names": customer_names,
            "platforms": platforms,
            "q": q,
            "unreturned": unreturned
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

    order_time: str = Form(""),
    customer_name: str = Form(""),
    platform: str = Form(""),
    tracking_number: str = Form(""),

    amount_rmb: str = Form(""),
    weight_kg: str = Form(""),

    remarks: str = Form(""),

    is_arrived: Optional[str] = Form(None),
    is_returned: Optional[str] = Form(None),

    q: str = Form(""),
    unreturned: str = Form(""),

    admin: str = Depends(
        verify_admin
    )
):

    old_order = fetch_order(
        order_id
    )

    if not old_order:

        raise HTTPException(
            status_code=404,
            detail="找不到訂單"
        )


    # =====================================================
    # 姓名
    # =====================================================

    customer_name = customer_name.strip()

    if not customer_name:

        raise HTTPException(
            status_code=400,
            detail="客戶姓名不能空白"
        )


    # =====================================================
    # 日期
    # =====================================================

    order_time = order_time.strip()

    if order_time:

        try:

            datetime.strptime(
                order_time,
                "%Y-%m-%d"
            )

        except ValueError:

            raise HTTPException(
                status_code=400,
                detail="日期格式錯誤"
            )

        order_time_value = order_time

    else:

        order_time_value = None


    # =====================================================
    # 人民幣金額
    # =====================================================

    amount_rmb = amount_rmb.strip()

    if amount_rmb == "":

        amount_value = None

    else:

        try:

            amount_value = Decimal(
                amount_rmb
            )

        except InvalidOperation:

            raise HTTPException(
                status_code=400,
                detail="人民幣金額格式錯誤"
            )

        if amount_value < 0:

            raise HTTPException(
                status_code=400,
                detail="人民幣金額不能小於 0"
            )


    # =====================================================
    # 重量
    # =====================================================

    weight_kg = weight_kg.strip()

    if weight_kg == "":

        weight_value = None

    else:

        try:

            weight_value = Decimal(
                weight_kg
            )

        except InvalidOperation:

            raise HTTPException(
                status_code=400,
                detail="重量格式錯誤"
            )

        if weight_value < 0:

            raise HTTPException(
                status_code=400,
                detail="重量不能小於 0"
            )


    # =====================================================
    # Checkbox
    # =====================================================

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


    # =====================================================
    # 文字欄位
    # =====================================================

    platform_value = (
        platform.strip()
        or None
    )

    tracking_value = (
        tracking_number.strip()
        or None
    )

    remarks_value = (
        remarks.strip()
        or None
    )


    # =====================================================
    # Update
    # =====================================================

    conn = get_db()

    try:

        with conn.cursor() as cursor:

            cursor.execute(
                """
                UPDATE orders

                SET
                    order_time = %s,
                    customer_name = %s,
                    platform = %s,
                    tracking_number = %s,
                    amount_rmb = %s,
                    weight_kg = %s,
                    remarks = %s,
                    is_arrived = %s,
                    is_returned = %s

                WHERE
                    order_id = %s
                """,
                (
                    order_time_value,
                    customer_name,
                    platform_value,
                    tracking_value,
                    amount_value,
                    weight_value,
                    remarks_value,
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


    # 編輯後直接刷新目前搜尋結果
    result = query_orders(
        q,
        unreturned
    )

    return templates.TemplateResponse(
        request=request,
        name="order_results.html",
        context={
            **result,
            "q": q,
            "unreturned": unreturned
        }
    )


# =========================================================
# Delete Order
# =========================================================

@app.post(
    "/orders/{order_id}/delete"
)
async def delete_order(
    request: Request,
    order_id: int,

    q: str = Form(""),
    unreturned: str = Form(""),

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


    conn = get_db()

    try:

        with conn.cursor() as cursor:

            cursor.execute(
                """
                DELETE FROM orders
                WHERE order_id = %s
                """,
                (order_id,)
            )

        conn.commit()

    except pymysql.err.IntegrityError:

        conn.rollback()

        raise HTTPException(
            status_code=409,
            detail="這筆訂單仍被其他資料引用，目前無法刪除"
        )

    except Exception:

        conn.rollback()

        raise

    finally:

        conn.close()


    result = query_orders(
        q,
        unreturned
    )

    return templates.TemplateResponse(
        request=request,
        name="order_results.html",
        context={
            **result,
            "q": q,
            "unreturned": unreturned
        }
    )
