import os
import json
import secrets
import threading
import logging
import re

from decimal import Decimal, InvalidOperation, ROUND_CEILING, ROUND_HALF_EVEN, ROUND_CEILING
from datetime import datetime
from zoneinfo import ZoneInfo
from typing import Optional
from urllib.parse import urlparse, unquote, quote

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
from fastapi.responses import RedirectResponse, Response
from shipping_excel import make_sellnow_xlsm, make_shipping_detail_xlsx
from profit_excel import make_profit_xlsx


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

    database_url = os.environ[
        "DATABASE_URL"
    ]

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
# Fetch Single Order
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
# Customer Suggestions
# =========================================================

def fetch_locked_order(cursor, order_id: int):
    # SELECT * keeps new order columns in future audit snapshots as well.
    cursor.execute(
        "SELECT * FROM orders WHERE order_id = %s FOR UPDATE",
        (order_id,)
    )
    return cursor.fetchone()


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
# Platform Suggestions
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
# Search Orders
# =========================================================

def query_orders(
    search_order_id="",
    search_customer_name="",
    search_tracking_number="",
    search_amount_rmb="",
    search_platform="",
    search_order_date="",
    search_arrived_status="",
    search_returned_status="",
    unreturned=""
):
    """訂單管理查詢。

    - 沒有任何篩選時：直接顯示最近 100 筆
    - 有篩選時：最多顯示 300 筆
    - unreturned 保留舊版相容性
    """

    conditions = []
    params = []
    search_error = ""

    search_order_id = str(search_order_id or "").strip()
    search_customer_name = str(search_customer_name or "").strip()
    search_tracking_number = str(search_tracking_number or "").strip()
    search_amount_rmb = str(search_amount_rmb or "").strip()
    search_platform = str(search_platform or "").strip()
    search_order_date = str(search_order_date or "").strip()
    search_arrived_status = str(search_arrived_status or "").strip()
    search_returned_status = str(search_returned_status or "").strip()
    unreturned = str(unreturned or "").strip()

    if search_order_id:
        try:
            order_id_value = int(search_order_id)
        except ValueError:
            search_error = "訂單編號只能輸入整數。"
        else:
            conditions.append("order_id = %s")
            params.append(order_id_value)

    if search_customer_name:
        conditions.append("customer_name LIKE %s")
        params.append(f"%{search_customer_name}%")

    if search_tracking_number:
        conditions.append("tracking_number LIKE %s")
        params.append(f"%{search_tracking_number}%")

    if search_amount_rmb and not search_error:
        try:
            amount_value = Decimal(search_amount_rmb)
        except InvalidOperation:
            search_error = "人民幣金額只能輸入數字。"
        else:
            conditions.append("amount_rmb = %s")
            params.append(amount_value)

    if search_platform:
        conditions.append("platform LIKE %s")
        params.append(f"%{search_platform}%")

    if search_order_date:
        try:
            datetime.strptime(search_order_date, "%Y-%m-%d")
        except ValueError:
            search_error = "下單日期格式錯誤。"
        else:
            conditions.append("DATE(order_time) = %s")
            params.append(search_order_date)

    if search_arrived_status == "arrived":
        conditions.append("COALESCE(is_arrived, 0) = 1")
    elif search_arrived_status == "unarrived":
        conditions.append("COALESCE(is_arrived, 0) = 0")

    # 新版三態篩選；同時保留舊 unreturned=1 相容
    if search_returned_status == "returned":
        conditions.append("COALESCE(is_returned, 0) = 1")
    elif search_returned_status == "unreturned":
        conditions.append("COALESCE(is_returned, 0) = 0")
    elif not search_returned_status:
        if unreturned == "1":
            conditions.append("COALESCE(is_returned, 0) = 0")
            conditions.append("COALESCE(order_status, '正常') <> '取消'")
        elif unreturned == "0":
            conditions.append("COALESCE(is_returned, 0) = 1")

    if search_error:
        return {
            "orders": [],
            "total_count": 0,
            "total_weight": 0.0,
            "showing_count": 0,
            "is_default_view": False,
            "search_error": search_error,
        }

    filtered = bool(conditions)
    where_sql = " AND ".join(conditions) if conditions else "1=1"
    limit_value = 300 if filtered else 100

    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT
                    COUNT(*) AS total_count,
                    COALESCE(SUM(COALESCE(weight_kg, 0)), 0) AS total_weight
                FROM orders
                WHERE {where_sql}
                """,
                params,
            )
            stats = cursor.fetchone() or {}

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
                    is_early_returned,
                    order_status
                FROM orders
                WHERE {where_sql}
                ORDER BY order_id DESC
                LIMIT {limit_value}
                """,
                params,
            )
            orders = cursor.fetchall()

        return {
            "orders": orders,
            "total_count": int(stats.get("total_count") or 0),
            "total_weight": float(stats.get("total_weight") or 0),
            "showing_count": len(orders),
            "is_default_view": not filtered,
            "search_error": "",
        }
    finally:
        conn.close()


# =========================================================
# Audit Log
# =========================================================

AUDIT_FIELD_LABELS = {

    "order_time":
        "下單日期",

    "customer_name":
        "客戶姓名",

    "platform":
        "平台",

    "tracking_number":
        "物流單號",

    "amount_rmb":
        "人民幣金額",

    "weight_kg":
        "重量",

    "remarks":
        "備註",

    "is_arrived":
        "已到貨",

    "is_returned":
        "已運回",

    "is_early_returned":
        "提前運回",

    "exchange_rate":
        "人民幣匯率",

    "member_level_snapshot":
        "會員等級",

    "original_service_fee":
        "原始手續費",

    "vip_discount_rate":
        "VIP 折扣率",

    "final_service_fee":
        "最終手續費",

    "extra_discount":
        "額外折扣",

    "order_status":
        "訂單狀態",

    "cancel_note":
        "取消原因"
}


def ensure_audit_table(
    cursor
):

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS order_audit_logs (

            id BIGINT UNSIGNED
                NOT NULL
                AUTO_INCREMENT,

            order_id INT
                NOT NULL,

            action VARCHAR(20)
                NOT NULL,

            admin_username VARCHAR(100)
                NOT NULL,

            before_data LONGTEXT
                NULL,

            after_data LONGTEXT
                NULL,

            created_at TIMESTAMP
                NOT NULL
                DEFAULT CURRENT_TIMESTAMP,

            PRIMARY KEY (id),

            INDEX idx_audit_order_id (
                order_id
            ),

            INDEX idx_audit_created_at (
                created_at
            )

        )
        ENGINE=InnoDB
        DEFAULT CHARSET=utf8mb4
        """
    )


# MySQL CREATE TABLE can implicitly commit. Never run it in an order change transaction.
_audit_table_ready = False
_audit_table_lock = threading.Lock()


def ensure_audit_storage():
    global _audit_table_ready
    if _audit_table_ready:
        return
    with _audit_table_lock:
        if _audit_table_ready:
            return
        conn = get_db()
        try:
            with conn.cursor() as cursor:
                ensure_audit_table(cursor)
            conn.commit()
            _audit_table_ready = True
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()


def order_to_json(
    data
):

    if data is None:
        return None

    return json.dumps(
        data,
        ensure_ascii=False,
        default=str
    )


def write_audit_log(

    cursor,

    order_id: int,

    action: str,

    admin_username: str,

    before_data=None,

    after_data=None
):

    cursor.execute(
        """
        INSERT INTO order_audit_logs
        (
            order_id,
            action,
            admin_username,
            before_data,
            after_data
        )

        VALUES
        (
            %s,
            %s,
            %s,
            %s,
            %s
        )
        """,
        (
            order_id,
            action,
            admin_username,
            order_to_json(
                before_data
            ),
            order_to_json(
                after_data
            )
        )
    )


def parse_audit_json(
    value
):

    if not value:
        return {}

    if isinstance(
        value,
        dict
    ):
        return value

    try:

        return json.loads(
            value
        )

    except Exception:

        return {}


def display_audit_value(
    value
):

    if value is None:
        return "—"

    if value == "":
        return "—"

    if value is True:
        return "是"

    if value is False:
        return "否"

    if value == 1:
        return "是"

    if value == 0:
        return "否"

    return str(
        value
    )


# =========================================================
# Dashboard
# =========================================================

def dashboard_stats():
    """首頁即時統計；訂單數字直接讀正式 orders，會員數字直接讀 members。"""
    taipei_now = datetime.now(ZoneInfo("Asia/Taipei"))
    month_start = taipei_now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    if month_start.month == 12:
        next_month = month_start.replace(year=month_start.year + 1, month=1)
    else:
        next_month = month_start.replace(month=month_start.month + 1)

    ready_filter = """
        COALESCE(o.is_arrived, 0) = 1
        AND (
            COALESCE(o.is_early_returned, 0) = 1
            OR NOT EXISTS (
                SELECT 1
                FROM orders pending
                WHERE pending.customer_name = o.customer_name
                  AND COALESCE(pending.is_arrived, 0) = 0
                  AND COALESCE(pending.is_returned, 0) = 0
                  AND COALESCE(pending.order_status, '正常') <> '取消'
            )
        )
    """

    stats = {
        "total_orders": 0,
        "ready_orders": 0,
        "ready_weight": 0.0,
        "month_purchase_orders": 0,
        "total_members": 0,
        "line_bound": 0,
        "binding_rate": 0.0,
        "month_label": f"{taipei_now.month} 月",
    }

    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                SELECT COUNT(*) AS cnt
                FROM orders
                WHERE COALESCE(order_status, '正常') <> '取消'
                """
            )
            stats["total_orders"] = int((cursor.fetchone() or {}).get("cnt") or 0)

            cursor.execute(
                f"""
                SELECT
                    COUNT(*) AS cnt,
                    COALESCE(SUM(COALESCE(o.weight_kg, 0)), 0) AS total_weight
                FROM orders o
                WHERE COALESCE(o.is_returned, 0) = 0
                  AND COALESCE(o.order_status, '正常') <> '取消'
                  AND o.customer_name IS NOT NULL
                  AND TRIM(o.customer_name) <> ''
                  AND ({ready_filter})
                """
            )
            ready_row = cursor.fetchone() or {}
            stats["ready_orders"] = int(ready_row.get("cnt") or 0)
            stats["ready_weight"] = float(ready_row.get("total_weight") or 0)

            cursor.execute(
                """
                SELECT COUNT(*) AS cnt
                FROM orders
                WHERE order_time >= %s
                  AND order_time < %s
                  AND COALESCE(order_status, '正常') <> '取消'
                  AND COALESCE(platform, '') <> '集運'
                """,
                (month_start.date(), next_month.date()),
            )
            stats["month_purchase_orders"] = int((cursor.fetchone() or {}).get("cnt") or 0)

            cursor.execute("SHOW TABLES LIKE 'members'")
            if cursor.fetchone():
                cursor.execute(
                    """
                    SELECT
                        COUNT(*) AS total_members,
                        SUM(
                            CASE
                                WHEN line_user_id IS NOT NULL
                                 AND TRIM(line_user_id) <> ''
                                THEN 1 ELSE 0
                            END
                        ) AS line_bound
                    FROM members
                    """
                )
                member_row = cursor.fetchone() or {}
                stats["total_members"] = int(member_row.get("total_members") or 0)
                stats["line_bound"] = int(member_row.get("line_bound") or 0)
                if stats["total_members"] > 0:
                    stats["binding_rate"] = (
                        stats["line_bound"] / stats["total_members"] * 100.0
                    )

        return stats
    finally:
        conn.close()


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
        context=dashboard_stats()
    )


# =========================================================
# Orders Page
# =========================================================

@app.get("/orders")
async def orders_page(
    request: Request,
    admin: str = Depends(verify_admin)
):
    result = query_orders()

    return templates.TemplateResponse(
        request=request,
        name="orders.html",
        context={
            **result,
            "platforms": fetch_platforms(),
            "search_order_id": "",
            "search_customer_name": "",
            "search_tracking_number": "",
            "search_amount_rmb": "",
            "search_platform": "",
            "search_order_date": "",
            "search_arrived_status": "",
            "search_returned_status": "",
            "unreturned": "",
        },
    )


# =========================================================
# Search Orders
# =========================================================

@app.get("/orders/search")
async def search_orders(
    request: Request,
    search_order_id: str = "",
    search_customer_name: str = "",
    search_tracking_number: str = "",
    search_amount_rmb: str = "",
    search_platform: str = "",
    search_order_date: str = "",
    search_arrived_status: str = "",
    search_returned_status: str = "",
    unreturned: str = "",
    admin: str = Depends(verify_admin),
):
    result = query_orders(
        search_order_id=search_order_id,
        search_customer_name=search_customer_name,
        search_tracking_number=search_tracking_number,
        search_amount_rmb=search_amount_rmb,
        search_platform=search_platform,
        search_order_date=search_order_date,
        search_arrived_status=search_arrived_status,
        search_returned_status=search_returned_status,
        unreturned=unreturned,
    )

    return templates.TemplateResponse(
        request=request,
        name="order_results.html",
        context={
            **result,
            "search_order_id": search_order_id,
            "search_customer_name": search_customer_name,
            "search_tracking_number": search_tracking_number,
            "search_amount_rmb": search_amount_rmb,
            "search_platform": search_platform,
            "search_order_date": search_order_date,
            "search_arrived_status": search_arrived_status,
            "search_returned_status": search_returned_status,
            "unreturned": unreturned,
        },
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

    search_order_id: str = "",

    search_customer_name: str = "",

    search_tracking_number: str = "",

    search_amount_rmb: str = "",

    search_platform: str = "",

    search_order_date: str = "",

    search_arrived_status: str = "",

    search_returned_status: str = "",

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

            "search_order_id":
                search_order_id,

            "search_customer_name":
                search_customer_name,

            "search_tracking_number":
                search_tracking_number,

            "search_amount_rmb":
                search_amount_rmb,

            "search_platform":
                search_platform,

            "search_order_date":
                search_order_date,

            "search_arrived_status":
                search_arrived_status,

            "search_returned_status":
                search_returned_status,

            "unreturned":
                unreturned
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

    search_order_id: str = "",

    search_customer_name: str = "",

    search_tracking_number: str = "",

    search_amount_rmb: str = "",

    search_platform: str = "",

    search_order_date: str = "",

    search_arrived_status: str = "",

    search_returned_status: str = "",

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


    customer_names = (
        fetch_customer_names()
    )

    platforms = (
        fetch_platforms()
    )


    return templates.TemplateResponse(
        request=request,
        name="order_edit.html",
        context={
            "order":
                order,

            "customer_names":
                customer_names,

            "platforms":
                platforms,

            "search_order_id":
                search_order_id,

            "search_customer_name":
                search_customer_name,

            "search_tracking_number":
                search_tracking_number,

            "search_amount_rmb":
                search_amount_rmb,

            "search_platform":
                search_platform,

            "search_order_date":
                search_order_date,

            "search_arrived_status":
                search_arrived_status,

            "search_returned_status":
                search_returned_status,

            "unreturned":
                unreturned
        }
    )


# =========================================================
# Update Order
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

    is_arrived: Optional[str] = Form(
        None
    ),

    is_returned: Optional[str] = Form(
        None
    ),

    search_order_id: str = Form(
        ""
    ),

    search_customer_name: str = Form(
        ""
    ),

    search_tracking_number: str = Form(
        ""
    ),

    search_amount_rmb: str = Form(
        ""
    ),

    search_platform: str = Form(
        ""
    ),

    search_order_date: str = Form(
        ""
    ),

    search_arrived_status: str = Form(
        ""
    ),

    search_returned_status: str = Form(
        ""
    ),

    unreturned: str = Form(
        ""
    ),

    admin: str = Depends(
        verify_admin
    )
):

    _check_same_origin(request)

    # =====================================================
    # Old Order
    # =====================================================



    # =====================================================
    # Customer
    # =====================================================

    customer_name = (
        customer_name.strip()
    )


    if not customer_name:

        raise HTTPException(
            status_code=400,
            detail="客戶姓名不能空白"
        )


    # =====================================================
    # Date
    # =====================================================

    order_time = (
        order_time.strip()
    )


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


        order_time_value = (
            order_time
        )

    else:

        order_time_value = None


    # =====================================================
    # Amount
    # =====================================================

    amount_rmb = (
        amount_rmb.strip()
    )


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
    # Weight
    # =====================================================

    weight_kg = (
        weight_kg.strip()
    )


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
    # Checkboxes
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
    # Text
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
    # Transaction
    # =====================================================

    ensure_audit_storage()

    conn = get_db()


    try:

        with conn.cursor() as cursor:
            old_order = fetch_locked_order(cursor, order_id)
            if not old_order:
                raise HTTPException(status_code=404, detail="找不到訂單")



            # Update
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


            # Keep the complete after snapshot (including future columns).
            new_order = fetch_locked_order(cursor, order_id)


            # Audit
            write_audit_log(
                cursor=
                    cursor,

                order_id=
                    order_id,

                action=
                    "UPDATE",

                admin_username=
                    admin,

                before_data=
                    old_order,

                after_data=
                    new_order
            )


        conn.commit()


    except Exception:

        conn.rollback()

        raise


    finally:

        conn.close()


    result = query_orders(

        search_order_id=
            search_order_id,

        search_customer_name=
            search_customer_name,

        search_tracking_number=
            search_tracking_number,

        search_amount_rmb=
            search_amount_rmb,

        search_platform=
            search_platform,

        search_order_date=
            search_order_date,

        search_arrived_status=
            search_arrived_status,

        search_returned_status=
            search_returned_status,

        unreturned=
            unreturned
    )


    return templates.TemplateResponse(
        request=request,
        name="order_results.html",
        context={
            **result,

            "search_order_id":
                search_order_id,

            "search_customer_name":
                search_customer_name,

            "search_tracking_number":
                search_tracking_number,

            "search_amount_rmb":
                search_amount_rmb,

            "search_platform":
                search_platform,

            "search_order_date":
                search_order_date,

            "search_arrived_status":
                search_arrived_status,

            "search_returned_status":
                search_returned_status,

            "unreturned":
                unreturned
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

    search_order_id: str = Form(
        ""
    ),

    search_customer_name: str = Form(
        ""
    ),

    search_tracking_number: str = Form(
        ""
    ),

    search_amount_rmb: str = Form(
        ""
    ),

    search_platform: str = Form(
        ""
    ),

    search_order_date: str = Form(
        ""
    ),

    search_arrived_status: str = Form(
        ""
    ),

    search_returned_status: str = Form(
        ""
    ),

    unreturned: str = Form(
        ""
    ),

    admin: str = Depends(
        verify_admin
    )
):

    _check_same_origin(request)



    ensure_audit_storage()

    conn = get_db()


    try:

        with conn.cursor() as cursor:
            old_order = fetch_locked_order(cursor, order_id)
            if not old_order:
                raise HTTPException(status_code=404, detail="找不到訂單")



            # Audit first
            write_audit_log(
                cursor=
                    cursor,

                order_id=
                    order_id,

                action=
                    "DELETE",

                admin_username=
                    admin,

                before_data=
                    old_order,

                after_data=
                    None
            )


            # Delete
            cursor.execute(
                """
                DELETE FROM orders

                WHERE order_id = %s
                """,
                (
                    order_id,
                )
            )
            if cursor.rowcount != 1:
                raise HTTPException(status_code=409, detail="訂單已被其他操作移除")


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

        search_order_id=
            search_order_id,

        search_customer_name=
            search_customer_name,

        search_tracking_number=
            search_tracking_number,

        search_amount_rmb=
            search_amount_rmb,

        search_platform=
            search_platform,

        search_order_date=
            search_order_date,

        search_arrived_status=
            search_arrived_status,

        search_returned_status=
            search_returned_status,

        unreturned=
            unreturned
    )


    return templates.TemplateResponse(
        request=request,
        name="order_results.html",
        context={
            **result,

            "search_order_id":
                search_order_id,

            "search_customer_name":
                search_customer_name,

            "search_tracking_number":
                search_tracking_number,

            "search_amount_rmb":
                search_amount_rmb,

            "search_platform":
                search_platform,

            "search_order_date":
                search_order_date,

            "search_arrived_status":
                search_arrived_status,

            "search_returned_status":
                search_returned_status,

            "unreturned":
                unreturned
        }
    )


# =========================================================
# Audit Logs / Restore
# =========================================================

@app.get("/audit-logs")
async def audit_logs_page(
    request: Request,
    restored: str = "",
    admin: str = Depends(verify_admin),
):
    ensure_audit_storage()
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """SELECT id, order_id, action, admin_username,
                          before_data, after_data, created_at
                   FROM order_audit_logs ORDER BY id DESC LIMIT 300"""
            )
            rows = cursor.fetchall()
            # Include older restores outside the 300 most recent records.
            cursor.execute(
                """SELECT order_id, before_data FROM order_audit_logs
                   WHERE action = 'RESTORE'"""
            )
            restore_rows = cursor.fetchall()
            # Current presence determines whether a delete is already restored.
            deleted_ids = [row["order_id"] for row in rows if row["action"] == "DELETE"]
            existing_order_ids = set()
            if deleted_ids:
                placeholders = ",".join(["%s"] * len(deleted_ids))
                cursor.execute(
                    f"SELECT order_id FROM orders WHERE order_id IN ({placeholders})",
                    deleted_ids,
                )
                existing_order_ids = {r["order_id"] for r in cursor.fetchall()}
    finally:
        conn.close()

    restored_log_ids = set()
    for record in restore_rows:
        ref = parse_audit_json(record["before_data"]).get("delete_log_id")
        if isinstance(ref, int):
            restored_log_ids.add(ref)

    logs = []
    for row in rows:
        before_data = parse_audit_json(row["before_data"])
        after_data = parse_audit_json(row["after_data"])
        changes = []
        if row["action"] == "UPDATE":
            for field, label in AUDIT_FIELD_LABELS.items():
                original = before_data.get(field)
                current = after_data.get(field)
                if str(original) != str(current):
                    changes.append({
                        "field": field,
                        "label": label,
                        "before": display_audit_value(original),
                        "after": display_audit_value(current),
                    })
        logs.append({
            **row,
            "before": before_data,
            "after": after_data,
            "changes": changes,
            "was_restored": row["id"] in restored_log_ids,
            "order_exists": row["order_id"] in existing_order_ids,
        })

    return templates.TemplateResponse(
        request=request,
        name="audit_logs.html",
        context={"logs": logs, "restored": restored},
    )


def _check_same_origin(request: Request):
    # Basic Auth credentials may be sent by browsers automatically; prevent cross-site form POSTs.
    source = request.headers.get("origin") or request.headers.get("referer")
    if not source:
        raise HTTPException(status_code=403, detail="缺少來源資訊，請從操作紀錄頁按復原")
    parsed = urlparse(source)
    expected_host = request.headers.get("host", "").lower()
    if parsed.scheme != "https" or parsed.netloc.lower() != expected_host:
        raise HTTPException(status_code=403, detail="僅允許從同一個後台網址操作")


@app.post("/audit-logs/{audit_id}/restore")
async def restore_deleted_order(
    request: Request,
    audit_id: int,
    admin: str = Depends(verify_admin),
):
    _check_same_origin(request)
    ensure_audit_storage()
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            # Lock this specific deletion record to serialize two simultaneous restore clicks.
            cursor.execute(
                """SELECT id, order_id, action, before_data
                   FROM order_audit_logs WHERE id = %s FOR UPDATE""",
                (audit_id,),
            )
            deletion = cursor.fetchone()
            if not deletion or deletion["action"] != "DELETE":
                raise HTTPException(status_code=404, detail="找不到有效的刪除紀錄")

            snapshot = parse_audit_json(deletion["before_data"])
            order_id = deletion["order_id"]
            if not snapshot or snapshot.get("order_id") != order_id:
                raise HTTPException(status_code=409, detail="刪除快照不完整，無法安全復原")

            # Refuse reusing a past deletion log even if this order was subsequently deleted again.
            cursor.execute(
                """SELECT before_data FROM order_audit_logs
                   WHERE action = 'RESTORE' AND order_id = %s""",
                (order_id,),
            )
            for entry in cursor.fetchall():
                if parse_audit_json(entry["before_data"]).get("delete_log_id") == audit_id:
                    raise HTTPException(status_code=409, detail="這筆刪除紀錄已經復原過")

            cursor.execute(
                "SELECT order_id FROM orders WHERE order_id = %s FOR UPDATE",
                (order_id,),
            )
            if cursor.fetchone():
                raise HTTPException(status_code=409, detail="相同訂單編號已存在，未覆蓋現有訂單")

            cursor.execute("SHOW COLUMNS FROM orders")
            schema = cursor.fetchall()
            live_columns = {
                c["Field"] for c in schema
                if "GENERATED" not in (c.get("Extra") or "").upper()
            }
            # A partial/outdated snapshot must not silently produce a different order.
            if set(snapshot) != live_columns:
                raise HTTPException(
                    status_code=409,
                    detail="刪除時的欄位與目前訂單表不同；請先備份並人工確認，未執行復原",
                )

            fields = [col["Field"] for col in schema if col["Field"] in live_columns]
            columns_sql = ", ".join("`" + field.replace("`", "``") + "`" for field in fields)
            placeholders = ", ".join(["%s"] * len(fields))
            cursor.execute(
                f"INSERT INTO orders ({columns_sql}) VALUES ({placeholders})",
                [snapshot[field] for field in fields],
            )
            new_order = fetch_locked_order(cursor, order_id)
            write_audit_log(
                cursor=cursor,
                order_id=order_id,
                action="RESTORE",
                admin_username=admin,
                before_data={"delete_log_id": audit_id},
                after_data=new_order,
            )
        conn.commit()
    except HTTPException:
        conn.rollback()
        raise
    except pymysql.err.IntegrityError as exc:
        conn.rollback()
        raise HTTPException(
            status_code=409,
            detail="復原與現有資料限制衝突（可能是編號、單號或關聯資料），未修改任何訂單",
        ) from exc
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return RedirectResponse(url=f"/audit-logs?restored={order_id}", status_code=303)


# =========================================================
# 新增訂單（與舊 Streamlit 的新增訂單公式一致）
# =========================================================

ORDER_PLATFORMS = (
    "集運", "拼多多", "淘寶", "閒魚", "1688", "微店",
    "小紅書", "抖音", "京東", "得物",
)
VIP_DISCOUNT = {
    "一般會員": Decimal("1.00"),
    "VIP1": Decimal("0.90"),
    "VIP2": Decimal("0.85"),
    "VIP3": Decimal("0.80"),
}


def member_level_for_name(cursor, name: str) -> str:
    if not name:
        return "一般會員"
    cursor.execute(
        "SELECT member_level FROM members WHERE customer_name = %s LIMIT 1",
        (name,),
    )
    member = cursor.fetchone()
    return str(member["member_level"]) if member and member.get("member_level") else "一般會員"


def calculate_add_fee(amount: Decimal, level: str, platform: str):
    """舊 app.py：集運免費；其餘 1–499→30、500–999→50、之後每 500 +50。
    舊版以 round() 做整數台幣銀行家捨入，這裡使用 Decimal ROUND_HALF_EVEN。
    """
    discount = VIP_DISCOUNT.get(level, Decimal("1.00"))
    if platform == "集運":
        return Decimal("0"), Decimal("0"), discount
    base = Decimal("30") if amount < 500 else Decimal(int(amount // 500) * 50)
    fee = (base * discount).quantize(Decimal("1"), rounding=ROUND_HALF_EVEN)
    return base, fee, discount


def valid_add_amount(raw: str, field: str, max_value: str) -> Decimal:
    try:
        value = Decimal(str(raw).strip())
    except (InvalidOperation, ValueError):
        raise HTTPException(status_code=422, detail=f"{field}格式不正確")
    if not value.is_finite() or value < 0 or value > Decimal(max_value):
        raise HTTPException(status_code=422, detail=f"{field}超出可接受範圍")
    if value.as_tuple().exponent < -2:
        raise HTTPException(status_code=422, detail=f"{field}最多兩位小數")
    return value.quantize(Decimal("0.01"))


def add_order_context(**updates):
    state = {
        "order_time": datetime.now(ZoneInfo("Asia/Taipei")).date().isoformat(),
        "customer_name": "",
        "platform": "集運",
        "tracking_number": "",
        "amount_rmb": "0.00",
        "weight_kg": "0.00",
        "is_arrived": False,
        "is_returned": False,
        "remarks": "",
        "keep_last_name": True,
        "success_id": None,
        "error_message": "",
    }
    state.update(updates)
    return state


def new_order_page_context(**updates):
    context = add_order_context(**updates)
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            # members 與 orders 的 customer_name 可能使用不同 collation；
            # 不在 SQL 用 UNION 合併，以免 MySQL 1271 Illegal mix of collations。
            names = set()
            for source in ("members", "orders"):
                cursor.execute(
                    f"SELECT DISTINCT customer_name FROM {source} "
                    "WHERE customer_name IS NOT NULL AND TRIM(customer_name) <> ''"
                )
                names.update(
                    row["customer_name"] for row in cursor.fetchall()
                    if row.get("customer_name")
                )
            context["customer_names"] = sorted(names)
            context["member_level"] = member_level_for_name(cursor, context["customer_name"])
    finally:
        conn.close()
    context["platforms"] = ORDER_PLATFORMS
    amount = context["amount_rmb"]
    try:
        amount_dec = valid_add_amount(str(amount), "人民幣金額", "99999999.99")
        base, fee, discount = calculate_add_fee(
            amount_dec, context["member_level"], context["platform"]
        )
        context.update({"base_fee": base, "calculated_fee": fee, "vip_discount": discount})
    except HTTPException:
        context.update({"base_fee": None, "calculated_fee": None, "vip_discount": None})
    return context


@app.get("/orders/new")
def new_order_page(request: Request, admin: str = Depends(verify_admin)):
    return templates.TemplateResponse(
        request=request, name="new_order.html", context=new_order_page_context()
    )


@app.get("/orders/new/fee")
def new_order_fee(
    request: Request,
    customer_name: str = "",
    platform: str = "集運",
    amount_rmb: str = "0",
    admin: str = Depends(verify_admin),
):
    if platform not in ORDER_PLATFORMS:
        return templates.TemplateResponse(
            request=request, name="new_order_fee.html",
            context={"error_message": "請選擇有效的平台"},
        )
    try:
        amount = valid_add_amount(amount_rmb or "0", "人民幣金額", "99999999.99")
    except HTTPException as exc:
        return templates.TemplateResponse(
            request=request, name="new_order_fee.html",
            context={"error_message": exc.detail},
        )
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            level = member_level_for_name(cursor, customer_name.strip())
    finally:
        conn.close()
    base, fee, discount = calculate_add_fee(amount, level, platform)
    return templates.TemplateResponse(
        request=request, name="new_order_fee.html",
        context={"member_level": level, "base_fee": base,
                 "calculated_fee": fee, "vip_discount": discount,
                 "error_message": ""},
    )


@app.post("/orders/new")
def create_order(
    request: Request,
    order_time: str = Form(""),
    customer_name: str = Form(""),
    platform: str = Form("集運"),
    tracking_number: str = Form(""),
    amount_rmb: str = Form("0"),
    weight_kg: str = Form("0"),
    is_arrived: Optional[str] = Form(None),
    is_returned: Optional[str] = Form(None),
    remarks: str = Form(""),
    keep_last_name: Optional[str] = Form(None),
    admin: str = Depends(verify_admin),
):
    _check_same_origin(request)
    name = customer_name.strip()
    tracking = tracking_number.strip()
    notes = remarks.strip()
    keep_name = keep_last_name == "1"
    inputs = dict(order_time=order_time, customer_name=name, platform=platform,
                  tracking_number=tracking, amount_rmb=amount_rmb,
                  weight_kg=weight_kg, is_arrived=is_arrived == "1",
                  is_returned=is_returned == "1", remarks=remarks,
                  keep_last_name=keep_name)

    def bad(message):
        return templates.TemplateResponse(
            request=request, name="new_order_panel.html",
            context=new_order_page_context(**inputs, error_message=message),
            status_code=200,  # HTMX replaces the form with a visible validation message.
        )

    if not name or len(name) > 50:
        return bad("請輸入客戶姓名（最多 50 字）")
    if platform not in ORDER_PLATFORMS:
        return bad("請選擇有效的平台")
    if len(tracking) > 50:
        return bad("物流單號最多 50 字")
    if len(notes) > 65535:
        return bad("備註內容過長")
    try:
        date_value = datetime.strptime(order_time, "%Y-%m-%d").date()
    except ValueError:
        return bad("下單日期格式不正確")
    try:
        amount = valid_add_amount(amount_rmb, "人民幣金額", "99999999.99")
        weight = valid_add_amount(weight_kg, "重量", "99999999.99")
    except HTTPException as exc:
        return bad(exc.detail)

    # DDL should not run in an order transaction; fail closed if auditing is unavailable.
    ensure_audit_storage()
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            level = member_level_for_name(cursor, name)
            original_fee, fee, discount = calculate_add_fee(amount, level, platform)
            # Mirror the old Streamlit INSERT column set and recompute the fee server-side.
            cursor.execute(
                """INSERT INTO orders
                   (order_time, customer_name, platform, tracking_number,
                    amount_rmb, weight_kg, is_arrived, is_returned, service_fee, remarks)
                   VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                (date_value, name, platform, tracking or None, amount, weight,
                 int(is_arrived == "1"), int(is_returned == "1"), fee, notes or None),
            )
            new_id = cursor.lastrowid
            cursor.execute(
                "INSERT IGNORE INTO members (customer_name) VALUES (%s)",
                (name,),
            )
            new_order = fetch_locked_order(cursor, new_id)
            if not new_order:
                raise RuntimeError("新增後無法讀取訂單，已取消交易")
            write_audit_log(cursor, new_id, "CREATE", admin, None, new_order)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return templates.TemplateResponse(
        request=request,
        name="new_order_panel.html",
        context=new_order_page_context(
            order_time=order_time, customer_name=name if keep_name else "",
            platform=platform, keep_last_name=keep_name, success_id=new_id,
        ),
    )


# =========================================================
# 入庫管理：沿用 Streamlit 解析與 0.05 kg 進位，增加預覽與交易紀錄
# =========================================================

INBOUND_PATTERNS = (
    re.compile(r"([A-Z]{1,3}\d{8,})[^\d\n]*入庫重量\s*([0-9]+(?:\.[0-9]+)?)\s*KG", re.I),
    re.compile(r"(\d{9,})[^\d\n]*入庫重量\s*([0-9]+(?:\.[0-9]+)?)\s*KG", re.I),
    re.compile(r"單號[:：]?\s*([A-Z0-9]{8,})[^\d\n]*重量[:：]?\s*([0-9]+(?:\.[0-9]+)?)", re.I),
)

INBOUND_LIMIT = 100
INBOUND_MAX_CHARS = 50000
INBOUND_MAX_LINES = 300
INBOUND_WEIGHT_MAX = Decimal("10000")
_failed_storage_ready = False
_failed_storage_lock = threading.Lock()


def ensure_failed_storage():
    """既有 failed_orders 相同結構；DDL 在寫入交易之外執行。"""
    global _failed_storage_ready
    if _failed_storage_ready:
        return
    with _failed_storage_lock:
        if _failed_storage_ready:
            return
        conn = get_db()
        try:
            with conn.cursor() as cursor:
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS failed_orders (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        tracking_number VARCHAR(64) NOT NULL,
                        weight_kg DECIMAL(10,3) NULL,
                        raw_message TEXT NULL,
                        retry_count INT NOT NULL DEFAULT 0,
                        last_error VARCHAR(255) NULL,
                        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
                            ON UPDATE CURRENT_TIMESTAMP,
                        UNIQUE KEY uk_tracking (tracking_number)
                    ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci
                """)
            conn.commit()
            _failed_storage_ready = True
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()


def inbound_weight(value):
    """舊版 round_weight：小於 0.1→0.1；其餘向上到 0.05。"""
    weight = Decimal(str(value))
    if not weight.is_finite() or weight < 0 or weight > INBOUND_WEIGHT_MAX:
        raise ValueError("重量超出可接受範圍")
    if weight < Decimal("0.1"):
        return Decimal("0.10")
    rounded = ((weight / Decimal("0.05")).to_integral_value(rounding=ROUND_CEILING)
               * Decimal("0.05"))
    if rounded > INBOUND_WEIGHT_MAX:
        raise ValueError("重量超出可接受範圍")
    return rounded.quantize(Decimal("0.01"))


def parse_inbound_text(raw_message):
    if len(raw_message) > INBOUND_MAX_CHARS:
        return {"items": [], "issues": ["貼上文字過長，請分批處理（每批不超過 5 萬字）"],
                "conflicts": [], "fatal": True, "duplicate_count": 0}
    lines = raw_message.splitlines()
    if len(lines) > INBOUND_MAX_LINES:
        return {"items": [], "issues": ["一次最多處理 300 行，請分批貼上"],
                "conflicts": [], "fatal": True, "duplicate_count": 0}
    items_by_tracking = {}
    issues = []
    conflicts = []
    duplicate_count = 0
    for number, line in enumerate(lines, 1):
        line = line.strip()
        if not line:
            continue
        matched = next((m for pattern in INBOUND_PATTERNS
                        if (m := pattern.search(line))), None)
        if not matched:
            issues.append(f"第 {number} 行未找到單號＋重量，已略過")
            continue
        tracking = matched.group(1).strip().upper()
        if len(tracking) > 50:
            issues.append(f"第 {number} 行物流單號過長，已略過")
            continue
        try:
            weight = inbound_weight(matched.group(2))
        except (ValueError, InvalidOperation):
            issues.append(f"第 {number} 行重量格式錯誤，已略過")
            continue
        old = items_by_tracking.get(tracking)
        if old is not None:
            duplicate_count += 1
            if old["weight_kg"] != weight:
                conflicts.append(f"第 {number} 行與第 {old['line']} 行的同一物流單號重量不同，請先確認")
            continue
        items_by_tracking[tracking] = {
            "line": number, "tracking_number": tracking,
            "weight_kg": weight, "raw_line": line,
        }
        if len(items_by_tracking) > INBOUND_LIMIT:
            return {"items": [], "issues": ["每批最多 100 個不同單號，請分批貼上"],
                    "conflicts": [], "fatal": True, "duplicate_count": duplicate_count}
    return {"items": list(items_by_tracking.values()), "issues": issues,
            "conflicts": conflicts, "fatal": False, "duplicate_count": duplicate_count}


def inbound_queue_data():
    ensure_failed_storage()
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute("SELECT COUNT(*) AS n FROM failed_orders")
            total = cursor.fetchone()["n"]
            cursor.execute("""SELECT id, tracking_number, weight_kg, raw_message,
                              retry_count, last_error, updated_at FROM failed_orders
                              ORDER BY updated_at DESC, id DESC LIMIT 100""")
            return cursor.fetchall(), total
    finally:
        conn.close()


def inbound_page_context(result=None, raw_message=""):
    queue, queue_total = inbound_queue_data()
    return {"queue": queue, "queue_total": queue_total,
            "result": result, "raw_message": raw_message}


def record_missing_inbound(cursor, tracking, weight, raw, reason):
    # Upsert matches the old Streamlit failed_orders retry queue.
    cursor.execute("""
        INSERT INTO failed_orders
            (tracking_number, weight_kg, raw_message, retry_count, last_error)
        VALUES (%s, %s, %s, 1, %s)
        ON DUPLICATE KEY UPDATE
            weight_kg = VALUES(weight_kg),
            raw_message = VALUES(raw_message),
            last_error = VALUES(last_error),
            retry_count = retry_count + 1,
            updated_at = CURRENT_TIMESTAMP
    """, (tracking, weight, raw[:5000], reason[:250]))


def apply_one_inbound(item, admin):
    """單號一個 transaction；相同單號只一筆記重量，其他筆為 0kg。"""
    tracking, weight = item["tracking_number"], item["weight_kg"]
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute("""SELECT * FROM orders WHERE tracking_number = %s
                              ORDER BY order_id ASC FOR UPDATE""", (tracking,))
            orders = cursor.fetchall()
            if not orders:
                record_missing_inbound(cursor, tracking, weight,
                                      item.get("raw_line", ""), "找不到對應訂單")
                conn.commit()
                return {"tracking_number": tracking, "weight_kg": weight,
                        "status": "待重試", "note": "找不到對應訂單，已加入佇列"}
            # 顯示該物流單號所屬的客戶（同單號可能綁定不同客戶）。
            customer_names = list(dict.fromkeys(
                (str(order.get("customer_name") or "").strip() or "未填姓名")
                for order in orders
            ))
            owners = "、".join(customer_names)
            primary_owner = str(orders[0].get("customer_name") or "").strip() or "未填姓名"
            changed = 0
            now = datetime.now(ZoneInfo("Asia/Taipei")).strftime("%Y-%m-%d %H:%M")
            for idx, before in enumerate(orders):
                target_weight = weight if idx == 0 else Decimal("0.00")
                if before.get("is_arrived") and before.get("weight_kg") is not None \
                        and Decimal(str(before["weight_kg"])) == target_weight:
                    continue
                label = f"主筆={weight}kg" if idx == 0 else "同單號=0kg"
                note = f"｜自動入庫({now}) {label}"
                cursor.execute("""UPDATE orders
                    SET is_arrived = 1, weight_kg = %s,
                        remarks = CONCAT(COALESCE(remarks, ''), %s)
                    WHERE order_id = %s""",
                    (target_weight, note, before["order_id"]))
                after = fetch_locked_order(cursor, before["order_id"])
                write_audit_log(cursor, before["order_id"], "UPDATE", admin, before, after)
                changed += 1
            # 清除同單號的舊失敗佇列，與更新同時 commit。
            cursor.execute("DELETE FROM failed_orders WHERE tracking_number = %s", (tracking,))
        conn.commit()
        return {"tracking_number": tracking, "weight_kg": weight,
                "status": "成功" if changed else "無需變更",
                "note": (f"客戶：{owners}；{len(orders)} 筆訂單，更新 {changed} 筆；"
                         f"主筆 #{orders[0]['order_id']}（{primary_owner}）")}
    except Exception:
        conn.rollback()
        logging.exception("入庫失敗，單號尾碼=%s", tracking[-4:])
        return {"tracking_number": tracking, "weight_kg": weight,
                "status": "失敗", "note": "更新失敗並已回滾，請查看 Railway Logs"}
    finally:
        conn.close()


@app.get("/inbound")
def inbound_page(request: Request, admin: str = Depends(verify_admin)):
    return templates.TemplateResponse(request=request, name="inbound.html",
                                      context=inbound_page_context())


@app.post("/inbound/preview")
def inbound_preview(request: Request, raw_message: str = Form(""),
                    admin: str = Depends(verify_admin)):
    _check_same_origin(request)
    parsed = parse_inbound_text(raw_message)
    if parsed["items"]:
        conn = get_db()
        try:
            with conn.cursor() as cursor:
                keys = [row["tracking_number"] for row in parsed["items"]]
                placeholders = ",".join(["%s"] * len(keys))
                cursor.execute(
                    "SELECT tracking_number, order_id, customer_name "
                    f"FROM orders WHERE tracking_number IN ({placeholders}) ORDER BY order_id ASC",
                    keys)
                matched_orders = {}
                for order in cursor.fetchall():
                    key = str(order["tracking_number"]).upper()
                    matched_orders.setdefault(key, []).append(order)
            for row in parsed["items"]:
                found = matched_orders.get(row["tracking_number"], [])
                row["matched_count"] = len(found)
                row["first_id"] = found[0]["order_id"] if found else None
                row["customer_names"] = list(dict.fromkeys(
                    str(order.get("customer_name") or "").strip() or "未填姓名"
                    for order in found
                ))
        finally:
            conn.close()
    return templates.TemplateResponse(request=request, name="inbound_preview.html",
                                      context={**parsed, "raw_message": raw_message})


@app.post("/inbound/apply")
def inbound_apply(request: Request, raw_message: str = Form(""),
                  admin: str = Depends(verify_admin)):
    _check_same_origin(request)
    parsed = parse_inbound_text(raw_message)
    if parsed["fatal"] or parsed["conflicts"] or not parsed["items"]:
        result = {"error": "解析有衝突或沒有有效單號，請重新預覽後再確認", "rows": []}
    else:
        ensure_audit_storage()
        ensure_failed_storage()
        rows = [apply_one_inbound(row, admin) for row in parsed["items"]]
        result = {"error": "", "rows": rows,
                  "success_count": sum(r["status"] in ("成功", "無需變更") for r in rows),
                  "pending_count": sum(r["status"] == "待重試" for r in rows)}
    return templates.TemplateResponse(request=request, name="inbound_body.html",
                                      context=inbound_page_context(result=result))


@app.post("/inbound/retry")
def inbound_retry(request: Request, admin: str = Depends(verify_admin)):
    _check_same_origin(request)
    ensure_audit_storage()
    ensure_failed_storage()
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute("""SELECT id, tracking_number, weight_kg, raw_message
                              FROM failed_orders ORDER BY id ASC LIMIT 100""")
            queued = cursor.fetchall()
    finally:
        conn.close()
    results = []
    for row in queued:
        try:
            weight = inbound_weight(row["weight_kg"])
        except (ValueError, TypeError, InvalidOperation):
            results.append({"tracking_number": row["tracking_number"], "weight_kg": "—",
                            "status": "失敗", "note": "佇列重量無效，請刪除後重新貼上"})
            continue
        results.append(apply_one_inbound({"tracking_number": row["tracking_number"],
                                          "weight_kg": weight,
                                          "raw_line": row.get("raw_message") or ""}, admin))
    result = {"error": "", "rows": results,
              "success_count": sum(r["status"] in ("成功", "無需變更") for r in results),
              "pending_count": sum(r["status"] == "待重試" for r in results)}
    return templates.TemplateResponse(request=request, name="inbound_body.html",
                                      context=inbound_page_context(result=result))


@app.post("/inbound/queue/{queue_id}/delete")
def inbound_delete_failed(request: Request, queue_id: int,
                          admin: str = Depends(verify_admin)):
    _check_same_origin(request)
    ensure_failed_storage()
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute("DELETE FROM failed_orders WHERE id = %s", (queue_id,))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return templates.TemplateResponse(request=request, name="inbound_body.html",
                                      context=inbound_page_context(
                                          result={"error": "", "rows": [], "message": "佇列項目已刪除"}))

# =========================================================
# 出貨管理：可出貨名單／姓名批次處理（保留舊 Streamlit 的備註標記）
# =========================================================

SHIPPING_DELAY_TAG = "[延後]"
SHIPPING_NOTIFY_TAG = "[已通知]"
SHIPPING_ACTIONS = {
    "returned": "標記已運回",
    "early": "標記提前運回",
    "delay": "延後運回",
    "undelay": "取消延後",
    "notify": "標記已通知",
    "unnotify": "取消已通知",
}


def shipping_rows(mode="ready", customer="", hide_delayed=False, hide_notified=False):
    """第一版以當前訂單快照產生名單；不會自動變更訂單。"""
    mode = mode if mode in ("ready", "customer") else "ready"
    customer = str(customer or "").strip()
    if len(customer) > 50:
        customer = customer[:50]
    if mode == "customer" and not customer:
        return [], 0

    ready_filter = """
        COALESCE(o.is_arrived, 0) = 1
        AND (
            COALESCE(o.is_early_returned, 0) = 1
            OR NOT EXISTS (
                SELECT 1 FROM orders pending
                WHERE pending.customer_name = o.customer_name
                  AND COALESCE(pending.is_arrived, 0) = 0
                  AND COALESCE(pending.is_returned, 0) = 0
                  AND COALESCE(pending.order_status, '正常') <> '取消'
            )
        )
    """
    where = [
        "COALESCE(o.is_returned, 0) = 0",
        "COALESCE(o.order_status, '正常') <> '取消'",
        "o.customer_name IS NOT NULL",
        "TRIM(o.customer_name) <> ''",
    ]
    params = []
    if mode == "ready":
        where.append(f"({ready_filter})")
    else:
        where.append("o.customer_name LIKE %s")
        params.append(f"%{customer}%")
    if hide_delayed:
        where.append("(o.remarks IS NULL OR o.remarks NOT LIKE %s)")
        params.append(f"%{SHIPPING_DELAY_TAG}%")
    if hide_notified:
        where.append("(o.remarks IS NULL OR o.remarks NOT LIKE %s)")
        params.append(f"%{SHIPPING_NOTIFY_TAG}%")
    condition = " AND ".join(where)
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""SELECT o.order_id, o.order_time, o.customer_name, o.platform,
                           o.tracking_number, o.weight_kg, o.is_arrived,
                           o.is_returned, o.is_early_returned, o.remarks
                    FROM orders o WHERE {condition}
                    ORDER BY o.customer_name, o.order_id DESC LIMIT 501""",
                params,
            )
            all_rows = cursor.fetchall()
    finally:
        conn.close()
    has_more = len(all_rows) > 500
    rows = [row for row in all_rows[:500] if not row.get("is_returned")]
    for row in rows:
        note = str(row.get("remarks") or "")
        row["delayed"] = SHIPPING_DELAY_TAG in note
        row["notified"] = SHIPPING_NOTIFY_TAG in note
        row["weight_num"] = float(row.get("weight_kg") or 0)
    return rows, int(has_more)


def shipping_billable_weight(weight, kind):
    """每位客戶、每一種包裹分類合併重量後，以 0.5 kg 進位。"""
    weight = max(Decimal("0"), Decimal(str(weight or 0)))
    if weight == 0:
        # 未記重的訂單不估運費，避免將未知重量視為已知費用。
        return Decimal("0")
    rounded = (weight / Decimal("0.5")).to_integral_value(rounding=ROUND_CEILING) * Decimal("0.5")
    return max(Decimal("1.0"), rounded) if kind == "forwarding" else max(Decimal("0.5"), rounded)


def shipping_fee_breakdown(rows):
    """同一客戶分成純集運/代購兩池，各自合重計費；不同客戶分別進位。"""
    grouped = {}
    for row in rows:
        if row.get("is_returned") or row.get("order_status") == "取消":
            continue
        name = row["customer_name"]
        g = grouped.setdefault(name, {
            "name": name, "count": 0, "weight": Decimal("0"),
            "forwarding_count": 0, "forwarding_weight": Decimal("0"),
            "purchase_count": 0, "purchase_weight": Decimal("0"),
        })
        kind = "forwarding" if row.get("platform") == "集運" else "purchase"
        weight = max(Decimal("0"), Decimal(str(row.get("weight_kg") or 0)))
        g["count"] += 1
        g["weight"] += weight
        g[f"{kind}_count"] += 1
        g[f"{kind}_weight"] += weight

    groups = []
    for g in grouped.values():
        g["forwarding_billed_weight"] = shipping_billable_weight(g["forwarding_weight"], "forwarding")
        g["purchase_billed_weight"] = shipping_billable_weight(g["purchase_weight"], "purchase")
        g["forwarding_fee"] = g["forwarding_billed_weight"] * Decimal("90")
        g["purchase_fee"] = g["purchase_billed_weight"] * Decimal("70")
        g["total_fee"] = g["forwarding_fee"] + g["purchase_fee"]
        groups.append(g)
    return groups


def shipping_context(mode="ready", customer="", hide_delayed="", hide_notified="", message="", error=""):
    mode = mode if mode in ("ready", "customer") else "ready"
    customer = str(customer or "").strip()[:50]
    hide_delayed = str(hide_delayed or "") == "1"
    hide_notified = str(hide_notified or "") == "1"
    rows, has_more = shipping_rows(mode, customer, hide_delayed, hide_notified)
    # 已運回訂單不進可選取清單、不計重量或運費（SQL 亦已排除）。
    rows = [row for row in rows if not row.get("is_returned")]
    total_weight = sum((max(Decimal("0"), Decimal(str(row.get("weight_kg") or 0))) for row in rows), Decimal("0"))
    groups = shipping_fee_breakdown(rows)
    total_fee = sum((g["total_fee"] for g in groups), Decimal("0"))
    return {
        "mode": mode,
        "customer": customer,
        "hide_delayed": hide_delayed,
        "hide_notified": hide_notified,
        "rows": rows,
        "count": len(rows),
        "total_weight": total_weight,
        "total_fee": total_fee,
        "groups": groups,
        "has_more": bool(has_more),
        "message": message,
        "error": error,
        "actions": SHIPPING_ACTIONS,
    }


def _shipping_update_text(remarks, action):
    before = str(remarks or "")
    if action == "delay" or action == "notify":
        tag = SHIPPING_DELAY_TAG if action == "delay" else SHIPPING_NOTIFY_TAG
        return before if tag in before else (before + " " + tag).strip()
    if action == "undelay" or action == "unnotify":
        tag = SHIPPING_DELAY_TAG if action == "undelay" else SHIPPING_NOTIFY_TAG
        return before.replace(tag, "").strip()
    return before


def apply_shipping_action(order_ids, action, admin, mode="ready", customer="",
                          hide_delayed=False, hide_notified=False, max_orders=100):
    """一批訂單同 transaction；無變更不寫 log。出錯時整批回滾。"""
    if action not in SHIPPING_ACTIONS:
        raise HTTPException(status_code=422, detail="無效的出貨操作")
    max_orders = max(1, min(int(max_orders or 100), 500))
    if len(order_ids) != len(set(order_ids)) or not 1 <= len(order_ids) <= max_orders:
        raise HTTPException(status_code=422, detail=f"每批請選擇 1～{max_orders} 筆不重複訂單")
    if any(oid <= 0 for oid in order_ids):
        raise HTTPException(status_code=422, detail="訂單編號不正確")
    if mode not in ("ready", "customer"):
        raise HTTPException(status_code=422, detail="出貨模式不正確")
    # 限制提交的訂單必須出現在本次查詢範圍，避免改造表單跨客戶批次操作。
    visible, truncated = shipping_rows(mode, customer, hide_delayed, hide_notified)
    allowed = {row["order_id"] for row in visible}
    if not set(order_ids).issubset(allowed):
        raise HTTPException(status_code=409, detail="訂單名單已變動，請刷新列表後重新勾選")
    ensure_audit_storage()
    conn = get_db()
    changed = 0
    try:
        with conn.cursor() as cursor:
            keys = sorted(order_ids)
            placeholders = ",".join(["%s"] * len(keys))
            cursor.execute(
                f"SELECT * FROM orders WHERE order_id IN ({placeholders}) ORDER BY order_id FOR UPDATE",
                keys,
            )
            before_rows = cursor.fetchall()
            if len(before_rows) != len(keys):
                raise HTTPException(status_code=409, detail="部分訂單已不存在，請重新查詢")
            for old in before_rows:
                oid = old["order_id"]
                if old.get("order_status") == "取消" or old.get("is_returned"):
                    raise HTTPException(status_code=409, detail=f"#{oid} 已取消或已運回，整批未更動")
                if action == "returned":
                    if not old.get("is_arrived"):
                        raise HTTPException(status_code=409, detail=f"#{oid} 尚未到貨，整批未更動")
                    if SHIPPING_DELAY_TAG in str(old.get("remarks") or ""):
                        raise HTTPException(status_code=409, detail=f"#{oid} 標記延後，請先取消延後再出貨")
                    if old.get("is_returned"):
                        continue
                    cursor.execute("UPDATE orders SET is_returned = 1 WHERE order_id = %s", (oid,))
                elif action == "early":
                    if old.get("is_early_returned"):
                        continue
                    cursor.execute("UPDATE orders SET is_early_returned = 1 WHERE order_id = %s", (oid,))
                else:
                    revised = _shipping_update_text(old.get("remarks"), action)
                    if revised == str(old.get("remarks") or ""):
                        continue
                    cursor.execute("UPDATE orders SET remarks = %s WHERE order_id = %s", (revised, oid))
                after = fetch_locked_order(cursor, oid)
                write_audit_log(cursor, oid, "UPDATE", admin, old, after)
                changed += 1
        conn.commit()
        return changed
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


@app.get("/shipping")
def shipping_page(request: Request, mode: str = "ready", customer: str = "",
                  hide_delayed: str = "", hide_notified: str = "",
                  admin: str = Depends(verify_admin)):
    context = shipping_context(mode, customer, hide_delayed, hide_notified)
    return templates.TemplateResponse(request=request, name="shipping.html", context=context)


@app.get("/shipping/list")
def shipping_list(request: Request, mode: str = "ready", customer: str = "",
                  hide_delayed: str = "", hide_notified: str = "",
                  admin: str = Depends(verify_admin)):
    context = shipping_context(mode, customer, hide_delayed, hide_notified)
    return templates.TemplateResponse(request=request, name="shipping_body.html", context=context)


@app.post("/shipping/apply")
def shipping_apply(request: Request, selected_order_ids: list[int] = Form([]),
                   action: str = Form(""), mode: str = Form("ready"),
                   customer: str = Form(""), hide_delayed: str = Form(""),
                   hide_notified: str = Form(""),
                   admin: str = Depends(verify_admin)):
    _check_same_origin(request)
    error = ""
    message = ""
    try:
        changed = apply_shipping_action(
            selected_order_ids, action, admin, mode, customer,
            hide_delayed == "1", hide_notified == "1",
        )
        message = f"{SHIPPING_ACTIONS[action]}：成功更新 {changed} 筆訂單，並記錄變更。" if changed else "選取的訂單沒有需要更新的欄位。"
    except HTTPException as exc:
        if exc.status_code not in (409, 422):
            raise
        error = str(exc.detail)
    context = shipping_context(mode, customer, hide_delayed, hide_notified,
                               message=message, error=error)
    return templates.TemplateResponse(request=request, name="shipping_body.html", context=context)


@app.post("/shipping/apply-customers")
def shipping_apply_customers(
    request: Request,
    selected_customers: list[str] = Form([]),
    action: str = Form(""),
    mode: str = Form("ready"),
    customer: str = Form(""),
    hide_delayed: str = Form(""),
    hide_notified: str = Form(""),
    admin: str = Depends(verify_admin),
):
    """將批次操作套用到勾選客戶在目前名單中的所有未運回訂單。"""
    _check_same_origin(request)
    error = ""
    message = ""
    try:
        names = [str(name or "").strip() for name in selected_customers]
        names = [name for name in names if name]
        if len(names) != len(set(names)) or not 1 <= len(names) <= 100:
            raise HTTPException(status_code=422, detail="每批請選擇 1～100 位不重複客戶")
        if action not in SHIPPING_ACTIONS:
            raise HTTPException(status_code=422, detail="請先選擇有效的批次操作")

        visible_rows, _ = shipping_rows(
            mode, customer, hide_delayed == "1", hide_notified == "1"
        )
        visible_names = {row["customer_name"] for row in visible_rows}
        if not set(names).issubset(visible_names):
            raise HTTPException(status_code=409, detail="客戶名單已變動，請更新列表後重新勾選")

        chosen = set(names)
        order_ids = [
            int(row["order_id"])
            for row in visible_rows
            if row["customer_name"] in chosen and not row.get("is_returned")
        ]
        if not order_ids:
            raise HTTPException(status_code=409, detail="勾選客戶目前沒有可批次處理的未運回訂單")
        if len(order_ids) > 500:
            raise HTTPException(status_code=422, detail="勾選客戶對應超過 500 筆訂單，請分批處理")

        changed = apply_shipping_action(
            order_ids, action, admin, mode, customer,
            hide_delayed == "1", hide_notified == "1",
            max_orders=500,
        )
        customer_count = len(names)
        message = (
            f"{SHIPPING_ACTIONS[action]}：已對 {customer_count} 位客戶執行，成功更新 {changed} 筆訂單，並記錄變更。"
            if changed else
            f"勾選的 {customer_count} 位客戶目前沒有需要更新的欄位。"
        )
    except HTTPException as exc:
        if exc.status_code not in (409, 422):
            raise
        error = str(exc.detail)

    context = shipping_context(
        mode, customer, hide_delayed, hide_notified,
        message=message, error=error
    )
    return templates.TemplateResponse(
        request=request, name="shipping_body.html", context=context
    )


# =========================================================
# 出貨 Excel：選客戶匯出統整／名下訂單明細，或只匯出勾選訂單
# =========================================================

@app.post("/shipping/export")
def shipping_export(
    request: Request,
    export_kind: str = Form(""),
    return_date: str = Form(""),
    selected_customers: list[str] = Form([]),
    selected_order_ids: list[int] = Form([]),
    mode: str = Form("ready"),
    customer: str = Form(""),
    hide_delayed: str = Form(""),
    hide_notified: str = Form(""),
    admin: str = Depends(verify_admin),
):
    _check_same_origin(request)
    if export_kind not in ("customer_summary", "customer_details", "selected_orders"):
        raise HTTPException(status_code=422, detail="匯出類型不正確")
    if mode not in ("ready", "customer"):
        raise HTTPException(status_code=422, detail="出貨模式不正確")

    # 所有出貨 Excel 都必須由使用者先指定本次運回日期。
    return_date = str(return_date or "").strip()
    try:
        ship_date = datetime.strptime(return_date, "%Y-%m-%d").date()
    except ValueError:
        raise HTTPException(status_code=422, detail="請先選擇正確的運回日期")

    if export_kind == "selected_orders":
        if not 1 <= len(selected_order_ids) <= 500 or len(set(selected_order_ids)) != len(selected_order_ids):
            raise HTTPException(status_code=422, detail="請先選取 1～500 筆不重複訂單")
    else:
        max_customers = 100 if export_kind == "customer_summary" else 500
        if not 1 <= len(selected_customers) <= max_customers or len(set(selected_customers)) != len(selected_customers):
            if export_kind == "customer_summary":
                raise HTTPException(status_code=422, detail="賣貨便統整表一次請選取 1～100 位不重複客戶")
            raise HTTPException(status_code=422, detail="請先選取 1～500 位不重複客戶")

    # 與目前畫面使用完全相同的查詢條件及前 500 筆範圍，避免偽造跨客戶匯出。
    rows, _has_more = shipping_rows(
        mode, customer, hide_delayed == "1", hide_notified == "1"
    )
    if export_kind == "selected_orders":
        allowed_ids = {row["order_id"] for row in rows}
        if not set(selected_order_ids).issubset(allowed_ids):
            raise HTTPException(status_code=409, detail="名單已變動，請更新後重新勾選訂單")
        chosen = set(selected_order_ids)
        export_rows = [row for row in rows if row["order_id"] in chosen]
    else:
        allowed_names = {row["customer_name"] for row in rows}
        if not set(selected_customers).issubset(allowed_names):
            raise HTTPException(status_code=409, detail="名單已變動，請更新後重新勾選客戶")
        chosen = set(selected_customers)
        export_rows = [row for row in rows if row["customer_name"] in chosen]

    # 統整表的價格仍依同一位客戶「純集運 / 代購」分池合重後計算。
    groups = shipping_fee_breakdown(export_rows)
    date_code = ship_date.strftime("%Y%m%d")

    try:
        if export_kind == "customer_summary":
            payload = make_sellnow_xlsm(groups, ship_date)
            filename = f"{date_code}賣貨便批次匯入.xlsm"
            media_type = "application/vnd.ms-excel.sheet.macroEnabled.12"
        else:
            payload = make_shipping_detail_xlsx(export_rows, ship_date)
            # 使用者指定的固定檔名格式：(運回日期)明細表
            filename = f"{date_code}明細表.xlsx"
            media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    except (ValueError, FileNotFoundError) as exc:
        raise HTTPException(status_code=422, detail=str(exc))

    return Response(
        content=payload,
        media_type=media_type,
        headers={
            "Content-Disposition": f"attachment; filename=shipping; filename*=UTF-8''{quote(filename)}",
            "Cache-Control": "no-store",
            "X-Content-Type-Options": "nosniff",
        },
    )



# =========================================================
# 前台運回申請管理
# =========================================================

RETURN_REQUEST_STATUS_LABELS = {
    "pending": "待處理",
    "processed": "已處理",
    "cancelled": "已取消",
}


def ensure_return_request_tables():
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS customer_return_requests (
                    request_id INT AUTO_INCREMENT PRIMARY KEY,
                    customer_name VARCHAR(255) NOT NULL,
                    selected_shipping_batch VARCHAR(255) NOT NULL,
                    delivery_method VARCHAR(50) NOT NULL DEFAULT '面交/自取',
                    total_count INT NOT NULL DEFAULT 0,
                    total_weight DECIMAL(10,3) NOT NULL DEFAULT 0,
                    estimated_fee DECIMAL(10,2) NOT NULL DEFAULT 0,
                    status ENUM('pending','processed','cancelled') NOT NULL DEFAULT 'pending',
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                    INDEX idx_return_request_status_created (status, created_at),
                    INDEX idx_return_request_customer (customer_name)
                ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS customer_return_request_items (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    request_id INT NOT NULL,
                    order_id INT NOT NULL,
                    tracking_number VARCHAR(255) NULL,
                    platform VARCHAR(50) NULL,
                    weight_kg DECIMAL(10,3) NULL,
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE KEY uk_request_order (request_id, order_id),
                    INDEX idx_return_item_request (request_id),
                    CONSTRAINT fk_return_req_items_request
                        FOREIGN KEY (request_id) REFERENCES customer_return_requests(request_id)
                        ON DELETE CASCADE
                ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci
                """
            )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def return_request_rows(status="pending", customer=""):
    ensure_return_request_tables()
    status = status if status in ("pending", "processed", "cancelled", "all") else "pending"
    customer = str(customer or "").strip()[:80]
    where = []
    params = []
    if status != "all":
        where.append("r.status = %s")
        params.append(status)
    if customer:
        where.append("r.customer_name LIKE %s")
        params.append(f"%{customer}%")
    where_sql = "WHERE " + " AND ".join(where) if where else ""
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT
                    r.request_id,
                    r.customer_name,
                    r.selected_shipping_batch,
                    r.delivery_method,
                    r.total_count,
                    r.total_weight,
                    r.estimated_fee,
                    r.status,
                    r.created_at,
                    r.updated_at
                FROM customer_return_requests r
                {where_sql}
                ORDER BY r.created_at DESC, r.request_id DESC
                LIMIT 300
                """,
                params,
            )
            rows = cursor.fetchall()
    finally:
        conn.close()
    for row in rows:
        row["total_weight_num"] = float(row.get("total_weight") or 0)
        row["estimated_fee_num"] = float(row.get("estimated_fee") or 0)
        row["status_label"] = RETURN_REQUEST_STATUS_LABELS.get(row.get("status"), row.get("status") or "—")
    return rows


def return_request_context(status="pending", customer="", message="", error=""):
    rows = return_request_rows(status, customer)
    return {
        "status": status if status in ("pending", "processed", "cancelled", "all") else "pending",
        "customer": str(customer or "").strip()[:80],
        "requests": rows,
        "total_count": sum(int(row.get("total_count") or 0) for row in rows),
        "total_weight": sum(float(row.get("total_weight") or 0) for row in rows),
        "total_fee": sum(float(row.get("estimated_fee") or 0) for row in rows),
        "message": message,
        "error": error,
    }


def fetch_return_request_detail(request_id: int):
    ensure_return_request_tables()
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                SELECT request_id, customer_name, selected_shipping_batch,
                       delivery_method, total_count, total_weight, estimated_fee,
                       status, created_at, updated_at
                FROM customer_return_requests
                WHERE request_id = %s
                LIMIT 1
                """,
                (request_id,),
            )
            request_row = cursor.fetchone()
            if not request_row:
                return None, []
            cursor.execute(
                """
                SELECT
                    i.order_id,
                    COALESCE(o.order_time, NULL) AS order_time,
                    COALESCE(o.customer_name, %s) AS customer_name,
                    COALESCE(o.platform, i.platform) AS platform,
                    COALESCE(o.tracking_number, i.tracking_number) AS tracking_number,
                    COALESCE(o.weight_kg, i.weight_kg, 0) AS weight_kg,
                    COALESCE(o.is_arrived, 0) AS is_arrived,
                    COALESCE(o.is_returned, 0) AS is_returned,
                    COALESCE(o.remarks, '') AS remarks
                FROM customer_return_request_items i
                LEFT JOIN orders o ON o.order_id = i.order_id
                WHERE i.request_id = %s
                ORDER BY i.order_id ASC
                """,
                (request_row["customer_name"], request_id),
            )
            items = cursor.fetchall()
    finally:
        conn.close()
    request_row["total_weight_num"] = float(request_row.get("total_weight") or 0)
    request_row["estimated_fee_num"] = float(request_row.get("estimated_fee") or 0)
    for item in items:
        item["weight_num"] = float(item.get("weight_kg") or 0)
    return request_row, items


@app.get("/shipping/requests")
def shipping_requests_page(
    request: Request,
    status: str = "pending",
    customer: str = "",
    message: str = "",
    admin: str = Depends(verify_admin),
):
    context = return_request_context(status, customer, message=message)
    return templates.TemplateResponse(request=request, name="shipping_requests.html", context=context)


@app.get("/shipping/requests/{request_id}/detail")
def shipping_request_detail(
    request: Request,
    request_id: int,
    admin: str = Depends(verify_admin),
):
    request_row, items = fetch_return_request_detail(request_id)
    if not request_row:
        raise HTTPException(status_code=404, detail="找不到這筆運回申請")
    return templates.TemplateResponse(
        request=request,
        name="shipping_request_detail.html",
        context={"request_row": request_row, "items": items},
    )


@app.post("/shipping/requests/apply")
def shipping_requests_apply(
    request: Request,
    selected_request_ids: list[int] = Form([]),
    action: str = Form(""),
    status_filter: str = Form("pending"),
    customer_filter: str = Form(""),
    admin: str = Depends(verify_admin),
):
    _check_same_origin(request)
    if action not in ("processed", "cancelled"):
        raise HTTPException(status_code=422, detail="申請操作不正確")
    ids = [int(x) for x in selected_request_ids]
    if len(ids) != len(set(ids)) or not 1 <= len(ids) <= 100:
        raise HTTPException(status_code=422, detail="每次請選擇 1～100 筆不重複申請")
    ensure_return_request_tables()
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            placeholders = ",".join(["%s"] * len(ids))
            cursor.execute(
                f"""
                SELECT request_id, status
                FROM customer_return_requests
                WHERE request_id IN ({placeholders})
                FOR UPDATE
                """,
                ids,
            )
            current = cursor.fetchall()
            if len(current) != len(ids):
                raise HTTPException(status_code=409, detail="部分申請已不存在，請重新整理")
            non_pending = [row["request_id"] for row in current if row.get("status") != "pending"]
            if non_pending:
                raise HTTPException(status_code=409, detail="部分申請已被處理，請重新整理後再操作")
            cursor.execute(
                f"UPDATE customer_return_requests SET status = %s WHERE request_id IN ({placeholders}) AND status = 'pending'",
                [action] + ids,
            )
            changed = cursor.rowcount
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    label = RETURN_REQUEST_STATUS_LABELS[action]
    # 回到待處理頁，避免使用者誤以為已處理的申請仍在待辦。
    target_status = "pending" if status_filter in ("pending", "all") else status_filter
    customer_filter = str(customer_filter or "").strip()
    qs = f"status={quote(target_status)}&customer={quote(customer_filter)}&message={quote(f'已將 {changed} 筆申請標記為{label}。')}"
    return RedirectResponse(url=f"/shipping/requests?{qs}", status_code=303)


# =========================================================
# 會員管理：沿用舊 Streamlit 的會員等級 / 備註 / LINE 綁定資訊
# =========================================================

MEMBER_LEVELS = ("一般會員", "VIP1", "VIP2", "VIP3")


def ensure_members_table_admin():
    """確保會員表與舊版需要的欄位存在；不建立餘額 / 儲值欄位。"""
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS members (
                    member_id INT AUTO_INCREMENT PRIMARY KEY,
                    customer_name VARCHAR(255) NOT NULL,
                    member_level VARCHAR(50) NOT NULL DEFAULT '一般會員',
                    note TEXT NULL,
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                    line_user_id VARCHAR(100) NULL,
                    line_name VARCHAR(100) NULL,
                    UNIQUE KEY uk_customer_name (customer_name)
                ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci
                """
            )
            cursor.execute("SHOW COLUMNS FROM members")
            existing = {row["Field"] for row in cursor.fetchall()}
            additions = {
                "member_level": "VARCHAR(50) NOT NULL DEFAULT '一般會員'",
                "note": "TEXT NULL",
                "created_at": "TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP",
                "updated_at": "TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP",
                "line_user_id": "VARCHAR(100) NULL",
                "line_name": "VARCHAR(100) NULL",
            }
            for name, ddl in additions.items():
                if name not in existing:
                    cursor.execute(f"ALTER TABLE members ADD COLUMN {name} {ddl}")
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def sync_members_from_orders_admin():
    """將 orders 中尚未存在的客戶同步進 members；不改既有會員資料。"""
    ensure_members_table_admin()
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                SELECT DISTINCT customer_name
                FROM orders
                WHERE customer_name IS NOT NULL
                  AND TRIM(customer_name) <> ''
                """
            )
            names = [row["customer_name"] for row in cursor.fetchall() if row.get("customer_name")]
            if names:
                cursor.executemany(
                    "INSERT IGNORE INTO members (customer_name) VALUES (%s)",
                    [(name,) for name in names],
                )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def member_order_stats(names):
    """避免 members / orders customer_name 不同 collation 的 JOIN 問題，以 Python 合併統計。"""
    names = [str(x) for x in names if x]
    if not names:
        return {}
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            result = {}
            # 依批次查詢，避免 IN 過長。
            for start in range(0, len(names), 200):
                batch = names[start:start + 200]
                placeholders = ",".join(["%s"] * len(batch))
                cursor.execute(
                    f"""
                    SELECT
                        customer_name,
                        COUNT(*) AS order_count,
                        SUM(CASE WHEN COALESCE(is_returned,0)=0 THEN 1 ELSE 0 END) AS unreturned_count,
                        COALESCE(SUM(weight_kg),0) AS total_weight,
                        MAX(order_time) AS last_order_date
                    FROM orders
                    WHERE customer_name IN ({placeholders})
                    GROUP BY customer_name
                    """,
                    batch,
                )
                for row in cursor.fetchall():
                    result[str(row["customer_name"])] = row
            return result
    finally:
        conn.close()


def member_rows(keyword="", level="all"):
    sync_members_from_orders_admin()
    keyword = str(keyword or "").strip()[:100]
    level = str(level or "all")
    if level not in ("all",) + MEMBER_LEVELS:
        level = "all"

    where = ["1=1"]
    params = []
    if keyword:
        where.append("(customer_name LIKE %s OR COALESCE(line_name,'') LIKE %s)")
        like = f"%{keyword}%"
        params.extend([like, like])
    if level != "all":
        where.append("member_level = %s")
        params.append(level)

    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT member_id, customer_name, member_level, note,
                       line_user_id, line_name, created_at, updated_at
                FROM members
                WHERE {' AND '.join(where)}
                ORDER BY updated_at DESC, member_id DESC
                LIMIT 500
                """,
                params,
            )
            rows = cursor.fetchall()
            cursor.execute(
                """
                SELECT member_level, COUNT(*) AS cnt
                FROM members
                GROUP BY member_level
                """
            )
            counts = {row["member_level"]: int(row["cnt"] or 0) for row in cursor.fetchall()}
    finally:
        conn.close()

    stats = member_order_stats([row.get("customer_name") for row in rows])
    for row in rows:
        st = stats.get(str(row.get("customer_name")), {})
        row["order_count"] = int(st.get("order_count") or 0)
        row["unreturned_count"] = int(st.get("unreturned_count") or 0)
        row["total_weight_num"] = float(st.get("total_weight") or 0)
        row["last_order_date"] = st.get("last_order_date")
        if row.get("member_level") not in MEMBER_LEVELS:
            row["member_level"] = "一般會員"
    return rows, counts, keyword, level


def fetch_member_admin(member_id: int):
    ensure_members_table_admin()
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                SELECT member_id, customer_name, member_level, note,
                       line_user_id, line_name, created_at, updated_at
                FROM members
                WHERE member_id = %s
                LIMIT 1
                """,
                (member_id,),
            )
            return cursor.fetchone()
    finally:
        conn.close()


def member_detail_context(member_id: int):
    member = fetch_member_admin(member_id)
    if not member:
        raise HTTPException(status_code=404, detail="找不到會員")
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                SELECT order_id, order_time, platform, tracking_number,
                       amount_rmb, weight_kg, is_arrived, is_returned, order_status
                FROM orders
                WHERE customer_name = %s
                ORDER BY order_time DESC, order_id DESC
                LIMIT 50
                """,
                (member["customer_name"],),
            )
            orders = cursor.fetchall()
            cursor.execute(
                """
                SELECT COUNT(*) AS order_count,
                       SUM(CASE WHEN COALESCE(is_returned,0)=0 THEN 1 ELSE 0 END) AS unreturned_count,
                       COALESCE(SUM(weight_kg),0) AS total_weight,
                       COALESCE(SUM(amount_rmb),0) AS total_rmb
                FROM orders
                WHERE customer_name = %s
                """,
                (member["customer_name"],),
            )
            stats = cursor.fetchone() or {}
    finally:
        conn.close()
    stats["order_count"] = int(stats.get("order_count") or 0)
    stats["unreturned_count"] = int(stats.get("unreturned_count") or 0)
    stats["total_weight_num"] = float(stats.get("total_weight") or 0)
    stats["total_rmb_num"] = float(stats.get("total_rmb") or 0)
    return {"member": member, "orders": orders, "stats": stats, "member_levels": MEMBER_LEVELS}


@app.get("/members")
def members_page(
    request: Request,
    keyword: str = "",
    level: str = "all",
    admin: str = Depends(verify_admin),
):
    rows, counts, keyword, level = member_rows(keyword, level)
    return templates.TemplateResponse(
        request=request,
        name="members.html",
        context={
            "members": rows,
            "level_counts": counts,
            "keyword": keyword,
            "level": level,
            "member_levels": MEMBER_LEVELS,
        },
    )


@app.get("/members/search")
def members_search(
    request: Request,
    keyword: str = "",
    level: str = "all",
    admin: str = Depends(verify_admin),
):
    rows, counts, keyword, level = member_rows(keyword, level)
    return templates.TemplateResponse(
        request=request,
        name="members_results.html",
        context={
            "members": rows,
            "level_counts": counts,
            "keyword": keyword,
            "level": level,
            "member_levels": MEMBER_LEVELS,
        },
    )


@app.get("/members/{member_id}/detail")
def member_detail(
    request: Request,
    member_id: int,
    admin: str = Depends(verify_admin),
):
    return templates.TemplateResponse(
        request=request,
        name="member_detail.html",
        context=member_detail_context(member_id),
    )


@app.get("/members/{member_id}/edit")
def member_edit_page(
    request: Request,
    member_id: int,
    admin: str = Depends(verify_admin),
):
    context = member_detail_context(member_id)
    return templates.TemplateResponse(request=request, name="member_edit.html", context=context)


@app.post("/members/{member_id}/edit")
def member_edit_save(
    request: Request,
    member_id: int,
    member_level: str = Form("一般會員"),
    note: str = Form(""),
    admin: str = Depends(verify_admin),
):
    _check_same_origin(request)
    if member_level not in MEMBER_LEVELS:
        raise HTTPException(status_code=422, detail="會員等級不正確")
    note_value = str(note or "").strip()
    if len(note_value) > 65535:
        raise HTTPException(status_code=422, detail="備註內容過長")

    ensure_members_table_admin()
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                "SELECT member_id FROM members WHERE member_id = %s FOR UPDATE",
                (member_id,),
            )
            if not cursor.fetchone():
                raise HTTPException(status_code=404, detail="找不到會員")
            cursor.execute(
                """
                UPDATE members
                SET member_level = %s,
                    note = %s
                WHERE member_id = %s
                """,
                (member_level, note_value or None, member_id),
            )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

    return templates.TemplateResponse(
        request=request,
        name="member_detail.html",
        context={**member_detail_context(member_id), "message": "會員資料已更新。"},
    )


# =========================================================
# 利潤報表 / 每月匯率設定 / 匯出
# =========================================================


def _profit_decimal(value, default="0"):
    try:
        return Decimal(str(value if value not in (None, "") else default))
    except (InvalidOperation, ValueError, TypeError):
        return Decimal(default)


def _profit_parse_date(value: str, fallback):
    try:
        return datetime.strptime(str(value or ""), "%Y-%m-%d").date()
    except ValueError:
        return fallback


def _profit_valid_month(value: str, fallback=None):
    value = str(value or "").strip()
    try:
        return datetime.strptime(value, "%Y-%m").strftime("%Y-%m")
    except ValueError:
        return fallback


def ensure_profit_monthly_rates_table():
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS profit_monthly_rates (
                    rate_month CHAR(7) NOT NULL PRIMARY KEY,
                    rmb_rate DECIMAL(10,4) NOT NULL,
                    payment_sell_rate DECIMAL(10,4) NOT NULL,
                    purchase_sell_rate DECIMAL(10,4) NOT NULL,
                    updated_by VARCHAR(100) NULL,
                    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
                        ON UPDATE CURRENT_TIMESTAMP
                ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci
                """
            )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _profit_rate_history(limit=36):
    ensure_profit_monthly_rates_table()
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                SELECT rate_month, rmb_rate, payment_sell_rate,
                       purchase_sell_rate, updated_by, updated_at
                FROM profit_monthly_rates
                ORDER BY rate_month DESC
                LIMIT %s
                """,
                (int(limit),),
            )
            return cursor.fetchall()
    finally:
        conn.close()


def _profit_rates_for_months(months):
    months = [m for m in dict.fromkeys(months) if _profit_valid_month(m)]
    if not months:
        return {}
    ensure_profit_monthly_rates_table()
    placeholders = ",".join(["%s"] * len(months))
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT rate_month, rmb_rate, payment_sell_rate, purchase_sell_rate,
                       updated_by, updated_at
                FROM profit_monthly_rates
                WHERE rate_month IN ({placeholders})
                """,
                months,
            )
            rows = cursor.fetchall()
    finally:
        conn.close()
    return {str(row["rate_month"]): row for row in rows}


def _profit_months_between(start_date, end_date):
    result = []
    cursor = start_date.replace(day=1)
    end_month = end_date.replace(day=1)
    while cursor <= end_month:
        result.append(cursor.strftime("%Y-%m"))
        if cursor.month == 12:
            cursor = cursor.replace(year=cursor.year + 1, month=1)
        else:
            cursor = cursor.replace(month=cursor.month + 1)
    return result


def _profit_date_bounds():
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                SELECT MIN(order_time) AS min_date, MAX(order_time) AS max_date
                FROM orders
                WHERE order_time IS NOT NULL
                """
            )
            row = cursor.fetchone() or {}
    finally:
        conn.close()
    return row.get("min_date"), row.get("max_date")


def _profit_rows(start_date, end_date, rates_by_month):
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                SELECT
                    order_id, order_time, customer_name, platform, tracking_number,
                    amount_rmb, service_fee, weight_kg, is_arrived, is_returned,
                    is_early_returned, remarks, order_status
                FROM orders
                WHERE order_time >= %s AND order_time <= %s
                ORDER BY order_time DESC, order_id DESC
                """,
                (start_date, end_date),
            )
            rows = cursor.fetchall()
    finally:
        conn.close()

    result = []
    for row in rows:
        amount = _profit_decimal(row.get("amount_rmb"))
        fee = _profit_decimal(row.get("service_fee")).quantize(Decimal("0.01"))
        weight = _profit_decimal(row.get("weight_kg"))
        order_time = row.get("order_time")
        month_key = order_time.strftime("%Y-%m") if order_time else ""
        rate_row = rates_by_month.get(month_key)
        is_payment = str(row.get("customer_name") or "").strip() == "代付"

        rmb_rate = None
        sell_rate = None
        rate_profit = None
        total_profit = None
        if rate_row:
            rmb_rate = _profit_decimal(rate_row.get("rmb_rate"))
            sell_rate = _profit_decimal(
                rate_row.get("payment_sell_rate") if is_payment
                else rate_row.get("purchase_sell_rate")
            )
            rate_profit = (amount * (sell_rate - rmb_rate)).quantize(Decimal("0.01"))
            total_profit = (rate_profit + fee).quantize(Decimal("0.01"))

        item = dict(row)
        item.update({
            "rate_month": month_key,
            "rate_missing": rate_row is None,
            "order_type": "代付" if is_payment else "代購",
            "amount_rmb_num": amount,
            "service_fee_num": fee,
            "weight_num": weight,
            "rmb_rate_num": rmb_rate,
            "sell_rate_num": sell_rate,
            "rate_profit_num": rate_profit,
            "total_profit_num": total_profit,
        })
        result.append(item)
    return result


def _profit_context(start_date="", end_date="", rate_month="", message="", error=""):
    taiwan_today = datetime.now(ZoneInfo("Asia/Taipei")).date()
    min_date, max_date = _profit_date_bounds()
    min_date = min_date or taiwan_today
    max_date = max_date or taiwan_today
    month_start = taiwan_today.replace(day=1)
    default_start = max(month_start, min_date)
    default_end = min(taiwan_today, max_date)
    if default_start > default_end:
        default_start = min_date
        default_end = max_date

    start = _profit_parse_date(start_date, default_start)
    end = _profit_parse_date(end_date, default_end)
    if start > end:
        start, end = end, start

    needed_months = _profit_months_between(start, end)
    rates_by_month = _profit_rates_for_months(needed_months)
    rows = _profit_rows(start, end, rates_by_month)
    months_with_orders = sorted({r["rate_month"] for r in rows if r.get("rate_month")})
    missing_months = [m for m in months_with_orders if m not in rates_by_month]
    can_calculate = not missing_months

    completed_rows = [r for r in rows if not r["rate_missing"]]
    payment_rows_complete = [r for r in completed_rows if r["order_type"] == "代付"]
    purchase_rows_complete = [r for r in completed_rows if r["order_type"] == "代購"]

    rate_profit = sum((r["rate_profit_num"] for r in completed_rows), Decimal("0"))
    service_fee = sum((r["service_fee_num"] for r in rows), Decimal("0"))
    total_profit = sum((r["total_profit_num"] for r in completed_rows), Decimal("0"))
    payment_profit = sum((r["total_profit_num"] for r in payment_rows_complete), Decimal("0"))
    purchase_profit = sum((r["total_profit_num"] for r in purchase_rows_complete), Decimal("0"))

    history = _profit_rate_history()
    history_map = {str(r["rate_month"]): r for r in history}
    chosen_month = _profit_valid_month(rate_month, taiwan_today.strftime("%Y-%m"))
    chosen_rate = history_map.get(chosen_month)

    return {
        "rows": rows[:500],
        "all_rows": rows,
        "total_count": len(rows),
        "has_more": len(rows) > 500,
        "start_date": start.isoformat(),
        "end_date": end.isoformat(),
        "min_date": min_date.isoformat(),
        "max_date": max_date.isoformat(),
        "rate_profit": float(rate_profit),
        "service_fee": float(service_fee),
        "total_profit": float(total_profit),
        "payment_count": len([r for r in rows if r["order_type"] == "代付"]),
        "purchase_count": len([r for r in rows if r["order_type"] == "代購"]),
        "payment_profit": float(payment_profit),
        "purchase_profit": float(purchase_profit),
        "can_calculate": can_calculate,
        "missing_months": missing_months,
        "used_months": months_with_orders,
        "rate_history": history,
        "rate_month": chosen_month,
        "selected_rate": chosen_rate,
        "message": str(message or ""),
        "error": str(error or ""),
    }


@app.get("/profit")
def profit_page(
    request: Request,
    start_date: str = "",
    end_date: str = "",
    rate_month: str = "",
    message: str = "",
    error: str = "",
    admin: str = Depends(verify_admin),
):
    context = _profit_context(start_date, end_date, rate_month, message, error)
    return templates.TemplateResponse(request=request, name="profit.html", context=context)


@app.post("/profit/rates/save")
def profit_rate_save(
    request: Request,
    rate_month: str = Form(""),
    rmb_rate: str = Form(""),
    payment_sell_rate: str = Form(""),
    purchase_sell_rate: str = Form(""),
    start_date: str = Form(""),
    end_date: str = Form(""),
    admin: str = Depends(verify_admin),
):
    _check_same_origin(request)
    month_key = _profit_valid_month(rate_month)
    if not month_key:
        raise HTTPException(status_code=422, detail="月份格式不正確")

    values = {
        "人民幣匯率": _profit_decimal(rmb_rate, "-1"),
        "代付定價匯率": _profit_decimal(payment_sell_rate, "-1"),
        "代購定價匯率": _profit_decimal(purchase_sell_rate, "-1"),
    }
    if any(v <= 0 for v in values.values()):
        raise HTTPException(status_code=422, detail="三個匯率都必須大於 0")
    if any(v > Decimal("20") for v in values.values()):
        raise HTTPException(status_code=422, detail="匯率數值看起來不合理，請確認後再儲存")

    ensure_profit_monthly_rates_table()
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO profit_monthly_rates
                    (rate_month, rmb_rate, payment_sell_rate, purchase_sell_rate, updated_by)
                VALUES (%s, %s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE
                    rmb_rate = VALUES(rmb_rate),
                    payment_sell_rate = VALUES(payment_sell_rate),
                    purchase_sell_rate = VALUES(purchase_sell_rate),
                    updated_by = VALUES(updated_by),
                    updated_at = CURRENT_TIMESTAMP
                """,
                (
                    month_key,
                    values["人民幣匯率"],
                    values["代付定價匯率"],
                    values["代購定價匯率"],
                    admin,
                ),
            )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

    qs = (
        f"rate_month={quote(month_key)}"
        f"&start_date={quote(str(start_date or ''))}"
        f"&end_date={quote(str(end_date or ''))}"
        f"&message={quote(month_key + ' 匯率設定已儲存。')}"
    )
    return RedirectResponse(url=f"/profit?{qs}", status_code=303)


@app.post("/profit/export")
def profit_export(
    request: Request,
    start_date: str = Form(""),
    end_date: str = Form(""),
    rate_month: str = Form(""),
    admin: str = Depends(verify_admin),
):
    _check_same_origin(request)
    context = _profit_context(start_date, end_date, rate_month)
    if context["missing_months"]:
        context["error"] = "請先設定以下月份的匯率，再匯出：" + "、".join(context["missing_months"])
        return templates.TemplateResponse(request=request, name="profit.html", context=context, status_code=422)

    rows = context["all_rows"]
    rate_records = _profit_rates_for_months(context["used_months"])
    data = make_profit_xlsx(
        rows,
        context["start_date"],
        context["end_date"],
        rate_records,
    )
    filename = f"代購利潤報表_{context['start_date'].replace('-', '')}_{context['end_date'].replace('-', '')}.xlsx"
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
    )


# =========================================================
# 前台設定 / 公告管理
# =========================================================

FRONTEND_DELIVERY_TYPES = {
    "home_delivery": "宅配",
    "shop_delivery": "賣貨便",
}


def ensure_frontend_config_tables_admin():
    """建立 / 補齊前台設定與船班資料表。"""
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS site_settings (
                    setting_key VARCHAR(100) PRIMARY KEY,
                    setting_value TEXT NULL,
                    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
                        ON UPDATE CURRENT_TIMESTAMP
                ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS shipping_batches (
                    batch_id INT AUTO_INCREMENT PRIMARY KEY,
                    batch_text VARCHAR(255) NOT NULL,
                    delivery_type VARCHAR(30) NOT NULL DEFAULT 'home_delivery',
                    sort_order INT NOT NULL DEFAULT 0,
                    is_active TINYINT(1) NOT NULL DEFAULT 1,
                    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
                        ON UPDATE CURRENT_TIMESTAMP,
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
                ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci
                """
            )

            cursor.execute("SHOW COLUMNS FROM shipping_batches")
            existing = {row["Field"] for row in cursor.fetchall()}
            additions = {
                "delivery_type": "VARCHAR(30) NOT NULL DEFAULT 'home_delivery'",
                "sort_order": "INT NOT NULL DEFAULT 0",
                "is_active": "TINYINT(1) NOT NULL DEFAULT 1",
                "updated_at": "TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP",
                "created_at": "TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP",
            }
            for name, ddl in additions.items():
                if name not in existing:
                    cursor.execute(f"ALTER TABLE shipping_batches ADD COLUMN {name} {ddl}")

            cursor.execute(
                """
                INSERT IGNORE INTO site_settings (setting_key, setting_value)
                VALUES ('current_exchange_rate', '4.78')
                """
            )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def frontend_settings_context(message="", error=""):
    ensure_frontend_config_tables_admin()
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                SELECT setting_key, setting_value, updated_at
                FROM site_settings
                WHERE setting_key IN ('orders_last_update_time', 'current_exchange_rate')
                """
            )
            setting_rows = cursor.fetchall()
            settings = {row["setting_key"]: row for row in setting_rows}

            cursor.execute(
                """
                SELECT batch_id, batch_text, delivery_type,
                       sort_order, is_active, updated_at, created_at
                FROM shipping_batches
                ORDER BY delivery_type ASC,
                         is_active DESC,
                         sort_order ASC,
                         batch_id DESC
                """
            )
            batches = cursor.fetchall()
    finally:
        conn.close()

    rate_value = settings.get("current_exchange_rate", {}).get("setting_value")
    try:
        current_rate = float(rate_value) if rate_value is not None else 4.78
    except (TypeError, ValueError):
        current_rate = 4.78

    for row in batches:
        row["delivery_label"] = FRONTEND_DELIVERY_TYPES.get(
            row.get("delivery_type"), "宅配"
        )
        row["is_active_bool"] = bool(row.get("is_active"))

    return {
        "orders_last_update_time": settings.get("orders_last_update_time", {}).get("setting_value") or "尚未設定",
        "orders_last_update_updated_at": settings.get("orders_last_update_time", {}).get("updated_at"),
        "current_rate": current_rate,
        "rate_updated_at": settings.get("current_exchange_rate", {}).get("updated_at"),
        "batches": batches,
        "delivery_types": FRONTEND_DELIVERY_TYPES,
        "message": str(message or ""),
        "error": str(error or ""),
    }


def _frontend_redirect(message="", error=""):
    parts = []
    if message:
        parts.append("message=" + quote(str(message)))
    if error:
        parts.append("error=" + quote(str(error)))
    suffix = ("?" + "&".join(parts)) if parts else ""
    return RedirectResponse(url="/frontend-settings" + suffix, status_code=303)


@app.get("/frontend-settings")
def frontend_settings_page(
    request: Request,
    message: str = "",
    error: str = "",
    admin: str = Depends(verify_admin),
):
    return templates.TemplateResponse(
        request=request,
        name="frontend_settings.html",
        context=frontend_settings_context(message=message, error=error),
    )


@app.post("/frontend-settings/update-order-time")
def frontend_update_order_time(
    request: Request,
    admin: str = Depends(verify_admin),
):
    _check_same_origin(request)
    ensure_frontend_config_tables_admin()
    now_str = datetime.now(ZoneInfo("Asia/Taipei")).strftime("%Y/%m/%d %H:%M")
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO site_settings (setting_key, setting_value)
                VALUES ('orders_last_update_time', %s)
                ON DUPLICATE KEY UPDATE setting_value = VALUES(setting_value)
                """,
                (now_str,),
            )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return _frontend_redirect(message=f"已更新前台訂單資料時間：{now_str}")


@app.post("/frontend-settings/exchange-rate")
def frontend_save_exchange_rate(
    request: Request,
    exchange_rate: str = Form(""),
    admin: str = Depends(verify_admin),
):
    _check_same_origin(request)
    ensure_frontend_config_tables_admin()
    try:
        rate = Decimal(str(exchange_rate).strip())
    except (InvalidOperation, ValueError):
        return _frontend_redirect(error="匯率格式不正確。")
    if rate <= 0 or rate > Decimal("20"):
        return _frontend_redirect(error="前台顯示匯率必須大於 0。")
    rate_text = format(rate.quantize(Decimal("0.01")), "f")

    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO site_settings (setting_key, setting_value)
                VALUES ('current_exchange_rate', %s)
                ON DUPLICATE KEY UPDATE setting_value = VALUES(setting_value)
                """,
                (rate_text,),
            )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return _frontend_redirect(message=f"前台顯示匯率已更新為 {rate_text}。")


@app.post("/frontend-settings/batches/add")
def frontend_add_shipping_batch(
    request: Request,
    delivery_type: str = Form("home_delivery"),
    batch_text: str = Form(""),
    sort_order: int = Form(0),
    admin: str = Depends(verify_admin),
):
    _check_same_origin(request)
    ensure_frontend_config_tables_admin()
    delivery_type = str(delivery_type or "")
    batch_text = str(batch_text or "").strip()
    if delivery_type not in FRONTEND_DELIVERY_TYPES:
        return _frontend_redirect(error="運回方式不正確。")
    if not batch_text:
        return _frontend_redirect(error="請輸入船班文字。")
    if len(batch_text) > 255:
        return _frontend_redirect(error="船班文字最多 255 個字。")
    sort_order = max(0, min(int(sort_order or 0), 9999))

    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO shipping_batches
                    (batch_text, delivery_type, sort_order, is_active)
                VALUES (%s, %s, %s, 1)
                """,
                (batch_text, delivery_type, sort_order),
            )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return _frontend_redirect(message="已新增船班。")


@app.post("/frontend-settings/batches/{batch_id}/save")
def frontend_save_shipping_batch(
    request: Request,
    batch_id: int,
    delivery_type: str = Form("home_delivery"),
    batch_text: str = Form(""),
    sort_order: int = Form(0),
    is_active: Optional[str] = Form(None),
    admin: str = Depends(verify_admin),
):
    _check_same_origin(request)
    ensure_frontend_config_tables_admin()
    delivery_type = str(delivery_type or "")
    batch_text = str(batch_text or "").strip()
    if delivery_type not in FRONTEND_DELIVERY_TYPES:
        return _frontend_redirect(error="運回方式不正確。")
    if not batch_text:
        return _frontend_redirect(error="船班文字不可空白。")
    if len(batch_text) > 255:
        return _frontend_redirect(error="船班文字最多 255 個字。")
    sort_order = max(0, min(int(sort_order or 0), 9999))
    active_value = 1 if is_active == "1" else 0

    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                UPDATE shipping_batches
                SET batch_text = %s,
                    delivery_type = %s,
                    sort_order = %s,
                    is_active = %s
                WHERE batch_id = %s
                """,
                (batch_text, delivery_type, sort_order, active_value, batch_id),
            )
            if cursor.rowcount == 0:
                cursor.execute("SELECT batch_id FROM shipping_batches WHERE batch_id = %s", (batch_id,))
                if not cursor.fetchone():
                    conn.rollback()
                    return _frontend_redirect(error="找不到這筆船班。")
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return _frontend_redirect(message=f"船班 #{batch_id} 已更新。")


@app.post("/frontend-settings/batches/{batch_id}/delete")
def frontend_delete_shipping_batch(
    request: Request,
    batch_id: int,
    admin: str = Depends(verify_admin),
):
    _check_same_origin(request)
    ensure_frontend_config_tables_admin()
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute("DELETE FROM shipping_batches WHERE batch_id = %s", (batch_id,))
            deleted = cursor.rowcount
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    if not deleted:
        return _frontend_redirect(error="找不到這筆船班。")
    return _frontend_redirect(message=f"船班 #{batch_id} 已刪除。")

# =========================================================
# 集運登記管理：沿用舊 Streamlit 前台登記 -> 建立集運訂單邏輯
# =========================================================

FORWARDING_STATUS_LABELS = {
    "pending": "待處理",
    "processed": "已處理",
    "cancelled": "已取消",
}


def ensure_forwarding_register_table_admin():
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS customer_forwarding_registers (
                    register_id INT AUTO_INCREMENT PRIMARY KEY,
                    customer_name VARCHAR(255) NOT NULL,
                    tracking_number VARCHAR(255) NOT NULL,
                    item_name VARCHAR(255) NOT NULL,
                    quantity INT NOT NULL DEFAULT 1,
                    unit_price_rmb DECIMAL(10,2) NOT NULL DEFAULT 0,
                    remarks TEXT NULL,
                    status ENUM('pending','processed','cancelled') NOT NULL DEFAULT 'pending',
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                    UNIQUE KEY uk_tracking_number (tracking_number)
                ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci
                """
            )
            cursor.execute("SHOW COLUMNS FROM customer_forwarding_registers")
            existing = {row["Field"] for row in cursor.fetchall()}
            additions = {
                "quantity": "INT NOT NULL DEFAULT 1",
                "unit_price_rmb": "DECIMAL(10,2) NOT NULL DEFAULT 0",
                "remarks": "TEXT NULL",
                "status": "ENUM('pending','processed','cancelled') NOT NULL DEFAULT 'pending'",
                "created_at": "TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP",
                "updated_at": "TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP",
            }
            for name, ddl in additions.items():
                if name not in existing:
                    cursor.execute(f"ALTER TABLE customer_forwarding_registers ADD COLUMN {name} {ddl}")
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def forwarding_register_rows(status="pending", keyword=""):
    ensure_forwarding_register_table_admin()
    status = str(status or "pending")
    if status not in ("all", "pending", "processed", "cancelled"):
        status = "pending"
    keyword = str(keyword or "").strip()[:100]

    where = ["1=1"]
    params = []
    if status != "all":
        where.append("status = %s")
        params.append(status)
    if keyword:
        like = f"%{keyword}%"
        where.append(
            "(customer_name LIKE %s OR tracking_number LIKE %s OR item_name LIKE %s OR COALESCE(remarks,'') LIKE %s)"
        )
        params.extend([like, like, like, like])

    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT register_id, customer_name, tracking_number, item_name,
                       quantity, unit_price_rmb, remarks, status, created_at, updated_at
                FROM customer_forwarding_registers
                WHERE {' AND '.join(where)}
                ORDER BY created_at DESC, register_id DESC
                LIMIT 500
                """,
                params,
            )
            rows = cursor.fetchall()
            cursor.execute(
                """
                SELECT status, COUNT(*) AS cnt
                FROM customer_forwarding_registers
                GROUP BY status
                """
            )
            counts = {row["status"]: int(row["cnt"] or 0) for row in cursor.fetchall()}
    finally:
        conn.close()

    for row in rows:
        row["status_label"] = FORWARDING_STATUS_LABELS.get(row.get("status"), row.get("status") or "—")
        row["quantity_num"] = int(row.get("quantity") or 0)
        row["unit_price_num"] = float(row.get("unit_price_rmb") or 0)
    return rows, counts, status, keyword


def _forwarding_redirect(status_filter="pending", keyword="", message="", error=""):
    parts = [f"status={quote(str(status_filter or 'pending'))}"]
    if keyword:
        parts.append("keyword=" + quote(str(keyword)))
    if message:
        parts.append("message=" + quote(str(message)))
    if error:
        parts.append("error=" + quote(str(error)))
    return RedirectResponse(url="/forwarding-registers?" + "&".join(parts), status_code=303)


@app.get("/forwarding-registers")
def forwarding_registers_page(
    request: Request,
    status: str = "pending",
    keyword: str = "",
    message: str = "",
    error: str = "",
    admin: str = Depends(verify_admin),
):
    rows, counts, status, keyword = forwarding_register_rows(status, keyword)
    return templates.TemplateResponse(
        request=request,
        name="forwarding_registers.html",
        context={
            "rows": rows,
            "counts": counts,
            "status": status,
            "keyword": keyword,
            "message": str(message or ""),
            "error": str(error or ""),
            "status_labels": FORWARDING_STATUS_LABELS,
        },
    )


@app.post("/forwarding-registers/apply")
def forwarding_registers_apply(
    request: Request,
    selected_register_ids: list[int] = Form([]),
    action: str = Form(""),
    status_filter: str = Form("pending"),
    keyword_filter: str = Form(""),
    admin: str = Depends(verify_admin),
):
    _check_same_origin(request)
    ensure_forwarding_register_table_admin()
    ids = []
    seen = set()
    for raw in selected_register_ids:
        try:
            rid = int(raw)
        except (TypeError, ValueError):
            continue
        if rid > 0 and rid not in seen:
            seen.add(rid)
            ids.append(rid)
    if not ids:
        return _forwarding_redirect(status_filter, keyword_filter, error="請先勾選至少一筆集運登記。")
    if len(ids) > 100:
        return _forwarding_redirect(status_filter, keyword_filter, error="一次最多處理 100 筆集運登記。")
    if action not in ("processed", "cancelled"):
        return _forwarding_redirect(status_filter, keyword_filter, error="批次操作不正確。")

    ensure_audit_storage()
    conn = get_db()
    created_count = 0
    duplicate_count = 0
    changed_count = 0
    skipped_count = 0
    invalid_count = 0
    try:
        with conn.cursor() as cursor:
            for register_id in ids:
                cursor.execute(
                    """
                    SELECT register_id, customer_name, tracking_number, item_name,
                           quantity, unit_price_rmb, remarks, status
                    FROM customer_forwarding_registers
                    WHERE register_id = %s
                    FOR UPDATE
                    """,
                    (register_id,),
                )
                row = cursor.fetchone()
                if not row or row.get("status") != "pending":
                    skipped_count += 1
                    continue

                if action == "cancelled":
                    cursor.execute(
                        "UPDATE customer_forwarding_registers SET status='cancelled' WHERE register_id=%s",
                        (register_id,),
                    )
                    changed_count += 1
                    continue

                customer_name = str(row.get("customer_name") or "").strip()
                tracking_number = str(row.get("tracking_number") or "").strip()
                item_name = str(row.get("item_name") or "").strip()
                remarks = str(row.get("remarks") or "").strip()
                quantity = int(row.get("quantity") or 1)
                unit_price = Decimal(str(row.get("unit_price_rmb") or 0))

                if not customer_name or not tracking_number or not item_name:
                    invalid_count += 1
                    continue

                cursor.execute(
                    "SELECT order_id FROM orders WHERE tracking_number = %s LIMIT 1",
                    (tracking_number,),
                )
                existing = cursor.fetchone()

                cursor.execute(
                    "UPDATE customer_forwarding_registers SET status='processed' WHERE register_id=%s",
                    (register_id,),
                )
                changed_count += 1

                if existing:
                    duplicate_count += 1
                    continue

                auto_remarks = (
                    f"前台集運登記｜內容物：{item_name}｜數量：{quantity}｜"
                    f"單價：{format(unit_price, 'f')} RMB"
                )
                if remarks:
                    auto_remarks += f"｜備註：{remarks}"

                order_date = datetime.now(ZoneInfo("Asia/Taipei")).date()
                cursor.execute(
                    """
                    INSERT INTO orders
                    (order_time, customer_name, platform, tracking_number,
                     amount_rmb, weight_kg, is_arrived, is_returned,
                     service_fee, remarks)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                    """,
                    (
                        order_date,
                        customer_name,
                        "集運",
                        tracking_number,
                        Decimal("0"),
                        Decimal("0"),
                        0,
                        0,
                        Decimal("0"),
                        auto_remarks,
                    ),
                )
                new_order_id = int(cursor.lastrowid)
                cursor.execute(
                    "INSERT IGNORE INTO members (customer_name) VALUES (%s)",
                    (customer_name,),
                )
                new_order = fetch_locked_order(cursor, new_order_id)
                if not new_order:
                    raise RuntimeError("集運登記建立訂單後無法讀取訂單，已取消交易")
                write_audit_log(cursor, new_order_id, "CREATE", admin, None, new_order)
                created_count += 1
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

    if action == "cancelled":
        msg = f"已取消 {changed_count} 筆集運登記。"
        if skipped_count:
            msg += f" 另有 {skipped_count} 筆因狀態已變更而略過。"
        return _forwarding_redirect(status_filter, keyword_filter, message=msg)

    msg = f"已處理 {changed_count} 筆集運登記，其中新增 {created_count} 筆集運訂單。"
    if duplicate_count:
        msg += f" {duplicate_count} 筆因 orders 已有相同快遞單號，未重複新增。"
    if invalid_count:
        msg += f" {invalid_count} 筆資料不完整，仍維持待處理。"
    if skipped_count:
        msg += f" {skipped_count} 筆因狀態已變更而略過。"
    return _forwarding_redirect(status_filter, keyword_filter, message=msg)



# =========================================================
# 匿名回饋管理
# 舊版管理頁欄位：id / created_at / content / status / staff_note
# 舊 feedback_store.py 未包含在目前專案，因此新版後台改用 MySQL 持久化。
# 若資料庫已有相容的 feedbacks 表，優先沿用；否則使用 anonymous_feedbacks。
# =========================================================

FEEDBACK_STATUSES = ("未處理", "已讀", "已回覆", "忽略")
FEEDBACK_TABLE_CANDIDATES = ("anonymous_feedbacks", "feedbacks")


def _feedback_table_has_required_columns(cursor, table_name: str) -> bool:
    if table_name not in FEEDBACK_TABLE_CANDIDATES:
        return False
    cursor.execute(f"SHOW COLUMNS FROM `{table_name}`")
    columns = {row["Field"] for row in cursor.fetchall()}
    return {"id", "content", "status", "staff_note", "created_at"}.issubset(columns)


def ensure_feedback_table_admin() -> str:
    """回傳實際使用的回饋表名，優先沿用相容舊表。"""
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            for table_name in FEEDBACK_TABLE_CANDIDATES:
                cursor.execute("SHOW TABLES LIKE %s", (table_name,))
                if cursor.fetchone() and _feedback_table_has_required_columns(cursor, table_name):
                    return table_name

            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS anonymous_feedbacks (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    content TEXT NOT NULL,
                    status VARCHAR(20) NOT NULL DEFAULT '未處理',
                    staff_note TEXT NULL,
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                    INDEX idx_feedback_status (status),
                    INDEX idx_feedback_created_at (created_at)
                ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci
                """
            )
        conn.commit()
        return "anonymous_feedbacks"
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def feedback_rows(keyword="", status="all"):
    table_name = ensure_feedback_table_admin()
    keyword = str(keyword or "").strip()[:120]
    status = str(status or "all")
    if status not in ("all",) + FEEDBACK_STATUSES:
        status = "all"

    where = ["1=1"]
    params = []
    if status != "all":
        where.append("status = %s")
        params.append(status)
    if keyword:
        like = f"%{keyword}%"
        where.append("(content LIKE %s OR COALESCE(staff_note,'') LIKE %s)")
        params.extend([like, like])

    conn = get_db()
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"""
                SELECT id, created_at, content, status, staff_note
                FROM `{table_name}`
                WHERE {' AND '.join(where)}
                ORDER BY created_at DESC, id DESC
                LIMIT 500
                """,
                params,
            )
            rows = cursor.fetchall()
            cursor.execute(
                f"""
                SELECT status, COUNT(*) AS cnt
                FROM `{table_name}`
                GROUP BY status
                """
            )
            counts = {row["status"]: int(row["cnt"] or 0) for row in cursor.fetchall()}
    finally:
        conn.close()

    for row in rows:
        if row.get("status") not in FEEDBACK_STATUSES:
            row["status"] = "未處理"
    return rows, counts, keyword, status


def _feedback_redirect(status_filter="all", keyword="", message="", error=""):
    parts = [f"status={quote(str(status_filter or 'all'))}"]
    if keyword:
        parts.append("keyword=" + quote(str(keyword)))
    if message:
        parts.append("message=" + quote(str(message)))
    if error:
        parts.append("error=" + quote(str(error)))
    return RedirectResponse(url="/feedbacks?" + "&".join(parts), status_code=303)


@app.get("/feedbacks")
def feedbacks_page(
    request: Request,
    keyword: str = "",
    status: str = "all",
    message: str = "",
    error: str = "",
    admin: str = Depends(verify_admin),
):
    rows, counts, keyword, status = feedback_rows(keyword, status)
    return templates.TemplateResponse(
        request=request,
        name="feedbacks.html",
        context={
            "rows": rows,
            "counts": counts,
            "keyword": keyword,
            "status": status,
            "message": str(message or ""),
            "error": str(error or ""),
            "feedback_statuses": FEEDBACK_STATUSES,
        },
    )


@app.post("/feedbacks/apply")
def feedbacks_apply(
    request: Request,
    selected_feedback_ids: list[int] = Form([]),
    new_status: str = Form(""),
    staff_note: str = Form(""),
    status_filter: str = Form("all"),
    keyword_filter: str = Form(""),
    admin: str = Depends(verify_admin),
):
    _check_same_origin(request)
    table_name = ensure_feedback_table_admin()

    ids = []
    seen = set()
    for raw in selected_feedback_ids:
        try:
            fid = int(raw)
        except (TypeError, ValueError):
            continue
        if fid > 0 and fid not in seen:
            seen.add(fid)
            ids.append(fid)

    if not ids:
        return _feedback_redirect(status_filter, keyword_filter, error="請先勾選至少一筆匿名回饋。")
    if len(ids) > 100:
        return _feedback_redirect(status_filter, keyword_filter, error="一次最多處理 100 筆匿名回饋。")
    if new_status not in ("已讀", "已回覆", "忽略"):
        return _feedback_redirect(status_filter, keyword_filter, error="請選擇要套用的狀態。")

    note_value = str(staff_note or "").strip()
    if len(note_value) > 65535:
        return _feedback_redirect(status_filter, keyword_filter, error="備註內容過長。")

    placeholders = ",".join(["%s"] * len(ids))
    conn = get_db()
    try:
        with conn.cursor() as cursor:
            # 舊版邏輯：有輸入備註就覆蓋 staff_note；未輸入則只改狀態。
            if note_value:
                cursor.execute(
                    f"UPDATE `{table_name}` SET status=%s, staff_note=%s WHERE id IN ({placeholders})",
                    [new_status, note_value] + ids,
                )
            else:
                cursor.execute(
                    f"UPDATE `{table_name}` SET status=%s WHERE id IN ({placeholders})",
                    [new_status] + ids,
                )
            changed = int(cursor.rowcount or 0)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

    return _feedback_redirect(
        status_filter,
        keyword_filter,
        message=f"已將 {changed} 筆回饋更新為「{new_status}」。",
    )
