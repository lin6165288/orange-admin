import os
import json
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

    unreturned=""
):

    conditions = []

    params = []


    # =====================================================
    # 訂單編號：精準搜尋
    # =====================================================

    search_order_id = str(
        search_order_id or ""
    ).strip()

    if search_order_id:

        try:

            order_id_value = int(
                search_order_id
            )

        except ValueError:

            return {
                "orders": [],
                "total_count": 0,
                "total_weight": 0
            }


        conditions.append(
            "order_id = %s"
        )

        params.append(
            order_id_value
        )


    # =====================================================
    # 姓名：部分搜尋
    # =====================================================

    search_customer_name = (
        search_customer_name or ""
    ).strip()

    if search_customer_name:

        conditions.append(
            "customer_name LIKE %s"
        )

        params.append(
            f"%{search_customer_name}%"
        )


    # =====================================================
    # 物流單號：部分搜尋
    # =====================================================

    search_tracking_number = (
        search_tracking_number or ""
    ).strip()

    if search_tracking_number:

        conditions.append(
            "tracking_number LIKE %s"
        )

        params.append(
            f"%{search_tracking_number}%"
        )


    # =====================================================
    # 人民幣金額：精準搜尋
    # =====================================================

    search_amount_rmb = str(
        search_amount_rmb or ""
    ).strip()

    if search_amount_rmb:

        try:

            amount_value = Decimal(
                search_amount_rmb
            )

        except InvalidOperation:

            return {
                "orders": [],
                "total_count": 0,
                "total_weight": 0
            }


        conditions.append(
            "amount_rmb = %s"
        )

        params.append(
            amount_value
        )


    # =====================================================
    # 平台：部分搜尋
    # =====================================================

    search_platform = (
        search_platform or ""
    ).strip()

    if search_platform:

        conditions.append(
            "platform LIKE %s"
        )

        params.append(
            f"%{search_platform}%"
        )


    # =====================================================
    # 下單日期
    # =====================================================

    search_order_date = (
        search_order_date or ""
    ).strip()

    if search_order_date:

        conditions.append(
            "order_time = %s"
        )

        params.append(
            search_order_date
        )


    # =====================================================
    # 只看未運回
    # =====================================================

    if unreturned == "1":

        conditions.append(
            "COALESCE(is_returned, 0) = 0"
        )

        conditions.append(
            "COALESCE(order_status, '正常') <> '取消'"
        )


    # =====================================================
    # 沒有搜尋條件
    # =====================================================

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


            # =================================================
            # Stats
            # =================================================

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
                )
                or 0
            )


            # =================================================
            # Orders
            # =================================================

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

                ORDER BY
                    order_id DESC

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

    ensure_audit_table(
        cursor
    )


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

    search_order_id: str = "",

    search_customer_name: str = "",

    search_tracking_number: str = "",

    search_amount_rmb: str = "",

    search_platform: str = "",

    search_order_date: str = "",

    unreturned: str = "",

    admin: str = Depends(
        verify_admin
    )
):

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

            "unreturned":
                unreturned
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

    search_order_id: str = "",

    search_customer_name: str = "",

    search_tracking_number: str = "",

    search_amount_rmb: str = "",

    search_platform: str = "",

    search_order_date: str = "",

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

    unreturned: str = Form(
        ""
    ),

    admin: str = Depends(
        verify_admin
    )
):

    # =====================================================
    # Old Order
    # =====================================================

    old_order = fetch_order(
        order_id
    )


    if not old_order:

        raise HTTPException(
            status_code=404,
            detail="找不到訂單"
        )


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

    conn = get_db()


    try:

        with conn.cursor() as cursor:


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


            # Fetch new version
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


            new_order = (
                cursor.fetchone()
            )


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

    unreturned: str = Form(
        ""
    ),

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


    conn = get_db()


    try:

        with conn.cursor() as cursor:


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

            "unreturned":
                unreturned
        }
    )


# =========================================================
# Audit Logs
# =========================================================

@app.get(
    "/audit-logs"
)
async def audit_logs_page(

    request: Request,

    admin: str = Depends(
        verify_admin
    )
):

    conn = get_db()


    try:

        with conn.cursor() as cursor:


            ensure_audit_table(
                cursor
            )


            cursor.execute(
                """
                SELECT
                    id,
                    order_id,
                    action,
                    admin_username,
                    before_data,
                    after_data,
                    created_at

                FROM order_audit_logs

                ORDER BY
                    id DESC

                LIMIT 300
                """
            )


            rows = cursor.fetchall()


        conn.commit()


    finally:

        conn.close()


    logs = []


    for row in rows:


        before_data = (
            parse_audit_json(
                row["before_data"]
            )
        )


        after_data = (
            parse_audit_json(
                row["after_data"]
            )
        )


        changes = []


        if row["action"] == "UPDATE":


            for (
                field,
                label
            ) in AUDIT_FIELD_LABELS.items():


                before_value = (
                    before_data.get(
                        field
                    )
                )


                after_value = (
                    after_data.get(
                        field
                    )
                )


                if str(
                    before_value
                ) != str(
                    after_value
                ):

                    changes.append(
                        {
                            "field":
                                field,

                            "label":
                                label,

                            "before":
                                display_audit_value(
                                    before_value
                                ),

                            "after":
                                display_audit_value(
                                    after_value
                                )
                        }
                    )


        logs.append(
            {
                "id":
                    row["id"],

                "order_id":
                    row["order_id"],

                "action":
                    row["action"],

                "admin_username":
                    row["admin_username"],

                "created_at":
                    row["created_at"],

                "before":
                    before_data,

                "after":
                    after_data,

                "changes":
                    changes
            }
        )


    return templates.TemplateResponse(
        request=request,
        name="audit_logs.html",
        context={
            "logs":
                logs
        }
    )
