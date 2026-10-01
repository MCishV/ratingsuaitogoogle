import json
import os
import traceback
from collections import Counter
from datetime import datetime, timedelta, timezone

import gspread
import vk_api
from google.oauth2.service_account import Credentials

GROUP_ID = int(os.environ["GROUP_ID"].strip())

SPREADSHEET_ID = os.environ["SPREADSHEET_ID"].strip()
GOOGLE_CREDENTIALS = json.loads(os.environ["GOOGLE_CREDENTIALS"])

# Названия листов (создаются, если их нет; иначе обновляются)
SHEET_YEAR = "Учебный год"
SHEET_MONTH = "Текущий месяц"

W_LIKE = 1
W_COMMENT = 2

POSTS_PAGE_SIZE = 100

# Сколько участников выгружать (None = всех)
MAX_ROWS = 300

COL_WIDTH_PLACE = 50
COL_WIDTH_NAME = 300
COL_WIDTH_NUM = 80

FORMULA_SEP = ";"

MOSCOW_TZ = timezone(timedelta(hours=3))

MONTHS_RU = [
    "январь", "февраль", "март", "апрель", "май", "июнь",
    "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь",
]

vk = vk_api.VkApi(token=os.environ["SERVICE_TOKEN"].strip()).get_api()


def period_bounds(now):
    """Начало учебного года (1 сентября) и начало текущего месяца, по МСК."""
    start_year = now.year if now.month >= 9 else now.year - 1

    year_from = datetime(start_year, 9, 1, tzinfo=MOSCOW_TZ)
    month_from = datetime(now.year, now.month, 1, tzinfo=MOSCOW_TZ)

    return year_from, month_from


def paged(method, per_page, **kw):
    offset = 0

    while True:
        r = method(count=per_page, offset=offset, **kw)

        yield from r["items"]

        offset += per_page

        if offset >= r["count"]:
            break


def get_posts_since(ts):
    posts = []
    offset = 0

    while True:
        r = vk.wall.get(
            owner_id=-GROUP_ID,
            count=POSTS_PAGE_SIZE,
            offset=offset
        )

        batch = r.get("items", [])

        if not batch:
            break

        stop = False

        for p in batch:
            post_ts = int(p.get("date", 0))
            post_dt = datetime.fromtimestamp(post_ts, MOSCOW_TZ)
            is_pinned = p.get("is_pinned", 0) == 1

            print(
                f'Пост {p["id"]}: '
                f'{post_dt.strftime("%d.%m.%Y %H:%M:%S")} МСК'
                f'{" [ЗАКРЕПЛЁН]" if is_pinned else ""}'
            )

            if is_pinned:
                continue

            if post_ts < ts:
                stop = True
                break

            posts.append(p)

        if stop:
            break

        offset += POSTS_PAGE_SIZE

        if offset >= r.get("count", 0):
            break

    return posts


def make_rows(likes, comments):
    rows = [
        (
            uid,
            likes[uid],
            comments[uid],
            likes[uid] * W_LIKE + comments[uid] * W_COMMENT
        )
        for uid in set(likes) | set(comments)
    ]

    rows.sort(key=lambda r: (-r[3], -r[2], r[0]))

    return rows


def calc(year_from, month_from):
    """Возвращает (рейтинг за учебный год, рейтинг за месяц).

    Месяц считается по дате публикации поста: учитываются лайки и комментарии
    под постами, опубликованными с 1-го числа текущего месяца.
    """
    year_likes, year_comments = Counter(), Counter()
    month_likes, month_comments = Counter(), Counter()

    month_ts = int(month_from.timestamp())

    print("Начало учебного года:", year_from.strftime("%d.%m.%Y"), "МСК")
    print("Начало месяца:", month_from.strftime("%d.%m.%Y"), "МСК")

    posts = get_posts_since(int(year_from.timestamp()))

    print(f"Постов найдено: {len(posts)}")

    for p in posts:
        post_id = p["id"]
        in_month = int(p.get("date", 0)) >= month_ts

        for uid in paged(
            vk.likes.getList,
            1000,
            type="post",
            owner_id=-GROUP_ID,
            item_id=post_id
        ):
            year_likes[uid] += 1

            if in_month:
                month_likes[uid] += 1

        for c in paged(
            vk.wall.getComments,
            100,
            owner_id=-GROUP_ID,
            post_id=post_id,
            thread_items_count=10
        ):
            authors = []

            if c.get("from_id", 0) > 0:
                authors.append(c["from_id"])

            for t in c.get("thread", {}).get("items", []):
                if t.get("from_id", 0) > 0:
                    authors.append(t["from_id"])

            for uid in authors:
                year_comments[uid] += 1

                if in_month:
                    month_comments[uid] += 1

    rows_year = make_rows(year_likes, year_comments)
    rows_month = make_rows(month_likes, month_comments)

    print(f"Участников за учебный год: {len(rows_year)}")
    print(f"Участников за месяц: {len(rows_month)}")

    return rows_year, rows_month


def load_users(uids):
    uids = list(uids)
    users = {}

    for i in range(0, len(uids), 1000):
        for u in vk.users.get(user_ids=uids[i:i + 1000], lang=0):
            users[u["id"]] = u

    return users


HEADER = ["#", "Участник", "❤️", "💬", "Баллы"]


def limit_rows(rows):
    return rows[:MAX_ROWS] if MAX_ROWS else rows


def build_table(rows, users):
    table = [HEADER]

    for place, (uid, likes_count, comments_count, points) in enumerate(limit_rows(rows), 1):
        user = users.get(uid)

        if user:
            name = f'{user["first_name"]} {user["last_name"]}'
        else:
            name = f"id{uid}"

        name = name.replace('"', '""')

        table.append([
            place,
            f'=HYPERLINK("https://vk.com/id{uid}"{FORMULA_SEP}"{name}")',
            likes_count,
            comments_count,
            points
        ])

    return table


def set_col_width(ws, col_index, pixels):
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId": ws.id,
                "dimension": "COLUMNS",
                "startIndex": col_index,
                "endIndex": col_index + 1
            },
            "properties": {"pixelSize": pixels},
            "fields": "pixelSize"
        }
    }


def open_spreadsheet():
    creds = Credentials.from_service_account_info(
        GOOGLE_CREDENTIALS,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )

    return gspread.authorize(creds).open_by_key(SPREADSHEET_ID)


def get_worksheet(sh, name):
    try:
        return sh.worksheet(name)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(title=name, rows=1000, cols=10)


def write_to_sheet(sh, sheet_name, title, table):
    ws = get_worksheet(sh, sheet_name)

    ws.clear()

    ws.update(range_name="A1", values=[[title]])

    ws.update(
        range_name="A3",
        values=table,
        value_input_option="USER_ENTERED"
    )

    last_row = len(table) + 2

    ws.format("A1", {"textFormat": {"bold": True, "fontSize": 14}})
    ws.format(
        "A3:E3",
        {
            "textFormat": {"bold": True},
            "horizontalAlignment": "CENTER",
            "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}
        }
    )

    # Номер и числа по центру, имя слева
    ws.format(f"A4:A{last_row}", {"horizontalAlignment": "CENTER"})
    ws.format(f"C4:E{last_row}", {"horizontalAlignment": "CENTER"})

    # Ширина: A узкий, B широкий, C–E средние
    sh.batch_update({
        "requests": [
            set_col_width(ws, 0, COL_WIDTH_PLACE),
            set_col_width(ws, 1, COL_WIDTH_NAME),
            set_col_width(ws, 2, COL_WIDTH_NUM),
            set_col_width(ws, 3, COL_WIDTH_NUM),
            set_col_width(ws, 4, COL_WIDTH_NUM),
        ]
    })

    ws.freeze(rows=3)


def main():
    now = datetime.now(MOSCOW_TZ)
    updated_at = now.strftime("%H:%M %d.%m.%Y")

    year_from, month_from = period_bounds(now)

    try:
        rows_year, rows_month = calc(year_from, month_from)

        # Имена грузим один раз для обоих рейтингов
        uids = {r[0] for r in limit_rows(rows_year)} | {r[0] for r in limit_rows(rows_month)}
        users = load_users(uids)

        table_year = build_table(rows_year, users)
        table_month = build_table(rows_month, users)

        title_year = (
            f"Рейтинг за учебный год {year_from.year}/{year_from.year + 1} "
            f"(с 01.09.{year_from.year}) на {updated_at}"
        )
        title_month = (
            f"Рейтинг за {MONTHS_RU[now.month - 1]} {now.year} на {updated_at}"
        )

        sh = open_spreadsheet()

        write_to_sheet(sh, SHEET_YEAR, title_year, table_year)
        write_to_sheet(sh, SHEET_MONTH, title_month, table_month)

        print("Таблица обновлена")

    except Exception:
        traceback.print_exc()
        raise SystemExit("Не удалось обновить таблицу")


if __name__ == "__main__":
    main()
