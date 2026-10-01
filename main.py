import json
import os
import traceback
from collections import Counter
from datetime import datetime, timedelta, timezone

import gspread
import vk_api
from google.oauth2.service_account import Credentials

GROUP_ID = int(os.environ["GROUP_ID"].strip())

# Google Sheets
SPREADSHEET_ID = os.environ["SPREADSHEET_ID"].strip()
SHEET_NAME = os.environ.get("SHEET_NAME", "Рейтинг").strip()
GOOGLE_CREDENTIALS = json.loads(os.environ["GOOGLE_CREDENTIALS"])  # JSON сервисного аккаунта

W_LIKE = 1
W_COMMENT = 2

POSTS_PAGE_SIZE = 100

# Сколько участников выгружать в таблицу (None = всех)
MAX_ROWS = 300

# Показывать фото в первом столбце (формула =IMAGE)
SHOW_PHOTO = True

MOSCOW_TZ = timezone(timedelta(hours=3))

RATING_FROM = datetime(2026, 9, 1, 0, 0, 0, tzinfo=MOSCOW_TZ)
RATING_FROM_TS = int(RATING_FROM.timestamp())

vk = vk_api.VkApi(token=os.environ["SERVICE_TOKEN"].strip()).get_api()


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


def calc():
    likes = Counter()
    comments = Counter()

    print("Дата начала рейтинга:", RATING_FROM.strftime("%d.%m.%Y %H:%M:%S"), "МСК")

    posts = get_posts_since(RATING_FROM_TS)

    print(f"Постов найдено: {len(posts)}")

    for p in posts:
        post_id = p["id"]

        for uid in paged(
            vk.likes.getList,
            1000,
            type="post",
            owner_id=-GROUP_ID,
            item_id=post_id
        ):
            likes[uid] += 1

        for c in paged(
            vk.wall.getComments,
            100,
            owner_id=-GROUP_ID,
            post_id=post_id,
            thread_items_count=10
        ):
            if c.get("from_id", 0) > 0:
                comments[c["from_id"]] += 1

            for t in c.get("thread", {}).get("items", []):
                if t.get("from_id", 0) > 0:
                    comments[t["from_id"]] += 1

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

    print(f"Участников с баллами: {len(rows)}")

    return rows


def load_users(uids):
    """Имена и фото через users.get (максимум 1000 id за запрос)."""
    users = {}

    for i in range(0, len(uids), 1000):
        chunk = uids[i:i + 1000]

        for u in vk.users.get(
            user_ids=chunk,
            fields="photo_100",
            lang=0
        ):
            users[u["id"]] = u

    return users


def build_table(rows):
    if MAX_ROWS:
        rows = rows[:MAX_ROWS]

    users = load_users([r[0] for r in rows])

    header = []
    if SHOW_PHOTO:
        header.append("Фото")
    header += ["Участник", "#", "❤️", "💬", "Баллы"]

    table = [header]

    for place, (uid, likes_count, comments_count, points) in enumerate(rows, 1):
        user = users.get(uid)

        if user:
            name = f'{user["first_name"]} {user["last_name"]}'
        else:
            name = f"id{uid}"

        name = name.replace('"', '""')

        row = []

        if SHOW_PHOTO:
            photo = (user or {}).get("photo_100")
            row.append(f'=IMAGE("{photo}")' if photo else "")

        row += [
            f'=HYPERLINK("https://vk.com/id{uid}","{name}")',
            place,
            likes_count,
            comments_count,
            points
        ]

        table.append(row)

    return table


def get_worksheet():
    creds = Credentials.from_service_account_info(
        GOOGLE_CREDENTIALS,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )

    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)

    try:
        return sh.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(title=SHEET_NAME, rows=1000, cols=10)


def write_to_sheet(table, updated_at):
    ws = get_worksheet()

    ws.clear()

    # Строка 1 — заголовок с временем обновления, таблица со строки 3
    ws.update(
        range_name="A1",
        values=[[f"Рейтинг на {updated_at}"]]
    )

    ws.update(
        range_name="A3",
        values=table,
        value_input_option="USER_ENTERED"
    )

    ncols = len(table[0])
    last_col = chr(ord("A") + ncols - 1)

    ws.format("A1", {"textFormat": {"bold": True, "fontSize": 14}})
    ws.format(
        f"A3:{last_col}3",
        {
            "textFormat": {"bold": True},
            "horizontalAlignment": "CENTER",
            "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}
        }
    )

    first_num_col = chr(ord("A") + (2 if SHOW_PHOTO else 1))
    ws.format(
        f"{first_num_col}4:{last_col}{len(table) + 2}",
        {"horizontalAlignment": "CENTER"}
    )

    # Высота строк под фото
    if SHOW_PHOTO and len(table) > 1:
        ws.spreadsheet.batch_update({
            "requests": [{
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": ws.id,
                        "dimension": "ROWS",
                        "startIndex": 3,
                        "endIndex": len(table) + 2
                    },
                    "properties": {"pixelSize": 60},
                    "fields": "pixelSize"
                }
            }]
        })

    ws.freeze(rows=3)


def main():
    updated_at = datetime.now(MOSCOW_TZ).strftime("%H:%M %d.%m.%Y")

    try:
        rows = calc()

        if not rows:
            table = [["Участник", "#", "❤️", "💬", "Баллы"]]
            print("Пока нет активности")
        else:
            table = build_table(rows)

        write_to_sheet(table, updated_at)

        print("Таблица обновлена")

    except Exception:
        traceback.print_exc()
        raise SystemExit("Не удалось обновить таблицу")


if __name__ == "__main__":
    main()
