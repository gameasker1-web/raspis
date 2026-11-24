"""
Telegram-бот расписания. Многопользовательский:
  /start -> меню: [Расписание] [Подписка вкл/выкл] [Выбрать группу]
Каждый пользователь выбирает свою группу из списка (по номеру: 480, 481, ...,
суффикс вида «РУ 26» не учитывается) и может подписаться на уведомления:
как только на Яндекс Диске появляется новый/изменённый файл расписания,
бот присылает подписчикам расписание их группы.

Установка:  pip install requests
Запуск:     BOT_TOKEN=... python bot.py
Необязательно: PUBLIC_KEY (ссылка на Диск), CHECK_EVERY (сек, по умолчанию 600),
               DATA_FILE (где хранить пользователей, по умолчанию data.json)
"""
import io
import json
import os
import re
import time
import traceback
import xml.etree.ElementTree as ET
import zipfile

import requests


def log(text):
    print(f"[{time.strftime('%d.%m %H:%M:%S')}] {text}", flush=True)

BOT_TOKEN = os.environ["BOT_TOKEN"]
PUBLIC_KEY = os.environ.get("PUBLIC_KEY", "https://disk.yandex.ru/d/g-E9DJTF1kYO2A")
CHECK_EVERY = int(os.environ.get("CHECK_EVERY", "600"))
CACHE_TTL = 120  # как часто (сек) по нажатию кнопки перепроверять Диск
DATA_FILE = os.environ.get("DATA_FILE", "data.json")

YA_API = "https://cloud-api.yandex.net/v1/disk/public/resources"
TG_API = f"https://api.telegram.org/bot{BOT_TOKEN}/"


# ---------------------------------------------------------------- Telegram
class TgError(Exception):
    def __init__(self, code, description):
        super().__init__(f"{code}: {description}")
        self.code = code
        self.description = description


def api(method, **params):
    for attempt in range(4):
        r = requests.post(TG_API + method, json=params, timeout=40)
        data = r.json()
        if data.get("ok"):
            return data["result"]
        if data.get("error_code") == 429 and attempt < 3:
            # Telegram просит подождать: ждём ровно столько, сколько он сказал
            wait = int(data.get("parameters", {}).get("retry_after", 5)) + 1
            print(f"{method}: лимит запросов, жду {wait} с")
            time.sleep(wait)
            continue
        raise TgError(data.get("error_code"), data.get("description"))


# ---------------------------------------------------------------- Яндекс Диск
def latest_file():
    """Самый свежий xlsx в папке (по дате ДД.ММ.ГГГГ в имени, иначе по created)."""
    r = requests.get(YA_API, params={"public_key": PUBLIC_KEY, "limit": 200}, timeout=30)
    r.raise_for_status()
    data = r.json()
    items = data.get("_embedded", {}).get("items", []) if "_embedded" in data else [data]
    files = [i for i in items if i.get("type") == "file" and i["name"].lower().endswith((".xlsx", ".xlsm"))]
    if not files:
        return None

    def key(i):
        m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", i["name"])
        d = (m.group(3), m.group(2), m.group(1)) if m else ("0", "0", "0")
        return (d, i.get("created", ""))

    return max(files, key=key)


def download(item):
    params = {"public_key": PUBLIC_KEY}
    if item.get("path"):
        params["path"] = item["path"]
    r = requests.get(YA_API + "/download", params=params, timeout=30)
    r.raise_for_status()
    f = requests.get(r.json()["href"], timeout=60)
    f.raise_for_status()
    return f.content


# ---------------------------------------------------------------- разбор Excel
def col_to_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def norm(s):
    return re.sub(r"\s+", "", str(s or "")).upper()


def clean(s):
    s = " ".join(str(s).replace("\n", " ").split())
    return re.sub(r"(\d)\s*-\s*(\d)", r"\1-\2", s)


def group_key(label):
    """'483 РУ 26' -> '483', '4 7 2' -> '472' (буквы и год не учитываются)."""
    m = re.search(r"\d+", re.sub(r"\s+", "", str(label or "")))
    return m.group(0) if m else None


def read_sheet(xlsx_bytes):
    """Читает первый лист напрямую из XML (openpyxl падает на файлах Яндекса)."""
    z = zipfile.ZipFile(io.BytesIO(xlsx_bytes))

    def local(tag):
        return tag.rsplit("}", 1)[-1]

    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root:
            shared.append("".join(t.text or "" for t in si.iter() if local(t.tag) == "t"))

    sheet_name = sorted(n for n in z.namelist() if n.startswith("xl/worksheets/sheet"))[0]
    root = ET.fromstring(z.read(sheet_name))
    cells, merges = {}, {}
    for el in root.iter():
        t = local(el.tag)
        if t == "c":
            m = re.match(r"([A-Z]+)(\d+)", el.get("r"))
            pos = (int(m.group(2)), col_to_num(m.group(1)))
            kind = el.get("t")
            text = None
            if kind == "inlineStr":
                text = "".join(x.text or "" for x in el.iter() if local(x.tag) == "t")
            else:
                v = next((x.text for x in el if local(x.tag) == "v"), None)
                if v is not None:
                    text = shared[int(v)] if kind == "s" else v
            if text not in (None, ""):
                cells[pos] = text
        elif t == "mergeCell":
            a, b = el.get("ref").split(":")
            ma = re.match(r"([A-Z]+)(\d+)", a)
            mb = re.match(r"([A-Z]+)(\d+)", b)
            merges[(int(ma.group(2)), col_to_num(ma.group(1)))] = int(mb.group(2)) - int(ma.group(2)) + 1
    return cells, merges


def parse_block(cells, merges, r0, c0):
    lines = []
    for r in range(r0 + 1, r0 + 8):
        subj = cells.get((r, c0 + 3))
        if not subj:
            continue
        span = merges.get((r, c0 + 3), 1)
        try:
            first = int(float(cells.get((r, c0 + 1), r - r0)))
        except ValueError:
            first = r - r0
        num = f"{first}" if span == 1 else f"{first}–{first + span - 1}"
        t = clean(cells.get((r, c0 + 2), ""))
        if span > 1:
            end_t = clean(cells.get((r + span - 1, c0 + 2), ""))
            if "-" in t and "-" in end_t:
                t = t.split("-")[0].strip() + "-" + end_t.split("-")[-1].strip()
        parts = [clean(subj)]
        if cells.get((r, c0 + 4)):
            parts.append(clean(cells[(r, c0 + 4)]))
        if cells.get((r, c0 + 5)):
            room = clean(cells[(r, c0 + 5)])
            if room.endswith(".0"):
                room = room[:-2]
            parts.append(f"каб. {room}")
        lines.append(f"{num}) {t} — " + ", ".join(parts))
    return lines


def parse_all(xlsx_bytes):
    """-> {'483': [строки пар], ...} для всех групп файла."""
    cells, merges = read_sheet(xlsx_bytes)
    groups = {}
    for (r, c), v in sorted(cells.items()):
        if norm(cells.get((r, c + 1))) == "ВРЕМЯ":
            key = group_key(v)
            if key:
                groups[key] = parse_block(cells, merges, r, c)
    return groups


# ---------------------------------------------------------------- кэш и состояние
CACHE = {"sig": None, "date": None, "groups": {}, "ts": 0}


def refresh(force=False):
    if not force and CACHE["groups"] and time.time() - CACHE["ts"] < CACHE_TTL:
        return
    item = latest_file()
    if item:
        sig = f'{item["name"]}|{item.get("md5")}'
        if sig != CACHE["sig"]:
            groups = parse_all(download(item))
            CACHE.update(sig=sig, date=item["name"].rsplit(".", 1)[0], groups=groups)
    CACHE["ts"] = time.time()


STATE = {"users": {}, "sig": None}


def load_state():
    try:
        with open(DATA_FILE, encoding="utf-8") as f:
            STATE.update(json.load(f))
    except Exception:
        pass


def save_state():
    try:
        tmp = DATA_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(STATE, f, ensure_ascii=False)
        os.replace(tmp, DATA_FILE)
    except Exception as e:  # нельзя записать файл — бот всё равно продолжает работать
        log(f"Не удалось сохранить {DATA_FILE}: {e} (настройки пропадут после перезапуска)")


def user(chat_id):
    return STATE["users"].setdefault(str(chat_id), {"group": None, "sub": False})


# ---------------------------------------------------------------- тексты и кнопки
def schedule_text(group):
    lines = CACHE["groups"].get(group)
    date = CACHE["date"]
    if lines is None:
        return f"Расписание {date}: группы {group} в файле нет."
    if not lines:
        return f"Расписание {date}, группа {group}: пар нет."
    return f"Расписание {date}, группа {group}:\n" + "\n".join(lines)


def menu_view(chat_id):
    u = user(chat_id)
    g = u["group"] or "не выбрана"
    sub = "включены" if u["sub"] else "выключены"
    text = f"Меню\nГруппа: {g}\nУведомления: {sub}"
    kb = [
        [{"text": "📅 Расписание", "callback_data": "show"}],
        [{"text": "🔕 Отключить уведомления" if u["sub"] else "🔔 Подписаться на уведомления",
          "callback_data": "sub"}],
        [{"text": f"👥 Группа: {u['group']}" if u["group"] else "👥 Выбрать группу",
          "callback_data": "groups"}],
    ]
    return text, {"inline_keyboard": kb}


def groups_view():
    keys = sorted(CACHE["groups"], key=lambda k: (len(k), k))
    kb, row = [], []
    for k in keys:
        row.append({"text": k, "callback_data": f"g:{k}"})
        if len(row) == 4:
            kb.append(row)
            row = []
    if row:
        kb.append(row)
    kb.append([{"text": "⬅️ Назад", "callback_data": "menu"}])
    return "Выберите группу:", {"inline_keyboard": kb}


def send(chat_id, text, markup=None):
    p = {"chat_id": chat_id, "text": text}
    if markup:
        p["reply_markup"] = markup
    return api("sendMessage", **p)


def edit(chat_id, message_id, view):
    text, markup = view
    try:
        api("editMessageText", chat_id=chat_id, message_id=message_id, text=text, reply_markup=markup)
    except TgError as e:
        if "not modified" not in (e.description or ""):
            raise


# ---------------------------------------------------------------- обработка обновлений
def show_schedule(chat_id):
    u = user(chat_id)
    if not u["group"]:
        send(chat_id, "Сначала выберите группу.", groups_view()[1])
        return
    send(chat_id, schedule_text(u["group"]))
    send(chat_id, *menu_view(chat_id))


def handle_message(msg):
    chat_id = msg["chat"]["id"]
    text = (msg.get("text") or "").strip()
    if text.startswith("/schedule"):
        try:
            refresh()
        except Exception as e:
            print("refresh:", e)
        show_schedule(chat_id)
    else:  # /start, /menu и любой другой текст
        user(chat_id)
        save_state()
        send(chat_id, *menu_view(chat_id))


def handle_callback(cq):
    chat_id = cq["message"]["chat"]["id"]
    mid = cq["message"]["message_id"]
    data = cq.get("data", "")
    u = user(chat_id)
    answer = None

    if data in ("show", "groups") or data.startswith("g:"):
        try:
            refresh()
        except Exception as e:
            print("refresh:", e)

    if data == "menu":
        edit(chat_id, mid, menu_view(chat_id))
    elif data == "show":
        show_schedule(chat_id)
    elif data == "groups":
        if CACHE["groups"]:
            edit(chat_id, mid, groups_view())
        else:
            answer = "Не удалось получить список групп, попробуйте позже."
    elif data.startswith("g:"):
        key = data[2:]
        if key in CACHE["groups"]:
            u["group"] = key
            save_state()
            answer = f"Группа {key} выбрана"
        edit(chat_id, mid, menu_view(chat_id))
    elif data == "sub":
        if not u["group"]:
            answer = "Сначала выберите группу"
            edit(chat_id, mid, groups_view() if CACHE["groups"] else menu_view(chat_id))
        else:
            u["sub"] = not u["sub"]
            save_state()
            answer = "Уведомления включены" if u["sub"] else "Уведомления отключены"
            edit(chat_id, mid, menu_view(chat_id))

    try:
        api("answerCallbackQuery", callback_query_id=cq["id"], **({"text": answer} if answer else {}))
    except TgError:
        pass


def handle(update):
    if "message" in update:
        handle_message(update["message"])
    elif "callback_query" in update:
        handle_callback(update["callback_query"])


# ---------------------------------------------------------------- рассылка
def check():
    """Проверяет Диск; если файл новый — рассылает подписчикам."""
    refresh(force=True)
    sig = CACHE["sig"]
    if not sig or sig == STATE.get("sig"):
        return
    first_run = STATE.get("sig") is None
    STATE["sig"] = sig
    save_state()
    if first_run:
        return  # не шлём старое расписание при самом первом запуске
    for chat_id, u in list(STATE["users"].items()):
        if not (u.get("sub") and u.get("group")):
            continue
        try:
            send(chat_id, schedule_text(u["group"]))
        except TgError as e:
            if e.code in (400, 403):  # бота заблокировали / чат удалён
                STATE["users"].pop(chat_id, None)
                save_state()
            else:
                print("send:", chat_id, e)
        except Exception as e:
            print("send:", chat_id, e)
        time.sleep(0.05)


def main():
    load_state()
    try:
        me = api("getMe")
        log(f"Бот запущен: @{me.get('username')}. Пользователей: {len(STATE['users'])}")
        log(f"Пишите именно этому боту: https://t.me/{me.get('username')}")
    except Exception as e:
        log(f"Не удалось подключиться к Telegram (проверьте токен и интернет/VPN): {e}")
        return
    try:
        api("deleteWebhook")  # если был включён вебхук, getUpdates не работает
    except Exception as e:
        log(f"deleteWebhook: {e}")
    try:
        api("setMyCommands", commands=[
            {"command": "start", "description": "Меню"},
            {"command": "schedule", "description": "Моё расписание"},
        ])
    except Exception as e:
        print("setMyCommands:", e)
    offset = None
    last_check = 0
    while True:
        if time.time() - last_check >= CHECK_EVERY:
            try:
                check()
                log(f"Бот работает. Расписание: {CACHE['date']}, групп: {len(CACHE['groups'])}, "
                    f"подписчиков: {sum(1 for u in STATE['users'].values() if u.get('sub'))}")
            except Exception as e:
                log(f"Ошибка проверки Диска: {e}")
            last_check = time.time()
        try:
            params = {"timeout": 25, "allowed_updates": ["message", "callback_query"]}
            if offset:
                params["offset"] = offset
            updates = api("getUpdates", **params)
        except Exception as e:
            if "409" in str(e):
                log("Запущена ещё одна копия бота с этим токеном — остановите лишние.")
            log(f"getUpdates: {e}")
            time.sleep(5)
            continue
        for u in updates:
            offset = u["update_id"] + 1
            if "message" in u:
                log(f"Сообщение от {u['message']['chat']['id']}: {u['message'].get('text')}")
            elif "callback_query" in u:
                log(f"Кнопка от {u['callback_query']['message']['chat']['id']}: {u['callback_query'].get('data')}")
            try:
                handle(u)
            except Exception:
                log("Ошибка обработки сообщения:\n" + traceback.format_exc())


if __name__ == "__main__":
    main()
