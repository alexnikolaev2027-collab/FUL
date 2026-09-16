import datetime
import io
import json
import os

import flet as ft
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

STORE_KEY = "fuel_app_data_v1"

MONTHS_RU = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
             "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
MONTHS_GEN = ["января", "февраля", "марта", "апреля", "мая", "июня",
              "июля", "августа", "сентября", "октября", "ноября", "декабря"]

CAR_COLORS = [ft.Colors.BLUE, ft.Colors.TEAL, ft.Colors.ORANGE, ft.Colors.PURPLE,
              ft.Colors.GREEN, ft.Colors.PINK, ft.Colors.INDIGO, ft.Colors.BROWN]


def car_color(car_id):
    # Детерминированный цвет по id машины, чтобы в гараже не путать
    # авто на глаз и чтобы цвет не менялся между запусками. Встроенный
    # hash() для строк рандомизирован между запусками Python — поэтому
    # считаем свою простую стабильную сумму.
    s = sum((i + 1) * ord(ch) for i, ch in enumerate(str(car_id)))
    return CAR_COLORS[s % len(CAR_COLORS)]


def to_float(value, default=0.0):
    try:
        return float(str(value).replace(",", ".").replace(" ", ""))
    except Exception:
        return default


def num_str(value):
    if not value:
        return ""
    return "%g" % value


def fmt_num(value, nd=2):
    return f"{value:,.{nd}f}".replace(",", " ").replace(".", ",")


def fmt_date_ru(iso):
    d = datetime.date.fromisoformat(iso)
    return "%02d.%02d.%04d" % (d.day, d.month, d.year)


def calc_day(car, entry):
    city_l = entry["km_city"] / 100.0 * car["norm_city"]
    hw_l = entry["km_hw"] / 100.0 * car["norm_hw"]
    idle_l = entry["idle"] * car["norm_idle"]
    idle_dg_l = entry.get("idle_diesel", 0.0) * car.get("norm_idle_diesel", 0.0)
    return city_l, hw_l, idle_l, idle_dg_l, city_l + hw_l + idle_l + idle_dg_l


def running_state(car, entries, before_iso=None):
    bal = car["start_fuel"]
    odo = car["start_odo"]
    for e in sorted(entries, key=lambda x: x["date"]):
        if before_iso and e["date"] >= before_iso:
            break
        bal += e["issued"] - calc_day(car, e)[4]
        if e["odo"] > 0:
            odo = e["odo"]
    return bal, odo


def days_in_month(year, month):
    nxt = datetime.date(year + (1 if month == 12 else 0), month % 12 + 1, 1)
    return (nxt - datetime.date(year, month, 1)).days


def month_rows(car, entries, year, month):
    start_iso = "%04d-%02d-01" % (year, month)
    bal, odo = running_state(car, entries, start_iso)
    start_bal, start_odo = bal, odo
    prefix = "%04d-%02d-" % (year, month)
    by_day = {}
    for e in entries:
        if e["date"].startswith(prefix):
            by_day[int(e["date"][8:10])] = e
    rows = []
    km_sum = 0.0
    for day in range(1, days_in_month(year, month) + 1):
        e = by_day.get(day)
        issued = e["issued"] if e else 0.0
        if e:
            city_l, hw_l, idle_l, idle_dg_l, total_l = calc_day(car, e)
        else:
            city_l = hw_l = idle_l = idle_dg_l = total_l = 0.0
        bal = bal + issued - total_l
        if e and e["odo"] > 0:
            odo = e["odo"]
        if e:
            km_sum += e["km_city"] + e["km_hw"]
        # "Возможный пробег на остатках топлива" и "возможное показание
        # спидометра" считаются от ТЕКУЩЕГО остатка в баке (а не от
        # свободного места) — это ровно та логика, что в исходной
        # таблице (столбцы Z и AC): сколько ещё можно проехать на том,
        # что реально в баке прямо сейчас, и на какой отметке спидометра
        # это закончится.
        possible_range = (bal * 100.0 / car["norm_city"]) if car.get("norm_city") else 0.0
        possible_odo = (odo + possible_range) if e else ""
        rows.append({
            "day": day,
            "odo": e["odo"] if (e and e["odo"] > 0) else "",
            "idle": e["idle"] if e else 0.0,
            "idle_diesel": e.get("idle_diesel", 0.0) if e else 0.0,
            "work_hours": e.get("work_hours", 0.0) if e else 0.0,
            "km_city": e["km_city"] if e else 0.0,
            "km_hw": e["km_hw"] if e else 0.0,
            "balance": bal,
            "issued": issued,
            "city_l": city_l, "hw_l": hw_l, "idle_l": idle_l, "idle_dg_l": idle_dg_l,
            "total_l": total_l,
            "km_sum": km_sum,
            "possible_range": possible_range,
            "possible_odo": possible_odo,
            "empty": e is None,
            "note": (e.get("note", "") if e else ""),
        })
    totals = {
        "issued": sum(r["issued"] for r in rows),
        "km_city": sum(r["km_city"] for r in rows),
        "km_hw": sum(r["km_hw"] for r in rows),
        "idle": sum(r["idle"] for r in rows),
        "idle_diesel": sum(r["idle_diesel"] for r in rows),
        "work_hours": sum(r["work_hours"] for r in rows),
        "city_l": sum(r["city_l"] for r in rows),
        "hw_l": sum(r["hw_l"] for r in rows),
        "idle_l": sum(r["idle_l"] for r in rows),
        "idle_dg_l": sum(r["idle_dg_l"] for r in rows),
        "total_l": sum(r["total_l"] for r in rows),
        "end_balance": bal,
        "end_odo": odo,
    }
    return rows, totals, start_bal, start_odo


def build_xlsx(car, year, month, rows, totals, start_bal, start_odo):
    # Раскладка и формулировки заголовков повторяют реальную путевую
    # таблицу (файл-образец), включая раздельный учёт стоянки с
    # двигателем и с дизель-генератором и перенос показаний на начало/
    # конец месяца. Декоративные объединения пустых ячеек из образца не
    # воспроизводятся — переносится только то, что реально заполняется
    # и на что потом смотрят при переносе в путевые листы. "Заметка"
    # добавлена последним столбцом, как и просили.
    wb = Workbook()
    ws = wb.active
    ws.title = (MONTHS_RU[month - 1] + " " + str(year)).upper()[:31]

    bold = Font(bold=True)
    small = Font(size=9)
    thin = Side(style="thin")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(wrap_text=True, vertical="center", horizontal="center")

    has_dg = bool(car.get("norm_idle_diesel"))

    ws.merge_cells("A1:B1")
    ws["A1"] = "Сведения на начало месяца"
    ws.merge_cells("C1:F1" if has_dg else "C1:E1")
    ws["C1"] = "Нормы расхода"
    ws.merge_cells(("G1:H1" if has_dg else "F1:G1"))
    ws["G1" if has_dg else "F1"] = "Сведения на конец месяца"
    tot_start = "I1" if has_dg else "H1"
    tot_end = "N1" if has_dg else "L1"
    ws.merge_cells("%s:%s" % (tot_start, tot_end))
    ws[tot_start] = "Итоги за месяц (переносится на след. месяц)"
    for addr in ("A1", "C1", ("G1" if has_dg else "F1"), tot_start):
        ws[addr].font = bold
        ws[addr].alignment = center

    top_headers = [
        ("A2", "Спидометр на начало"),
        ("B2", "Остаток топлива на начало, л"),
        ("C2", "Норма город, л/100км"),
        ("D2", "Норма трасса, л/100км"),
        ("E2", "Норма стоянка (двиг.), л/ч"),
    ]
    col = 6
    if has_dg:
        top_headers.append((get_column_letter(col) + "2", "Норма стоянка (дизель-ген.), л/ч"))
        col += 1
    top_headers += [
        (get_column_letter(col) + "2", "Спидометр на конец"),
        (get_column_letter(col + 1) + "2", "Остаток топлива на конец, л"),
        (get_column_letter(col + 2) + "2", "Топливо за месяц с остатком, л"),
        (get_column_letter(col + 3) + "2", "Выдано топлива, л"),
        (get_column_letter(col + 4) + "2", "Пробег, км"),
        (get_column_letter(col + 5) + "2", "Расход топлива, л"),
        (get_column_letter(col + 6) + "2", "Стоянка (двиг.), ч"),
    ]
    last_top_col = col + 6
    if has_dg:
        top_headers.append((get_column_letter(col + 7) + "2", "Стоянка (дизель-ген.), ч"))
        last_top_col = col + 7
    for addr, text in top_headers:
        ws[addr] = text
        ws[addr].font = small
        ws[addr].alignment = center

    end_odo_col = col
    end_bal_col = col + 1
    month_fuel_col = col + 2
    issued_col = col + 3
    km_col = col + 4
    consum_col = col + 5
    idle_col = col + 6
    dg_col = col + 7 if has_dg else None

    ws["A3"] = start_odo
    ws["B3"] = round(start_bal, 2)
    ws["C3"] = car["norm_city"]
    ws["D3"] = car["norm_hw"]
    ws["E3"] = car["norm_idle"]
    if has_dg:
        ws["F3"] = car["norm_idle_diesel"]
    ws.cell(row=3, column=end_odo_col, value=totals["end_odo"])
    ws.cell(row=3, column=end_bal_col, value=round(totals["end_balance"], 2))
    ws.cell(row=3, column=month_fuel_col, value=round(start_bal + totals["issued"], 2))
    ws.cell(row=3, column=issued_col, value=round(totals["issued"], 2))
    ws.cell(row=3, column=km_col, value=round(totals["km_city"] + totals["km_hw"], 2))
    ws.cell(row=3, column=consum_col, value=round(totals["total_l"], 2))
    ws.cell(row=3, column=idle_col, value=round(totals["idle"], 2))
    if has_dg:
        ws.cell(row=3, column=dg_col, value=round(totals["idle_diesel"], 2))
    for c in range(1, last_top_col + 1):
        ws.cell(row=3, column=c).border = box

    ws["%s5" % get_column_letter(1)] = "Ёмкость бака, л"
    ws["%s6" % get_column_letter(1)] = car["tank"]
    ws["A5"].font = small
    ws["A6"].border = box

    header_row = 8
    ws.merge_cells(start_row=header_row - 1, start_column=1, end_row=header_row - 1,
                    end_column=last_top_col)
    ws.cell(row=header_row - 1, column=1,
            value="Накопительные данные за месяц (переносится в путевые листы построчно)")
    ws.cell(row=header_row - 1, column=1).font = bold

    headers = [
        "Дата", "Спидометр на конец смены, км", "Общее время в наряде, ч",
        "Стоянка с двигателем, ч",
    ]
    if has_dg:
        headers.append("Стоянка с дизель-генератором, ч")
    headers += [
        "Пробег всего, км", "в городе, км", "за городом, км",
        "Остаток топлива в баке, л", "Выдано топлива, л",
        "Расход всего, л", "расход в городе, л", "расход за городом, л",
        "расход на стоянке (двиг.), л",
    ]
    if has_dg:
        headers.append("расход на стоянке (дизель-ген.), л")
    headers += [
        "Сумма пробега, км", "Свободное место в баке, л",
        "Возможный пробег на остатках топлива, км",
        "Возможное показание спидометра на остатках",
        "Заметка",
    ]
    for c, text in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=c, value=text)
        cell.font = small
        cell.alignment = center
        cell.border = box
    ws.row_dimensions[header_row].height = 46

    r = header_row
    for row in rows:
        r += 1
        km_day = row["km_city"] + row["km_hw"]
        values = [
            row["day"],
            row["odo"] if row["odo"] != "" else None,
            row["work_hours"],
            row["idle"],
        ]
        if has_dg:
            values.append(row["idle_diesel"])
        values += [
            km_day, row["km_city"], row["km_hw"],
            round(row["balance"], 2), row["issued"],
            round(row["total_l"], 2), round(row["city_l"], 2), round(row["hw_l"], 2),
            round(row["idle_l"], 2),
        ]
        if has_dg:
            values.append(round(row["idle_dg_l"], 2))
        values += [
            round(row["km_sum"], 2),
            round(car["tank"] - row["balance"], 2),
            round(row["possible_range"], 2) if row["possible_odo"] != "" else "",
            round(row["possible_odo"], 0) if row["possible_odo"] != "" else "",
            row["note"],
        ]
        for c, v in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = box
            cell.font = Font(size=10)

    r += 1
    total_values = ["Итого", "", round(totals["work_hours"], 2), round(totals["idle"], 2)]
    if has_dg:
        total_values.append(round(totals["idle_diesel"], 2))
    total_values += [
        round(totals["km_city"] + totals["km_hw"], 2),
        round(totals["km_city"], 2), round(totals["km_hw"], 2),
        round(totals["end_balance"], 2), round(totals["issued"], 2),
        round(totals["total_l"], 2), round(totals["city_l"], 2), round(totals["hw_l"], 2),
        round(totals["idle_l"], 2),
    ]
    if has_dg:
        total_values.append(round(totals["idle_dg_l"], 2))
    total_values += ["", "", "", "", ""]
    for c, v in enumerate(total_values, start=1):
        cell = ws.cell(row=r, column=c, value=v)
        cell.border = box
        cell.font = bold

    widths = [6, 12, 10, 10]
    if has_dg:
        widths.append(11)
    widths += [10, 10, 10, 12, 10, 10, 11, 11, 11]
    if has_dg:
        widths.append(12)
    widths += [11, 12, 13, 14, 32]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A%d" % (header_row + 1)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------- хранилище данных ----------
# page.client_storage / page.shared_preferences ненадёжны между версиями
# Flet и платформами сборки (в частности, в мобильных сборках через
# serious_python их может не быть вовсе). Вместо этого пишем обычный
# JSON-файл в постоянную папку приложения FLET_APP_STORAGE_DATA — эта
# переменная окружения официально предоставляется Flet на всех
# платформах (Android/iOS/desktop/web) и гарантированно доступна.

APP_DATA_DIR = os.getenv("FLET_APP_STORAGE_DATA") or "."
DATA_FILE_PATH = os.path.join(APP_DATA_DIR, STORE_KEY + ".json")


def storage_get():
    if not os.path.exists(DATA_FILE_PATH):
        return None
    try:
        with open(DATA_FILE_PATH, "r", encoding="utf-8") as f:
            return f.read()
    except Exception:
        return None


def storage_set(value):
    os.makedirs(APP_DATA_DIR, exist_ok=True)
    with open(DATA_FILE_PATH, "w", encoding="utf-8") as f:
        f.write(value)


async def main(page: ft.Page):
    page.title = "Учёт топлива"
    page.theme_mode = ft.ThemeMode.SYSTEM

    # FilePicker в Flet 1.0 работает как сервис с асинхронными методами
    # (save_file_async вместо колбэка on_result).
    file_picker = ft.FilePicker()
    if hasattr(page, "services"):
        page.services.append(file_picker)
    else:
        page.overlay.append(file_picker)

    state = {
        "cars": [],
        "active_id": None,
        "entries": {},
        "tab": 0,
        "year": datetime.date.today().year,
        "month": datetime.date.today().month,
        "selected_date": datetime.date.today().isoformat(),
    }

    def new_car(name):
        return {
            "id": "car_%d" % int(datetime.datetime.now().timestamp() * 1000),
            "name": name,
            "plate": "",
            "tank": 75.0,
            "norm_city": 14.2,
            "norm_hw": 12.4,
            "norm_idle": 1.16,
            "norm_idle_diesel": 0.0,
            "start_odo": 0.0,
            "start_fuel": 0.0,
        }

    async def save_all():
        storage_set(json.dumps(
            {"cars": state["cars"], "active_id": state["active_id"], "entries": state["entries"]},
            ensure_ascii=False))

    async def load_all():
        raw = storage_get()
        data = json.loads(raw) if raw else {}
        state["cars"] = data.get("cars", [])
        state["active_id"] = data.get("active_id")
        state["entries"] = data.get("entries", {})
        if not state["cars"]:
            car = new_car("Основное ТС")
            state["cars"] = [car]
            state["active_id"] = car["id"]
            await save_all()

    def active_car():
        for c in state["cars"]:
            if c["id"] == state["active_id"]:
                return c
        return state["cars"][0]

    def car_entries(car_id=None):
        return state["entries"].setdefault(car_id or state["active_id"], [])

    # Разные сборки Flet по-разному реализуют показ оверлеев (диалог,
    # шторка меню, снекбар): где-то show_x(control) принимает контрол
    # аргументом, где-то show_x() — без аргументов и показывает то, что
    # уже лежит в соответствующем свойстве страницы (page.dialog,
    # page.drawer...), а где-то таких методов вообще нет и работает
    # только универсальный page.open()/page.close() (это оказалось
    # верно и для диалогов, и — после исправления — для бокового меню).
    # Эти две обёртки пробуют все правдоподобные варианты по очереди и
    # только в крайнем случае откатываются на самый старый способ
    # (.open=True + page.overlay), поэтому дальше по коду что открытие,
    # что закрытие не зависят от того, какая именно сигнатура в этой
    # сборке. ВАЖНО: и диалоги, и боковое меню (drawer) обязаны идти
    # через эти же обёртки — отдельная, "самодельная" реализация именно
    # для drawer раньше не доходила до фолбэка page.open() и поэтому
    # меню не открывалось по нажатию на иконку.
    def _show_overlay(page_prop, method_name, control):
        try:
            setattr(page, page_prop, control)
        except Exception:
            pass
        try:
            control.open = True
        except Exception:
            pass
        fn = getattr(page, method_name, None)
        if fn is not None:
            try:
                fn(control)
                page.update()
                return
            except TypeError:
                try:
                    fn()
                    page.update()
                    return
                except TypeError:
                    pass
        try:
            page.open(control)
            page.update()
            return
        except AttributeError:
            pass
        page.overlay.append(control)
        page.update()

    def _hide_overlay(method_name, control):
        if control is not None:
            try:
                control.open = False
            except Exception:
                pass
        fn = getattr(page, method_name, None)
        if fn is not None:
            try:
                fn()
                page.update()
                return
            except TypeError:
                try:
                    fn(control)
                    page.update()
                    return
                except TypeError:
                    pass
        try:
            page.close(control)
            page.update()
            return
        except AttributeError:
            pass
        page.update()

    def open_dialog(dlg):
        _show_overlay("dialog", "show_dialog", dlg)

    def close_dialog(dlg=None):
        _hide_overlay("pop_dialog", dlg)

    def snack(msg):
        sb = ft.SnackBar(ft.Text(msg))
        _show_overlay("dialog", "show_dialog", sb)

    def safe_shadow(blur, dy, opacity=0.10):
        # BoxShadow/Offset — не гарантированно одинаковые в разных
        # сборках Flet. Возвращает None, если конструктор не подошёл,
        # чтобы вызывающий код мог подстраховаться рамкой вместо тени.
        try:
            return ft.BoxShadow(blur_radius=blur, spread_radius=0,
                                color=ft.Colors.with_opacity(opacity, ft.Colors.BLACK),
                                offset=ft.Offset(0, dy))
        except Exception:
            return None

    def safe_container(shadow, **kwargs):
        # Строит Container с тенью, если параметр shadow= вообще
        # принимается в этой сборке Flet; если нет (или сама тень не
        # собралась выше) — тот же контейнер, но с тонкой рамкой вместо
        # тени, чтобы карточка всё равно была визуально видна. border
        # мог быть передан явно (например, синяя рамка "сегодня") —
        # в этом случае резервную рамку не подставляем.
        if shadow is None:
            if kwargs.get("border") is None:
                kwargs["border"] = ft.Border.all(1, ft.Colors.OUTLINE_VARIANT)
            return ft.Container(**kwargs)
        try:
            return ft.Container(shadow=shadow, **kwargs)
        except Exception:
            if kwargs.get("border") is None:
                kwargs["border"] = ft.Border.all(1, ft.Colors.OUTLINE_VARIANT)
            return ft.Container(**kwargs)

    def card_container(content, padding=14, bgcolor=None):
        # Общий "современный" стиль карточки: скруглённые углы + мягкая
        # тень вместо стандартной плоской рамки/жёсткой Card-тени.
        return safe_container(
            safe_shadow(14, 4),
            padding=padding, border_radius=18,
            bgcolor=bgcolor or ft.Colors.SURFACE,
            content=content)

    def info_row(label, ctrl):
        return ft.Row(
            [ft.Text(label, expand=True, size=14, color=ft.Colors.GREY_700), ctrl],
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN)

    def kpi_card(label, value):
        c = card_container(ft.Column([
            ft.Text(label, size=12, color=ft.Colors.GREY_600),
            ft.Text(value, size=26, weight=ft.FontWeight.W_600),
        ], spacing=2), padding=14)
        c.expand = True
        return c

    def fuel_gauge(balance, tank):
        # Дуговой индикатор остатка топлива — на глаз видно заполнение
        # бака, а не только цифру. Stack/ProgressRing тоже подстрахованы:
        # если в сборке их нет или другие параметры — просто текстовая
        # версия того же самого (без визуального кольца).
        pct = max(0.0, min(1.0, (balance / tank) if tank else 0.0))
        color = ft.Colors.RED_400 if pct < 0.15 else (
            ft.Colors.AMBER_600 if pct < 0.35 else ft.Colors.GREEN_600)
        try:
            return ft.Stack([
                ft.ProgressRing(value=pct, width=104, height=104, stroke_width=10,
                                color=color, bgcolor=ft.Colors.with_opacity(0.12, color)),
                ft.Container(
                    width=104, height=104, alignment=ft.Alignment.center,
                    content=ft.Column([
                        ft.Text(str(round(pct * 100)) + "%", size=20, weight=ft.FontWeight.W_600),
                        ft.Text(fmt_num(balance) + " л", size=11, color=ft.Colors.GREY_600),
                    ], spacing=0, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                ),
            ], width=104, height=104)
        except Exception:
            return ft.Container(
                width=104, height=104, alignment=ft.Alignment.center,
                content=ft.Column([
                    ft.Text(str(round(pct * 100)) + "%", size=22, weight=ft.FontWeight.W_600,
                            color=color),
                    ft.Text(fmt_num(balance) + " л", size=11, color=ft.Colors.GREY_600),
                ], spacing=0, horizontal_alignment=ft.CrossAxisAlignment.CENTER))

    def sparkline(values, height=44, color=None):
        # Простой спарклайн без графической библиотеки: ряд тонких
        # столбиков высотой пропорционально расходу за день. Пустые
        # (нулевые) дни рисуются минимальной "точкой", чтобы видеть
        # пробелы в ряду.
        color = color or ft.Colors.PRIMARY
        peak = max(values) if values and max(values) > 0 else 1.0
        bars = []
        for v in values:
            h = max(2, round((v / peak) * height)) if v else 2
            bars.append(ft.Container(
                width=5, height=h, bgcolor=color if v else ft.Colors.GREY_300,
                border_radius=3, alignment=ft.Alignment.bottom_center))
        return ft.Row(bars, spacing=3, height=height,
                      alignment=ft.MainAxisAlignment.START,
                      vertical_alignment=ft.CrossAxisAlignment.END)

    # ---------- экран 1: запись за день ----------

    def entry_view():
        car = active_car()
        entries = car_entries()
        iso = state["selected_date"]
        existing = next((e for e in entries if e["date"] == iso), None)
        bal_before, odo_before = running_state(car, entries, iso)

        date_tf = ft.TextField(label="Дата (ДД.ММ.ГГГГ)", value=fmt_date_ru(iso))
        odo_tf = ft.TextField(label="Спидометр на конец смены, км",
                              value=num_str(existing["odo"]) if existing else "",
                              keyboard_type=ft.KeyboardType.NUMBER)
        odo_hint = ft.Text(
            ("Предыдущий спидометр: %s км" % num_str(odo_before)) if odo_before else
            "Предыдущих записей нет — впишите пробег вручную",
            size=12, color=ft.Colors.GREY_600)
        total_km_tf = ft.TextField(
            label="Пробег всего, км (считается по спидометру)",
            value=num_str(existing["km_city"] + existing["km_hw"]) if existing else "",
            keyboard_type=ft.KeyboardType.NUMBER, expand=True)
        kmh_tf = ft.TextField(label="Пробег по трассе, км",
                              value=num_str(existing["km_hw"]) if existing else "",
                              keyboard_type=ft.KeyboardType.NUMBER, expand=True)
        res_city_km = ft.Text(size=18, weight=ft.FontWeight.W_600)
        res_hw_km = ft.Text(size=18, weight=ft.FontWeight.W_600)
        breakdown_card = ft.Container(
            padding=12,
            border=ft.Border.all(1, ft.Colors.OUTLINE_VARIANT),
            border_radius=10,
            content=ft.Row([
                ft.Column(
                    [ft.Text("Город", size=12, color=ft.Colors.GREY_600), res_city_km],
                    expand=True, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                ft.VerticalDivider(width=1),
                ft.Column(
                    [ft.Text("Трасса", size=12, color=ft.Colors.GREY_600), res_hw_km],
                    expand=True, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
            ]))
        has_dg = bool(car.get("norm_idle_diesel"))
        idle_tf = ft.TextField(label="Стоянка с двигателем, ч",
                               value=num_str(existing["idle"]) if existing else "",
                               keyboard_type=ft.KeyboardType.NUMBER, expand=True)
        idle_dg_tf = ft.TextField(label="Стоянка с дизель-генератором, ч",
                                  value=num_str(existing.get("idle_diesel", 0)) if existing else "",
                                  keyboard_type=ft.KeyboardType.NUMBER, expand=True)
        work_tf = ft.TextField(label="Общее время в наряде, ч",
                               value=num_str(existing.get("work_hours", 0)) if existing else "",
                               keyboard_type=ft.KeyboardType.NUMBER, expand=True)
        iss_tf = ft.TextField(label="Выдано топлива, л",
                              value=num_str(existing["issued"]) if existing else "",
                              keyboard_type=ft.KeyboardType.NUMBER)
        note_tf = ft.TextField(label="Заметка — куда ездил",
                               value=existing.get("note", "") if existing else "",
                               multiline=True, min_lines=1, max_lines=4)

        res_total = ft.Text(weight=ft.FontWeight.W_500, size=14)
        res_bal = ft.Text(weight=ft.FontWeight.W_500, size=14)
        res_free = ft.Text(size=14)
        res_range = ft.Text(size=14)
        res_possible_range = ft.Text(size=14)
        res_possible_odo = ft.Text(size=14)

        def recalc(*_):
            total_km = to_float(total_km_tf.value)
            hw_km = to_float(kmh_tf.value)
            city_km = max(total_km - hw_km, 0.0)
            temp = {"km_city": city_km, "km_hw": hw_km, "idle": to_float(idle_tf.value),
                    "idle_diesel": to_float(idle_dg_tf.value)}
            _c, _h, _i, _dg, total_l = calc_day(car, temp)
            bal = bal_before + to_float(iss_tf.value) - total_l
            free = car["tank"] - bal
            rng = free / (car["norm_city"] / 100.0) if car["norm_city"] else 0.0
            # Возможный пробег/показание спидометра на ОСТАТКАХ топлива —
            # та же логика, что в исходной таблице (столбцы "Возможный
            # пробег" и "Возможное показание спидометра"): считается от
            # текущего остатка в баке, а не от свободного места.
            cur_odo = to_float(odo_tf.value) or odo_before
            possible_range = (bal * 100.0 / car["norm_city"]) if car["norm_city"] else 0.0
            possible_odo = cur_odo + possible_range
            res_city_km.value = fmt_num(city_km, 0) + " км"
            res_hw_km.value = fmt_num(hw_km, 0) + " км"
            res_total.value = fmt_num(total_l) + " л"
            res_bal.value = fmt_num(bal) + " л"
            res_free.value = fmt_num(free) + " л"
            res_range.value = fmt_num(rng, 0) + " км"
            res_possible_range.value = fmt_num(possible_range, 0) + " км"
            res_possible_odo.value = fmt_num(possible_odo, 0) + " км"
            page.update()

        change_fields = [odo_tf, total_km_tf, kmh_tf, idle_tf, work_tf, iss_tf]
        if has_dg:
            change_fields.append(idle_dg_tf)
        for tf in change_fields:
            tf.on_change = recalc

        def on_odo_change(*_):
            new_odo = to_float(odo_tf.value)
            # Если знаем предыдущий спидометр — сами считаем пробег как
            # разницу, чтобы не вводить его вручную дважды. Поле
            # "Пробег всего" при этом остаётся редактируемым: если
            # разница посчиталась неверно (например, обнулился счётчик),
            # её можно поправить руками.
            if new_odo and odo_before and new_odo > odo_before:
                total_km_tf.value = num_str(new_odo - odo_before)
                odo_hint.value = "Пробег с прошлой записи: %s км" % num_str(new_odo - odo_before)
            elif new_odo and odo_before and new_odo <= odo_before:
                odo_hint.value = "Спидометр не больше предыдущего (%s км) — впишите пробег вручную" % num_str(odo_before)
            recalc()

        odo_tf.on_change = on_odo_change

        def on_date(*_):
            try:
                d = datetime.datetime.strptime(date_tf.value.strip(), "%d.%m.%Y").date()
            except Exception:
                snack("Проверьте дату: формат ДД.ММ.ГГГГ")
                return
            state["selected_date"] = d.isoformat()
            render()

        date_tf.on_blur = on_date

        async def save(*_):
            try:
                d = datetime.datetime.strptime(date_tf.value.strip(), "%d.%m.%Y").date()
            except Exception:
                snack("Проверьте дату: формат ДД.ММ.ГГГГ")
                return
            total_km = to_float(total_km_tf.value)
            hw_km = to_float(kmh_tf.value)
            new_odo = to_float(odo_tf.value)
            odo_warning = (new_odo and odo_before and new_odo < odo_before)
            ent = {
                "date": d.isoformat(),
                "odo": new_odo,
                "km_city": max(total_km - hw_km, 0.0),
                "km_hw": hw_km,
                "idle": to_float(idle_tf.value),
                "idle_diesel": to_float(idle_dg_tf.value),
                "work_hours": to_float(work_tf.value),
                "issued": to_float(iss_tf.value),
                "note": note_tf.value.strip() if note_tf.value else "",
            }
            lst = car_entries()
            for i, e in enumerate(lst):
                if e["date"] == ent["date"]:
                    lst[i] = ent
                    break
            else:
                lst.append(ent)
            state["selected_date"] = ent["date"]
            await save_all()
            if odo_warning:
                snack("Сохранено, но спидометр (%s км) меньше предыдущего показания (%s км) — проверьте" % (
                    num_str(new_odo), num_str(odo_before)))
            else:
                snack("Сохранено: " + fmt_date_ru(ent["date"]))
            render()

        async def do_delete(*_):
            lst = car_entries()
            lst[:] = [e for e in lst if e["date"] != iso]
            await save_all()
            snack("Запись удалена")
            render()

        def ask_delete(*_):
            async def confirm(e):
                close_dialog(dlg)
                await do_delete()
            dlg = ft.AlertDialog(
                title=ft.Text("Удалить запись?"),
                content=ft.Text("Данные за " + fmt_date_ru(iso) + " будут удалены."),
                actions=[
                    ft.TextButton("Отмена", on_click=lambda e: close_dialog(dlg)),
                    ft.TextButton("Удалить", on_click=confirm),
                ],
            )
            open_dialog(dlg)

        buttons = [ft.Button(content="Сохранить", icon=ft.Icons.SAVE, on_click=save, expand=True)]
        if existing:
            buttons.append(ft.OutlinedButton("Удалить", icon=ft.Icons.DELETE_OUTLINE, on_click=ask_delete))

        result_card = ft.Card(content=ft.Container(
            padding=12,
            content=ft.Column([
                info_row("Расход за день", res_total),
                info_row("Остаток в баке", res_bal),
                info_row("Свободно в баке (" + num_str(car["tank"]) + " л)", res_free),
                info_row("Возможный пробег на свободное место", res_range),
                ft.Divider(height=8),
                info_row("Возможный пробег на остатках топлива", res_possible_range),
                info_row("Возможное показание спидометра на остатках", res_possible_odo),
            ], spacing=6)))

        recalc()
        idle_row = [work_tf, idle_tf]
        if has_dg:
            idle_row.append(idle_dg_tf)
        return ft.ListView(
            controls=[date_tf, odo_tf, odo_hint,
                      ft.Row([total_km_tf, kmh_tf], spacing=10),
                      breakdown_card,
                      ft.Row(idle_row, spacing=10),
                      iss_tf,
                      note_tf,
                      result_card,
                      ft.Row(buttons)],
            expand=True, spacing=10, padding=ft.Padding.all(12))

    # ---------- экран 2: месяц ----------

    def month_view():
        car = active_car()
        y, m = state["year"], state["month"]
        rows, totals, _sb, _so = month_rows(car, car_entries(), y, m)
        working = [r for r in rows if not r["empty"]]

        def shift(delta, *_):
            idx = state["year"] * 12 + (state["month"] - 1) + delta
            state["year"] = idx // 12
            state["month"] = idx % 12 + 1
            render()

        async def del_entry_for(del_iso):
            lst = car_entries()
            lst[:] = [x for x in lst if x["date"] != del_iso]
            await save_all()
            render()

        def ask_del(e, del_iso):
            async def confirm(ev):
                close_dialog(dlg)
                await del_entry_for(del_iso)
            dlg = ft.AlertDialog(
                title=ft.Text("Удалить запись?"),
                content=ft.Text("Данные за " + fmt_date_ru(del_iso) + " будут удалены."),
                actions=[
                    ft.TextButton("Отмена", on_click=lambda ev: close_dialog(dlg)),
                    ft.TextButton("Удалить", on_click=confirm),
                ],
            )
            open_dialog(dlg)

        def open_day(e, open_iso):
            state["selected_date"] = open_iso
            state["tab"] = 0
            render()

        lv = ft.ListView(expand=True, spacing=8, padding=ft.Padding.all(12))
        lv.controls.append(ft.Row([
            ft.IconButton(ft.Icons.CHEVRON_LEFT, on_click=lambda e: shift(-1)),
            ft.Text(MONTHS_RU[m - 1] + " " + str(y), size=17,
                    weight=ft.FontWeight.W_500, expand=True, text_align=ft.TextAlign.CENTER),
            ft.IconButton(ft.Icons.CHEVRON_RIGHT, on_click=lambda e: shift(1)),
        ]))
        if not working:
            lv.controls.append(ft.Container(
                padding=40,
                content=ft.Column([
                    ft.Icon(ft.Icons.EVENT_NOTE, size=40, color=ft.Colors.GREY_400),
                    ft.Text("Нет записей за этот месяц", weight=ft.FontWeight.W_500,
                            text_align=ft.TextAlign.CENTER),
                    ft.Text("Добавьте первую запись на вкладке «Запись»",
                            text_align=ft.TextAlign.CENTER, color=ft.Colors.GREY_600, size=12),
                ], spacing=6, horizontal_alignment=ft.CrossAxisAlignment.CENTER)))
        today_iso = datetime.date.today().isoformat()
        for r in reversed(working):
            iso = "%04d-%02d-%02d" % (y, m, r["day"])
            sub = "%s км · выдано %s л · стоянка %s ч" % (
                fmt_num(r["km_city"] + r["km_hw"], 0),
                fmt_num(r["issued"]),
                fmt_num(r["idle"]))
            note = (r.get("note") or "").strip()
            if note:
                preview = note if len(note) <= 40 else note[:37] + "…"
                sub += "\n📝 " + preview
            is_today = (iso == today_iso)
            open_handler = lambda e, s=iso: open_day(e, s)

            def make_dismiss_handler(diso):
                async def handler(e):
                    await del_entry_for(diso)
                    snack("Запись за " + fmt_date_ru(diso) + " удалена")
                return handler

            row_shadow = None if is_today else safe_shadow(10, 3, opacity=0.08)
            row_border = ft.Border.all(2, ft.Colors.PRIMARY) if is_today else None
            row_card = safe_container(
                row_shadow,
                border=row_border, border_radius=16, bgcolor=ft.Colors.SURFACE,
                on_click=open_handler, ink=True,
                content=ft.ListTile(
                    title=ft.Text(str(r["day"]) + " " + MONTHS_GEN[m - 1] +
                                 (" · сегодня" if is_today else ""), size=14),
                    subtitle=ft.Text(sub, size=12),
                    trailing=ft.Row([
                        ft.Text(fmt_num(r["total_l"]) + " л", weight=ft.FontWeight.W_500, size=14),
                        ft.IconButton(ft.Icons.DELETE_OUTLINE, icon_size=18,
                                      on_click=lambda e, s=iso: ask_del(e, s)),
                    ], width=128),
                    on_click=open_handler,
                ),
            )
            # Свайп влево для удаления — приятный современный жест поверх
            # обычной кнопки-корзины. Если в этой сборке Flet нет
            # Dismissible (или другой набор параметров) — просто
            # показываем карточку без свайпа, кнопка-корзина всё равно
            # работает. Ловим Exception широко: тут это безопасно (запись
            # ещё не менялась), а поломка тут не должна ронять весь экран.
            try:
                lv.controls.append(ft.Dismissible(
                    key="entry_" + iso,
                    content=row_card,
                    dismiss_direction=ft.DismissDirection.END_TO_START,
                    background=ft.Container(
                        border_radius=16, bgcolor=ft.Colors.RED_400,
                        padding=ft.Padding.all(16),
                        alignment=ft.Alignment.center_right,
                        content=ft.Icon(ft.Icons.DELETE_OUTLINE, color=ft.Colors.WHITE)),
                    on_dismiss=make_dismiss_handler(iso),
                ))
            except Exception:
                lv.controls.append(row_card)
        return lv

    # ---------- экран 3: сводка ----------

    def summary_view():
        car = active_car()
        y, m = state["year"], state["month"]
        rows, totals, start_bal, start_odo = month_rows(car, car_entries(), y, m)
        km_total = totals["km_city"] + totals["km_hw"]

        async def export_click(e):
            data = build_xlsx(car, y, m, rows, totals, start_bal, start_odo)
            fname = "Топливо_%s_%d.xlsx" % (MONTHS_RU[m - 1], y)
            try:
                path = await file_picker.save_file_async(
                    file_name=fname, allowed_extensions=["xlsx"])
            except (AttributeError, TypeError):
                # Совместимость со старым (не-async) FilePicker API или
                # другим набором именованных аргументов в этой сборке.
                path = file_picker.save_file(file_name=fname, allowed_extensions=["xlsx"])
            if path:
                try:
                    with open(path, "wb") as fh:
                        fh.write(data)
                    snack("Сохранено: " + path)
                except Exception as ex:
                    snack("Ошибка сохранения: " + str(ex))
            page.update()

        lv = ft.ListView(expand=True, spacing=10, padding=ft.Padding.all(12))
        lv.controls.append(ft.Text(MONTHS_RU[m - 1] + " " + str(y), size=17,
                                   weight=ft.FontWeight.W_500))

        lv.controls.append(card_container(ft.Row([
            fuel_gauge(totals["end_balance"], car["tank"]),
            ft.Column([
                ft.Text("Остаток в баке", size=12, color=ft.Colors.GREY_600),
                ft.Text(fmt_num(totals["end_balance"]) + " л из " + num_str(car["tank"]) + " л",
                        size=15, weight=ft.FontWeight.W_500),
                ft.Text("Пробег за месяц: " + fmt_num(km_total, 0) + " км", size=13),
                ft.Text("Расход дней с данными:", size=12, color=ft.Colors.GREY_600),
                sparkline([r["total_l"] for r in rows if not r["empty"]]),
            ], spacing=4, expand=True),
        ], spacing=16, vertical_alignment=ft.CrossAxisAlignment.CENTER)))

        lv.controls.append(ft.Row([
            kpi_card("Остаток на конец", fmt_num(totals["end_balance"]) + " л"),
            kpi_card("Пробег за месяц", fmt_num(km_total, 0) + " км"),
        ]))
        lv.controls.append(ft.Row([
            kpi_card("Топливо выдано", fmt_num(totals["issued"]) + " л"),
            kpi_card("Расход всего", fmt_num(totals["total_l"]) + " л"),
        ]))
        lv.controls.append(card_container(ft.Column([
            info_row("Город", fmt_num(totals["km_city"], 0) + " км → " + fmt_num(totals["city_l"]) + " л"),
            info_row("Трасса", fmt_num(totals["km_hw"], 0) + " км → " + fmt_num(totals["hw_l"]) + " л"),
            info_row("Стоянка (двигатель)", fmt_num(totals["idle"]) + " ч → " + fmt_num(totals["idle_l"]) + " л"),
        ] + ([
            info_row("Стоянка (дизель-ген.)", fmt_num(totals["idle_diesel"]) + " ч → " + fmt_num(totals["idle_dg_l"]) + " л"),
        ] if car.get("norm_idle_diesel") else []) + [
            ft.Divider(height=8),
            info_row("На начало месяца", fmt_num(start_bal) + " л · " + fmt_num(start_odo, 0) + " км"),
            info_row("Нормы (город/трасса/стоянка)",
                     num_str(car["norm_city"]) + " / " + num_str(car["norm_hw"]) + " / " + num_str(car["norm_idle"])),
        ], spacing=6)))
        lv.controls.append(ft.Button(content="Экспорт в Excel", icon=ft.Icons.DOWNLOAD,
                                     on_click=export_click))
        lv.controls.append(ft.OutlinedButton("Параметры авто", icon=ft.Icons.SETTINGS,
                                             on_click=lambda e: open_car_dialog(active_car())))
        return lv

    # ---------- гараж ----------

    def open_car_dialog(car=None):
        is_new = car is None
        c = dict(car) if car else new_car("")

        # Все поля — на всю ширину, друг под другом. Раньше пары полей
        # (нормы город/трасса, спидометр/топливо на старте) стояли в
        # одном Row без expand, из-за чего на мобильном экране второе
        # поле в паре просто не помещалось и было невозможно в него
        # попасть — этой парой Row вообще больше не пользуемся.
        name_tf = ft.TextField(label="Название / марка", value=c["name"])
        plate_tf = ft.TextField(label="Госномер", value=c["plate"])
        tank_tf = ft.TextField(label="Ёмкость бака, л", value=num_str(c["tank"]),
                               keyboard_type=ft.KeyboardType.NUMBER)
        ncity_tf = ft.TextField(label="Норма город, л/100км", value=num_str(c["norm_city"]),
                                keyboard_type=ft.KeyboardType.NUMBER)
        nhw_tf = ft.TextField(label="Норма трасса, л/100км", value=num_str(c["norm_hw"]),
                              keyboard_type=ft.KeyboardType.NUMBER)
        nidle_tf = ft.TextField(label="Норма стоянка (двигатель), л/час", value=num_str(c["norm_idle"]),
                                keyboard_type=ft.KeyboardType.NUMBER)
        nidle_dg_tf = ft.TextField(
            label="Норма стоянка (дизель-генератор), л/час — оставьте пустым, если нет",
            value=num_str(c.get("norm_idle_diesel", 0.0)),
            keyboard_type=ft.KeyboardType.NUMBER)
        sodo_tf = ft.TextField(label="Спидометр на старте, км", value=num_str(c["start_odo"]),
                               keyboard_type=ft.KeyboardType.NUMBER)
        sfuel_tf = ft.TextField(label="Топливо на старте, л", value=num_str(c["start_fuel"]),
                                keyboard_type=ft.KeyboardType.NUMBER)

        async def save(e):
            if not name_tf.value.strip() and not plate_tf.value.strip():
                snack("Укажите название или госномер")
                return
            c["name"] = name_tf.value.strip() or "Авто"
            c["plate"] = plate_tf.value.strip()
            c["tank"] = to_float(tank_tf.value, 75.0)
            c["norm_city"] = to_float(ncity_tf.value)
            c["norm_hw"] = to_float(nhw_tf.value)
            c["norm_idle"] = to_float(nidle_tf.value)
            c["norm_idle_diesel"] = to_float(nidle_dg_tf.value)
            c["start_odo"] = to_float(sodo_tf.value)
            c["start_fuel"] = to_float(sfuel_tf.value)
            if is_new:
                state["cars"].append(c)
                state["active_id"] = c["id"]
            else:
                car.update(c)
            await save_all()
            close_dialog(dlg)
            snack("Сохранено")
            render()

        async def do_remove():
            state["cars"] = [x for x in state["cars"] if x["id"] != c["id"]]
            state["entries"].pop(c["id"], None)
            state["active_id"] = state["cars"][0]["id"]
            await save_all()
            close_dialog(dlg)
            snack("Авто удалено")
            render()

        def remove(e):
            if len(state["cars"]) <= 1:
                snack("Нельзя удалить единственное авто")
                return

            async def confirm(ev):
                close_dialog(confirm_dlg)
                await do_remove()

            confirm_dlg = ft.AlertDialog(
                title=ft.Text("Удалить авто?"),
                content=ft.Text("«%s» и все записи по нему будут удалены безвозвратно." % c["name"]),
                actions=[
                    ft.TextButton("Отмена", on_click=lambda ev: close_dialog(confirm_dlg)),
                    ft.TextButton("Удалить", on_click=confirm),
                ],
            )
            open_dialog(confirm_dlg)

        actions = [
            ft.TextButton("Отмена", on_click=lambda e: close_dialog(dlg)),
            ft.Button(content="Сохранить", on_click=save),
        ]
        if not is_new:
            actions.append(ft.TextButton("Удалить авто", on_click=remove))

        # Диалог подстраивается под размер экрана, а не под жёстко
        # заданные 400x340 — на телефоне с крупным шрифтом это и было
        # причиной того, что нижние поля просто обрезались.
        dlg_width = min((page.width or 380) - 32, 420)
        dlg_height = min((page.height or 700) - 140, 620)
        dlg = ft.AlertDialog(
            title=ft.Text("Новое авто" if is_new else "Параметры авто"),
            content=ft.Column([
                name_tf, plate_tf, tank_tf,
                ncity_tf, nhw_tf, nidle_tf, nidle_dg_tf,
                sodo_tf, sfuel_tf,
            ], height=dlg_height, width=dlg_width, scroll=ft.ScrollMode.AUTO, spacing=10),
            actions=actions,
            actions_alignment=ft.MainAxisAlignment.END,
        )
        open_dialog(dlg)

    def open_garage(*_):
        def make_select(cid):
            async def handler(e):
                if cid != state["active_id"]:
                    state["active_id"] = cid
                    await save_all()
                close_dialog(dlg)
                render()
            return handler

        async def add(e):
            c = new_car((name_new.value or "").strip() or "Новое авто")
            c["plate"] = (plate_new.value or "").strip()
            state["cars"].append(c)
            state["active_id"] = c["id"]
            await save_all()
            close_dialog(dlg)
            render()
            open_car_dialog(active_car())

        lv = ft.ListView(spacing=8, height=min(len(state["cars"]) * 66 + 10, 260))
        for c in state["cars"]:
            selected = c["id"] == state["active_id"]
            handler = make_select(c["id"])
            row = ft.ListTile(
                leading=ft.CircleAvatar(
                    bgcolor=car_color(c["id"]),
                    content=ft.Icon(ft.Icons.DIRECTIONS_CAR, color=ft.Colors.WHITE, size=18)),
                title=ft.Text(c["name"] + ((" · " + c["plate"]) if c["plate"] else ""), size=14),
                subtitle=ft.Text("бак %s л · нормы %s / %s / %s" % (
                    num_str(c["tank"]), num_str(c["norm_city"]),
                    num_str(c["norm_hw"]), num_str(c["norm_idle"])), size=12),
                trailing=ft.Icon(ft.Icons.CHECK_CIRCLE, size=18) if selected else None,
                on_click=handler,
            )
            # on_click ставим и на ListTile, и на оборачивающий Container —
            # в некоторых сборках Flet клик по вложенному ListTile не
            # всплывает до родителя, тогда сработает обработчик контейнера.
            lv.controls.append(ft.Container(
                border=ft.Border.all(1, ft.Colors.OUTLINE_VARIANT),
                border_radius=10,
                bgcolor=ft.Colors.SECONDARY_CONTAINER if selected else None,
                content=row,
                on_click=handler,
                ink=True,
            ))

        name_new = ft.TextField(label="Название нового авто", value="Новое авто")
        plate_new = ft.TextField(label="Госномер")

        dlg_width = min((page.width or 380) - 32, 420)
        dlg = ft.AlertDialog(
            title=ft.Text("Гараж"),
            content=ft.Column([
                ft.Text("Нажмите на авто, чтобы сделать его активным", size=12,
                        color=ft.Colors.GREY_600),
                lv,
                ft.Divider(height=8),
                ft.Text("Добавить автомобиль", weight=ft.FontWeight.W_500),
                name_new,
                plate_new,
                ft.Button(content="Добавить и настроить", icon=ft.Icons.ADD, on_click=add),
            ], height=min((page.height or 700) - 160, 520), width=dlg_width,
               scroll=ft.ScrollMode.AUTO, spacing=10),
            actions=[ft.TextButton("Закрыть", on_click=lambda e: close_dialog(dlg))],
        )
        open_dialog(dlg)

    # ---------- каркас ----------

    body = ft.Column(expand=True)

    # Боковое меню (Drawer) — полностью заменяет нижнюю навигацию.
    # Пустой NavigationDrawer создаём один раз, а его содержимое
    # (controls) перестраиваем при каждом render(), чтобы подсвечивать
    # активный пункт и актуальное название машины.
    drawer = ft.NavigationDrawer(controls=[])

    # ИСПРАВЛЕНО (настоящая причина): в актуальном Flet открытие и
    # закрытие NavigationDrawer делается ТОЛЬКО через асинхронные методы
    # страницы — await page.show_drawer() / await page.close_drawer().
    # У NavigationDrawer, в отличие от AlertDialog, нет свойства "open",
    # поэтому старый трюк "control.open = True; page.update()" на него
    # не действует. Предыдущая версия вызывала show_drawer()/close_drawer()
    # синхронно: вызов асинхронного метода без await молча создаёт
    # корутину и тут же её отбрасывает — ни исключения, ни эффекта.
    # Поэтому меню и не открывалось. Теперь вызываем с await, а если в
    # какой-то сборке Flet этих методов нет — подстраховываемся старым
    # способом на всякий случай.
    async def open_drawer(e=None):
        page.drawer = drawer
        fn = getattr(page, "show_drawer", None)
        if fn is not None:
            try:
                result = fn()
                if hasattr(result, "__await__"):
                    await result
                page.update()
                return
            except Exception:
                pass
        try:
            drawer.open = True
            page.update()
        except Exception:
            pass

    async def close_drawer(e=None):
        fn = getattr(page, "close_drawer", None)
        if fn is not None:
            try:
                result = fn()
                if hasattr(result, "__await__"):
                    await result
                page.update()
                return
            except Exception:
                pass
        try:
            drawer.open = False
            page.update()
        except Exception:
            pass

    def build_drawer_items():
        car = active_car()

        def go(idx):
            async def handler(e):
                await close_drawer()
                switch_tab(idx)
            return handler

        async def go_garage(e):
            await close_drawer()
            open_garage()

        async def go_car_settings(e):
            await close_drawer()
            open_car_dialog(active_car())

        async def go_export(e):
            await close_drawer()
            switch_tab(2)  # вкладка «Сводка» — там кнопка «Экспорт в Excel»

        items = [
            ft.Container(
                padding=ft.Padding.all(20),
                content=ft.Column([
                    ft.CircleAvatar(bgcolor=car_color(car["id"]),
                                    content=ft.Icon(ft.Icons.LOCAL_GAS_STATION,
                                                    color=ft.Colors.WHITE, size=20)),
                    ft.Text(car["name"] + ((" · " + car["plate"]) if car["plate"] else ""),
                            size=16, weight=ft.FontWeight.W_600),
                    ft.Text("Учёт топлива", size=12, color=ft.Colors.GREY_600),
                ], spacing=4),
            ),
            ft.Divider(height=1),
        ]

        nav_defs = [("Запись", ft.Icons.EDIT_NOTE, 0),
                    ("Месяц", ft.Icons.CALENDAR_MONTH, 1),
                    ("Сводка", ft.Icons.INSIGHTS, 2)]
        for label, icon, idx in nav_defs:
            selected = state["tab"] == idx
            items.append(ft.Container(
                padding=ft.Padding.all(4),
                content=ft.Container(
                    bgcolor=ft.Colors.SECONDARY_CONTAINER if selected else None,
                    border_radius=12,
                    content=ft.ListTile(
                        leading=ft.Icon(icon),
                        title=ft.Text(label, weight=ft.FontWeight.W_500 if selected else None),
                        on_click=go(idx),
                    ),
                ),
            ))

        items.append(ft.Divider(height=1))
        for label, icon, handler, sub in [
            ("Гараж", ft.Icons.GARAGE, go_garage, "%d авто" % len(state["cars"])),
            ("Параметры авто", ft.Icons.SETTINGS, go_car_settings, None),
            ("Экспорт в Excel", ft.Icons.DOWNLOAD, go_export, None),
        ]:
            items.append(ft.Container(
                padding=ft.Padding.all(4),
                content=ft.ListTile(
                    leading=ft.Icon(icon),
                    title=ft.Text(label),
                    subtitle=ft.Text(sub, size=11) if sub else None,
                    on_click=handler,
                ),
            ))
        return items

    def render():
        builders = (entry_view, month_view, summary_view)
        body.controls.clear()
        body.controls.append(builders[state["tab"]]())
        car = active_car()
        page.appbar.title = ft.Text(car["name"] + ((" · " + car["plate"]) if car["plate"] else ""))
        drawer.controls = build_drawer_items()
        page.update()

    def switch_tab(idx):
        state["tab"] = idx
        render()

    page.appbar = ft.AppBar(
        leading=ft.IconButton(ft.Icons.MENU, tooltip="Меню", on_click=open_drawer),
        title=ft.Text("Учёт топлива"),
    )
    page.drawer = drawer
    page.add(body)

    await load_all()
    render()


ft.run(main)
