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
    return city_l, hw_l, idle_l, city_l + hw_l + idle_l


def running_state(car, entries, before_iso=None):
    bal = car["start_fuel"]
    odo = car["start_odo"]
    for e in sorted(entries, key=lambda x: x["date"]):
        if before_iso and e["date"] >= before_iso:
            break
        bal += e["issued"] - calc_day(car, e)[3]
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
            city_l, hw_l, idle_l, total_l = calc_day(car, e)
        else:
            city_l = hw_l = idle_l = total_l = 0.0
        bal = bal + issued - total_l
        if e and e["odo"] > 0:
            odo = e["odo"]
        if e:
            km_sum += e["km_city"] + e["km_hw"]
        rows.append({
            "day": day,
            "odo": e["odo"] if (e and e["odo"] > 0) else "",
            "idle": e["idle"] if e else 0.0,
            "work_hours": e.get("work_hours", 0.0) if e else 0.0,
            "km_city": e["km_city"] if e else 0.0,
            "km_hw": e["km_hw"] if e else 0.0,
            "balance": bal,
            "issued": issued,
            "city_l": city_l, "hw_l": hw_l, "idle_l": idle_l, "total_l": total_l,
            "km_sum": km_sum,
            "empty": e is None,
        })
    totals = {
        "issued": sum(r["issued"] for r in rows),
        "km_city": sum(r["km_city"] for r in rows),
        "km_hw": sum(r["km_hw"] for r in rows),
        "idle": sum(r["idle"] for r in rows),
        "work_hours": sum(r["work_hours"] for r in rows),
        "city_l": sum(r["city_l"] for r in rows),
        "hw_l": sum(r["hw_l"] for r in rows),
        "idle_l": sum(r["idle_l"] for r in rows),
        "total_l": sum(r["total_l"] for r in rows),
        "end_balance": bal,
        "end_odo": odo,
    }
    return rows, totals, start_bal, start_odo


def build_xlsx(car, year, month, rows, totals, start_bal, start_odo):
    wb = Workbook()
    ws = wb.active
    ws.title = (MONTHS_RU[month - 1] + " " + str(year)).upper()[:31]

    bold = Font(bold=True)
    small = Font(size=9)
    thin = Side(style="thin")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(wrap_text=True, vertical="center", horizontal="center")

    ws.merge_cells("A1:C1")
    ws["A1"] = "Сведения на начало месяца"
    ws.merge_cells("F1:H1")
    ws["F1"] = "Нормы расхода"
    ws.merge_cells("J1:L1")
    ws["J1"] = "Сведения на конец данного месяца"
    ws.merge_cells("N1:R1")
    ws["N1"] = "Итоги за месяц"
    for addr in ("A1", "F1", "J1", "N1"):
        ws[addr].font = bold
        ws[addr].alignment = center

    top_headers = [
        ("A3", "Показания спидометра на начало"),
        ("C3", "Остаток топлива на начало, л"),
        ("F3", "Норма по городу, л/100км"),
        ("G3", "Норма за городом, л/100км"),
        ("H3", "Норма за 1 час стоянки, л"),
        ("J3", "Показания спидометра на конец"),
        ("K3", "Остаток топлива на конец, л"),
        ("N3", "Топливо за месяц с остатком, л"),
        ("O3", "Выдано топлива, л"),
        ("P3", "Пробег, км"),
        ("Q3", "Расход топлива, л"),
        ("R3", "Стоянка с двигателем, ч"),
    ]
    for addr, text in top_headers:
        ws[addr] = text
        ws[addr].font = small
        ws[addr].alignment = center

    ws["F5"] = car["norm_city"]
    ws["G5"] = car["norm_hw"]
    ws["H5"] = car["norm_idle"]
    ws["A6"] = start_odo
    ws["C6"] = round(start_bal, 2)
    ws["J6"] = totals["end_odo"]
    ws["K6"] = round(totals["end_balance"], 2)
    ws["N5"] = round(start_bal + totals["issued"], 2)
    ws["O5"] = round(totals["issued"], 2)
    ws["P5"] = round(totals["km_city"] + totals["km_hw"], 2)
    ws["Q5"] = round(totals["total_l"], 2)
    ws["R5"] = round(totals["idle"], 2)
    for addr in ("F5", "G5", "H5", "A6", "C6", "J6", "K6", "N5", "O5", "P5", "Q5", "R5"):
        ws[addr].border = box

    ws.merge_cells("A8:N8")
    ws["A8"] = "Накопительные данные за текущий месяц"
    ws["A8"].font = bold

    headers = [
        "Дата", "Спидометр на конец смены, км", "Общее рабочее время, ч",
        "Стоянка с двигателем, ч",
        "Пробег всего, км", "в городе, км", "за городом, км",
        "Остаток топлива в баке, л", "Выдано топлива, л",
        "Расход всего, л", "расход в городе, л", "расход за городом, л",
        "расход на стоянке, л", "Сумма пробега, км", "Свободное место в баке, л",
    ]
    header_row = 9
    for col, text in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=text)
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
            km_day,
            row["km_city"],
            row["km_hw"],
            round(row["balance"], 2),
            row["issued"],
            round(row["total_l"], 2),
            round(row["city_l"], 2),
            round(row["hw_l"], 2),
            round(row["idle_l"], 2),
            round(row["km_sum"], 2),
            round(car["tank"] - row["balance"], 2),
        ]
        for col, v in enumerate(values, start=1):
            cell = ws.cell(row=r, column=col, value=v)
            cell.border = box
            cell.font = Font(size=10)

    r += 1
    total_values = [
        "Итого", "",
        round(totals["work_hours"], 2),
        round(totals["idle"], 2),
        round(totals["km_city"] + totals["km_hw"], 2),
        round(totals["km_city"], 2),
        round(totals["km_hw"], 2),
        round(totals["end_balance"], 2),
        round(totals["issued"], 2),
        round(totals["total_l"], 2),
        round(totals["city_l"], 2),
        round(totals["hw_l"], 2),
        round(totals["idle_l"], 2),
        "", "",
    ]
    for col, v in enumerate(total_values, start=1):
        cell = ws.cell(row=r, column=col, value=v)
        cell.border = box
        cell.font = bold

    widths = [6, 12, 10, 10, 10, 10, 10, 12, 10, 10, 11, 11, 11, 11, 12]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A10"
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

    def snack(msg):
        page.show_dialog(ft.SnackBar(ft.Text(msg)))

    def info_row(label, ctrl):
        return ft.Row(
            [ft.Text(label, expand=True, size=14, color=ft.Colors.GREY_700), ctrl],
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN)

    def kpi_card(label, value):
        return ft.Card(
            content=ft.Container(
                padding=12,
                content=ft.Column([
                    ft.Text(label, size=12, color=ft.Colors.GREY_600),
                    ft.Text(value, size=22, weight=ft.FontWeight.W_500),
                ], spacing=2),
            ),
            expand=True)

    # ---------- экран 1: запись за день ----------

    def entry_view():
        car = active_car()
        entries = car_entries()
        iso = state["selected_date"]
        existing = next((e for e in entries if e["date"] == iso), None)
        bal_before, _odo_before = running_state(car, entries, iso)

        date_tf = ft.TextField(label="Дата (ДД.ММ.ГГГГ)", value=fmt_date_ru(iso))
        odo_tf = ft.TextField(label="Спидометр на конец смены, км",
                              value=num_str(existing["odo"]) if existing else "",
                              keyboard_type=ft.KeyboardType.NUMBER)
        total_km_tf = ft.TextField(
            label="Пробег всего, км",
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
        idle_tf = ft.TextField(label="Стоянка с двигателем, ч",
                               value=num_str(existing["idle"]) if existing else "",
                               keyboard_type=ft.KeyboardType.NUMBER, expand=True)
        work_tf = ft.TextField(label="Общее рабочее время, ч",
                               value=num_str(existing.get("work_hours", 0)) if existing else "",
                               keyboard_type=ft.KeyboardType.NUMBER, expand=True)
        iss_tf = ft.TextField(label="Выдано топлива, л",
                              value=num_str(existing["issued"]) if existing else "",
                              keyboard_type=ft.KeyboardType.NUMBER)

        res_total = ft.Text(weight=ft.FontWeight.W_500, size=14)
        res_bal = ft.Text(weight=ft.FontWeight.W_500, size=14)
        res_free = ft.Text(size=14)
        res_range = ft.Text(size=14)

        def recalc(*_):
            total_km = to_float(total_km_tf.value)
            hw_km = to_float(kmh_tf.value)
            city_km = max(total_km - hw_km, 0.0)
            temp = {"km_city": city_km, "km_hw": hw_km, "idle": to_float(idle_tf.value)}
            _c, _h, _i, total_l = calc_day(car, temp)
            bal = bal_before + to_float(iss_tf.value) - total_l
            free = car["tank"] - bal
            rng = free / (car["norm_city"] / 100.0) if car["norm_city"] else 0.0
            res_city_km.value = fmt_num(city_km, 0) + " км"
            res_hw_km.value = fmt_num(hw_km, 0) + " км"
            res_total.value = fmt_num(total_l) + " л"
            res_bal.value = fmt_num(bal) + " л"
            res_free.value = fmt_num(free) + " л"
            res_range.value = fmt_num(rng, 0) + " км"
            page.update()

        for tf in (odo_tf, total_km_tf, kmh_tf, idle_tf, work_tf, iss_tf):
            tf.on_change = recalc

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
            ent = {
                "date": d.isoformat(),
                "odo": to_float(odo_tf.value),
                "km_city": max(total_km - hw_km, 0.0),
                "km_hw": hw_km,
                "idle": to_float(idle_tf.value),
                "work_hours": to_float(work_tf.value),
                "issued": to_float(iss_tf.value),
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
                page.pop_dialog()
                await do_delete()
            dlg = ft.AlertDialog(
                title=ft.Text("Удалить запись?"),
                content=ft.Text("Данные за " + fmt_date_ru(iso) + " будут удалены."),
                actions=[
                    ft.TextButton("Отмена", on_click=lambda e: page.pop_dialog()),
                    ft.TextButton("Удалить", on_click=confirm),
                ],
            )
            page.show_dialog(dlg)

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
            ], spacing=6)))

        recalc()
        return ft.ListView(
            controls=[date_tf, odo_tf,
                      ft.Row([total_km_tf, kmh_tf], spacing=10),
                      breakdown_card,
                      ft.Row([work_tf, idle_tf], spacing=10),
                      iss_tf,
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
                page.pop_dialog()
                await del_entry_for(del_iso)
            dlg = ft.AlertDialog(
                title=ft.Text("Удалить запись?"),
                content=ft.Text("Данные за " + fmt_date_ru(del_iso) + " будут удалены."),
                actions=[
                    ft.TextButton("Отмена", on_click=lambda ev: page.pop_dialog()),
                    ft.TextButton("Удалить", on_click=confirm),
                ],
            )
            page.show_dialog(dlg)

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
                content=ft.Text("Нет записей за этот месяц. Добавьте на вкладке «Запись».",
                                text_align=ft.TextAlign.CENTER, color=ft.Colors.GREY_600)))
        for r in reversed(working):
            iso = "%04d-%02d-%02d" % (y, m, r["day"])
            sub = "%s км · выдано %s л · стоянка %s ч" % (
                fmt_num(r["km_city"] + r["km_hw"], 0),
                fmt_num(r["issued"]),
                fmt_num(r["idle"]))
            lv.controls.append(ft.Container(
                border=ft.Border.all(1, ft.Colors.OUTLINE_VARIANT),
                border_radius=10,
                content=ft.ListTile(
                    title=ft.Text(str(r["day"]) + " " + MONTHS_GEN[m - 1], size=14),
                    subtitle=ft.Text(sub, size=12),
                    trailing=ft.Row([
                        ft.Text(fmt_num(r["total_l"]) + " л", weight=ft.FontWeight.W_500, size=14),
                        ft.IconButton(ft.Icons.DELETE_OUTLINE, icon_size=18,
                                      on_click=lambda e, s=iso: ask_del(e, s)),
                    ], width=128),
                    on_click=lambda e, s=iso: open_day(e, s),
                ),
            ))
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
            except AttributeError:
                # Совместимость со старым (не-async) FilePicker API
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
        lv.controls.append(ft.Row([
            kpi_card("Остаток на конец", fmt_num(totals["end_balance"]) + " л"),
            kpi_card("Пробег за месяц", fmt_num(km_total, 0) + " км"),
        ]))
        lv.controls.append(ft.Row([
            kpi_card("Топливо выдано", fmt_num(totals["issued"]) + " л"),
            kpi_card("Расход всего", fmt_num(totals["total_l"]) + " л"),
        ]))
        lv.controls.append(ft.Card(content=ft.Container(
            padding=12,
            content=ft.Column([
                info_row("Город", fmt_num(totals["km_city"], 0) + " км → " + fmt_num(totals["city_l"]) + " л"),
                info_row("Трасса", fmt_num(totals["km_hw"], 0) + " км → " + fmt_num(totals["hw_l"]) + " л"),
                info_row("Стоянка", fmt_num(totals["idle"]) + " ч → " + fmt_num(totals["idle_l"]) + " л"),
                ft.Divider(height=8),
                info_row("На начало месяца", fmt_num(start_bal) + " л · " + fmt_num(start_odo, 0) + " км"),
                info_row("Нормы (город/трасса/стоянка)",
                         num_str(car["norm_city"]) + " / " + num_str(car["norm_hw"]) + " / " + num_str(car["norm_idle"])),
            ], spacing=6))))
        lv.controls.append(ft.Button(content="Экспорт в Excel", icon=ft.Icons.DOWNLOAD,
                                     on_click=export_click))
        lv.controls.append(ft.OutlinedButton("Параметры авто", icon=ft.Icons.SETTINGS,
                                             on_click=lambda e: open_car_dialog(active_car())))
        return lv

    # ---------- гараж ----------

    def open_car_dialog(car=None):
        is_new = car is None
        c = dict(car) if car else new_car("")

        name_tf = ft.TextField(label="Название / марка", value=c["name"])
        plate_tf = ft.TextField(label="Госномер", value=c["plate"])
        tank_tf = ft.TextField(label="Ёмкость бака, л", value=num_str(c["tank"]),
                               keyboard_type=ft.KeyboardType.NUMBER)
        ncity_tf = ft.TextField(label="Норма город, л/100км", value=num_str(c["norm_city"]),
                                keyboard_type=ft.KeyboardType.NUMBER)
        nhw_tf = ft.TextField(label="Норма трасса, л/100км", value=num_str(c["norm_hw"]),
                              keyboard_type=ft.KeyboardType.NUMBER)
        nidle_tf = ft.TextField(label="Норма стоянка, л/час", value=num_str(c["norm_idle"]),
                                keyboard_type=ft.KeyboardType.NUMBER)
        sodo_tf = ft.TextField(label="Спидометр на старте, км", value=num_str(c["start_odo"]),
                               keyboard_type=ft.KeyboardType.NUMBER)
        sfuel_tf = ft.TextField(label="Топливо на старте, л", value=num_str(c["start_fuel"]),
                                keyboard_type=ft.KeyboardType.NUMBER)

        async def save(e):
            c["name"] = name_tf.value.strip() or "Авто"
            c["plate"] = plate_tf.value.strip()
            c["tank"] = to_float(tank_tf.value, 75.0)
            c["norm_city"] = to_float(ncity_tf.value)
            c["norm_hw"] = to_float(nhw_tf.value)
            c["norm_idle"] = to_float(nidle_tf.value)
            c["start_odo"] = to_float(sodo_tf.value)
            c["start_fuel"] = to_float(sfuel_tf.value)
            if is_new:
                state["cars"].append(c)
                state["active_id"] = c["id"]
            else:
                car.update(c)
            await save_all()
            page.pop_dialog()
            render()

        async def remove(e):
            if len(state["cars"]) <= 1:
                snack("Нельзя удалить единственное авто")
                return
            state["cars"] = [x for x in state["cars"] if x["id"] != c["id"]]
            state["entries"].pop(c["id"], None)
            state["active_id"] = state["cars"][0]["id"]
            await save_all()
            page.pop_dialog()
            render()

        actions = [
            ft.TextButton("Отмена", on_click=lambda e: page.pop_dialog()),
            ft.Button(content="Сохранить", on_click=save),
        ]
        if not is_new:
            actions.append(ft.TextButton("Удалить авто", on_click=remove))
        dlg = ft.AlertDialog(
            title=ft.Text("Новое авто" if is_new else "Параметры авто"),
            content=ft.Column([
                name_tf, plate_tf, tank_tf,
                ft.Row([ncity_tf, nhw_tf]),
                nidle_tf,
                ft.Row([sodo_tf, sfuel_tf]),
            ], height=400, width=340, scroll=ft.ScrollMode.AUTO, spacing=10),
            actions=actions,
            actions_alignment=ft.MainAxisAlignment.END,
        )
        page.show_dialog(dlg)

    def open_garage(*_):
        async def select(cid, dlg):
            state["active_id"] = cid
            await save_all()
            page.pop_dialog()
            render()

        async def add(e):
            c = new_car((name_new.value or "").strip() or "Новое авто")
            c["plate"] = (plate_new.value or "").strip()
            state["cars"].append(c)
            state["active_id"] = c["id"]
            await save_all()
            page.pop_dialog()
            render()
            open_car_dialog(active_car())

        lv = ft.ListView(spacing=8, height=190)
        for c in state["cars"]:
            selected = c["id"] == state["active_id"]
            lv.controls.append(ft.Container(
                border=ft.Border.all(1, ft.Colors.OUTLINE_VARIANT),
                border_radius=10,
                bgcolor=ft.Colors.SECONDARY_CONTAINER if selected else None,
                content=ft.ListTile(
                    leading=ft.Icon(ft.Icons.DIRECTIONS_CAR),
                    title=ft.Text(c["name"] + ((" · " + c["plate"]) if c["plate"] else ""), size=14),
                    subtitle=ft.Text("бак %s л · нормы %s / %s / %s" % (
                        num_str(c["tank"]), num_str(c["norm_city"]),
                        num_str(c["norm_hw"]), num_str(c["norm_idle"])), size=12),
                    trailing=ft.Icon(ft.Icons.CHECK_CIRCLE, size=18) if selected else None,
                    on_click=lambda e, cid=c["id"]: select(cid, dlg),
                ),
            ))

        name_new = ft.TextField(label="Название нового авто", value="Новое авто")
        plate_new = ft.TextField(label="Госномер")

        dlg = ft.AlertDialog(
            title=ft.Text("Гараж"),
            content=ft.Column([
                lv,
                ft.Divider(height=8),
                ft.Text("Добавить автомобиль", weight=ft.FontWeight.W_500),
                name_new,
                plate_new,
                ft.Button(content="Добавить и настроить", icon=ft.Icons.ADD, on_click=add),
            ], height=470, width=340, scroll=ft.ScrollMode.AUTO, spacing=10),
        )
        page.show_dialog(dlg)

    # ---------- каркас ----------

    body = ft.Column(expand=True)

    def render():
        builders = (entry_view, month_view, summary_view)
        body.controls.clear()
        body.controls.append(builders[state["tab"]]())
        car = active_car()
        page.appbar.title = ft.Text(car["name"] + ((" · " + car["plate"]) if car["plate"] else ""))
        page.navigation_bar.selected_index = state["tab"]
        page.update()

    def switch_tab(idx):
        state["tab"] = idx
        render()

    page.appbar = ft.AppBar(
        title=ft.Text("Учёт топлива"),
        actions=[ft.IconButton(ft.Icons.GARAGE, tooltip="Гараж", on_click=open_garage)],
    )
    page.navigation_bar = ft.NavigationBar(
        destinations=[
            ft.NavigationBarDestination(icon=ft.Icons.EDIT_NOTE, label="Запись"),
            ft.NavigationBarDestination(icon=ft.Icons.CALENDAR_MONTH, label="Месяц"),
            ft.NavigationBarDestination(icon=ft.Icons.INSIGHTS, label="Сводка"),
        ],
        on_change=lambda e: switch_tab(e.control.selected_index),
    )
    page.add(body)

    await load_all()
    render()


ft.run(main)
