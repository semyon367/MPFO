import io
import re
from collections import defaultdict
from datetime import date, datetime, timedelta

import openpyxl
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.styles.borders import Border, Side

st.set_page_config(page_title="МП Инспектор — Итоги по ФО", page_icon="📊", layout="wide")

# ================== Конфигурация ==================
SHEET_NAME = "Детализация МП Инспектор"

COLUMN_KEYWORDS = {
    "subjekt":      ["субъект рф"],
    "podrazd":      ["подразделение"],
    "vid_nadzora":  ["вид надзора"],
    "nom_knm":      ["номер кнм"],
    "vid":          ["вид"],
    "status":       ["статус кнм"],
    "narusheniya":  ["нарушения выявлены"],
    "proverka_ogv": ["проверка огв/омсу"],
    "knd":          ["кнд"],
    "ssylki":       ["ссылки на файлы"],
    "date_act":     ["дата составления акта о результате кнм", "дата составления акта"],
    "s_vks":        ["с вкс", "вкс"],
}

DISTRICTS = [
    ("ЦФО", "Центральный ФО", [
        "УНДиПР ГУ МЧС России по Тверской области",
        "УНДиПР ГУ МЧС России по Курской области",
        "УНДиПР ГУ МЧС России по г. Москве",
        "УНДиПР ГУ МЧС России по Московской области",
        "УНДиПР ГУ МЧС России по Владимирской области",
        "УНДиПР ГУ МЧС России по Тамбовской области",
        "УНДиПР ГУ МЧС России по Тульской области",
        "УНДиПР ГУ МЧС России по Липецкой области",
        "УНДиПР ГУ МЧС России по Рязанской области",
        "УНДиПР ГУ МЧС России по Костромской области",
        "УНДиПР ГУ МЧС России по Ярославской области",
        "УНДиПР ГУ МЧС России по Ивановской области",
        "УНДиПР ГУ МЧС России по Воронежской области",
        "УНДиПР ГУ МЧС России по Калужской области",
        "УНДиПР ГУ МЧС России по Белгородской области",
        "УНДиПР ГУ МЧС России по Брянской области",
        "УНДиПР ГУ МЧС России по Смоленской области",
        "УНДПР ГУ МЧС России по Орловской области",
    ]),
    ("СЗФО", "Северо-Западный ФО", [
        "УНДПР Главного управления МЧС России по г. Санкт-Петербургу",
        "УНДиПР ГУ МЧС России по Ленинградской области",
        "УНДиПР ГУ МЧС России по Калининградской области",
        "УНДиПР ГУ МЧС России по Псковской области",
        "УНДиПР ГУ МЧС России по Республике Коми",
        "УНДиПР ГУ МЧС России по Архангельской области",
        "УНДиПР ГУ МЧС России по Вологодской области",
        "УНДиПР ГУ МЧС России по Новгородской области",
        "УНДиПР ГУ МЧС России по Республике Карелия",
        "УНДиПР ГУ МЧС России по Мурманской области",
        "ОНДиПР ГУ МЧС России по Ненецкому автономному округу",
    ]),
    ("СКФО", "Северо-Кавказский ФО", [
        "УНДиПР ГУ МЧС России по Кабардино-Балкарской Республике",
        "УНДиПР ГУ МЧС России по Республике Северная Осетия - Алания",
        "УНДиПР ГУ МЧС России по Республике Дагестан",
        "УНДиПР ГУ МЧС России по Карачаево-Черкесской Республике",
        "УНДиПР ГУ МЧС России по Ставропольскому краю",
        "УНДиПР ГУ МЧС России по Республике Ингушетия",
        "УНДиПР ГУ МЧС России по Чеченской Республике",
    ]),
    ("ЮФО", "Южный ФО", [
        "УНДиПР ГУ МЧС России по г. Севастополю",
        "УНДиПР ГУ МЧС России по Волгоградской области",
        "УНДиПР ГУ МЧС России по Ростовской области",
        "УНДиПР ГУ МЧС России по Республике Адыгея",
        "УНДиПР ГУ МЧС России по Астраханской области",
        "УНДиПР ГУ МЧС России по Республике Крым",
        "УНДиПР ГУ МЧС России по Республике Калмыкия",
        "УНДиПР ГУ МЧС России по Краснодарскому краю",
    ]),
    ("ПФО", "Приволжский ФО", [
        "УНДиПР ГУ МЧС России по Пензенской области",
        "УНДиПР ГУ МЧС России по Оренбургской области",
        "УНДиПР ГУ МЧС России по Ульяновской области",
        "УНДиПР ГУ МЧС России по Республике Башкортостан",
        "УНДиПР ГУ МЧС России по Удмуртской Республике",
        "УНДиПР ГУ МЧС России по Самарской области",
        "УНДиПР ГУ МЧС России по Нижегородской области",
        "УНДиПР ГУ МЧС России по Кировской области",
        "УНДиПР ГУ МЧС России по Пермскому краю",
        "УНДиПР ГУ МЧС России по Саратовской области",
        "УНДиПР ГУ МЧС России по Республике Мордовия",
        "УНДиПР ГУ МЧС России по Чувашской Республике - Чувашии",
        "УНДиПР Главного управления МЧС России по Республике Марий Эл",
        "УНДиПР ГУ МЧС России по Республике Татарстан",
    ]),
    ("ДФО", "Дальневосточный ФО", [
        "УНДиПР ГУ МЧС России по Чукотскому АО",
        "УНДиПР ГУ МЧС России по Забайкальскому краю",
        "УНДиПР ГУ МЧС России по Сахалинской области",
        "УНДиПР ГУ МЧС России по Приморскому краю",
        "УНДиПР ГУ МЧС России по Еврейской АО",
        "УНДиПР ГУ МЧС России по Камчатскому краю",
        "УНДиПР ГУ МЧС России по Республике Саха (Якутия)",
        "УНДиПР Главного управления МЧС России по Республике Бурятия",
        "УНДиПР ГУ МЧС России по Хабаровскому краю",
        "УНДиПР ГУ МЧС России по Магаданской области",
        "УНДиПР ГУ МЧС России по Амурской области",
    ]),
    ("СФО", "Сибирский ФО", [
        "УНДиПР ГУ МЧС России по Республике Тыва",
        "УНДиПР ГУ МЧС России по Новосибирской области",
        "УНДиПР ГУ МЧС России по Кемеровской области - Кузбассу",
        "УНДиПР ГУ МЧС России по Красноярскому краю",
        "УНДиПР ГУ МЧС России по Томской области",
        "УНДиПР ГУ МЧС России по Республике Алтай",
        "УНДиПР ГУ МЧС России по Омской области",
        "УНДиПР ГУ МЧС России по Республике Хакасия",
        "УНДиПР ГУ МЧС России по Алтайскому краю",
        "УНДиПР ГУ МЧС России по Иркутской области",
    ]),
    ("УФО", "Уральский ФО", [
        "УНДиПР ГУ МЧС России по Курганской области",
        "УНДиПР ГУ МЧС России по Свердловской области",
        "УНДиПР ГУ МЧС России по Челябинской области",
        "УНДиПР ГУ МЧС России по Ямало-Ненецкому АО",
        "УНДиПР ГУ МЧС России по Ханты-Мансийскому АО - Югре",
        "УНДиПР ГУ МЧС России по Тюменской области",
    ]),
]

NEW_REGIONS = [
    "УНДиПР ГУ МЧС России по Донецкой Народной Республике",
    "УНДиПР ГУ МЧС России по Запорожской области",
    "УНДиПР ГУ МЧС России по Херсонской области",
    "УНДиПР ГУ МЧС России по Луганской Народной Республике",
]

ALL_SUBJECTS = {subject for _, _, subjects in DISTRICTS for subject in subjects}
ALL_SUBJECTS.update(NEW_REGIONS)
CANONICAL_SUBJECTS = {
    re.sub(r"\s+", " ", subject.strip()).lower(): subject
    for subject in ALL_SUBJECTS
}


# ================== Вспомогательные функции ==================

def normalize_str(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip()).lower()


def find_column_index(headers, possible_names):
    headers_norm = [normalize_str(h) for h in headers]
    possible_norm = [normalize_str(n) for n in possible_names]
    for idx, norm in enumerate(headers_norm):
        if norm in possible_norm:
            return idx
    for idx, norm in enumerate(headers_norm):
        for pname in possible_norm:
            if pname in norm:
                return idx
    return None


def parse_date(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        value = value.strip()
        for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                continue
    return None


def load_data(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(
            f"Лист «{SHEET_NAME}» не найден.\n"
            f"Доступные листы: {', '.join(wb.sheetnames)}"
        )
    ws = wb[SHEET_NAME]
    headers = [cell.value if cell.value else "" for cell in ws[1]]

    col_indices = {}
    warnings_ = []
    missing = []

    for key, possible_names in COLUMN_KEYWORDS.items():
        idx = find_column_index(headers, possible_names)
        if idx is None:
            if key == "podrazd" and len(headers) > 17:
                idx = 17
                warnings_.append("Столбец «подразделение» не найден по имени — используем позицию 18 (индекс 17)")
            else:
                missing.append(f"«{key}» (искали: {possible_names})")
        col_indices[key] = idx

    if missing:
        raise ValueError("Не найдены обязательные столбцы:\n• " + "\n• ".join(missing))

    data = [
        row for row in ws.iter_rows(min_row=2, values_only=True)
        if not all(cell is None for cell in row)
    ]
    return data, col_indices, warnings_


def filter_by_date(data, col_idx, date_from, date_to):
    date_col = col_idx["date_act"]
    filtered = []
    skipped_out = 0
    skipped_inv = 0
    for row in data:
        parsed = parse_date(row[date_col])
        if parsed is None:
            skipped_inv += 1
            continue
        if date_from <= parsed <= date_to:
            filtered.append(row)
        else:
            skipped_out += 1
    return filtered, skipped_out, skipped_inv


def calculate_metrics_by_subject(data, col_idx):
    subj_col        = col_idx["subjekt"]
    knm_col         = col_idx["nom_knm"]
    vid_col         = col_idx["vid"]
    status_col      = col_idx["status"]
    proverka_col    = col_idx["proverka_ogv"]
    vid_nadzora_col = col_idx["vid_nadzora"]
    knd_col         = col_idx["knd"]
    nar_col         = col_idx["narusheniya"]
    vks_col         = col_idx["s_vks"]
    ssylki_col      = col_idx["ssylki"]

    allowed_vids = {"", "выездная проверка", "рейдовый осмотр", "инспекционный визит"}
    knm_info = {}

    for row in data:
        reasons_base = []
        if normalize_str(row[status_col]) != "завершена":
            reasons_base.append("status")
        if normalize_str(row[proverka_col]) != "нет":
            reasons_base.append("proverka")
        if normalize_str(row[vid_nadzora_col]) == "гнго":
            reasons_base.append("vid_nadzora")

        subject_raw  = str(row[subj_col]).strip() if row[subj_col] else ""
        subject_name = CANONICAL_SUBJECTS.get(normalize_str(subject_raw), "")
        knm = row[knm_col]
        if not row[subj_col] or not knm:
            reasons_base.append("empty")

        if reasons_base or subject_name not in ALL_SUBJECTS:
            continue

        vid_val   = normalize_str(row[vid_col])  if row[vid_col]  else ""
        knd_str   = normalize_str(row[knd_col])  if row[knd_col]  else ""
        nar_str   = normalize_str(row[nar_col])  if row[nar_col]  else ""
        vks_str   = normalize_str(row[vks_col])  if row[vks_col]  else ""
        ssylki_ok = row[ssylki_col] is not None and str(row[ssylki_col]).strip() != ""

        if knm not in knm_info:
            knm_info[knm] = {
                "subject":   subject_name,
                "vks_denom": False,
                "vks_num":   False,
                "och_denom": False,
                "och_num":   False,
                "och_nar":   False,
            }

        info = knm_info[knm]
        if vid_val in allowed_vids:
            info["vks_denom"] = True
            if vks_str == "да" and ssylki_ok:
                info["vks_num"] = True

        if vid_val in allowed_vids and "осмотр" in knd_str:
            info["och_denom"] = True
            if vks_str == "нет" and ssylki_ok:
                info["och_num"] = True
            if nar_str == "да":
                info["och_nar"] = True

    metrics = defaultdict(lambda: [set(), set(), set(), set(), set()])
    for knm, info in knm_info.items():
        subject = info["subject"]
        if info["vks_denom"]: metrics[subject][0].add(knm)
        if info["vks_num"]:   metrics[subject][1].add(knm)
        if info["och_denom"]: metrics[subject][2].add(knm)
        if info["och_num"]:   metrics[subject][3].add(knm)
        if info["och_nar"]:   metrics[subject][4].add(knm)

    return metrics


def make_subject_rows(subjects, metrics_year, metrics_week):
    rows = []
    for subject in subjects:
        y = metrics_year.get(subject, [set(), set(), set(), set(), set()])
        w = metrics_week.get(subject, [set(), set(), set(), set(), set()])
        rows.append({
            "subject":        subject,
            "vks_total_year": len(y[0]),
            "vks_mp_year":    len(y[1]),
            "vks_total_week": len(w[0]),
            "vks_mp_week":    len(w[1]),
            "och_total_year": len(y[2]),
            "och_mp_year":    len(y[3]),
            "och_nar_year":   len(y[4]),
            "och_total_week": len(w[2]),
            "och_mp_week":    len(w[3]),
            "och_nar_week":   len(w[4]),
        })
    return rows


def fmt_ratio(num, den):
    pct = num / den * 100 if den > 0 else 0.0
    return f"{num} ({pct:.2f}%)"


def fmt_och(total, nar):
    return f"{total} ({nar})"


def build_excel(district_rows, selected_date, week_start):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    blue_fill   = PatternFill("solid", start_color="4472C4", end_color="4472C4")
    yellow_fill = PatternFill("solid", start_color="FFFF00", end_color="FFFF00")
    red_fill    = PatternFill("solid", start_color="FF0000", end_color="FF0000")
    thin   = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def style_header_row(ws, row_num):
        for cell in ws[row_num]:
            cell.fill      = blue_fill
            cell.font      = Font(bold=True, color="FFFFFF", name="Arial", size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border    = border

    for sheet_name, district_name, rows in district_rows:
        ws = wb.create_sheet(title=sheet_name)

        ws.append(["Субъект", "ВКС", "", "", "", "ОЧНЫЕ", "", "", ""])
        ws.merge_cells("A1:A2")
        ws.merge_cells("B1:E1")
        ws.merge_cells("F1:I1")

        ws.append([
            "",
            "Всего в ААС КНД с начала года",
            "Всего в МП Инспектор с начала года",
            "Всего в ААС КНД за прошедшую неделю",
            "Всего в МП Инспектор за прошедшую неделю",
            "Всего в ААС КНД с начала года (из них с нарушениями)",
            "Всего в МП Инспектор с начала года",
            "Всего в ААС КНД за прошедшую неделю (из них с нарушениями)",
            "Всего в МП Инспектор за прошедшую неделю",
        ])

        style_header_row(ws, 1)
        style_header_row(ws, 2)
        ws.row_dimensions[1].height = 24
        ws.row_dimensions[2].height = 42

        ws.append([district_name, "", "", "", "", "", "", "", ""])
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=9)
        district_cell = ws.cell(3, 1)
        district_cell.font      = Font(bold=True, name="Arial", size=10)
        district_cell.alignment = Alignment(horizontal="center", vertical="center")

        for row_data in rows:
            ws.append([
                row_data["subject"],
                row_data["vks_total_year"],
                fmt_ratio(row_data["vks_mp_year"],  row_data["vks_total_year"]),
                row_data["vks_total_week"],
                fmt_ratio(row_data["vks_mp_week"],  row_data["vks_total_week"]),
                fmt_och(row_data["och_total_year"], row_data["och_nar_year"]),
                fmt_ratio(row_data["och_mp_year"],  row_data["och_total_year"]),
                fmt_och(row_data["och_total_week"], row_data["och_nar_week"]),
                fmt_ratio(row_data["och_mp_week"],  row_data["och_total_week"]),
            ])

        tv_y  = sum(r["vks_total_year"] for r in rows)
        tvmp_y= sum(r["vks_mp_year"]    for r in rows)
        tv_w  = sum(r["vks_total_week"] for r in rows)
        tvmp_w= sum(r["vks_mp_week"]    for r in rows)
        to_y  = sum(r["och_total_year"] for r in rows)
        tomp_y= sum(r["och_mp_year"]    for r in rows)
        tn_y  = sum(r["och_nar_year"]   for r in rows)
        to_w  = sum(r["och_total_week"] for r in rows)
        tomp_w= sum(r["och_mp_week"]    for r in rows)
        tn_w  = sum(r["och_nar_week"]   for r in rows)

        ws.append([
            f"Итого за {district_name}",
            tv_y,  fmt_ratio(tvmp_y, tv_y),
            tv_w,  fmt_ratio(tvmp_w, tv_w),
            fmt_och(to_y, tn_y), fmt_ratio(tomp_y, to_y),
            fmt_och(to_w, tn_w), fmt_ratio(tomp_w, to_w),
        ])
        total_row_idx = ws.max_row

        ws.column_dimensions["A"].width = 60.89
        ws.column_dimensions["B"].width = 9.33
        ws.column_dimensions["C"].width = 13.22
        ws.column_dimensions["D"].width = 8.89
        ws.column_dimensions["E"].width = 13.22
        ws.column_dimensions["F"].width = 13.44
        ws.column_dimensions["G"].width = 16.44
        ws.column_dimensions["H"].width = 9.44
        ws.column_dimensions["I"].width = 14.33

        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=9):
            for cell in row:
                cell.border = border
                if cell.row >= 4:
                    if cell.column == 1:
                        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                    else:
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    cell.font = Font(name="Arial", size=10)

        for row_idx, row_data in enumerate(rows, start=4):
            if row_data["vks_total_year"] > 0 and row_data["vks_mp_year"] / row_data["vks_total_year"] < 0.10:
                ws.cell(row_idx, 3).fill = yellow_fill
            if row_data["vks_total_week"] > 0 and row_data["vks_mp_week"] / row_data["vks_total_week"] < 0.10:
                ws.cell(row_idx, 5).fill = red_fill
            if row_data["och_total_year"] > 0 and row_data["och_mp_year"] / row_data["och_total_year"] < 0.80:
                ws.cell(row_idx, 7).fill = yellow_fill
            if row_data["och_total_week"] > 0 and row_data["och_mp_week"] / row_data["och_total_week"] < 0.80:
                ws.cell(row_idx, 9).fill = red_fill

        for cell in ws[total_row_idx]:
            cell.font      = Font(bold=True, name="Arial", size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.cell(total_row_idx, 1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        if tv_y  > 0 and tvmp_y / tv_y  < 0.10: ws.cell(total_row_idx, 3).fill = yellow_fill
        if tv_w  > 0 and tvmp_w / tv_w  < 0.10: ws.cell(total_row_idx, 5).fill = red_fill
        if to_y  > 0 and tomp_y / to_y  < 0.80: ws.cell(total_row_idx, 7).fill = yellow_fill
        if to_w  > 0 and tomp_w / to_w  < 0.80: ws.cell(total_row_idx, 9).fill = red_fill

        info_row = ws.max_row + 2
        ws.cell(info_row,     1, f"Дата отчёта: {selected_date.strftime('%d.%m.%Y')}")
        ws.cell(info_row + 1, 1, f"Прошедшая неделя: {week_start.strftime('%d.%m.%Y')} - {selected_date.strftime('%d.%m.%Y')}")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ================== Streamlit UI ==================

st.title("📊 МП Инспектор — Итоги по федеральным округам")
st.markdown("Загрузите файл выгрузки АИС КНД, выберите дату отчёта и получите готовый Excel по всем ФО.")

uploaded = st.file_uploader(
    "**Шаг 1 — Загрузите файл выгрузки (.xlsx / .xlsm)**",
    type=["xlsx", "xlsm"],
)

if uploaded:
    file_bytes = uploaded.read()

    with st.spinner("Читаем файл..."):
        try:
            data, col_idx, warnings_ = load_data(file_bytes)
        except Exception as e:
            st.error(f"Ошибка при загрузке: {e}")
            st.stop()

    for w in warnings_:
        st.warning(f"⚠️ {w}")

    st.success(f"✅ Файл загружен. Строк данных: **{len(data)}**")

    st.markdown("---")
    st.markdown("**Шаг 2 — Выберите дату отчёта**")
    st.caption("Период с начала года считается автоматически (01.01 → выбранная дата). Неделя — последние 7 дней включая выбранную дату.")

    selected_date = st.date_input("Дата отчёта", value=date.today())

    year_start = date(selected_date.year, 1, 1)
    week_start = selected_date - timedelta(days=6)

    st.info(
        f"📅 Период с начала года: **{year_start.strftime('%d.%m.%Y')}** — **{selected_date.strftime('%d.%m.%Y')}**  \n"
        f"📅 Прошедшая неделя: **{week_start.strftime('%d.%m.%Y')}** — **{selected_date.strftime('%d.%m.%Y')}**"
    )

    st.markdown("---")
    if st.button("🚀 Запустить расчёт", type="primary", use_container_width=True):

        with st.spinner("Фильтруем данные по периодам..."):
            ytd_data,  ytd_out,  ytd_inv  = filter_by_date(data, col_idx, year_start,  selected_date)
            week_data, week_out, week_inv = filter_by_date(data, col_idx, week_start,   selected_date)

        col1, col2 = st.columns(2)
        with col1:
            st.metric("Строк в периоде (год)",    len(ytd_data))
            st.caption(f"Вне периода: {ytd_out} | Без даты: {ytd_inv}")
        with col2:
            st.metric("Строк в периоде (неделя)", len(week_data))
            st.caption(f"Вне периода: {week_out} | Без даты: {week_inv}")

        if len(ytd_data) == 0 and len(week_data) == 0:
            st.error("Данных за оба периода нет. Проверьте файл и дату.")
            st.stop()

        with st.spinner("Считаем показатели..."):
            metrics_year = calculate_metrics_by_subject(ytd_data,  col_idx)
            metrics_week = calculate_metrics_by_subject(week_data, col_idx)

            district_rows = [
                (short, full, make_subject_rows(subjects, metrics_year, metrics_week))
                for short, full, subjects in DISTRICTS
            ]

        # Сводная таблица в интерфейсе
        st.markdown("---")
        st.subheader("📈 Сводка по федеральным округам")

        import pandas as pd
        summary = []
        for short_name, full_name, rows in district_rows:
            tv_y  = sum(r["vks_total_year"] for r in rows)
            tvmp_y= sum(r["vks_mp_year"]    for r in rows)
            to_y  = sum(r["och_total_year"] for r in rows)
            tomp_y= sum(r["och_mp_year"]    for r in rows)
            tv_w  = sum(r["vks_total_week"] for r in rows)
            tvmp_w= sum(r["vks_mp_week"]    for r in rows)
            to_w  = sum(r["och_total_week"] for r in rows)
            tomp_w= sum(r["och_mp_week"]    for r in rows)
            summary.append({
                "ФО":                  full_name,
                "ВКС год, всего":      tv_y,
                "ВКС год, с МП (%)":   f"{tvmp_y} ({tvmp_y/tv_y*100:.1f}%)" if tv_y else "—",
                "ВКС нед., всего":     tv_w,
                "ВКС нед., с МП (%)":  f"{tvmp_w} ({tvmp_w/tv_w*100:.1f}%)" if tv_w else "—",
                "Очные год, всего":    to_y,
                "Очные год, с МП (%)": f"{tomp_y} ({tomp_y/to_y*100:.1f}%)" if to_y else "—",
                "Очные нед., всего":   to_w,
                "Очные нед., с МП (%)":f"{tomp_w} ({tomp_w/to_w*100:.1f}%)" if to_w else "—",
            })
        st.dataframe(pd.DataFrame(summary), use_container_width=True, hide_index=True)

        with st.spinner("Формируем Excel-файл..."):
            excel_bytes = build_excel(district_rows, selected_date, week_start)

        filename = f"итоги_фо_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.success("✅ Готово!")
        st.download_button(
            label="⬇️ Скачать результат (.xlsx)",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
