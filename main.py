"""
Сервис обработки отчёта по инвентаризации для n8n.

Принимает xlsx-файл (POST /process, multipart/form-data, поле "file"),
делает свод по листу "Данные КРГ" в разрезе Товарной группы (Дельта,
СуммаРозница, СуммаПриход), записывает свод в столбцы T/U/V листа
"Отчет по ТГ" и прибавляет эти значения к столбцам I/J/K.
Формулы и Excel-таблица на листе "Отчет по ТГ" не трогаются и остаются
рабочими (openpyxl меняет только конкретные ячейки-значения).

Запуск локально:  uvicorn main:app --host 0.0.0.0 --port 8000
"""

from io import BytesIO

import openpyxl
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import StreamingResponse

app = FastAPI(title="Inventory Report Processor")

SHEET_DATA = "Данные КРГ"
SHEET_REPORT = "Отчет по ТГ"

# 0-based индексы столбцов на листе "Данные КРГ"
COL_DATA_MAGAZIN = 0   # A
COL_DATA_TG = 1        # B
COL_DATA_DELTA = 11    # L
COL_DATA_SUM_ROZ = 12  # M
COL_DATA_SUM_PRIH = 13  # N

# 0-based индексы столбцов на листе "Отчет по ТГ"
COL_REPORT_MAGAZIN = 0   # A
COL_REPORT_TG = 2        # C
COL_REPORT_I = 8         # I - Фактический остаток, шт.
COL_REPORT_J = 9         # J - Фактический остаток, розн. руб.
COL_REPORT_K = 10        # K - Фактический остаток, приходные руб.
COL_REPORT_T = 19        # T - Дельта КРГ, шт
COL_REPORT_U = 20        # U - Дельта КРГ, розн. руб.
COL_REPORT_V = 21        # V - Дельта КРГ, приходные руб.


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/process")
async def process_report(file: UploadFile = File(...)):
    if not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(400, "Ожидается файл .xlsx")

    contents = await file.read()

    try:
        wb = openpyxl.load_workbook(BytesIO(contents), data_only=False)
    except Exception as e:
        raise HTTPException(400, f"Не удалось открыть файл: {e}")

    for sheet in (SHEET_DATA, SHEET_REPORT):
        if sheet not in wb.sheetnames:
            raise HTTPException(
                400, f"На листе не найдена вкладка '{sheet}'. Есть: {wb.sheetnames}"
            )

    ws_data = wb[SHEET_DATA]
    ws_report = wb[SHEET_REPORT]

    # 1. Свод по (Магазин, Товарная группа)
    pivot = {}
    for row in ws_data.iter_rows(min_row=2, values_only=False):
        magazin = row[COL_DATA_MAGAZIN].value
        tg = row[COL_DATA_TG].value
        if tg is None:
            continue
        delta = row[COL_DATA_DELTA].value or 0
        sum_roz = row[COL_DATA_SUM_ROZ].value or 0
        sum_prih = row[COL_DATA_SUM_PRIH].value or 0

        key = (magazin, tg)
        agg = pivot.setdefault(key, {"delta": 0.0, "sum_roz": 0.0, "sum_prih": 0.0})
        agg["delta"] += float(delta)
        agg["sum_roz"] += float(sum_roz)
        agg["sum_prih"] += float(sum_prih)

    # 2. Проход по "Отчет по ТГ": запись T/U/V, прибавление к I/J/K
    current_magazin = None
    rows_updated = 0
    for row in ws_report.iter_rows(min_row=2):
        magazin_cell = row[COL_REPORT_MAGAZIN].value
        tg = row[COL_REPORT_TG].value

        if magazin_cell:
            current_magazin = magazin_cell

        # пропускаем пустые и итоговые строки ("Итого по магазину ...")
        if not tg or "Итого" in str(tg):
            continue

        key = (current_magazin, tg)
        agg = pivot.get(key, {"delta": 0.0, "sum_roz": 0.0, "sum_prih": 0.0})

        row[COL_REPORT_T].value = agg["delta"]
        row[COL_REPORT_U].value = agg["sum_roz"]
        row[COL_REPORT_V].value = agg["sum_prih"]

        i_val = row[COL_REPORT_I].value
        j_val = row[COL_REPORT_J].value
        k_val = row[COL_REPORT_K].value
        i_val = i_val if isinstance(i_val, (int, float)) else 0
        j_val = j_val if isinstance(j_val, (int, float)) else 0
        k_val = k_val if isinstance(k_val, (int, float)) else 0

        row[COL_REPORT_I].value = i_val + agg["delta"]
        row[COL_REPORT_J].value = j_val + agg["sum_roz"]
        row[COL_REPORT_K].value = k_val + agg["sum_prih"]

        rows_updated += 1

    # заставляем Excel пересчитать формулы (N-S) при открытии файла пользователем
    wb.calculation.fullCalcOnLoad = True

    out = BytesIO()
    wb.save(out)
    out.seek(0)

    out_filename = file.filename.rsplit(".", 1)[0] + "_готово.xlsx"
    # кириллица в имени файла не кодируется как latin-1 в обычном
    # filename=, поэтому используем filename* (RFC 5987) + ASCII-фолбэк
    from urllib.parse import quote

    ascii_fallback = "report_done.xlsx"
    encoded_name = quote(out_filename)

    return StreamingResponse(
        out,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": (
                f'attachment; filename="{ascii_fallback}"; '
                f"filename*=UTF-8''{encoded_name}"
            ),
            "X-Rows-Updated": str(rows_updated),
        },
    )
