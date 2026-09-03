from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
import openpyxl
from io import BytesIO

app = FastAPI()


def find_col(ws, header_name):
    """Ищет номер столбца по названию заголовка в первой строке"""
    for cell in ws[1]:
        if cell.value and str(cell.value).strip() == header_name:
            return cell.column
    return None


@app.post("/process")
async def process_excel(file: UploadFile = File(...)):
    contents = await file.read()

    # ============================================================
    # ПРОХОД 1: Читаем вычисленные значения формул (data_only=True)
    # Нужен для ячеек I, J, K листа "Готовый отчет по ТГ",
    # если там формулы — получим их числовые результаты
    # ============================================================
    wb_vals = openpyxl.load_workbook(BytesIO(contents), data_only=True)
    ws3_vals = wb_vals["Готовый отчет по ТГ"]

    tg_col_3v = find_col(ws3_vals, "Товарная группа")
    existing_values = {}

    if tg_col_3v:
        for row in ws3_vals.iter_rows(min_row=2):
            tg_cell = row[tg_col_3v - 1]  # openpyxl: 1-based, list: 0-based
            if not tg_cell.value:
                continue
            tg_name = str(tg_cell.value).strip()
            val_i = row[8].value   # Столбец I (9-й, индекс 8)
            val_j = row[9].value   # Столбец J (10-й, индекс 9)
            val_k = row[10].value  # Столбец K (11-й, индекс 10)
            existing_values[tg_name] = {
                "i": float(val_i) if val_i is not None else 0,
                "j": float(val_j) if val_j is not None else 0,
                "k": float(val_k) if val_k is not None else 0,
            }
    wb_vals.close()

    # ============================================================
    # ПРОХОД 2: Открываем с data_only=False (формулы = строки)
    # Модифицируем нужные ячейки. Все формулы, которые мы НЕ трогаем,
    # физически остаются в файле неизменными.
    # ============================================================
    wb = openpyxl.load_workbook(BytesIO(contents), data_only=False)

    # --- ЛИСТ 1: "Данные КРГ" → Сводная по ТГ ---
    ws1 = wb["Данные КРГ"]

    tg_col_1 = find_col(ws1, "Товарная группа")
    delta_col = find_col(ws1, "Дельта")          # ⚠️ Подставь точное название заголовка!
    retail_col = find_col(ws1, "СуммаРозница")    # ⚠️ Подставь точное название заголовка!
    purchase_col = find_col(ws1, "СуммаПриход")   # ⚠️ Подставь точное название заголовка!

    summary = {}

    for row in ws1.iter_rows(min_row=2):
        if not tg_col_1:
            continue
        tg_cell = row[tg_col_1 - 1]
        if not tg_cell.value:
            continue
        tg_name = str(tg_cell.value).strip()

        def safe_float(cell):
            try:
                return float(cell.value) if cell.value is not None else 0
            except (ValueError, TypeError):
                return 0

        d = safe_float(row[delta_col - 1]) if delta_col else 0
        r = safe_float(row[retail_col - 1]) if retail_col else 0
        p = safe_float(row[purchase_col - 1]) if purchase_col else 0

        if tg_name not in summary:
            summary[tg_name] = {"delta": 0, "retail": 0, "purchase": 0}
        summary[tg_name]["delta"] += d
        summary[tg_name]["retail"] += r
        summary[tg_name]["purchase"] += p

    # --- ЛИСТ 2: "Отчет по ТГ" → Вставка в столбцы T(20), U(21), V(22) ---
    ws2 = wb["Отчет по ТГ"]
    tg_col_2 = find_col(ws2, "Товарная группа")

    if tg_col_2:
        for row in ws2.iter_rows(min_row=2):
            tg_cell = row[tg_col_2 - 1]
            if not tg_cell.value:
                continue
            tg_name = str(tg_cell.value).strip()
            if tg_name in summary:
                row[19].value = summary[tg_name]["delta"]     # T = 20, индекс 19
                row[20].value = summary[tg_name]["retail"]    # U = 21, индекс 20
                row[21].value = summary[tg_name]["purchase"]  # V = 22, индекс 21

    # --- ЛИСТ 3: "Готовый отчет по ТГ" → Прибавление к I(9), J(10), K(11) ---
    ws3 = wb["Готовый отчет по ТГ"]
    tg_col_3 = find_col(ws3, "Товарная группа")

    if tg_col_3:
        for row in ws3.iter_rows(min_row=2):
            tg_cell = row[tg_col_3 - 1]
            if not tg_cell.value:
                continue
            tg_name = str(tg_cell.value).strip()
            if tg_name in summary:
                base = existing_values.get(tg_name, {"i": 0, "j": 0, "k": 0})
                row[8].value = base["i"] + summary[tg_name]["delta"]      # I = 9
                row[9].value = base["j"] + summary[tg_name]["retail"]     # J = 10
                row[10].value = base["k"] + summary[tg_name]["purchase"]  # K = 11

    # ============================================================
    # Сохраняем и возвращаем файл
    # ============================================================
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    wb.close()

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="Result.xlsx"'}
    )
