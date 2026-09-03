from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
import openpyxl
from io import BytesIO
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI()

def find_col(ws, header_name):
    """Ищет номер столбца по названию заголовка"""
    logger.info(f"Ищу столбец: {header_name}")
    for cell in ws[1]:
        if cell.value and str(cell.value).strip().lower() == header_name.strip().lower():
            logger.info(f"Найден столбец {header_name} в позиции {cell.column}")
            return cell.column
    logger.warning(f"Столбец {header_name} НЕ НАЙДЕН!")
    return None

@app.post("/process")
async def process_excel(file: UploadFile = File(...)):
    try:
        logger.info(f"📥 Получен файл: {file.filename}, размер: {len(await file.read())} байт")
        contents = await file.read()
        
        # ПРОХОД 1: Читаем значения
        logger.info("🔄 Открываю файл для чтения (data_only=True)...")
        wb_vals = openpyxl.load_workbook(BytesIO(contents), data_only=True)
        logger.info(f" Листы: {wb_vals.sheetnames}")
        
        logger.info("🔍 Открываю лист 'Отчет по ТГ'...")
        ws3_vals = wb_vals["Отчет по ТГ"]
        
        logger.info("📊 Читаю существующие значения I, J, K...")
        tg_col_3v = find_col(ws3_vals, "Товарная группа")
        existing_values = {}
        
        if tg_col_3v:
            for row in ws3_vals.iter_rows(min_row=2):
                tg_cell = row[tg_col_3v - 1]
                if not tg_cell.value:
                    continue
                tg_name = str(tg_cell.value).strip()
                val_i = row[8].value if len(row) >= 9 else 0
                val_j = row[9].value if len(row) >= 10 else 0
                val_k = row[10].value if len(row) >= 11 else 0
                
                existing_values[tg_name] = {
                    "i": float(val_i) if val_i is not None else 0,
                    "j": float(val_j) if val_j is not None else 0,
                    "k": float(val_k) if val_k is not None else 0,
                }
        logger.info(f"✅ Прочитано {len(existing_values)} строк из 'Отчет по ТГ'")
        wb_vals.close()

        # ПРОХОД 2: Модификация
        logger.info("🔄 Открываю файл для записи (data_only=False)...")
        wb = openpyxl.load_workbook(BytesIO(contents), data_only=False)
        
        logger.info(" Обрабатываю лист 'Данные КРГ'...")
        ws1 = wb["Данные КРГ"]
        tg_col_1 = find_col(ws1, "Товарная группа")
        delta_col = find_col(ws1, "Дельта, шт.")
        retail_col = find_col(ws1, "Сумма дельта средневзвешенная, руб.")
        purchase_col = find_col(ws1, "Сумма дельта приходная, руб.")

        summary = {}
        if tg_col_1:
            for row in ws1.iter_rows(min_row=2):
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
        
        logger.info(f"✅ Сводная по {len(summary)} товарным группам")

        # Запись данных
        logger.info("✏️ Обновляю лист 'Отчет по ТГ'...")
        ws2 = wb["Отчет по ТГ"]
        tg_col_2 = find_col(ws2, "Товарная группа")

        if tg_col_2:
            updated_count = 0
            for row in ws2.iter_rows(min_row=2):
                tg_cell = row[tg_col_2 - 1]
                if not tg_cell.value:
                    continue
                tg_name = str(tg_cell.value).strip()
                
                if tg_name in summary:
                    if len(row) >= 22:
                        row[19].value = summary[tg_name]["delta"]
                        row[20].value = summary[tg_name]["retail"]
                        row[21].value = summary[tg_name]["purchase"]
                    
                    base = existing_values.get(tg_name, {"i": 0, "j": 0, "k": 0})
                    if len(row) >= 11:
                        row[8].value = base["i"] + summary[tg_name]["delta"]
                        row[9].value = base["j"] + summary[tg_name]["retail"]
                        row[10].value = base["k"] + summary[tg_name]["purchase"]
                    updated_count += 1
            
            logger.info(f"✅ Обновлено {updated_count} строк")

        # Сохранение
        logger.info("💾 Сохраняю файл...")
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        wb.close()
        logger.info("✅ Файл сохранен успешно!")

        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": 'attachment; filename="Result.xlsx"'}
        )
    
    except Exception as e:
        logger.error(f"❌ КРИТИЧЕСКАЯ ОШИБКА: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        raise
