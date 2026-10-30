import customtkinter as ctk
from tkinter import filedialog
import os
import re
import glob
import pandas as pd
import fitz
import threading
import openpyxl
from pypdf import PdfReader

# Настройки внешнего вида
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

# --- ГЛАВНАЯ ФУНКЦИЯ ПРОЦЕССА СВЕРКИ, СОРТИРОВКИ PDF И 1С (ВКЛАДКА 1) ---
def process_full_workflow(scans_excel_path, direct_scans_text, pdf_dir, output_dir, check_excel_path, log_callback, progress_callback):
    
    def clean_val(v):
        if pd.isna(v): return ""
        s = str(v)
        s = re.sub(r'[\s\xa0\u200b\ufeff]', '', s)
        if s.endswith('.0'): s = s[:-2]
        if s.lower() == 'nan': return ""
        return s

    def fix_layout(s):
        if not re.search(r'[а-яА-ЯёЁ]', s):
            return s
        ru_chars = 'йцукенгшщзхъфывапролджэячсмитьбюЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭЯЧСМИТЬБЮёЁ"№;:?,.'
        en_chars = 'qwertyuiop[]asdfghjkl;\'zxcvbnm,.QWERTYUIOP{}ASDFGHJKL:"ZXCVBNM<>`~@#$^&?/'
        return s.translate(str.maketrans(ru_chars, en_chars))

    # 1. Загружаем справочник (УМНЫЙ ПОИСК КОЛОНОК ПО НАЗВАНИЯМ)
    map_gtin = {}
    map_common = {}
    map_ozn = {}
    
    if check_excel_path and os.path.exists(check_excel_path):
        log_callback(f"Загружаем справочник: {os.path.basename(check_excel_path)}")
        try:
            df_db_raw = pd.read_excel(check_excel_path, sheet_name=0, header=None, dtype=str)
            
            header_row_idx = -1
            col_map = {}
            
            # Сканируем первые 10 строк, чтобы найти шапку таблицы
            for idx, row in df_db_raw.head(10).iterrows():
                row_vals_full = [str(x).strip().lower() if pd.notna(x) else "" for x in row.values]
                if 'баркод' in row_vals_full and 'gtin' in row_vals_full:
                    header_row_idx = idx
                    for col_idx, val in enumerate(row_vals_full):
                        if val == 'артикул продавца': col_map['Seller_Art'] = col_idx
                        elif val == 'артикул wb': col_map['WB_Art'] = col_idx
                        elif val == 'баркод': col_map['Barcode_Common'] = col_idx
                        elif val == 'gtin': col_map['GTIN'] = col_idx
                        elif val == 'название товара': col_map['Name'] = col_idx
                        elif val == 'шк 2': col_map['SK_2'] = col_idx
                        elif val == 'артикул': col_map['Ozon_Art'] = col_idx
                    break
                    
            if header_row_idx == -1:
                log_callback("❌ Ошибка: Не найдена строка с заголовками (Баркод, GTIN). Проверьте структуру файла!")
                self.safe_stage_qr("Ошибка чтения базы")
                return
                
            log_callback("Шапка таблицы найдена. Индексируем товары...")
            
            # Читаем данные, начиная со строки ПОСЛЕ найденной шапки
            for idx, row in df_db_raw.iloc[header_row_idx+1:].iterrows():
                def get_val(col_key):
                    c_idx = col_map.get(col_key)
                    if c_idx is not None and c_idx < len(row):
                        return clean_val(row.iloc[c_idx])
                    return ""
                
                b_common = get_val('Barcode_Common')
                gtin = get_val('GTIN')
                seller_art = get_val('Seller_Art')
                
                if gtin and len(gtin) in [13, 14]:
                    gtin = gtin.zfill(14)
                ozn = get_val('SK_2')
                
                # Если строка пустая (нет даже баркода), пропускаем её
                if not b_common and not gtin:
                    continue
                
                info = {
                    'barcode_common': b_common,
                    'gtin': gtin,
                    'seller_art': seller_art,
                    'wb_art': get_val('WB_Art'),
                    'ozon_art': get_val('Ozon_Art'),
                    'name': get_val('Name'),
                    'sk_2': ozn
                }
                
                if b_common: map_common[b_common] = info
                if gtin: map_gtin[gtin] = info
                if ozn: map_ozn[ozn] = info
                
            log_callback(f"Справочник успешно загружен (уникальных позиций: {len(map_common)}).")
        except Exception as e:
            log_callback(f"Ошибка при чтении файла сверки: {e}")

    # 2. Читаем отсканированные данные (Прямой ввод ИЛИ из Excel)
    scanned_sequence = []
    
    if direct_scans_text.strip():
        log_callback("Чтение данных из поля прямого сканирования...")
        raw_lines = direct_scans_text.split('\n')
        for line in raw_lines:
            v_str = clean_val(line)
            if v_str:
                v_str = fix_layout(v_str)
                scanned_sequence.append(v_str)
        
        try:
            raw_out_path = os.path.join(output_dir, "My_Raw_Scans.xlsx")
            wb_raw = openpyxl.Workbook()
            ws_raw = wb_raw.active
            ws_raw.title = "Сырые сканы"
            ws_raw.append(["Отсканированный Штрихкод / КИЗ"])
            for s in scanned_sequence:
                cell = ws_raw.cell(row=ws_raw.max_row + 1, column=1, value=str(s))
                cell.number_format = '@' 
            wb_raw.save(raw_out_path)
            log_callback(f"Ваши сканы сохранены в отдельный файл: My_Raw_Scans.xlsx")
        except Exception as e:
            log_callback(f"Ошибка при сохранении сырых сканов: {e}")
            
    elif scans_excel_path and os.path.exists(scans_excel_path):
        log_callback("Чтение файла с отсканированными данными...")
        try:
            xls_scans = pd.read_excel(scans_excel_path, sheet_name=None, header=None, dtype=str)
            for _, s_df in xls_scans.items():
                for _, row in s_df.iterrows():
                    for val in row.dropna():
                        v_str = clean_val(val)
                        if v_str:
                            v_str = fix_layout(v_str)
                            scanned_sequence.append(v_str)
            log_callback(f"Считано {len(scanned_sequence)} записей из файла сканов.")
        except Exception as e:
            log_callback(f"Ошибка при чтении файла сканов: {e}")
    else:
        log_callback("ВНИМАНИЕ: Нет данных для обработки (поле пустое и файл не выбран).")
        return

    # 2.5 Группируем сканы в пары
    paired_scans = []
    temp_pair = {'kiz': "", 'barcode': ""}
    
    for token in scanned_sequence:
        if "отсканированный" in str(token).lower() or "штрихкод" in str(token).lower() or "киз" in str(token).lower():
            continue

        clean_t = re.sub(r'\s+', '', token).replace('(01)', '01').replace('(21)', '21')
        is_kiz = bool(re.search(r'01\d{14}21', clean_t))
        
        if is_kiz:
            if temp_pair['kiz']: 
                paired_scans.append(temp_pair)
                temp_pair = {'kiz': "", 'barcode': ""}
            temp_pair['kiz'] = clean_t
        else:
            if temp_pair['barcode']: 
                paired_scans.append(temp_pair)
                temp_pair = {'kiz': "", 'barcode': ""}
            temp_pair['barcode'] = clean_t
            
        if temp_pair['kiz'] and temp_pair['barcode']:
            paired_scans.append(temp_pair)
            temp_pair = {'kiz': "", 'barcode': ""}
            
    if temp_pair['kiz'] or temp_pair['barcode']:
        paired_scans.append(temp_pair)

    log_callback(f"Сформировано пар (КИЗ + Баркод): {len(paired_scans)}")

    # 2.6 ЖЕЛЕЗНАЯ ЛОГИКА ПОИСКА И ОТСЛЕЖИВАНИЯ ДУБЛИКАТОВ
    records_for_1c = []
    lost_items = []
    
    pdf_kiz_masks = {} 
    matched_for_1c = 0
    kiz_seen = {} 

    for scan_idx, pair in enumerate(paired_scans, start=1):
        s_kiz = pair['kiz']
        s_bar = pair['barcode']
        
        info = None
        lookup_gtin = ""
        short_serial = ""
        
        is_pair_dup = False
        is_kiz_only_dup = False
        
        if s_kiz:
            if s_kiz in kiz_seen:
                kiz_seen[s_kiz]['count'] += 1
                first_bar = kiz_seen[s_kiz]['barcode']
                if s_bar == first_bar:
                    is_pair_dup = True
                else:
                    is_kiz_only_dup = True
            
            m = re.search(r'01(\d{14})21(.{1,15})', s_kiz)
            if m:
                lookup_gtin = m.group(1)
                serial_raw = m.group(2)[:13]
                serial_clean = re.sub(r'[^a-zA-Z0-9]', '', serial_raw).lower()
                short_serial = serial_clean[:6] 
                
                if not (is_pair_dup or is_kiz_only_dup):
                    if lookup_gtin not in pdf_kiz_masks:
                        pdf_kiz_masks[lookup_gtin] = []
                    pdf_kiz_masks[lookup_gtin].append({'serial': short_serial, 'row': scan_idx})
                
        if s_bar in map_common: info = map_common[s_bar]
        elif s_bar in map_gtin: info = map_gtin[s_bar]
        elif s_bar.zfill(14) in map_gtin: info = map_gtin[s_bar.zfill(14)]
        elif s_bar in map_ozn: info = map_ozn[s_bar]
        elif lookup_gtin in map_gtin: info = map_gtin[lookup_gtin]
            
        if info:
            b_common = info['barcode_common'] if info['barcode_common'] else "НЕТ ОБЩЕГО БАРКОДА В БАЗЕ"
            
            records_for_1c.append({
                'Отсканированный ШК': s_bar if s_bar else "НЕТ",
                'Баркод общий (для 1С)': b_common,
                'ШК 2 (Ozon)': info['sk_2'],
                'КИЗ (Честный ЗНАК)': s_kiz,
                'Артикул продавца': info['seller_art'],
                'Артикул WB': info['wb_art'],
                'Артикул Ozon': info['ozon_art'],
                'Наименование товара': info['name'],
                'Количество': 1,
                'no_kiz': not bool(s_kiz),
                'no_barcode': not bool(s_bar),
                'is_pair_dup': is_pair_dup,
                'is_kiz_only_dup': is_kiz_only_dup
            })
            matched_for_1c += 1
            
            if s_kiz:
                if not (is_pair_dup or is_kiz_only_dup):
                    kiz_seen[s_kiz] = {'row': scan_idx, 'barcode': s_bar, 'count': 1}
                else:
                    first_r = kiz_seen[s_kiz]['row']
                    if is_pair_dup:
                        lost_items.append({
                            'Причина': f"Дубликат ПАРЫ (Скан №{scan_idx})",
                            'Штрихкод / GTIN': b_common,
                            'Детали': f"Оригинал был в скане №{first_r}. Точный повтор."
                        })
                    else:
                        lost_items.append({
                            'Причина': f"Конфликт КИЗ (Скан №{scan_idx})",
                            'Штрихкод / GTIN': b_common,
                            'Детали': f"Оригинал в скане №{first_r}. КИЗ совпал, но Баркод другой!"
                        })
        else:
            if s_kiz and (is_pair_dup or is_kiz_only_dup):
                first_r = kiz_seen[s_kiz]['row']
                lost_items.append({
                    'Причина': f"Дубликат (Скан №{scan_idx}) + Нет в базе",
                    'Штрихкод / GTIN': s_bar if s_bar else lookup_gtin,
                    'Детали': f"Оригинал был в скане №{first_r}. КИЗ: {s_kiz}"
                })
            else:
                lost_items.append({
                    'Причина': f"Не найден в базе (Скан №{scan_idx})",
                    'Штрихкод / GTIN': s_bar if s_bar else lookup_gtin,
                    'Детали': s_kiz
                })

    log_callback(f"Успешно сопоставлено для 1С: {matched_for_1c}")
    
    total_kizs = sum(len(v) for v in pdf_kiz_masks.values())
    log_callback(f"Уникальных КИЗов для вырезки из PDF: {total_kizs} шт.")
    log_callback("Начинаем сканирование и сортировку PDF-файлов...\n")

    # 3. Сканируем PDF-файлы
    pdf_files = glob.glob(os.path.join(pdf_dir, '*.pdf'))
    pdf_files = [f for f in pdf_files if not os.path.basename(f).startswith('Matched_') and not os.path.basename(f).startswith('All_Matched_')]
    
    doc_male = fitz.open()
    doc_female = fitz.open()
    doc_other = fitz.open()
    
    male_count = 0
    female_count = 0
    other_count = 0
    
    for i, pdf_path in enumerate(pdf_files):
        filename = os.path.basename(pdf_path)
        progress_callback(i + 1, len(pdf_files), filename)
        
        try:
            doc = fitz.open(pdf_path)
            new_doc = fitz.open()
            file_matched_count = 0
            
            for page_num in range(len(doc)):
                if sum(len(v) for v in pdf_kiz_masks.values()) == 0:
                    break 
                    
                page = doc[page_num]
                raw_text = page.get_text()
                clean_text = re.sub(r'[^a-zA-Z0-9]', '', raw_text).lower()
                
                is_match = False
                
                for target_gtin, serials in pdf_kiz_masks.items():
                    if target_gtin.lower() in clean_text:
                        matched_serial_info = None
                        for s_info in serials:
                            if s_info['serial'] in clean_text:
                                is_match = True
                                matched_serial_info = s_info
                                break
                        if matched_serial_info:
                            serials.remove(matched_serial_info)
                            break
                            
                if is_match:
                    new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
                    text_lower = raw_text.lower()
                    file_lower = filename.lower()
                    if 'женск' in text_lower or 'жен.' in text_lower or 'w-' in text_lower or 'женск' in file_lower:
                        doc_female.insert_pdf(doc, from_page=page_num, to_page=page_num)
                        female_count += 1
                    elif 'мужск' in text_lower or 'муж.' in text_lower or 'мужск' in file_lower:
                        doc_male.insert_pdf(doc, from_page=page_num, to_page=page_num)
                        male_count += 1
                    else:
                        doc_other.insert_pdf(doc, from_page=page_num, to_page=page_num)
                        other_count += 1
                    file_matched_count += 1
                            
            if file_matched_count > 0:
                output_filename = f'Matched_{filename}'
                new_doc.save(os.path.join(output_dir, output_filename))
                log_callback(f"Успех! Извлечено {file_matched_count} нужных стр. из файла {filename}")
                
            new_doc.close()
            doc.close()
                
        except Exception as e:
            log_callback(f"Ошибка в файле {filename}: {e}")
            
    if male_count > 0:
        doc_male.save(os.path.join(output_dir, "All_Matched_Мужские.pdf"))
        log_callback(f"\n[PDF ИТОГ] Мужские этикетки ({male_count} стр.) -> All_Matched_Мужские.pdf")
    if female_count > 0:
        doc_female.save(os.path.join(output_dir, "All_Matched_Женские.pdf"))
        log_callback(f"[PDF ИТОГ] Женские этикетки ({female_count} стр.) -> All_Matched_Женские.pdf")
    if other_count > 0:
        doc_other.save(os.path.join(output_dir, "All_Matched_Неопределенные.pdf"))
        log_callback(f"[PDF ИТОГ] Неопределенные ({other_count} стр.) -> All_Matched_Неопределенные.pdf")
        
    doc_male.close()
    doc_female.close()
    doc_other.close()

    for target_gtin, serials in pdf_kiz_masks.items():
        for s_info in serials:
            lost_items.append({
                'Причина': f"Нет в PDF (Скан № {s_info['row']})",
                'Штрихкод / GTIN': target_gtin,
                'Детали': f"Хвост серийника: {s_info['serial']}"
            })

    # 4. Генерируем файл для 1С И ВКЛАДКУ ПОТЕРЯШЕК
    if check_excel_path and (records_for_1c or lost_items):
        output_1c = os.path.join(output_dir, "For_1C_Upload.xlsx")
        wb_1c = openpyxl.Workbook()
        
        ws1 = wb_1c.active
        ws1.title = "Детализированный"
        ws1.append(['Отсканированный ШК', 'Баркод общий (для 1С)', 'ШК 2 (Ozon)', 'КИЗ (Честный ЗНАК)', 'Артикул продавца', 'Артикул WB', 'Артикул Ozon', 'Наименование товара', 'Количество'])
        
        fill_no_kiz = openpyxl.styles.PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid") 
        fill_no_barcode = openpyxl.styles.PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid") 
        fill_kiz_dup = openpyxl.styles.PatternFill(start_color="FF9900", end_color="FF9900", fill_type="solid") 
        fill_pair_dup = openpyxl.styles.PatternFill(start_color="FF00FF", end_color="FF00FF", fill_type="solid") 
        
        for i, item in enumerate(records_for_1c, start=2):
            ws1.append([
                item['Отсканированный ШК'], item['Баркод общий (для 1С)'], item['ШК 2 (Ozon)'], 
                item['КИЗ (Честный ЗНАК)'], item['Артикул продавца'], item['Артикул WB'], 
                item['Артикул Ozon'], item['Наименование товара'], 1
            ])
            
            fill_to_apply = None
            if item.get('is_pair_dup'):
                fill_to_apply = fill_pair_dup
            elif item.get('is_kiz_only_dup'):
                fill_to_apply = fill_kiz_dup
            elif item.get('no_barcode'):
                fill_to_apply = fill_no_barcode
            elif item.get('no_kiz'):
                fill_to_apply = fill_no_kiz
                
            if fill_to_apply:
                for col in range(1, 10):
                    ws1.cell(row=i, column=col).fill = fill_to_apply
            
        ws2 = wb_1c.create_sheet(title="Для загрузки в 1С")
        ws2.append(['Штрихкод', 'Количество'])
        
        if records_for_1c:
            df_det = pd.DataFrame(records_for_1c)
            df_valid_all = df_det[(df_det['no_barcode'] == False) & (df_det['Баркод общий (для 1С)'] != "НЕТ ОБЩЕГО БАРКОДА В БАЗЕ")]
            
            if not df_valid_all.empty:
                df_good = df_valid_all[df_valid_all['is_pair_dup'] == False]
                df_pair_dups = df_valid_all[df_valid_all['is_pair_dup'] == True]
                
                if not df_good.empty:
                    df_summary = df_good.groupby('Баркод общий (для 1С)')['Количество'].sum().reset_index()
                    for _, row in df_summary.iterrows():
                        ws2.append([row['Баркод общий (для 1С)'], row['Количество']])
                
                if not df_pair_dups.empty:
                    ws2.append([]) 
                    ws2.append(['--- ВНИМАНИЕ: ОТКЛОНЕННЫЕ ДУБЛИКАТЫ ПАР ---', ''])
                    warn_cell = ws2.cell(row=ws2.max_row, column=1)
                    warn_cell.font = openpyxl.styles.Font(bold=True, color="FF0000")
                    
                    df_dup_summary = df_pair_dups.groupby('Баркод общий (для 1С)')['Количество'].sum().reset_index()
                    for _, row in df_dup_summary.iterrows():
                        ws2.append([row['Баркод общий (для 1С)'], row['Количество']])

        ws3 = wb_1c.create_sheet(title="Потеряшки")
        ws3.append(['Причина', 'Штрихкод / GTIN', 'Детали'])
        for item in lost_items:
            ws3.append([item['Причина'], item['Штрихкод / GTIN'], item['Детали']])
            
        header_fill = openpyxl.styles.PatternFill(start_color="3B5998", end_color="3B5998", fill_type="solid")
        header_font = openpyxl.styles.Font(bold=True, color="FFFFFF")
        thin_border = openpyxl.styles.Border(
            left=openpyxl.styles.Side(style='thin', color="DDDDDD"), 
            right=openpyxl.styles.Side(style='thin', color="DDDDDD"), 
            top=openpyxl.styles.Side(style='thin', color="DDDDDD"), 
            bottom=openpyxl.styles.Side(style='thin', color="DDDDDD")
        )
        
        for ws in [ws1, ws2, ws3]:
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=1, column=col)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                for cell in row:
                    cell.border = thin_border
                    cell.alignment = openpyxl.styles.Alignment(vertical="center", horizontal="left" if cell.column >= 3 else "center")

        ws1.column_dimensions['A'].width = 22
        ws1.column_dimensions['B'].width = 22
        ws1.column_dimensions['C'].width = 18
        ws1.column_dimensions['D'].width = 38
        ws1.column_dimensions['E'].width = 16
        ws1.column_dimensions['F'].width = 16
        ws1.column_dimensions['G'].width = 16
        ws1.column_dimensions['H'].width = 45
        ws1.column_dimensions['I'].width = 12
        
        ws2.column_dimensions['A'].width = 45
        ws2.column_dimensions['B'].width = 15
        
        ws3.column_dimensions['A'].width = 35
        ws3.column_dimensions['B'].width = 25
        ws3.column_dimensions['C'].width = 50

        wb_1c.save(output_1c)
        log_callback(f"\n[1С ИТОГ] Отчет сохранен (Потеряшки: {len(lost_items)} шт):\n{output_1c}")
    elif check_excel_path:
        log_callback("\n[1С ИТОГ] Справочник выбран, но ничего не найдено.")


# --- ФУНКЦИИ ВТОРОЙ ВКЛАДКИ (ПАРСЕР OZON PDF) ---
def parse_orders_from_pdf(pdf_path: str) -> pd.DataFrame:
    reader = PdfReader(pdf_path)
    full_text = ""
    for page in reader.pages:
        text = page.extract_text()
        if text:
            full_text += text + "\n"
    lines = full_text.split('\n')
    skus = []
    orders = []
    products = []
    
    regex_sku = r'((?:[A-Za-z]+[-_]?)?\d+(?:[-_][A-Za-z]+)?(?:[\s-]*\([^)]+\))?[\s_]+\d{2,3}(?:/\d{2,3})?(?:\s*\([^)]+\))?[-_]?|\b\d{5,10}\b)'
    
    for i, line in enumerate(lines):
        line_clean = line
        line_clean = re.sub(r'\b([A-Za-z0-9\-]*(\d))\s+\2_(\d+)\s+\3\b', r'\1_\3', line_clean)
        
        m_order = re.search(r'\d{6,15}-\d{4}-\d{1,3}', line_clean)
        if m_order:
            orders.append((i, m_order.group(0)))
            line_clean = line_clean.replace(m_order.group(0), '')
            
        m_sku = re.search(regex_sku, line_clean)
        if m_sku:
            sku_val = m_sku.group(1).strip()
            sku_val = re.sub(r'(\d)\s+(\d)', r'\1_\2', sku_val)
            skus.append((i, sku_val))
            
        prod = line_clean.replace('|', '').strip()
        if m_sku:
            prod = prod.replace(m_sku.group(0), '').strip()
        prod = re.sub(r'\s+\d+\s+\d{4}$', '', prod).strip()
        prod = re.sub(r'^[-_\s\d]+', '', prod).strip()
        if prod and len(prod) > 2 and not prod.replace('.', '', 1).isdigit():
            if not re.match(r'^(шт|руб|\d+\s*шт|\d+\s*руб)$', prod, re.IGNORECASE):
                products.append((i, prod))
                
    records = []
    for sku_idx, sku_val in skus:
        closest_order = min(orders, key=lambda x: abs(x[0] - sku_idx)) if orders else (0, '')
        closest_product = min(products, key=lambda x: abs(x[0] - sku_idx)) if products else (0, 'Товар не указан')
        order_val = closest_order[1]
        label = order_val.split('-')[0][-4:] if '-' in order_val else ''
        records.append({
            'Номер отправления': order_val,
            'Товар': closest_product[1],
            'Артикул': sku_val,
            'Кол-во': 1,
            'Этикетка': label
        })
    df = pd.DataFrame(records)
    if not df.empty:
        df.index = df.index + 1
        df.index.name = '№'
    return df

def save_to_formatted_excel(df: pd.DataFrame, output_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Заказы"
    headers = ['№', 'Номер отправления', 'Товар', 'Артикул', 'Кол-во', 'Этикетка']
    ws.append(headers)
    for row in df.itertuples():
        ws.append([row.Index, row[1], row[2], row[3], row[4], row[5]])
    header_font = openpyxl.styles.Font(bold=True, color="FFFFFF")
    header_fill = openpyxl.styles.PatternFill(start_color="3B5998", end_color="3B5998", fill_type="solid")
    thin_border = openpyxl.styles.Border(
        left=openpyxl.styles.Side(style='thin', color="DDDDDD"), 
        right=openpyxl.styles.Side(style='thin', color="DDDDDD"), 
        top=openpyxl.styles.Side(style='thin', color="DDDDDD"), 
        bottom=openpyxl.styles.Side(style='thin', color="DDDDDD")
    )
    for col in range(1, 7):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 45
    ws.column_dimensions['D'].width = 19
    ws.column_dimensions['E'].width = 10
    ws.column_dimensions['F'].width = 12
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=6):
        for cell in row:
            cell.border = thin_border
            cell.alignment = openpyxl.styles.Alignment(vertical="center", horizontal="left" if cell.column == 3 else "center")
            if cell.row % 2 == 0:
                cell.fill = openpyxl.styles.PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
    ws_summary = wb.create_sheet(title="Сводка")
    summary_df = df.groupby('Артикул')['Кол-во'].sum().reset_index()
    summary_df = summary_df.sort_values('Артикул')
    ws_summary.append(['Названия строк', 'Сумма по полю Кол-во'])
    total_qty = 0
    for _, row in summary_df.iterrows():
        ws_summary.append([row['Артикул'], row['Кол-во']])
        total_qty += row['Кол-во']
    ws_summary.append(['Общий итог', total_qty])
    for col in range(1, 3):
        cell = ws_summary.cell(row=1, column=col)
        cell.font = openpyxl.styles.Font(bold=True)
        cell.fill = openpyxl.styles.PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        cell.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
    last_row = ws_summary.max_row
    for col in range(1, 3):
        cell = ws_summary.cell(row=last_row, column=col)
        cell.font = openpyxl.styles.Font(bold=True)
        cell.fill = openpyxl.styles.PatternFill(start_color="A6A6A6", end_color="A6A6A6", fill_type="solid")
    ws_summary.column_dimensions['A'].width = 25
    ws_summary.column_dimensions['B'].width = 25
    for row in ws_summary.iter_rows(min_row=1, max_row=ws_summary.max_row, min_col=1, max_col=2):
        for cell in row:
            cell.border = thin_border
            cell.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
    wb.save(output_path)


# --- ГЛАВНОЕ ОКНО ПРИЛОЖЕНИЯ ---
class MainApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Инструменты логистики")
        self.geometry("850x920")

        self.tabview = ctk.CTkTabview(self, width=810, height=880)
        self.tabview.pack(padx=20, pady=10, fill="both", expand=True)

        self.tab_qr = self.tabview.add("Сверка и 1С")
        self.tab_ozon = self.tabview.add("Парсер Ozon PDF")
        self.tab_split = self.tabview.add("Разделение папки PDF")

        self.init_qr_tab()
        self.init_ozon_tab()
        self.init_split_tab()

    def init_qr_tab(self):
        header = ctk.CTkLabel(self.tab_qr, text="Обработка сканов, PDF и генерация 1С", font=ctk.CTkFont(size=22, weight="bold"), text_color="#90CAF9")
        header.pack(pady=(10, 10))

        files_frame = ctk.CTkFrame(self.tab_qr, fg_color="transparent")
        files_frame.pack(fill="x", padx=10, pady=5)

        self.lbl_scan = ctk.CTkLabel(files_frame, text="1. Прямое сканирование (поставьте сюда курсор и пикайте сканером):", font=ctk.CTkFont(weight="bold"))
        self.lbl_scan.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 2))
        self.textbox_direct_scans = ctk.CTkTextbox(files_frame, width=730, height=120, fg_color="#2b2b2b", border_width=2, border_color="#555555")
        self.textbox_direct_scans.grid(row=1, column=0, columnspan=2, pady=(0, 15))

        self.btn_scans = ctk.CTkButton(files_frame, text="ИЛИ выбрать Excel-скан", width=180, command=self.pick_scans_file)
        self.btn_scans.grid(row=2, column=0, padx=(0, 10), pady=8)
        self.entry_scans = ctk.CTkEntry(files_frame, placeholder_text="Файл не выбран (оставьте пустым, если сканируете выше)", width=540)
        self.entry_scans.grid(row=2, column=1, pady=8)

        self.btn_pdf_in = ctk.CTkButton(files_frame, text="Исходные PDF", width=180, command=self.pick_pdf_in)
        self.btn_pdf_in.grid(row=3, column=0, padx=(0, 10), pady=8)
        self.entry_pdf_in = ctk.CTkEntry(files_frame, placeholder_text="Папка с исходными PDF не выбрана", width=540)
        self.entry_pdf_in.grid(row=3, column=1, pady=8)

        self.btn_check = ctk.CTkButton(files_frame, text="Файл для сверки", width=180, command=self.pick_check_file)
        self.btn_check.grid(row=4, column=0, padx=(0, 10), pady=8)
        self.entry_check = ctk.CTkEntry(files_frame, placeholder_text="Файл справочника базы (необязательно)", width=540)
        self.entry_check.grid(row=4, column=1, pady=8)

        self.btn_pdf_out = ctk.CTkButton(files_frame, text="Куда сохранить?", width=180, command=self.pick_pdf_out)
        self.btn_pdf_out.grid(row=5, column=0, padx=(0, 10), pady=8)
        self.entry_pdf_out = ctk.CTkEntry(files_frame, placeholder_text="Папка для сохранения результатов не выбрана", width=540)
        self.entry_pdf_out.grid(row=5, column=1, pady=8)

        self.btn_start = ctk.CTkButton(self.tab_qr, text="Запустить обработку", height=42, width=280, font=ctk.CTkFont(size=15), command=self.start_qr_processing)
        self.btn_start.pack(pady=12)

        self.lbl_stage = ctk.CTkLabel(self.tab_qr, text="Стадия: Готово к работе", text_color="#ffcc00", font=ctk.CTkFont(weight="bold"))
        self.lbl_stage.pack(pady=3)

        self.progress_bar = ctk.CTkProgressBar(self.tab_qr, width=400)
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=(3, 0))
        self.progress_bar.pack_forget()

        self.lbl_progress = ctk.CTkLabel(self.tab_qr, text="")
        self.lbl_progress.pack(pady=(0, 3))
        self.lbl_progress.pack_forget()

        self.textbox_qr = ctk.CTkTextbox(self.tab_qr, width=750, height=180, fg_color="#1e1e1e")
        self.textbox_qr.pack(padx=10, pady=5, fill="both", expand=True)

    def pick_scans_file(self):
        filepath = filedialog.askopenfilename(title="Выберите файл сканов", filetypes=[("Excel", "*.xlsx *.xls")])
        if filepath:
            self.entry_scans.delete(0, 'end')
            self.entry_scans.insert(0, filepath)

    def pick_pdf_in(self):
        dirpath = filedialog.askdirectory(title="Выберите папку PDF")
        if dirpath:
            self.entry_pdf_in.delete(0, 'end')
            self.entry_pdf_in.insert(0, dirpath)

    def pick_check_file(self):
        filepath = filedialog.askopenfilename(title="Выберите базу (справочник)", filetypes=[("Excel", "*.xlsx *.xls")])
        if filepath:
            self.entry_check.delete(0, 'end')
            self.entry_check.insert(0, filepath)

    def pick_pdf_out(self):
        dirpath = filedialog.askdirectory(title="Папка для сохранения")
        if dirpath:
            self.entry_pdf_out.delete(0, 'end')
            self.entry_pdf_out.insert(0, dirpath)

    def safe_log_qr(self, msg):
        self.after(0, lambda: self._append_log_qr(msg))
        
    def _append_log_qr(self, msg):
        self.textbox_qr.insert("end", msg + "\n")
        self.textbox_qr.see("end")

    def safe_stage_qr(self, msg):
        self.after(0, lambda: self.lbl_stage.configure(text=f"Стадия: {msg}"))

    def safe_progress_qr(self, current, total, filename):
        def update():
            self.progress_bar.set(current / total if total > 0 else 0)
            self.lbl_progress.configure(text=f"{current} / {total} файлов")
            self.lbl_stage.configure(text=f"Проверяю: {filename}")
        self.after(0, update)

    def start_qr_processing(self):
        scans_path = self.entry_scans.get()
        direct_scans_text = self.textbox_direct_scans.get("1.0", "end-1c")
        
        pdf_in = self.entry_pdf_in.get()
        check_path = self.entry_check.get()
        pdf_out = self.entry_pdf_out.get()

        if (not scans_path and not direct_scans_text.strip()) or not pdf_in or not pdf_out:
            self.safe_stage_qr("Ошибка: Заполните обязательные поля или отсканируйте штрихкоды!")
            return

        self.btn_start.configure(state="disabled")
        self.progress_bar.pack(pady=(3, 0))
        self.lbl_progress.pack(pady=(0, 3))
        self.progress_bar.set(0)
        self.textbox_qr.delete("1.0", "end")

        threading.Thread(target=self.run_full_script, args=(scans_path, direct_scans_text, pdf_in, pdf_out, check_path), daemon=True).start()

    def run_full_script(self, scans_path, direct_scans_text, current_pdf_dir, current_output_dir, check_path):
        self.safe_stage_qr("Выполнение процесса...")
        self.safe_log_qr("Запуск обработки...")
        try:
            process_full_workflow(
                scans_path, direct_scans_text, current_pdf_dir, current_output_dir, check_path, self.safe_log_qr, self.safe_progress_qr
            )
            self.safe_stage_qr("Готово!")
        except Exception as e:
            self.safe_log_qr(f"Ошибка: {e}")
            self.safe_stage_qr("Произошла ошибка")
        finally:
            self.after(0, lambda: self.btn_start.configure(state="normal"))
            self.after(0, self.progress_bar.pack_forget)
            self.after(0, self.lbl_progress.pack_forget)


    def init_ozon_tab(self):
        header = ctk.CTkLabel(self.tab_ozon, text="Парсер накладных Ozon в Excel", font=ctk.CTkFont(size=22, weight="bold"), text_color="#90CAF9")
        header.pack(pady=(10, 15))
        ozon_frame = ctk.CTkFrame(self.tab_ozon, fg_color="transparent")
        ozon_frame.pack(fill="x", padx=10, pady=5)
        self.btn_ozon_pdf = ctk.CTkButton(ozon_frame, text="Выбрать PDF", width=150, command=self.pick_ozon_pdf)
        self.btn_ozon_pdf.grid(row=0, column=0, padx=(0, 10), pady=10)
        self.entry_ozon_pdf = ctk.CTkEntry(ozon_frame, placeholder_text="PDF файл с накладной Ozon не выбран", width=570)
        self.entry_ozon_pdf.grid(row=0, column=1, pady=10)
        self.btn_ozon_out = ctk.CTkButton(ozon_frame, text="Куда сохранить?", width=150, command=self.pick_ozon_out)
        self.btn_ozon_out.grid(row=1, column=0, padx=(0, 10), pady=10)
        self.entry_ozon_out = ctk.CTkEntry(ozon_frame, placeholder_text="Папка для сохранения результата не выбрана", width=570)
        self.entry_ozon_out.grid(row=1, column=1, pady=10)
        self.btn_start_ozon = ctk.CTkButton(self.tab_ozon, text="Запустить парсинг PDF", height=42, width=250, font=ctk.CTkFont(size=15), command=self.start_ozon_processing)
        self.btn_start_ozon.pack(pady=15)
        self.lbl_stage_ozon = ctk.CTkLabel(self.tab_ozon, text="Стадия: Готово к работе", text_color="#ffcc00", font=ctk.CTkFont(weight="bold"))
        self.lbl_stage_ozon.pack(pady=5)
        self.textbox_ozon = ctk.CTkTextbox(self.tab_ozon, width=750, height=310, fg_color="#1e1e1e")
        self.textbox_ozon.pack(padx=10, pady=10, fill="both", expand=True)

    def pick_ozon_pdf(self):
        filepath = filedialog.askopenfilename(title="Выберите PDF файл Ozon", filetypes=[("PDF", "*.pdf")])
        if filepath:
            self.entry_ozon_pdf.delete(0, 'end')
            self.entry_ozon_pdf.insert(0, filepath)

    def pick_ozon_out(self):
        dirpath = filedialog.askdirectory(title="Выберите папку для сохранения отчета")
        if dirpath:
            self.entry_ozon_out.delete(0, 'end')
            self.entry_ozon_out.insert(0, dirpath)

    def safe_log_ozon(self, msg):
        self.after(0, lambda: self._append_log_ozon(msg))

    def _append_log_ozon(self, msg):
        self.textbox_ozon.insert("end", msg + "\n")
        self.textbox_ozon.see("end")

    def safe_stage_ozon(self, msg):
        self.after(0, lambda: self.lbl_stage_ozon.configure(text=f"Стадия: {msg}"))

    def start_ozon_processing(self):
        pdf_path = self.entry_ozon_pdf.get()
        out_dir = self.entry_ozon_out.get()
        if not pdf_path or not out_dir:
            self.safe_stage_ozon("Ошибка: Выберите PDF-файл и папку сохранения!")
            return
        self.btn_start_ozon.configure(state="disabled")
        self.textbox_ozon.delete("1.0", "end")
        threading.Thread(target=self.run_ozon_script, args=(pdf_path, out_dir), daemon=True).start()

    def run_ozon_script(self, pdf_path, out_dir):
        self.safe_stage_ozon("Парсинг PDF...")
        self.safe_log_ozon(f"Читаем файл: {os.path.basename(pdf_path)}")
        try:
            df_orders = parse_orders_from_pdf(pdf_path)
            if df_orders.empty:
                self.safe_log_ozon("Не удалось найти заказы по заданным шаблонам.")
                self.safe_stage_ozon("Завершено с предупреждением")
            else:
                output_xlsx = os.path.join(out_dir, "Ozon_Orders_List.xlsx")
                save_to_formatted_excel(df_orders, output_xlsx)
                self.safe_log_ozon(f"Успешно обработано позиций: {len(df_orders)}")
                self.safe_log_ozon(f"Файл сохранен: {output_xlsx}")
                self.safe_stage_ozon("Успешно завершено!")
        except Exception as e:
            self.safe_log_ozon(f"Ошибка при обработке PDF: {e}")
            self.safe_stage_ozon("Произошла ошибка")
        finally:
            self.after(0, lambda: self.btn_start_ozon.configure(state="normal"))

    def init_split_tab(self):
        header = ctk.CTkLabel(self.tab_split, text="Массовое разделение папки PDF", font=ctk.CTkFont(size=22, weight="bold"), text_color="#90CAF9")
        header.pack(pady=(10, 15))
        split_frame = ctk.CTkFrame(self.tab_split, fg_color="transparent")
        split_frame.pack(fill="x", padx=10, pady=5)
        self.btn_split_in = ctk.CTkButton(split_frame, text="Папка с PDF", width=150, command=self.pick_split_in)
        self.btn_split_in.grid(row=0, column=0, padx=(0, 10), pady=10)
        self.entry_split_in = ctk.CTkEntry(split_frame, placeholder_text="Папка с исходными PDF не выбрана", width=570)
        self.entry_split_in.grid(row=0, column=1, pady=10)
        self.btn_split_out = ctk.CTkButton(split_frame, text="Куда сохранить?", width=150, command=self.pick_split_out)
        self.btn_split_out.grid(row=1, column=0, padx=(0, 10), pady=10)
        self.entry_split_out = ctk.CTkEntry(split_frame, placeholder_text="Папка для сохранения не выбрана", width=570)
        self.entry_split_out.grid(row=1, column=1, pady=10)
        self.btn_start_split = ctk.CTkButton(self.tab_split, text="Запустить массовое разделение", height=42, width=280, font=ctk.CTkFont(size=15), command=self.start_split_processing)
        self.btn_start_split.pack(pady=15)
        self.lbl_stage_split = ctk.CTkLabel(self.tab_split, text="Стадия: Готово к работе", text_color="#ffcc00", font=ctk.CTkFont(weight="bold"))
        self.lbl_stage_split.pack(pady=5)
        self.textbox_split = ctk.CTkTextbox(self.tab_split, width=750, height=310, fg_color="#1e1e1e")
        self.textbox_split.pack(padx=10, pady=10, fill="both", expand=True)

    def pick_split_in(self):
        dirpath = filedialog.askdirectory(title="Выберите папку с PDF")
        if dirpath:
            self.entry_split_in.delete(0, 'end')
            self.entry_split_in.insert(0, dirpath)

    def pick_split_out(self):
        dirpath = filedialog.askdirectory(title="Папка для сохранения")
        if dirpath:
            self.entry_split_out.delete(0, 'end')
            self.entry_split_out.insert(0, dirpath)

    def safe_log_split(self, msg):
        self.after(0, lambda: self._append_log_split(msg))

    def _append_log_split(self, msg):
        self.textbox_split.insert("end", msg + "\n")
        self.textbox_split.see("end")

    def safe_stage_split(self, msg):
        self.after(0, lambda: self.lbl_stage_split.configure(text=f"Стадия: {msg}"))

    def start_split_processing(self):
        pdf_dir = self.entry_split_in.get()
        out_dir = self.entry_split_out.get()
        if not pdf_dir or not out_dir:
            self.safe_stage_split("Ошибка: Выберите папку с PDF и папку сохранения!")
            return
        self.btn_start_split.configure(state="disabled")
        self.textbox_split.delete("1.0", "end")
        threading.Thread(target=self.run_folder_split_script, args=(pdf_dir, out_dir), daemon=True).start()

    def run_folder_split_script(self, pdf_dir, out_dir):
        self.safe_stage_split("Обработка папок...")
        self.safe_log_split(f"Сканируем папку и все её подпапки:\n{pdf_dir}")
        try:
            search_pattern = os.path.join(pdf_dir, '**', '*.pdf')
            pdf_files = glob.glob(search_pattern, recursive=True)
            
            if not pdf_files:
                self.safe_log_split("В выбранной директории и её подпапках нет PDF-файлов!")
                self.safe_stage_split("Завершено")
                return

            master_male = fitz.open()
            master_female = fitz.open()
            master_other = fitz.open()
            total_male, total_female, total_other = 0, 0, 0

            for pdf_path in pdf_files:
                filename = os.path.basename(pdf_path)
                current_dir = os.path.dirname(pdf_path) 
                base_name = os.path.splitext(filename)[0]
                
                if base_name.endswith('_Мужские') or base_name.endswith('_Женские') or base_name.endswith('_Неопределенные') or base_name.startswith('All_Folder_'):
                    continue

                self.safe_log_split(f"\nЧитаем: {filename}\n(В папке: {current_dir})")
                
                doc = fitz.open(pdf_path)
                file_male = fitz.open()
                file_female = fitz.open()
                file_other = fitz.open()
                f_male, f_female, f_other = 0, 0, 0
                
                for page_num in range(len(doc)):
                    text = doc[page_num].get_text().lower()
                    if 'женск' in text or 'жен.' in text or 'w-' in text:
                        file_female.insert_pdf(doc, from_page=page_num, to_page=page_num)
                        master_female.insert_pdf(doc, from_page=page_num, to_page=page_num)
                        f_female += 1
                        total_female += 1
                    elif 'мужск' in text or 'муж.' in text:
                        file_male.insert_pdf(doc, from_page=page_num, to_page=page_num)
                        master_male.insert_pdf(doc, from_page=page_num, to_page=page_num)
                        f_male += 1
                        total_male += 1
                    else:
                        file_other.insert_pdf(doc, from_page=page_num, to_page=page_num)
                        master_other.insert_pdf(doc, from_page=page_num, to_page=page_num)
                        f_other += 1
                        total_other += 1
                        
                if f_male > 0: file_male.save(os.path.join(current_dir, f"{base_name}_Мужские.pdf"))
                if f_female > 0: file_female.save(os.path.join(current_dir, f"{base_name}_Женские.pdf"))
                if f_other > 0: file_other.save(os.path.join(current_dir, f"{base_name}_Неопределенные.pdf"))
                
                self.safe_log_split(f"  -> Муж: {f_male}, Жен: {f_female}, Др: {f_other}")
                
                file_male.close()
                file_female.close()
                file_other.close()
                doc.close()
                
            if total_male > 0:
                master_male.save(os.path.join(out_dir, "All_Folder_Мужские.pdf"))
                self.safe_log_split(f"\n[ОБЩИЙ ИТОГ] All_Folder_Мужские.pdf ({total_male} стр.) сохранен в папку вывода.")
            if total_female > 0:
                master_female.save(os.path.join(out_dir, "All_Folder_Женские.pdf"))
                self.safe_log_split(f"[ОБЩИЙ ИТОГ] All_Folder_Женские.pdf ({total_female} стр.)")
            if total_other > 0:
                master_other.save(os.path.join(out_dir, "All_Folder_Неопределенные.pdf"))
                self.safe_log_split(f"[ОБЩИЙ ИТОГ] All_Folder_Неопределенные.pdf ({total_other} стр.)")
                
            master_male.close()
            master_female.close()
            master_other.close()
            
            self.safe_log_split("\nМассовое разделение папок успешно завершено!")
            self.safe_stage_split("Успешно завершено!")
        except Exception as e:
            self.safe_log_split(f"Ошибка: {e}")
            self.safe_stage_split("Произошла ошибка")
        finally:
            self.after(0, lambda: self.btn_start_split.configure(state="normal"))

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
