import openpyxl

def convert_modbus_map(path_map="",name_new_map="modbus_for_panel", name_sheet = "Шаблон 2"):
    try:
        workbook = openpyxl.load_workbook(path_map)
    except Exception as e:
        print(f"Ошибка при открытии файла: {e}")
        return False

    if name_sheet not in workbook.sheetnames:
        print(f"Лист '{name_sheet}' не найден в таблице")
        return False

    ws_data = workbook[name_sheet]

    # Именованные столбцы для ясности
    ADDR_LABEL = 'K'  # Адресная метка
    DEVICE_NAME = 'J'  # Имя устройства
    BITMAP_TYPE = 'L'  # Тип: BITMAP или нет
    DATA_TYPE = 'D'  # Тип данных (BOOL и др.)
    REG_ADDR = 'B'  # Адрес регистра
    SIGNAL_NAME = 'A'  # Имя сигнала

    new_workbook = openpyxl.Workbook()
    new_sheet = new_workbook.active
    new_sheet.title = "Sheet 1"

    name_plc = "PLC"
    row_counter = 1

    def get_function_type(data_type):
        return "4x_Bit" if data_type == "BOOL" else "4x"

    def convert_data_type(data_type: str) -> str:
        print(data_type)
        mapping = {
            "FLOAT(4 byte)": "32-bit Float",
            "INT(2 byte)": "16-bit Unsigned",
            # Можно добавить другие типы при необходимости
        }
        return mapping.get(data_type.strip(), "Неопределенный")

    max_row = ws_data.max_row
    for row in range(2, max_row + 1):
        addr_label_val = ws_data[f"{ADDR_LABEL}{row}"].value
        if not addr_label_val:  # Пропускаем пустые строки
            continue

        bitmap_type_val = ws_data[f"{BITMAP_TYPE}{row}"].value
        device_name_val = ws_data[f"{DEVICE_NAME}{row}"].value or ""
        reg_addr_val = ws_data[f"{REG_ADDR}{row}"].value
        data_type_val = ws_data[f"{DATA_TYPE}{row}"].value
        panel_data_type = convert_data_type(data_type_val)
        signal_name_val = ws_data[f"{SIGNAL_NAME}{row}"].value

        # Определяем тип функции
        func_type = get_function_type(data_type_val)

        if bitmap_type_val == "BITMAP":
            # Обработка BITMAP: разбиваем по строкам
            try:
                cell_bitmap = {}
                for line in str(addr_label_val).strip().split('\n'):
                    if '-' in line:
                        bit, name_part = line.split('-', 1)
                        cell_bitmap[int(bit)] = name_part.strip()
                # Сортируем по биту
                for bit, bit_name in sorted(cell_bitmap.items()):
                    signal_full_name = f"{bit_name}_{device_name_val}" if device_name_val else bit_name

                    new_sheet.cell(row=row_counter, column=1, value=signal_full_name)  # A: имя сигнала
                    new_sheet.cell(row=row_counter, column=2, value=name_plc)  # B: имя устройства
                    new_sheet.cell(row=row_counter, column=3, value="4x_Bit")  # C: тип (всегда 4x_Bit для BITMAP)
                    new_sheet.cell(row=row_counter, column=4, value=f"{reg_addr_val}.{bit}")  # D: адрес
                    new_sheet.cell(row=row_counter, column=5, value=signal_name_val)
                    new_sheet.cell(row=row_counter, column=6, value=panel_data_type)

                    row_counter += 1
            except Exception as e:
                print(f"Ошибка обработки BITMAP в строке {row}: {e}")
                continue
        else:
            # Обычная строка
            signal_full_name = f"{addr_label_val}_{device_name_val}" if device_name_val else addr_label_val

            new_sheet.cell(row=row_counter, column=1, value=signal_full_name)
            new_sheet.cell(row=row_counter, column=2, value=name_plc)
            new_sheet.cell(row=row_counter, column=3, value=func_type)
            new_sheet.cell(row=row_counter, column=4, value=reg_addr_val)
            new_sheet.cell(row=row_counter, column=5, value=signal_name_val)
            new_sheet.cell(row=row_counter, column=6, value=panel_data_type)

            row_counter += 1

    # Сохранение
    try:
        new_workbook.save(name_new_map + ".xlsx")  # или просто name_new_map, если уже с .xlsx
        print(f"Файл успешно сохранён как {name_new_map}.xlsx")
        return True
    except Exception as e:
        print(f"Ошибка при сохранении файла: {e}")
        return False



if __name__ == "__main__" :
    convert_modbus_map("modbus_map.xlsx", name_sheet="Шаблон")