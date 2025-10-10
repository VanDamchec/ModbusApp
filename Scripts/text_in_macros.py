import openpyxl
import os
import argparse

# Allowed signal types
ALLOWED_SIGNALS = {"AI", "AO", "DI", "DO", "MTR", "MTRPID", "DVLV", "AVLV"}

def generate_case_file(input_excel, signal_name, output_file="output_code"):
    if signal_name not in ALLOWED_SIGNALS:
        print(f"Ошибка: '{signal_name}' не является допустимым сигналом.")
        print(f"Допустимые значения: {', '.join(sorted(ALLOWED_SIGNALS))}")
        return

    file_name = f"{output_file}_{signal_name}.txt"

    # Проверка файла
    if not os.path.exists(input_excel):
        print(f"Файл {input_excel} не найден.")
        return

    # Открываем Excel
    try:
        wb = openpyxl.load_workbook(input_excel, read_only=True)
    except Exception as e:
        print(f"Ошибка при открытии Excel-файла: {e}")
        return

    if signal_name not in wb.sheetnames:
        print(f"Лист '{signal_name}' не найден в файле {input_excel}.")
        print(f"Доступные листы: {', '.join(wb.sheetnames)}")
        return

    ws = wb[signal_name]
    signals = []
    for row in ws.iter_rows(min_col=1, max_col=1, min_row=2):  # skip header
        cell_value = row[0].value
        if cell_value is not None:
            signals.append(str(cell_value))

    if not signals:
        print(f"В листе '{signal_name}' не найдено данных (столбец A пуст).")
        return

    # Генерируем текст
    try:
        with open(file_name, "w", encoding="utf-8") as f:
            for i, text in enumerate(signals, 1):
                f.write(f"case {i}\n")
                f.write(f'    String2Unicode("{text}", unicode_str[0])\n')
                f.write(f"    break\n\n")
        print(f"Файл {file_name} успешно создан с {len(signals)} case-ами.")
    except Exception as e:
        print(f"Ошибка при записи файла: {e}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Генерация case-блоков для заданного типа сигнала из Excel-файла."
    )
    parser.add_argument(
        "input_excel",
        help="Путь к входному Excel-файлу (например, signals.xlsx)"
    )
    parser.add_argument(
        "signal_name",
        choices=sorted(ALLOWED_SIGNALS),
        help="Тип сигнала: " + ", ".join(sorted(ALLOWED_SIGNALS))
    )
    parser.add_argument(
        "--output_file",
        default="output_code",
        help="Базовое имя выходного файла (по умолчанию: output_code)"
    )

    args = parser.parse_args()

    generate_case_file(
        input_excel=args.input_excel,
        signal_name=args.signal_name,
        output_file=args.output_file
    )