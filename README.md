# Excel Data Processor

## Description

This Python program processes data from one Excel file (`sheet.xlsx`) and writes the extracted data to another Excel file (`test.xlsx`). It utilizes the `openpyxl` library to interact with Excel files.

---

# Обработчик данных Excel

## Описание

Эта программа на Python обрабатывает данные из одного файла Excel (`sheet.xlsx`) и записывает извлеченные данные в другой файл Excel (`test.xlsx`). Для работы с файлами Excel используется библиотека `openpyxl`.

---

## Requirements / Требования

- **English:**
  - Python 3.x
  - `openpyxl` library (install it via `pip install openpyxl`)

- **Русский:**
  - Python 3.x
  - Библиотека `openpyxl` (установить с помощью `pip install openpyxl`)

---

## How It Works / Как работает программа

1. **Input File (`sheet.xlsx`) / Входной файл (`sheet.xlsx`):**
   - **English:**  
     - The program reads data from columns "F" and "G" for rows 3 to 41.
     - If both cells in a row contain numbers, they are converted to integers and stored as tuples `(F, G)`.
     - If either cell is empty, the program appends a placeholder tuple `(' ', ' ')` to the data list.
     - Rows with incompatible data types are skipped without causing errors.
   - **Русский:**  
     - Программа считывает данные из столбцов "F" и "G" для строк с 3 по 41.
     - Если обе ячейки в строке содержат числа, они преобразуются в целые числа и сохраняются как кортеж `(F, G)`.
     - Если хотя бы одна ячейка пустая, в список данных добавляется кортеж-заполнитель `(' ', ' ')`.
     - Строки с несовместимыми типами данных пропускаются без ошибок.

2. **Output File (`test.xlsx`) / Выходной файл (`test.xlsx`):**
   - **English:**  
     - The extracted data is written to columns "E", "F", and "G" in `test.xlsx`.
     - Column "E" contains a sequential counter starting from 1.
     - Columns "F" and "G" contain the processed values from `sheet.xlsx`.
   - **Русский:**  
     - Извлеченные данные записываются в столбцы "E", "F" и "G" в файле `test.xlsx`.
     - Столбец "E" содержит последовательный счетчик, начинающийся с 1.
     - Столбцы "F" и "G" содержат обработанные значения из `sheet.xlsx`.

---

## File Descriptions / Описание файлов

- **English:**
  - `sheet.xlsx`: The source Excel file containing the input data.
  - `test.xlsx`: The destination Excel file where the processed data is written.

- **Русский:**
  - `sheet.xlsx`: Исходный файл Excel, содержащий входные данные.
  - `test.xlsx`: Файл Excel, в который записываются обработанные данные.

---

## How to Use / Как использовать

1. **English:**  
   - Place the source file (`sheet.xlsx`) and the destination file (`test.xlsx`) in the same directory as the script.  
   - Run the script using `python script_name.py`.  

2. **Русский:**  
   - Поместите исходный файл (`sheet.xlsx`) и файл для записи данных (`test.xlsx`) в ту же папку, где находится скрипт.  
   - Запустите скрипт с помощью команды `python script_name.py`.  

---

## Notes / Примечания

- **English:** Ensure that both Excel files exist in the working directory before running the script.  
- **Русский:** Убедитесь, что оба файла Excel находятся в рабочей директории перед запуском скрипта.  
