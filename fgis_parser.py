import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time
import tkinter as tk
from tkinter import filedialog
from webdriver_manager.chrome import ChromeDriverManager

def select_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Выберите файл Excel",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    return file_path

def reload_initial_page(driver, max_attempts=3):
    attempts = 0
    while attempts < max_attempts:
        try:
            driver.get('https://fgis.gost.ru/fundmetrology/registry/4')
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'input.form-control'))
            )
            return True
        except Exception as e:
            attempts += 1
            print(f"Попытка {attempts} перезагрузки страницы не удалась: {e}")
            time.sleep(5)
    return False

def process_record(driver, data, index, df, excel_file_path):
    try:
        # Ожидание и заполнение поля ввода
        input_field = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input.form-control'))
        )
        input_field.clear()
        input_field.send_keys(data)
        input_field.send_keys(Keys.RETURN)

        # Ожидание появления элемента с текстом
        WebDriverWait(driver, 30).until(
            EC.text_to_be_present_in_element((By.XPATH, f"//span[text()='{data}']"), data)
        )

        # Ожидание таблицы
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'reactable-data'))
        )

        # Поиск и клик по кнопке
        btt = driver.find_elements(By.CSS_SELECTOR, 'tbody.reactable-data tr')
        for bt in btt:
            reg_cell = bt.find_element(By.CSS_SELECTOR, 'td:nth-child(1) span')
            if reg_cell.text.strip() == data:
                button = bt.find_element(By.CSS_SELECTOR, 'button.btn.btn-xs.btn-info')
                button.click()
                break

        # Ожидание детальной информации
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'reactable-data'))
        )

        # Извлечение данных
        rows = driver.find_elements(By.CSS_SELECTOR, 'tbody.reactable-data tr')
        if len(rows) >= 2:
            second_row_cells = rows[1].find_elements(By.CSS_SELECTOR, 'td[label="Значение"]')
            if second_row_cells:
                df.at[index, 'name'] = second_row_cells[0].text

        manufacturer_elements = driver.find_elements(By.CSS_SELECTOR, 'td[label="Название"]')
        for element in manufacturer_elements:
            if element.text == "Изготовитель":
                parent_row = element.find_element(By.XPATH, '..')
                manufacturer_span = parent_row.find_element(By.CSS_SELECTOR, 'td[label="Значение"] span')
                if manufacturer_span:
                    df.at[index, 'manufacturer'] = manufacturer_span.text

        mpi_elements = driver.find_elements(By.CSS_SELECTOR, 'td[label="Название"]')
        for mpi in mpi_elements:
            if mpi.text == "МПИ":
                parent_row = mpi.find_element(By.XPATH, '..')
                second_mpi_cells = parent_row.find_elements(By.CSS_SELECTOR, 'td[label="Значение"]')
                if second_mpi_cells:
                    df.at[index, 'mpi'] = second_mpi_cells[0].text

        df.to_excel(excel_file_path, index=False)
        print(f"Данные для записи {index + 1} успешно обработаны.")
        return True

    except Exception as e:
        print(f"Ошибка при обработке записи {index + 1} ({data}): {e}")
        return False

# Основной код
chrome_options = Options()
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

excel_file_path = select_file()
if not excel_file_path:
    print("Файл не выбран. Программа завершена.")
    driver.quit()
    exit()

df = pd.read_excel(excel_file_path)

# Инициализация столбцов, если их нет
for col in ['name', 'mpi', 'manufacturer']:
    if col not in df.columns:
        df[col] = None

if 'Рег. Номер' not in df.columns:
    print("Отсутствует столбец 'Рег. Номер'")
    driver.quit()
    exit()

data_list = df['Рег. Номер'].dropna().tolist()

if not reload_initial_page(driver):
    print("Не удалось загрузить начальную страницу")
    driver.quit()
    exit()

max_retries = 3  # Максимальное количество попыток обработки одной записи

for index, data in enumerate(data_list):
    if pd.notna(df.at[index, 'name']) and pd.notna(df.at[index, 'mpi']) and pd.notna(df.at[index, 'manufacturer']):
        print(f"Запись {index + 1} уже обработана, пропускаем.")
        continue

    print(f"\nНачало обработки записи {index + 1}: {data}")
    retry_count = 0
    success = False

    while not success and retry_count < max_retries:
        if retry_count > 0:
            print(f"Повторная попытка {retry_count} для записи {index + 1}")

        success = process_record(driver, data, index, df, excel_file_path)
        
        if success:
            # После успешной обработки возвращаемся на начальную страницу
            if not reload_initial_page(driver):
                print("Не удалось вернуться на начальную страницу")
                break
        else:
            retry_count += 1
            if not reload_initial_page(driver):
                print("Не удалось перезагрузить страницу. Прекращаем попытки.")
                break
            time.sleep(2)

    if not success:
        print(f"Не удалось обработать запись {index + 1} после {max_retries} попыток")

driver.quit()
print("Обработка всех записей завершена.")