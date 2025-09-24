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
        # Очищаем предыдущие значения
        df.at[index, 'name'] = None
        df.at[index, 'mpi'] = None
        df.at[index, 'manufacturer'] = None

        # Поиск и ввод регистрационного номера
        input_field = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, 'input.form-control'))
)
        input_field.clear()
        input_field.send_keys(data)  # data содержит текущий регистрационный номер (например "3345-09")
        input_field.send_keys(Keys.RETURN)

        # Клик по кнопке поиска
        search_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '.fa.fa-search'))
        )
        search_button.click()

#########
        # Находит конкретную строку с нужным номером
        try:
            # Используем более точный XPath для поиска
            target_row = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, f'//tr[.//td[1][normalize-space()="{data}"]]'))
            )
            
            # Прокручиваем к найденной строке
            driver.execute_script("arguments[0].scrollIntoView();", target_row)
            time.sleep(0.5)
            
            # Двойной клик по строке
            webdriver.ActionChains(driver).double_click(target_row).perform()
            
        except Exception as e:
            print(f"Не удалось найти строку с номером {data}: {e}")
            # Для отладки выведем все номера из таблицы
            all_rows = driver.find_elements(By.CSS_SELECTOR, 'table.table tr')
            print("Найденные номера в таблице:")
            for row in all_rows[1:]:  # Пропускаем заголовок
                cells = row.find_elements(By.TAG_NAME, 'td')
                if cells:
                    print(f"- {cells[0].text.strip()}")

#########
        # Извлечение наименования
        try:
            name_table = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'table.table.table-striped.table-2columns'))
            )
            rows = name_table.find_elements(By.CSS_SELECTOR, 'tbody tr')
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, 'td')
                if len(cells) >= 2 and cells[0].text.strip() == "Наименование":
                    df.at[index, 'name'] = cells[1].text.strip()
                    break
        except Exception as e:
            print(f"Ошибка при извлечении наименования: {e}")

        # Извлечение МПИ (с возможностью пропуска если вкладка не найдена)
        try:
            # Пытаемся найти вкладку
            mpi_tab = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, 
                    '//div[contains(@class, "tabhead") and contains(., "Межповерочный интервал")]'))
            )
            mpi_tab.click()
            
            # Пытаемся загрузить таблицу
            mpi_table = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'table.table-striped.subtabTable'))
            )

            # Список для хранения всех найденных значений МПИ
            mpi_values = []
            
            # Анализ строк таблицы
            mpi_rows = mpi_table.find_elements(By.CSS_SELECTOR, 'tr.borderBetweenChildren')
            
            for row in mpi_rows:
                try:
                    cells = row.find_elements(By.TAG_NAME, 'td')
                    if len(cells) >= 2 and any(word in cells[1].text.lower() for word in ['год', 'лет', 'день']):
                        mpi_value = cells[1].text.strip()
                        mpi_values.append(mpi_value)
                        print(f"Найден МПИ: {mpi_value}")
                except Exception as e:
                    continue
            
            # Сохраняем все значения через запятую или выбираем нужное
            if mpi_values:
                df.at[index, 'mpi'] = ', '.join(mpi_values)  # или выбрать первое/последнее/максимальное значение
                print(f"Все найденные МПИ: {df.at[index, 'mpi']}")

        except Exception as e:
            print(f"Вкладка МПИ не найдена или другие проблемы с МПИ: {e}")

#########
        # Извлечение информации о производителе
        try:
            manufacturer_tab = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, 
                    '//div[contains(@class, "tabhead") and contains(., "Изготовители")]'))
            )
            manufacturer_tab.click()

            manufacturer_table = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'table.table-striped.subtabTable'))
            )

            manufacturers = []
            manufacturer_rows = manufacturer_table.find_elements(By.CSS_SELECTOR, 'tr.borderBetweenChildren')
            
            for row in manufacturer_rows:
                try:
                    # Пропускаем строку с заголовками
                    if "Наименование организации" in row.text and "ИНН" in row.text:
                        continue
                        
                    cells = row.find_elements(By.TAG_NAME, 'td')
                    if len(cells) >= 3:  # Проверяем наличие нужных столбцов
                        name = cells[0].text.strip()
                        address = cells[2].text.strip()
                        
                        if name and address:
                            manufacturer_info = f"{name}, {address}"
                            manufacturers.append(manufacturer_info)
                        elif name:
                            manufacturers.append(name)
                            
                except Exception as e:
                    print(f"Ошибка обработки строки производителя: {e}")
                    continue
            
            if manufacturers:
                df.at[index, 'manufacturer'] = " | ".join(manufacturers)
                print(f"Найдены производители: {df.at[index, 'manufacturer']}")
            else:
                print("Не удалось извлечь данные о производителях")
                df.at[index, 'manufacturer'] = None

        except Exception as e:
            print(f"Ошибка при извлечении информации о производителях: {e}")
            df.at[index, 'manufacturer'] = None

        # Сохранение в Excel
        try:
            df.to_excel(excel_file_path, index=False, engine='openpyxl')
            print(f"Данные для {data} успешно сохранены.")
            return True
        except Exception as e:
            print(f"Ошибка при сохранении в Excel: {e}")
            return False

    except Exception as e:
        print(f"Общая ошибка при обработке записи {index + 1} ({data}): {e}")
        return False

# Основной код
if __name__ == "__main__":
    chrome_options = Options()
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    excel_file_path = select_file()
    if not excel_file_path:
        print("Файл не выбран. Программа завершена.")
        driver.quit()
        exit()

    df = pd.read_excel(excel_file_path)

    # Инициализация столбцов
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

    max_retries = 3

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
    print("\nОбработка всех записей завершена.")