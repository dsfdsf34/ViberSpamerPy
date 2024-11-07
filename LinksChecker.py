import time
import keyboard
import os
import threading
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

print("F8 Пауза/Возобновление скрипта")

# Путь к файлу настроек
settings_file = "settings.txt"
pause_flag = False


def toggle_pause():
    """Функция для переключения состояния паузы."""
    global pause_flag
    while True:
        if keyboard.is_pressed('F8'):  # Если нажата клавиша F8
            pause_flag = not pause_flag  # Меняем состояние паузы
            print("Пауза!" if pause_flag else "Продолжение...")
            time.sleep(1)  # Предотвращаем многократное срабатывание


def read_settings():
    """Чтение настроек из файла."""
    if os.path.exists(settings_file):
        with open(settings_file, "r") as file:
            path_to_excel = file.readline().strip()
            start_row = int(file.readline().strip())
            return path_to_excel, start_row
    else:
        return None, None


def save_settings(path_to_excel, start_row):
    """Сохранение настроек в файл."""
    with open(settings_file, "w") as file:
        file.write(f"{path_to_excel}\n{start_row}\n")


def check_viber_group_status(link, driver):
    # Открытие ссылки на новой вкладке
    try:
        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[-1])  # Переход на новую вкладку
        driver.get(link)
        time.sleep(2)  # Время ожидания загрузки страницы

        # Проверка наличия текста для недействительного приглашения
        invalid_texts = [
            "Ссылка приглашения неактивна",
            "Срок действия приглашения истек",
            "группа не существует",
            "Страница не найдена"
        ]
        page_source = driver.page_source
        if any(text in page_source for text in invalid_texts):
            return "Неактивна"
        else:
            return "Активна"
    except Exception as e:
        return f"Ошибка при проверке ссылки {link}: {e}"
    finally:
        # Закрытие текущей вкладки и возврат к первой
        try:
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
        except Exception as close_error:
            print(f"Ошибка при закрытии вкладки или переключении: {close_error}")


def main():
    # Запрос на использование файла настроек
    use_settings = input("Хотите использовать файл настроек? (y/n): ").lower()

    if use_settings == "y":
        path_to_excel, start_row = read_settings()
        if path_to_excel and start_row:
            print(f"Используем настройки: путь к файлу Excel - {path_to_excel}, начало с строки {start_row}")
        else:
            print("Настройки не найдены. Пожалуйста, укажите путь и начальную строку вручную.")
            path_to_excel = input("Введите путь к файлу Excel с ссылками: ")
            start_row = int(input(f"С какой строки начать проверку? (по умолчанию 2): ") or 2)
            save_settings(path_to_excel, start_row)
    else:
        # Если не используется файл настроек, запрашиваем путь и начальную строку вручную
        path_to_excel = input("Введите путь к файлу Excel с ссылками: ")
        start_row = int(input(f"С какой строки начать проверку? (по умолчанию 2): ") or 2)

    # Загрузка файла Excel
    try:
        workbook = load_workbook(path_to_excel)
        sheet = workbook.active
    except Exception as e:
        print(f"Ошибка при загрузке файла Excel: {e}")
        return

    # Запрос на удаление неактивных ссылок
    delete_links = input("Удалять неактивные ссылки из таблицы? (y/n): ").lower() == "y"

    # Создание текстового файла для неактивных ссылок
    with open("inactive_links.txt", "w") as inactive_file:
        # Инициализация драйвера
        try:
            driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
        except Exception as e:
            print(f"Ошибка при запуске браузера: {e}")
            return

        # Запуск потока для отслеживания нажатия клавиши F8 (паузу можно будет включать и выключать в любой момент)
        pause_thread = threading.Thread(target=toggle_pause)
        pause_thread.daemon = True  # Поток будет завершаться вместе с основным процессом
        pause_thread.start()

        # Проверка каждой ссылки
        rows_to_delete = []
        for row_idx, row in enumerate(sheet.iter_rows(min_row=start_row, max_col=1, values_only=True), start=start_row):
            link = row[0]
            if link:
                print(f"Проверка строки {row_idx}: {link}")

                # Проверка паузы
                while pause_flag:
                    time.sleep(0.1)  # Ожидание, пока пауза активна

                status = check_viber_group_status(link, driver)
                print(f"Ссылка: {link} — Статус: {status}")

                # Если ссылка неактивна, записываем в файл и отмечаем для удаления
                if status == "Неактивна":
                    inactive_file.write(f"Строка {row_idx}: {link}\n")
                    rows_to_delete.append(row_idx)  # Отметка строки для удаления

                # Пауза, если нажата клавиша F8
                if keyboard.is_pressed('F8'):
                    print("Пауза... Нажмите F8 для продолжения.")
                    while not keyboard.is_pressed('F8'):  # Ожидание, пока не нажмется F8
                        time.sleep(0.1)

        # Удаление неактивных ссылок из Excel, если выбрано
        if delete_links:
            for row_idx in sorted(rows_to_delete, reverse=True):
                sheet.delete_rows(row_idx)

            # Сохранение изменений в файле Excel
            workbook.save(path_to_excel)
            print(f"Неактивные ссылки удалены и изменения сохранены в {path_to_excel}.")
        else:
            print("Неактивные ссылки не были удалены из таблицы.")

        driver.quit()  # Закрытие окна браузера

    print("Проверка завершена. Неактивные ссылки сохранены в inactive_links.txt.")


if __name__ == "__main__":
    main()
