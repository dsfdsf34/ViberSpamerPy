import pandas as pd
import pyautogui
import time
import keyboard
import json
import os
import sys  # Для выхода из скрипта

print(f"F8 для паузы/возобновления во время цикла")
print(f"F9 для сохранения текущих координат курсора")
print(f"F10 для прерывания скрипта")

# Глобальная переменная для хранения координат
current_coordinates = None
last_processed_rows = []  # Глобальная переменная для хранения последних обработанных строк


# Функция, вызываемая по нажатию горячей клавиши F9, для сохранения текущих координат
def save_current_coordinates():
    global current_coordinates
    current_coordinates = pyautogui.position()
    print(f"Координаты сохранены: {current_coordinates}")


keyboard.add_hotkey('f9', save_current_coordinates)


# Функция для получения координат
def get_coordinates(action_name, window_number):
    global current_coordinates
    print(
        f"Наведите курсор на расположение для '{action_name}' в окне {window_number + 1} и нажмите F9 для сохранения координат.")

    # Ждём, пока координаты не будут сохранены
    while current_coordinates is None:
        time.sleep(0.1)

    coords = current_coordinates
    current_coordinates = None  # Сбрасываем глобальные координаты после использования
    print(f"Координаты для '{action_name}' в окне {window_number + 1}: {coords}")
    return coords


# Функция для загрузки координат из файла
def load_coordinates(file_path='coords.json'):
    if os.path.exists(file_path):
        with open(file_path, 'r') as file:
            return json.load(file)
    return {}


# Функция для сохранения координат в файл
def save_coordinates(coords, file_path='coords.json'):
    # Загружаем текущие координаты
    existing_coords = load_coordinates(file_path)

    # Обновляем только нужные координаты
    for key in coords:
        existing_coords[key] = coords[key]

    with open(file_path, 'w') as file:
        json.dump(existing_coords, file)


# Функция для сохранения состояния (пути к файлу и строк)
def save_state(excel_path, start_rows, state_file='state.json'):
    state = {
        "excel_path": excel_path,
        "start_rows": start_rows
    }
    with open(state_file, 'w') as file:
        json.dump(state, file)


# Функция для загрузки состояния
def load_state(state_file='state.json'):
    if os.path.exists(state_file):
        with open(state_file, 'r') as file:
            return json.load(file)
    return {"excel_path": "", "start_rows": []}


# Основная функция автоматизации
def automate_viber_process(excel_file, total_cycles, start_rows, window_count, coords):
    global last_processed_rows  # Указываем, что это глобальная переменная

    # Чтение данных из Excel
    df = pd.read_excel(excel_file)

    # Список для хранения последних обработанных строк для каждого окна
    last_processed_rows = [0] * window_count
    processed_count = [0] * window_count  # Счетчик обработанных строк для каждого окна

    for cycle in range(total_cycles):
        for window_index in range(window_count):
            # Проверяем, остались ли строки для обработки
            if start_rows[window_index] >= len(df):
                print(f"Нет больше строк для обработки в окне {window_index + 1}. Пропуск...")
                continue

            # Проверка на паузу
            if paused:
                print("Скрипт приостановлен. Нажмите F8 для возобновления.")
                while paused:
                    time.sleep(1)  # Ждём, пока не нажмут F8

            # Получаем ссылку из Excel для текущего окна
            group_link = df.iloc[start_rows[window_index], 0]  # Предполагаем, что ссылки в первом столбце
            print(f"Обработка окна {window_index + 1}, строка {start_rows[window_index] + 1}")

            # 2. Активация нужного окна Viber
            pyautogui.click(coords['window'][window_index])
            time.sleep(0.5)

            # 3. Вставка ссылки и отправка сообщения
            pyautogui.write(group_link)
            pyautogui.press('enter')
            time.sleep(2)

            # 4. Клик на отправленной ссылке
            pyautogui.click(coords['link'][window_index])
            time.sleep(2)

            # Нажимаем Enter после клика на отправленной ссылке
            pyautogui.press('enter')
            time.sleep(1)

            # 5. Нажать кнопку "Присоединиться"
            pyautogui.click(coords['join'][window_index])
            time.sleep(2)

            # 6. Открываем поиск контактов
            pyautogui.click(coords['search'][window_index])
            time.sleep(1)

            # 7. Выбор контакта
            pyautogui.click(coords['select'][window_index])
            time.sleep(1)

            # Увеличиваем номер строки для текущего окна
            last_processed_rows[window_index] = start_rows[window_index] + 1
            start_rows[window_index] += 1

            # Увеличиваем счетчик обработанных строк
            processed_count[window_index] += 1
            print(f"Окно {window_index + 1}: +{processed_count[window_index]} строк обработано")

    print("Все циклы завершены.")
    print("Последние обработанные строки для каждого окна:")
    for i in range(window_count):
        print(f"Окно {i + 1}: строка {last_processed_rows[i]}")
    print("Время завершения:", time.strftime("%H:%M:%S"))


# Флаг для паузы
paused = False

# Флаг для отслеживания вывода сообщений о паузе
pause_message_displayed = False


def toggle_pause():
    global paused, pause_message_displayed
    paused = not paused  # Переключаем состояние паузы

    if paused and not pause_message_displayed:
        pause_message_displayed = True
    elif not paused and pause_message_displayed:
        print("Скрипт возобновлен.")
        pause_message_displayed = False


def stop_script():
    print("Скрипт прерван. Текущее состояние:")
    for i in range(len(last_processed_rows)):
        print(f"Окно {i + 1}: последняя обработанная строка — {last_processed_rows[i]}")
    print("Время прерывания:", time.strftime("%H:%M:%S"))
    sys.exit(0)  # Завершение выполнения скрипта


if __name__ == "__main__":
    # Устанавливаем горячую клавишу F8 для начала и паузы скрипта
    keyboard.add_hotkey('f8', toggle_pause)

    # Устанавливаем горячую клавишу F10 для остановки скрипта
    keyboard.add_hotkey('f10', stop_script)

    # Загрузка состояния и координат из файлов
    state = load_state()
    saved_coords = load_coordinates()

    # Запрашиваем количество окон Viber
    window_count = int(input("Введите количество окон Viber (от 1 до 6): "))

    # Запрашиваем путь к файлу Excel или используем последний
    excel_file = input(
        f"Введите путь к файлу Excel (или нажмите Enter, чтобы использовать последний: '{state['excel_path']}'): ")
    if not excel_file:
        excel_file = state['excel_path']

    # Запрашиваем начальные строки для каждого окна или используем последние
    start_rows = []
    for i in range(window_count):
        last_row = state['start_rows'][i] if i < len(state['start_rows']) else 0
        start_row = input(
            f"Введите начальную строку для окна {i + 1} (или нажмите Enter, чтобы использовать последнюю: {last_row}): ")
        start_rows.append(int(start_row) if start_row else last_row)

    # Запрашиваем количество циклов
    total_cycles = int(input("Введите количество циклов: "))

    # Проверяем, использовать ли сохраненные координаты для всех окон сразу
    use_saved = input("Использовать сохраненные координаты для всех окон? (y/n): ").strip().lower()

    coords = {
        'window': [],
        'link': [],
        'join': [],
        'search': [],
        'select': []
    }

    if use_saved == 'y' and all(str(i) in saved_coords for i in range(window_count)):
        # Загружаем координаты для всех окон, если все окна имеют сохраненные координаты
        for i in range(window_count):
            coords['window'].append(tuple(saved_coords[str(i)]['window']))
            coords['link'].append(tuple(saved_coords[str(i)]['link']))
            coords['join'].append(tuple(saved_coords[str(i)]['join']))
            coords['search'].append(tuple(saved_coords[str(i)]['search']))
            coords['select'].append(tuple(saved_coords[str(i)]['select']))
    else:
        # Получаем координаты для каждого окна
        for i in range(window_count):
            overwrite = input(f"Перезаписать координаты для окна {i + 1}? (y/n): ").strip().lower()
            if overwrite == 'y':
                coords['window'].append(get_coordinates("активация окна", i))
                coords['link'].append(get_coordinates("клик на ссылку", i))
                coords['join'].append(get_coordinates("клик на кнопку 'Присоединиться'", i))
                coords['search'].append(get_coordinates("клик для поиска", i))
                coords['select'].append(get_coordinates("клик для выбора контакта", i))
            else:
                coords['window'].append(tuple(saved_coords[str(i)]['window']))
                coords['link'].append(tuple(saved_coords[str(i)]['link']))
                coords['join'].append(tuple(saved_coords[str(i)]['join']))
                coords['search'].append(tuple(saved_coords[str(i)]['search']))
                coords['select'].append(tuple(saved_coords[str(i)]['select']))

        # Сохраняем координаты для каждого окна
        for i in range(window_count):
            save_coordinates({
                str(i): {
                    'window': coords['window'][i],
                    'link': coords['link'][i],
                    'join': coords['join'][i],
                    'search': coords['search'][i],
                    'select': coords['select'][i],
                }
            })

    # Сохраняем состояние
    save_state(excel_file, start_rows)

    # Пауза 3 секунды перед началом выполнения
    print("Начинаем через 3 секунды...")
    time.sleep(3)

    # Запускаем автоматизацию
    automate_viber_process(excel_file, total_cycles, start_rows, window_count, coords)
