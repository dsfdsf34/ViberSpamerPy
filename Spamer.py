import pyautogui
import time
import keyboard
import json
import os

# Файл для сохранения координат
COORDINATES_FILE = 'coordinates.json'

def save_coordinates(arrow_coords, message_coords, scrollbar_coords, drag_coords=None):
    coordinates = {
        'arrow_coords': arrow_coords,
        'message_coords': message_coords,
        'scrollbar_coords': scrollbar_coords,
        'drag_coords': drag_coords  # Может быть None, если не используется метод перетаскивания
    }
    with open(COORDINATES_FILE, 'w') as f:
        json.dump(coordinates, f)

def load_coordinates():
    if os.path.exists(COORDINATES_FILE):
        with open(COORDINATES_FILE, 'r') as f:
            coordinates = json.load(f)
        return coordinates
    return None

def wait_for_f8(prompt):
    print(prompt + " (наведите курсор и нажмите F8 для сохранения координат)")
    keyboard.wait('f8')
    pos = pyautogui.position()
    print("Сохранено:", pos)
    return pos

def drag_scrollbar(scrollbar_top_x, scrollbar_top_y, scrollbar_bottom_y):
    # Перетаскиваем скроллбар вниз
    pyautogui.click(scrollbar_top_x, scrollbar_top_y)  # Клик по верхней части скроллбара
    time.sleep(0.2)
    pyautogui.dragTo(scrollbar_top_x, scrollbar_bottom_y, duration=1)  # Перемещаемся к нижней части скроллбара
    time.sleep(0.5)

def click_messages(window, message_coords):
    # Выбираем 12 сообщений поочерёдно с кликами и прокруткой вверх
    for i in range(12):
        if keyboard.is_pressed('f8'):
            print("Скрипт остановлен.")
            return False
        x, y = message_coords[window]
        print(f"Клик по сообщению {i + 1}...")
        pyautogui.click(x, y)
        time.sleep(0.2)
        print("Прокручиваем вверх...")
        pyautogui.scroll(90)
        time.sleep(0.2)
    return True

def drag_select_messages(window, drag_coords):
    # Выделение сообщений перетаскиванием мыши (для двух сообщений)
    drag_start_x, drag_start_y, drag_end_x, drag_end_y = drag_coords[window]
    print("Выделение сообщений (перетаскиванием)...")
    pyautogui.moveTo(drag_start_x, drag_start_y)
    pyautogui.mouseDown()
    time.sleep(0.2)
    pyautogui.moveTo(drag_end_x, drag_end_y, duration=1)
    pyautogui.mouseUp()
    time.sleep(0.5)

def forward_messages(num_cycles, num_windows, arrow_coords, message_coords, scrollbar_coords, selection_method, drag_coords):
    print("Начинается выполнение скрипта. У вас есть 1 секунда, чтобы переключиться на Viber.")
    time.sleep(1)
    for cycle in range(num_cycles):
        for window in range(num_windows):
            print(f"\nЦикл {cycle + 1}, Окно {window + 1}")
            if keyboard.is_pressed('f8'):
                print("Скрипт остановлен.")
                return

            # 1. Выделение сообщений
            if selection_method == 2:
                drag_select_messages(window, drag_coords)
            elif selection_method == 1:
                # Если выбран метод клика, просто фокусируемся на области сообщений
                print("Фокусировка на области сообщений (клик)...")
                x, y = message_coords[window]
                pyautogui.click(x, y)
                time.sleep(0.2)

            # 2. Нажатие на стрелку "Переслать"
            arrow_x, arrow_y = arrow_coords[window]
            print("Нажатие на стрелку 'Переслать'...")
            pyautogui.click(arrow_x, arrow_y)
            time.sleep(5)

            # 3. Прокрутка скроллбара в самый низ
            scrollbar_top_x, scrollbar_top_y, scrollbar_bottom_y = scrollbar_coords[window]
            print("Прокрутка скроллбара в самый низ...")
            drag_scrollbar(scrollbar_top_x, scrollbar_top_y, scrollbar_bottom_y)

            # 4. Выбор 12 сообщений поочерёдно с прокруткой вверх
            if selection_method == 1:
                if not click_messages(window, message_coords):
                    return
            elif selection_method == 2:
                # Если метод выделения drag использовался, то дополнительно выбираем сообщения по клику
                if not click_messages(window, message_coords):
                    return

            # 5. Нажатие Enter
            print("Нажатие 'enter'...")
            pyautogui.press('enter')
            time.sleep(5)
        if num_windows > 1:
            print("Переход к следующему окну...")
            time.sleep(5)
        else:
            print("Повторяем цикл для одного окна...")
            time.sleep(5)

if __name__ == "__main__":
    # В начале скрипта спрашиваем, сколько окон используем
    try:
        num_windows = int(input("Введите количество окон Viber (1 или 2): "))
        if num_windows not in [1, 2]:
            print("Неверное значение. Будет использовано 1 окно.")
            num_windows = 1
    except ValueError:
        print("Неверное значение. Будет использовано 1 окно.")
        num_windows = 1

    # Инициализация переменных для каждого окна
    arrow_coords = [None] * num_windows         # Координаты для кнопки "Переслать"
    message_coords = [None] * num_windows         # Координаты для клика по сообщению
    scrollbar_coords = [None] * num_windows       # Координаты скроллбара (верхняя и нижняя точки)
    drag_coords = [None] * num_windows            # Координаты для перетаскивания (выделение сообщений)

    # Выбор метода выделения сообщений
    try:
        selection_method = int(input("Выберите метод выделения сообщений:\n"
                                       "1 - Фокусировка (клик по сообщению)\n"
                                       "2 - Перетаскивание для выделения двух сообщений\n"
                                       "Ваш выбор (1 или 2): "))
    except ValueError:
        selection_method = 1

    # Запрос на загрузку сохранённых координат
    if input("Хотите загрузить сохраненные координаты? (y/n): ").strip().lower() == 'y':
        coordinates = load_coordinates()
        if coordinates:
            arrow_coords = coordinates['arrow_coords']
            message_coords = coordinates['message_coords']
            scrollbar_coords = coordinates['scrollbar_coords']
            if selection_method == 2:
                drag_coords = coordinates.get('drag_coords')
            print("Координаты загружены.")
        else:
            print("Сохраненные координаты не найдены. Введите координаты вручную.")

    # Если координаты не загружены, запрашиваем их с помощью F8
    for window in range(num_windows):
        if arrow_coords[window] is None:
            arrow_coords[window] = wait_for_f8(f"\nОкно {window + 1}: наведите курсор на стрелку 'Переслать'")
        if message_coords[window] is None:
            message_coords[window] = wait_for_f8(f"Окно {window + 1}: наведите курсор на сообщение (для клика)")
        if scrollbar_coords[window] is None:
            pos_top = wait_for_f8(f"Окно {window + 1}: наведите курсор на верхнюю часть скроллбара")
            pos_bottom = wait_for_f8(f"Окно {window + 1}: наведите курсор на нижнюю часть скроллбара")
            scrollbar_coords[window] = (pos_top.x, pos_top.y, pos_bottom.y)
        if selection_method == 2 and drag_coords[window] is None:
            pos_drag_start = wait_for_f8(f"Окно {window + 1}: наведите курсор на верхнее сообщение для выделения")
            pos_drag_end = wait_for_f8(f"Окно {window + 1}: наведите курсор на нижнее сообщение для выделения")
            drag_coords[window] = (pos_drag_start.x, pos_drag_start.y, pos_drag_end.x, pos_drag_end.y)

    # Сохраняем все координаты
    save_coordinates(arrow_coords, message_coords, scrollbar_coords, drag_coords)

    num_cycles = int(input("\nВведите количество циклов: "))
    forward_messages(num_cycles, num_windows, arrow_coords, message_coords,
                     scrollbar_coords, selection_method, drag_coords)
