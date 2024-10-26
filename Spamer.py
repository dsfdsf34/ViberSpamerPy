import pyautogui
import time
import keyboard
import json
import os

# Файл для сохранения координат
COORDINATES_FILE = 'coordinates.json'


def save_coordinates(arrow_coords, message_coords, scrollbar_coords):
    # Сохраняем координаты в JSON файл
    coordinates = {
        'arrow_coords': arrow_coords,
        'message_coords': message_coords,
        'scrollbar_coords': scrollbar_coords
    }
    with open(COORDINATES_FILE, 'w') as f:
        json.dump(coordinates, f)


def load_coordinates():
    # Загружаем координаты из JSON файла, если файл существует
    if os.path.exists(COORDINATES_FILE):
        with open(COORDINATES_FILE, 'r') as f:
            coordinates = json.load(f)
        return coordinates
    return None


def drag_scrollbar(scrollbar_top_x, scrollbar_top_y, scrollbar_bottom_y):
    # Перетаскиваем скроллбар вниз
    pyautogui.click(scrollbar_top_x, scrollbar_top_y)  # Клик на верхнюю часть скроллбара
    time.sleep(0.2)
    # Перетаскиваем вниз
    pyautogui.dragTo(scrollbar_top_x, scrollbar_bottom_y, duration=1)  # Перемещение вниз к нижней части скроллбара
    time.sleep(0.5)


def forward_messages(num_cycles, num_windows, arrow_coords, message_coords, scrollbar_coords):
    print("Начинается выполнение скрипта. У вас есть 1 сек, чтобы переключиться на Viber.")
    time.sleep(1)  # Время для переключения на Viber

    for cycle in range(num_cycles):
        for window in range(num_windows):
            print(f"Цикл {cycle + 1}, Окно {window + 1}")

            # Проверка нажатия клавиши F8 для остановки
            if keyboard.is_pressed('f8'):
                print("Скрипт остановлен.")
                return

            # Нажимаем на стрелку "Переслать"
            arrow_x, arrow_y = arrow_coords[window]
            print("Клик на стрелку 'Переслать'...")
            pyautogui.click(arrow_x, arrow_y)
            time.sleep(5)  # Задержка после клика

            # Перетаскиваем скроллбар вниз
            scrollbar_top_x, scrollbar_top_y, scrollbar_bottom_y = scrollbar_coords[window]
            print("Перетаскиваем скроллбар вниз...")
            drag_scrollbar(scrollbar_top_x, scrollbar_top_y, scrollbar_bottom_y)

            # Пересылаем сообщения
            for i in range(12):  # Пересылаем до 12 сообщений
                # Проверка нажатия клавиши F8 для остановки
                if keyboard.is_pressed('f8'):
                    print("Скрипт остановлен.")
                    return

                # Кликаем по сообщению
                message_x, message_y = message_coords[window]
                print(f"Клик по сообщению {i + 1}...")
                pyautogui.click(message_x, message_y)
                time.sleep(0.2)  # Задержка для удобства

                # Прокручиваем вверх после клика
                print("Прокручиваем вверх...")
                pyautogui.scroll(90)  # Прокрутка вверх на 100 пикселей
                time.sleep(0.2)  # Задержка для удобства

            # Нажимаем 'enter' после 8 кликов
            print("Нажимаем 'enter'...")
            pyautogui.press('enter')
            time.sleep(5)  # Небольшая пауза перед переходом к следующему окну или повтором цикла

        # Если есть больше одного окна, мы можем перейти ко второму
        if num_windows > 1:
            print("Переход ко второму окну...")
            time.sleep(5)  # Пауза перед переключением окна
        else:
            print("Повторяем цикл для одного окна...")
            time.sleep(5)  # Пауза перед повтором цикла


if __name__ == "__main__":
    # Инициализация переменных
    arrow_coords = [None, None]  # Координаты для 2 окон
    message_coords = [None, None]  # Координаты сообщений для 2 окон
    scrollbar_coords = [None, None]  # Координаты скроллбара для 2 окон

    # Запрос на загрузку координат
    load_coordinates_answer = input("Хотите загрузить сохраненные координаты? (y/n): ").strip().lower()

    # Проверяем ответ пользователя
    if load_coordinates_answer == 'y':
        coordinates = load_coordinates()
        if coordinates:
            arrow_coords = coordinates['arrow_coords']
            message_coords = coordinates['message_coords']
            scrollbar_coords = coordinates['scrollbar_coords']
            print("Координаты загружены.")
        else:
            print("Сохраненные координаты не найдены. Пожалуйста, введите координаты вручную.")

    for window in range(2):  # Запрашиваем координаты для двух окон
        if arrow_coords[window] is None or message_coords[window] is None or scrollbar_coords[window] is None:
            print(f"Укажите координаты для окна {window + 1}:")
            print("Наведите курсор на стрелку 'Переслать'.")
            time.sleep(5)  # Время для указания координат
            arrow_coords[window] = pyautogui.position()
            print("Координаты стрелки 'Переслать' сохранены:", arrow_coords[window])

            print("Теперь наведите курсор на сообщение.")
            time.sleep(5)  # Время для указания координат
            message_coords[window] = pyautogui.position()
            print("Координаты сообщения сохранены:", message_coords[window])

            print("Теперь наведите курсор на верхнюю часть скроллбара.")
            time.sleep(5)  # Время для указания координат
            scrollbar_top_x, scrollbar_top_y = pyautogui.position()
            print("Координаты верхней части скроллбара сохранены:", scrollbar_top_x, scrollbar_top_y)

            print("Теперь наведите курсор на нижнюю часть скроллбара.")
            time.sleep(5)  # Время для указания координат
            scrollbar_bottom_x, scrollbar_bottom_y = pyautogui.position()
            print("Координаты нижней части скроллбара сохранены:", scrollbar_bottom_x, scrollbar_bottom_y)

            # Сохраняем координаты для текущего окна
            scrollbar_coords[window] = (scrollbar_top_x, scrollbar_top_y, scrollbar_bottom_y)

    # Сохраняем все координаты
    save_coordinates(arrow_coords, message_coords, scrollbar_coords)

    # Запрашиваем ввод пользователя
    num_cycles = int(input("Введите количество циклов: "))
    num_windows = int(input("Введите количество окон Viber (1 или 2): "))

    if num_windows not in [1, 2]:
        print("Пожалуйста, выберите 1 или 2 окна Viber.")
    else:
        forward_messages(num_cycles, num_windows, arrow_coords, message_coords, scrollbar_coords)
