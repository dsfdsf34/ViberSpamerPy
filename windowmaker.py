import json
import os
import pygetwindow as gw
import win32gui
import keyboard

# Имя файла для сохранения данных
WINDOW_POSITIONS_FILE = "window_positions.json"


def is_window_visible(hwnd):
    """Проверяет, является ли окно развернутым и видимым."""
    return win32gui.IsWindowVisible(hwnd) and not win32gui.IsIconic(hwnd)


def is_system_window(hwnd):
    """Проверяет, является ли окно системным."""
    class_name = win32gui.GetClassName(hwnd)
    system_classes = ["ApplicationFrameWindow", "Progman", "Windows.UI.Core.CoreWindow"]
    return class_name in system_classes


def get_windows_in_taskbar_order():
    """Возвращает список видимых окон в порядке панели задач."""
    hwnd_list = []

    def enum_windows_proc(hwnd, lParam):
        if is_window_visible(hwnd) and not is_system_window(hwnd):
            title = win32gui.GetWindowText(hwnd).strip()
            if title:  # Только окна с заголовками
                hwnd_list.append((hwnd, title))

    win32gui.EnumWindows(enum_windows_proc, None)
    return hwnd_list


def save_window_positions():
    """Сохраняет позиции окон в порядке панели задач в файл."""
    hwnd_list = get_windows_in_taskbar_order()
    window_positions = []

    for hwnd, title in hwnd_list:
        try:
            win = gw.Window(hwnd)
            window_positions.append({
                "title": title,
                "x": win.left,
                "y": win.top,
                "width": win.width,
                "height": win.height,
            })
        except Exception as e:
            print(f"Не удалось получить информацию об окне '{title}': {e}")

    with open(WINDOW_POSITIONS_FILE, "w", encoding="utf-8") as file:
        json.dump(window_positions, file, indent=4)

    print(f"Позиции окон сохранены в файл '{WINDOW_POSITIONS_FILE}'.")


def restore_window_positions():
    """Восстанавливает окна из сохраненных позиций."""
    if not os.path.exists(WINDOW_POSITIONS_FILE):
        print(f"Файл '{WINDOW_POSITIONS_FILE}' не найден. Сначала сохраните позиции окон.")
        return

    with open(WINDOW_POSITIONS_FILE, "r", encoding="utf-8") as file:
        saved_positions = json.load(file)

    open_windows = gw.getAllWindows()  # Все текущие окна

    for saved_pos in saved_positions:
        title = saved_pos["title"]
        matching_windows = [win for win in open_windows if win.title == title]

        if matching_windows:
            # Берем первое совпавшее окно
            win = matching_windows[0]
            open_windows.remove(win)  # Удаляем из списка обработанных окон
            try:
                win.moveTo(saved_pos["x"], saved_pos["y"])
                win.resizeTo(saved_pos["width"], saved_pos["height"])
                print(f"Окно '{title}' перемещено в сохраненную позицию.")
            except Exception as e:
                print(f"Не удалось переместить окно '{title}': {e}")
        else:
            print(f"Окно '{title}' не найдено. Возможно, оно закрыто или перезапущено.")


def main():
    print("Горячие клавиши:")
    print("- Нажмите 'F7' для сохранения текущих позиций окон")
    print("- Нажмите 'F8' для восстановления позиций окон из сохраненных")
    print("- Нажмите 'F9' для выхода из программы")

    # Обработчики горячих клавиш
    keyboard.add_hotkey("F7", save_window_positions)
    keyboard.add_hotkey("F8", restore_window_positions)
    keyboard.add_hotkey("F9", exit)

    print("Ожидание ввода горячих клавиш...")
    keyboard.wait("F9")  # Программа завершится, когда нажмете F9


if __name__ == "__main__":
    main()
