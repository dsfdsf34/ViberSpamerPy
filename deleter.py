import os
import json
import shutil

def load_settings(settings_file="set.json"):
    """Загружает настройки из файла."""
    if os.path.exists(settings_file):
        with open(settings_file, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_settings(settings, settings_file="set.json"):
    """Сохраняет настройки в файл."""
    with open(settings_file, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=4)

def find_folders_to_delete(root_folder, folder_names):
    """
    Находит пути всех папок, которые будут удалены.

    :param root_folder: Строка. Путь к общей папке.
    :param folder_names: Список строк. Названия папок для удаления.
    :return: Список строк. Полные пути папок для удаления.
    """
    folders_to_delete = []
    for foldername, dirnames, _ in os.walk(root_folder, topdown=False):
        for dirname in dirnames:
            if dirname in folder_names:
                folder_path = os.path.join(foldername, dirname)
                folders_to_delete.append(folder_path)
    return folders_to_delete

def delete_folders(folders_to_delete):
    """
    Удаляет папки из списка.

    :param folders_to_delete: Список строк. Полные пути папок для удаления.
    """
    for folder_path in folders_to_delete:
        try:
            shutil.rmtree(folder_path)
            print(f"Удалена папка: {folder_path}")
        except Exception as e:
            print(f"Ошибка при удалении папки {folder_path}: {e}")

def main():
    settings = load_settings()

    # Запрос основной папки у пользователя
    root_folder = settings.get("root_folder")
    if root_folder:
        use_saved = input(f"Использовать сохранённый путь: {root_folder}? (y/n): ").strip().lower()
        if use_saved != "y":
            root_folder = None

    if not root_folder:
        root_folder = input("Введите путь к основной папке: ").strip()
        settings["root_folder"] = root_folder

    # Запрос названий папок для удаления
    folder_names = settings.get("folder_names")
    if folder_names:
        use_saved_folders = input(f"Использовать сохранённые папки для удаления: {', '.join(folder_names)}? (y/n): ").strip().lower()
        if use_saved_folders != "y":
            folder_names = None

    if not folder_names:
        folder_names = input("Введите названия папок для удаления (через запятую): ").strip().split(",")
        folder_names = [name.strip() for name in folder_names]
        settings["folder_names"] = folder_names

    # Сохранение настроек
    save_settings(settings)

    # Найти папки для удаления
    folders_to_delete = find_folders_to_delete(root_folder, folder_names)

    # Показать папки, которые будут удалены
    if folders_to_delete:
        print("Будут удалены следующие папки:")
        for folder in folders_to_delete:
            print(folder)
        confirm = input("Вы уверены, что хотите удалить эти папки? (y/n): ").strip().lower()
        if confirm == "y":
            # Удаление папок
            delete_folders(folders_to_delete)
            print("Удаление завершено.")
        else:
            print("Удаление отменено.")
    else:
        print("Папки для удаления не найдены.")

if __name__ == "__main__":
    main()
