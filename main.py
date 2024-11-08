import subprocess
import sys

def main():
    # Путь к виртуальному окружению, измените путь на свой
    venv_path = 'C:/Users/mitsumi/PycharmProjects/ViberSpamBot/.venv/Scripts/activate'

    while True:
        print("\nВыберите, какой скрипт хотите выполнить:")
        print("1 - Joiner.py")
        print("2 - LinksChecker.py")
        print("3 - Spamer.py")
        print("4 - OpenLinks.py")
        print("0 - Завершить работу")

        choice = input("Введите номер выбранного скрипта: ")

        # Проверка выбора и определение пути к скрипту
        script_path = None
        if choice == "1":
            script_path = 'C:/Users/mitsumi/PycharmProjects/ViberSpamBot/Joiner.py'
        elif choice == "2":
            script_path = 'C:/Users/mitsumi/PycharmProjects/ViberSpamBot/LinksChecker.py'
        elif choice == "3":
            script_path = 'C:/Users/mitsumi/PycharmProjects/ViberSpamBot/Spamer.py'
        elif choice == "4":
            script_path = 'C:/Users/mitsumi/PycharmProjects/ViberSpamBot/OpenLinks.py'
        elif choice == "0":
            print("Завершение работы программы.")
            sys.exit()
        else:
            print("Неверный выбор. Попробуйте снова.")
            continue

        # Запросить пользователя, хочет ли он использовать права администратора
        admin_choice = input("Хотите ли вы запустить скрипт с правами администратора? (y/n): ").lower()

        if admin_choice == 'y':
            # Запуск с правами администратора
            subprocess.run(
                f'runas /user:mitsumi "cmd /k {venv_path} && python {script_path}"',
                shell=True
            )
        else:
            # Открытие выбранного скрипта в новом окне командной строки с активированным виртуальным окружением
            subprocess.run(
                f'start cmd /k "{venv_path} && python {script_path}"',
                shell=True
            )

        # После запуска скрипта, больше не возвращаемся к выбору
        print("Скрипт запущен в новом окне с использованием виртуального окружения.")
        continue

if __name__ == "__main__":
    main()
