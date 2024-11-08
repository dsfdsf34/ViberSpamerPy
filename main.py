import subprocess
import sys

def main():
    print("Выберите, какой скрипт хотите выполнить:")
    print("1 - Joiner.py")
    print("2 - LinksChecker.py")
    print("3 - Spamer.py")
    print("0 - Выйти из программы")

    choice = input("Введите номер выбранного скрипта: ")

    if choice == "1":
        process = subprocess.Popen(['python', 'C:/Users/mitsumi/PycharmProjects/ViberSpamBot/Joiner.py'])
    elif choice == "2":
        process = subprocess.Popen(['python', 'C:/Users/mitsumi/PycharmProjects/ViberSpamBot/LinksChecker.py'])
    elif choice == "3":
        process = subprocess.Popen(['python', 'C:/Users/mitsumi/PycharmProjects/ViberSpamBot/Spamer.py'])
    elif choice == "0":
        print("Завершение работы программы.")
        sys.exit()  # Завершает выполнение программы
    else:
        print("Неверный выбор. Попробуйте снова.")
        main()  # Повторить запрос, если выбор некорректен

    # Цикл для проверки ввода команды для завершения работы
    while True:
        print("Нажмите 'q' для остановки текущего скрипта.")
        user_input = input()
        if user_input.lower() == 'q':
            print("Останавливаем выполнение скрипта...")
            process.terminate()  # Завершаем процесс
            print("Скрипт остановлен.")
            break  # Прерываем цикл и возвращаемся к выбору

if __name__ == "__main__":
    main()
