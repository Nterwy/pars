# Import the os module for interacting with the operating system
import os
import sys
import tkinter as tk
from tkinter import messagebox

# Import the time module for time-related operations (though not used in the current code)
import time
import pickle
import random
import requests
import pandas as pd

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager  # Ensure webdriver_manager is installed
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup

def get_base_dir():
    if getattr(sys, 'frozen', False):  # Если запущено как exe
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

base_dir = get_base_dir()

# Создаем окно
# Создание главного окна
root = tk.Tk()
root.title("Программа с интерфейсом")
root.geometry("300x430")  # Размер окна
root.title("Scrapper")
root.configure(bg="black")  # Цвет фона окна
instruction = tk.Label(root, text="Ключевые слова:", font=("Times New Roman", 15), bg="white", fg="black")
instruction.place(x=150, y=270, anchor="center")  # Размещение текста в центре окна
EnterTextKeyWords = tk.Entry(root, font=("Times New Roman", 15), bg="white", fg="black")
EnterTextKeyWords.place(x=150, y=300, anchor="center")  # Поле ввода текста

results_file = os.path.join(base_dir, "video_results.xlsx")
existing_file = os.path.join(base_dir, "existing_videos.xlsx")
cookies_path = os.path.join(base_dir, "cookies", "cookies.pkl")

# Массив для хранения ключевых слов
words = [
    "farmer","farm","smalltown","harvest","tractor","barn","foreclosure","sheriffsale",
    "flood","storm","auction","abandonedfarm","veteran","soldier","marine","army",
    "navy","SEAL","general","colonel","helicopter","convoy","president","governor",
    "senator","billionaire","millionaire","biker","bikergang","outlaw","thug","robber",
    "truckdriver","trucker","lawyer","advisor","broker","dishwasher","waitress","widow",
    "senior","elder","woman","harass","bully","mock","threaten","trap","surround",
    "martialart","combat","sniper","sharpshooter","cowboy","gunslinger","DeltaForce",
    "SpecialOp","CIA","mechanic","luxurycar","blackSUV","aircraft","typhoon","pilot",
    "militarytruck","blizzard","rancher","private","train","railway","crash","foodtruck",
    "celebritychef","submarine","liferaft","militaryjeep","pararescue","airforce","chopper",
    "humvee","blackhawk","RollsRoyce","firefighting","hospital","janitor","officer","banker",
    "security","lifepurpose","stranded","rescue","medicine","lostchild","shelter",
    "militarypolice","truckroute","service","medal","celebrity","stormvictim","junkyard",
    "memorial","oldk9","heartattack","heartful","savinglife","homelessveteran","roadsidehelp",
    "airline","kickoffflight","elderlyman","shockindustry","teenbully","bikershowup",
    "dinerowner","helicopterland","cafe","violin","roomsilent","wheelchair","lineup",
    "lifechangingcall","brokewoman","luxurycararound","airstrip","lockheed","c5msupergalaxy",
    "farmhouse","germangirl","heroism","waitresskindness","billionairechange","ranchowner",
    "saving","veteranfamily","militarysupport","secretfound","haircut","fired","nursedonate",
    "patient","mother","necklace","simplewoman","busstopjacket","ceo","changedlife",
    "discrimination","restaurant","poorgirl","beggar","luckychanged","restaurantgasp",
    "hungryveteran","trooparrive","storyrevealed","teacher","billionairehelicopter","job",
    "blackSUVstorm","repair","truthrevealed","stranger","lifetransform","tankrollup","meal",
    "convoy","girl","strangerhelp","singledad","daycaremistake","lovefound","grandma",
    "necklace2","moment","divorce","textrevenge","friend","jobchange","janitortruth",
    "collegefriend","shockreaction"
]
# Фукнция для записи EnterTextKeyWords в массив words
def EnterTextKeyWordsToArray():
    global words
    # Получаем текст из поля ввода
    input_text = EnterTextKeyWords.get()
    # Разбиваем текст по запятым и убираем пробелы, приводим к нижнему регистру (опционально)
    words = [word.strip() for word in input_text.split(",") if word.strip()]
    print(words)  # Выводим массив в консоль для проверки


def button1_click():
    messagebox.showinfo("Открыть браузер", "Вы Открываете браузер")
    driver.get('https://www.google.com/')
    # Ждем некоторое время (если нужно)
    time.sleep(1)
    driver.get('https://www.youtube.com/')


def save_cookies(driver, path=cookies_path):
    messagebox.showinfo("Сохранить куки", "Вы Сохраняете куки")
    cookies = driver.get_cookies()
    unique_cookies = {cookie['name']: cookie for cookie in cookies}

    with open(path, "wb") as f:
        pickle.dump(list(unique_cookies.values()), f)
    print(f"Сохранено {len(unique_cookies)} уникальных куки")


def load_cookies(driver, path=cookies_path):
    messagebox.showinfo("Загрузить куки", "Вы Загружаете куки")
    driver.delete_all_cookies()

    with open(path, "rb") as f:
        cookies = pickle.load(f)

    loaded_names = set()
    for cookie in cookies:
        if cookie['name'] in loaded_names:
            continue
        try:
            driver.add_cookie(cookie)
            loaded_names.add(cookie['name'])
        except Exception as e:
            print(f"Ошибка: {e}")

    driver.refresh()

# # Используя Селениум, получаем DOM-структуру страницы и сохраняем её в файл
# def get_dom_structure(driver, filename='index.html'):
#     try:
#         full_dom = driver.page_source
#         print(f"Длина DOM: {len(full_dom)} символов")
#         if not full_dom.strip():
#             print("Предупреждение: получен пустой DOM")
#             return False

#         # Формируем путь к папке TryToParsing
#         save_dir = os.path.join(os.getcwd(), "TryToParsing")
#         # Создаем папку, если она не существует
#         if not os.path.exists(save_dir):
#             os.makedirs(save_dir)
#         # Формируем полный путь к файлу
#         file_path = os.path.join(save_dir, filename)

#         print(f"Сохранение DOM в {file_path}")
#         with open(file_path, 'w', encoding='utf-8') as file:
#             file.write(full_dom)
#         print(f"DOM успешно сохранён в {file_path}")
#         return True
#     except Exception as e:
#         print(f"Ошибка при получении DOM: {str(e)}")
#         return False



# Функция для поиска видео по ключевым словам с помощью Selenium

def Search_byWords(driver, words):
    try:
        links = driver.find_elements('xpath', "//div//div//a[@id='video-title-link']")
        titles = driver.find_elements('xpath', "//a[@id='video-title-link']//yt-formatted-string[@id='video-title']")
        Amount_of_Views = driver.find_elements('xpath', '//div[@id="metadata-line"]/span[1]')
        links_to_channels = driver.find_elements('xpath', '//div[@id = "container"]//div[@id = "text-container"]//a')

        print(f"Found titles: {len(titles)}")
        print(f"Found links: {len(links)}")
        print(f"Found views: {len(Amount_of_Views)}")
        print(f"Found channels: {len(links_to_channels)}")

        video_data = []

        # Определяем минимальную длину из всех списков
        min_length = min(len(titles), len(links), len(Amount_of_Views), len(links_to_channels))

        for i in range(min_length):  # Итерируемся по минимальному размеру
            title_text = titles[i].text.strip()
            if not title_text:
                continue

            if any(word.lower() in title_text.lower() for word in words):
                try:
                    print(f"Title: {title_text}")
                    print(f"Link: {links[i].get_attribute('href')}")
                    print('-' * 20)

                    video_data.append({
                        "Название видео": title_text,
                        "Ссылка": links[i].get_attribute('href'),
                        "Количество просмотров": Amount_of_Views[i].text,
                        "Ссылка на канал": links_to_channels[i].get_attribute('href')
                    })
                except UnicodeEncodeError as e:
                    print(f"Ошибка вывода заголовка: {e}")
                except Exception as e:
                    print(f"Ошибка при обработке элемента {i}: {e}")

        if video_data:
            df = pd.DataFrame(video_data)
            if os.path.exists(results_file):
                existing_df = pd.read_excel(results_file)
                combined_df = pd.concat([existing_df, df], ignore_index=True)
                combined_df.to_excel(results_file, index=False)
            else:
                df.to_excel(results_file, index=False)
            print("Data successfully saved to video_results.xlsx")
        else:
            print("No data to write to Excel")

    except Exception as e:
        print(f"An error occurred: {e}")


def DeleteExistingVideos():

    # Проверяем, существуют ли файлы
    if not os.path.exists(results_file):
        messagebox.showerror("Ошибка", f"Файл {results_file} не найден.")
        return
    if not os.path.exists(existing_file):
        messagebox.showerror("Ошибка", f"Файл {existing_file} не найден.")
        return

    # Загружаем данные из файлов
    try:
        df_results = pd.read_excel(results_file)
        df_existing = pd.read_excel(existing_file)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка при чтении файлов: {e}")
        return

    # Собираем все значения из всех ячеек existing_videos.xlsx
    existing_values = set()
    for row in df_existing.values: # Перебираем все строки
        for cell in row: # Перебираем все ячейки в строке
            if pd.notna(cell): # Проверяем, что ячейка не пуста
                existing_values.add(str(cell).strip().lower()) # Добавляем значение в множество

    # Функция для проверки, содержит ли строка совпадения с existing_values
    def has_matching_value(row):
        for cell in row:
            if pd.notna(cell) and str(cell).strip().lower() in existing_values:
                return True
        return False

    # Удаляем строки с совпадающими значениями
    df_results_filtered = df_results[~df_results.apply(has_matching_value, axis=1)]

    # Удаляем дубликаты по четвертому столбцу (с индексом 3)
    if df_results_filtered.shape[1] >= 4:
        df_results_filtered = df_results_filtered.drop_duplicates(subset=df_results_filtered.columns[3])
    else:
        messagebox.showwarning("Предупреждение", "В файле недостаточно столбцов для удаления дубликатов по четвёртому столбцу.")

    # Удаляем дубликаты по столбцу со ссылками на YouTube-каналы (предположим, это столбец с индексом 2)
    channel_column_index = 2  # Укажите индекс столбца, где находятся ссылки на каналы
    if df_results_filtered.shape[1] > channel_column_index:
        df_results_filtered = df_results_filtered.drop_duplicates(subset=df_results_filtered.columns[channel_column_index])
    else:
        messagebox.showwarning("Предупреждение", "В файле недостаточно столбцов для удаления дубликатов по столбцу ссылок на каналы.")

    # Сохраняем обновленный файл
    try:
        df_results_filtered.to_excel(results_file, index=False)
        messagebox.showinfo("Успех", "Файл video_results.xlsx обновлен: удалены совпадения, дубликаты по четвертому столбцу и дубликаты ссылок на каналы.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка при сохранении файла: {e}")


options = Options()
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36")

service = Service(executable_path=ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 30, poll_frequency = 1)

print(driver.get_cookies())

driver.refresh()

# Создание кнопок и их размещение
button1 = tk.Button(root, text="Открыть браузер", command=button1_click)
button1.pack(pady=10)

button2 = tk.Button(root, text="Сохранить куки", command=lambda: save_cookies(driver, cookies_path))
button2.pack(pady=10)

button3 = tk.Button(root, text="Загрузить куки", command=lambda: load_cookies(driver, cookies_path))
button3.pack(pady=10)

# button4 = tk.Button(root, text="Получить DOM", command=lambda: get_dom_structure(driver, filename='index.html'))
# button4.pack(pady=10)

button5 = tk.Button(root, text="Записать все ссылки и названия", command=lambda: Search_byWords(driver, words))
button5.pack(pady=10)

button6 = tk.Button(root, text="Удалить существующие видео", command=DeleteExistingVideos)
button6.pack(pady=10)

button7 = tk.Button(root, text="Добавить новые слова", command=EnterTextKeyWordsToArray)
button7.pack(pady=90)
# Запуск основного цикла обработки событий
# Обработчик закрытия окна
def on_closing():
    try:
        driver.quit()
    except:
        pass
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)

root.mainloop()
