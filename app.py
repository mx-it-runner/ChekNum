import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
import pandas as pd
import sqlite3
import re

# В ТТК нет изначально плейсхолдера поэтому прописываем эту функциональность вручную
class PlaceholderEntry(ttk.Entry):
    def __init__(self, master=None, placeholder="", *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.placeholder = placeholder
        self.placeholder_color = "grey"
        self.normal_color = self["foreground"]
        self.bind("<FocusIn>", self._clear_placeholder)
        self.bind("<FocusOut>", self._add_placeholder)
        self._add_placeholder()

    def _clear_placeholder(self, event):
        if self["foreground"] == self.placeholder_color:
            self.delete(0, tk.END)
            self["foreground"] = self.normal_color

    def _add_placeholder(self, event=None):
        if not self.get():
            self.insert(0, self.placeholder)
            self["foreground"] = self.placeholder_color

# Объявляем глобалье списки и переменные
phone_num = []
all_num = 0
complit_num = 0
error_num = 0
global_data = []
unprocessed_results = []

# Массовая обработка
def read_file_numbers():
    
    filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")])
    
    global phone_num
    global error_num
    global unprocessed_results
    global all_num

    unprocessed_results = []
    phone_num = []       #-----------#
    data_set = []        #--- clear--#
    error_num = 0        #----list---#
    all_num = 0          #-----------#
    
    if filename.endswith('.xls') or filename.endswith('.xlsx'):
        df = pd.read_excel(filename, dtype=str, header=None)
        if not df.empty:
            # Путь к файлу для отображения
            file_label.config(text="Выбранный файл:\n" + filename)  
            
            data_set.extend(df.iloc[:, 0].tolist())
            for num in data_set:
                # Удаляем пробелы и другие нежелательные символы
                num = re.sub(r'\s+|-|\.|,|\(|\)|\[|\]|"|\'', '', num)
                if len(str(num)) == 10 and int(num[0:3]) > 899:
                    num = '7' + num
                    all_num += 1
                    phone_num.append(num)
                elif len(str(num)) == 11 and int(num[1:4]) > 899:
                    all_num += 1
                    phone_num.append(num)
                else:
                    unprocessed_results.append(num)
                    error_num += 1 
    else:
        messagebox.showerror("Ошибка", "Неподдерживаемый формат файла.")

def chek_numbers():
    global phone_num
    global global_data
    global unprocessed_results
    global error_num
    global complit_num

    global_data = []
    error_num = 0
    complit_num = 0

    conn = sqlite3.connect('operator_and_region.db')
    cursor = conn.cursor()

    for num in phone_num:
        kod_operatora = num[1:4]
        nomer =  num[4:11]

        cursor.execute("SELECT * FROM data WHERE cod=?", (kod_operatora,))
        results = cursor.fetchall()

        found = False # Флаг, указывающий, был ли найден оператор для номера

        for result in results:
            ot, do = result[1], result[2]  # Извлечение значений "От" и "До" из результата

            if int(ot) <= int(nomer) <= int(do):
                operator, region = result[3], result[4]
                global_data.append((num, operator, region))
                found = True # Устанавливаем флаг, что оператор был найден
                complit_num += 1
                break

        if not found:
            unprocessed_results.append(num)
            error_num += 1

    # Запуск окна с результатами
    if global_data: # Используем global_data для отображения только правильно обработанных номеров
        open_results_window(global_data)
    else:
        messagebox.showwarning("Предупреждение", "ВЫБЕРИТЕ КОРРЕКТНЫЙ ФАЙЛ ДЛЯ ОБРАБОТКИ")
    
    # Закрываем соединение с базой данных
    conn.close()

def open_results_window(results):
    results_window = tk.Toplevel()
    results_window.title("Результаты")
    results_window.geometry("490x350")

    label_fullprocessed = tk.Label(results_window, text=f"Кол-во номеров: {all_num}")
    label_fullprocessed.pack()
    
    label_processed = tk.Label(results_window, text=f"Успешно обработано: {complit_num}")
    label_processed.pack()

    label_unprocessed = tk.Label(results_window, text=f"Не удалось обработать: {error_num}")
    label_unprocessed.pack()

    # Создание и настройка фрейма для кнопок "Скачать"

    download_button = ttk.Button(results_window, text="Сохранить результаты", command=save_to_excel)
    download_button.pack(pady=5)

    download_error_button = ttk.Button(results_window, text="Сохранить необработанные номера", command=save_unprocessed)
    download_error_button.pack(pady=5)

    # Создание и настройка текстового виджета для результатов
    results_text = tk.Text(results_window, font=("Arial", 12))
    scrollbar = tk.Scrollbar(results_window, command=results_text.yview)
    scrollbar.pack(side="right", fill="y", pady=(10, 10))
    results_text.config(yscrollcommand=scrollbar.set)
    results_text.pack(expand=True, fill="both", padx=(10, 0), pady=(10, 10))

    
    # Заполнение текстового виджета результатами
    for result in results:
        number, operator, region = result  # Разделение кортежа на отдельные элементы
        formatted_result = f"Номер: {number}\nОператор: {operator}\nРегион: {region}\n\n"
        results_text.insert(tk.END, formatted_result)


def save_to_excel():
    global global_data

    df = pd.DataFrame(global_data, columns=['Номер', 'Оператор', 'Регион'])

    filename = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[("Excel Files", "*.xlsx")])
    
    if filename:
        df.to_excel(filename, index=False)
        messagebox.showinfo("Успех", "Результаты успешно сохранены в файле Excel.")

def save_unprocessed():
    if unprocessed_results:
        filename = filedialog.asksaveasfilename(defaultextension='.txt', filetypes=[("Text Files", "*.txt")])
        if filename:
            with open(filename, 'w') as file:
                for num in unprocessed_results:
                    file.write(num + '\n')
            messagebox.showinfo("Успех", "Необработанные номера успешно сохранены в файле.")
    else:
        messagebox.showerror("Ошибка", "Нет необработанных номеров для сохранения.")

# Поштучная обработка
def process_number():
    global global_data
    global_data = []

    conn = sqlite3.connect('operator_and_region.db')
    cursor = conn.cursor()

    number = entry.get()

    if len(number) != 11 or int(number[1:4]) <= 899:
        messagebox.showerror("Ошибка", "Номер должен быть 11-значным и начинаться с кода оператора > 899")
        return

    number = re.sub(r'\s+|-|\.|,|\(|\)|\[|\]|"|\'', '', number)
    kod_operatora = number[1:4]
    nomer = number[4:11]

    cursor.execute("SELECT * FROM data WHERE cod=?", (kod_operatora,))
    results = cursor.fetchall()

    found = False

    for result in results:
        ot, do = result[1], result[2]  # Извлечение значений "От" и "До" из результата

        if int(ot) <= int(nomer) <= int(do):
            operator, region = result[3], result[4]
            global_data.append((number, operator, region))
            found = True  # Устанавливаем флаг, что оператор был найден
            break

    if not found:
        messagebox.showerror("Ошибка", "Нет данных о номере")
        return

    result_text = ""
    for item in global_data:
        result_text += "{}\n{}\n{}\n".format(*item)  # Разделение каждого атрибута на новую строку
    result_label.config(text=result_text)


def validate_number(char):
    return char.isdigit()



# Сервис
def show_help():
    help_window = tk.Toplevel()
    help_window.title("Помощь")
    help_window.geometry("600x600")
    help_window.resizable(False, False) 

    help_text = """
    Добро пожаловать в программу "ChekNum"!

    Эта программа предназначена для определения оператора 
    и территориальности абонентских номеров телефонов.

    Чтобы выполнить поштучную обработку номеров
    выполните следующие действия:

        1. Введите номер телефона в поле для ввода.

        2. Нажмите на кнопку "Обработать".

        3. Результат будет предоставлен ниже.

    Чтобы выполнить комплексную обработку номеров
    выполните следующие действия:

        1. Нажмите кнопку "Загрузить файл" и выберите файл формата excel 
           с абонентскими номерами.

        2. Нажмите кнопку "Обработать", чтобы выполнить обработку номеров.
           Результаты будут отображены в новом окне.

        4. При необходимости сохраните нужные данные на ПК с помощью кнопок 
           "Сохранить результаты" и "Сохранить необработанные номера".

    Если у вас возникнут вопросы, обращайтесь к разработчику.

    Приятного использования!
    """
    
    help_label = tk.Label(help_window, text=help_text, justify="left", font=("Arial", 12))
    help_label.pack()

def close_window():
    root.destroy()





# Главное окно
root = tk.Tk()
root.title("ChekNum")
root.iconbitmap(default="phone.ico")
root.geometry("300x600")
root.resizable(False, False) 

# Подключаем модуль ttk (Themed Tkinter) 
style = ttk.Style()
style.theme_use('clam')  # Темы: 'clam', 'alt', 'default', 'classic', 'vista', 'xpnative'
style.configure('TButton', padding=6, relief="ridge", background="#ccc", foreground="#333", font=("Helvetica", 12))
style.configure("Title.TLabel", background="#f0f0f0", foreground="#333333", font=("Helvetica", 14, "bold"))
pady = 10 # Задаем отступы для всех кнопок сразу



# ---------------------------------------------------
separator = ttk.Separator(root, orient='horizontal')
separator.pack(fill='x', pady=pady)
# ---------------------------------------------------


# Поштучная обработка 
single_processing_label = ttk.Label(root, text="Поштучная обработка", style="Title.TLabel")
single_processing_label.pack(pady=5)

validate_cmd = root.register(validate_number)
entry = ttk.Entry(root, validate="key", validatecommand=(validate_cmd, '%S'))
entry.pack(pady=pady)

process_button = ttk.Button(root, text="Обработать", command=process_number)
process_button.pack(pady=pady)

result_label = tk.Label(root, text=" ")
result_label.pack(pady=pady)


# ---------------------------------------------------
separator = ttk.Separator(root, orient='horizontal')
separator.pack(fill='x', pady=pady)
# ---------------------------------------------------


# Массовая обработка

bulk_processing_label = ttk.Label(root, text="Комплексная обработка", style="Title.TLabel")
bulk_processing_label.pack(pady=5)

button1 = ttk.Button(root, text="Загрузить файл", command=read_file_numbers)
button1.pack(pady=pady)

file_label = tk.Label(root, text="Файл для обработки - не выбран !")
file_label.pack(pady=pady)

button2 = ttk.Button(root, text="Обработать", command=chek_numbers)
button2.pack(pady=pady)


# ---------------------------------------------------
separator = ttk.Separator(root, orient='horizontal')
separator.pack(fill='x', pady=pady)
# ---------------------------------------------------


# Сервис
button_help = ttk.Button(root, text="Помощь", command=show_help)
button_help.pack(pady=pady)

close_button = ttk.Button(root, text="Выход", command=close_window)
close_button.pack(pady=pady)



# Запуск проекта
root.mainloop()
