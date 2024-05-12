import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import pandas as pd
import sqlite3

phone_num = []
complite_num = 0
error_num = 0
global_data = []
unprocessed_results = []

def read_file_numbers():
    
    filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")])
    
    global phone_num
    global error_num
    global unprocessed_results
    global complite_num

    phone_num = []       #-----------#
    data_set = []        #--- clear--#
    error_num = 0        #----list---#
    complite_num = 0     #-----------#
    
    if filename.endswith('.xls') or filename.endswith('.xlsx'):
        df = pd.read_excel(filename, dtype=str, header=None)
        if not df.empty:
            # Путь к файлу для отображения
            file_label.config(text="Выбранный файл:\n" + filename)  
            
            data_set.extend(df.iloc[:, 0].tolist())
            for num in data_set:
                if len(str(num)) == 10 and int(num[0:3]) > 899:
                    num = '7' + num
                    complite_num += 1
                    phone_num.append(num)
                elif len(str(num)) == 11 and int(num[1:4]) > 899:
                    complite_num += 1
                    phone_num.append(num)
                else:
                    unprocessed_results.append(num)
                    error_num += 1
    else:
        messagebox.showerror("Ошибка", "Неподдерживаемый формат файла.")

    print("Абонентские номера из файла:", phone_num)
    print('Правильно указанных номеров:', complite_num)
    print('Не правильно указанных номеров:', error_num)

def chek_numbers():
    global phone_num
    global global_data
    global unprocessed_results
    global error_num

    matching_numbers = []

    conn = sqlite3.connect('operator_and_region.db')
    cursor = conn.cursor()

    for num in phone_num:
        kod_operatora = num[1:4]
        nomer =  num[4:11]

        cursor.execute("SELECT * FROM data WHERE cod=?", (kod_operatora,))
        results = cursor.fetchall()

        for result in results:
            ot, do = result[1], result[2]  # Извлечение значений "От" и "До" из результата

            if int(ot) <= int(nomer) <= int(do):
                operator, region = result[3], result[4]
                global_data.append((num, operator, region))
                matching_numbers.append((num, operator, region))

    # Запуск окна с результатами
    if matching_numbers:
        open_results_window(matching_numbers)
    else:
        messagebox.showinfo("Результаты", "Нет подходящих номеров для обработки")
    
    # Закрываем соединение с базой данных
    conn.close()

def open_results_window(results):
    # Создание нового окна
    results_window = tk.Toplevel()
    results_window.title("Результаты")
    results_window.geometry("525x300")

    label_processed = tk.Label(results_window, text=f"Успешно обработано: {complite_num}")
    label_processed.pack()

    label_unprocessed = tk.Label(results_window, text=f"Необработанные: {error_num}")
    label_unprocessed.pack()

    download_button = tk.Button(results_window, text="Скачать обработанные номера", command=save_to_excel)
    download_button.pack()

    download_error_button = tk.Button(results_window, text="Скачать НЕ обработанные номера", command=save_unprocessed)
    download_error_button.pack()

    # Создание текстового виджета и запись результатов
    results_text = tk.Text(results_window)
    results_text.pack(expand=True, fill="both")
    for result in results:
        results_text.insert(tk.END, f"{result}\n")

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



# Создание главного окна
root = tk.Tk()
root.title("ChekNum")
root.iconbitmap(default="phone.ico")
root.geometry("300x250")

button = tk.Button(root, text="Загрузить файл", command=read_file_numbers)
button.pack()

file_label = tk.Label(root, text="Файл для обработки не выбран!!!")
file_label.pack()

button = tk.Button(root, text="Обработать", command=chek_numbers)
button.pack()

root.mainloop()
