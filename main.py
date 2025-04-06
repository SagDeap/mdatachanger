### main.py
# -*- coding: cp1251 -*-
import os
from tkinter import Tk, filedialog, messagebox
from metadata import process_files

def select_files():
    
    ini_file = filedialog.askopenfilename(
        title="Выберите .ini файл",
        filetypes=[("INI файлы", "*.ini")]
    )
    if ini_file:
        directory = filedialog.askdirectory(title="Папка с файлами на обработку")
        if directory:
            try:
                processed, failed = process_files(directory, ini_file)

                
                success_files = "\n".join(processed)
                failed_files = "\n".join(failed)
                
                message = (
                    f"Обработка завершена\n\n"
                    f"Успешно обработанные файлы:\n{success_files if processed else 'Нет обработанных файлов.'}\n\n"
                    f"Не удалось обработать:\n{failed_files if failed else 'Все файлы обработаны успешно.'}"
                )

                messagebox.showinfo("Результаты обработки", message)
            except Exception as e:
                messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")
        else:
            messagebox.showwarning("Внимание", "Папка не выбрана.")
    else:
        messagebox.showwarning("Внимание", "INI файл не выбран.")

if __name__ == "__main__":
    root = Tk()
    root.withdraw()  # Скрыть главное окно
    select_files()