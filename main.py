### main.py
# -*- coding: cp1251 -*-
import os
from tkinter import Tk, filedialog, messagebox
from metadata import process_files

def select_files():
    
    ini_file = filedialog.askopenfilename(
        title="�������� .ini ����",
        filetypes=[("INI �����", "*.ini")]
    )
    if ini_file:
        directory = filedialog.askdirectory(title="����� � ������� �� ���������")
        if directory:
            try:
                processed, failed = process_files(directory, ini_file)

                
                success_files = "\n".join(processed)
                failed_files = "\n".join(failed)
                
                message = (
                    f"��������� ���������\n\n"
                    f"������� ������������ �����:\n{success_files if processed else '��� ������������ ������.'}\n\n"
                    f"�� ������� ����������:\n{failed_files if failed else '��� ����� ���������� �������.'}"
                )

                messagebox.showinfo("���������� ���������", message)
            except Exception as e:
                messagebox.showerror("������", f"��������� ������: {e}")
        else:
            messagebox.showwarning("��������", "����� �� �������.")
    else:
        messagebox.showwarning("��������", "INI ���� �� ������.")

if __name__ == "__main__":
    root = Tk()
    root.withdraw()  # ������ ������� ����
    select_files()