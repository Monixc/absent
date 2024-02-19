import pandas as pd
import tkinter as tk
from tkinter import messagebox

def save_to_csv():
    try:
        filename = file_name_entry.get()
        columns = [int(i) for i in column_number_entry.get().split(',')]

        df = pd.read_excel('출석.xlsx', header=None)

        df = df.loc[4:318] 

        absent_students = {f'출석{i+1}': df[df[column] == "."][0] for i, column in enumerate(columns)}

        absent_students_df = pd.DataFrame(absent_students)
        absent_students_df.to_csv(filename, index=False)

        messagebox.showinfo("완료", "CSV 파일이 성공적으로 생성되었습니다.")
    except Exception as e:

        messagebox.showerror("오류", str(e))

root = tk.Tk()


file_name_label = tk.Label(root, text="파일명:")
file_name_label.pack()
file_name_entry = tk.Entry(root)
file_name_entry.pack()

column_number_label = tk.Label(root, text="열 번호 (콤마로 구분. 예: 1,2,3):")
column_number_label.pack()
column_number_entry = tk.Entry(root)
column_number_entry.pack()

save_button = tk.Button(root, text="CSV 파일 생성", command=save_to_csv)
save_button.pack()

root.mainloop()
