import pyodbc
import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from PIL import Image, ImageTk

class App(ctk.CTk):
    conn = pyodbc.connect('Driver={SQL Server};Server=XTREMELITE-PC;Database=test;UID=;PWD=')
    cursor = conn.cursor()
    def __init__(self):
        super().__init__()

        #fonts
        self.font1 = ctk.CTkFont(size=20)
        self.font2 = ctk.CTkFont(size=30)
        self.font3 = ctk.CTkFont(size=15)
        #buttons
        self.users_button = ctk.CTkButton(self,
                                          text='Выбор таблицы\n"Пользователи"', command=self.users_data, corner_radius=60, font=self.font2
                                          )
        self.users_button.grid(row=1, column=0, sticky="n")

        self.product_button = ctk.CTkButton(self,
                                            text='Выбор таблицы\n"Продукты"', command=self.product_data, corner_radius=60, font=self.font2
                                            )
        self.product_button.grid(row=2, column=0, sticky="n")
        #gridSettings
        self.grid_columnconfigure((0,1,2), weight=1)
        self.grid_rowconfigure((0,1,2), weight=2)
        self.grid_rowconfigure((3,7), weight=1)

        self.product_data()
    #methods
    def sort_column(self, treeview, col, reverse):
        def convert(value):
            if value.isdigit():
                return int(value)
            elif value.replace('.', '', 1).isdigit():
                return float(value)
            elif value.lower() == 'true':
                return 1
            elif value.lower() == 'false':
                return 0
            else:
                return value

        data = [(convert(treeview.set(child, col)), child) for child in treeview.get_children('')]
        data.sort(key=lambda x: x[0], reverse=reverse)
        for index, item in enumerate(data):
            treeview.move(item[1], '', index)
        treeview.heading(col, command=lambda: self.sort_column(treeview, col, not reverse))

    def get_img(self):
        try:
            self.cursor.execute('SELECT Image FROM Лист1$ WHERE ID = ?', self.selected_id)
            img_row = self.cursor.fetchone()

            if img_row:
                self.image_filename = img_row[0]
                return self.image_filename

            self.conn.commit()

        except Exception as e:
            print(e)
            messagebox.showerror("Ошибка", f"{e}")

    def show_image_in_tkinter(self):
        try:
            self.image_filename = self.get_img()
            if self.image_filename:
                width = 200
                height = 200
                image = Image.open(f"images\\{self.image_filename}")
                image = image.resize((width, height))
                self.tk_image = ImageTk.PhotoImage(image)
                self.img_label = tk.Label(self,  image=self.tk_image)
                self.img_label.grid(row=7, column=2, padx=10, sticky="wn")
            else:
                return

        except Exception as e:
            print(e)
            messagebox.showerror("Ошибка", f"{e}")

    def users_data(self):
        try:
            self.users_data_active = True
            self.cursor.execute("select * from UBusers$")
            users_bd = self.cursor.fetchall()

            if hasattr(self, 'product_ttk_list'):
                self.product_ttk_list.destroy()
            columns = ("ID", "ФИО", "Место работы", "Зарплата")
            self.users_ttk_list = ttk.Treeview(self)
            self.users_ttk_list.grid(row=0, column=1, columnspan=2, rowspan=3, sticky="nsew", padx=10)
            self.users_ttk_list.heading("#0", text="")
            self.users_ttk_list.column("#0", width=0, stretch=tk.NO)
            self.users_ttk_list["columns"] = columns

            self.bottom_plug = ctk.CTkLabel(self, fg_color="grey", height=200, width=200, text="")
            self.bottom_plug.grid(row=3, column=1, pady=5, padx=10, rowspan=5, columnspan=2)

            for col in columns:
                self.users_ttk_list.heading(col, text=col)
                self.users_ttk_list.heading(col, command=lambda c=col: self.sort_column(self.users_ttk_list, c, False))

            for item in users_bd:
                row_id = item[0]
                value = (row_id, *item[1:])
                self.users_ttk_list.insert("", "end", values=value)

            self.conn.commit()

        except Exception as e:
            print(e)
            messagebox.showerror("Ошибка", f"{e}")

        finally:
            self.users_data_active = False
            self.bottom_plug.grid_remove()

    def product_data(self):
        try:
            self.product_data_active = True
            self.cursor.execute('SELECT * FROM Лист1$')
            self.product_bd = self.cursor.fetchall()
            if hasattr(self, 'users_ttk_list'):
                self.users_ttk_list.destroy()

            columns = (
                "ID", "Артикул", "Наименование", "Ед. Измерения", "Стоимость", "Макс. размер скидки", "Производитель",
                "Поставщик", "Категория товара", "Скидка", "Кол-во", "Описание")

            self.product_ttk_list = ttk.Treeview(self)
            self.product_ttk_list.grid(row=0, column=1,columnspan=2, rowspan=3, sticky="nsew", padx=10)
            self.product_ttk_list.heading("#0", text="")
            self.product_ttk_list.column("#0", width=0, stretch=tk.NO)
            self.product_ttk_list["columns"] = columns
            self.product_ttk_list.column("ID", width=2)
            self.product_ttk_list.column("Артикул", width=60)
            self.product_ttk_list.column("Наименование", width=120)
            self.product_ttk_list.column("Ед. Измерения", width=100)
            self.product_ttk_list.column("Стоимость", width=70)
            self.product_ttk_list.column("Макс. размер скидки", width=130)
            self.product_ttk_list.column("Производитель", width=100)
            self.product_ttk_list.column("Поставщик", width=80)
            self.product_ttk_list.column("Категория товара", width=110)
            self.product_ttk_list.column("Скидка", width=60)
            self.product_ttk_list.column("Кол-во", width=60)
            self.product_ttk_list.column("Описание", width=250)
            self.product_ttk_list.bind("<<TreeviewSelect>>", self.inf_labels)

            for col in columns:
                self.product_ttk_list.heading(col, text=col)
                self.product_ttk_list.heading(col, command=lambda c=col: self.sort_column(self.product_ttk_list, c, False))

            for self.row in self.product_bd:
                self.row_id = self.row[0]
                self.value = (self.row_id, *self.row[1:])
                self.product_ttk_list.insert("", "end", values=self.value)

            self.conn.commit()
        except Exception as e:
            print(e)
            messagebox.showerror("Ошибка", f"{e}")

        finally:
            self.product_data_active = False

    def inf_labels(self, event):
        try:
            self.inf_labels_active = True
            self.selected_item = self.product_ttk_list.focus()
            self.item = self.product_ttk_list.item(self.selected_item)
            self.description = self.item['values'][11]
            self.article = self.item['values'][1]
            self.name = self.item['values'][2]

            if hasattr(self, 'description_label'):
                self.description_label.destroy()

            if hasattr(self, 'article_label'):
                self.article_label.destroy()

            if hasattr(self, 'name_label'):
                self.name_label.destroy()

            if hasattr(self, 'img_label'):
                self.img_label.destroy()

            self.name_label = ctk.CTkLabel(self, text=self.name, wraplength=200, font=self.font3, fg_color="grey", corner_radius=90)
            self.name_label.grid(row=3, column=1, padx=10, sticky="w")

            self.description_label = ctk.CTkLabel(self, text=self.description, wraplength=1000, font=self.font3, fg_color="grey", corner_radius=90)
            self.description_label.grid(row=4, column=1, padx=10, sticky="w", columnspan=2)

            self.article_label = ctk.CTkLabel(self, text=self.article, font=self.font3, fg_color="grey", corner_radius=90)
            self.article_label.grid(row=3, column=1, pady=5, padx=10)

            self.selected_id = self.item['values'][0]
            self.show_image_in_tkinter()

        except Exception as e:
            print(e)
            messagebox.showerror("Ошибка", f"{e}")

        finally:
            self.inf_labels_active = False



if __name__ == '__main__':
    app = App()
    app.title("Product App")
    app.geometry("1600x800")
    app.mainloop()