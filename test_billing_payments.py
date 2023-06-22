import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog


class BillingApp:
    def init(self, root):
        self.root = root
        self.root.title("Билинговый инструмент")

        # Создание вкладок
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        # Вкладки для каждого файла Excel
        self.tab1 = ttk.Frame(self.notebook)
        self.tab2 = ttk.Frame(self.notebook)
        self.tab3 = ttk.Frame(self.notebook)

        self.notebook.add(self.tab1, text="Файл Excel 1")
        self.notebook.add(self.tab2, text="Файл Excel 2")
        self.notebook.add(self.tab3, text="Файл Excel 3")

        # Кнопки для загрузки файлов
        self.load_button1 = tk.Button(self.tab1, text="Загрузить файл Excel 1", command=lambda: self.load_file(1))
        self.load_button1.pack(pady=10)

        self.load_button2 = tk.Button(self.tab2, text="Загрузить файл Excel 2", command=lambda: self.load_file(2))
        self.load_button2.pack(pady=10)

        self.load_button3 = tk.Button(self.tab3, text="Загрузить файл Excel 3", command=lambda: self.load_file(3))
        self.load_button3.pack(pady=10)

        # Treeviews для отображения данных таблиц
        self.tree1 = self.create_treeview(self.tab1)
        self.tree2 = self.create_treeview(self.tab2)
        self.tree3 = self.create_treeview(self.tab3)

    def create_treeview(self, parent):
        tree = ttk.Treeview(parent, show="headings")
        tree.pack(pady=20, padx=20, fill=tk.BOTH, expand=True)
        tree.bind("<Double-1>", self.on_item_double_click)

        # Добавление стилей для выравнивания по центру
        style = ttk.Style()
        style.configure("Treeview.Heading", align="center")
        style.configure("Treeview", rowheight=30, font=('Helvetica', 10))

        return tree

    def on_item_double_click(self, event):
        tree = event.widget
        item = tree.selection()[0]
        col = tree.identify_column(event.x)
        col = col.split("#")[-1]
        old_value = tree.item(item, 'values')[int(col) - 1]
        new_value = simpledialog.askstring("Редактирование", f"Изменить значение:", initialvalue=old_value)

        if new_value is not None:
            values = list(tree.item(item, 'values'))
            values[int(col) - 1] = new_value
            tree.item(item, values=values)

    def load_file(self, file_number):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        tree = [self.tree1, self.tree2, self.tree3][file_number - 1]

        if file_path:
            try:
                # Загрузка данных из Excel файла
                data = pd.read_excel(file_path, engine='openpyxl')

                # Очистка старых данных в Treeview
                for col in tree["columns"]:
                    tree.heading(col, text="")
                    tree.column(col, width=0)
                    tree.delete(col)
                tree["columns"] = []
                tree.delete(*tree.get_children())

                # Добавление новых данных в Treeview
                columns = data.columns.tolist()
                tree["columns"] = columns
                for col in columns:
                    tree.column(col, width=100, anchor="center")
                    tree.heading(col, text=col, anchor="center")

                for index, row in data.iterrows():
                    tree.insert("", tk.END, values=list(row))

                messagebox.showinfo("Успех", f"Файл {file_number} успешно загружен!")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при загрузке файла {file_number}: {e}")


if name == "main":
    root = tk.Tk()
    app = BillingApp(root)
    root.mainloop()
