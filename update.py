Привет!
Мой старый код:
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QTabWidget, QTableWidget, QTableWidgetItem, QPushButton, QTextEdit, QProgressBar, QFileDialog
from PyQt5.QtCore import Qt
import pandas as pd

class LogWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.console = QTextEdit()
        self.console.setReadOnly(True)
        layout = QVBoxLayout()
        layout.addWidget(self.console)
        self.setLayout(layout)

    def append_log(self, message):
        self.console.append(message)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.p2p_file_name = None
        self.report_file_name = None
        self.account_table_file_name = None
        self.log_window = LogWindow()
        
        self.progress = QProgressBar()
        
        self.notebook = QTabWidget()
        
        upload_p2p_btn = QPushButton('Загрузить файл P2P')
        upload_p2p_btn.clicked.connect(self.upload_p2p_file)
        
        upload_report_btn = QPushButton('Загрузить файл Report_tab')
        upload_report_btn.clicked.connect(self.upload_report_file)
        
        upload_account_table_btn = QPushButton('Загрузить таблицу счетов')
        upload_account_table_btn.clicked.connect(self.upload_account_table)
        
        compare_files_btn = QPushButton('Сравнить файлы')
        compare_files_btn.clicked.connect(self.compare_files)
        
        check_accounts_btn = QPushButton('Проверить счета')
        check_accounts_btn.clicked.connect(self.check_account_numbers)
        
        layout = QVBoxLayout()
        layout.addWidget(upload_p2p_btn)
        layout.addWidget(upload_report_btn)
        layout.addWidget(upload_account_table_btn)
        layout.addWidget(compare_files_btn)
        layout.addWidget(check_accounts_btn)
        
        central_widget = QWidget()
        central_layout = QVBoxLayout()
        central_layout.addLayout(layout)
        central_layout.addWidget(self.notebook)
        central_widget.setLayout(central_layout)
        
        self.setCentralWidget(central_widget)

    def upload_p2p_file(self):
      file_dialog = QFileDialog.getOpenFileName(self, 'Выберите файл P2P', '', 'Excel Files (*.xlsx)')
      if file_dialog[0]:
          self.p2p_file_name = file_dialog[0]
          self.log_window.append_log(f"Загружен файл P2P: {self.p2p_file_name}")
          
          # Отображение загруженной таблицы в редактируемом QTableWidget
          df_p2p = pd.read_excel(self.p2p_file_name)
          rows, cols = df_p2p.shape
          
          table_widget = QTableWidget(rows, cols)  # Создаем QTableWidget с нужным количеством строк и столбцов
          
          for j in range(cols):
              header_item = QTableWidgetItem(df_p2p.columns[j])
              table_widget.setHorizontalHeaderItem(j, header_item)

          for i in range(rows):
              for j in range(cols):
                  item = QTableWidgetItem(str(df_p2p.iloc[i, j]))  # Создаем QTableWidgetItem с данными из DataFrame
                  table_widget.setItem(i, j, item)  # Устанавливаем QTableWidgetItem в ячейку таблицы

          self.notebook.addTab(table_widget, 'P2P (Редактирование)')

    def upload_report_file(self):
      file_dialog = QFileDialog.getOpenFileName(self, 'Выберите файл Report_tab', '', 'Excel Files (*.xlsx)')
      if file_dialog[0]:
          self.report_file_name = file_dialog[0]
          self.log_window.append_log(f"Загружен файл Report_tab: {self.report_file_name}")
          
          # Отображение загруженной таблицы в редактируемом QTableWidget
          df_report = pd.read_excel(self.report_file_name)
          rows, cols = df_report.shape
          
          table_widget = QTableWidget(rows, cols)  # Создаем QTableWidget с нужным количеством строк и столбцов
          
          for j in range(cols):
              header_item = QTableWidgetItem(df_report.columns[j])
              table_widget.setHorizontalHeaderItem(j, header_item)

          for i in range(rows):
              for j in range(cols):
                  item = QTableWidgetItem(str(df_report.iloc[i, j]))  # Создаем QTableWidgetItem с данными из DataFrame
                  table_widget.setItem(i, j, item)  # Устанавливаем QTableWidgetItem в ячейку таблицы

          self.notebook.addTab(table_widget, 'Report_tab (Редактирование)')

    def compare_files(self):
        if self.p2p_file_name and self.report_file_name:
            self.log_window.append_log("Сравнение файлов начато...")

            # Загрузка данных из файлов
            df_p2p = pd.read_excel(self.p2p_file_name)
            df_report = pd.read_excel(self.report_file_name)

            # Сравнение файлов и вывод результатов
            common_columns = 'Номер транзакции'  # Используйте соответствующее название столбца

            df_common = pd.merge(df_p2p, df_report, on=common_columns)
            df_no_p2p = df_p2p[~df_p2p["Номер транзакции"].isin(df_report["Номер транзакции"])]
            df_no_reportid = df_report[~df_report["Номер транзакции"].isin(df_p2p["Номер транзакции"])].copy()
            df_no_reportid["Метод"].fillna("", inplace=True)

            # Создание новых QTableWidget для отображения результатов
            table_common = QTableWidget(df_common.shape[0], df_common.shape[1])
            table_no_p2p = QTableWidget(df_no_p2p.shape[0], df_no_p2p.shape[1])
            table_no_reportid = QTableWidget(df_no_reportid.shape[0], df_no_reportid.shape[1])  # Создание таблицы для вкладки "Отсутствующие в Report_tab"

            # Установка заголовков столбцов
            for j in range(df_common.shape[1]):
                header_item = QTableWidgetItem(df_common.columns[j])
                table_common.setHorizontalHeaderItem(j, header_item)

            for j in range(df_no_p2p.shape[1]):
                header_item = QTableWidgetItem(df_no_p2p.columns[j])
                table_no_p2p.setHorizontalHeaderItem(j, header_item)

            for j in range(df_no_reportid.shape[1]):
                header_item = QTableWidgetItem(df_no_reportid.columns[j])
                table_no_reportid.setHorizontalHeaderItem(j, header_item)

            # Заполнение таблиц данными из DataFrame
            for i in range(df_common.shape[0]):
                for j in range(df_common.shape[1]):
                    item = QTableWidgetItem(str(df_common.iloc[i, j]))
                    table_common.setItem(i, j, item)

            for i in range(df_no_p2p.shape[0]):
                for j in range(df_no_p2p.shape[1]):
                    item = QTableWidgetItem(str(df_no_p2p.iloc[i, j]))
                    table_no_p2p.setItem(i, j, item)

            for i in range(df_no_reportid.shape[0]):  # Заполнение таблицы для вкладки "Отсутствующие в Report_tab"
                for j in range(df_no_reportid.shape[1]):
                    item = QTableWidgetItem(str(df_no_reportid.iloc[i, j]))
                    table_no_reportid.setItem(i, j, item)
            
            # Добавление таблиц в соответствующие вкладки с именами колонок
            self.notebook.addTab(table_common, 'Совпадающие данные')
            self.notebook.addTab(table_no_p2p, 'Отсутствующие в P2P')
            self.notebook.addTab(table_no_reportid, 'Отсутствующие в Report_tab')  # Добавление таблицы в соответствующую вкладку

            unique_values = df_no_reportid.iloc[:, 2].unique()  # Получение уникальных значений из третьей колонки

            # Создание новых вкладок на основе уникальных значений
            for value in unique_values:
                df_value = df_no_reportid[df_no_reportid.iloc[:, 2] == value]
                table_value = QTableWidget(df_value.shape[0], df_value.shape[1])

                # Установка заголовков столбцов
                for j in range(df_value.shape[1]):
                    header_item = QTableWidgetItem(df_value.columns[j])
                    table_value.setHorizontalHeaderItem(j, header_item)

                # Заполнение таблицы данными из DataFrame
                for i in range(df_value.shape[0]):
                    for j in range(df_value.shape[1]):
                        item = QTableWidgetItem(str(df_value.iloc[i, j]))
                        table_value.setItem(i, j, item)

                self.notebook.addTab(table_value, str(value))

            self.log_window.append_log("Сравнение файлов завершено.")
        else:
            self.log_window.append_log("Необходимо загрузить файлы P2P и Report_tab для сравнения.")

    def upload_account_table(self):
      file_dialog = QFileDialog.getOpenFileName(self, 'Выберите таблицу счетов', '', 'Excel Files (*.xlsx)')
      if file_dialog[0]:
          self.account_table_file_name = file_dialog[0]
          self.log_window.append_log(f"Загружена таблица счетов: {self.account_table_file_name}")
          
          # Отображение загруженной таблицы в редактируемом QTableWidget
          df_account_table = pd.read_excel(self.account_table_file_name)
          rows, cols = df_account_table.shape
          
          table_widget = QTableWidget(rows, cols)  # Создаем QTableWidget с нужным количеством строк и столбцов
          
          for j in range(cols):
              header_item = QTableWidgetItem(df_account_table.columns[j])
              table_widget.setHorizontalHeaderItem(j, header_item)

          for i in range(rows):
              for j in range(cols):
                  item = QTableWidgetItem(str(df_account_table.iloc[i, j]))  # Создаем QTableWidgetItem с данными из DataFrame
                  table_widget.setItem(i, j, item)  # Устанавливаем QTableWidgetItem в ячейку таблицы

          self.notebook.addTab(table_widget, 'Таблица счетов (Редактирование)')

    def check_account_numbers(self):
        if self.report_file_name and self.account_table_file_name:
            # Загрузка данных из файла отчета
            df_report = pd.read_excel(self.report_file_name)

            # Загрузка данных из таблицы счетов
            df_account_table = pd.read_excel(self.account_table_file_name)

            # Проверка наличия счетов в отчете
            df_missing_accounts = df_account_table[~df_account_table['Счет'].isin(df_report['Счет'])]

            # Сохраняем найденные счета в отдельном файле
            df_missing_accounts.to_excel('missing_accounts.xlsx', sheet_name='Счета не найдены', index=False)

            self.log_window.append_log("Найденные счета сохранены в файле missing_accounts.xlsx")
        else:
            self.log_window.append_log("Необходимо загрузить файл Report_tab и таблицу счетов для проверки.")

app = QApplication(sys.argv)
window = MainWindow()
window.show()

log_window = LogWindow()
window.notebook.addTab(log_window, 'Лог')

sys.exit(app.exec_())

Мой новый код с исправлениями:
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QTabWidget, QTableWidget, QTableWidgetItem, QPushButton, QTextEdit, QProgressBar, QFileDialog
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import pandas as pd

class LogWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.console = QTextEdit()
        self.console.setReadOnly(True)
        layout = QVBoxLayout()
        layout.addWidget(self.console)
        self.setLayout(layout)

    def append_log(self, message):
        self.console.append(message)

class ComparisonThread(QThread):
    finished = pyqtSignal()

    def __init__(self):
        super().__init__()

    def run(self):
        # Реализуйте функционал сравнения файлов и другие операции здесь
        # После выполнения задачи эмитируйте сигнал finished
        self.finished.emit()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.p2p_file_name = None
        self.report_file_name = None
        self.account_table_file_name = None
        self.log_window = LogWindow()
        
        self.progress = QProgressBar()
        
        self.notebook = QTabWidget()
        
        upload_p2p_btn = QPushButton('Загрузить файл P2P')
        upload_p2p_btn.clicked.connect(self.upload_p2p_file)
        
        upload_report_btn = QPushButton('Загрузить файл Report_tab')
        upload_report_btn.clicked.connect(self.upload_report_file)
        
        upload_account_table_btn = QPushButton('Загрузить таблицу счетов')
        upload_account_table_btn.clicked.connect(self.upload_account_table)
        
        compare_files_btn = QPushButton('Сравнить файлы')
        compare_files_btn.clicked.connect(self.compare_files)
        
        check_accounts_btn = QPushButton('Проверить счета')
        check_accounts_btn.clicked.connect(self.check_account_numbers)
        
        layout = QVBoxLayout()
        layout.addWidget(upload_p2p_btn)
        layout.addWidget(upload_report_btn)
        layout.addWidget(upload_account_table_btn)
        layout.addWidget(compare_files_btn)
        layout.addWidget(check_accounts_btn)
        
        central_widget = QWidget()
        central_layout = QVBoxLayout()
        central_layout.addLayout(layout)
        central_layout.addWidget(self.notebook)
        central_widget.setLayout(central_layout)
        
        self.setCentralWidget(central_widget)

    def upload_p2p_file(self):
      file_dialog = QFileDialog.getOpenFileName(self, 'Выберите файл P2P', '', 'Excel Files (*.xlsx)')
      if file_dialog[0]:
          self.p2p_file_name = file_dialog[0]
          self.log_window.append_log(f"Загружен файл P2P: {self.p2p_file_name}")
          
          # Отображение загруженной таблицы в редактируемом QTableWidget
          df_p2p = pd.read_excel(self.p2p_file_name)
          rows, cols = df_p2p.shape
          
          table_widget = QTableWidget(rows, cols)  # Создаем QTableWidget с нужным количеством строк и столбцов
          
          for j in range(cols):
              header_item = QTableWidgetItem(df_p2p.columns[j])
              table_widget.setHorizontalHeaderItem(j, header_item)

          for i in range(rows):
              for j in range(cols):
                  item = QTableWidgetItem(str(df_p2p.iloc[i, j]))  # Создаем QTableWidgetItem с данными из DataFrame
                  table_widget.setItem(i, j, item)  # Устанавливаем QTableWidgetItem в ячейку таблицы

          self.notebook.addTab(table_widget, 'P2P (Редактирование)')

    def upload_report_file(self):
      file_dialog = QFileDialog.getOpenFileName(self, 'Выберите файл Report_tab', '', 'Excel Files (*.xlsx)')
      if file_dialog[0]:
          self.report_file_name = file_dialog[0]
          self.log_window.append_log(f"Загружен файл Report_tab: {self.report_file_name}")
          
          # Отображение загруженной таблицы в редактируемом QTableWidget
          df_report = pd.read_excel(self.report_file_name)
          rows, cols = df_report.shape
          
          table_widget = QTableWidget(rows, cols)  # Создаем QTableWidget с нужным количеством строк и столбцов
          
          for j in range(cols):
              header_item = QTableWidgetItem(df_report.columns[j])
              table_widget.setHorizontalHeaderItem(j, header_item)

          for i in range(rows):
              for j in range(cols):
                  item = QTableWidgetItem(str(df_report.iloc[i, j]))  # Создаем QTableWidgetItem с данными из DataFrame
                  table_widget.setItem(i, j, item)  # Устанавливаем QTableWidgetItem в ячейку таблицы

          self.notebook.addTab(table_widget, 'Report_tab (Редактирование)')

    def compare_files(self):
        if self.p2p_file_name and self.report_file_name:
            self.log_window.append_log("Сравнение файлов начато...")
            
            # Создаем экземпляр ComparisonThread и подключаем его сигналы к соответствующим слотам
            self.comparison_thread = ComparisonThread()
            self.comparison_thread.finished.connect(self.handle_comparison_finished)
            
            # Запускаем поток
            self.comparison_thread.start()
        else:
            self.log_window.append_log("Необходимо загрузить файлы P2P и Report_tab для сравнения.")

    def handle_comparison_finished(self):
        # Обработчик завершения сравнения файлов
        self.log_window.append_log("Сравнение файлов завершено.")
        
        # Добавьте здесь код для обновления данных в таблицах P2P и Report_tab на основе изменений, внесенных пользователем через QTableWidget
        
        # Повторите цикл сравнения файлов, если необходимо
        
        # Создайте отчетный файл

    def upload_account_table(self):
      file_dialog = QFileDialog.getOpenFileName(self, 'Выберите таблицу счетов', '', 'Excel Files (*.xlsx)')
      if file_dialog[0]:
          self.account_table_file_name = file_dialog[0]
          self.log_window.append_log(f"Загружена таблица счетов: {self.account_table_file_name}")
          
          # Отображение загруженной таблицы в редактируемом QTableWidget
          df_account_table = pd.read_excel(self.account_table_file_name)
          rows, cols = df_account_table.shape
          
          table_widget = QTableWidget(rows, cols)  # Создаем QTableWidget с нужным количеством строк и столбцов
          
          for j in range(cols):
              header_item = QTableWidgetItem(df_account_table.columns[j])
              table_widget.setHorizontalHeaderItem(j, header_item)

          for i in range(rows):
              for j in range(cols):
                  item = QTableWidgetItem(str(df_account_table.iloc[i, j]))  # Создаем QTableWidgetItem с данными из DataFrame
                  table_widget.setItem(i, j, item)  # Устанавливаем QTableWidgetItem в ячейку таблицы

          self.notebook.addTab(table_widget, 'Таблица счетов (Редактирование)')

    def check_account_numbers(self):
        if self.report_file_name and self.account_table_file_name:
            # Загрузка данных из файла отчета
            df_report = pd.read_excel(self.report_file_name)
            
            # Загрузка данных из таблицы счетов
            df_account_table = pd.read_excel(self.account_table_file_name)
            
            # Проверка наличия счетов в отчете
            df_missing_accounts = df_account_table[~df_account_table['Счет'].isin(df_report['Счет'])]
            
            # Сохраняем найденные счета в отдельном файле
            df_missing_accounts.to_excel('missing_accounts.xlsx', sheet_name='Счета не найдены', index=False)
            
            self.log_window.append_log("Найденные счета сохранены в файле missing_accounts.xlsx")
        else:
            self.log_window.append_log("Необходимо загрузить файл Report_tab и таблицу счетов для проверки.")

    def closeEvent(self, event):
        # Переопределяем метод closeEvent для корректного завершения потока перед закрытием приложения
        self.comparison_thread.quit()
        self.comparison_thread.wait()
        event.accept()

app = QApplication(sys.argv)
window = MainWindow()
window.show()

log_window = LogWindow()
window.notebook.addTab(log_window, 'Лог')

sys.exit(app.exec_())

Помоги мне пожалуйста добавить новый код в старый. Напиши соедененный готовый код программы пожалуйста.
