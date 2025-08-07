import sys
from PyQt6.QtWidgets import QApplication, QWidget, QFileDialog

import pandas

from ds_main import Ui_MainForm


class ParsWindow(QWidget, Ui_MainForm):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.pushButton_choose_file.clicked.connect(self.choose_file)
        self.pushButton_ready.clicked.connect(self.run_pars)
        self.lineEdit_inn.returnPressed.connect(self.run_pars)
        self.pushButton_info.clicked.connect(lambda: self.textEdit_output.setText("Эта программа поможет тебе найти информацию о компании по ИНН в твоём файле. \n" \
        "Загрузи файл, воспользовавшись кнопкой с тремя точками, введи ИНН и нажми кнопку 'Готово'. В этом же поле ты увидишь инфо о компании.\n" \
        "Важно: в твоём файле должны присутствовать столбцы 'ИНН', 'Наименование / ФИО', 'Телефон', 'E-mail', 'WWW'\n\n" \
        "Автор приложения: Татьяна Перова"))
        self.pushButton_close.clicked.connect(lambda: self.close())

    def choose_file(self):
        self.file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл для парсинга",  # заголовок окна
            "",                # начальная директория (когда пусто используется текущая папка)
            "*.xlsx *.csv"  # фильтр файлов
        )

        self.lineEdit_path.setText(self.file_path)
        

    def run_pars(self):

        # Проверяем выбран ли файл
        if not hasattr(self, 'file_path') or not self.file_path:
            self.textEdit_output.setText("Нужно сначала выбрать файл!")
            return

        # Проверяем корректный ли ИНН
        try:
            inn = int(self.lineEdit_inn.text())
        except ValueError:
            self.textEdit_output.setText("ИНН может быть только числом!")
            return
        
        try:
            df = detTable(self.file_path)

            df['ИНН'] = df['ИНН'].astype(int)
            row = df[df['ИНН'] == inn]
            if not row.empty:
                lines = []
                columns_for_show = ['Наименование / ФИО', 'Телефон', 'E-mail', 'WWW']
                for column, value in row.iloc[0].items():
                        if column in columns_for_show:
                            lines.append(f"{column}: {value if not pandas.isna(value) else ''}")
                            self.res = "\n".join(lines)
                        else:
                            continue
            else: self.res = f"Организации с ИНН {inn} в файле не найдено."

            self.textEdit_output.setText(self.res)
        except:
            self.textEdit_output.setText("Вероятно ты выбрал файл неподходящей структуры. Чтобы узнать больше нажми кнопку 'О программе' внизу окна.")


def detTable(file_path):

    # Прочитаем Excel целиком, без заголовков
    df_raw = pandas.read_excel(file_path, header=None)

    # Найдём строку с максимальным числом непустых значений (это заголовки)
    header_row = df_raw.notna().sum(axis=1).idxmax()

    # Прочитаем файл снова, используя найденную строку как заголовки
    df = pandas.read_excel(file_path, header=header_row)

    return(df)



app = QApplication(sys.argv) # Создаём экземпляр приложения на базе фреймворка Qt
win = ParsWindow() # Создаём элемент класса Widget
win.show() # Показываем окно на экране
sys.exit(app.exec()) # Запускаем основной цикл обработки событий, в котором приложение будет "жить", пока пользователь его не закроет