"""
GUI程序，同时也是主逻辑模块。
"""
import datetime
import sys
from PySide6.QtWidgets import QApplication, QMainWindow
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QFile, QIODevice
from GetData import GetData
from MarkQualifiedRows import MarkQualifiedRows
from CalculateSuccessRate import CalculateSuccessRate


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        ui_file_name = "stock.ui"
        ui_file = QFile(ui_file_name)
        if not ui_file.open(QIODevice.ReadOnly):
            print(f"Cannot open {ui_file_name}: {ui_file.errorString()}")
            sys.exit(-1)
        loader = QUiLoader()
        self.ui = loader.load(ui_file)
        ui_file.close()
        if not self.ui:
            print(loader.errorString())
            sys.exit(-1)

        # Set window title
        self.ui.setWindowTitle("Stock Data Viewer")
        # fix window size
        size = self.ui.size()
        w, h = size.width(), size.height()
        self.ui.setFixedSize(w, h)

        # 一些需要初始化指定值的ui元素
        end= datetime.datetime.today()
        beg = end - datetime.timedelta(days=365 * 2)
        self.ui.date_beg.setDate(beg)
        self.ui.date_end.setDate(end)

        # Connect button click signal to slot
        self.ui.btn_start.clicked.connect(self.handle_btn_start_clicked)
        self.ui.btn_firstlogic.clicked.connect(self.handle_btn_firstlogic_clicked)

        # 主逻辑涉及到的跨类变量
        self.getdata_output_filename = ''
        self.firstlogic_output_filename = ''

    @staticmethod
    def date_to_string(dt):
        """

        :param dt: Qdate
        :return: string like 20220317
        """
        dt_year, dt_month, dt_day = dt.year(), dt.month(), dt.day()
        dt_year = str(dt_year)
        if dt_month < 10:
            dt_month = "0" + str(dt_month)
        else:
            dt_month = str(dt_month)
        if dt_day < 10:
            dt_day = "0" + str(dt_day)
        else:
            dt_day = str(dt_day)
        dt = dt_year + dt_month + dt_day
        return dt

    def handle_btn_start_clicked(self):
        getdata_instance = GetData()

        # read data from ui file
        stock_code = self.ui.lineEdit_stock_code.text()
        date_beg = self.ui.date_beg.date()
        date_end = self.ui.date_end.date()
        date_beg = self.date_to_string(date_beg)
        date_end = self.date_to_string(date_end)

        getdata_instance.set_attris(stock_code, date_beg, date_end)
        getdata_instance.get_data()
        getdata_instance.save_data_to_file()
        result_prompt = getdata_instance.get_result_prompt()
        self.ui.text_browser.setText(result_prompt)
        self.getdata_output_filename = getdata_instance.get_output_filename()

    def handle_btn_firstlogic_clicked(self):
        # --- 筛选并标记可投资的row
        markrows = MarkQualifiedRows()
        markrows.run(filename=self.getdata_output_filename)
        self.firstlogic_output_filename = markrows.get_output_filename()
        prompt = markrows.get_prompt()
        self.ui.text_browser.append(prompt)

        # --- 计算成功率
        target = self.ui.spinBox_target.value()
        days_hold = self.ui.spinBox_dayshold.value()
        cal = CalculateSuccessRate(target=target, days_hold=days_hold)
        cal.run(filename=self.firstlogic_output_filename)
        prompt = cal.get_prompt()
        self.ui.text_browser.append(prompt)

    def closeEvent(self, event):
        print("closeEvent!!!!!")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.ui.show()
    sys.exit(app.exec())
