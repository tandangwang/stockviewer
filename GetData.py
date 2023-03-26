"""
本类不能直接运行，需要在Window.py中以实际运行。
功能：
获取指定股票代码、指定开始、结束日期间的日K数据。
"""
# 导入 efinance 库
import datetime

import efinance as ef


class GetData:
    def __init__(self):
        self.stock_code = None
        self.end = None
        self.beg = None
        self.df = None
        self.n_rows = 0
        self.n_cols = 0
        self.result_prompt = "null"
        self.output_filename = ''
        self.save_file_prompt = ''

    def get_begin_end_date(self):
        return self.beg, self.end

    def set_attris(self, stock_code, beg, end):
        self.stock_code = stock_code
        self.beg = beg
        self.end = end
        self.df = None
        self.n_rows = 0
        self.n_cols = 0

    def get_data(self):
        self.df = ef.stock.get_quote_history(self.stock_code, beg=self.beg, end=self.end)
        self.n_rows, self.n_cols = self.df.shape

    def get_result_prompt(self):
        if self.df is None:
            prompt = "尚未获取数据"
            return None
        now = datetime.datetime.now()
        prompt = str(now) + "\n"
        prompt = prompt + "本次共获取 " + str(self.n_rows) + " 行 " + str(self.n_cols) + " 列"
        prompt += self.save_file_prompt
        return prompt

    def get_num_rows_cols(self):
        return self.n_rows, self.n_cols

    def save_data_to_file(self):
        self.output_filename = fr"result_{self.stock_code}_{self.beg}_{self.end}.xlsx"
        self.df.to_excel(self.output_filename)
        self.save_file_prompt += "\n" + f"{self.output_filename} 已保存至本地"

    def get_output_filename(self):
        return self.output_filename
