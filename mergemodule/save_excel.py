import os

import pandas as pd


class SaveExcel:
    def __init__(self, path, file_name, sheet_name: list, arr: list):
        """
        
        :param path: 文件夹名字
        :param file_name: 文件名
        :param sheet_name:
        :param arr:
        """
        self.path = path
        self.file_name = file_name
        self.sheet_name = sheet_name
        self.arr = arr
    
    def save_excel(self):
        
        if not os.path.exists(self.path):
            os.makedirs(self.path)
        
        file_path = pd.ExcelWriter(self.path + '/' + f'{self.file_name}.xlsx')
        
        for i in range(2):
            pf = pd.DataFrame(self.arr[i])
            pf.to_excel(file_path, encoding='utf-8', index=False, sheet_name=self.sheet_name[i])
            file_path.save()
