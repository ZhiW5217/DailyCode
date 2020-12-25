import pandas as pd
import os


class ExcelMerge:
    def __init__(self):
        self.path = 'excels'
        self.orders = []
        self.pp_eval = []
    
    def merge(self):
        # 获取excels 下所有的文件
        for root_dir, sub_dir, files in os.walk(self.path):  # root_dir 绝对路径 sub_dir 相对路径 files 文件列表
            for file in files:
                file_name = self.path + '/' + file  # 拼接文件路径
                
                # 读取Excel
                xlsx = pd.ExcelFile(file_name, engine='openpyxl')
                df = pd.read_excel(xlsx, sheet_name=["刷单", "PP测评"], header=1)
                
                #  删除空白行  axis = 0 以行为轴  axis = 1 以列为轴
                order_df = df["刷单"].dropna(axis=0, how='all')
                pp_df = df["PP测评"].dropna(axis=0, how='all')
                
                # 遍历索引
                for i in order_df.index:
                    #  获取索引的内容
                    data = order_df.loc[i]
                    
                    # 转字典
                    data = data.to_dict()
                    
                    # 存入列表
                    self.orders.append(data)
                
                for j in pp_df.index:
                    data = pp_df.loc[j]
                    data = data.to_dict()
                    self.pp_eval.append(data)
        return [self.orders, self.pp_eval]
# if __name__ == '__main__':
#     es = ExcelMerge()
#     arr = es.merge()
#     print(arr)
