import pandas as pd
import os


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


class ExcelMerge:
    def __init__(self):
        self.path = 'excels'
        self.orders = []
        self.pp_eval = []
    
    def merge(self):
        # 获取excels 下所有的文件
        if not os.path.exists(self.path):
            raise Exception
        for root_dir, sub_dir, files in os.walk(self.path):  # root_dir 绝对路径 sub_dir 相对路径 files 文件列表
            print(files)
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


class DataFilter:
    def __init__(self):
        
        self.xlsx1 = pd.ExcelFile('销售订单总历史记录.xlsx', engine='openpyxl')
        self.xlsx2 = pd.ExcelFile('merge/推广费报表(合并).xlsx', engine='openpyxl')
        self.df1 = pd.read_excel(self.xlsx1, usecols='D,F')
        self.df2 = pd.read_excel(self.xlsx2, sheet_name=["刷单", "PP测评"])
        self.orders = [i.replace(" ", "") for i in
                       self.df1["采购订单/支票编号"].astype(str) + self.df1["AIO Account"].astype(str)]
    
    # 刷单筛选逻辑
    def orders_screen(self):
        """
        刷单表格订单+店铺数据提取
        格式:028-0317077-4205948Foot-Care-DE
        :return:
        """
        order_arr = []
        order_in_arr = []
        for i in self.df2["刷单"].index:
            item = str(self.df2["刷单"]['订单号'].loc[i]) + str(self.df2["刷单"]["店铺（与系统要一致）"].loc[i]).replace(
                " ", "")
            
            if len(item) >= 23 and item not in self.orders:
                data = self.df2["刷单"].loc[i]
                data = data.to_dict()
                print(data)
                if type(data['订单生成日期']) is str:
                    
                    data['订单生成日期'] = None
                order_arr.append(data)
            else:
                data = self.df2["刷单"].loc[i]
                data = data.to_dict()
                if type(data['订单生成日期']) is str:
                    data['订单生成日期'] = None
                order_in_arr.append(data)
        return order_arr, order_in_arr
    
    def pp_evaluation(self):
        """
        pp评测订单+店铺数据提取
        格式:028-0317077-4205948Foot-Care-DE
        :return:
        """
        pp_arr = []  # 匹配失败
        pp_in_arr = []  # 匹配成功
        for i in self.df2["PP测评"].index:
            item = self.df2["PP测评"]["订单号"].loc[i] + self.df2["PP测评"]["店铺（与系统要一致）"].loc[i]
            if len(item) >= 23 and item not in self.orders:
                data = self.df2["PP测评"].loc[i]
                data = data.to_dict()
                pp_arr.append(data)
            else:
                data = self.df2["PP测评"].loc[i]
                data = data.to_dict()
                pp_in_arr.append(data)
        return pp_arr, pp_in_arr
    
    def run_filter(self):
        order_arr, order_in_arr = self.orders_screen()
        pp_arr, pp_in_arr = self.pp_evaluation()
        arr = [order_arr, pp_arr]
        in_arr = [order_in_arr, pp_in_arr]
        return arr, in_arr


if __name__ == '__main__':
    sheet_name = ["刷单", "PP测评"]  # 固定表格名
    while True:
        print("*" * 4 + "运行程序前请先删除历史生成文件夹\'filter\'和\'merge\'" + "*" * 4, "\n")
        print("*" * 6 + "确保待处理文件放入\'excels\'文件夹中" + "*" * 6, "\n")
        if input('输入回车继续程序:') != '':
            break
        try:
            excel_merge = ExcelMerge()
            print("*" * 14 + "读取数据中" + "*" * 14, "\n")
            merge_arr = excel_merge.merge()
            print("*" * 14 + "正在进行合并" + "*" * 14, "\n")
            save_merge = SaveExcel('merge', "推广费报表(合并)", sheet_name, merge_arr)
            save_merge.save_excel()
            print("*" * 16 + "合并成功" + "*" * 16, "\n")
        except Exception as e:
            print("ERROR:文件读取失败！请检查excels文件夹文件格式是否正确！", "\n")
            print(e)
            if input('输入任意键退出程序:') != 'asdadada2df':
                break
        
        try:
            print("*" * 12 + "正在读取历史记录" + "*" * 12, "\n")
            print("*" * 10 + "数据量较大，请耐心等待！" + "*" * 10, "\n")
            filter = DataFilter()
            print("*" * 16 + "读取成功" + "*" * 16, "\n")
            print("*" * 14 + "正在进行数据筛选！" + "*" * 14, "\n")
            arr, in_arr = filter.run_filter()
            print("*" * 16 + '筛选成功' + "*" * 16, "\n")
            print("*" * 16 + "正在保存" + "*" * 16, "\n")
            
            filter_arr = SaveExcel('filter', '匹配失败', sheet_name, arr)
            filter_arr.save_excel()
            filter_in_arr = SaveExcel('filter', '匹配成功', sheet_name, in_arr)
            filter_in_arr.save_excel()
            
            print("*" * 6 + f"数据成功保存到filter文件夹中，请查看" + "*" * 6, "\n")
            if input('输入任意键退出程序:') != 'asdadada2df':
                break
        except Exception as e:
            print("ERROR:请检查销售订单总历史记录.xlsx是否存在", "\n")
            print(e)
        if input('输入任意键退出程序:') != 'asdadada2df':
            break
