import pandas as pd


class DataFilter:
    def __init__(self):
        
        self.xlsx1 = pd.ExcelFile('销售订单总历史记录.xlsx', engine='openpyxl')
        self.xlsx2 = pd.ExcelFile('merge/推广费报表(合并).xlsx', engine='openpyxl')
        self.df1 = pd.read_excel(self.xlsx1, usecols='D,F')
        self.df2 = pd.read_excel(self.xlsx2, sheet_name=["刷单", "PP测评"])
        # todo 去除for 直接合并
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
