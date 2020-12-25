from mergemodule.save_excel import SaveExcel
from mergemodule.data_filter import DataFilter
from mergemodule.excel_merge import ExcelMerge

if __name__ == '__main__':
    sheet_name = ["刷单", "PP测评"]  # 固定表格名
    while True:
        print("*" * 4 + "运行程序前请先删除历史生成文件夹\'filter\'和\'merge\'" + "*" * 4, "\n")
        print("*" * 6 + "确保待处理文件放入\'excels\'文件夹中" + "*" * 6, "\n")
        if input('输入y(小写)继续执行程序,其他键退出程序:') != 'y':
            break
        excel_merge = ExcelMerge()
        print("*" * 14 + "读取数据中" + "*" * 14, "\n")
        
        print("*" * 14 + "正在进行合并" + "*" * 14, "\n")
        try:
            merge_arr = excel_merge.merge()
            save_merge = SaveExcel('merge', "推广费报表(合并)", sheet_name, merge_arr)
            save_merge.save_excel()
            print("*" * 16 + "合并成功" + "*" * 16, "\n")
        except Exception as e:
            print("ERROR:文件读取失败！请检查excels文件夹文件格式是否正确！", "\n")
            print(e)
            if input('输入y(小写)重新执行程序,其他键退出程序:') == 'y':
                continue
            else:
                break
        
        try:
            print("*" * 12 + "正在读取历史记录" + "*" * 12, "\n")
            print("*" * 10 + "数据量较大，请耐心等待！" + "*" * 10, "\n")
            # todo 耗时过长 有待优化
            filter = DataFilter()
            print("*" * 16 + "读取成功" + "*" * 16, "\n")
            print("*" * 14 + "正在进行数据筛选！" + "*" * 14, "\n")
            arr, in_arr = filter.run_filter()
            print("*" * 16 + "正在保存" + "*" * 16, "\n")
            
            filter_arr = SaveExcel('filter', '匹配失败', sheet_name, arr)
            filter_arr.save_excel()
            filter_in_arr = SaveExcel('filter', '匹配成功', sheet_name, in_arr)
            filter_in_arr.save_excel()
            print("*" * 16 + '筛选成功' + "*" * 16)
            print("*" * 6 + f"数据成功保存到filter文件夹中，请查看" + "*" * 6, "\n")
            if input('输入y(小写)重新执行程序,其他键退出程序:') == 'y':
                continue
            else:
                break
        except Exception as e:
            print("ERROR:请检查销售订单总历史记录.xlsx是否存在", "\n")
            print(e)
            if input('输入y(小写)重新执行程序,其他键退出程序:') == 'y':
                continue
            else:
                break
