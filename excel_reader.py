import pandas as pd
import os
import sys

def read_excel_data(file_path):
    """
    从Excel文件中读取品牌和商品名称两列的内容
    
    参数:
        file_path (str): Excel文件的路径
    
    返回:
        pandas.DataFrame: 包含品牌和商品名称的数据框
    """
    try:
        # 检查文件是否存在
        if not os.path.exists(file_path):
            print(f"错误: 文件 '{file_path}' 不存在")
            return None
        
        # 检查文件扩展名
        if not file_path.endswith(('.xlsx', '.xls')):
            print(f"错误: 文件 '{file_path}' 不是Excel文件")
            return None
        
        # 读取Excel文件
        df = pd.read_excel(file_path)
        
        # 检查是否包含必要的列
        required_columns = ['品牌', '商品名称']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"错误: Excel文件缺少以下列: {', '.join(missing_columns)}")
            print(f"可用的列: {', '.join(df.columns)}")
            return None
        
        # 只保留品牌和商品名称两列
        result_df = df[['品牌', '商品名称']]
        
        # 删除空值行
        result_df = result_df.dropna()
        
        print(f"成功读取 {len(result_df)} 条数据")
        return result_df
    
    except Exception as e:
        print(f"读取Excel文件时发生错误: {str(e)}")
        return None

def main():
    # 检查命令行参数
    if len(sys.argv) < 2:
        print("使用方法: python excel_reader.py <Excel文件路径>")
        return
    
    # 从命令行参数获取文件路径
    file_path = sys.argv[1]
    
    # 读取数据
    data = read_excel_data(file_path)
    
    if data is not None:
        # 打印所有数据
        print("\n所有品牌和商品名称数据:")
        pd.set_option('display.max_rows', None)  # 显示所有行
        pd.set_option('display.max_columns', None)  # 显示所有列
        pd.set_option('display.width', None)  # 自动调整显示宽度
        pd.set_option('display.max_colwidth', None)  # 显示完整的列内容
        print(data)

if __name__ == "__main__":
    main()