import pandas as pd
import os

# 检查另一个可能包含原始数据的文件
source_file = '/app/data/files/男女歌手歌曲分离结果（两个表格）.xlsx'
if os.path.exists(source_file):
    try:
        # 使用pandas读取文件
        df_source = pd.read_excel(source_file, sheet_name=0)  # 读取第一个工作表
        
        print(f"源文件第一个工作表包含 {len(df_source)} 行，{len(df_source.columns)} 列")
        print("\n列名：")
        print(df_source.columns.tolist())
        
        # 显示前5行数据
        print("\n前5行数据：")
        print(df_source.head())
        
        # 尝试读取第二个工作表
        try:
            df_source2 = pd.read_excel(source_file, sheet_name=1)
            print(f"\n源文件第二个工作表包含 {len(df_source2)} 行，{len(df_source2.columns)} 列")
            print("列名：")
            print(df_source2.columns.tolist())
            print("\n前5行数据：")
            print(df_source2.head())
        except Exception as e:
            print(f"读取第二个工作表时出错：{e}")
    except Exception as e:
        print(f"读取源文件时出错：{e}")
else:
    print(f"源文件 {source_file} 不存在")