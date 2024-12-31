import pandas as pd

def read_excel_headers(file_path, header_row=0):
    """读取Excel文件的表头
    
    参数:
        file_path: Excel文件路径
        header_row: 表头所在行号（默认为0，即第一行）
        
    返回:
        成功返回表头列表，失败返回None
    """
    try:
        # 读取Excel文件，指定header行
        df = pd.read_excel(file_path, header=header_row)
        # 获取列名（表头）
        return df.columns.tolist()
    except Exception as e:
        print(f"读取Excel文件出错：{str(e)}")
        return None

def main():
    # 直接指定两个Excel文件的路径
    file1_path = "***"  # 替换为你的第一个Excel文件路径
    file2_path = "***"  # 替换为你的第二个Excel文件路径
    
    print("\n=== 第一个Excel文件的表头 ===")
    headers1 = read_excel_headers(file1_path, header_row=1)  # 使用第二行作为表头（索引为1）
    if headers1:
        for i, header in enumerate(headers1, 1):
            print(f"{i}. {header}")
    
    print("\n=== 第二个Excel文件的表头 ===")
    headers2 = read_excel_headers(file2_path)  # 第二个文件仍使用第一行作为表头
    if headers2:
        for i, header in enumerate(headers2, 1):
            print(f"{i}. {header}")

if __name__ == "__main__":
    main()
