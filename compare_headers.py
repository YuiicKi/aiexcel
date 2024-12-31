import pandas as pd
from openai import OpenAI
import os

# 当前使用的模型
CURRENT_MODEL = "***"

# 性能测试数据保存路径
def get_performance_data_file():
    """根据模型名生成性能数据文件路径"""
    return f"performance_data_{CURRENT_MODEL}.json"

def load_performance_data():
    try:
        file_path = get_performance_data_file()
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return {
            "runs": [],
            "total_average": 0,
            "model": CURRENT_MODEL
        }

def save_performance_data(data):
    file_path = get_performance_data_file()
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def read_excel_headers(file_path, header_row=0):
    """读取Excel文件的表头"""
    try:
        df = pd.read_excel(file_path, header=header_row)
        return df.columns.tolist()
    except Exception as e:
        print(f"读取Excel文件出错：{str(e)}")
        return None

def compare_headers_with_ai(headers1, headers2, is_comparing_2_and_3=False, is_comparing_1_and_3=False):
    """使用AI比较两个表的表头并建立映射关系"""
    
    # 初始化API客户端
    client = OpenAI(
        api_key="***",
        base_url="***"
    )
    
    # 构建提示信息
    headers1_str = "\n".join([f"{i+1}. {h}" for i, h in enumerate(headers1)])
    headers2_str = "\n".join([f"{i+1}. {h}" for i, h in enumerate(headers2)])
    
    # 根据不同的比较场景选择不同的提示信息
    if is_comparing_2_and_3:
        prompt = f"""请分析以下两个表格的表头，并建立它们之间的对应关系。
这是表2的表头：
{headers1_str}

这是表3的表头：
{headers2_str}

请分析每个表头的含义，找出相同或相似的字段。对于每个找到的对应关系，请使用以下格式输出：
表2的第X列 对应 表3的第Y列

注意：
1. 只输出确定的对应关系
2. 每行只写一个对应关系
3. 不需要解释原因
4. 如果找不到对应关系，就不要输出
5. 请特别注意"学号"列的对应关系，这个是关键"""
    elif is_comparing_1_and_3:
        prompt = f"""请分析以下两个表格的表头，并建立它们之间的对应关系。
这是表1的表头：
{headers1_str}

这是表3的表头：
{headers2_str}

请分析每个表头的含义，找出相同或相似的字段。对于每个找到的对应关系，请使用以下格式输出：
表1的第X列 对应 表3的第Y列

注意：
1. 只输出确定的对应关系
2. 每行只写一个对应关系
3. 不需要解释原因
4. 如果找不到对应关系，就不要输出
5. 请特别注意"学号"列的对应关系，这个是关键"""
    else:
        prompt = f"""请分析以下两个表格的表头，并建立它们之间的对应关系。
这是表1的表头：
{headers1_str}

这是表2的表头：
{headers2_str}

请分析每个表头的含义，找出相同或相似的字段。对于每个找到的对应关系，请使用以下格式输出：
表1的第X列 对应 表2的第Y列

注意：
1. 只输出确定的对应关系
2. 每行只写一个对应关系
3. 不需要解释原因
4. 如果找不到对应关系，就不要输出
5. 请特别注意"学号"列的对应关系，这个是关键"""

    try:
        # 调用API进行分析
        response = client.chat.completions.create(
            model="***",
            messages=[
                {"role": "system", "content": "你是一个专门分析Excel表头并建立映射关系的助手。请直接返回映射关系，不要返回其他内容。"},
                {"role": "user", "content": prompt}
            ]
        )
        
        # 返回API响应
        return response.choices[0].message.content
        
    except Exception as e:
        print(f"调用AI API出错：{str(e)}")
        return None

def run_performance_test(headers1, headers2, is_comparing_2_and_3=False, is_comparing_1_and_3=False):
    """
    运行性能测试
    """
    performance_data = load_performance_data()
    run_times = []
    
    print(f"\n开始性能测试... (使用模型: {CURRENT_MODEL})")
    final_result = None
    for i in range(10):
        print(f"\n运行测试 {i+1}/10")
        start_time = time.time()
        result = compare_headers_with_ai(headers1, headers2, is_comparing_2_and_3, is_comparing_1_and_3)
        end_time = time.time()
        
        elapsed_time = end_time - start_time
        run_times.append(elapsed_time)
        print(f"耗时: {elapsed_time:.2f} 秒")
        
        # 记录运行数据
        run_data = {
            "run_number": i + 1,
            "time": elapsed_time,
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            "comparison_type": "表1和表3" if is_comparing_1_and_3 else "表1和表2"
        }
        performance_data["runs"].append(run_data)
        final_result = result
    
    # 计算总平均耗时
    all_times = [run["time"] for run in performance_data["runs"]]
    total_average = sum(all_times) / len(all_times)
    performance_data["total_average"] = total_average
    performance_data["model"] = CURRENT_MODEL
    
    # 保存数据
    save_performance_data(performance_data)
    
    # 只打印总平均耗时
    if is_comparing_1_and_3:
        print(f"\n所有比较的平均耗时: {total_average:.2f} 秒")
        print(f"性能数据保存到: {get_performance_data_file()}")
    
    return final_result

def main():
    # 指定Excel文件路径
    file1_path = "***"
    file2_path = "***"
    file3_path = "***"
    
    # 读取表头
    headers1 = read_excel_headers(file1_path, header_row=1)
    headers2 = read_excel_headers(file2_path)
    headers3 = read_excel_headers(file3_path, header_row=1)
    
    if headers1 and headers2 and headers3:
        print("\n=== 表1表头 ===")
        for i, header in enumerate(headers1, 1):
            print(f"{i}. {header}")
        
        print("\n=== 表2表头 ===")
        for i, header in enumerate(headers2, 1):
            print(f"{i}. {header}")
        
        print("\n=== 表3表头 ===")
        for i, header in enumerate(headers3, 1):
            print(f"{i}. {header}")
        
        print("\n=== AI分析结果: 表1和表2对应关系 ===")
        result1 = run_performance_test(headers1, headers2)
        if not result1 or not result1.strip():
            print("无法获取表1和表2之间的对应关系")
            
        print("\n=== AI分析结果: 表1和表3对应关系 ===")
        result3 = run_performance_test(headers1, headers3, is_comparing_1_and_3=True)
        if not result3 or not result3.strip():
            print("无法获取表1和表3之间的对应关系")

if __name__ == "__main__":
    main()
