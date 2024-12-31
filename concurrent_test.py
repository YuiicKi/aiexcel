import multiprocessing
from compare_headers import read_excel_headers, compare_headers_with_ai
from tqdm import tqdm
import json
from collections import Counter
import time
import os

def process_single_comparison(headers1, headers3):
    try:
        result = compare_headers_with_ai(headers1, headers3, is_comparing_2_and_3=False, is_comparing_1_and_3=True)
        return result.strip() if result else ""
    except Exception as e:
        return f"Error: {str(e)}"

def run_concurrent_tests(total_tests=500):
    # 读取表头
    file1_path = "/Users/yuii/Downloads/aiexcel/副本xx班xx假期离返校去向表.xlsx"
    file3_path = "/Users/yuii/Downloads/aiexcel/人工智能-34人-教务原始数据.xlsx"
    
    headers1 = read_excel_headers(file1_path, header_row=1)
    headers3 = read_excel_headers(file3_path, header_row=1)
    
    if not headers1 or not headers3:
        print("无法读取表头数据")
        return

    # 创建进程池
    num_processes = min(multiprocessing.cpu_count() * 2, total_tests)  # 使用CPU核心数的2倍作为进程数
    pool = multiprocessing.Pool(processes=num_processes)
    
    print(f"使用 {num_processes} 个进程进行并发测试")
    
    # 创建参数列表
    args = [(headers1, headers3) for _ in range(total_tests)]
    
    # 使用进度条显示进度
    start_time = time.time()
    results = []
    with tqdm(total=total_tests, desc="Processing comparisons") as pbar:
        for result in pool.starmap(process_single_comparison, args):
            results.append(result)
            pbar.update()
    
    # 关闭进程池
    pool.close()
    pool.join()
    
    end_time = time.time()
    
    # 分析结果
    result_counter = Counter(results)
    total_time = end_time - start_time
    
    # 输出分析报告
    print("\n=== 测试结果分析 ===")
    print(f"总测试次数: {total_tests}")
    print(f"总耗时: {total_time:.2f} 秒")
    print(f"平均每次耗时: {(total_time/total_tests):.2f} 秒")
    print(f"进程数: {num_processes}")
    print("\n不同结果的分布:")
    
    # 将结果保存到文件
    output_data = {
        "test_info": {
            "total_tests": total_tests,
            "total_time": total_time,
            "avg_time": total_time/total_tests,
            "num_processes": num_processes
        },
        "results": {str(k): v for k, v in result_counter.items()}
    }
    
    # 保存详细结果到JSON文件
    with open('concurrent_test_results.json', 'w', encoding='utf-8') as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)
    
    # 打印结果分布
    for result, count in result_counter.items():
        percentage = (count / total_tests) * 100
        print(f"\n出现 {count} 次 ({percentage:.2f}%):")
        print(result)

if __name__ == "__main__":
    multiprocessing.freeze_support()  # Windows支持
    run_concurrent_tests(total_tests=500)
