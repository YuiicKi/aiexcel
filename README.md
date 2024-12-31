# AI-Driven Excel Processing Tool

This project provides an intelligent solution for processing and merging Excel files with different formats using AI technology. It's particularly useful for scenarios where multiple departments require the same data in different formats.

## Features

- **Intelligent Column Mapping**: Uses AI to understand and map columns between different Excel formats
- **Automated Data Processing**: Merges multiple Excel files with different formats into a standardized output
- **Smart Caching**: Caches mapping results for improved performance
- **Format Standardization**: Automatically standardizes data formats (dates, student IDs, etc.)
- **Error Handling**: Robust error handling for various edge cases

## Project Structure

```
aiexcel/
├── merge_excel.py         # Main script for Excel file processing
├── compare_headers.py     # AI-powered header comparison logic
├── mapping_cache.py       # Caching mechanism for column mappings
├── read_excel_headers.py  # Excel header reading utilities
└── requirements.txt       # Project dependencies
```

## Requirements

- Python 3.8+
- Required packages:
  - pandas
  - openpyxl
  - openai

## Installation

1. Clone the repository
2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Place your Excel files in the project directory
2. Run the main script:
```bash
python merge_excel.py
```

## Configuration

- Configure API settings in `merge_excel.py`
- Adjust cache settings in `mapping_cache.py`
- Modify column mappings in `compare_headers.py`

## Features in Detail

### Intelligent Column Mapping
- Semantic understanding of column headers
- Automatic mapping between different formats
- Support for various naming conventions

### Data Processing
- Automated merging of multiple Excel files
- Format standardization
- Data validation and error checking

### Performance Optimization
- Smart caching mechanism
- Efficient data processing
- Minimal memory footprint

## Future Development

- Cloud integration
- Multi-user support
- Advanced analytics features
- API endpoint support

## Contributing

Contributions are welcome! Please feel free to submit pull requests.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

---

# AI驱动的Excel处理工具

[English](#ai-driven-excel-processing-tool) | [中文](#ai驱动的excel处理工具)

本项目提供了一个智能的解决方案，使用AI技术处理和合并不同格式的Excel文件。特别适用于多个部门需要使用相同数据但格式不同的场景。

## 功能特点

- **智能列映射**：使用AI技术理解和映射不同Excel格式之间的列
- **自动化数据处理**：将不同格式的多个Excel文件合并为标准化输出
- **智能缓存**：缓存映射结果以提高性能
- **格式标准化**：自动标准化数据格式（日期、学号等）
- **错误处理**：针对各种边缘情况的健壮错误处理

## 项目结构

```
aiexcel/
├── merge_excel.py         # Excel文件处理的主要脚本
├── compare_headers.py     # 基于AI的表头比较逻辑
├── mapping_cache.py       # 列映射的缓存机制
├── read_excel_headers.py  # Excel表头读取工具
└── requirements.txt       # 项目依赖
```

## 系统要求

- Python 3.8+
- 依赖包：
  - pandas
  - openpyxl
  - openai

## 安装步骤

1. 克隆仓库
2. 安装依赖：
```bash
pip install -r requirements.txt
```

## 使用方法

1. 将Excel文件放在项目目录中
2. 运行主脚本：
```bash
python merge_excel.py
```

## 配置说明

- 在 `merge_excel.py` 中配置API设置
- 在 `mapping_cache.py` 中调整缓存设置
- 在 `compare_headers.py` 中修改列映射

## 详细功能

### 智能列映射
- 表头的语义理解
- 不同格式之间的自动映射
- 支持各种命名规范

### 数据处理
- 自动合并多个Excel文件
- 格式标准化
- 数据验证和错误检查

### 性能优化
- 智能缓存机制
- 高效的数据处理
- 最小内存占用

## 未来开发计划

- 云端集成
- 多用户支持
- 高级分析功能
- API接口支持

## 参与贡献

欢迎提交Pull Request来改进这个项目！

## 许可证

本项目采用MIT许可证 - 详见LICENSE文件
