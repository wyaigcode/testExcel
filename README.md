# 📊 Excel 报表合并/拆分工具 (testExcel)

一个功能强大的 Python 命令行工具，用于批量合并和拆分 Excel 报表文件。

## ✨ 功能特性

### 合并功能 (Merge)
- **按行追加合并**：将多个 Excel 文件的数据纵向拼接到一个 Sheet 中，自动添加来源标识
- **按 Sheet 合并**:：将多个 Excel 文件合并为一个文件，每个源文件作为独立的 Sheet
- **文件夹批量合并**:：一键合并整个文件夹中的所有 Excel 文件

### 拆分功能 (Split)
- **按 Sheet 拆分**:：将一个多 Sheet 的 Excel 文件拆分为多个独立文件
- **按列值拆分**：根据指定列的值（如部门、地区）自动拆分为多个文件

### 其他
- **文件信息查看**：快速查看Excel 文件的 Sheet 列表、行列数等信息

## 🛠 安装

```bash
# 克隆项目
git clone https://github.com/wyaigcode/testExcel.git
cd testExcel

# 安装依赖
pip install -r requirements.txt
```

## 📕 使用方法

### 命令行用法

#### 合并文件

```bash
# 方式1：按行追加合并（多个文件纵向拼接)
python main.py merge --mode row --files data1.xlsx data2.xlsx -o merged.xlsx

# 方式2：按 Sheet 合并（每个文件为独立的 Sheet）
python main.py merge --mode sheet --files data1.xlsx data2.xlsx -o merged.xlsx

# 方式3：合并文件夹中的所有 Excel 文件
python main.py merge --mode row --folder ./sample_data -o merged.xlsx
```

#### 拆分文件

```bash
# 方式1：按 Sheet 拆分（每个 Sheet 生成一个文件）
python main.py split --mode sheet --file multi_sheet.xlsx -o ./output

# 方式2：按列值拆分（根据指定列的值分组）
python main.py split --mode column --file data.xlsx --column 部门 -o ./output
```

#### 查看文件信息

```bash
python main.py info --file data.xlsx
```

### Python 代码调用

```python
from excel_tool import ExcelMerger, ExcelSplitter, ExcelInfo

# 合并 - 按行追加
ExcelMerger.merge_by_row(
    file_list=['data1.xlsx', 'data2.xlsx'],
    output_file='merged.xlsx',
    add_source=True
)

# 合并 - 按 Sheet
ExcelMerger.merge_by_sheet(
    file_list=['data1.xlsx', 'data2.xlsx'],
    output_file='merged.xlsx'
)

# 拆分 - 按 Sheet
ExcelSplitter.split_by_sheet(
    input_file='multi_sheet.xlsx',
    output_dir='./output'
)

# 拆分 - 按列值
ExcelSplitter.split_by_column(
    input_file='data.xlsx',
    column_name='部门',
    output_dir='./output'
)

# 查看文件信息
info = ExcelInfo.get_info('data.xlsx')
print(info)
```

## 📁 目录结构

```
testExcel/
├── excel_tool.py         # 核心模块 - 合并/拆分逻辑
├── main.py              # 命令行入口 (CLI)
├── generate_samples.py    # 示例数据生成脚本
├── requirements.txt     # 依赖列表
├── README.md             # 项目文椄
└── .gitignore            # Git 忽略规则
```

## 📄 License

MIT License
