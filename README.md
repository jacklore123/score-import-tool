# 成绩匹配导入工具

一个面向广西综合评价系统的成绩整理工具，用于把五岳、好分数等来源的原始成绩表，按学生姓名匹配到综合评价系统模板中，并自动导出可直接上传的表格。

## 功能

- 支持 `.xlsx`、`.xls`、`.zip` 和文件夹拖拽导入。
- 自动按“姓名”精确匹配学生，等价于 Excel `VLOOKUP(..., 0)`。
- 支持一份或多份原始表、系统表批量处理。
- 自动识别原始成绩表中的“姓名”列和“成绩”列，也允许手动配置列名。
- 自动处理缺考、空值、`/`、`-`、`缺考` 等标记。
- 自动根据学生名单配对，即使文件名不一致也能匹配。
- 处理重名学生并给出警告。
- 输出结果自动加时间戳，避免覆盖旧文件。
- 支持拖拽操作（安装 `tkinterdnd2` 后启用）。

## 运行环境

- Windows 10/11
- Python 3.10 或更高版本
- Tkinter（Python 自带）

## 安装依赖

```powershell
pip install -r requirements.txt
```

## 运行

```powershell
python main.py
```

也可以直接运行：

```powershell
python score_match_import_tool.py
```

## 使用流程

1. 准备“原始成绩表”：五岳或好分数导出的 Excel、压缩包或文件夹。
2. 准备“系统模板”：从综合评价系统下载的成绩导入模板。
3. 打开程序，把原始表和系统模板分别拖到对应区域。
4. 确认姓名列、成绩列、备注列配置。
5. 选择导出文件夹。
6. 点击“开始匹配”，等待生成结果。

更详细的说明见 [docs/使用说明.md](docs/使用说明.md)。

## 打包为 EXE

```powershell
pip install pyinstaller
pyinstaller --onefile --noconsole --name 成绩匹配导入工具 main.py
```

详细说明见 [docs/构建EXE.md](docs/构建EXE.md)。

## 目录结构

```text
.
├── main.py
├── score_match_import_tool.py
├── requirements.txt
├── README.md
├── LICENSE
├── .gitignore
└── docs/
    ├── 使用说明.md
    ├── 数据格式说明.md
    └── 构建EXE.md
```

## 隐私说明

本工具为本地运行，不会上传或发送学生成绩数据。原始成绩表、系统模板和导出结果都保存在本机，请妥善保管。

## License

[MIT](LICENSE)
