# 构建 EXE

## 安装打包工具

```powershell
pip install pyinstaller
```

## 构建图形界面 EXE

```powershell
pyinstaller --onefile --noconsole --name 成绩匹配导入工具 main.py
```

构建结果在 `dist` 文件夹中。

## 常见问题

### 杀毒软件误报

PyInstaller 单文件 EXE 有时会被杀毒软件误报，可将 `build` 和 `dist` 加入信任区。

### 拖拽功能不可用

请确认已安装 `tkinterdnd2`：

```powershell
pip install tkinterdnd2
```

### 中文路径

程序支持中文路径。若旧系统控制台乱码，可先运行：

```powershell
chcp 65001
```
