## 发票打印小助手

# 发票打印小助手

一个简单的PDF到Word转换工具，可以将PDF文件的第一页转换为图片并插入到Word文档中，特别适用于发票等文档的转换。

## 功能

- 选择一个或多个PDF文件
- 将每个PDF的第一页转换为图片
- 将图片插入到Word文档中，每页最多两张图片
- 自动设置Word文档页边距为1.27厘米
- 自动打开输出文件夹

## 依赖库

- PyMuPDF (fitz) - PDF处理
- Pillow - 图像处理
- python-docx - Word文档生成
- tkinter - GUI界面

## 使用方法

1. 运行 `main.py`
2. 选择要转换的PDF文件
3. 程序会自动将PDF第一页转换为图片并插入到Word文档中
4. 输出文件位于 `output` 文件夹中，以当前时间命名

## 打包为可执行文件

本项目可以使用PyInstaller打包为Windows可执行文件：

```bash
# 安装依赖
pip install -r requirements.txt

# 安装PyInstaller
pip install PyInstaller

# 打包为可执行文件
pyinstaller --onefile --windowed --name=PDFConvertor main.py
```

打包后的可执行文件将位于 `dist` 文件夹中。

## 文件说明

- `main.py` - 主程序文件
- `requirements.txt` - 项目依赖
- `setup.py` - 打包配置文件
- `build_exe.py` - 自动打包脚本
- `output/` - 输出文件夹