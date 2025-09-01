# HTML 转换器

一个基于 Python 的图形界面工具，用于将 HTML 文件转换为 DOCX 或 TXT 格式。

## 功能特性

- 📄 支持 HTML 到 DOCX 格式转换
- 📝 支持 HTML 到 TXT 格式转换
- 🖼️ 图形用户界面，操作简单
- 🎯 自动提取纯文本内容，去除脚本和样式
- 📁 自定义输出目录和文件名

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

1. **运行程序**：
   ```bash
   python html_converter.py
   ```
   或者双击 `run.bat` 文件

2. **选择文件**：点击"浏览..."按钮选择要转换的 HTML 文件

3. **设置输出**：
   - 选择输出目录
   - 输入输出文件名
   - 选择转换格式（DOCX 或 TXT）

4. **开始转换**：点击"开始转换"按钮

## 支持的格式

- **DOCX**: Microsoft Word 文档格式，保留基本的标题和段落结构
- **TXT**: 纯文本格式，提取 HTML 中的文本内容

## 项目结构

```
html_converter/
├── html_converter.py  # 主程序文件
├── requirements.txt   # 依赖包列表
├── run.bat           # Windows 启动脚本
└── README.md         # 项目说明文档
```

## 依赖库

- `beautifulsoup4==4.12.2` - HTML 解析
- `python-docx==1.1.0` - DOCX 文档操作
- `tkinter` - 图形界面（Python 标准库）

## 系统要求

- Python 3.6+
- Windows 操作系统（支持其他系统但界面可能有所不同）

## 许可证

MIT License

