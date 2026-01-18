# Excel数据汇总工具

一个基于Flask的Web应用，用于合并和汇总Excel文件数据。

## 功能特性

- 📊 支持 `.xlsx` 和 `.xls` 格式
- 🔍 智能识别复杂的多行标题结构
- 🎯 灵活的字段映射和选择
- 📑 支持多Sheet文件处理
- 🌐 Web界面操作，简单易用

## 使用方法

### 方式1：直接运行（开发环境）

```bash
# 安装依赖
pip install -r requirements.txt

# 运行服务
python app.py
```

访问 http://localhost:5000

### 方式2：Windows EXE（已打包）

1. 下载 `Excel数据汇总工具.exe`
2. 双击运行
3. 浏览器会自动打开 http://localhost:5000

## 打包说明

本项目使用 GitHub Actions 自动打包 Windows EXE 文件。

### 触发打包

- 推送到分支 `cursor/excelmergertool-windows-exe-da3a` 会自动触发
- 或在 GitHub Actions 页面手动触发

### 下载打包产物

1. 进入 GitHub Actions 页面
2. 选择最新的 workflow 运行记录
3. 在 Artifacts 部分下载 `Excel数据汇总工具-Windows`

## 项目结构

```
ExcelMergerTool/
├── app.py                 # Flask后端主程序
├── templates/
│   └── index.html        # 前端界面
├── requirements.txt       # Python依赖
└── .github/workflows/
    └── build-windows-exe.yml  # GitHub Actions配置
```

## 依赖

- Flask 3.0.0
- pandas 2.1.3
- openpyxl 3.1.2
- xlrd 2.0.1
- flask-cors 4.0.0
- pyinstaller 6.3.0（仅用于打包）

## License

MIT
