# GitHub Actions 自动打包 Windows EXE 使用指南

## 📋 前提条件

1. **GitHub账号**（如果没有，去 https://github.com 注册）
2. **Git已安装**（macOS通常自带）

## 🚀 快速开始（3步）

### 步骤1：创建GitHub仓库

1. 登录GitHub
2. 点击右上角 "+" → "New repository"
3. 仓库名称：`ExcelMergerTool`（或任意名称）
4. 选择 **Private**（私有，保护代码）
5. 点击 "Create repository"

### 步骤2：上传代码到GitHub

在终端中执行：

```bash
cd ~/Desktop/SOHUProgram/HealthSup/SohuSport/ExcelMergerTool

# 初始化git（如果还没有）
git init

# 添加所有文件
git add .

# 提交
git commit -m "Initial commit: Excel数据汇总工具"

# 添加远程仓库（替换YOUR_USERNAME为您的GitHub用户名）
git remote add origin https://github.com/YOUR_USERNAME/ExcelMergerTool.git

# 推送到GitHub
git branch -M main
git push -u origin main
```

### 步骤3：触发打包

**方法1：手动触发（推荐）**
1. 打开GitHub仓库页面
2. 点击 "Actions" 标签
3. 选择 "Build Windows EXE" workflow
4. 点击 "Run workflow" 按钮
5. 选择分支（通常是main），点击绿色按钮
6. 等待打包完成（约5-10分钟）

**方法2：自动触发**
- 每次推送代码到main分支时自动打包
- 或者创建Release时自动打包

## 📥 下载打包好的exe文件

打包完成后：

1. 在 "Actions" 页面找到最新的workflow运行记录
2. 点击进入详情页
3. 在 "Artifacts" 部分找到 "Excel数据汇总工具-Windows"
4. 点击下载，解压后得到 `Excel数据汇总工具.exe`

## 🔄 更新代码后重新打包

如果修改了代码，需要重新打包：

```bash
# 修改代码后
git add .
git commit -m "更新说明"
git push
```

然后按照步骤3手动触发打包，或等待自动触发。

## ⚙️ 自定义配置

### 修改触发条件

编辑 `.github/workflows/build-windows-exe.yml`：

- **只手动触发**：删除 `push:` 部分
- **修改分支**：修改 `branches:` 下的分支名
- **修改Python版本**：修改 `python-version: '3.9'`

### 添加其他功能

可以在workflow中添加：
- 自动创建Release
- 自动上传到网盘
- 发送通知邮件

## ❓ 常见问题

### Q1: 打包失败怎么办？

**检查：**
1. 查看Actions页面的错误日志
2. 检查代码是否有语法错误
3. 检查依赖是否正确

**常见错误：**
- 缺少依赖：在requirements.txt中添加
- 路径错误：检查templates文件夹是否存在

### Q2: 如何查看打包日志？

在Actions页面点击workflow运行记录，可以看到详细的打包日志。

### Q3: 打包需要多长时间？

通常5-10分钟，取决于：
- GitHub Actions的负载
- 依赖下载速度
- 代码复杂度

### Q4: 可以同时打包多个版本吗？

可以！创建多个workflow文件，分别打包不同配置。

### Q5: 打包产物保存多久？

默认30天，可以在workflow文件中修改 `retention-days: 30`

## 💡 提示

1. **私有仓库**：建议使用私有仓库保护代码
2. **定期清理**：定期下载并保存exe文件，避免过期
3. **版本管理**：每次打包可以打tag标记版本
4. **自动化**：设置自动打包，每次代码更新自动生成新版本

## 🎯 完整流程示例

```bash
# 1. 本地修改代码
# ... 编辑文件 ...

# 2. 提交并推送
git add .
git commit -m "修复bug"
git push

# 3. 在GitHub上触发打包（或等待自动触发）

# 4. 等待5-10分钟

# 5. 下载exe文件

# 6. 发送给朋友使用
```

---

**现在就开始使用GitHub Actions自动打包吧！** 🎉
