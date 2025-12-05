# go-license

A GUI-based license analysis tool for Go and Node.js projects

一个基于GUI的Go和Node.js项目许可证分析工具

## Features 功能特性

- **Dual Support** 双重支持：支持解析 Go 模块 (go.mod) 和 Node.js 项目 (package.json)
- **Multi-source Metadata** 多源元数据：从 pkg.go.dev 和 npm registry 获取许可证信息
- **Rich Information** 丰富信息：提取许可证、作者、描述、版权、仓库链接等详细信息
- **Excel Export** Excel导出：生成格式化的Excel报告，便于查看和管理
- **GUI Interface** 图形界面：使用文件选择对话框，用户友好
- **Progress Tracking** 进度跟踪：实时显示处理进度

## Usage 使用方法

1. Run the program:
```bash
go run main.go
```

2. 选择文件类型：
   - 对于 Go 项目，选择 `go.mod` 文件
   - 对于 Node.js 项目，选择 `package.json` 文件

The tool will automatically detect the file type and process accordingly.
工具会自动检测文件类型并进行相应处理。

## Output 输出内容

### For Go modules (go.mod):
生成的Excel文件 `{module-name}-api_license.xlsx` 包含：
- **Name** - 包名称
- **License** - 许可证类型
- **PackageVersion** - 包版本
- **LicenseURL** - 许可证URL
- **Author** - 作者
- **Description** - 描述
- **Copyright** - 版权信息
- **PackageURL** - 包URL
- **GitHubURL** - GitHub链接
- **RepositoryType** - 仓库类型

### For Node.js projects (package.json):
生成的Excel文件 `{package-name}-ui_license.xlsx` 包含：
- **Module Name** - 模块名称 (包含版本)
- **License** - 许可证类型
- **Repository** - 仓库地址
- **License URL** - 许可证URL
- **Author** - 作者
- **Description** - 描述
- **Copyright** - 版权信息
- **GitHub URL** - GitHub链接
- **Module Name (No Version)** - 模块名称（不含版本）
- **Version** - 版本号

## Requirements 环境要求

- Go 1.24.0 or higher / Go 1.24.0 或更高版本

## Installation 安装步骤

```bash
git clone <repository-url>
cd go-license
go mod tidy
go run main.go
```

## Build Binary 构建可执行文件

```bash
go build -o go-license.exe main.go
```

## Dependencies 依赖库

- **[htmlquery](https://github.com/antchfx/htmlquery)** - HTML parsing library / HTML解析库
- **[zenity](https://github.com/ncruces/zenity)** - Cross-platform GUI dialogs / 跨平台GUI对话框
- **[excelize](https://github.com/xuri/excelize/v2)** - Excel file operations / Excel文件操作
- **[golang.org/x/mod](https://golang.org/x/mod)** - Go module parsing / Go模块解析

## Technical Details 技术细节

### Data Sources 数据源
- **Go modules**: https://pkg.go.dev/
- **Node.js packages**: https://registry.npmjs.org/

### Error Handling 错误处理
- Network requests use context with 10-second timeout
- 网络请求使用带有10秒超时的上下文
- Graceful handling of missing metadata
- 优雅处理缺失的元数据
- User-friendly error messages with zenity dialogs
- 使用zenity对话框显示用户友好的错误消息

### License URL Generation 许可证URL生成
The tool generates license URLs using: https://licenses.nuget.org/{LICENSE_TYPE}
工具使用以下格式生成许可证URL：https://licenses.nuget.org/{许可证类型}

## Project Evolution 项目演进

该项目经过多次重构和优化：
- 改进了依赖信息获取逻辑
- 优化了错误处理和取消逻辑
- 支持多种项目类型（Go 和 npm）
- 提供更丰富的输出格式和字段
- 添加了图形用户界面支持

## Author 作者

License Tool / 许可证工具
