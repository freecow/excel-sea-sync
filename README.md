# Excel to SeaTable 数据同步工具

## 🚀 项目简介

这是一个专门用于将Excel文件数据同步到SeaTable表格的自动化工具。支持多表同步、字段映射、数据类型转换以及分批处理大量数据。

### ✨ 主要功能

- 📊 **Excel多表同步**：支持从单个Excel文件中同步多个工作表到SeaTable
- 🔄 **智能字段映射**：灵活的字段名称映射，适应不同的表格结构
- 🎯 **数据类型转换**：自动处理日期、数字等数据类型格式化
- ⚡ **分批处理**：支持大数据量的分批上传，避免超时问题
- 🎮 **交互式选择**：提供友好的命令行界面选择不同的同步任务
- 🌍 **跨平台支持**：支持Windows、Linux、macOS三大平台
- 🔧 **灵活配置**：支持环境变量和JSON配置文件

## 📋 系统要求

- Python 3.9+
- 网络连接（访问SeaTable服务）
- Excel文件（.xlsx格式）

## 🛠️ 安装与配置

### 方式一：使用编译好的可执行文件（推荐）

1. **下载可执行文件**
   
   从[Releases页面](../../releases)下载对应平台的可执行文件：
   - Windows: `excel-sea-sync-windows.exe`
   - Linux: `excel-sea-sync-linux`
   - macOS: `excel-sea-sync-macos`

2. **配置环境变量**
   
   创建`.env`文件（复制.env.example）：
   ```bash
   cp .env.example .env
   ```
   
   编辑`.env`文件，填入你的配置：
   ```env
   # SeaTable 服务器地址
   SEATABLE_SERVER_URL=https://cloud.seatable.cn
   
   # 不同项目的API Token
   SEATABLE_BH_GOV_TOKEN=你的政企项目token
   SEATABLE_BH_STAR_TOKEN=你的卫星项目token
   SEATABLE_BH_YWL_TOKEN=你的云未来项目token
   SEATABLE_BH_YXD_TOKEN=你的云现代项目token
   ```

3. **运行程序**
   ```bash
   # Windows
   excel-sea-sync-windows.exe
   
   # Linux/macOS
   ./excel-sea-sync-linux
   ./excel-sea-sync-macos
   ```

### 方式二：源码安装

1. **克隆仓库**
   ```bash
   git clone <repository-url>
   cd excel-sea-sync
   ```

2. **安装依赖**
   ```bash
   pip install -r requirements.txt
   ```

3. **配置环境变量**（同上）

4. **运行程序**
   ```bash
   python excel-sync.py
   ```

## 🔧 配置说明

### 环境变量配置

在`.env`文件中配置以下变量：

| 变量名 | 说明 | 示例 |
|--------|------|------|
| `SEATABLE_SERVER_URL` | SeaTable服务器地址 | `https://cloud.seatable.cn` |
| `SEATABLE_BH_GOV_TOKEN` | 博浩政企项目API Token | `your_token_here` |
| `SEATABLE_BH_STAR_TOKEN` | 博浩卫星项目API Token | `your_token_here` |
| `SEATABLE_BH_YWL_TOKEN` | 博浩云未来项目API Token | `your_token_here` |
| `SEATABLE_BH_YXD_TOKEN` | 博浩云现代项目API Token | `your_token_here` |

### JSON配置文件

项目包含四个预配置的JSON文件：

- `memo-bh-gov.json` - 博浩政企数据同步配置
- `memo-bh-star.json` - 博浩卫星数据同步配置
- `memo-bh-ywl.json` - 博浩云未来数据同步配置
- `memo-bh-yxd.json` - 博浩云现代数据同步配置

#### 配置文件结构

```json
{
  "chunk_size": 300,
  "excel_config": {
    "file_path": "/path/to/your/excel/file.xlsx"
  },
  "tables": [
    {
      "seatable": {
        "table_name": "目标表名",
        "name_column": "主键列名"
      },
      "excel_sheet": "Excel工作表名",
      "start_row": 1,
      "field_mappings": {
        "Excel列名1": "SeaTable字段名1",
        "Excel列名2": "SeaTable字段名2"
      },
      "data_types": {
        "SeaTable字段名1": "date",
        "SeaTable字段名2": "number"
      }
    }
  ]
}
```

#### 配置参数说明

| 参数 | 说明 | 类型 | 示例 |
|------|------|------|------|
| `chunk_size` | 分批上传的数据量 | 数字 | `300` |
| `excel_config.file_path` | Excel文件完整路径 | 字符串 | `/path/to/file.xlsx` |
| `seatable.table_name` | SeaTable中的目标表名 | 字符串 | `合同档案` |
| `seatable.name_column` | 主键列名（用于标识） | 字符串 | `BH合同编号` |
| `excel_sheet` | Excel中的工作表名 | 字符串 | `合同档案` |
| `start_row` | 数据开始行号（1开始） | 数字 | `1` |
| `field_mappings` | Excel列到SeaTable字段的映射 | 对象 | 见上方示例 |
| `data_types` | 字段数据类型定义 | 对象 | `{"字段名": "date"}` |

#### 支持的数据类型

- `date` - 日期类型，格式：YYYY-MM-DD
- `number` - 数字类型，带千分位逗号格式
- 默认为文本类型

## 🚀 使用指南

### 基本使用流程

1. **启动程序**
   ```bash
   ./excel-sea-sync
   ```

2. **选择同步任务**
   
   程序会显示可用的同步任务菜单：
   ```
   ===== Excel同步任务选择 =====
   1. 博浩政企数据同步
   2. 博浩卫星数据同步
   3. 博浩云未来数据同步
   4. 博浩云现代数据同步
   
   0. 退出程序
   
   请选择要执行的同步任务 (1):
   ```

3. **执行同步**
   
   选择任务后，程序会：
   - 加载相应的配置文件
   - 验证API Token
   - 连接SeaTable服务
   - 清空目标表格
   - 读取Excel数据
   - 分批上传到SeaTable

### 运行示例

```bash
$ ./excel-sea-sync

===== Excel同步任务选择 =====
1. 博浩政企数据同步
2. 博浩卫星数据同步
3. 博浩云未来数据同步
4. 博浩云现代数据同步

0. 退出程序

请选择要执行的同步任务 (1): 1

已选择: 博浩政企数据同步
配置文件: memo-bh-gov.json

开始同步表格: 合同档案
Excel工作表: 合同档案
Found 1500 rows to delete in table '合同档案'
Table '合同档案' has been successfully cleared.
Reading Excel sheet: 合同档案
Processing data...
Inserting 850 rows into SeaTable...
Table '合同档案' sync completed successfully!

开始同步表格: 订单档案
Excel工作表: 订单档案
...

所有表格同步完成！
```

## 📝 配置自定义项目

### 创建新的配置文件

1. **复制现有配置**
   ```bash
   cp memo-bh-gov.json my-project.json
   ```

2. **修改配置内容**
   
   编辑`my-project.json`，修改以下内容：
   - Excel文件路径
   - 表格映射关系
   - 字段映射
   - 数据类型定义

3. **添加环境变量**
   
   在`.env`文件中添加新的API Token：
   ```env
   SEATABLE_MY_PROJECT_TOKEN=your_new_token
   ```

4. **修改代码**
   
   在`excel-sync.py`的`select_configuration()`函数中添加新选项：
   ```python
   config_options = {
       # ... 现有选项 ...
       5: {
           "name": "我的项目数据同步",
           "config_file": "my-project.json",
           "api_token_env": "SEATABLE_MY_PROJECT_TOKEN"
       }
   }
   ```

### 字段映射配置技巧

1. **字段名完全匹配**
   ```json
   "field_mappings": {
     "Excel中的列名": "SeaTable中的字段名"
   }
   ```

2. **处理特殊字符**
   - Excel列名可以包含空格、中文
   - SeaTable字段名建议使用英文或中文
   - 映射关系必须精确匹配

3. **数据类型处理**
   ```json
   "data_types": {
     "金额字段": "number",      // 自动添加千分位逗号
     "日期字段": "date",        // 转换为YYYY-MM-DD格式
     "普通字段": ""             // 不指定类型，保持原始文本
   }
   ```

## 🐛 常见问题与解决方案

### 1. 连接问题

**问题**：无法连接到SeaTable服务
```
ERROR: 获取可用表失败: HTTP 401
```

**解决方案**：
- 检查API Token是否正确
- 确认SeaTable服务器地址是否正确
- 验证网络连接是否正常

### 2. Excel文件问题

**问题**：Excel文件读取失败
```
ERROR: No such file or directory: '/path/to/file.xlsx'
```

**解决方案**：
- 确认Excel文件路径是否正确
- 检查文件是否存在
- 确认文件格式为.xlsx
- 检查文件是否被其他程序占用

### 3. 字段映射问题

**问题**：某些字段数据丢失
```
KeyError: '字段名不存在'
```

**解决方案**：
- 检查Excel列名是否与配置文件中的映射一致
- 确认Excel工作表名称正确
- 验证起始行号设置

### 4. 数据类型问题

**问题**：日期或数字格式错误

**解决方案**：
- 检查Excel中的数据格式
- 确认`data_types`配置正确
- 对于日期字段，确保Excel中是日期格式
- 对于数字字段，确保Excel中是数值格式

### 5. 权限问题

**问题**：无法写入SeaTable表格
```
ERROR: Permission denied
```

**解决方案**：
- 确认API Token具有表格写入权限
- 检查表格是否存在
- 确认表格名称拼写正确

### 6. 大数据量处理

**问题**：数据量大时出现超时

**解决方案**：
- 减小`chunk_size`值（如从300改为100）
- 分批次处理大文件
- 检查网络稳定性

## 🔧 开发与构建

### 本地开发

1. **设置开发环境**
   ```bash
   git clone <repository-url>
   cd excel-sea-sync
   pip install -r requirements.txt
   ```

2. **运行测试**
   ```bash
   python excel-sync.py
   ```

### 构建可执行文件

#### 本地构建

```bash
# 通用构建
python build_standalone.py

# Windows专用构建
python build_windows_ci.py
```

#### GitHub Actions自动构建

项目配置了GitHub Actions，会在以下情况自动构建：
- 推送到main/master分支
- 创建Pull Request
- 发布Release

构建产物会上传到Actions页面，Release时会自动附加到发布页面。

### 项目结构

```
excel-sea-sync/
├── excel-sync.py              # 主程序文件
├── requirements.txt           # Python依赖
├── .env.example              # 环境变量模板
├── README.md                 # 项目文档
├── memo-bh-gov.json          # 政企项目配置
├── memo-bh-star.json         # 卫星项目配置
├── memo-bh-ywl.json          # 云未来项目配置
├── memo-bh-yxd.json          # 云现代项目配置
├── build_standalone.py       # 通用构建脚本
├── build_windows_ci.py       # Windows构建脚本
└── .github/
    └── workflows/
        └── build.yml         # GitHub Actions配置
```

## 📄 许可证

本项目采用 [MIT License](LICENSE) 许可证。

## 🤝 贡献指南

欢迎提交Issue和Pull Request来改进这个项目！

### 提交Issue

如果您遇到问题或有功能建议，请：
1. 检查是否已有类似的Issue
2. 提供详细的问题描述
3. 包含错误信息和环境信息
4. 如果可能，提供复现步骤

### 提交代码

1. Fork此仓库
2. 创建功能分支 (`git checkout -b feature/amazing-feature`)
3. 提交更改 (`git commit -m 'Add some amazing feature'`)
4. 推送到分支 (`git push origin feature/amazing-feature`)
5. 创建Pull Request

## 📞 支持与联系

如果您需要帮助或有任何问题，请：
- 查看本README的常见问题部分
- 创建GitHub Issue
- 发送邮件至：[your-email@example.com]

## 🔄 更新日志

### v1.0.0 (2025-01-XX)
- 初始版本发布
- 支持Excel到SeaTable的数据同步
- 多项目配置支持
- 跨平台可执行文件
- 完整的字段映射和数据类型转换

---

**感谢使用Excel to SeaTable数据同步工具！** 🎉