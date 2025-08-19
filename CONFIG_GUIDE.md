# Excel-SeaTable 同步工具配置说明

## 简介

本工具用于将Excel文件中的数据同步到SeaTable表格中，支持多表同步和字段映射。


## 运行方式

#### Windows:
```cmd
excel-sea-sync.exe
```

#### Linux/macOS:
```bash
./excel-sea-sync
```

运行后会显示同步任务选择菜单，根据提示选择对应的任务即可。

## 配置文件说明

### .env 环境变量配置（推荐）

创建 `.env` 文件，配置SeaTable连接信息：

```env
# SeaTable服务器地址
SEATABLE_SERVER_URL=https://cloud.seatable.cn

# 各项目的API Token（统一自动化命名规则）
SEATABLE_MEMO_BH_GOV_TOKEN=your_gov_token_here
SEATABLE_MEMO_BH_STAR_TOKEN=your_star_token_here
SEATABLE_MEMO_BH_YWL_TOKEN=your_ywl_token_here
SEATABLE_MEMO_BH_YXD_TOKEN=your_yxd_token_here
```

### JSON 配置文件

项目包含4个预配置的JSON文件：

- `memo-bh-gov.json` - 博浩政企数据同步配置
- `memo-bh-star.json` - 博浩卫星数据同步配置
- `memo-bh-ywl.json` - 博浩云未来数据同步配置
- `memo-bh-yxd.json` - 博浩云现代数据同步配置

#### JSON配置文件结构

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
        "Excel列名": "SeaTable字段名"
      },
      "data_types": {
        "字段名": "数据类型"
      }
    }
  ]
}
```

#### 配置参数说明

- `chunk_size`: 分批处理的数据行数，建议300
- `excel_config.file_path`: Excel文件的完整路径
- `tables`: 表格同步配置数组
  - `seatable.table_name`: SeaTable中的目标表名
  - `seatable.name_column`: 主键列名
  - `excel_sheet`: Excel中的工作表名
  - `start_row`: 数据开始行号（1为第一行）
  - `field_mappings`: Excel列名到SeaTable字段名的映射
  - `data_types`: 字段数据类型定义
    - `"number"`: 数字类型（自动格式化为千分位）
    - `"date"`: 日期类型（格式化为YYYY-MM-DD）

## 使用流程

1. **准备配置**
   - 创建 `.env` 文件并填入SeaTable连接信息
   - 确认JSON配置文件中的Excel文件路径正确

2. **运行同步工具**
   ```bash
   excel-sea-sync.exe
   ```

3. **选择同步任务**
   - 根据提示选择对应的项目（1-4）
   - 程序会自动加载对应的配置文件

4. **执行同步**
   - 程序会逐个同步配置中的表格
   - 同步前会清空目标表格数据
   - 显示同步进度和结果

## 注意事项

- 确保网络能访问SeaTable服务器
- 确保API Token有相应表格的读写权限
- 同步操作会清空目标表格的现有数据
- Excel文件必须是.xlsx格式
- 建议在同步前备份重要数据

## 故障排除

### 常见错误

1. **找不到环境变量**
   ```
   错误：未找到环境变量 SEATABLE_BH_XXX_TOKEN
   ```
   解决：检查 `.env` 文件是否存在且包含正确的Token

2. **Excel文件路径错误**
   ```
   Error during sync for table 'xxx': [Errno 2] No such file or directory
   ```
   解决：检查JSON配置文件中的 `excel_config.file_path` 路径是否正确

3. **SeaTable连接失败**
   ```
   获取可用表失败
   ```
   解决：检查网络连接和SeaTable服务器地址、API Token是否正确

## Token环境变量命名规则

所有Token环境变量都遵循统一的自动化规则：
- **格式**：`SEATABLE_{配置文件名大写}_TOKEN`
- **配置文件名处理**：移除.json后缀，将 `-` 替换为 `_`，转为大写

### 示例：
- `memo-bh-gov.json` → `SEATABLE_MEMO_BH_GOV_TOKEN`
- `memo-bh-star.json` → `SEATABLE_MEMO_BH_STAR_TOKEN`
- `memo-bh-test.json` → `SEATABLE_MEMO_BH_TEST_TOKEN`
- `my-project.json` → `SEATABLE_MY_PROJECT_TOKEN`

### 重要提醒：
如果你已经有旧的.env文件，请按照新的命名规则更新Token名称。

