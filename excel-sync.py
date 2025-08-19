"""
Excel to SeaTable 数据同步工具
创建日期：2025-05-24
功能：将Excel文件中的数据同步到SeaTable表格中
支持多表同步和字段映射
"""

import json
import pandas as pd
from seatable_api import Base
from dotenv import load_dotenv
import os
import sys
from datetime import datetime
import numpy as np
import glob
import re

# 加载 .env 文件中的环境变量
load_dotenv()

# SeaTable配置
seatable_config = {
    'server_url': os.getenv('SEATABLE_SERVER_URL'),
    'api_token': os.getenv('SEATABLE_API_TOKEN')
}

# 全局变量
chunk_size = None
tables_config = None
config = None

def load_config(config_file):
    """加载配置文件"""
    global chunk_size, tables_config, config
    with open(config_file, 'r', encoding='utf-8') as f:
        loaded_config = json.load(f)
    chunk_size = loaded_config['chunk_size']
    tables_config = loaded_config['tables']
    config = loaded_config

def clear_table(base, table_name):
    """清空SeaTable表格"""
    # 获取所有行
    rows = base.list_rows(table_name)
    if not rows:
        print(f"Table '{table_name}' is already empty.")
        return
        
    row_ids = [row['_id'] for row in rows]
    print(f"Found {len(row_ids)} rows to delete in table '{table_name}'")

    # 批量删除行
    for i in range(0, len(row_ids), chunk_size):
        chunk = row_ids[i:i + chunk_size]
        try:
            base.batch_delete_rows(table_name, chunk)
        except Exception as e:
            print(f"Error deleting chunk: {e}")
            continue

    # 验证表是否真的被清空
    remaining_rows = base.list_rows(table_name)
    if remaining_rows:
        print(f"Warning: Table '{table_name}' still has {len(remaining_rows)} rows remaining after deletion attempt.")
        # 如果还有剩余行，尝试再次删除
        remaining_ids = [row['_id'] for row in remaining_rows]
        try:
            base.batch_delete_rows(table_name, remaining_ids)
        except Exception as e:
            print(f"Error deleting remaining rows: {e}")
    else:
        print(f"Table '{table_name}' has been successfully cleared.")

def process_excel_data(df, field_mappings, data_types):
    """处理Excel数据，转换为SeaTable格式"""
    processed_data = []
    
    for _, row in df.iterrows():
        row_data = {}
        for excel_col, seatable_field in field_mappings.items():
            value = row[excel_col]
            
            # 处理数据类型
            if seatable_field in data_types:
                if data_types[seatable_field] == 'number':
                    # 处理数字格式
                    if pd.notna(value):
                        if isinstance(value, (int, float)):
                            value = f"{value:,.2f}"
                        else:
                            try:
                                value = f"{float(value):,.2f}"
                            except:
                                value = str(value)
                elif data_types[seatable_field] == 'date':
                    # 处理日期格式
                    if pd.notna(value):
                        if isinstance(value, datetime):
                            value = value.strftime('%Y-%m-%d')
                        else:
                            try:
                                value = pd.to_datetime(value).strftime('%Y-%m-%d')
                            except:
                                value = str(value)
            
            # 处理空值
            if pd.isna(value):
                value = ""
            
            row_data[seatable_field] = value
        
        processed_data.append(row_data)
    
    return processed_data

def insert_data_into_seatable(base, data, table_name, chunk_size):
    """将数据分批插入到SeaTable"""
    for i in range(0, len(data), chunk_size):
        chunk = data[i:i + chunk_size]
        try:
            base.batch_append_rows(table_name, chunk)
        except Exception as e:
            print(f"Error inserting chunk: {e}")
            continue

def sync_table(base, table_config, excel_file):
    """同步单个表格的数据"""
    table_name = table_config['seatable']['table_name']
    sheet_name = table_config['excel_sheet']
    start_row = table_config['start_row']
    
    print(f"\n开始同步表格: {table_name}")
    print(f"Excel工作表: {sheet_name}")
    
    # 清空表格
    clear_table(base, table_name)
    
    try:
        # 读取Excel文件
        print(f"Reading Excel sheet: {sheet_name}")
        df = pd.read_excel(
            excel_file,
            sheet_name=sheet_name,
            skiprows=start_row - 1
        )
        
        # 处理数据
        print("Processing data...")
        processed_data = process_excel_data(
            df, 
            table_config['field_mappings'],
            table_config.get('data_types', {})
        )
        
        # 插入数据到SeaTable
        print(f"Inserting {len(processed_data)} rows into SeaTable...")
        insert_data_into_seatable(base, processed_data, table_name, chunk_size)
        
        print(f"Table '{table_name}' sync completed successfully!")
        
    except Exception as e:
        print(f"Error during sync for table '{table_name}': {e}")

def sync_excel():
    """主同步函数"""
    # 连接到SeaTable
    base = Base(seatable_config['api_token'], seatable_config['server_url'])
    base.auth()
    #Seatable需采用API网关，否则无法找到指定的表
    #base.use_api_gateway = False

    print("SeaTable配置：", seatable_config)
    try:
        metadata = base.get_metadata()
        print("可用表：", [t['name'] for t in metadata['tables']])
    except Exception as e:
        print("获取可用表失败：", e)
        import traceback
        traceback.print_exc()

    excel_file = config['excel_config']['file_path']

    # 同步每个表格
    for table_config in tables_config:
        print("当前同步表名：", table_config['seatable']['table_name'])
        sync_table(base, table_config, excel_file)
    
    print("\n所有表格同步完成！")

def get_config_name_from_file(config_file):
    """从配置文件名生成友好的显示名称"""
    # 移除.json扩展名
    name = os.path.splitext(config_file)[0]
    
    # 定义名称映射规则
    name_mappings = {
        'memo-bh-gov': '博浩政企数据同步',
        'memo-bh-star': '博浩卫星数据同步', 
        'memo-bh-ywl': '博浩云未来数据同步',
        'memo-bh-yxd': '博浩云现代数据同步'
    }
    
    # 如果有预定义的映射，使用映射名称，否则使用文件名
    return name_mappings.get(name, name.replace('-', ' ').replace('_', ' ').title())

def get_token_env_from_file(config_file):
    """从配置文件名生成对应的环境变量名"""
    # 移除.json扩展名
    name = os.path.splitext(config_file)[0]
    
    # 统一使用自动化规则生成环境变量名
    # 格式: SEATABLE_{NAME}_TOKEN
    env_name = name.upper().replace('-', '_').replace(' ', '_')
    return f"SEATABLE_{env_name}_TOKEN"

def discover_config_files():
    """自动发现目录下的JSON配置文件"""
    # 查找当前目录下所有.json文件
    json_files = glob.glob("*.json")
    
    # 过滤掉可能不是配置文件的json文件（比如package.json等）
    config_files = []
    for json_file in json_files:
        try:
            # 尝试读取文件，检查是否包含必要的配置字段
            with open(json_file, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
                
            # 检查是否包含必要的配置字段
            if 'tables' in config_data and 'excel_config' in config_data:
                config_files.append(json_file)
        except (json.JSONDecodeError, KeyError, FileNotFoundError):
            # 如果不是有效的配置文件，跳过
            continue
    
    return sorted(config_files)

def select_configuration():
    """选择配置文件"""
    print("\n===== Excel同步任务选择 =====")
    
    # 自动发现配置文件
    config_files = discover_config_files()
    
    if not config_files:
        print("错误：未找到有效的JSON配置文件")
        print("请确保当前目录下包含有效的配置文件（*.json）")
        sys.exit(1)
    
    # 构建配置选项
    config_options = {}
    
    # 添加配置文件选项
    for i, config_file in enumerate(config_files, 1):
        config_name = get_config_name_from_file(config_file)
        token_env = get_token_env_from_file(config_file)
        
        config_options[i] = {
            "name": config_name,
            "config_file": config_file,
            "api_token_env": token_env
        }
    
    # 添加退出选项
    config_options[0] = {
        "name": "退出程序",
        "config_file": None,
        "api_token_env": None
    }
    
    # 显示选项
    print(f"\n发现 {len(config_files)} 个配置文件：")
    for key, value in config_options.items():
        if key == 0:
            print(f"\n{key}. {value['name']}")
        else:
            print(f"{key}. {value['name']} ({value['config_file']})")
    
    # 获取用户选择
    while True:
        try:
            default_choice = 1 if config_files else 0
            choice_input = input(f"\n请选择要执行的同步任务 ({default_choice}): ").strip()
            
            # 如果用户直接按回车，使用默认选择
            if not choice_input:
                choice = default_choice
            else:
                choice = int(choice_input)
                
            if choice in config_options:
                break
            else:
                print("无效选择，请输入有效的数字")
        except ValueError:
            print("请输入有效的数字")
        except KeyboardInterrupt:
            print("\n\n程序已被用户中断")
            sys.exit(0)
    
    # 检查是否选择退出
    if choice == 0:
        print("\n程序已退出")
        sys.exit(0)
    
    selected_config = config_options[choice]
    print(f"\n已选择: {selected_config['name']}")
    print(f"配置文件: {selected_config['config_file']}")
    print(f"环境变量: {selected_config['api_token_env']}")
    
    return selected_config

if __name__ == '__main__':
    # 选择配置
    selected_config = select_configuration()
    
    # 加载配置
    load_config(selected_config['config_file'])
    
    # 动态设置 token
    api_token = os.getenv(selected_config['api_token_env'])
    if not api_token:
        print(f"错误：未找到环境变量 {selected_config['api_token_env']}")
        print("请确保 .env 文件中包含正确的 token")
        exit(1)
    seatable_config['api_token'] = api_token
    
    # 执行同步
    sync_excel() 