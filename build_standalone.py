#!/usr/bin/env python3
"""
独立打包脚本
创建完全自包含的可执行文件，包含所有配置文件
"""

import os
import sys
import shutil
import subprocess
import json

def create_standalone_build():
    print("====================================")
    print("   Building SeaTable Sync Tool")
    print("====================================")
    
    # 1. Install dependencies
    print("1. Installing dependencies...")
    subprocess.run([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"], check=True)
    subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
    
    # 2. Clean old files
    print("\n2. Cleaning previous build files...")
    for folder in ["dist", "build"]:
        if os.path.exists(folder):
            shutil.rmtree(folder)
    
    for file in os.listdir("."):
        if file.endswith(".spec"):
            os.remove(file)
    
    # 3. Collect config files
    config_files = []
    if os.path.exists("config"):
        config_files = [f"config/{f}" for f in os.listdir("config") if f.endswith(".json")]
    print(f"\n3. Found config files: {', '.join(config_files)}")
    
    # 4. 构建PyInstaller命令
    cmd = [
        "pyinstaller",
        "--onefile",
        "--console",
        "--name", "excel-sea-sync",
        "--noupx",  # 禁用UPX压缩，避免DLL加载问题
        "--clean",  # 清理缓存
        "--hidden-import", "seatable_api",
        "--hidden-import", "pandas",
        "--hidden-import", "numpy",
        "--hidden-import", "openpyxl",
        "--hidden-import", "dotenv",
        "--hidden-import", "json",
        "--hidden-import", "datetime"
    ]
    
    # Windows特定选项
    if sys.platform.startswith("win"):
        cmd.extend([
            "--collect-all", "seatable_api",  # 收集所有seatable_api依赖
            "--collect-all", "pandas",        # 收集所有pandas依赖
            "--collect-all", "openpyxl",      # 收集所有openpyxl依赖
            "--noconsole" if "--noconsole" in sys.argv else "--console"
        ])
    
    # 添加JSON配置文件
    json_files = [f for f in os.listdir(".") if f.endswith(".json")]
    for json_file in json_files:
        cmd.extend(["--add-data", f"{json_file}:."])
    
    # 添加.env示例文件
    if os.path.exists(".env.example"):
        cmd.extend(["--add-data", ".env.example:."])
    
    # 添加主文件
    cmd.append("excel-sync.py")
    
    print(f"\n4. Executing build command...")
    print(f"Command: {' '.join(cmd)}")
    
    try:
        subprocess.run(cmd, check=True)
        print("\n[OK] Build successful!")
    except subprocess.CalledProcessError as e:
        print(f"\n[ERROR] Build failed: {e}")
        return False
    
    # 5. Create deployment package
    print("\n5. Creating deployment package...")
    
    # Create deployment directory
    deploy_dir = "seatable-sync-deploy"
    if os.path.exists(deploy_dir):
        shutil.rmtree(deploy_dir)
    os.makedirs(deploy_dir)
    
    # Copy executable file
    exe_name = "excel-sea-sync.exe" if sys.platform.startswith("win") else "excel-sea-sync"
    src_exe = os.path.join("dist", exe_name)
    dst_exe = os.path.join(deploy_dir, exe_name)
    
    if os.path.exists(src_exe):
        shutil.copy2(src_exe, dst_exe)
        print(f"[OK] Copied executable: {exe_name}")
    else:
        print(f"[ERROR] Executable not found: {src_exe}")
        return False
    
    # Copy JSON config files
    json_files = [f for f in os.listdir(".") if f.endswith(".json")]
    for json_file in json_files:
        shutil.copy2(json_file, deploy_dir)
        print(f"[OK] Copied config file: {json_file}")
    
    # Copy .env example file
    if os.path.exists(".env.example"):
        shutil.copy2(".env.example", deploy_dir)
        print("[OK] Copied .env example file")
    
    # Copy documentation
    if os.path.exists("README.md"):
        shutil.copy2("README.md", deploy_dir)
    if os.path.exists("PREPROCESS_GUIDE.md"):
        shutil.copy2("PREPROCESS_GUIDE.md", deploy_dir)
    
    # 创建使用说明
    readme_content = """# Excel to SeaTable 同步工具部署包

## 使用步骤：

1. 配置环境变量（推荐）：
   cp .env.example .env
   编辑 .env 文件，填入你的SeaTable Token

2. 运行同步工具：
   # Windows:
   excel-sea-sync.exe
   
   # Linux/macOS:
   ./excel-sea-sync

## 配置方式：

### .env文件配置（推荐）
1. 复制 .env.example 为 .env
2. 编辑 .env 文件，填入配置信息：
   - SEATABLE_SERVER_URL=你的SeaTable服务器地址
   - SEATABLE_BH_GOV_TOKEN=博浩政企项目Token
   - SEATABLE_BH_STAR_TOKEN=博浩卫星项目Token
   - SEATABLE_BH_YWL_TOKEN=博浩云未来项目Token
   - SEATABLE_BH_YXD_TOKEN=博浩云现代项目Token
3. 直接运行: ./excel-sea-sync

## 配置文件说明：

- memo-bh-gov.json: 博浩政企数据同步配置
- memo-bh-star.json: 博浩卫星数据同步配置  
- memo-bh-ywl.json: 博浩云未来数据同步配置
- memo-bh-yxd.json: 博浩云现代数据同步配置

## 功能特性：

- 支持Excel文件多表同步到SeaTable
- 支持字段映射和数据类型转换
- 支持分批处理大量数据
- 交互式配置选择
- 跨平台支持（Windows, Linux, macOS）

## 注意事项：

- 确保网络能访问SeaTable服务
- 确保API Token有相应的表格权限
- 同步前会清空目标表格数据
- Excel文件路径在配置文件中指定
- 支持.xlsx格式Excel文件
"""
    
    with open(os.path.join(deploy_dir, "USAGE.txt"), "w", encoding="utf-8") as f:
        f.write(readme_content)
    
    print("\n====================================")
    print("[SUCCESS] Excel-SeaTable Sync tool package created successfully!")
    print(f"Package location: {deploy_dir}/")
    print(f"Executable: {deploy_dir}/{exe_name}")
    print("Share the entire folder with your team")
    print("====================================")
    
    return True

if __name__ == "__main__":
    success = create_standalone_build()
    if not success:
        sys.exit(1)