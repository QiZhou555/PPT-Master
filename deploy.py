#!/usr/bin/env python3
# -*- coding: utf-8 -*-
'''
部署脚本 - 用于将应用部署到指定服务器
'''

import os
import sys
import argparse
import subprocess
import shutil
import logging

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger('deploy')

# 定义需要部署的文件和目录
DEPLOY_FILES = [
    'app.py',
    'config.ini',
    'requirements.txt',
    'templates',
    'ppt_files',
    'uploads'
]

def validate_source():
    """验证源文件是否存在"""
    missing_files = []
    for item in DEPLOY_FILES:
        if not os.path.exists(item):
            if not os.path.isdir(item):
                missing_files.append(item)
    
    if missing_files:
        logger.error(f"缺少必要的文件或目录: {', '.join(missing_files)}")
        return False
    return True

def create_deployment_package():
    """创建部署包"""
    package_dir = 'deployment_package'
    
    # 如果包目录存在，删除它
    if os.path.exists(package_dir):
        shutil.rmtree(package_dir)
    
    # 创建包目录
    os.makedirs(package_dir)
    
    # 复制文件和目录
    for item in DEPLOY_FILES:
        dest_path = os.path.join(package_dir, os.path.basename(item))
        if os.path.isfile(item):
            shutil.copy2(item, dest_path)
            logger.info(f"复制文件: {item} -> {dest_path}")
        elif os.path.isdir(item):
            # 确保目标目录不存在
            if os.path.exists(dest_path):
                shutil.rmtree(dest_path)
            os.makedirs(dest_path, exist_ok=True)
            for root, dirs, files in os.walk(item):
                # 计算相对路径
                rel_path = os.path.relpath(root, item)
                if rel_path != '.':
                    target_dir = os.path.join(dest_path, rel_path)
                else:
                    target_dir = dest_path
                
                # 创建目标目录
                os.makedirs(target_dir, exist_ok=True)
                
                # 复制文件
                for file in files:
                    src_file = os.path.join(root, file)
                    dst_file = os.path.join(target_dir, file)
                    shutil.copy2(src_file, dst_file)
            logger.info(f"复制目录: {item} -> {dest_path}")
    
    # 创建启动脚本
    start_script = os.path.join(package_dir, 'start_server.bat')
    with open(start_script, 'w', encoding='utf-8') as f:
        f.write('''@echo off
set FLASK_SECRET_KEY=oeasy
set API_KEY=ms-4b457ae8-4cfd-4504-8ec2-8dc2fb930454
set API_BASE_URL=https://api-inference.modelscope.cn/v1
set FLASK_HOST=0.0.0.0
set FLASK_PORT=5001

echo 正在启动PPT生成服务器...
echo 请确保已安装必要的依赖: pip install -r requirements.txt
python app.py
pause
''')
    
    logger.info(f"创建启动脚本: {start_script}")
    logger.info(f"部署包已创建: {package_dir}")
    return package_dir

def deploy_to_server(package_dir, server_url, server_username=None, server_password=None):
    """部署到远程服务器"""
    logger.info(f"准备部署到服务器: {server_url}")
    logger.info("注意: 以下是部署说明，您需要手动执行或根据实际情况调整:")
    logger.info("\n部署步骤:")
    logger.info("1. 将 'deployment_package' 目录复制到目标服务器")
    logger.info("2. 在目标服务器上安装依赖: pip install -r requirements.txt")
    logger.info("3. 执行启动脚本: start_server.bat (Windows) 或运行: python app.py")
    logger.info("4. 访问应用: http://<服务器IP>:5001")
    logger.info("\n如果需要通过SSH/SFTP上传，可以使用以下命令:")
    if server_username:
        logger.info(f"scp -r {package_dir} {server_username}@{server_url}:/path/to/deployment/")
    else:
        logger.info(f"使用文件传输工具将 {package_dir} 目录上传到服务器")

def main():
    parser = argparse.ArgumentParser(description='部署PPT生成应用到服务器')
    parser.add_argument('--server', type=str, default='192.144.142.60', help='目标服务器地址')
    parser.add_argument('--username', type=str, help='服务器用户名')
    parser.add_argument('--password', type=str, help='服务器密码')
    args = parser.parse_args()
    
    # 验证源文件
    if not validate_source():
        sys.exit(1)
    
    # 创建部署包
    package_dir = create_deployment_package()
    
    # 部署到服务器
    deploy_to_server(package_dir, args.server, args.username, args.password)
    
    logger.info("\n部署准备完成！请按照上述说明将应用部署到目标服务器。")

if __name__ == '__main__':
    main()