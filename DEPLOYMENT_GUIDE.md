# PPT自动生成系统部署指南

## 项目概述

这是一个基于Flask的PPT自动生成系统，支持通过文生文API生成结构化PPT内容，并可导出为PPTX或TXT格式。本指南将帮助您将应用部署到指定服务器。

## 部署准备

已成功创建部署包，包含以下内容：
- 主应用文件：`app.py`
- 配置文件：`config.ini`
- 依赖文件：`requirements.txt`
- 模板目录：`templates/`
- 文件存储目录：`ppt_files/` 和 `uploads/`
- 启动脚本：`start_server.bat`

## 部署步骤

### 1. 将部署包上传到目标服务器

使用您常用的文件传输工具（如FTP、SFTP、WinSCP等）将`deployment_package`目录上传到目标服务器（192.144.142.60）。

推荐上传路径：`/path/to/deployment/deployment_package`

### 2. 安装Python环境

确保目标服务器已安装Python 3.6或更高版本：

```bash
# 检查Python版本
python --version
```

如果未安装Python，请先安装Python。

### 3. 安装依赖

进入部署包目录，安装必要的Python依赖：

```bash
cd /path/to/deployment/deployment_package
pip install -r requirements.txt
```

### 4. 配置服务器

配置文件`config.ini`已预先设置好，包含以下配置：

```ini
[server]
host = 0.0.0.0  # 监听所有网络接口
port = 5001     # 服务端口
secret_key = oeasy  # 您提供的密钥

[api]
base_url = https://api-inference.modelscope.cn/v1  # 文生文API地址
api_key = ms-4b457ae8-4cfd-4504-8ec2-8dc2fb930454  # API密钥
```

如需修改配置，请直接编辑`config.ini`文件。

### 5. 启动服务

#### Windows服务器：

双击运行`start_server.bat`脚本，或在命令行中执行：

```cmd
start_server.bat
```

#### Linux服务器：

使用Python直接启动应用：

```bash
python app.py
```

### 6. 访问应用

服务启动后，可通过以下地址访问：
- 服务器本地：`http://127.0.0.1:5001`
- 网络访问：`http://192.144.142.60:5001`

## 注意事项

1. **防火墙配置**：确保目标服务器的防火墙已开放5001端口。

2. **生产环境建议**：对于生产环境部署，建议使用Gunicorn或uWSGI作为WSGI服务器，并配置Nginx作为反向代理。

3. **持久化运行**：
   - Windows：可使用任务计划程序设置开机自启
   - Linux：可使用systemd服务或nohup命令

4. **日志管理**：应用运行日志默认输出到控制台，生产环境建议配置日志文件。

## 故障排除

### 常见问题

1. **服务无法启动**：
   - 检查端口是否被占用：`netstat -ano | findstr :5001`
   - 检查Python依赖是否安装成功
   - 查看错误日志获取详细信息

2. **API调用失败**：
   - 检查网络连接是否正常
   - 验证API密钥是否正确
   - 查看应用日志中的错误信息

3. **PPT生成失败**：
   - 确保已安装python-pptx：`pip install python-pptx`
   - 检查文件权限是否正确

## 技术支持

如有任何部署或运行问题，请根据应用日志中的错误信息进行排查，或联系技术支持。

---

**部署信息**
- 目标服务器：192.144.142.60
- 服务端口：5001
- 密钥：oeasy
- 部署包路径：deployment_package