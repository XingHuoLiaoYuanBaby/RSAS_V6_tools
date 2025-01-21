# RSAS V6 离线提取端口扫描报告

## 功能说明
本工具用于处理绿盟 RSAS V6 设备的漏洞扫描报告，提取其中所有的端口和指纹信息。

漏扫导出的包为excel包，且包含综述和主机报表。

任意报表模板均可。

## 主要功能
- 自动处理 pending 目录下的所有符合命名规则的 ZIP 文件
- 提取每个主机的基本信息（IP、主机名、操作系统、扫描时间）
- 提取所有开放端口信息（端口号、协议、服务、状态）
- 支持处理端口范围（如：80-89）
- 生成统一格式的 Excel 报告

## 注意事项
1. pending于脚本同一级目录下，RSAS的xls漏扫结果放在pending文件夹下，脚本会自动删除index.xls
2. 导出数据的时候 导出xls格式 带主机报表 如果同时导出多个文件选批量导出不要选合并导出
3. 正常导出来的数据 文件名是 任务序号_任务名称_导出时间_xls.zip
4. 直接把zip文件放到pending文件夹里然后运行脚本就行了
5. 不需要再解压
6. 导出的数据表中null的为不存在开放端口

## 使用说明

### 1. 文件准备
1. 将需要处理的 ZIP 文件放入 `pending` 目录（首次运行会自动创建）
2. ZIP 文件命名规则：`数字_名称_年_月_日_xls.zip` 或 `数字_名称_年_月_日_excel.zip`
   - 示例：`1_扫描任务_2024_01_25_xls.zip`

### 2. 运行程序
- 直接运行 exe 文件即可
- 程序会自动处理 pending 目录下的所有符合命名规则的文件
- 处理完成的 Excel 文件将保存在同一目录下

### 3. 输出结果
生成的 Excel 文件包含以下字段：
- taskid：任务ID
- ip：主机IP地址
- hostname：主机名
- system_type：操作系统类型
- scan_time：扫描时间
- port：端口号
- protocol：协议
- service：服务名称
- status：状态
- ip:port：IP和端口组合

## 开发环境
- Python 3.12
- 依赖包：
  - openpyxl==3.1.2
  - xlrd
  - colorama

## 打包说明

### 1. 环境准备

####  安装必要的包
pip install pyinstaller
pip install openpyxl==3.1.2
pip install xlrd
pip install colorama


### 2. 打包步骤
pyinstaller --name "RSAS端口扫描报告提取工具" --onefile --collect-all openpyxl RSAS_V6设备离线提取端口扫描报告_ver20250121.py
pyinstaller --name "IP地址段统计工具" --onefile --windowed --clean ip_asset_check.py