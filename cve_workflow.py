"""
CVE数据抓取与处理工具
作者：Cursor AI 协助开发
功能：自动抓取指定日期范围内的CVE漏洞数据，并生成格式化的Excel报告
"""

import requests  # 用于发送HTTP请求
import json  # 用于处理JSON数据
import re  # 用于正则表达式处理
import pandas as pd  # 用于数据处理和Excel生成
from openpyxl import load_workbook  # 用于Excel文件操作
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side  # 用于Excel样式设置
from datetime import datetime  # 用于日期处理
import brotli  # 用于解压缩响应数据

def get_date_input():
    """
    获取用户输入的日期范围
    返回：
        tuple: (开始日期, 结束日期)
    """
    while True:
        try:
            # 获取开始日期
            start_date = input("请输入开始日期 (格式: YYYY-MM-DD): ")
            datetime.strptime(start_date, '%Y-%m-%d')
            
            # 获取结束日期
            end_date = input("请输入结束日期 (格式: YYYY-MM-DD): ")
            datetime.strptime(end_date, '%Y-%m-%d')
            
            # 验证结束日期是否晚于开始日期
            if datetime.strptime(end_date, '%Y-%m-%d') < datetime.strptime(start_date, '%Y-%m-%d'):
                print("错误：结束日期不能早于开始日期，请重新输入")
                continue
                
            return start_date, end_date
            
        except ValueError:
            print("错误：日期格式不正确，请使用 YYYY-MM-DD 格式")
            continue

def is_valid_component_match(text, component_name):
    """
    验证文本中是否真正包含目标组件（使用单词边界匹配）
    参数：
        text: 要检查的文本
        component_name: 组件名称
    返回：
        bool: 是否为有效匹配
    """
    # 转换为小写进行比较
    text_lower = text.lower()
    component_lower = component_name.lower()
    
    # 使用正则表达式进行单词边界匹配
    # \b 表示单词边界，确保匹配的是完整单词而不是子字符串
    pattern = r'\b' + re.escape(component_lower) + r'\b'
    
    return bool(re.search(pattern, text_lower))

def get_component_aliases(component_name):
    """
    获取组件的常见别名和相关搜索词
    参数：
        component_name: 基础组件名称
    返回：
        list: 包含原名称和别名的列表
    """
    # 定义常见组件的别名映射
    alias_map = {
        'poi': ['apache poi', 'poi library'],
        'log4j': ['apache log4j', 'log4j2'],
        'spring': ['spring framework', 'springframework'],
        'struts': ['apache struts', 'struts2'],
        'jackson': ['jackson-databind', 'fasterxml jackson'],
        'gson': ['google gson'],
        'fastjson': ['alibaba fastjson'],
        'mysql': ['mysql connector', 'mysql-connector'],
        'redis': ['redis server'],
        'nginx': ['nginx server'],
        'apache': ['apache httpd', 'apache server'],
        'tomcat': ['apache tomcat'],
        'elasticsearch': ['elastic elasticsearch'],
        'kibana': ['elastic kibana'],
        'mongodb': ['mongo database'],
        'postgresql': ['postgres', 'postgresql database'],
        'node': ['node.js', 'nodejs'],
        'react': ['react.js', 'reactjs'],
        'vue': ['vue.js', 'vuejs'],
        'angular': ['angular.js', 'angularjs'],
        'jquery': ['jquery library'],
        'bootstrap': ['bootstrap framework'],
        'django': ['django framework'],
        'flask': ['flask framework'],
        'express': ['express.js', 'expressjs'],
        'laravel': ['laravel framework'],
        'symfony': ['symfony framework'],
        'wordpress': ['wordpress cms'],
        'drupal': ['drupal cms'],
        'joomla': ['joomla cms']
    }
    
    component_lower = component_name.lower()
    aliases = alias_map.get(component_lower, [])
    
    # 始终包含原始名称
    return [component_name] + aliases

def validate_cve_relevance(cve_data, component_name):
    """
    验证CVE数据是否真正与目标组件相关
    参数：
        cve_data: CVE数据字典
        component_name: 目标组件名称
    返回：
        bool: 是否相关
    """
    # 检查的字段列表
    fields_to_check = [
        cve_data.get('description', ''),
        cve_data.get('id', ''),
        ' '.join(cve_data.get('references', [])) if cve_data.get('references') else ''
    ]
    
    # 获取组件的所有可能名称（包括别名）
    all_names = get_component_aliases(component_name)
    
    # 检查是否有任何一个名称在任何字段中匹配
    for field_text in fields_to_check:
        if field_text:
            for name in all_names:
                if is_valid_component_match(field_text, name):
                    return True
    
    return False

def read_components_with_config(filename):
    """
    从文本文件中读取组件列表，支持配置格式
    支持格式：
    1. 简单格式：每行一个组件名
    2. 配置格式：组件名:搜索词1,搜索词2
    
    参数：
        filename: 组件列表文件名
    返回：
        list: 组件配置列表，每个元素包含{'name': '组件名', 'search_terms': [搜索词列表]}
    """
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            lines = [line.strip() for line in f.readlines()]
            lines = [line for line in lines if line and not line.startswith('#')]  # 过滤空行和注释
        
        components = []
        for line in lines:
            if ':' in line:
                # 配置格式：组件名:搜索词1,搜索词2
                name, search_terms = line.split(':', 1)
                name = name.strip()
                search_terms = [term.strip() for term in search_terms.split(',')]
            else:
                # 简单格式：只有组件名
                name = line
                search_terms = get_component_aliases(name)
            
            components.append({
                'name': name,
                'search_terms': search_terms
            })
        
        return components
    except Exception as e:
        print(f"读取组件文件出错: {e}")
        return []

def fetch_cve_data(component, start_date, end_date):
    """
    从cvefeed.io获取指定组件的CVE数据，使用改进的搜索策略
    参数：
        component: 组件名称
        start_date: 开始日期
        end_date: 结束日期
    返回：
        dict: CVE数据字典，包含组件名称和漏洞列表
    """
    # 获取组件的搜索别名
    search_terms = get_component_aliases(component)
    
    # 优先使用更具体的搜索词（通常别名更准确）
    primary_search_term = search_terms[1] if len(search_terms) > 1 else search_terms[0]
    
    # 设置请求URL和参数
    url = "https://cvefeed.io/api/advanced-search"
    params = {
        "keyword": primary_search_term,  # 使用更精确的搜索词
        "published_after": start_date,
        "published_before": end_date,
        "last_modified_after": start_date,
        "last_modified_before": end_date,
        "cvss_min": "7.00",  # 只获取CVSS评分大于等于7.0的漏洞
        "cvss_max": "10.00",  # 只获取CVSS评分小于等于10.0的漏洞
        "order_by": "-published"  # 按发布时间降序排序
    }

    # 设置请求头，模拟浏览器请求
    headers = {
        "Host": "cvefeed.io",
        "Connection": "keep-alive",
        "sec-ch-ua-platform": "Windows",
        "X-Requested-With": "XMLHttpRequest",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36",
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "sec-ch-ua": '"Google Chrome";v="135", "Not-A.Brand";v="8", "Chromium";v="135"',
        "sec-ch-ua-mobile": "?0",
        "Sec-Fetch-Site": "same-origin",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Dest": "empty",
        "Accept-Encoding": "gzip, deflate, br",
        "Accept-Language": "zh-CN,zh;q=0.9"
    }

    try:
        print(f"\n正在获取组件 {component} 的CVE数据...")
        print(f"使用搜索词: {primary_search_term}")
        # 发送GET请求
        response = requests.get(url, params=params, headers=headers)
        
        # 打印调试信息
        print(f"请求URL: {response.url}")
        print(f"状态码: {response.status_code}")
        
        # 尝试使用brotli解压缩
        try:
            content = brotli.decompress(response.content).decode('utf-8')
            print("使用brotli解压缩成功")
        except:
            # 如果brotli解压缩失败，尝试直接解码
            content = response.content.decode('utf-8')
            print("使用直接解码")
        
        # 检查响应状态
        response.raise_for_status()
        
        # 解析JSON响应
        data = json.loads(content)
        
        # 添加组件名称到结果中
        data['component'] = component
        
        return data
        
    except requests.exceptions.RequestException as e:
        print(f"请求出错: {e}")
        return None
    except json.JSONDecodeError as e:
        print(f"JSON解析错误: {e}")
        print("原始响应内容:")
        print(content)
        return None
    except Exception as e:
        print(f"其他错误: {e}")
        return None

def clean_html_tags(text):
    """
    清理文本中的HTML标签
    参数：
        text: 包含HTML标签的文本
    返回：
        str: 清理后的纯文本
    """
    clean = re.compile('<.*?>')
    return re.sub(clean, '', text)

def clean_description(text):
    """
    清理漏洞描述文本，移除多余空白和换行
    参数：
        text: 原始描述文本
    返回：
        str: 清理后的描述文本
    """
    # 移除HTML标签
    text = clean_html_tags(text)
    # 替换所有空白字符为单个空格
    text = re.sub(r'\s+', ' ', text)
    # 移除首尾空格
    return text.strip()

def process_cve_data(json_data):
    """
    处理CVE数据，提取关键信息并进行相关性验证
    参数：
        json_data: 原始JSON数据
    返回：
        list: 处理后的CVE数据列表
    """
    results = []
    
    for component_data in json_data:
        component = component_data['component']
        cve_entries = []
        
        print(f"\n正在处理组件 {component} 的漏洞数据...")
        original_count = len(component_data['results'])
        filtered_count = 0
        
        for cve in component_data['results']:
            # 验证CVE是否真正与目标组件相关
            if validate_cve_relevance(cve, component):
                cve_entry = {
                    'CVE编号': cve['id'],
                    '发布时间': cve['published'].split('T')[0],
                    '漏洞描述': clean_description(cve['description'])
                }
                cve_entries.append(cve_entry)
                filtered_count += 1
            else:
                print(f"过滤掉不相关的漏洞: {cve['id']}")
        
        print(f"组件 {component}: 原始漏洞 {original_count} 个，过滤后 {filtered_count} 个")
        
        # 只在有有效漏洞时才添加结果
        if cve_entries:
            result = {
                '组件名称': component,
                '漏洞列表': cve_entries
            }
            results.append(result)
    
    return results

def format_worksheet(ws, data_length):
    """
    设置Excel工作表格式
    参数：
        ws: Excel工作表对象
        data_length: 数据行数
    """
    # 设置列宽
    ws.column_dimensions['A'].width = 15  # 组件名称列
    ws.column_dimensions['B'].width = 20  # CVE编号列
    ws.column_dimensions['C'].width = 15  # 发布时间列
    ws.column_dimensions['D'].width = 100  # 漏洞描述列

    # 设置标题行格式
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")  # 蓝色背景
    header_font = Font(name='微软雅黑', size=11, bold=True, color="FFFFFF")  # 白色字体
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # 设置标题行
    headers = ['组件名称', 'CVE编号', '发布时间', '漏洞描述']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    # 设置数据行格式
    data_font = Font(name='微软雅黑', size=10)
    for row in range(2, data_length + 2):
        for col in range(1, 5):
            cell = ws.cell(row=row, column=col)
            cell.font = data_font
            cell.border = thin_border
            # 为漏洞描述列设置自动换行
            if col == 4:
                cell.alignment = Alignment(wrapText=True, vertical='center')
            else:
                cell.alignment = Alignment(horizontal='center', vertical='center')

def merge_component_cells(ws, data_length):
    """
    合并Excel中相同组件名称的单元格
    参数：
        ws: Excel工作表对象
        data_length: 数据行数
    """
    current_component = None
    start_row = 2
    
    for row in range(2, data_length + 2):
        component = ws.cell(row=row, column=1).value
        
        if current_component is None:
            current_component = component
            start_row = row
            continue
            
        if component != current_component or row == data_length + 1:
            if start_row < row:
                # 合并单元格
                ws.merge_cells(f'A{start_row}:A{row-1}')
                merged_cell = ws.cell(row=start_row, column=1)
                merged_cell.alignment = Alignment(horizontal='center', vertical='center')
            
            current_component = component
            start_row = row
    
    # 处理最后一组
    if start_row < data_length + 1:
        ws.merge_cells(f'A{start_row}:A{data_length+1}')
        merged_cell = ws.cell(row=start_row, column=1)
        merged_cell.alignment = Alignment(horizontal='center', vertical='center')

def convert_to_excel(data):
    """
    将处理后的数据转换为Excel格式
    参数：
        data: 处理后的CVE数据
    """
    # 创建数据列表
    rows = []
    for component in data:
        component_name = component['组件名称']
        for vuln in component['漏洞列表']:
            rows.append({
                '组件名称': component_name,
                'CVE编号': vuln['CVE编号'],
                '发布时间': vuln['发布时间'],
                '漏洞描述': vuln['漏洞描述']
            })
    
    if not rows:
        print("警告：没有找到任何数据")
        return
    
    # 创建DataFrame
    df = pd.DataFrame(rows)
    
    # 保存为Excel文件
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_file = f'cve_results_{timestamp}.xlsx'
    df.to_excel(excel_file, index=False)
    
    # 加载工作簿进行格式调整
    wb = load_workbook(excel_file)
    ws = wb.active
    
    # 设置工作表格式
    format_worksheet(ws, len(rows))
    
    # 合并组件名称列的相同单元格
    merge_component_cells(ws, len(rows))
    
    # 调整行高
    for row in range(1, len(rows) + 2):
        ws.row_dimensions[row].height = 30
    
    # 保存修改后的Excel文件
    wb.save(excel_file)
    print(f"结果已成功保存到 {excel_file} 文件中")

def main():
    """
    主函数，程序入口
    """
    print("欢迎使用CVE数据抓取和处理工具")
    print("=" * 50)
    
    # 获取用户输入的日期
    start_date, end_date = get_date_input()
    print(f"\n已选择日期范围: {start_date} 至 {end_date}")
    
    # 读取组件列表
    components = read_components_with_config('components.txt')
    if not components:
        print("没有找到任何组件名称")
        return
    
    print(f"\n找到 {len(components)} 个组件")
    print("组件列表:", [comp['name'] for comp in components])
    
    # 确认是否继续
    confirm = input("\n是否开始抓取数据？(y/n): ")
    if confirm.lower() != 'y':
        print("操作已取消")
        return
    
    # 存储所有结果
    all_results = []
    
    # 依次处理每个组件
    for component in components:
        result = fetch_cve_data(component['name'], start_date, end_date)
        if result:
            all_results.append(result)
            print(f"成功获取组件 {component['name']} 的数据")
        else:
            print(f"获取组件 {component['name']} 的数据失败")
    
    if not all_results:
        print("\n没有成功获取任何数据")
        return
    
    # 处理数据并转换为Excel
    print("\n开始处理数据并生成Excel文件...")
    processed_data = process_cve_data(all_results)
    convert_to_excel(processed_data)

if __name__ == "__main__":
    main() 