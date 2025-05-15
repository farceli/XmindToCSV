import json
import zipfile
import os
import pandas as pd
import shutil
import re

# pip install pandas

# 解压 .xmind 文件
def extract_content_json(xmind_file, extract_to='xmind_temp'):
    with zipfile.ZipFile(xmind_file, 'r') as zip_ref:
        zip_ref.extractall(extract_to)
    return os.path.join(extract_to, 'content.json')

# 从 marker 中提取优先级
def extract_priority(topic):
    markers = topic.get('markers', [])
    for marker in markers:
        if isinstance(marker, str):
            target = marker
        elif isinstance(marker, dict) and 'markerId' in marker:
            target = marker['markerId']
        else:
            continue
        match = re.match(r'priority-(\d)', target)
        if match:
            return match.group(1)
    return ''

# 递归解析节点，路径中每一层记录标题和优先级
def parse_topic(topic, path=None, result=None):
    if path is None:
        path = []
    if result is None:
        result = []

    title = topic.get('title', '')
    note = topic.get('notes', {}).get('plain', {}).get('content', '')
    priority = extract_priority(topic)
    current_path = path + [{'title': title, 'priority': priority}]

    # 只在叶子节点记录
    if not topic.get('children', {}).get('attached'):
        result.append({
            'path': current_path,
            'note': note
        })
    else:
        for child in topic['children']['attached']:
            parse_topic(child, current_path, result)

    return result

# 根据映射配置生成 CSV 每一行
def map_to_columns(parsed_data, column_mapping):
    rows = []
    for item in parsed_data:
        row = {}
        path = item['path']
        for col_name, config in column_mapping.items():
            if isinstance(config, dict):
                # ✅ 新增固定值支持
                if 'value' in config:
                    row[col_name] = config['value']
                else:
                    level = config.get('level')
                    if level is not None and level - 1 < len(path):
                        node = path[level - 1]
                        if config.get('priority'):
                            row[col_name] = node.get('priority', '')
                        else:
                            row[col_name] = node.get('title', '')
                    else:
                        row[col_name] = ''
            elif config == 'note':
                row[col_name] = item.get('note', '')
            else:
                row[col_name] = ''
        rows.append(row)
    return rows

# 写入 CSV
def save_to_csv(rows, csv_file_path, column_headers):
    df = pd.DataFrame(rows, columns=column_headers)
    df.to_csv(csv_file_path, index=False, encoding='utf-8-sig')
    print(f"✅ CSV 已生成：{csv_file_path}")

# 主函数
def main():
    xmind_file = 'test_points.xmind'  # 替换为你的 XMind 文件
    csv_file = 'test_cases.csv'       # 输出的 CSV 文件
    extract_dir = 'xmind_temp'

    # ✅ 列名 → 层级配置。支持指定优先级提取
    column_mapping = {
        '所属模块': {'level': 1},
        '用例标题': {'level': 2},
        '前置条件': {'level': 3},
        '步骤': {'level': 3},
        '预期': {'level': 3},
        '关键词': {'value': ''},
        '优先级': {'level': 3, 'priority': True},
        '用例类型': {'value': '功能测试'},
        '适用阶段': {'value': '功能测试阶段'},
    }

    column_headers = list(column_mapping.keys())

    # 解压读取
    content_json_path = extract_content_json(xmind_file, extract_dir)
    with open(content_json_path, 'r', encoding='utf-8') as f:
        content = json.load(f)

    # 解析所有 sheet
    all_data = []
    for sheet in content:
        root_topic = sheet.get('rootTopic')
        if root_topic:
            all_data.extend(parse_topic(root_topic))

    # 映射 & 导出
    mapped_rows = map_to_columns(all_data, column_mapping)
    save_to_csv(mapped_rows, csv_file, column_headers)

    shutil.rmtree(extract_dir)

if __name__ == '__main__':
    main()
