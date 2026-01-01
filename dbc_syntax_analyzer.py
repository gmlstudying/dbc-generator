#!/usr/bin/env python3
"""
DBC文件语法分析工具
用于详细分析DBC文件的语法结构，找出所有可能的语法错误
"""

import re

def analyze_dbc_file(dbc_file):
    """分析DBC文件语法"""
    print(f"分析DBC文件: {dbc_file}")
    print("=" * 50)
    
    with open(dbc_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 1. 检查文件结构
    print("1. 检查文件结构")
    required_sections = ['VERSION', 'NS_', 'BS_', 'BU_']
    for section in required_sections:
        if section not in content:
            print(f"   - ✗ 缺少必要的部分: {section}")
        else:
            print(f"   - ✓ 找到部分: {section}")
    
    # 2. 检查BO_行
    print("\n2. 检查BO_行")
    bo_pattern = r'^BO_\s+(\d+)\s+(\w+)\s*:\s*(\d+)\s+(\w+)\s*$'
    bo_matches = re.findall(bo_pattern, content, re.MULTILINE)
    print(f"   - 找到 {len(bo_matches)} 个消息定义")
    for match in bo_matches:
        print(f"     - BO_ {match[0]} {match[1]}: {match[2]} {match[3]}")
    
    # 3. 检查SG_行
    print("\n3. 检查SG_行")
    sg_pattern = r'^\s*SG_\s+(\w+)\s*:\s*(\d+)\|(\d+)@(\d)([+-])\s+\(([^)]+)\)\s+\[([^\]]+)\]\s+"([^"]*)"\s+([\w\s]+)\s*$'
    sg_matches = re.findall(sg_pattern, content, re.MULTILINE)
    print(f"   - 找到 {len(sg_matches)} 个信号定义")
    for match in sg_matches:
        print(f"     - SG_ {match[0]}: {match[1]}|{match[2]}@{match[3]}{match[4]} ({match[5]}) [{match[6]}] \"{match[7]}\" {match[8]}")
    
    # 4. 检查BU_行
    print("\n4. 检查BU_行")
    bu_pattern = r'^BU_:\s+(.+)\s*$'
    bu_match = re.search(bu_pattern, content, re.MULTILINE)
    if bu_match:
        nodes = bu_match.group(1).split()
        print(f"   - 找到 {len(nodes)} 个节点定义: {nodes}")
    else:
        print("   - ✗ 缺少BU_行")
    
    # 5. 查找可能的语法错误
    print("\n5. 查找可能的语法错误")
    lines = content.split('\n')
    
    # 检查SG_行的具体格式
    for i, line in enumerate(lines):
        stripped_line = line.strip()
        # 确保只检查真正的信号定义行，而不是NS_部分的关键字
        if stripped_line.startswith('SG_') and not stripped_line.startswith('SG_MUL_VAL_') and not stripped_line.startswith('SGTYPE_') and not stripped_line.startswith('SIGTYPE_'):
            # 检查信号定义格式
            if not re.match(sg_pattern, line):
                print(f"   - ✗ 第{i+1}行SG_格式错误: {line.strip()}")
                # 详细分析错误
                if '@' not in line:
                    print(f"     - 缺少@符号")
                elif '(' not in line:
                    print(f"     - 缺少(符号")
                elif ')' not in line:
                    print(f"     - 缺少)符号")
                elif '[' not in line:
                    print(f"     - 缺少[符号")
                elif ']' not in line:
                    print(f"     - 缺少]符号")
                elif '"' not in line:
                    print(f"     - 缺少\"符号")
    
    print("\n" + "=" * 50)
    print("分析完成")

if __name__ == '__main__':
    import sys
    if len(sys.argv) != 2:
        print(f"用法: {sys.argv[0]} <dbc_file>")
        sys.exit(1)
    
    analyze_dbc_file(sys.argv[1])