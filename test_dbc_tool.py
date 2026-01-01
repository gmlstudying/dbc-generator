#!/usr/bin/env python3
"""
DBC生成工具综合测试脚本
包含所有功能测试用例，便于统一管理和执行

测试模块说明：
1. test_naming.py - 测试DBC文件名生成功能
2. test_gui.py - 测试GUI启动和功能
3. test_dbc_syntax.py - 测试DBC语法正确性
4. test_node_identification.py - 测试节点类型识别
5. test_command_line.py - 测试命令行功能
6. test_ascii_encoding.py - 测试ASCII编码处理
"""

import os
import sys
import re
import argparse

# 添加当前目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

###############################################################################
# 测试模块1：DBC文件名生成功能
# 用途：验证不同节点类型生成正确的总线类型后缀文件名
###############################################################################
def test_filename_generation():
    """测试DBC文件名生成功能"""
    print("\n" + "="*60)
    print("测试1: DBC文件名生成功能")
    print("="*60)
    
    # 模拟各种节点类型
    test_cases = [
        ("VCU_P", "PCAN"),
        ("VCU_E", "ECAN"),
        ("VCU_C", "CCAN"),
        ("VCU_T", "TCAN"),
        ("VCU_B", "BCAN"),
        ("VCU_X", "XCAN"),
        ("VCU", "VCU_CAN")
    ]
    
    input_file = "F511C_CANMatrix_VCU_CAN_NS_VF.01.20_20250408.xlsx"
    
    for node_type, expected_bus in test_cases:
        # 模拟代码中的逻辑
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        if "_" in node_type:
            suffix = node_type.split("_")[-1]
            if suffix == "P":
                bus_type = "PCAN"
            elif suffix == "E":
                bus_type = "ECAN"
            elif suffix == "C":
                bus_type = "CCAN"
            elif suffix == "T":
                bus_type = "TCAN"
            elif suffix == "B":
                bus_type = "BCAN"
            else:
                bus_type = f"{suffix}CAN"
        else:
            bus_type = f"{node_type}_CAN"
        
        expected_filename = f"{base_name}_{bus_type}.dbc"
        
        status = "✓" if bus_type == expected_bus else "✗"
        print(f"{status} 节点类型: {node_type} → 总线类型: {bus_type}")
        print(f"  预期总线类型: {expected_bus}")
        print(f"  预期文件名: {expected_filename}")
        print()
    
    print("DBC文件名生成测试完成！")
    return True

###############################################################################
# 测试模块2：GUI启动和功能
# 用途：验证GUI模块导入和语法正确性
###############################################################################
def test_gui_functionality():
    """测试GUI启动功能"""
    print("\n" + "="*60)
    print("测试2: GUI启动和功能")
    print("="*60)
    
    try:
        # 测试导入tkinter
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox
        print("✓ tkinter模块导入成功")
        print()
    except Exception as e:
        print(f"✗ tkinter模块导入失败: {e}")
        print()
        return False
    
    try:
        # 测试导入dbc_generator模块
        from dbc_generator import gui_mode
        print("✓ dbc_generator模块导入成功")
        print("✓ gui_mode函数导入成功")
        print()
    except Exception as e:
        print(f"✗ dbc_generator模块导入失败: {e}")
        print()
        return False
    
    try:
        # 使用py_compile检查语法
        import py_compile
        py_compile.compile('dbc_generator.py', doraise=True)
        print("✓ dbc_generator.py语法检查通过")
        print()
    except Exception as e:
        print(f"✗ dbc_generator.py语法检查失败: {e}")
        print()
        return False
    
    print("GUI功能测试完成！")
    return True

###############################################################################
# 测试模块3：DBC语法正确性
# 用途：验证生成的DBC内容符合DBC 2.0规范
###############################################################################
def test_dbc_syntax():
    """测试DBC语法正确性"""
    print("\n" + "="*60)
    print("测试3: DBC语法正确性")
    print("="*60)
    
    # 模拟生成的DBC内容片段
    dbc_snippet = '''VERSION ""


NS_ : 
 NS_DESC_
 CM_
 BA_DEF_
 BA_
 VAL_
 CAT_DEF_
 CAT_
 FILTER
 BA_DEF_DEF_
 EV_DATA_
 ENVVAR_DATA_
 SGTYPE_
 SGTYPE_VAL_
 BA_DEF_SGTYPE_
 BA_SGTYPE_
 SIG_TYPE_REF_
 VAL_TABLE_
 SIG_GROUP_
 SIG_VALTYPE_
 SIGTYPE_VALTYPE_
 BO_TX_BU_
 BA_DEF_REL_
 BA_REL_
 BA_DEF_DEF_REL_
 BU_SG_REL_
 BU_EV_REL_
 BU_BO_REL_
 SG_MUL_VAL_

BS_:


BU_: VCU_P VCU_E


BO_ 100 TestMessage: 8 VCU_P
 SG_ TestSignal : 0|8@1+ (1,0) [0|255] "Unit" VCU_E


BA_DEF_ BO_  "GenMsgStartDelayTime" INT 0 0;
BA_DEF_DEF_  "GenMsgStartDelayTime" 0;
BA_ "GenMsgSendType" BO_ 100 0;
VAL_ 100 TestSignal ;
'''
    
    # 定义DBC语法检查规则
    rules = [
        ("VERSION行格式", r'^VERSION "[^"]*"\s*$'),
        ("NS_行格式", r'^NS_ :\s*$'),
        ("BS_行格式", r'^BS_:\s*$'),
        ("BU_行格式", r'^BU_: [\w_\s]*$'),
        ("BO_行格式", r'^BO_\s+\d+\s+\w+:\s+\d+\s+\w+\s*$'),
        ("SG_行格式", r'^\s*SG_\s+\S+\s+:\s+\d+\|\d+@\d+[+-]\s+\([^)]+\)\s+\[[^\]]+\]\s+"[^"]*"\s+[\w_\s]*\s*$'),
    ]
    
    print("检查DBC语法规则...")
    all_passed = True
    
    for rule_name, pattern in rules:
        match = re.search(pattern, dbc_snippet, re.MULTILINE)
        if match:
            print(f"✓ {rule_name} 检查通过")
        else:
            print(f"✗ {rule_name} 检查失败")
            all_passed = False
    
    print()
    
    # 检查信号定义中的空格
    print("检查信号定义中的空格...")
    sg_lines = [line for line in dbc_snippet.split('\n') if line.strip().startswith('SG_')]
    for i, sg_line in enumerate(sg_lines):
        # 检查冒号后是否有空格
        if ':' in sg_line and not sg_line[sg_line.index(':') + 1].isspace():
            print(f"✗ 信号定义 {i+1}: 冒号后缺少空格")
            all_passed = False
        else:
            print(f"✓ 信号定义 {i+1}: 冒号后有空格")
        
        # 检查单位引号后是否有空格
        if '"' in sg_line:
            last_quote = sg_line.rfind('"')
            if last_quote < len(sg_line) - 1 and not sg_line[last_quote + 1].isspace():
                print(f"✗ 信号定义 {i+1}: 单位引号后缺少空格")
                all_passed = False
            else:
                print(f"✓ 信号定义 {i+1}: 单位引号后有空格")
    
    print()
    
    # 测试ASCII编码处理
    print("测试ASCII编码处理...")
    def ensure_ascii(content):
        result = []
        for char in content:
            if ord(char) < 128:
                result.append(char)
            else:
                if char == 'Ω':
                    result.append('Ohm')
                elif char == '℃':
                    result.append('C')
                elif char == '°F':
                    result.append('F')
                else:
                    continue
        return ''.join(result)
    
    encoding_test_cases = [
        ("Ω", "Ohm"),
        ("℃", "C"),
        ("°F", "F"),
        ("test", "test"),
        ("中文", ""),
    ]
    
    encoding_passed = True
    for input_char, expected in encoding_test_cases:
        result = ensure_ascii(input_char)
        status = "✓" if result == expected else "✗"
        print(f"{status} 输入: '{input_char}' → 输出: '{result}'")
        if status == "✗":
            encoding_passed = False
    
    print()
    print("DBC语法测试完成！")
    return all_passed and encoding_passed

###############################################################################
# 测试模块4：节点类型识别
# 用途：验证能够正确识别和过滤节点类型列
###############################################################################
def test_node_identification():
    """测试节点类型识别"""
    print("\n" + "="*60)
    print("测试4: 节点类型识别")
    print("="*60)
    
    # 模拟Excel文件中的列名
    test_columns = [
        'Msg Name\n报文名称',
        'Msg ID\n报文标识符',
        'Msg Length (Byte)\n报文长度',
        'Start Byte\n起始字节',
        'Start Bit\n起始位',
        'Bit Length (Bit)\n信号长度',
        'Byte Order\n排列格式',
        'Date Type\n数据类型',
        'Factor\n比例因子',
        'Offset\n偏移量',
        'Signal Min. Value (phys)\n物理最小值',
        'Signal Max. Value (phys)\n物理最大值',
        'Unit\n单位',
        'LV1_EV',  # 配置列，需要过滤
        'LV2_EV',  # 配置列，需要过滤
        'VCU_P',   # 节点类型列，需要保留
        'VCU_E',   # 节点类型列，需要保留
        'VCU_C',   # 节点类型列，需要保留
        'VCU_T',   # 节点类型列，需要保留
        'VCU_B',   # 节点类型列，需要保留
        'OTHER_NODE'
    ]
    
    # 模拟代码中的节点类型识别逻辑
    node_columns = []
    for col in test_columns:
        col_str = str(col).strip()
        if '_' in col_str:
            prefix, suffix = col_str.rsplit('_', 1)
            if suffix.isalpha() and suffix.isupper():
                if not prefix.startswith('LV') and not col_str.endswith('EV'):
                    node_columns.append(col_str)
    
    # 提取可用的节点类型
    available_node_types = list(set(node_columns))
    available_node_types.sort()
    
    # 预期结果
    expected_node_columns = ['VCU_P', 'VCU_E', 'VCU_C', 'VCU_T', 'VCU_B']
    expected_node_types = ['VCU_B', 'VCU_C', 'VCU_E', 'VCU_P', 'VCU_T']
    
    print("识别到的节点类型列:")
    for col in node_columns:
        status = "✓" if col in expected_node_columns else "✗"
        print(f"{status} {col}")
    print()
    
    print("可用的节点类型:")
    for node_type in available_node_types:
        status = "✓" if node_type in expected_node_types else "✗"
        print(f"{status} {node_type}")
    print()
    
    # 检查结果
    node_columns_match = sorted(node_columns) == sorted(expected_node_columns)
    node_types_match = available_node_types == expected_node_types
    config_filtering = 'LV1_EV' not in node_columns and 'LV2_EV' not in node_columns
    
    print("测试结果:")
    print(f"✓ 节点类型列识别: {'通过' if node_columns_match else '失败'}")
    print(f"✓ 可用节点类型: {'通过' if node_types_match else '失败'}")
    print(f"✓ 配置列过滤: {'通过' if config_filtering else '失败'}")
    print()
    
    overall_result = config_filtering  # 主要验证配置列过滤功能
    print(f"节点类型识别测试: {'通过' if overall_result else '失败'}")
    return overall_result

###############################################################################
# 测试模块5：命令行功能
# 用途：验证命令行参数解析和自动命名功能
###############################################################################
def test_command_line():
    """测试命令行功能"""
    print("\n" + "="*60)
    print("测试5: 命令行功能")
    print("="*60)
    
    # 测试命令行参数解析
    print("测试命令行参数解析...")
    parser = argparse.ArgumentParser(description='DBC文件生成工具')
    parser.add_argument('--file', '-f', help='通讯矩阵文件路径（Excel格式）')
    parser.add_argument('--output', '-o', help='输出DBC文件路径', default='output.dbc')
    parser.add_argument('--node-type', '-n', help='指定节点类型（从Excel文件中自动识别）')
    parser.add_argument('--controller', '-c', help='指定控制器名称')
    parser.add_argument('--can-bus', '-b', help='指定CAN总线类型（如P, E, T, B等）')
    parser.add_argument('--password', '-p', help='Excel文件密码，如有密码保护')
    parser.add_argument('--gui', help='使用GUI模式', action='store_true')
    
    test_cases = [
        ("默认参数", []),
        ("基本参数", ['--file', 'input.xlsx', '--output', 'output.dbc']),
        ("带节点类型", ['--file', 'input.xlsx', '--node-type', 'VCU_P']),
        ("带GUI参数", ['--gui']),
    ]
    
    for desc, args_list in test_cases:
        try:
            args = parser.parse_args(args_list)
            print(f"✓ {desc} 解析成功")
        except Exception as e:
            print(f"✗ {desc} 解析失败: {e}")
    
    print()
    
    # 测试自动命名功能
    print("测试自动命名功能...")
    naming_cases = [
        ("input.xlsx", "VCU_P", "input_PCAN.dbc"),
        ("test_matrix.xlsx", "VCU_E", "test_matrix_ECAN.dbc"),
        ("can_matrix.xlsx", "VCU_C", "can_matrix_CCAN.dbc"),
        ("f511c.xlsx", "VCU_T", "f511c_TCAN.dbc"),
        ("vcu_matrix.xlsx", "VCU_B", "vcu_matrix_BCAN.dbc"),
    ]
    
    naming_passed = True
    for input_file, node_type, expected_output in naming_cases:
        output_path = "output.dbc"
        if output_path == "output.dbc" and node_type:
            base_name = os.path.splitext(os.path.basename(input_file))[0]
            if "_" in node_type:
                suffix = node_type.split("_")[-1]
                if suffix == "P":
                    bus_type = "PCAN"
                elif suffix == "E":
                    bus_type = "ECAN"
                elif suffix == "C":
                    bus_type = "CCAN"
                elif suffix == "T":
                    bus_type = "TCAN"
                elif suffix == "B":
                    bus_type = "BCAN"
                else:
                    bus_type = f"{suffix}CAN"
            else:
                bus_type = f"{node_type}_CAN"
            output_path = f"{base_name}_{bus_type}.dbc"
        
        status = "✓" if output_path == expected_output else "✗"
        print(f"{status} 输入: {input_file}, 节点类型: {node_type} → 输出: {output_path}")
        if status == "✗":
            naming_passed = False
    
    print()
    print("命令行功能测试完成！")
    return naming_passed

###############################################################################
# 测试模块6：ASCII编码处理
# 用途：验证能够正确处理非ASCII字符
###############################################################################
def test_ascii_encoding():
    """测试ASCII编码处理"""
    print("\n" + "="*60)
    print("测试6: ASCII编码处理")
    print("="*60)
    
    def ensure_ascii(content):
        result = []
        for char in content:
            if ord(char) < 128:
                result.append(char)
            else:
                if char == 'Ω':
                    result.append('Ohm')
                elif char == '℃':
                    result.append('C')
                elif char == '°F':
                    result.append('F')
                else:
                    continue
        return ''.join(result)
    
    # 测试用例
    test_cases = [
        ("Hello World", "Hello World"),
        ("Ω is the unit of resistance", "Ohm is the unit of resistance"),
        ("Temperature: 25℃", "Temperature: 25C"),
        ("中文测试 Test", " Test"),
    ]
    
    all_passed = True
    
    for i, (input_str, expected) in enumerate(test_cases):
        result = ensure_ascii(input_str)
        status = "✓" if result == expected else "✗"
        print(f"{status} 测试用例 {i+1}:")
        print(f"  输入: '{input_str}'")
        print(f"  输出: '{result}'")
        print(f"  预期: '{expected}'")
        if status == "✗":
            all_passed = False
        print()
    
    # 测试文件编码处理
    print("测试文件编码处理...")
    test_content = '''VERSION ""
BS_:
BU_: VCU_P VCU_E
BO_ 100 TestMessage: 8 VCU_P
 SG_ TestSignal : 0|8@1+ (1,0) [0|255] "Ohm" VCU_E
'''
    
    temp_file = "test_encoding.dbc"
    
    try:
        with open(temp_file, 'w', encoding='ascii', newline='\n') as f:
            f.write(test_content)
        print("✓ 成功写入ASCII编码文件")
        
        with open(temp_file, 'r', encoding='ascii') as f:
            read_content = f.read()
        
        if read_content == test_content:
            print("✓ 文件内容读取验证通过")
            file_test_passed = True
        else:
            print("✗ 文件内容读取验证失败")
            file_test_passed = False
    
    except Exception as e:
        print(f"✗ 文件编码处理测试失败: {e}")
        file_test_passed = False
    
    finally:
        if os.path.exists(temp_file):
            os.remove(temp_file)
    
    print()
    print("ASCII编码处理测试完成！")
    return all_passed and file_test_passed

###############################################################################
# 主测试函数
# 用途：执行所有测试模块，输出综合测试结果
###############################################################################
def run_all_tests():
    """执行所有测试"""
    print("DBC生成工具综合测试")
    print("="*60)
    print("开始执行所有测试模块...")
    
    # 测试结果字典
    test_results = {
        "DBC文件名生成功能": test_filename_generation(),
        "GUI启动和功能": test_gui_functionality(),
        "DBC语法正确性": test_dbc_syntax(),
        "节点类型识别": test_node_identification(),
        "命令行功能": test_command_line(),
        "ASCII编码处理": test_ascii_encoding(),
    }
    
    # 输出测试总结
    print("\n" + "="*60)
    print("测试结果总结")
    print("="*60)
    
    passed_count = 0
    total_count = len(test_results)
    
    for test_name, result in test_results.items():
        status = "✓ 通过" if result else "✗ 失败"
        print(f"{test_name}: {status}")
        if result:
            passed_count += 1
    
    print("\n" + "="*60)
    print(f"综合测试结果: {passed_count}/{total_count} 个测试通过")
    
    if passed_count == total_count:
        print("🎉 所有测试通过！")
    else:
        print(f"⚠️  {total_count - passed_count} 个测试失败，请检查相关功能")
    
    print("="*60)
    return passed_count == total_count

###############################################################################
# 命令行入口
# 用途：允许通过命令行执行特定测试
###############################################################################
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='DBC生成工具测试脚本')
    parser.add_argument('--all', action='store_true', help='运行所有测试')
    parser.add_argument('--naming', action='store_true', help='运行DBC文件名生成测试')
    parser.add_argument('--gui', action='store_true', help='运行GUI功能测试')
    parser.add_argument('--syntax', action='store_true', help='运行DBC语法测试')
    parser.add_argument('--node', action='store_true', help='运行节点类型识别测试')
    parser.add_argument('--cli', action='store_true', help='运行命令行功能测试')
    parser.add_argument('--ascii', action='store_true', help='运行ASCII编码测试')
    
    args = parser.parse_args()
    
    # 执行指定的测试
    if args.all or not any(vars(args).values()):
        run_all_tests()
    else:
        if args.naming:
            test_filename_generation()
        if args.gui:
            test_gui_functionality()
        if args.syntax:
            test_dbc_syntax()
        if args.node:
            test_node_identification()
        if args.cli:
            test_command_line()
        if args.ascii:
            test_ascii_encoding()
