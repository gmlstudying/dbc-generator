# 命令行命令模块
"""
实现DBC生成器的命令行接口
"""

import argparse
import sys
from core.excel_handler import ExcelHandler
from core.dbc_generator import DBCGenerator


def parse_args():
    """
    解析命令行参数
    
    Returns:
        argparse.Namespace: 解析后的参数对象
    """
    parser = argparse.ArgumentParser(
        description='DBC生成器 - 从Excel通讯矩阵生成DBC文件',
        formatter_class=argparse.RawTextHelpFormatter
    )
    
    # 必选参数
    parser.add_argument(
        '-i', '--input',
        required=True,
        help='通讯矩阵Excel文件路径'
    )
    
    parser.add_argument(
        '-o', '--output',
        required=True,
        help='输出DBC文件路径'
    )
    
    # 可选参数
    parser.add_argument(
        '-p', '--password',
        help='Excel文件密码（如果有密码保护）'
    )
    
    parser.add_argument(
        '-n', '--node-type',
        help='节点类型'
    )
    
    parser.add_argument(
        '-c', '--controller-name',
        help='控制器名称'
    )
    
    parser.add_argument(
        '-b', '--bus-type',
        help='目标总线类型（如C、E、P等）'
    )
    
    parser.add_argument(
        '--canfd',
        action='store_true',
        help='生成CANFD格式的DBC文件'
    )
    
    parser.add_argument(
        '-v', '--version',
        action='version',
        version='DBC Generator v1.0'
    )
    
    return parser.parse_args()


def main():
    """
    命令行主函数
    
    Returns:
        int: 退出码，0表示成功，非0表示失败
    """
    # 确保输出能在控制台显示
    import os
    os.system("cls" if os.name == "nt" else "clear")
    
    print("DBC生成器 v1.0")
    print("=" * 50)
    print("正在初始化...")
    
    args = None
    
    try:
        # 解析命令行参数
        args = parse_args()
        
        # 创建核心组件实例
        excel_handler = ExcelHandler()
        dbc_generator = DBCGenerator()
            
        print(f"\nDBC生成器 v1.0")
        print(f"=" * 50)
        print(f"输入文件: {args.input}")
        print(f"输出文件: {args.output}")
        print(f"=" * 50)
            
        # 1. 读取Excel文件
        print("1. 正在读取Excel文件...")
        if not excel_handler.read_matrix_file(args.input, password=args.password):
            print("\n✗ 读取Excel文件失败")
            print("\n按任意键退出...")
            input()
            return 1
        
        # 2. 提取节点列
        print("2. 正在提取节点列...")
        node_columns = excel_handler.extract_node_columns()
        if not node_columns:
            print("\n✗ 未能识别到节点类型列")
            print("\n按任意键退出...")
            input()
            return 1
        
        # 3. 设置生成器数据
        print("3. 正在设置生成器数据...")
        dbc_generator.set_data(excel_handler.get_matrix_data(), node_columns)
        
        # 4. 生成DBC文件
        print("4. 正在生成DBC文件...")
        if dbc_generator.generate_dbc(
            args.output,
            node_type=args.node_type,
            controller_name=args.controller_name,
            target_bus_type=args.bus_type,
            is_canfd=args.canfd
        ):
            print(f"\n✓ DBC文件生成成功")
            print(f"输出路径: {args.output}")
            print("\n按任意键退出...")
            input()
            return 0
        else:
            print("\n✗ 生成DBC文件失败")
            print("\n按任意键退出...")
            input()
            return 1
            
    except KeyboardInterrupt:
        print("\n\n✗ 操作被用户中断")
        print("\n按任意键退出...")
        input()
        return 1
    except Exception as e:
        # 捕获所有其他错误
        if args is None:
            # 参数解析错误
            print(f"\n✗ 参数解析错误: {e}")
            print("\n使用 -h 或 --help 查看帮助信息")
            print("\n示例用法:")
            print("  dbc_generator.exe -i input.xlsx -o output.dbc")
        else:
            # 其他运行时错误
            print(f"\n✗ 运行时错误: {e}")
        
        print("\n按任意键退出...")
        input()
        return 1


if __name__ == '__main__':
    sys.exit(main())