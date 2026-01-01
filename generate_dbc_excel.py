#!/usr/bin/env python3
"""
DBC Excel文件生成工具
用于生成符合DBC生成规范的Excel模板文件
"""

import argparse
import pandas as pd
import os
from typing import Dict, List

class DBCExcelGenerator:
    """DBC Excel文件生成器"""
    
    # DBC Excel文件各Sheet的预期结构定义
    EXPECTED_STRUCTURES = {
        'Nodes': {
            'required_columns': ['NodeName'],
            'optional_columns': ['Comment'],
            'column_types': {
                'NodeName': str,
                'Comment': str
            }
        },
        'Messages': {
            'required_columns': ['MessageID', 'MessageName', 'DLC', 'Sender'],
            'optional_columns': ['CycleTime', 'Comment'],
            'column_types': {
                'MessageID': (int, str),  # 支持整数或十六进制字符串
                'MessageName': str,
                'DLC': int,
                'Sender': str,
                'CycleTime': (int, float),
                'Comment': str
            }
        },
        'Signals': {
            'required_columns': ['MessageName', 'SignalName', 'StartBit', 'SignalLength', 
                               'ByteOrder', 'ValueType', 'Factor', 'Offset', 'Receiver'],
            'optional_columns': ['Min', 'Max', 'Unit', 'Comment'],
            'column_types': {
                'MessageName': str,
                'SignalName': str,
                'StartBit': int,
                'SignalLength': int,
                'ByteOrder': str,
                'ValueType': str,
                'Factor': (int, float),
                'Offset': (int, float),
                'Min': (int, float),
                'Max': (int, float),
                'Unit': str,
                'Receiver': str,
                'Comment': str
            }
        },
        'ValueTables': {
            'required_columns': ['TableName', 'Value', 'Description'],
            'optional_columns': [],
            'column_types': {
                'TableName': str,
                'Value': int,
                'Description': str
            }
        },
        'SignalValueTables': {
            'required_columns': ['SignalName', 'TableName'],
            'optional_columns': [],
            'column_types': {
                'SignalName': str,
                'TableName': str
            }
        }
    }
    
    def __init__(self, output_file: str):
        """
        初始化生成器
        
        Args:
            output_file: 输出Excel文件路径
        """
        self.output_file = output_file
    
    def generate_excel(self) -> bool:
        """
        生成DBC Excel文件
        
        Returns:
            bool: 是否成功生成
        """
        try:
            # 创建Excel写入器
            with pd.ExcelWriter(self.output_file, engine='xlsxwriter') as writer:
                # 为每个Sheet生成数据
                for sheet_name in self.EXPECTED_STRUCTURES:
                    structure = self.EXPECTED_STRUCTURES[sheet_name]
                    
                    # 合并必填列和可选列
                    all_columns = structure['required_columns'] + structure['optional_columns']
                    
                    # 创建空DataFrame
                    df = pd.DataFrame(columns=all_columns)
                    
                    # 写入Sheet
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                
            print(f"成功生成DBC Excel文件: {self.output_file}")
            return True
        except Exception as e:
            print(f"生成Excel文件失败: {str(e)}")
            return False
    
    def generate_with_example_data(self) -> bool:
        """
        生成带有示例数据的DBC Excel文件
        
        Returns:
            bool: 是否成功生成
        """
        try:
            # 创建Excel写入器
            with pd.ExcelWriter(self.output_file, engine='xlsxwriter') as writer:
                # 生成Nodes Sheet示例数据
                nodes_data = {
                    'NodeName': ['ECU1', 'ECU2', 'ECU3'],
                    'Comment': ['发动机控制单元', '变速箱控制单元', '车身控制单元']
                }
                pd.DataFrame(nodes_data).to_excel(writer, sheet_name='Nodes', index=False)
                
                # 生成Messages Sheet示例数据
                messages_data = {
                    'MessageID': [0x100, 0x200, 0x300],
                    'MessageName': ['EngineStatus', 'TransmissionStatus', 'BodyControl'],
                    'DLC': [8, 6, 4],
                    'Sender': ['ECU1', 'ECU2', 'ECU3'],
                    'CycleTime': [100, 200, 500],
                    'Comment': ['发动机状态信息', '变速箱状态信息', '车身控制信息']
                }
                pd.DataFrame(messages_data).to_excel(writer, sheet_name='Messages', index=False)
                
                # 生成Signals Sheet示例数据
                signals_data = {
                    'MessageName': ['EngineStatus', 'EngineStatus', 'TransmissionStatus', 'BodyControl'],
                    'SignalName': ['EngineSpeed', 'EngineTemp', 'GearPosition', 'LightStatus'],
                    'StartBit': [0, 16, 0, 0],
                    'SignalLength': [16, 8, 4, 8],
                    'ByteOrder': ['Motorola', 'Motorola', 'Intel', 'Intel'],
                    'ValueType': ['Unsigned', 'Signed', 'Unsigned', 'Unsigned'],
                    'Factor': [0.1, 1.0, 1.0, 1.0],
                    'Offset': [0, 0, 0, 0],
                    'Receiver': ['ECU2,ECU3', 'ECU2', 'ECU1,ECU3', 'ECU1,ECU2'],
                    'Min': [0.0, -40.0, 0.0, 0.0],
                    'Max': [8000.0, 120.0, 6.0, 255.0],
                    'Unit': ['rpm', '°C', 'gear', ''],
                    'Comment': ['发动机转速', '发动机温度', '档位位置', '灯光状态']
                }
                pd.DataFrame(signals_data).to_excel(writer, sheet_name='Signals', index=False)
                
                # 生成ValueTables Sheet示例数据
                value_tables_data = {
                    'TableName': ['GearPosition', 'GearPosition', 'LightStatus', 'LightStatus'],
                    'Value': [0, 1, 0, 1],
                    'Description': ['Neutral', 'Drive', 'Off', 'On']
                }
                pd.DataFrame(value_tables_data).to_excel(writer, sheet_name='ValueTables', index=False)
                
                # 生成SignalValueTables Sheet示例数据
                signal_value_tables_data = {
                    'SignalName': ['GearPosition', 'LightStatus'],
                    'TableName': ['GearPosition', 'LightStatus']
                }
                pd.DataFrame(signal_value_tables_data).to_excel(writer, sheet_name='SignalValueTables', index=False)
            
            print(f"成功生成带有示例数据的DBC Excel文件: {self.output_file}")
            return True
        except Exception as e:
            print(f"生成Excel文件失败: {str(e)}")
            return False

def main():
    """主函数"""
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='DBC Excel文件生成工具')
    parser.add_argument('--output', '-o', help='输出Excel文件路径', default='dbc_template.xlsx')
    parser.add_argument('--with-examples', '-e', help='生成带有示例数据的Excel文件', action='store_true')
    
    args = parser.parse_args()
    
    # 移除路径中的引号
    output_file = args.output.strip('"\'')
    
    # 创建生成器实例
    generator = DBCExcelGenerator(output_file)
    
    # 生成Excel文件
    if args.with_examples:
        generator.generate_with_example_data()
    else:
        generator.generate_excel()

if __name__ == '__main__':
    main()