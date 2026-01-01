#!/usr/bin/env python3
"""
通讯矩阵到DBC Excel模板转换工具
用于将通讯矩阵转换为符合DBC生成规范的Excel文件
"""

import argparse
import pandas as pd
import os
from typing import Dict, List, Optional, Any
import json

class DBCMatrixConverter:
    """通讯矩阵到DBC Excel模板转换器"""
    
    # DBC Excel文件各Sheet的预期结构定义
    DBC_STRUCTURES = {
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
    
    def __init__(self, matrix_file: str, output_file: str):
        """
        初始化转换器
        
        Args:
            matrix_file: 通讯矩阵文件路径
            output_file: 输出DBC Excel文件路径
        """
        self.matrix_file = matrix_file
        self.output_file = output_file
        self.matrix_data = {}
        self.mapping_config = {}
    
    def load_matrix(self) -> bool:
        """
        加载通讯矩阵文件
        
        Returns:
            bool: 是否成功加载
        """
        try:
            if not os.path.exists(self.matrix_file):
                print(f"通讯矩阵文件不存在: {self.matrix_file}")
                return False
            
            # 支持Excel和CSV文件
            if self.matrix_file.endswith('.xlsx') or self.matrix_file.endswith('.xls'):
                xl = pd.ExcelFile(self.matrix_file)
                for sheet_name in xl.sheet_names:
                    self.matrix_data[sheet_name] = pd.read_excel(self.matrix_file, sheet_name=sheet_name)
            elif self.matrix_file.endswith('.csv'):
                self.matrix_data['Matrix'] = pd.read_csv(self.matrix_file)
            else:
                print(f"不支持的文件格式: {self.matrix_file}")
                return False
            
            print(f"成功加载通讯矩阵文件: {self.matrix_file}")
            print(f"包含的工作表: {list(self.matrix_data.keys())}")
            return True
        except Exception as e:
            print(f"加载通讯矩阵文件失败: {str(e)}")
            return False
    
    def load_mapping_config(self, config_file: str) -> bool:
        """
        加载映射配置文件
        
        Args:
            config_file: 映射配置文件路径
            
        Returns:
            bool: 是否成功加载
        """
        try:
            if not os.path.exists(config_file):
                print(f"映射配置文件不存在: {config_file}")
                return False
            
            with open(config_file, 'r', encoding='utf-8') as f:
                self.mapping_config = json.load(f)
            
            print(f"成功加载映射配置: {config_file}")
            return True
        except Exception as e:
            print(f"加载映射配置失败: {str(e)}")
            return False
    
    def generate_default_mapping(self, output_config: str) -> bool:
        """
        生成默认映射配置文件
        
        Args:
            output_config: 输出配置文件路径
            
        Returns:
            bool: 是否成功生成
        """
        try:
            if not self.matrix_data:
                print("请先加载通讯矩阵文件")
                return False
            
            # 为每个工作表生成映射配置
            default_config = {
                'mapping': {},
                'options': {
                    'default_byte_order': 'Motorola',
                    'default_value_type': 'Unsigned',
                    'default_factor': 1.0,
                    'default_offset': 0.0
                }
            }
            
            # 遍历通讯矩阵的所有工作表
            for sheet_name, df in self.matrix_data.items():
                # 为每个工作表生成默认映射
                # 这里假设通讯矩阵是单工作表结构，包含所有信号信息
                if sheet_name == 'Matrix' or len(self.matrix_data) == 1:
                    # 生成Signals映射
                    signals_mapping = {}
                    for col in df.columns:
                        # 尝试自动匹配列名
                        if 'signal' in col.lower() and 'name' in col.lower():
                            signals_mapping[col] = 'SignalName'
                        elif 'message' in col.lower() and 'name' in col.lower():
                            signals_mapping[col] = 'MessageName'
                        elif 'message' in col.lower() and 'id' in col.lower():
                            signals_mapping[col] = 'MessageID'
                        elif 'dbc' in col.lower() or 'dlc' in col.lower():
                            signals_mapping[col] = 'DLC'
                        elif 'sender' in col.lower():
                            signals_mapping[col] = 'Sender'
                        elif 'receiver' in col.lower():
                            signals_mapping[col] = 'Receiver'
                        elif 'start' in col.lower() and 'bit' in col.lower():
                            signals_mapping[col] = 'StartBit'
                        elif 'signal' in col.lower() and 'length' in col.lower():
                            signals_mapping[col] = 'SignalLength'
                        elif 'byte' in col.lower() and 'order' in col.lower():
                            signals_mapping[col] = 'ByteOrder'
                        elif 'value' in col.lower() and 'type' in col.lower():
                            signals_mapping[col] = 'ValueType'
                        elif 'factor' in col.lower():
                            signals_mapping[col] = 'Factor'
                        elif 'offset' in col.lower():
                            signals_mapping[col] = 'Offset'
                        elif 'min' in col.lower():
                            signals_mapping[col] = 'Min'
                        elif 'max' in col.lower():
                            signals_mapping[col] = 'Max'
                        elif 'unit' in col.lower():
                            signals_mapping[col] = 'Unit'
                        elif 'comment' in col.lower():
                            signals_mapping[col] = 'Comment'
                    
                    default_config['mapping']['Signals'] = {
                        'source_sheet': sheet_name,
                        'column_mapping': signals_mapping
                    }
                    
                    # 生成Messages映射（从Signals中提取唯一消息）
                    messages_column_mapping = {}
                    for dbc_col in ['MessageID', 'MessageName', 'DLC', 'Sender']:
                        # 查找对应的源列
                        source_cols = [k for k in signals_mapping if dbc_col in signals_mapping[k]]
                        if source_cols:
                            messages_column_mapping[dbc_col] = source_cols[0]  # 取第一个匹配的列
                    
                    default_config['mapping']['Messages'] = {
                        'source_sheet': sheet_name,
                        'column_mapping': messages_column_mapping
                    }
                    
                    # 生成Nodes映射（从Signals中提取唯一节点）
                    nodes_column_mapping = {}
                    # 查找对应的源列
                    source_cols = [k for k in signals_mapping if 'Sender' in signals_mapping[k]]
                    if source_cols:
                        nodes_column_mapping['NodeName'] = source_cols[0]  # 取第一个匹配的列
                    
                    default_config['mapping']['Nodes'] = {
                        'source_sheet': sheet_name,
                        'column_mapping': nodes_column_mapping
                    }
                    break
            
            # 保存默认配置
            with open(output_config, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, ensure_ascii=False, indent=4)
            
            print(f"成功生成默认映射配置: {output_config}")
            print("请根据实际情况修改配置文件，然后重新运行转换")
            return True
        except Exception as e:
            print(f"生成默认映射配置失败: {str(e)}")
            return False
    
    def convert(self) -> bool:
        """
        执行通讯矩阵到DBC Excel的转换
        
        Returns:
            bool: 是否成功转换
        """
        try:
            if not self.matrix_data:
                print("请先加载通讯矩阵文件")
                return False
            
            if not self.mapping_config:
                print("请先加载映射配置文件")
                return False
            
            # 创建Excel写入器
            with pd.ExcelWriter(self.output_file, engine='xlsxwriter') as writer:
                # 处理每个DBC工作表
                for dbc_sheet in self.DBC_STRUCTURES:
                    if dbc_sheet in self.mapping_config['mapping']:
                        # 获取映射配置
                        mapping = self.mapping_config['mapping'][dbc_sheet]
                        source_sheet = mapping['source_sheet']
                        column_mapping = mapping['column_mapping']
                        
                        if source_sheet not in self.matrix_data:
                            print(f"映射配置中的源工作表不存在: {source_sheet}")
                            continue
                        
                        # 获取源数据
                        source_df = self.matrix_data[source_sheet].copy()
                        
                        # 创建目标DataFrame
                        target_columns = self.DBC_STRUCTURES[dbc_sheet]['required_columns'] + self.DBC_STRUCTURES[dbc_sheet]['optional_columns']
                        target_df = pd.DataFrame(columns=target_columns)
                        
                        # 映射列数据
                        for source_col, target_col in column_mapping.items():
                            if source_col in source_df.columns and target_col in target_df.columns:
                                target_df[target_col] = source_df[source_col]
                        
                        # 处理默认值
                        options = self.mapping_config.get('options', {})
                        if dbc_sheet == 'Signals':
                            # 设置默认ByteOrder
                            if 'ByteOrder' in target_df.columns:
                                target_df['ByteOrder'] = target_df['ByteOrder'].fillna(options.get('default_byte_order', 'Motorola'))
                            
                            # 设置默认ValueType
                            if 'ValueType' in target_df.columns:
                                target_df['ValueType'] = target_df['ValueType'].fillna(options.get('default_value_type', 'Unsigned'))
                            
                            # 设置默认Factor
                            if 'Factor' in target_df.columns:
                                target_df['Factor'] = target_df['Factor'].fillna(options.get('default_factor', 1.0))
                            
                            # 设置默认Offset
                            if 'Offset' in target_df.columns:
                                target_df['Offset'] = target_df['Offset'].fillna(options.get('default_offset', 0.0))
                        
                        # 去重处理
                        if dbc_sheet == 'Messages':
                            # 按MessageID和MessageName去重
                            target_df.drop_duplicates(subset=['MessageID', 'MessageName'], inplace=True)
                        elif dbc_sheet == 'Nodes':
                            # 按NodeName去重
                            target_df.drop_duplicates(subset=['NodeName'], inplace=True)
                        
                        # 写入目标工作表
                        target_df.to_excel(writer, sheet_name=dbc_sheet, index=False)
                        print(f"成功转换工作表: {source_sheet} -> {dbc_sheet}")
                    else:
                        # 对于没有映射的DBC工作表，创建空表
                        target_columns = self.DBC_STRUCTURES[dbc_sheet]['required_columns'] + self.DBC_STRUCTURES[dbc_sheet]['optional_columns']
                        target_df = pd.DataFrame(columns=target_columns)
                        target_df.to_excel(writer, sheet_name=dbc_sheet, index=False)
                        print(f"创建空工作表: {dbc_sheet}")
            
            print(f"成功生成DBC Excel文件: {self.output_file}")
            return True
        except Exception as e:
            print(f"转换失败: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def show_matrix_info(self) -> bool:
        """
        显示通讯矩阵的基本信息
        
        Returns:
            bool: 是否成功显示
        """
        if not self.matrix_data:
            print("请先加载通讯矩阵文件")
            return False
        
        print("通讯矩阵基本信息:")
        for sheet_name, df in self.matrix_data.items():
            print(f"\n工作表: {sheet_name}")
            print(f"  行数: {len(df)}")
            print(f"  列数: {len(df.columns)}")
            print(f"  列名: {list(df.columns)}")
            print(f"  数据示例:")
            print(df.head(3).to_string())
        
        return True

def main():
    """主函数"""
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='通讯矩阵到DBC Excel模板转换工具')
    parser.add_argument('--matrix', '-m', help='通讯矩阵文件路径（支持Excel和CSV）')
    parser.add_argument('--output', '-o', help='输出DBC Excel文件路径', default='dbc_converted.xlsx')
    parser.add_argument('--config', '-c', help='映射配置文件路径')
    parser.add_argument('--generate-config', '-g', help='生成默认映射配置文件', action='store_true')
    parser.add_argument('--show-info', '-i', help='显示通讯矩阵信息', action='store_true')
    
    args = parser.parse_args()
    
    if not args.matrix:
        print("请指定通讯矩阵文件路径")
        return
    
    # 创建转换器实例
    converter = DBCMatrixConverter(args.matrix, args.output)
    
    # 加载通讯矩阵
    if not converter.load_matrix():
        return
    
    # 显示通讯矩阵信息
    if args.show_info:
        converter.show_matrix_info()
        return
    
    # 生成默认映射配置
    if args.generate_config:
        config_file = args.config if args.config else 'dbc_mapping.json'
        converter.generate_default_mapping(config_file)
        return
    
    # 加载映射配置
    if args.config:
        if not converter.load_mapping_config(args.config):
            return
    else:
        # 自动生成并使用默认配置
        default_config = 'dbc_mapping_auto.json'
        if converter.generate_default_mapping(default_config):
            if converter.load_mapping_config(default_config):
                print("使用自动生成的映射配置进行转换")
            else:
                return
        else:
            return
    
    # 执行转换
    converter.convert()

if __name__ == '__main__':
    main()