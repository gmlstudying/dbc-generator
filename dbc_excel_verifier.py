#!/usr/bin/env python3
"""
DBC Excel文件格式验证工具
用于验证Excel文件是否符合DBC生成的文件格式规范
"""

import argparse
import pandas as pd
import os
from typing import Dict, List, Tuple, Optional

class DBCExcelVerifier:
    """DBC Excel文件验证器"""
    
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
    
    def __init__(self, excel_file: str):
        """
        初始化验证器
        
        Args:
            excel_file: Excel文件路径
        """
        self.excel_file = excel_file
        self.errors = []
        self.warnings = []
        self.sheet_names = []
        
    def load_excel(self) -> bool:
        """
        加载Excel文件并获取Sheet名称
        
        Returns:
            bool: 是否成功加载
        """
        try:
            if not os.path.exists(self.excel_file):
                self.errors.append(f"Excel文件不存在: {self.excel_file}")
                return False
            
            xl = pd.ExcelFile(self.excel_file)
            self.sheet_names = xl.sheet_names
            return True
        except Exception as e:
            self.errors.append(f"加载Excel文件失败: {str(e)}")
            return False
    
    def verify_sheet(self, sheet_name: str, expected_type: Optional[str] = None) -> bool:
        """
        验证指定的Sheet

        Args:
            sheet_name: 要验证的Sheet名称
            expected_type: 预期的Sheet类型（如果不指定，将根据Sheet名称或列结构自动判断）

        Returns:
            bool: 是否通过验证
        """
        # 清空错误和警告列表
        self.errors = []
        self.warnings = []

        # 加载Excel文件
        if not self.load_excel():
            return False
        
        # 检查Sheet是否存在
        if sheet_name not in self.sheet_names:
            self.errors.append(f"Sheet '{sheet_name}' 不存在于Excel文件中")
            return False
        
        # 确定Sheet类型
        sheet_type = expected_type
        
        # 读取Sheet数据，用于后续类型识别和验证
        try:
            df = pd.read_excel(self.excel_file, sheet_name=sheet_name)
        except Exception as e:
            self.errors.append(f"读取Sheet '{sheet_name}' 失败: {str(e)}")
            return False
        
        # 如果未指定类型，尝试识别
        if not sheet_type:
            # 1. 尝试根据Sheet名称匹配支持的类型
            sheet_name_lower = sheet_name.lower()
            for supported_type in self.EXPECTED_STRUCTURES:
                if supported_type.lower() in sheet_name_lower:
                    sheet_type = supported_type
                    break
            
            # 2. 如果名称匹配失败，尝试根据列结构识别
            if not sheet_type:
                sheet_type = self._detect_sheet_type_by_columns(df.columns.tolist())
            
            # 3. 如果仍未匹配到，提供可用类型列表
            if not sheet_type:
                supported_types = ', '.join(self.EXPECTED_STRUCTURES.keys())
                self.errors.append(f"Sheet '{sheet_name}' 类型无法自动识别，请使用 --type 参数指定类型。支持的类型: {supported_types}")
                return False
        
        # 检查指定的类型是否支持
        if sheet_type not in self.EXPECTED_STRUCTURES:
            supported_types = ', '.join(self.EXPECTED_STRUCTURES.keys())
            self.errors.append(f"不支持的Sheet类型: {sheet_type}。支持的类型: {supported_types}")
            return False
        
        # 获取预期结构
        expected = self.EXPECTED_STRUCTURES[sheet_type]
        
        # 1. 验证必填列是否存在
        missing_columns = [col for col in expected['required_columns'] if col not in df.columns]
        if missing_columns:
            self.errors.append(f"Sheet '{sheet_name}' 缺少必填列: {', '.join(missing_columns)}")
        
        # 2. 验证列数据类型
        for col in df.columns:
            if col in expected['column_types']:
                expected_type = expected['column_types'][col]
                actual_type = df[col].dtype
                
                # 检查实际类型是否与预期类型兼容
                if isinstance(expected_type, tuple):
                    # 支持多种类型
                    is_compatible = any(pd.api.types.is_dtype_equal(actual_type, t) or \
                                      pd.api.types.is_dtype_equal(actual_type, pd.Series([1]).astype(t).dtype) \
                                      for t in expected_type)
                else:
                    # 单一类型
                    is_compatible = pd.api.types.is_dtype_equal(actual_type, expected_type) or \
                                   pd.api.types.is_dtype_equal(actual_type, pd.Series([1]).astype(expected_type).dtype)
                
                if not is_compatible:
                    self.warnings.append(f"Sheet '{sheet_name}' 列 '{col}' 数据类型不符合预期: 预期 {expected_type}, 实际 {actual_type}")
        
        # 3. 执行特定Sheet的额外验证
        self._perform_specific_validations(sheet_name, sheet_type, df)
        
        return len(self.errors) == 0
    
    def _detect_sheet_type_by_columns(self, columns: List[str]) -> Optional[str]:
        """
        根据列名列表检测Sheet类型
        
        Args:
            columns: Sheet的列名列表
            
        Returns:
            Optional[str]: 检测到的Sheet类型，如果无法检测则返回None
        """
        columns_lower = [col.lower() for col in columns]
        best_match = None
        best_score = 0
        
        # 为每种类型计算匹配分数
        for sheet_type, config in self.EXPECTED_STRUCTURES.items():
            score = 0
            
            # 计算必填列匹配分数（权重更高）
            for req_col in config['required_columns']:
                if req_col.lower() in columns_lower:
                    score += 3
            
            # 计算可选列匹配分数
            for opt_col in config['optional_columns']:
                if opt_col.lower() in columns_lower:
                    score += 1
            
            # 计算总匹配率
            total_columns = len(config['required_columns']) + len(config['optional_columns'])
            if total_columns > 0:
                match_rate = score / (total_columns * 3)  # 最大可能分数是总列数 * 3
            else:
                match_rate = 0
            
            # 更新最佳匹配
            if match_rate > best_score:
                best_score = match_rate
                best_match = sheet_type
        
        # 只有当匹配率超过一定阈值时才返回结果
        if best_score > 0.5:
            return best_match
        
        return None
    
    def _perform_specific_validations(self, sheet_name: str, sheet_type: str, df: pd.DataFrame):
        """
        执行特定Sheet类型的额外验证
        
        Args:
            sheet_name: Sheet名称
            sheet_type: Sheet类型
            df: Sheet数据
        """
        if sheet_type == 'Messages':
            # 验证DLC值范围（0-8）
            if 'DLC' in df.columns:
                # 过滤掉NaN值并转换为数值类型
                dlc_col = pd.to_numeric(df['DLC'], errors='coerce')
                invalid_dlc_mask = (~dlc_col.between(0, 8)) & (~dlc_col.isna())
                invalid_dlc = df[invalid_dlc_mask]
                if not invalid_dlc.empty:
                    self.errors.append(f"Sheet '{sheet_name}' 存在无效的DLC值（必须在0-8之间）: {', '.join(map(str, invalid_dlc['DLC'].tolist()))}")

            # 验证消息ID唯一性
            if 'MessageID' in df.columns:
                # 过滤掉NaN值后再检查重复
                non_nan_ids = df[df['MessageID'].notna()]
                duplicate_ids = non_nan_ids[non_nan_ids.duplicated('MessageID', keep=False)]['MessageID'].drop_duplicates().tolist()
                if duplicate_ids:
                    self.errors.append(f"Sheet '{sheet_name}' 存在重复的MessageID: {', '.join(map(str, duplicate_ids))}")
        
        elif sheet_type == 'Signals':
            # 验证ByteOrder值
            if 'ByteOrder' in df.columns:
                valid_byte_orders = ['Motorola', 'Intel']
                invalid_byte_orders = df[~df['ByteOrder'].isin(valid_byte_orders)]
                if not invalid_byte_orders.empty:
                    self.errors.append(f"Sheet '{sheet_name}' 存在无效的ByteOrder值（必须是'Motorola'或'Intel'）: {', '.join(invalid_byte_orders['ByteOrder'].tolist())}")
            
            # 验证ValueType值
            if 'ValueType' in df.columns:
                valid_value_types = ['Signed', 'Unsigned']
                invalid_value_types = df[~df['ValueType'].isin(valid_value_types)]
                if not invalid_value_types.empty:
                    self.errors.append(f"Sheet '{sheet_name}' 存在无效的ValueType值（必须是'Signed'或'Unsigned'）: {', '.join(invalid_value_types['ValueType'].tolist())}")
            
            # 验证StartBit和SignalLength是否为非负数
            if 'StartBit' in df.columns:
                invalid_start_bits = df[df['StartBit'] < 0]
                if not invalid_start_bits.empty:
                    self.errors.append(f"Sheet '{sheet_name}' 存在无效的StartBit值（必须是非负数）: {', '.join(map(str, invalid_start_bits['StartBit'].tolist()))}")
            
            if 'SignalLength' in df.columns:
                invalid_lengths = df[df['SignalLength'] <= 0]
                if not invalid_lengths.empty:
                    self.errors.append(f"Sheet '{sheet_name}' 存在无效的SignalLength值（必须是正数）: {', '.join(map(str, invalid_lengths['SignalLength'].tolist()))}")
    
    def generate_report(self) -> str:
        """
        生成验证报告
        
        Returns:
            str: 验证报告
        """
        report = []
        report.append(f"DBC Excel文件验证报告")
        report.append(f"文件路径: {self.excel_file}")
        report.append("=" * 50)
        
        if self.errors:
            report.append("\n错误:")
            for i, error in enumerate(self.errors, 1):
                report.append(f"  {i}. {error}")
        
        if self.warnings:
            report.append("\n警告:")
            for i, warning in enumerate(self.warnings, 1):
                report.append(f"  {i}. {warning}")
        
        if not self.errors and not self.warnings:
            report.append("\n验证通过: 所有检查项均符合DBC Excel文件格式规范")
        
        report.append("\n" + "=" * 50)
        
        return "\n".join(report)
    
    def get_sheet_info(self) -> Dict[str, List[str]]:
        """
        获取Excel文件中所有Sheet的信息
        
        Returns:
            Dict[str, List[str]]: 包含Sheet名称和支持的Sheet类型的字典
        """
        if not self.load_excel():
            return {'sheet_names': [], 'supported_types': list(self.EXPECTED_STRUCTURES.keys())}
        
        return {
            'sheet_names': self.sheet_names,
            'supported_types': list(self.EXPECTED_STRUCTURES.keys())
        }

def main():
    """主函数"""
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='DBC Excel文件格式验证工具')
    parser.add_argument('--file', '-f', help='Excel文件路径')
    parser.add_argument('--sheet', '-s', help='要验证的Sheet名称')
    parser.add_argument('--type', '-t', help='预期的Sheet类型（如果不指定，将根据Sheet名称自动判断或交互式选择）')
    
    args = parser.parse_args()
    
    # 交互式输入
    file_path = args.file
    if not file_path:
        file_path = input("请输入Excel文件路径: ")
    
    # 移除路径中的引号
    file_path = file_path.strip('"\'')
    
    # 创建验证器实例
    verifier = DBCExcelVerifier(file_path)
    
    # 加载Excel文件，获取Sheet信息
    if not verifier.load_excel():
        # 加载失败，直接执行验证（会生成错误报告）
        report = verifier.generate_report()
        print(report)
        exit(1)
    
    # 选择Sheet
    sheet_name = args.sheet
    if not sheet_name:
        print(f"\nExcel文件中包含以下Sheet:")
        for i, sheet in enumerate(verifier.sheet_names, 1):
            print(f"  {i}. {sheet}")
        
        while True:
            try:
                choice = int(input(f"请选择要验证的Sheet (1-{len(verifier.sheet_names)}): ")) - 1
                if 0 <= choice < len(verifier.sheet_names):
                    sheet_name = verifier.sheet_names[choice]
                    break
                else:
                    print(f"无效选项，请输入1-{len(verifier.sheet_names)}之间的数字")
            except ValueError:
                print("请输入有效的数字")
    
    # 移除Sheet名称中的引号（如果有的话）
    sheet_name = sheet_name.strip('"\'')
    
    # 选择Sheet类型
    sheet_type = args.type
    if not sheet_type:
        # 尝试自动识别类型
        sheet_name_lower = sheet_name.lower()
        for supported_type in verifier.EXPECTED_STRUCTURES:
            if supported_type.lower() in sheet_name_lower:
                sheet_type = supported_type
                print(f"\n自动识别Sheet类型为: {sheet_type}")
                break
        
        # 如果仍未识别到，提供交互式选择
        if not sheet_type:
            print(f"\n无法自动识别Sheet '{sheet_name}' 的类型，请从以下选项中选择:")
            supported_types = list(verifier.EXPECTED_STRUCTURES.keys())
            for i, type_name in enumerate(supported_types, 1):
                print(f"  {i}. {type_name}")
            
            while True:
                try:
                    choice = int(input(f"请输入选项 (1-{len(supported_types)}): ")) - 1
                    if 0 <= choice < len(supported_types):
                        sheet_type = supported_types[choice]
                        break
                    else:
                        print(f"无效选项，请输入1-{len(supported_types)}之间的数字")
                except ValueError:
                    print("请输入有效的数字")
    
    # 执行验证
    passed = verifier.verify_sheet(sheet_name, sheet_type)
    
    # 生成报告
    report = verifier.generate_report()
    print(report)
    
    # 返回适当的退出码
    exit(0 if passed else 1)

if __name__ == '__main__':
    main()