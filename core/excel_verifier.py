# Excel验证器模块
"""
负责验证Excel文件是否符合DBC生成的格式规范
"""

import pandas as pd
import os
from typing import Dict, List, Tuple, Optional

class ExcelVerifier:
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
        self.sheet_data = {}
    
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
    
    def _is_single_sheet_dbc(self, df: pd.DataFrame) -> bool:
        """
        检测是否为单sheet DBC文件
        
        Args:
            df: Sheet数据
            
        Returns:
            bool: 是否为单sheet DBC文件
        """
        # 检查是否包含DBC元素标识列
        # 单sheet DBC文件通常有一个列用于标识元素类型（NODE, MESSAGE, SIGNAL等）
        
        # 将列名转为小写
        columns_lower = [str(col).lower() for col in df.columns]
        
        # 检查是否包含元素类型相关的列
        element_indicators = ['element', 'type', 'dbc', 'node', 'message', 'signal']
        for col in columns_lower:
            if any(indicator in col for indicator in element_indicators):
                return True
        
        # 检查数据中是否包含DBC元素关键字
        # 检查前20行数据中是否包含DBC元素标识
        for i in range(min(20, len(df))):
            row = df.iloc[i].dropna().tolist()
            for cell in row:
                if isinstance(cell, str):
                    cell_lower = cell.lower()
                    if any(indicator in cell_lower for indicator in ['node', 'message', 'signal', 'dbc_element']):
                        return True
        
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
            
            # 3. 如果仍未匹配到，尝试检测是否为单sheet DBC文件
            if not sheet_type:
                if self._is_single_sheet_dbc(df):
                    sheet_type = 'SINGLE_SHEET_DBC'
                else:
                    supported_types = ', '.join(self.EXPECTED_STRUCTURES.keys()) + ', SINGLE_SHEET_DBC'
                    self.errors.append(f"Sheet '{sheet_name}' 类型无法自动识别，请使用 --type 参数指定类型。支持的类型: {supported_types}")
                    return False
        
        # 执行验证
        if sheet_type == 'SINGLE_SHEET_DBC':
            # 验证单sheet DBC格式
            return self._verify_single_sheet_dbc(sheet_name, df)
        elif sheet_type in self.EXPECTED_STRUCTURES:
            # 验证传统多sheet格式
            return self._verify_traditional_sheet(sheet_name, sheet_type, df)
        else:
            supported_types = ', '.join(self.EXPECTED_STRUCTURES.keys()) + ', SINGLE_SHEET_DBC'
            self.errors.append(f"不支持的Sheet类型: {sheet_type}。支持的类型: {supported_types}")
            return False
    
    def _verify_traditional_sheet(self, sheet_name: str, sheet_type: str, df: pd.DataFrame) -> bool:
        """
        验证传统多sheet格式的Sheet
        
        Args:
            sheet_name: Sheet名称
            sheet_type: Sheet类型
            df: Sheet数据
            
        Returns:
            bool: 是否通过验证
        """
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
    
    def _verify_single_sheet_dbc(self, sheet_name: str, df: pd.DataFrame) -> bool:
        """
        验证单sheet DBC格式
        
        Args:
            sheet_name: Sheet名称
            df: Sheet数据
            
        Returns:
            bool: 是否通过验证
        """
        print(f"\n检测到单sheet DBC格式，开始验证Sheet '{sheet_name}'...")
        
        # 重置索引，方便处理
        df = df.reset_index(drop=True)
        
        # 统计各元素类型数量
        element_counts = {
            'nodes': 0,
            'messages': 0,
            'signals': 0
        }
        
        # 遍历所有行，识别和验证DBC元素
        for index, row in df.iterrows():
            # 跳过空行
            if row.isnull().all():
                continue
            
            # 将行数据转换为列表
            row_values = row.dropna().tolist()
            if not row_values:
                continue
            
            # 检查第一列是否包含元素类型标识
            first_cell = str(row_values[0]).lower()
            
            # 识别元素类型
            if any(indicator in first_cell for indicator in ['node', 'ecu']):
                # 验证节点
                if self._verify_single_sheet_node(row, index):
                    element_counts['nodes'] += 1
            elif any(indicator in first_cell for indicator in ['message', 'msg', 'frame']):
                # 验证消息
                if self._verify_single_sheet_message(row, index):
                    element_counts['messages'] += 1
            elif any(indicator in first_cell for indicator in ['signal', 'sig']):
                # 验证信号
                if self._verify_single_sheet_signal(row, index):
                    element_counts['signals'] += 1
        
        # 输出验证结果统计
        print(f"\n单sheet DBC验证结果统计：")
        print(f"- 节点数量: {element_counts['nodes']}")
        print(f"- 消息数量: {element_counts['messages']}")
        print(f"- 信号数量: {element_counts['signals']}")
        
        return len(self.errors) == 0
    
    def _verify_single_sheet_node(self, row: pd.Series, row_index: int) -> bool:
        """
        验证单sheet中的节点元素
        
        Args:
            row: 行数据
            row_index: 行索引
            
        Returns:
            bool: 是否通过验证
        """
        # 验证节点
        # 节点至少需要包含节点名称
        row_values = row.dropna().tolist()
        if len(row_values) < 2:
            self.errors.append(f"行 {row_index+2}: 节点定义不完整，缺少必要信息")
            return False
        
        node_name = str(row_values[1]).strip()
        if not node_name:
            self.errors.append(f"行 {row_index+2}: 节点名称为空")
            return False
        
        return True
    
    def _verify_single_sheet_message(self, row: pd.Series, row_index: int) -> bool:
        """
        验证单sheet中的消息元素
        
        Args:
            row: 行数据
            row_index: 行索引
            
        Returns:
            bool: 是否通过验证
        """
        # 验证消息
        # 消息至少需要包含消息ID、名称、DLC和发送节点
        row_values = row.dropna().tolist()
        if len(row_values) < 5:
            self.errors.append(f"行 {row_index+2}: 消息定义不完整，缺少必要信息")
            return False
        
        message_id = str(row_values[1]).strip()
        message_name = str(row_values[2]).strip()
        dlc = str(row_values[3]).strip()
        sender = str(row_values[4]).strip()
        
        # 验证消息ID
        if not message_id:
            self.errors.append(f"行 {row_index+2}: 消息ID为空")
            return False
        
        # 验证消息名称
        if not message_name:
            self.errors.append(f"行 {row_index+2}: 消息名称为空")
            return False
        
        # 验证DLC
        try:
            dlc_int = int(dlc)
            if not 0 <= dlc_int <= 8:
                self.errors.append(f"行 {row_index+2}: 消息 '{message_name}' DLC值 {dlc} 无效，必须在0-8之间")
                return False
        except ValueError:
            self.errors.append(f"行 {row_index+2}: 消息 '{message_name}' DLC值 '{dlc}' 不是有效整数")
            return False
        
        # 验证发送节点
        if not sender:
            self.errors.append(f"行 {row_index+2}: 消息 '{message_name}' 发送节点为空")
            return False
        
        return True
    
    def _verify_single_sheet_signal(self, row: pd.Series, row_index: int) -> bool:
        """
        验证单sheet中的信号元素
        
        Args:
            row: 行数据
            row_index: 行索引
            
        Returns:
            bool: 是否通过验证
        """
        # 验证信号
        # 信号至少需要包含信号名称、所属消息、起始位、长度、字节序和值类型
        row_values = row.dropna().tolist()
        if len(row_values) < 6:
            self.errors.append(f"行 {row_index+2}: 信号定义不完整，缺少必要信息")
            return False
        
        signal_name = str(row_values[1]).strip()
        message_name = str(row_values[2]).strip()
        start_bit = str(row_values[3]).strip()
        signal_length = str(row_values[4]).strip()
        byte_order = str(row_values[5]).strip()
        value_type = str(row_values[6]).strip() if len(row_values) > 6 else ""
        
        # 验证信号名称
        if not signal_name:
            self.errors.append(f"行 {row_index+2}: 信号名称为空")
            return False
        
        # 验证所属消息
        if not message_name:
            self.errors.append(f"行 {row_index+2}: 信号 '{signal_name}' 所属消息为空")
            return False
        
        # 验证起始位
        try:
            start_bit_int = int(start_bit)
            if start_bit_int < 0:
                self.errors.append(f"行 {row_index+2}: 信号 '{signal_name}' 起始位 {start_bit} 无效，必须是非负数")
                return False
        except ValueError:
            self.errors.append(f"行 {row_index+2}: 信号 '{signal_name}' 起始位 '{start_bit}' 不是有效整数")
            return False
        
        # 验证信号长度
        try:
            signal_length_int = int(signal_length)
            if signal_length_int <= 0 or signal_length_int > 64:
                self.errors.append(f"行 {row_index+2}: 信号 '{signal_name}' 长度 {signal_length} 无效，必须在1-64之间")
                return False
        except ValueError:
            self.errors.append(f"行 {row_index+2}: 信号 '{signal_name}' 长度 '{signal_length}' 不是有效整数")
            return False
        
        # 验证字节序
        if byte_order:
            valid_byte_orders = ['motorola', 'intel']
            if byte_order.lower() not in valid_byte_orders:
                self.errors.append(f"行 {row_index+2}: 信号 '{signal_name}' 字节序 '{byte_order}' 无效，必须是'Motorola'或'Intel'")
                return False
        
        # 验证值类型
        if value_type:
            valid_value_types = ['signed', 'unsigned']
            if value_type.lower() not in valid_value_types:
                self.errors.append(f"行 {row_index+2}: 信号 '{signal_name}' 值类型 '{value_type}' 无效，必须是'Signed'或'Unsigned'")
                return False
        
        return True
    
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