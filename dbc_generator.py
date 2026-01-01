#!/usr/bin/env python3
"""
DBC文件生成工具
用于将通讯矩阵或DBC Excel文件转换为DBC格式文件
"""

import argparse
import pandas as pd
import os
from typing import Dict, List, Tuple, Optional
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sys

class DBCGenerator:
    """DBC文件生成器"""
    
    def __init__(self, matrix_file: str):
        """
        初始化生成器
        
        Args:
            matrix_file: 通讯矩阵文件路径
        """
        self.matrix_file = matrix_file
        self.matrix_data = None
        self.dbc_content = ""
        # 自动识别的节点类型列
        self.node_columns = []
        # 自动识别的节点类型
        self.available_node_types = []
        # 用户指定的控制器名称
        self.controller_name = None
        # 用户指定的CAN总线类型
        self.can_bus_type = None
    
    def load_matrix(self, password: str = None) -> bool:
        """
        加载通讯矩阵文件
        
        Args:
            password: Excel文件密码，如有密码保护
            
        Returns:
            bool: 是否成功加载
        """
        try:
            if not os.path.exists(self.matrix_file):
                print(f"通讯矩阵文件不存在: {self.matrix_file}")
                return False
            
            # 读取Excel文件的Matrix工作表
            data_read_success = False
            
            try:
                # 先尝试无密码使用xlwings读取（更兼容各种Excel格式）
                import xlwings as xw
                
                # 启动Excel应用程序（不可见模式）
                app = None
                try:
                    app = xw.App(visible=False, add_book=False)
                    app.display_alerts = False
                    
                    # 密码列表，用于尝试解锁受保护的Excel文件
                    password_list = []
                    
                    # 如果用户提供了密码，先尝试该密码
                    if password:
                        password_list.append(password)
                    
                    # 添加默认尝试的三个密码
                    password_list.extend(['SGMW5050', 'Sgmw5050', 'sgmw5050'])
                    
                    # 去重，避免重复尝试相同密码
                    password_list = list(set(password_list))
                    
                    wb = None
                    password_success = False
                    
                    # 依次尝试每个密码
                    for pwd in password_list:
                        try:
                            print(f"  尝试密码: {pwd}")
                            wb = app.books.open(self.matrix_file, password=pwd)
                            password_success = True
                            print(f"  ✓ 密码 '{pwd}' 成功解锁文件")
                            break
                        except Exception as pwd_error:
                            print(f"  ✗ 密码 '{pwd}' 解锁失败: {pwd_error}")
                            continue
                    
                    # 如果所有密码都失败，抛出错误
                    if not password_success:
                        print(f"✗ 所有密码都尝试失败，无法打开受保护的Excel文件")
                        return False
                    
                    # 获取Matrix工作表
                    if 'Matrix' not in [sheet.name for sheet in wb.sheets]:
                        print(f"Excel文件中不存在Matrix工作表")
                        return False
                    
                    # 将工作表转换为pandas DataFrame
                    sheet = wb.sheets['Matrix']
                    
                    # 获取整个工作表的数据范围
                    used_range = sheet.used_range
                    
                    # 获取所有数据
                    data = used_range.value
                    
                    # 第一行作为列名，从第二行开始作为数据
                    if len(data) >= 2:
                        columns = data[0]
                        rows = data[1:]
                        self.matrix_data = pd.DataFrame(rows, columns=columns)
                        print("✓ 使用xlwings成功读取Excel文件")
                        data_read_success = True
                    else:
                        print(f"Excel文件中数据不足")
                        return False
                    
                    # 关闭工作簿
                    wb.close()
                    wb = None
                    
                    # 成功读取数据后，尝试关闭应用程序
                    if app:
                        app.quit()
                        app = None
                except Exception as wb_error:
                    # 如果已经成功读取数据，忽略关闭时的错误
                    if hasattr(self, 'matrix_data') and self.matrix_data is not None:
                        print("✓ 数据已成功读取，忽略关闭错误")
                        data_read_success = True
                        # 尝试安全关闭
                        if app:
                            try:
                                app.quit()
                            except:
                                pass
                    else:
                        raise wb_error  # 数据读取失败，抛出错误
            except Exception as e1:
                # xlwings读取失败，尝试使用pandas读取
                print(f"✗ xlwings读取失败: {e1}")
                print("尝试使用pandas读取Excel文件...")
                try:
                    # 尝试使用pandas读取，支持密码列表
                    password_list = []
                    if password:
                        password_list.append(password)
                    password_list.extend(['SGMW5050', 'Sgmw5050', 'sgmw5050'])
                    password_list = list(set(password_list))
                    
                    for pwd in password_list:
                        try:
                            print(f"  尝试密码: {pwd}")
                            self.matrix_data = pd.read_excel(self.matrix_file, sheet_name='Matrix', password=pwd)
                            print("✓ 使用pandas成功读取Excel文件")
                            data_read_success = True
                            break
                        except Exception as pwd_error:
                            print(f"  ✗ 密码 '{pwd}' 解锁失败: {pwd_error}")
                            continue
                    
                    # 如果所有密码都失败
                    if not data_read_success:
                        print(f"✗ 所有密码都尝试失败，无法打开受保护的Excel文件")
                        return False
                except Exception as e2:
                    # 所有方法都失败
                    print(f"✗ pandas读取失败: {e2}")
                    return False
            
            # 只有成功读取数据后才继续处理
            if data_read_success:
                print(f"成功加载通讯矩阵文件: {self.matrix_file}")
                print(f"数据行数: {len(self.matrix_data)}")
                print(f"列名: {list(self.matrix_data.columns)}")
                
                # 自动识别节点类型列
                self.node_columns = []
                
                for col in self.matrix_data.columns:
                    col_str = str(col).strip()
                    
                    # 直接检查列名是否为节点类型列（包含下划线且后缀为大写字母）
                    if '_' in col_str:
                        # 处理可能包含多个下划线的情况
                        parts = col_str.split('_')
                        suffix = parts[-1]
                        prefix = '_'.join(parts[:-1])
                        
                        # 检查是否为节点类型列
                        if suffix.isalpha() and suffix.isupper():
                            # 过滤配置相关列（LV开头或EV结尾）
                            if not prefix.startswith('LV') and not col_str.endswith('EV'):
                                # 添加到节点类型列表
                                self.node_columns.append(col_str)
                
                # 确保PCU_P和PCU_E被包含
                if 'PCU_P' in self.matrix_data.columns:
                    self.node_columns.append('PCU_P')
                if 'PCU_E' in self.matrix_data.columns:
                    self.node_columns.append('PCU_E')
                if 'VCU_P' in self.matrix_data.columns:
                    self.node_columns.append('VCU_P')
                if 'VCU_E' in self.matrix_data.columns:
                    self.node_columns.append('VCU_E')
                if 'VCU_C' in self.matrix_data.columns:
                    self.node_columns.append('VCU_C')
                if 'VCU_T' in self.matrix_data.columns:
                    self.node_columns.append('VCU_T')
                if 'VCU_B' in self.matrix_data.columns:
                    self.node_columns.append('VCU_B')
                
                # 去重并排序
                self.node_columns = list(set(self.node_columns))
                self.node_columns.sort()
                
                # 提取可用的节点类型
                self.available_node_types = self.node_columns.copy()
                
                print(f"自动识别的节点类型列: {self.node_columns}")
                print(f"可用的节点类型: {self.available_node_types}")
            
            return True
        except Exception as e:
            print(f"加载通讯矩阵文件失败: {str(e)}")
            return False
    
    def _parse_hex_value(self, value: str) -> int:
        """
        解析十六进制值
        
        Args:
            value: 十六进制字符串
            
        Returns:
            int: 解析后的整数值
        """
        try:
            if isinstance(value, str):
                value = value.strip().upper()
                if value.startswith('0X'):
                    return int(value, 16)
                elif value.startswith('0B'):
                    return int(value, 2)
            return int(value)
        except (ValueError, TypeError):
            return 0
    
    def _format_message_id(self, message_id: str) -> str:
        """
        格式化消息ID
        
        Args:
            message_id: 消息ID
            
        Returns:
            str: 格式化后的消息ID
        """
        message_id = str(message_id).strip().upper()
        if message_id.startswith('0X'):
            return message_id
        else:
            return f'0x{message_id}'
    
    def _extract_message_info(self) -> Dict[str, Dict]:
        """
        提取消息信息
        
        Returns:
            Dict[str, Dict]: 消息信息字典，键为消息名称
        """
        messages = {}
        
        for index, row in self.matrix_data.iterrows():
            msg_name = str(row['Msg Name\n报文名称']).strip() if pd.notna(row['Msg Name\n报文名称']) else ''
            
            if msg_name and msg_name not in messages:
                msg_id = self._format_message_id(row['Msg ID\n报文标识符'])
                msg_len = int(row['Msg Length (Byte)\n报文长度']) if pd.notna(row['Msg Length (Byte)\n报文长度']) else 8
                
                # 查找发送节点 - 使用自动识别的节点类型列
                sender = 'Unknown'  # 默认发送节点
                for col in self.node_columns:
                    if col in row and pd.notna(row[col]) and row[col] == 'Tx':
                        sender = col
                        break
                
                # 灵活处理列名，使用try-except避免KeyError
                cycle_time = 0
                try:
                    # 尝试不同的列名变体
                    for col in ['Msg Cycle Time (ms)\n报文周期时间', 'Msg Cycle Time (ms)\n报文 周期时间']:
                        if col in row and pd.notna(row[col]):
                            cycle_time = int(row[col])
                            break
                except (ValueError, TypeError):
                    cycle_time = 0
                
                messages[msg_name] = {
                    'id': msg_id,
                    'dlc': msg_len,
                    'sender': sender,
                    'cycle_time': cycle_time,
                    'signals': []
                }
        
        return messages
    
    def _extract_signal_info(self) -> List[Dict]:
        """
        提取信号信息
        
        Returns:
            List[Dict]: 信号信息列表
        """
        signals = []
        current_msg_name = ''
        
        for index, row in self.matrix_data.iterrows():
            # 获取当前行的消息名称，如果为空则使用前一行的消息名称
            msg_name = str(row['Msg Name\n报文名称']).strip() if pd.notna(row['Msg Name\n报文名称']) else ''
            if msg_name:
                current_msg_name = msg_name
            
            sig_name = str(row['Signal Name\n信号名称']).strip() if pd.notna(row['Signal Name\n信号名称']) else ''
            
            if current_msg_name and sig_name:
                # 计算起始位（Byte Order: Intel格式）
                start_byte = int(row['Start Byte\n起始字节']) if pd.notna(row['Start Byte\n起始字节']) else 0
                start_bit = int(row['Start Bit\n起始位']) if pd.notna(row['Start Bit\n起始位']) else 0
                # Intel格式：起始位 = 起始字节 * 8 + 起始位
                intel_start_bit = start_byte * 8 + start_bit
                
                # 信号长度
                sig_len = int(row['Bit Length (Bit)\n信号长度']) if pd.notna(row['Bit Length (Bit)\n信号长度']) else 8
                
                # 字节顺序
                byte_order = str(row['Byte Order\n排列格式']).strip() if pd.notna(row['Byte Order\n排列格式']) else 'Intel'
                
                # 数据类型
                value_type = str(row['Date Type\n数据类型']).strip() if pd.notna(row['Date Type\n数据类型']) else 'Unsigned'
                is_signed = 1 if value_type.lower() == 'signed' else 0
                
                # 灵活处理列名，使用try-except避免KeyError
                def get_column_value(row, col_names, default_value, conversion_func=lambda x: x):
                    """获取列值，支持多个列名变体"""
                    for col in col_names:
                        if col in row and pd.notna(row[col]):
                            try:
                                return conversion_func(row[col])
                            except (ValueError, TypeError):
                                return default_value
                    return default_value
                
                # 比例因子和偏移量
                factor = get_column_value(row, ['Factor\n比例因子', 'Factor\n比例因子 '], 1.0, float)
                offset = get_column_value(row, ['Offset\n偏移量', 'Offset\n偏移 量'], 0.0, float)
                
                # 物理最值
                min_phys = get_column_value(row, ['Signal Min. Value (phys)\n物理最小值'], 0.0, float)
                max_phys = get_column_value(row, ['Signal Max. Value (phys)\n物理最大值'], 0.0, float)
                
                # 单位
                unit = get_column_value(row, ['Unit\n单位'], '', str).strip()
                
                # 接收节点 - 使用自动识别的节点类型列
                receivers = []
                for col in self.node_columns:
                    if col in row and pd.notna(row[col]):
                        if row[col] == 'Rx':
                            receivers.append(col)
                
                signals.append({
                    'msg_name': current_msg_name,
                    'sig_name': sig_name,
                    'start_bit': intel_start_bit,
                    'signal_length': sig_len,
                    'byte_order': byte_order,
                    'is_signed': is_signed,
                    'factor': factor,
                    'offset': offset,
                    'min_phys': min_phys,
                    'max_phys': max_phys,
                    'unit': unit,
                    'receivers': receivers
                })
        
        return signals
    
    def generate_dbc(self, output_file: str, node_type: str = None, controller_name: str = None, can_bus_type: str = None) -> bool:
        """
        生成DBC文件
        
        Args:
            output_file: 输出DBC文件路径
            node_type: 指定节点类型，None表示不限制
            controller_name: 用户指定的控制器名称
            can_bus_type: 用户指定的CAN总线类型
            
        Returns:
            bool: 是否成功生成
        """
        try:
            if self.matrix_data is None:
                print("请先加载通讯矩阵文件")
                return False
            
            # 保存用户指定的控制器名称和CAN总线类型
            self.controller_name = controller_name
            self.can_bus_type = can_bus_type
            
            # 提取消息和信号信息
            messages = self._extract_message_info()
            signals = self._extract_signal_info()
            
            # 将信号分配到对应的消息
            for signal in signals:
                msg_name = signal['msg_name']
                if msg_name in messages:
                    messages[msg_name]['signals'].append(signal)
            
            # 如果指定了节点类型，过滤消息和信号
            filtered_messages = {}
            if node_type:
                for msg_name, msg_info in messages.items():
                    # 检查消息是否与指定节点类型相关
                    msg_related = False
                    
                    # 检查发送节点是否匹配
                    is_sender = msg_info['sender'] == node_type
                    if is_sender:
                        msg_related = True
                    
                    # 收集相关的信号
                    filtered_signals = []
                    for signal in msg_info['signals']:
                        # 检查信号的接收节点是否包含指定节点类型，或者该节点是发送节点
                        if node_type in signal['receivers'] or is_sender:
                            filtered_signals.append(signal)
                            msg_related = True
                    
                    if msg_related and filtered_signals:
                        filtered_messages[msg_name] = msg_info.copy()
                        filtered_messages[msg_name]['signals'] = filtered_signals
            else:
                filtered_messages = messages
            
            # 生成DBC文件内容
            # 严格遵循DBC 2.0规范，确保文件能被各种DBC编辑器正确打开
            dbc_content = 'VERSION ""\n\n\n'
            dbc_content += 'NS_ : \n NS_DESC_\n CM_\n BA_DEF_\n BA_\n VAL_\n CAT_DEF_\n CAT_\n FILTER\n BA_DEF_DEF_\n EV_DATA_\n ENVVAR_DATA_\n SGTYPE_\n SGTYPE_VAL_\n BA_DEF_SGTYPE_\n BA_SGTYPE_\n SIG_TYPE_REF_\n VAL_TABLE_\n SIG_GROUP_\n SIG_VALTYPE_\n SIGTYPE_VALTYPE_\n BO_TX_BU_\n BA_DEF_REL_\n BA_REL_\n BA_DEF_DEF_REL_\n BU_SG_REL_\n BU_EV_REL_\n BU_BO_REL_\n SG_MUL_VAL_\n\n'
            
            # 添加BS_部分，与参考DBC文件格式一致
            dbc_content += 'BS_:\n\n\n'
            
            # 添加节点定义 - 包含所有使用的节点类型
            if node_type:
                # 指定了节点类型，只添加该节点类型
                dbc_content += f'BU_: {node_type}\n\n\n'
            else:
                # 未指定节点类型，添加所有可用的节点类型
                if self.available_node_types:
                    # 所有可用节点类型用空格分隔
                    dbc_content += f'BU_: {' '.join(self.available_node_types)}\n\n\n'
                else:
                    # 默认节点类型
                    dbc_content += 'BU_: PCU_P\n\n\n'
            
            # 添加消息定义
            for msg_name, msg_info in filtered_messages.items():
                # 转换消息ID为十进制
                msg_id_dec = int(msg_info['id'], 16)
                
                # 使用Vector__XXX作为默认发送节点，与用户提供的DBC格式一致
                sender = msg_info['sender'] if msg_info['sender'] != 'Unknown' else 'Vector__XXX'
                
                # 消息定义行 - 严格按照DBC语法格式
                dbc_content += f'BO_ {msg_id_dec} {msg_name}: {msg_info["dlc"]} {sender}\n'
                
                # 添加信号定义
                for signal in msg_info['signals']:
                    # 格式化数值，去除不必要的小数位
                    def format_num(num):
                        if num == int(num):
                            return f"{int(num)}"
                        else:
                            # 最多保留3位小数
                            return f"{num:.3f}".rstrip('0').rstrip('.')
                    
                    # 字节顺序：0表示Motorola（大端），1表示Intel（小端）
                    # 从Excel中读取的byte_order是字符串，需要转换为数字
                    # 处理'Motorola Msb'、'Motorola'、'Intel'等格式
                    byte_order_str = signal['byte_order'].lower()
                    byte_order_val = 1 if 'intel' in byte_order_str else 0
                    
                    # 数据类型：+表示无符号，-表示有符号
                    value_type = '-' if signal['is_signed'] == 1 else '+'
                    
                    # 单位处理：空单位用空字符串
                    unit = signal['unit'] if signal['unit'] else ''
                    
                    # 如果指定了节点类型，只保留该节点类型作为接收节点
                    receivers = [node_type] if node_type else (self.available_node_types if self.available_node_types else ['PCU_P'])
                    
                    # 构建信号定义行 - 严格按照DBC语法格式
                    # 注意：冒号后有空格，单位引号后有空格
                    signal_def = f'SG_ {signal["sig_name"]} : {format_num(signal["start_bit"])}|{format_num(signal["signal_length"])}@{byte_order_val}{value_type} ({format_num(signal["factor"])},{format_num(signal["offset"])}) [{format_num(signal["min_phys"])}|{format_num(signal["max_phys"])}] "{unit}" '
                    # 单位引号后有空格，与参考DBC格式一致
                    signal_def += f'{" ".join(receivers)}\n'
                    
                    dbc_content += f'{signal_def}'
                
                dbc_content += '\n\n'
            
            # 添加BA_DEF_和BA_DEF_DEF_部分，与参考DBC格式一致
            dbc_content += '''BA_DEF_ BO_  "GenMsgStartDelayTime" INT 0 0;
BA_DEF_ BO_  "GenMsgDelayTime" INT 0 0;
BA_DEF_ BO_  "GenMsgNrOfRepetition" INT 0 0;
BA_DEF_ BO_  "GenMsgCycleTimeFast" INT 0 0;
BA_DEF_ BO_  "GenMsgCycleTime" INT 0 0;
BA_DEF_ BO_  "GenMsgSendType" ENUM  "Cyclic","NotUsed","NotUsed","NotUsed","NotUsed","NotUsed","NotUsed","IfActive","NoMsgSendType","NotUsed";
BA_DEF_ SG_  "GenSigStartValue" INT 0 0;
BA_DEF_ SG_  "GenSigInactiveValue" INT 0 0;
BA_DEF_ SG_  "GenSigCycleTimeActive" INT 0 0;
BA_DEF_ SG_  "GenSigCycleTime" INT 0 0;
BA_DEF_ SG_  "GenSigSendType" ENUM  "Cyclic","OnWrite","OnWriteWithRepetition","OnChange","OnChangeWithRepetition","IfActive","IfActiveWithRepetition","NoSigSendType","NotUsed","NotUsed","NotUsed","NotUsed","NotUsed";
BA_DEF_  "Baudrate" INT 0 1000000;
BA_DEF_  "BusType" STRING ;
BA_DEF_  "NmType" STRING ;
BA_DEF_  "Manufacturer" STRING ;
BA_DEF_ BO_  "TpTxIndex" INT 0 255;
BA_DEF_ BU_  "NodeLayerModules" STRING ;
BA_DEF_ BU_  "NmStationAddress" HEX 0 255;
BA_DEF_ BU_  "NmNode" ENUM  "no","yes";
BA_DEF_ BO_  "NmMessage" ENUM  "no","yes";
BA_DEF_  "NmAsrWaitBusSleepTime" INT 0 65535;
BA_DEF_  "NmAsrTimeoutTime" INT 1 65535;
BA_DEF_  "NmAsrRepeatMessageTime" INT 0 65535;
BA_DEF_ BU_  "NmAsrNodeIdentifier" HEX 0 255;
BA_DEF_ BU_  "NmAsrNode" ENUM  "no","yes";
BA_DEF_  "NmAsrMessageCount" INT 1 256;
BA_DEF_ BO_  "NmAsrMessage" ENUM  "no","yes";
BA_DEF_ BU_  "NmAsrCanMsgReducedTime" INT 1 65535;
BA_DEF_  "NmAsrCanMsgCycleTime" INT 1 65535;
BA_DEF_ BU_  "NmAsrCanMsgCycleOffset" INT 0 65535;
BA_DEF_  "NmAsrBaseAddress" HEX 0 2047;
BA_DEF_ BU_  "ILUsed" ENUM  "no","yes";
BA_DEF_  "ILTxTimeout" INT 0 65535;
BA_DEF_ SG_  "GenSigTimeoutValue" INT 0 65535;
BA_DEF_ SG_  "GenSigTimeoutTime" INT 0 65535;
BA_DEF_ BO_  "GenMsgILSupport" ENUM  "no","yes";
BA_DEF_ BO_  "GenMsgFastOnStart" INT 0 65535;
BA_DEF_ BO_  "DiagUudtResponse" ENUM  "false","true";
BA_DEF_ BO_  "DiagUudResponse" ENUM  "False","True";
BA_DEF_ BO_  "DiagState" ENUM  "no","yes";
BA_DEF_ BO_  "DiagResponse" ENUM  "no","yes";
BA_DEF_ BO_  "DiagRequest" ENUM  "no","yes";
BA_DEF_  "DBName" STRING ;
BA_DEF_DEF_  "GenMsgStartDelayTime" 0;
BA_DEF_DEF_  "GenMsgDelayTime" 0;
BA_DEF_DEF_  "GenMsgNrOfRepetition" 0;
BA_DEF_DEF_  "GenMsgCycleTimeFast" 0;
BA_DEF_DEF_  "GenMsgCycleTime" 0;
BA_DEF_DEF_  "GenMsgSendType" "Cyclic";
BA_DEF_DEF_  "GenSigStartValue" 0;
BA_DEF_DEF_  "GenSigInactiveValue" 0;
BA_DEF_DEF_  "GenSigCycleTimeActive" 0;
BA_DEF_DEF_  "GenSigCycleTime" 0;
BA_DEF_DEF_  "GenSigSendType" "Cyclic";
BA_DEF_DEF_  "Baudrate" 500000;
BA_DEF_DEF_  "BusType" "";
BA_DEF_DEF_  "NmType" "";
BA_DEF_DEF_  "Manufacturer" "Vector";
BA_DEF_DEF_  "TpTxIndex" 0;
BA_DEF_DEF_  "NodeLayerModules" " ";
BA_DEF_DEF_  "NmStationAddress" 0;
BA_DEF_DEF_  "NmNode" "no";
BA_DEF_DEF_  "NmMessage" "no";
BA_DEF_DEF_  "NmAsrWaitBusSleepTime" 1500;
BA_DEF_DEF_  "NmAsrTimeoutTime" 2000;
BA_DEF_DEF_  "NmAsrRepeatMessageTime" 3200;
BA_DEF_DEF_  "NmAsrNodeIdentifier" 50;
BA_DEF_DEF_  "NmAsrNode" "no";
BA_DEF_DEF_  "NmAsrMessageCount" 128;
BA_DEF_DEF_  "NmAsrMessage" "no";
BA_DEF_DEF_  "NmAsrCanMsgReducedTime" 320;
BA_DEF_DEF_  "NmAsrCanMsgCycleTime" 640;
BA_DEF_DEF_  "NmAsrCanMsgCycleOffset" 0;
BA_DEF_DEF_  "NmAsrBaseAddress" 1280;
BA_DEF_DEF_  "ILUsed" "no";
BA_DEF_DEF_  "ILTxTimeout" 0;
BA_DEF_DEF_  "GenSigTimeoutValue" 0;
BA_DEF_DEF_  "GenSigTimeoutTime" 0;
BA_DEF_DEF_  "GenMsgILSupport" "no";
BA_DEF_DEF_  "GenMsgFastOnStart" 0;
BA_DEF_DEF_  "DiagUudtResponse" "false";
BA_DEF_DEF_  "DiagUudResponse" "False";
BA_DEF_DEF_  "DiagState" "no";
BA_DEF_DEF_  "DiagResponse" "no";
BA_DEF_DEF_  "DiagRequest" "no";
BA_DEF_DEF_  "DBName" "";
BA_ "Manufacturer" "Vector";
BA_ "NmType" "NmAsr";
BA_ "BusType" "CAN";
BA_ "Baudrate" 500000;
BA_ "NmAsrWaitBusSleepTime" 2000;
BA_ "DBName" "Generated_DBC_File";

'''
            
            # 添加BA_部分，为消息和信号添加属性
            for msg_name, msg_info in filtered_messages.items():
                msg_id_dec = int(msg_info['id'], 16)
                dbc_content += f'BA_ "GenMsgSendType" BO_ {msg_id_dec} 0;\n'
                dbc_content += f'BA_ "GenMsgCycleTime" BO_ {msg_id_dec} 100;\n'
                
                for signal in msg_info['signals']:
                    dbc_content += f'BA_ "GenSigStartValue" SG_ {msg_id_dec} {signal["sig_name"]} 0;\n'
            
            # 添加VAL_部分，与参考DBC格式一致
            dbc_content += '\n'
            for msg_name, msg_info in filtered_messages.items():
                msg_id_dec = int(msg_info['id'], 16)
                for signal in msg_info['signals']:
                    dbc_content += f'VAL_ {msg_id_dec} {signal["sig_name"]} ;\n'
            
            dbc_content += '\n'
            
            # 确保输出文件带有.dbc后缀
            if not output_file.lower().endswith('.dbc'):
                output_file = f"{output_file}.dbc"
            
            # 确保输出目录存在
            output_dir = os.path.dirname(output_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
                print(f"已创建输出目录: {output_dir}")
            
            # 确保DBC文件使用LF行尾符，符合DBC规范
            # 替换所有\r\n为\n，然后替换剩余的\r为\n
            # 标准化行尾符为LF
            dbc_content = dbc_content.replace('\r\n', '\n').replace('\r', '\n')
            
            # 确保DBC内容只包含ASCII字符，替换或移除非ASCII字符
            def ensure_ascii(content):
                result = []
                for char in content:
                    if ord(char) < 128:
                        result.append(char)
                    else:
                        # 替换常见的非ASCII字符为ASCII等效字符
                        if char == 'Ω':
                            result.append('Ohm')
                        else:
                            # 移除其他非ASCII字符
                            continue
                return ''.join(result)
            
            # 确保内容只包含ASCII字符
            dbc_content = ensure_ascii(dbc_content)
            
            # 写入DBC文件，使用ASCII编码，确保兼容DBC规范
            with open(output_file, 'w', encoding='ascii', newline='\n') as f:
                f.write(dbc_content)
            
            print(f"成功生成DBC文件: {output_file}")
            total_signals = sum(len(msg['signals']) for msg in filtered_messages.values())
            print(f"生成的DBC文件包含 {len(filtered_messages)} 个消息和 {total_signals} 个信号")
            if node_type:
                print(f"DBC文件限制为 {node_type} 节点类型")
            return True
        except Exception as e:
            print(f"生成DBC文件失败: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def convert_from_excel(self, output_file: str, node_type: str = None, controller_name: str = None, can_bus_type: str = None, password: str = None) -> bool:
        """
        从Excel文件转换为DBC文件
        
        Args:
            output_file: 输出DBC文件路径
            node_type: 指定节点类型，None表示不限制
            controller_name: 用户指定的控制器名称
            can_bus_type: 用户指定的CAN总线类型
            password: Excel文件密码，如有密码保护
            
        Returns:
            bool: 是否成功转换
        """
        # 加载Excel文件
        if not self.load_matrix(password):
            return False
        
        # 生成DBC文件
        return self.generate_dbc(output_file, node_type, controller_name, can_bus_type)

def gui_mode():
    """GUI交互模式"""
    root = tk.Tk()
    root.title("DBC文件生成工具")
    root.geometry("600x500")
    root.resizable(True, True)
    
    # 设置字体
    font = ("Arial", 10)
    
    # 创建主框架
    main_frame = tk.Frame(root, padx=20, pady=20)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # 创建滚动条
    scrollbar = tk.Scrollbar(main_frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    # 创建画布
    canvas = tk.Canvas(main_frame, yscrollcommand=scrollbar.set)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    # 配置滚动条
    scrollbar.config(command=canvas.yview)
    
    # 创建内容框架
    content_frame = tk.Frame(canvas)
    canvas.create_window((0, 0), window=content_frame, anchor=tk.NW)
    
    # 标题
    title_label = tk.Label(content_frame, text="DBC文件生成工具", font=("Arial", 14, "bold"))
    title_label.pack(pady=10)
    
    # 输入文件路径
    file_frame = tk.Frame(content_frame)
    file_frame.pack(fill=tk.X, pady=10)
    
    file_label = tk.Label(file_frame, text="通讯矩阵文件:", font=font, width=15, anchor=tk.W)
    file_label.pack(side=tk.LEFT)
    
    file_path_var = tk.StringVar()
    file_entry = tk.Entry(file_frame, textvariable=file_path_var, font=font, width=40)
    file_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
    
    def browse_file():
        """浏览文件"""
        file_path = filedialog.askopenfilename(
            title="选择通讯矩阵文件",
            filetypes=[("Excel Files", "*.xlsx;*.xls")]
        )
        if file_path:
            file_path_var.set(file_path)
    
    browse_btn = tk.Button(file_frame, text="浏览", command=browse_file, font=font, width=10)
    browse_btn.pack(side=tk.LEFT)
    
    # 输出文件路径
    output_frame = tk.Frame(content_frame)
    output_frame.pack(fill=tk.X, pady=10)
    
    output_label = tk.Label(output_frame, text="输出DBC文件:", font=font, width=15, anchor=tk.W)
    output_label.pack(side=tk.LEFT)
    
    output_path_var = tk.StringVar()
    output_entry = tk.Entry(output_frame, textvariable=output_path_var, font=font, width=40)
    output_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
    
    def browse_output():
        """浏览输出路径"""
        output_path = filedialog.asksaveasfilename(
            title="保存DBC文件",
            defaultextension=".dbc",
            filetypes=[("DBC Files", "*.dbc")]
        )
        if output_path:
            output_path_var.set(output_path)
    
    output_btn = tk.Button(output_frame, text="浏览", command=browse_output, font=font, width=10)
    output_btn.pack(side=tk.LEFT)
    
    # 节点类型选择
    node_frame = tk.Frame(content_frame)
    node_frame.pack(fill=tk.X, pady=10)
    
    node_label = tk.Label(node_frame, text="节点类型:", font=font, width=15, anchor=tk.W)
    node_label.pack(side=tk.LEFT)
    
    node_type_var = tk.StringVar()
    node_type_combo = ttk.Combobox(node_frame, textvariable=node_type_var, font=font, width=30)
    node_type_combo.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
    
    # 总线类型映射信息
    bus_info_label = tk.Label(content_frame, text="总线类型映射: VCU_P → PCAN, VCU_E → ECAN", font=font, fg="blue")
    bus_info_label.pack(fill=tk.X, pady=5)
    
    # 密码输入
    password_frame = tk.Frame(content_frame)
    password_frame.pack(fill=tk.X, pady=10)
    
    password_label = tk.Label(password_frame, text="Excel密码:", font=font, width=15, anchor=tk.W)
    password_label.pack(side=tk.LEFT)
    
    password_var = tk.StringVar()
    password_entry = tk.Entry(password_frame, textvariable=password_var, show="*", font=font, width=40)
    password_entry.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
    
    # 操作按钮框架
    button_frame = tk.Frame(content_frame)
    button_frame.pack(fill=tk.X, pady=15)
    
    def load_file():
        """加载文件，预览节点类型"""
        try:
            # 清空日志
            log_text.delete(1.0, tk.END)
            
            # 重定向print
            sys.stdout = PrintRedirector(log_text)
            sys.stderr = PrintRedirector(log_text)
            
            # 获取参数
            file_path = file_path_var.get().strip()
            password = password_var.get().strip()
            password = password if password else None
            
            # 验证输入
            if not file_path:
                messagebox.showerror("错误", "请选择通讯矩阵文件")
                return
            
            if not os.path.exists(file_path):
                messagebox.showerror("错误", f"文件不存在: {file_path}")
                return
            
            # 创建DBC生成器实例
            original_print(f"正在加载文件: {file_path}")
            generator = DBCGenerator(file_path)
            
            # 加载Excel文件
            original_print("正在解析Excel文件...")
            if generator.load_matrix(password):
                original_print(f"文件加载成功！")
                available_types = generator.available_node_types
                if available_types:
                    original_print(f"可用的节点类型: {', '.join(available_types)}")
                    # 更新节点类型下拉列表
                    node_type_combo['values'] = available_types
                    node_type_combo.set(available_types[0] if available_types else "")
                    messagebox.showinfo("成功", f"文件加载成功！\n可用的节点类型: {', '.join(available_types)}")
                else:
                    original_print("未识别到节点类型列")
                    messagebox.showwarning("警告", "未识别到节点类型列，请检查文件格式")
            else:
                messagebox.showerror("错误", "加载Excel文件失败，请检查文件格式和密码")
        except Exception as e:
            original_print(f"发生错误: {str(e)}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("错误", f"加载文件失败: {str(e)}")
        finally:
            # 恢复原始print
            sys.stdout = sys.__stdout__
            sys.stderr = sys.__stderr__
    
    # 日志输出
    log_frame = tk.Frame(content_frame)
    log_frame.pack(fill=tk.BOTH, pady=10, expand=True)
    
    log_label = tk.Label(log_frame, text="日志输出:", font=font, anchor=tk.W)
    log_label.pack(fill=tk.X)
    
    log_text = tk.Text(log_frame, font=font, height=15, wrap=tk.WORD, yscrollcommand=scrollbar.set)
    log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    
    # 重定向print到日志
    class PrintRedirector:
        def __init__(self, text_widget):
            self.text_widget = text_widget
        
        def write(self, text):
            self.text_widget.insert(tk.END, text)
            self.text_widget.see(tk.END)
        
        def flush(self):
            pass
    
    # 保存原始print
    original_print = print
    
    def generate_dbc():
        """生成DBC文件"""
        try:
            # 清空日志
            log_text.delete(1.0, tk.END)
            
            # 重定向print
            sys.stdout = PrintRedirector(log_text)
            sys.stderr = PrintRedirector(log_text)
            
            # 获取参数
            file_path = file_path_var.get().strip()
            output_path = output_path_var.get().strip()
            node_type = node_type_var.get().strip()
            password = password_var.get().strip()
            password = password if password else None
            
            # 验证输入
            if not file_path:
                messagebox.showerror("错误", "请选择通讯矩阵文件")
                return
            
            if not os.path.exists(file_path):
                messagebox.showerror("错误", f"文件不存在: {file_path}")
                return
            
            if not output_path:
                # 自动根据节点类型命名
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                # 总线类型映射 - 支持多种总线类型
                # 从节点类型中提取总线类型后缀
                if node_type:
                    # 支持 VCU_P, VCU_E, VCU_C 等格式
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
                        # 直接使用节点类型作为总线类型
                        bus_type = f"{node_type}_CAN"
                else:
                    bus_type = "CAN"
                output_path = f"{base_name}_{bus_type}.dbc"
            
            # 确保输出目录存在
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            # 创建DBC生成器实例
            original_print(f"正在加载文件: {file_path}")
            generator = DBCGenerator(file_path)
            
            # 加载Excel文件
            original_print("正在解析Excel文件...")
            if not generator.load_matrix(password):
                messagebox.showerror("错误", "加载Excel文件失败，请检查文件格式和密码")
                return
            
            # 显示可用节点类型
            available_types = generator.available_node_types
            if available_types:
                original_print(f"可用的节点类型: {', '.join(available_types)}")
            
            # 生成DBC文件
            original_print("正在生成DBC文件...")
            if generator.generate_dbc(output_path, node_type):
                messagebox.showinfo("成功", f"DBC文件生成成功: {output_path}")
            else:
                messagebox.showerror("错误", "生成DBC文件失败")
                return
            
        except Exception as e:
            original_print(f"发生错误: {str(e)}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("错误", f"生成失败: {str(e)}")
        finally:
            # 恢复原始print
            sys.stdout = sys.__stdout__
            sys.stderr = sys.__stderr__
    
    load_btn = tk.Button(button_frame, text="加载文件", command=load_file, font=font, width=15, bg="#2196F3", fg="white")
    load_btn.pack(side=tk.LEFT, padx=5)
    
    generate_btn = tk.Button(button_frame, text="生成DBC文件", command=generate_dbc, font=("Arial", 11, "bold"), bg="#4CAF50", fg="white", width=20)
    generate_btn.pack(side=tk.LEFT, padx=5)
    
    # 退出按钮
    def exit_app():
        """退出应用"""
        root.destroy()
    
    exit_btn = tk.Button(content_frame, text="退出", command=exit_app, font=font, width=20)
    exit_btn.pack(pady=20)
    
    # 绑定关闭事件
    def on_closing():
        """关闭事件"""
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # 更新滚动区域
    def update_scrollregion():
        """更新滚动区域"""
        content_frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))
    
    content_frame.bind("<Configure>", lambda e: update_scrollregion())
    
    # 运行主循环
    root.mainloop()

def main():
    """主函数"""
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='DBC文件生成工具')
    parser.add_argument('--file', '-f', help='通讯矩阵文件路径（Excel格式）')
    parser.add_argument('--output', '-o', help='输出DBC文件路径', default='output.dbc')
    parser.add_argument('--node-type', '-n', help='指定节点类型（从Excel文件中自动识别）')
    parser.add_argument('--controller', '-c', help='指定控制器名称')
    parser.add_argument('--can-bus', '-b', help='指定CAN总线类型（如P, E, T, B等）')
    parser.add_argument('--password', '-p', help='Excel文件密码，如有密码保护')
    parser.add_argument('--gui', help='使用GUI模式', action='store_true')
    
    args = parser.parse_args()
    
    # 如果指定了--gui参数，使用GUI模式
    if args.gui or len(sys.argv) == 1:  # 双击运行时使用GUI模式
        gui_mode()
    else:
        # 命令行模式
        file_path = args.file.strip('"\'') if args.file else None
        output_path = args.output.strip('"\'')
        
        if not file_path:
            print("错误: 请指定通讯矩阵文件路径")
            print("使用示例: python dbc_generator.py --file input.xlsx --output output.dbc")
            return
        
        # 如果使用默认输出路径，根据节点类型自动命名
        if output_path == "output.dbc" and args.node_type:
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            # 总线类型映射 - 支持多种总线类型
            # 从节点类型中提取总线类型后缀
            node_type = args.node_type
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
                # 直接使用节点类型作为总线类型
                bus_type = f"{node_type}_CAN"
            output_path = f"{base_name}_{bus_type}.dbc"
        
        # 创建DBC生成器实例
        generator = DBCGenerator(file_path)
        
        # 加载Excel文件
        if not generator.load_matrix(args.password):
            print("错误: 加载Excel文件失败")
            return
        
        # 验证节点类型
        node_type = args.node_type
        if node_type and node_type not in generator.available_node_types:
            print(f"警告：指定的节点类型 '{node_type}' 不在自动识别的列表中")
            print(f"可用的节点类型: {', '.join(generator.available_node_types)}")
            # 允许使用用户指定的节点类型
        
        # 生成DBC文件
        if generator.generate_dbc(output_path, node_type, args.controller, args.can_bus):
            print("DBC文件生成成功")
        else:
            print("DBC文件生成失败")
    
    # 命令行模式下等待用户按任意键退出（如果不是GUI模式）
    if len(sys.argv) > 1 and not args.gui:
        print("\n转换完成！")
        input("按任意键退出...")

if __name__ == '__main__':
    main()