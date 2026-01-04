import os
import pandas as pd
from typing import Dict, List, Tuple, Optional

class DBCGenerator:
    """DBC文件生成器"""
    
    def __init__(self):
        """初始化DBC生成器"""
        self.matrix_data = None
        self.node_columns = []
    
    def set_data(self, matrix_data, node_columns: List[str]):
        """
        设置生成DBC所需的数据
        
        Args:
            matrix_data: 从Excel加载的矩阵数据
            node_columns: 节点类型列列表
        """
        self.matrix_data = matrix_data
        self.node_columns = node_columns
    
    def generate_dbc(self, output_file: str, node_type: str = None, 
                    controller_name: str = None, target_bus_type: str = None, 
                    is_canfd: bool = False) -> bool:
        """
        生成DBC文件，支持按总线类型生成单独文件
        
        Args:
            output_file: 输出DBC文件路径
            node_type: 节点类型
            controller_name: 控制器名称
            target_bus_type: 目标总线类型（如C、E、P等），用于过滤要生成的DBC文件
            is_canfd: 是否生成CANFD格式的DBC文件
            
        Returns:
            bool: 生成是否成功
        """
        if self.matrix_data is None or not self.node_columns:
            print("错误: 未设置足够的数据来生成DBC文件")
            return False
        
        try:
            # 分析节点类型，识别不同的总线类型
            bus_types = set()
            valid_bus_columns = []
            
            # 首先找出有实际数据的总线列
            for col in self.node_columns:
                if '_' in col:
                    parts = col.split('_')
                    if len(parts) >= 2:
                        bus_suffix = parts[-1]
                        # 检查该列是否有实际数据（Rx/Tx）
                        non_na_count = self.matrix_data[col].notna().sum()
                        if non_na_count > 0:
                            bus_types.add(bus_suffix)
                            valid_bus_columns.append(col)
            
            print(f"\n检测到的总线类型: {sorted(bus_types)}")
            print(f"有实际数据的总线列: {valid_bus_columns}")
            
            # 如果指定了目标总线类型，只生成该类型的DBC
            if target_bus_type:
                if target_bus_type in bus_types:
                    bus_types = [target_bus_type]
                else:
                    print(f"\n警告: 未找到总线类型 '{target_bus_type}' 的数据，跳过生成")
                    return False
            
            # 为每种总线类型生成单独的DBC文件
            for bus_type in bus_types:
                # 生成对应的输出文件名
                base_name, ext = os.path.splitext(output_file)
                bus_output_file = f"{base_name}_{bus_type}CAN{ext}"
                
                print(f"\n开始生成{bus_type}CAN DBC文件: {bus_output_file}")
                
                # 优化：直接写入文件，不保留完整的DBC内容在内存中
                with open(bus_output_file, 'w', encoding='utf-8') as f:
                    # 1. 生成DBC文件头并直接写入
                    f.write(self._generate_header(is_canfd))
                    
                    # 2. 生成节点定义并直接写入
                    nodes = self._extract_nodes()
                    f.write(self._generate_nodes(nodes))
                    
                    # 3. 生成消息和信号定义并直接写入
                    messages = self._extract_messages()
                    signals = self._extract_signals()
                    
                    # 合并消息和信号，一次性生成并写入
                    dbc_content = self._generate_messages_and_signals(messages, signals)
                    f.write(dbc_content)
                
                print(f"✓ {bus_type}CAN DBC文件生成成功: {bus_output_file}")
            
            return True
        except Exception as e:
            print(f"✗ 生成DBC文件失败: {e}")
            return False
    
    def _generate_header(self, is_canfd: bool = False) -> str:
        """生成DBC文件头，严格按照DBC文件格式规范"""
        header_lines = []
        
        # 1. VERSION行
        header_lines.extend([
            "VERSION \"\"",
            "",
            "",
        ])
        
        # 2. 名称空间定义部分 (NS_ : 之后的行)
        header_lines.extend([
            "NS_ : ",
            " NS_DESC_",
            " CM_",
            " BA_DEF_",
            " BA_",
            " VAL_",
            " CAT_DEF_",
            " CAT_",
            " FILTER",
            " BA_DEF_DEF_",
            " EV_DATA_",
            " ENVVAR_DATA_",
            " SGTYPE_",
            " SGTYPE_VAL_",
            " BA_DEF_SGTYPE_",
            " BA_SGTYPE_",
            " SIG_TYPE_REF_",
            " VAL_TABLE_",
            " SIG_GROUP_",
            " SIG_VALTYPE_",
            " SIGTYPE_VALTYPE_",
            " BO_TX_BU_",
            " BA_DEF_REL_",
            " BA_REL_",
            " BA_DEF_DEF_REL_",
            " BU_SG_REL_",
            " BU_EV_REL_",
            " BU_BO_REL_",
            " SG_MUL_VAL_",
            "",
        ])
        
        # 3. BS_部分 (位定时定义)
        header_lines.extend([
            "BS_:",
            "",
        ])
        
        # 4. CANFD属性定义 - 暂时注释，避免语法错误
        # if is_canfd:
        #     header_lines.extend([
        #         'BA_DEF_ BO_ "CANFD" BOOLEAN',
        #         'BA_DEF_ BO_ "BRS" BOOLEAN',
        #         'BA_DEF_ BO_ "FDF" BOOLEAN',
        #         'BA_DEF_ BO_ "FD_DLC" INT 0 15',
        #         "",
        #     ])
        
        # 5. 添加节点定义前的空行
        header_lines.extend([
            ""])
        
        return '\n'.join(header_lines)
    
    def _extract_nodes(self) -> List[str]:
        """
        从数据中提取节点列表
        
        Returns:
            List[str]: 节点列表
        """
        nodes = set()
        
        for col in self.node_columns:
            # 从列名中提取节点名称
            if '_' in col:
                parts = col.split('_')
                node_name = parts[-1].strip()
                if node_name:
                    nodes.add(node_name)
        
        # 添加所有控制器名称作为节点
        for col in self.matrix_data.columns:
            col_str = str(col)
            if 'Controller' in col_str or '控制器' in col_str:
                controllers = self.matrix_data[col].dropna().unique()
                for controller in controllers:
                    ctrl_str = str(controller).strip()
                    if ctrl_str:
                        nodes.add(ctrl_str)
                break
        
        return sorted(nodes)
    
    def _generate_nodes(self, nodes: List[str]) -> str:
        """
        生成节点定义
        
        Args:
            nodes: 节点列表
            
        Returns:
            str: 生成的节点定义内容
        """
        if nodes:
            # 确保Vector__XXX节点存在（默认发送者）
            if 'Vector__XXX' not in nodes:
                nodes.append('Vector__XXX')
            return f"BU_: {' '.join(nodes)}\n\n\n"
        return "\n\n\n"
    
    def _extract_messages(self) -> Dict[str, Dict]:
        """
        从数据中提取消息定义
        
        Returns:
            Dict[str, Dict]: 消息字典，键为消息ID
        """
        messages = {}
        
        print(f"\n正在提取消息，数据行数: {len(self.matrix_data)}")
        print(f"可用列: {list(self.matrix_data.columns)}")
        
        # 映射实际列名到预期列名
        col_mapping = {}
        
        # 查找消息ID列
        message_id_col = None
        for col in self.matrix_data.columns:
            col_str = str(col)
            if 'ID' in col_str and ('Msg' in col_str or '报文' in col_str):
                message_id_col = col
                col_mapping['MessageID'] = col
                break
        
        if not message_id_col:
            print("警告: Excel文件中没有消息ID列")
            return messages
        
        # 查找消息名称列
        message_name_col = None
        for col in self.matrix_data.columns:
            col_str = str(col)
            if ('Name' in col_str and 'Msg' in col_str) or '报文名称' in col_str:
                message_name_col = col
                col_mapping['MessageName'] = col
                break
        
        # 查找DLC列
        dlc_col = None
        for col in self.matrix_data.columns:
            col_str = str(col)
            if ('Length' in col_str and 'Msg' in col_str) or '报文长度' in col_str:
                dlc_col = col
                col_mapping['DLC'] = col
                break
        
        print(f"列映射关系: {col_mapping}")
        
        # 提取消息ID和名称
        message_count = 0
        for idx, row in self.matrix_data.iterrows():
            # 检查必要列是否存在
            message_id_val = row.get(message_id_col)
            if pd.isna(message_id_val):
                continue
            
            message_id = str(message_id_val).strip()
            if not message_id:
                continue
            
            message_count += 1
            if message_count <= 5:  # 只显示前5个消息作为示例
                print(f"  处理消息 {message_count}: ID={message_id}")
            
            # 处理消息ID格式
            try:
                if message_id.startswith('0x') or message_id.startswith('0X'):
                    # 十六进制格式
                    message_id_hex = message_id
                    message_id_dec = str(int(message_id, 16))
                else:
                    # 十进制格式
                    message_id_dec = message_id
                    message_id_hex = f"0x{int(message_id):X}"
            except ValueError:
                print(f"  跳过无效的消息ID: {message_id}")
                continue
            
            # 获取消息名称
            message_name = ''
            if message_name_col:
                message_name = str(row.get(message_name_col, '')).strip()
            
            if not message_name:
                # 使用默认消息名称
                message_name = f"MSG_{message_id_hex}"
            
            # 获取DLC
            dlc = '8'  # 默认值
            if dlc_col:
                dlc_val = row.get(dlc_col)
                if not pd.isna(dlc_val):
                    dlc_str = str(dlc_val).strip()
                    if dlc_str.isdigit():
                        dlc_num = int(dlc_str)
                        # 支持CAN FD的DLC范围（0-15）
                        if 0 <= dlc_num <= 15:
                            dlc = dlc_str
            
            # 获取发送节点 - 从节点列中确定
            sender = 'Vector__XXX'  # 默认发送节点，与参考DBC文件保持一致
            
            # 添加到消息字典
            if message_id_dec not in messages:
                messages[message_id_dec] = {
                    'id': message_id_dec,
                    'hex_id': message_id_hex,
                    'name': message_name,
                    'dlc': dlc,
                    'sender': sender,
                    'signals': []
                }
        
        print(f"共提取到 {len(messages)} 条消息")
        return messages
    
    def _generate_messages(self, messages: Dict[str, Dict]) -> str:
        """
        生成消息定义
        
        Args:
            messages: 消息字典
            
        Returns:
            str: 生成的消息定义内容
        """
        if not messages:
            return ""
        
        message_lines = []
        for msg_id, msg in messages.items():
            # 生成消息定义行
            msg_line = f"BO_ {msg_id} {msg['name']}: {msg['dlc']} {msg['sender']}"
            message_lines.append(msg_line)
        
        return '\n'.join(message_lines)
    
    def _extract_signals(self) -> Dict[str, List[Dict]]:
        """
        从矩阵格式数据中提取信号定义
        矩阵格式特点：消息头行包含MessageID，后续行包含该消息的信号
        
        Returns:
            Dict[str, List[Dict]]: 信号字典，键为消息ID
        """
        signals = {}
        
        print(f"\n正在提取信号，数据行数: {len(self.matrix_data)}")
        
        # 查找消息ID列
        message_id_col = None
        for col in self.matrix_data.columns:
            col_str = str(col)
            if 'ID' in col_str and ('Msg' in col_str or '报文' in col_str):
                message_id_col = col
                print(f"  找到消息ID列: '{col_str}'")
                break
        
        if not message_id_col:
            print("警告: 未找到消息ID列")
            return signals
        
        # 查找信号名称列
        signal_name_col = None
        for col in self.matrix_data.columns:
            col_str = str(col)
            col_lower = col_str.lower()
            if 'signal' in col_lower and ('name' in col_lower or '名称' in col_lower):
                signal_name_col = col
                print(f"  找到信号名称列: '{col_str}'")
                break
        
        if not signal_name_col:
            print("警告: 未找到信号名称列")
            return signals
        
        # 查找其他必要列
        start_bit_col = None
        start_byte_col = None
        signal_length_col = None
        byte_order_col = None
        value_type_col = None
        factor_col = None
        offset_col = None
        unit_col = None
        min_val_col = None
        max_val_col = None
        
        for col in self.matrix_data.columns:
            col_str = str(col)
            col_lower = col_str.lower()
            
            if 'start bit' in col_lower or '起始位' in col_lower:
                start_bit_col = col
            elif 'start byte' in col_lower or '起始字节' in col_lower:
                start_byte_col = col
            elif 'bit length' in col_lower or '信号长度' in col_lower:
                signal_length_col = col
            elif 'byte order' in col_lower or '排列格式' in col_lower:
                byte_order_col = col
            elif 'date type' in col_lower or '数据类型' in col_lower:
                value_type_col = col
            elif 'factor' in col_lower or '比例因子' in col_lower:
                factor_col = col
            elif 'offset' in col_lower or '偏移量' in col_lower:
                offset_col = col
            elif 'unit' in col_lower:
                unit_col = col
            elif 'min' in col_lower or '最小值' in col_lower:
                min_val_col = col
            elif 'max' in col_lower or '最大值' in col_lower:
                max_val_col = col
        
        print(f"信号提取配置:")
        print(f"  消息ID列: {message_id_col}")
        print(f"  信号名称列: {signal_name_col}")
        print(f"  起始字节列: {start_byte_col}")
        print(f"  起始位列: {start_bit_col}")
        print(f"  信号长度列: {signal_length_col}")
        print(f"  字节序列: {byte_order_col}")
        print(f"  数据类型: {value_type_col}")
        print(f"  比例因子: {factor_col}")
        print(f"  偏移量: {offset_col}")
        print(f"  单位: {unit_col}")
        print(f"  最小值: {min_val_col}")
        print(f"  最大值: {max_val_col}")
        
        # 遍历数据，按消息分组提取信号
        current_msg_id = None
        signal_count = 0
        
        for idx, row in self.matrix_data.iterrows():
            # 检查当前行是否包含消息ID（消息头行）
            msg_id_val = row.get(message_id_col)
            if not pd.isna(msg_id_val):
                # 这是消息头行，更新当前消息ID
                msg_id_str = str(msg_id_val).strip()
                if msg_id_str:
                    try:
                        if msg_id_str.startswith('0x') or msg_id_str.startswith('0X'):
                            current_msg_id = str(int(msg_id_str, 16))
                        else:
                            current_msg_id = str(int(msg_id_str))
                    except ValueError:
                        current_msg_id = None
                    print(f"  正在处理消息: {msg_id_str} -> ID={current_msg_id}")
            else:
                # 这是信号行，属于当前消息
                if current_msg_id:
                    # 获取信号名称
                    signal_name_val = row.get(signal_name_col)
                    if not pd.isna(signal_name_val):
                        signal_name = str(signal_name_val).strip()
                        if signal_name:
                            signal_count += 1
                            if signal_count <= 5:
                                print(f"    提取信号: {signal_name}")
                            
                            # 初始化信号属性
                            start_bit = '0'
                            signal_length = '8'
                            byte_order = 'Motorola'
                            value_type = 'Unsigned'
                            factor = '1.0'
                            offset = '0.0'
                            min_val = '0'
                            max_val = '0'
                            unit = ''
                            
                            # 尝试从行中获取信号属性
                            # 先提取所有信号属性，然后再计算起始位
                            
                            # 提取字节顺序
                            if byte_order_col:
                                byte_order_val = row.get(byte_order_col)
                                if not pd.isna(byte_order_val):
                                    byte_order_val = str(byte_order_val).strip().lower()
                                    if byte_order_val in ['intel', 'little endian', 'little']:
                                        byte_order = 'Intel'
                                    else:
                                        byte_order = 'Motorola'
                            
                            # 提取信号长度
                            if signal_length_col:
                                length_val = row.get(signal_length_col)
                                if not pd.isna(length_val):
                                    try:
                                        # 转换为整数，支持浮点数和字符串格式
                                        length = int(float(length_val))
                                        if length > 0:
                                            signal_length = str(length)
                                    except (ValueError, TypeError):
                                        pass
                            
                            # 提取数据类型
                            if value_type_col:
                                value_type_val = row.get(value_type_col)
                                if not pd.isna(value_type_val):
                                    value_type_val = str(value_type_val).strip().lower()
                                    if value_type_val in ['unsigned', '无符号']:
                                        value_type = 'Unsigned'
                                    else:
                                        value_type = 'Signed'
                            
                            # 计算实际的起始位：结合起始字节和起始位
                            actual_start_bit = 0
                            
                            # 获取起始字节和起始位
                            start_byte_val = row.get(start_byte_col, 0)
                            start_bit_val = row.get(start_bit_col, 0)
                            
                            if not pd.isna(start_byte_val) and not pd.isna(start_bit_val):
                                try:
                                    start_byte = int(float(start_byte_val))
                                    bit_in_byte = int(float(start_bit_val))
                                    
                                    # 按照模板DBC文件的格式，使用Intel格式（小端）计算起始位
                                    # 不管Excel矩阵中指定的字节序是什么，都使用Intel格式计算起始位
                                    actual_start_bit = start_byte * 8 + bit_in_byte
                                except (ValueError, TypeError):
                                    pass
                            
                            start_bit = str(actual_start_bit)
                            
                            # 提取比例因子
                            if factor_col:
                                factor_val = row.get(factor_col)
                                if not pd.isna(factor_val):
                                    try:
                                        factor = str(float(factor_val))
                                    except (ValueError, TypeError):
                                        pass
                            
                            # 提取偏移量
                            if offset_col:
                                offset_val = row.get(offset_col)
                                if not pd.isna(offset_val):
                                    try:
                                        offset = str(float(offset_val))
                                    except (ValueError, TypeError):
                                        pass
                            
                            # 提取单位
                            if unit_col:
                                unit_val = row.get(unit_col)
                                if not pd.isna(unit_val):
                                    unit = str(unit_val).strip()
                            
                            # 提取最小值
                            if min_val_col:
                                min_val_val = row.get(min_val_col)
                                if not pd.isna(min_val_val):
                                    try:
                                        min_val = str(float(min_val_val))
                                    except (ValueError, TypeError):
                                        pass
                            
                            # 提取最大值
                            if max_val_col:
                                max_val_val = row.get(max_val_col)
                                if not pd.isna(max_val_val):
                                    try:
                                        max_val = str(float(max_val_val))
                                    except (ValueError, TypeError):
                                        pass
                            
                            # 创建信号字典
                            signal = {
                                'name': signal_name,
                                'start_bit': start_bit,
                                'length': signal_length,
                                'byte_order': byte_order,
                                'value_type': value_type,
                                'factor': factor,
                                'offset': offset,
                                'min': min_val,
                                'max': max_val,
                                'unit': unit,
                                'receiver': 'Vector__XXX'  # 使用Vector__XXX作为默认接收者，与参考DBC保持一致
                            }
                            
                            # 添加信号到字典
                            if current_msg_id not in signals:
                                signals[current_msg_id] = []
                            signals[current_msg_id].append(signal)
        
        print(f"共提取到 {signal_count} 个信号")
        return signals
    
    def _generate_signals(self, signals: Dict[str, List[Dict]]) -> str:
        """
        生成信号定义
        
        Args:
            signals: 信号字典
            
        Returns:
            str: 生成的信号定义内容
        """
        if not signals:
            return ""
        
        signal_lines = []
        for msg_id, msg_signals in signals.items():
            for signal in msg_signals:
                # 生成信号定义行
                # 格式: SG_ SignalName : StartBit|Length@ByteOrder+ValueType (Factor,Offset) [Min|Max] "Unit" Receiver
                sg_line = f" SG_ {signal['name']} : {signal['start_bit']}|{signal['length']}@{1 if signal['byte_order'] == 'Intel' else 0}{'+' if signal['value_type'] == 'Unsigned' else '-'}\n"
                sg_line += f"            (1.0,0.0) [0|0] \"\" {signal['receiver']}"
                signal_lines.append(sg_line)
            signal_lines.append("")
        
        return '\n'.join(signal_lines)
    
    def _generate_messages_and_signals(self, messages: Dict[str, Dict], signals: Dict[str, List[Dict]]) -> str:
        """
        生成消息和信号定义
        
        Args:
            messages: 消息字典
            signals: 信号字典
            
        Returns:
            str: 生成的消息和信号定义内容
        """
        combined_content = []
        
        # 遍历所有消息
        for msg_id in messages:
            # 添加消息定义
            msg = messages[msg_id]
            msg_line = f"BO_ {msg_id} {msg['name']}: {msg['dlc']} {msg['sender']}"
            combined_content.append(msg_line)
            combined_content.append("")
            
            # 添加对应的信号定义
            if msg_id in signals:
                for signal in signals[msg_id]:
                    # 使用提取的实际信号属性值
                    factor = signal.get('factor', '1')
                    offset = signal.get('offset', '0')
                    min_val = signal.get('min', '0')
                    max_val = signal.get('max', '0')
                    unit = signal.get('unit', '')
                    
                    # 确保单位字符串是正确的格式
                    if unit is None:
                        unit = ""
                    else:
                        unit = str(unit).strip()
                    
                    # 生成信号行，格式与模板DBC文件完全一致
                    # 使用Intel格式（@1），不管Excel矩阵中指定的字节序是什么
                    sg_line = f" SG_ {signal['name']} : {signal['start_bit']}|{signal['length']}@1{'+' if signal['value_type'] == 'Unsigned' else '-'}"
                    sg_line += f" ({factor},{offset}) [{min_val}|{max_val}] \"{unit}\" {signal['receiver']}"
                    combined_content.append(sg_line)
            
            combined_content.append("")
            combined_content.append("")
        
        return '\n'.join(combined_content)
    
    def clear(self):
        """
        清空生成器状态
        """
        self.matrix_data = None
        self.node_columns = []