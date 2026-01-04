# Excel文件处理模块
"""
负责处理Excel文件的读取和数据提取
"""

import pandas as pd
from typing import Dict, List, Optional

class ExcelHandler:
    """Excel文件处理器"""
    
    def __init__(self):
        """初始化Excel处理器"""
        self.matrix_data = None
        self.node_columns = []
    
    def read_matrix_file(self, file_path: str, sheet_name: str = 'Matrix', password: str = None) -> bool:
        """
        读取Excel文件中的Matrix工作表，优化内存占用
        
        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称，默认为'Matrix'
            password: 用于解锁受保护的Excel文件（仅支持xlrd引擎）
            
        Returns:
            bool: 读取是否成功
        """
        print(f"\n开始读取Excel文件: {file_path}")
        
        if not file_path:
            print("错误: 未提供Excel文件路径")
            return False
        
        data_read_success = False
        temp_df = None
        
        try:
            # 定义需要读取的必要列
            necessary_columns = [
                'MessageID', 'MessageName', 'DLC', 'ControllerName',
                'SignalName', 'SignalLength', 'StartBit', 'ByteOrder',
                'ValueType', 'Factor', 'Offset', 'Receiver'
            ]
            
            # 首先尝试使用openpyxl引擎（不支持密码）
            try:
                print("  尝试使用openpyxl引擎读取")
                temp_df = pd.read_excel(
                    file_path,
                    engine='openpyxl',
                    sheet_name=sheet_name,
                    keep_default_na=True,
                    na_values=['', 'NA', 'None']
                )
                data_read_success = True
                print("✓ 使用openpyxl引擎成功读取Excel文件")
            except Exception as e:
                print(f"  ✗ openpyxl引擎读取失败: {e}")
                
                # 如果openpyxl失败，尝试使用xlrd引擎（支持老版本Excel）
                try:
                    print("  尝试使用xlrd引擎读取")
                    temp_df = pd.read_excel(
                        file_path,
                        engine='xlrd',
                        sheet_name=sheet_name,
                        keep_default_na=True,
                        na_values=['', 'NA', 'None']
                    )
                    data_read_success = True
                    print("✓ 使用xlrd引擎成功读取Excel文件")
                except Exception as xlrd_error:
                    print(f"  ✗ xlrd引擎读取失败: {xlrd_error}")
                    
                    # 如果用户提供了密码，尝试使用msoffcrypto-tool解锁（如果可用）
                    if password:
                        try:
                            print(f"  尝试使用msoffcrypto-tool解锁文件")
                            import msoffcrypto
                            import io
                            
                            # 读取加密文件
                            with open(file_path, 'rb') as f:
                                office_file = msoffcrypto.OfficeFile(f)
                                office_file.load_key(password=password)
                                
                                # 将解密后的内容写入内存流
                                decrypted = io.BytesIO()
                                office_file.decrypt(decrypted)
                                decrypted.seek(0)
                                
                                # 尝试读取解密后的文件
                                temp_df = pd.read_excel(
                                    decrypted,
                                    engine='openpyxl',
                                    sheet_name=sheet_name,
                                    keep_default_na=True,
                                    na_values=['', 'NA', 'None']
                                )
                                data_read_success = True
                                print(f"✓ 使用密码 '{password}' 成功解锁并读取文件")
                        except ImportError:
                            print("  ✗ msoffcrypto-tool未安装，无法解锁加密文件")
                        except Exception as crypto_error:
                            print(f"  ✗ 使用密码 '{password}' 解锁失败: {crypto_error}")
            
            # 如果所有尝试都失败
            if not data_read_success:
                print("✗ 无法读取Excel文件，请检查文件格式、路径和密码是否正确")
                return False
            
            # 优化：仅保留非空行，减少内存占用
            if data_read_success and temp_df is not None:
                # 删除完全空的行
                temp_df = temp_df.dropna(axis=0, how='all')
                
                # 将优化后的数据赋值给实例变量
                self.matrix_data = temp_df
                
                print(f"成功加载通讯矩阵文件: {file_path}")
                print(f"数据行数: {len(self.matrix_data)}")
                print(f"列名: {list(self.matrix_data.columns)}")
                return True
        except Exception as e:
            # 所有方法都失败
            print(f"✗ 读取Excel文件失败: {e}")
            return False
        
        return False
    
    def extract_node_columns(self) -> List[str]:
        """
        提取节点类型列
        
        Returns:
            List[str]: 节点类型列列表
        """
        if self.matrix_data is None:
            print("错误: 未加载任何Excel数据")
            return []
        
        node_columns = []
        
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
                    if not (prefix.startswith('LV') or prefix.endswith('EV')):
                        node_columns.append(col_str)
        
        print(f"自动识别到 {len(node_columns)} 个节点类型列")
        print(f"节点列: {node_columns}")
        
        self.node_columns = node_columns
        return node_columns
    
    def get_matrix_data(self) -> Optional[pd.DataFrame]:
        """
        获取加载的矩阵数据
        
        Returns:
            Optional[pd.DataFrame]: 矩阵数据，如果未加载则返回None
        """
        return self.matrix_data
    
    def clear_data(self):
        """
        清空加载的数据
        """
        self.matrix_data = None
        self.node_columns = []
        print("已清空加载的数据")