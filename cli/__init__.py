# CLI模块
"""
提供命令行接口，让用户可以通过命令行使用DBC生成器
"""

from .commands import parse_args, main

__all__ = ['parse_args', 'main']