#!/usr/bin/env python3
# 主程序入口
"""
DBC生成器主程序入口，只支持命令行模式
"""

import sys


def main():
    """
    主程序入口
    """
    # 只支持命令行模式，直接调用CLI主函数
    from cli.commands import main as cli_main
    return cli_main()


if __name__ == "__main__":
    sys.exit(main())