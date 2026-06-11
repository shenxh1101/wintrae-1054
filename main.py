#!/usr/bin/env python3
"""
公共卫生管理自动化工具
重点人群随访资料整理工具
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gui import MainWindow


def main():
    app = MainWindow()
    app.run()


if __name__ == '__main__':
    main()
