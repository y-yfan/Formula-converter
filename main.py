# -*- coding: utf-8 -*-
"""公式转换工具入口"""
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))
from formula_converter_gui import main

if __name__ == '__main__':
    main()
