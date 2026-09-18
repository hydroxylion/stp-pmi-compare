# -*- coding: utf-8 -*-
"""Streamlit 启动器（供 PyCharm 运行配置调用）。

等价于命令行：streamlit run pmi_compare_work.py
好处：不依赖 IDE 的「模块名」运行模式，PyCharm / 命令行都能稳定启动。
"""
import os
import sys

from streamlit.web import cli as stcli

APP_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pmi_compare_work.py")

if __name__ == "__main__":
    sys.argv = ["streamlit", "run", APP_PATH] + sys.argv[1:]
    sys.exit(stcli.main())
