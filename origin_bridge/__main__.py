# -*- coding: utf-8 -*-
"""支持 ``python -m origin_bridge`` 方式调用命令行工具。"""

import sys

from .cli import main

if __name__ == "__main__":
    sys.exit(main())
