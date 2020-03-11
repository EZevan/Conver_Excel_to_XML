# coding:utf-8

import os
import sys
reload(sys)

from enum import Enum
sys.setdefaultencoding('utf-8')


class Significance(Enum):
    high = "高"
    medium = "中"
    low = "低"


class ExecMode(Enum):
    auto = "自动"
    manual = "手动"
