import io
import base64

import numpy as np
import pandas as pd
from scipy.special import exp1


def parse_content(contents, filename):
    """解析数据表"""
    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    if 'csv' in filename:
        return pd.read_csv(io.StringIO(decoded.decode('utf-8')))
    elif 'xls' in filename:
        return pd.read_excel(io.BytesIO(decoded))


def theis_fit(t, s, Q, r, lgs, lgtr2):
    """配线法"""
    x = np.linspace(-1, 4, 51)  # 等差数列
    line_x = np.float_power(10, x)
    line_y = exp1(1 / line_x)

    tr2 = t / r ** 2

    scatter_x = tr2 * np.float_power(10, lgtr2)
    scatter_y = s * np.float_power(10, lgs)
    return line_x, line_y, scatter_x, scatter_y


def line_fit(t, s, t0, slope):
    """直线图解法"""
    line_x = t
    line_y = slope * np.log10(t / t0)
    scatter_x = t
    scatter_y = s
    return line_x, line_y, scatter_x, scatter_y


def theis_calculate(Q, lgs, lgtr2):
    """配线法计算导水系数和储水系数"""
    T = Q / 4 / np.pi * np.float_power(10, lgs)
    S = 4 * T * np.float_power(10, -lgtr2)
    return T, S


def line_calculate(Q, r, t0, slope):
    """直线图解法计算导水系数和储水系数"""
    T = 0.183 * Q / slope
    S = 2.25 * T * t0 / r ** 2
    return T, S
