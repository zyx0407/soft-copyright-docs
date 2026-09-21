"""描述性统计、异常值筛查与线性回归。

只用标准库实现，不依赖 numpy，方便在没有科学计算环境的机器上跑。
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field


@dataclass
class Summary:
    """一列数据的描述性统计结果。"""

    name: str
    count: int = 0
    mean: float = 0.0
    std: float = 0.0
    median: float = 0.0
    q1: float = 0.0
    q3: float = 0.0
    minimum: float = 0.0
    maximum: float = 0.0
    cv: float = 0.0
    outliers: list[float] = field(default_factory=list)

    @property
    def iqr(self) -> float:
        return self.q3 - self.q1


def to_numbers(values) -> list[float]:
    """把任意序列转成浮点数列表，跳过无法转换的项。"""
    out = []
    for v in values:
        if v is None or v == "":
            continue
        try:
            out.append(float(v))
        except (TypeError, ValueError):
            continue
    return out


def percentile(sorted_values: list[float], q: float) -> float:
    """线性插值法求分位数。q 取 0~1。"""
    if not sorted_values:
        raise ValueError("空序列没有分位数")
    if len(sorted_values) == 1:
        return sorted_values[0]

    pos = (len(sorted_values) - 1) * q
    low = math.floor(pos)
    high = math.ceil(pos)
    if low == high:
        return sorted_values[int(pos)]

    weight = pos - low
    return sorted_values[low] * (1 - weight) + sorted_values[high] * weight


def summarize(name: str, values) -> Summary:
    """计算一列的完整描述性统计。"""
    nums = to_numbers(values)
    if not nums:
        return Summary(name=name)

    ordered = sorted(nums)
    n = len(ordered)
    mean = sum(ordered) / n

    if n > 1:
        variance = sum((x - mean) ** 2 for x in ordered) / (n - 1)
        std = math.sqrt(variance)
    else:
        std = 0.0

    result = Summary(
        name=name,
        count=n,
        mean=mean,
        std=std,
        median=percentile(ordered, 0.5),
        q1=percentile(ordered, 0.25),
        q3=percentile(ordered, 0.75),
        minimum=ordered[0],
        maximum=ordered[-1],
    )
    result.cv = (std / mean) if mean else 0.0
    return result


def find_outliers(values, method: str = "iqr", k: float = 1.5) -> list[float]:
    """找出异常值。

    method='iqr'：超出 Q1 - k*IQR 或 Q3 + k*IQR 的点（默认 k=1.5）
    method='sigma'：偏离均值超过 k 倍标准差的点（默认 k=1.5，常用 3）
    """
    nums = to_numbers(values)
    if len(nums) < 4:
        return []

    if method == "iqr":
        s = summarize("", nums)
        low = s.q1 - k * s.iqr
        high = s.q3 + k * s.iqr
        return [x for x in nums if x < low or x > high]

    if method == "sigma":
        if len(nums) < 2:
            return []
        mean = sum(nums) / len(nums)
        std = math.sqrt(sum((x - mean) ** 2 for x in nums) / (len(nums) - 1))
        if std == 0:
            return []
        return [x for x in nums if abs(x - mean) > k * std]

    raise ValueError(f"不认识的判据：{method}（只支持 'iqr' 和 'sigma'）")


def t_test(group_a, group_b) -> dict:
    """两独立样本 t 检验（假定方差齐）。"""
    a = to_numbers(group_a)
    b = to_numbers(group_b)
    if len(a) < 2 or len(b) < 2:
        raise ValueError("每组至少需要 2 个数据点")

    ma, mb = sum(a) / len(a), sum(b) / len(b)
    va = sum((x - ma) ** 2 for x in a) / (len(a) - 1)
    vb = sum((x - mb) ** 2 for x in b) / (len(b) - 1)

    pooled = (((len(a) - 1) * va + (len(b) - 1) * vb)
              / (len(a) + len(b) - 2))
    if pooled <= 0:
        return {"t": 0.0, "df": len(a) + len(b) - 2, "p": 1.0,
                "mean_a": ma, "mean_b": mb, "significant": False}

    se = math.sqrt(pooled * (1 / len(a) + 1 / len(b)))
    t = (ma - mb) / se
    df = len(a) + len(b) - 2
    p = 2 * (1 - _t_cdf(abs(t), df))

    return {
        "t": t,
        "df": df,
        "p": p,
        "mean_a": ma,
        "mean_b": mb,
        "significant": p < 0.05,
    }


def linear_regression(x_values, y_values) -> dict:
    """一元线性回归，返回斜率、截距、R² 与残差。"""
    xs = to_numbers(x_values)
    ys = to_numbers(y_values)
    n = min(len(xs), len(ys))
    if n < 3:
        raise ValueError("线性回归至少需要 3 组配对数据")

    xs, ys = xs[:n], ys[:n]
    mx, my = sum(xs) / n, sum(ys) / n

    sxx = sum((x - mx) ** 2 for x in xs)
    if sxx == 0:
        raise ValueError("自变量没有任何变化，无法拟合")

    sxy = sum((xs[i] - mx) * (ys[i] - my) for i in range(n))
    slope = sxy / sxx
    intercept = my - slope * mx

    residuals = [ys[i] - (slope * xs[i] + intercept) for i in range(n)]
    ss_res = sum(r ** 2 for r in residuals)
    ss_tot = sum((y - my) ** 2 for y in ys)
    r2 = 1 - ss_res / ss_tot if ss_tot else 0.0

    return {
        "slope": slope,
        "intercept": intercept,
        "r2": r2,
        "residuals": residuals,
        "n": n,
    }


def _t_cdf(t: float, df: int) -> float:
    """t 分布的累积分布函数，用不完全贝塔函数的连分式近似。"""
    x = df / (df + t * t)
    ib = _betacf(df / 2, 0.5, x)
    if ib == 0:
        return 0.5
    p = 0.5 * ib
    return 1 - p if t > 0 else p


def _betacf(a: float, b: float, x: float, max_iter: int = 200) -> float:
    """不完全贝塔函数的连分式展开（Lentz 算法）。"""
    tiny = 1e-30
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c = 1.0
    d = 1.0 - qab * x / qap
    if abs(d) < tiny:
        d = tiny
    d = 1.0 / d
    h = d

    for m in range(1, max_iter + 1):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d
        if abs(d) < tiny:
            d = tiny
        c = 1.0 + aa / c
        if abs(c) < tiny:
            c = tiny
        d = 1.0 / d
        h *= d * c

        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d
        if abs(d) < tiny:
            d = tiny
        c = 1.0 + aa / c
        if abs(c) < tiny:
            c = tiny
        d = 1.0 / d
        delta = d * c
        h *= delta

        if abs(delta - 1.0) < 1e-10:
            break

    return h
