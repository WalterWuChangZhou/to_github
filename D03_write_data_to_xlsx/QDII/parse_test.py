import pandas as pd
from datetime import datetime, date
import locale

# 设置英文 locale（可选）
try:
    locale.setlocale(locale.LC_TIME, 'en_US.UTF-8')
except:
    pass  # 如果系统不支持，继续


def parse_date(value, dayfirst=False):
    """
    通用日期解析器

    Args:
        value: 任意类型的输入
        dayfirst: bool, 默认 False（月-日优先），设为 True 表示日-月优先

    Returns:
        datetime.date 对象，解析失败返回 None
    """
    # 1. datetime 对象
    if isinstance(value, datetime):
        return value.date()

    # 2. date 对象
    elif isinstance(value, date):
        return value

    # 3. Pandas Timestamp
    elif isinstance(value, pd.Timestamp):
        return value.date()

    # 4. 字符串
    elif isinstance(value, str):
        value = value.strip()
        if not value:
            return None

        # 常用日期格式（按优先级排列）
        formats = [
            '%Y-%m-%d',
            '%Y/%m/%d',
            '%m/%d/%Y' if not dayfirst else '%d/%m/%Y',  # 根据参数选择
            '%d/%m/%Y' if not dayfirst else '%m/%d/%Y',
            '%Y%m%d',
            '%b %d, %Y',  # 英文月份缩写
            '%B %d, %Y',  # 英文月份全称
        ]

        for fmt in formats:
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                continue

        # pandas 万能解析（支持更多格式）
        try:
            return pd.to_datetime(value, dayfirst=dayfirst).date()
        except (ValueError, TypeError):
            return None

    # 5. 数字（Excel 日期序列号）
    elif isinstance(value, (int, float)):
        # Excel 日期序列号范围：1-100000（1900-2173年）
        if 1 <= value <= 100000:
            try:
                # Excel 1900 日期系统（修正闰年bug）
                return pd.to_datetime(value, unit='D', origin='1899-12-30').date()
            except (ValueError, OverflowError):
                return None
        else:
            return None

    # 6. 其他类型（最后的尝试）
    else:
        try:
            return pd.to_datetime(str(value)).date()
        except:
            return None


# 测试
if __name__ == "__main__":
    test_cases = [
        ("2026-05-10", "2026-05-10"),
        ("2026/05/10", "2026-05-10"),
        ("05/10/2026", "2026-05-10"),
        ("10/05/2026", "2026-10-05"),  # 月-日优先
        ("20260510", "2026-05-10"),
        ("Jun 26, 2026", "2026-06-26"),
        ("June 26, 2026", "2026-06-26"),
        (46100, "2026-04-01"),  # Excel 日期序列号
        (datetime(2026, 5, 10), "2026-05-10"),
        (date(2026, 5, 10), "2026-05-10"),
        (pd.Timestamp('2026-05-10'), "2026-05-10"),
        (None, None),
        ("invalid", None),
    ]

    print("默认模式（月-日优先）：")
    for input_val, expected in test_cases:
        result = parse_date(input_val)
        result_str = result.strftime('%Y-%m-%d') if result else None
        status = "✅" if result_str == expected else "❌"
        print(f"  {status} {repr(input_val):<25} -> {result_str}")

    print("\n日-月优先模式：")
    result = parse_date("10/05/2026", dayfirst=True)
    print(f"  ✅ '10/05/2026' -> {result.strftime('%Y-%m-%d')} (日-月优先)")

    my_datetime = datetime.today()
    my_datetime2 = datetime.now().date()
    print(my_datetime)
    print(my_datetime2)