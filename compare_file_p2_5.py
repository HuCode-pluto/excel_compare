# -*- coding: utf-8 -*-
from __future__ import with_statement
import xlrd
import sys
import datetime


# ========== 👇 通用转换函数（所有任务可复用） ==========
from xlrd.timemachine import unicode


def trim_str(s):
    """去掉前后空格，处理空值"""
    if s is None or s == "":
        return None
    return unicode(s).strip()


def parse_date(d):
    """统一日期格式，支持 Excel 日期数字和字符串"""
    if d is None or d == "":
        return None
    if isinstance(d, float):
        try:
            t = xlrd.xldate_as_tuple(d, 0)
            return datetime.date(t[0], t[1], t[2])
        except:
            return None
    s = unicode(d).strip()
    for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"]:
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except:
            continue
    return None


def parse_percent(p):
    """统一百分比："98.5%" 和 0.985 都转成 98.5"""
    if p is None or p == "":
        return None
    from xlrd.timemachine import unicode
    s = unicode(p).strip()
    if s.endswith("%"):
        try:
            return float(s[:-1])
        except:
            return None
    try:
        return float(s) * 100
    except:
        return None


def normalize_fac(f):
    """厂家名称归一化"""
    if f is None or f == "":
        return None
    s = unicode(f).strip()
    if u"南瑞" in s:
        return u"南瑞"
    if u"许继" in s:
        return u"许继"
    if u"四方" in s:
        return u"四方"
    return s


# ========== 👇 批量任务配置区（只需要改这里） ==========
# 每个 {} 代表一对对比任务，复制粘贴即可添加新任务
TASKS = [
    # 任务1：终端信息对比
    {
        "name": "终端基础信息对比",  # 任务名称，输出时显示
        "file1": "old_terminal.xls",  # 第一个文件路径
        "file2": "new_terminal.xls",  # 第二个文件路径
        "col_map": {  # 列映射：{旧表列名: 新表列名}
            "终端ID": "termid",
            "终端厂家": "term_fac",
            "安装日期": "install_date",
            "在线率": "online_rate",
        },
        "transform_funcs": {  # 该任务专属转换函数
            "终端厂家": normalize_fac,
            "安装日期": parse_date,
            "在线率": parse_percent,
        },
        "key_col": "终端ID"  # 主键列
    },
    # 任务2：厂家资质对比
    {
        "name": "厂家资质信息对比",
        "file1": "old_factory.xls",
        "file2": "new_factory.xls",
        "col_map": {
            "厂家编号": "factory_id",
            "厂家名称": "factory_name",
            "资质等级": "level",
            "有效期": "valid_date",
        },
        "transform_funcs": {
            "厂家名称": trim_str,
            "资质等级": trim_str,
            "有效期": parse_date,
        },
        "key_col": "厂家编号"
    },
    # 任务3：设备台账对比（继续往下加即可）
    # {
    #     "name": "设备台账对比",
    #     "file1": "old_device.xls",
    #     "file2": "new_device.xls",
    #     "col_map": {
    #         "设备编号": "dev_id",
    #         "设备名称": "dev_name",
    #         "安装位置": "location",
    #     },
    #     "transform_funcs": {
    #         "设备名称": trim_str,
    #         "安装位置": trim_str,
    #     },
    #     "key_col": "设备编号"
    # },
]


# ========== 👆 配置结束，下面的不用改 ==========


def read_excel(file_path, col_names, transform_funcs):
    """读取 Excel 文件，返回 {主键: {列名: 转换后的值}}"""
    try:
        workbook = xlrd.open_workbook(file_path)
        sheet = workbook.sheet_by_index(0)
        headers = [unicode(cell.value).strip() for cell in sheet.row(0)]

        # 检查列是否存在
        col_index = {}
        for name in col_names:
            if name not in headers:
                print("错误：文件 %s 中找不到列 '%s'" % (file_path, name))
                return None
            col_index[name] = headers.index(name)

        data = {}
        for row_idx in range(1, sheet.nrows):
            row = sheet.row(row_idx)
            row_data = {}
            for name in col_names:
                cell_value = row[col_index[name]].value
                # 应用转换函数
                if name in transform_funcs:
                    row_data[name] = transform_funcs[name](cell_value)
                else:
                    row_data[name] = cell_value

            key = row_data[col_names[0]]  # 临时用第一列当key，后面会替换
            if key is None:
                continue
            data[key] = row_data

        return data
    except Exception as e:
        print("读取文件 %s 失败：%s" % (file_path, str(e)))
        return None


def compare_single_task(task):
    """对比单个任务"""
    print
    ("\n" + "=" * 70)
    print
    ("【任务：%s】" % task["name"])
    print
    ("对比文件：%s vs %s" % (task["file1"], task["file2"]))
    print
    ("=" * 70)

    # 读取两个文件
    df1 = read_excel(task["file1"], task["col_map"].keys(), task["transform_funcs"])
    df2 = read_excel(task["file2"], task["col_map"].values(), task["transform_funcs"])

    if df1 is None or df2 is None:
        print("❌ 任务执行失败，跳过")
        return False

    # 重命名 df2 的列
    reverse_map = dict((v, k) for k, v in task["col_map"].items())
    df2_renamed = {}
    for key, row in df2.items():
        new_row = {}
        for old_col, new_col in reverse_map.items():
            new_row[new_col] = row[old_col]
        df2_renamed[key] = new_row
    df2 = df2_renamed

    # 重新设置主键
    key_col = task["key_col"]
    df1_keyed = {}
    for row in df1.values():
        key = row[key_col]
        if key is not None:
            df1_keyed[key] = row

    df2_keyed = {}
    for row in df2.values():
        key = row[key_col]
        if key is not None:
            df2_keyed[key] = row

    # 找出差异
    keys1 = set(df1_keyed.keys())
    keys2 = set(df2_keyed.keys())
    only1 = sorted(keys1 - keys2)
    only2 = sorted(keys2 - keys1)

    if len(only1) > 0:
        print("\n❌ 仅在第一个文件的行（主键）：")
        for key in only1[:10]:  # 最多显示10个，太多省略
            print    ("  %s" % key)
        if len(only1) > 10:
            print    ("  ... 共 %d 条" % len(only1))

    if len(only2) > 0:
        print("\n❌ 仅在第二个文件的行（主键）：")
        for key in only2[:10]:
            print    ("  %s" % key)
        if len(only2) > 10:
            print    ("  ... 共 %d 条" % len(only2))

    # 对比单元格差异
    common_keys = sorted(keys1 & keys2)
    diff_count = 0
    print
    ("\n✅ 共同行的单元格差异（处理后）：")
    print
    ("-" * 70)
    print
    ("%-15s %-15s %-20s %-20s" % ("主键", "列名", "第一个文件", "第二个文件"))
    print
    ("-" * 70)

    for key in common_keys:
        row1 = df1_keyed[key]
        row2 = df2_keyed[key]
        for col in task["col_map"].keys():
            val1 = row1[col]
            val2 = row2[col]
            if val1 != val2:
                diff_count += 1
                if diff_count <= 20:  # 最多显示20条差异
                    print    ("%-15s %-15s %-20s %-20s" % (
                        unicode(key),
                        unicode(col),
                        unicode(val1) if val1 is not None else "空",
                        unicode(val2) if val2 is not None else "空"
                    ))

    if diff_count == 0:
        print("所有要对比的列，处理后完全一致！")
    elif diff_count > 20:
        print("  ... 共 %d 条差异" % diff_count)

    print("\n统计：")
    print("  第一个文件总行数：%d" % len(df1_keyed))
    print("  第二个文件总行数：%d" % len(df2_keyed))
    print("  共同行数：%d" % len(common_keys))
    print("  差异单元格数：%d" % diff_count)

    return diff_count == 0 and len(only1) == 0 and len(only2) == 0


def main():
    print("=" * 70)
    print("批量 Excel 对比工具（Python 2.5.2 兼容版）")
    print("共加载 %d 个对比任务" % len(TASKS))
    print("=" * 70)

    success_count = 0
    for i, task in enumerate(TASKS):
        print("\n[%d/%d] 正在执行..." % (i + 1, len(TASKS)))
        is_ok = compare_single_task(task)
        if is_ok:
            success_count += 1

    print("\n" + "=" * 70)
    print("全部任务执行完成！")
    print("成功：%d/%d" % (success_count, len(TASKS)))
    print("=" * 70)


if __name__ == "__main__":
    main()