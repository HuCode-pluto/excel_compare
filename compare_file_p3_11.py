# -*- coding: utf-8 -*-
import pandas as pd
import sys
import datetime


# ========== 👇 通用转换函数（所有任务可复用） ==========
def trim_str(s):
    """去掉前后空格，处理空值"""
    if pd.isna(s):
        return None
    return str(s).strip()


def parse_date(d):
    """统一日期格式，支持 Excel 日期数字和字符串"""
    if pd.isna(d):
        return None
    return pd.to_datetime(d, errors='coerce').date()


def parse_percent(p):
    """统一百分比："98.5%" 和 0.985 都转成 98.5"""
    if pd.isna(p):
        return None
    s = str(p).strip()
    if s.endswith('%'):
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
    if pd.isna(f):
        return None
    s = str(f).strip()
    if "南瑞" in s:
        return "南瑞"
    if "许继" in s:
        return "许继"
    if "四方" in s:
        return "四方"
    return s


# ========== 👇 批量任务配置区（只需要改这里） ==========
# 每个 {} 代表一对对比任务，复制粘贴即可添加新任务
TASKS = [
    # 任务1：终端信息对比
    {
        "name": "终端基础信息对比",  # 任务名称，输出时显示
        "file1": "old_terminal.xlsx",  # 第一个文件路径（支持xls/xlsx）
        "file2": "new_terminal.xlsx",  # 第二个文件路径
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
        "file1": "old_factory.xlsx",
        "file2": "new_factory.xlsx",
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
    #     "file1": "old_device.xlsx",
    #     "file2": "new_device.xlsx",
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
        # 自动识别 xls/xlsx，pandas 3.11 完美支持
        df = pd.read_excel(file_path)
        df = df[col_names].copy()

        # 应用转换函数
        for col, func in transform_funcs.items():
            df[col] = df[col].apply(func)

        # 按主键索引
        df = df.set_index(KEY_COL, drop=False)
        return df
    except Exception as e:
        print(f"错误：读取文件 {file_path} 失败：{str(e)}")
        return None


def compare_single_task(task):
    """对比单个任务"""
    print("\n" + "=" * 70)
    print(f"【任务：{task['name']}】")
    print(f"对比文件：{task['file1']} vs {task['file2']}")
    print("=" * 70)

    # 读取两个文件
    df1 = pd.read_excel(task["file1"])
    df2 = pd.read_excel(task["file2"])

    # 重命名列，统一
    df1 = df1[list(task["col_map"].keys())].rename(columns={v: k for k, v in task["col_map"].items()})
    df2 = df2[list(task["col_map"].values())].rename(columns={v: k for k, v in task["col_map"].items()})

    # 应用转换函数
    for col, func in task["transform_funcs"].items():
        df1[col] = df1[col].apply(func)
        df2[col] = df2[col].apply(func)

    # 按主键排序
    key_col = task["key_col"]
    df1 = df1.set_index(key_col).sort_index()
    df2 = df2.set_index(key_col).sort_index()

    # 找出只在一个文件的行
    only1 = sorted(set(df1.index) - set(df2.index))
    only2 = sorted(set(df2.index) - set(df1.index))

    if len(only1) > 0:
        print(f"\n❌ 仅在第一个文件的行（主键）：")
        for key in only1[:10]:
            print(f"  {key}")
        if len(only1) > 10:
            print(f"  ... 共 {len(only1)} 条")

    if len(only2) > 0:
        print(f"\n❌ 仅在第二个文件的行（主键）：")
        for key in only2[:10]:
            print(f"  {key}")
        if len(only2) > 10:
            print(f"  ... 共 {len(only2)} 条")

    # 对比单元格差异
    common = df1.index.intersection(df2.index)
    diff = df1.loc[common].compare(df2.loc[common], keep_shape=True, keep_equal=True)
    diff_count = diff.notna().any(axis=1).sum()

    if diff_count == 0:
        print(f"\n✅ 所有要对比的列，处理后完全一致！")
    else:
        print(f"\n✅ 共同行的单元格差异（处理后）：")
        print("-" * 70)
        print(diff.dropna(how='all'))

    print(f"\n统计：")
    print(f"  第一个文件总行数：{len(df1)}")
    print(f"  第二个文件总行数：{len(df2)}")
    print(f"  共同行数：{len(common)}")
    print(f"  差异单元格数：{diff_count}")

    return diff_count == 0 and len(only1) == 0 and len(only2) == 0


def main():
    print("=" * 70)
    print("批量 Excel 对比工具（Python 3.11 专属版，支持xlsx）")
    print(f"共加载 {len(TASKS)} 个对比任务")
    print("=" * 70)

    success_count = 0
    for i, task in enumerate(TASKS):
        print(f"\n[{i + 1}/{len(TASKS)}] 正在执行...")
        try:
            is_ok = compare_single_task(task)
            if is_ok:
                success_count += 1
        except Exception as e:
            print(f"任务 {task['name']} 执行失败：{str(e)}")

    print("\n" + "=" * 70)
    print("全部任务执行完成！")
    print(f"成功：{success_count}/{len(TASKS)}")
    print("=" * 70)


if __name__ == "__main__":
    main()