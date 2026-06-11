# -*- coding: utf-8 -*-
import pandas as pd
import sys
import os
import traceback


# ========== 通用转换&清洗函数 ==========
def full_clean(s):
    """
    深度清洗：去除所有空格、制表符、全角空格、空白字符，统一为纯字符串
    解决肉眼一致、隐形字符导致的误判
    """
    if pd.isna(s):
        return ""
    s = str(s)
    # 常规半角空格、制表符、换行
    s = s.replace(" ", "").replace("\t", "").replace("\n", "").replace("\r", "")
    # 全角空格
    s = s.replace("　", "")
    return s


def trim_str_length(s, max_length=19):
    if pd.isna(s):
        return ""
    result = str(s).strip()
    if max_length is not None:
        result = result[:max_length]
    return result


def convert_station_type(s):
    """开关站类型 中文/数字 统一映射"""
    if pd.isna(s):
        return ""
    s = full_clean(s)
    if s in ("开关站",):
        return "0"
    elif s in ("箱式变电站",):
        return "1"
    if s == "0.0" or s == "0":
        return "0"
    elif s == "1.0" or s == "1":
        return "1"
    return s


# ========== 批量任务配置区 ==========
# ignore_only_row: True=忽略独有行，只对比共同数据；False=严格校验行数一致
TASKS = [
    # 任务1：配网馈线表实时库和达梦对比
    {
        "name": "配网馈线表实时库和达梦对比",
        "file1": "C:\\Users\\13303\\Desktop\\kx610.xls",
        "file2": "C:\\Users\\13303\\Desktop\\kx610_dm.xls",
        "col_map": {
            "馈线ID号": "ID",
            "馈线名称": "NAME",
            # "图形名": "GRAPH_NAME",
        },
        "transform_funcs": {
            "馈线ID号": lambda x: trim_str_length(x, 19),
            "馈线名称": full_clean,
            # "图形名": full_clean,
        },
        "key_col": "馈线ID号",
        "ignore_only_row": True   # 忽略独有行
    },
    # 任务2：开关站实时库商用库对比
    {
        "name": "开关站实时库商用库对比",
        "file1": "C:\\Users\\13303\\Desktop\\kgz610.xls",
        "file2": "C:\\Users\\13303\\Desktop\\kgz610_dm.xls",
        "col_map": {
            "开关站ID号": "ID",
            "开关站名称": "NAME",
            "所属馈线": "feeder_name",
            # "开关站类型": "COMBINED_STATE",
        },
        "transform_funcs": {
            "开关站ID号": lambda x: trim_str_length(x, 19),
            "开关站名称": full_clean,
            "所属馈线": full_clean,
            # "开关站类型": convert_station_type,
        },
        "key_col": "开关站ID号",
        "ignore_only_row": True   # 忽略独有行
    },
]


# ========== 文件读取 ==========
def load_excel(file_path):
    if not os.path.exists(file_path):
        print(f"错误：文件不存在 {file_path}")
        return None

    suffix = os.path.splitext(file_path)[1].lower()
    try:
        if suffix == ".xls":
            df = pd.read_excel(file_path, engine="xlrd")
        elif suffix == ".xlsx":
            df = pd.read_excel(file_path, engine="openpyxl")
        else:
            df = pd.read_excel(file_path)

        df.columns = [str(c).strip() for c in df.columns]
        print(f"【调试】{os.path.basename(file_path)} 实际表头：{list(df.columns)}")
        return df
    except Exception as e1:
        # 兼容伪Excel(TSV)，多编码+制表符分隔
        try:
            enc_list = ["utf-8", "gb18030", "gbk"]
            for enc in enc_list:
                try:
                    df = pd.read_csv(file_path, encoding=enc, sep="\t")
                    df.columns = [str(c).strip() for c in df.columns]
                    print(f"【调试】使用编码 {enc} 读取成功")
                    print(f"【调试】{os.path.basename(file_path)} 实际表头：{list(df.columns)}")
                    return df
                except:
                    continue
            return None
        except Exception as e2:
            print(f"读取失败，详细错误：")
            traceback.print_exc()
            return None


# ========== 单任务对比逻辑 ==========
def compare_single_task(task):
    print("\n" + "=" * 70)
    print(f"【任务：{task['name']}】")
    print(f"对比文件：{task['file1']} vs {task['file2']}")
    print("=" * 70)

    try:
        df1 = load_excel(task["file1"])
        df2 = load_excel(task["file2"])
        if df1 is None or df2 is None:
            return False

        # 校验列是否存在
        map_keys = list(task["col_map"].keys())
        map_vals = list(task["col_map"].values())
        missing1 = [c for c in map_keys if c not in df1.columns]
        missing2 = [c for c in map_vals if c not in df2.columns]
        if missing1 or missing2:
            if missing1:
                print(f"❌ 文件1缺失列：{missing1}")
            if missing2:
                print(f"❌ 文件2缺失列：{missing2}")
            return False

        # 筛选列 + 统一列名（全部统一为文件1的列名，这样后面的transform才能找到列）
        df1 = df1[map_keys]  # 文件1的列名不变！
        rev_map = {v: k for k, v in task["col_map"].items()}
        df2 = df2[map_vals].rename(columns=rev_map)  # 只改文件2的列名，改成文件1的

        # 应用清洗/转换函数
        for col, func in task["transform_funcs"].items():
            df1[col] = df1[col].apply(func)
            df2[col] = df2[col].apply(func)

        key_col = task["key_col"]
        df1 = df1.set_index(key_col).sort_index()
        df2 = df2.set_index(key_col).sort_index()

        # 独有行
        only1 = sorted(set(df1.index) - set(df2.index))
        only2 = sorted(set(df2.index) - set(df1.index))

        if len(only1) > 0:
            print(f"\n❌ 仅在第一个文件的主键({len(only1)}条)：")
            for k in only1[:10]:
                print(f"  {k}")
            if len(only1) > 10:
                print(f"  ... 省略剩余 {len(only1)-10} 条")

        if len(only2) > 0:
            print(f"\n❌ 仅在第二个文件的主键({len(only2)}条)：")
            for k in only2[:10]:
                print(f"  {k}")
            if len(only2) > 10:
                print(f"  ... 省略剩余 {len(only2)-10} 条")

        # 共同行对比
        common_idx = df1.index.intersection(df2.index)
        df_common1 = df1.loc[common_idx]
        df_common2 = df2.loc[common_idx]

        # diff_df = df_common1.compare(df_common2, keep_shape=True, keep_equal=True)
        diff_df = df_common1.compare(df_common2, keep_shape=True)
        diff_count = diff_df.notna().any(axis=1).sum()

        if diff_count == 0:
            print(f"\n✅ 共同行【所有字段完全一致】")
        else:
            print(f"\n⚠️ 共同行存在 {diff_count} 处真实数据差异：")
            print("-" * 70)
            print(diff_df.dropna(how='all'))

        # 统计信息
        print(f"\n【统计】")
        print(f"  文件1总行数：{len(df1)}")
        print(f"  文件2总行数：{len(df2)}")
        print(f"  共同行数：{len(common_idx)}")
        print(f"  独有行数：文件1:{len(only1)} | 文件2:{len(only2)}")
        print(f"  数据差异行数：{diff_count}")

        # 判定任务成功
        ignore_only = task.get("ignore_only_row", False)
        if ignore_only:
            # 只看共同行是否无差异
            return diff_count == 0
        else:
            # 严格模式：无独有行 + 无差异
            return (len(only1) == 0 and len(only2) == 0 and diff_count == 0)

    except Exception as e:
        print(f"❌ 任务执行异常：{str(e)}")
        traceback.print_exc()
        return False


def main():
    print("=" * 70)
    print("批量Excel对比工具（深度清洗+区分独有行）")
    print(f"共加载 {len(TASKS)} 个对比任务")
    print("=" * 70)

    success_count = 0
    for idx, task in enumerate(TASKS):
        print(f"\n[{idx+1}/{len(TASKS)}] 开始执行：{task['name']}")
        ok = compare_single_task(task)
        if ok:
            success_count += 1

    print("\n" + "=" * 70)
    print(f"执行汇总：成功 {success_count} / 总计 {len(TASKS)}")
    print("=" * 70)


if __name__ == "__main__":
    main()