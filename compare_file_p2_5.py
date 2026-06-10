#!/usr/bin/env python2.7
# -*- coding: utf-8 -*-
"""
批量 Excel 对比工具（Python 2.5 兼容版）
========================================
适配说明：
  1. f-string 全部替换为 %% 格式化
  2. print() 函数 → print 语句
  3. except X as e → except X, e
  4. pd.isna() → pd.isnull()
  5. 字典/集合推导式 → generator + 构造器
  6. 全角字符使用 unicode 字面量 u"　"
  7. df.compare() → 手动 diff_mask（pandas 1.1.0 才引入）
  8. sys.exit() 返回退出码，方便 CI/脚本判断
"""
import pandas as pd
import sys
import os
import traceback


# ========== 通用工具函数 ==========
def safe_unicode(s):
    """
    安全转为 unicode，兼容 str(bytes) / unicode / 数字 / None 等各种输入。
    Python 2 下 unicode() 默认用 ASCII 解码 bytes，遇到中文会炸，所以先判断类型。
    """
    if isinstance(s, unicode):
        return s
    if isinstance(s, str):
        # bytes → unicode，优先 utf-8，回退 gb18030
        for enc in (u"utf-8", u"gb18030"):
            try:
                return s.decode(enc)
            except (UnicodeDecodeError, UnicodeError):
                continue
        return s.decode(u"utf-8", u"replace")
    # 数字、float 等
    return unicode(s)


# ========== 通用转换 & 清洗函数 ==========
def full_clean(s):
    """
    深度清洗：去除所有空格、制表符、全角空格、空白字符，统一为纯字符串
    解决肉眼一致、隐形字符导致的误判
    """
    if pd.isnull(s):
        return u""
    s = safe_unicode(s)
    # 常规半角空格、制表符、换行
    s = s.replace(u" ", u"").replace(u"\t", u"").replace(u"\n", u"").replace(u"\r", u"")
    # 全角空格（U+3000）
    s = s.replace(u"　", u"")
    return s


def trim_str_length(s, max_length=19):
    """截断字符串到指定长度（用于适配数据库字段长度限制）"""
    if pd.isnull(s):
        return u""
    result = safe_unicode(s).strip()
    if max_length is not None:
        result = result[:max_length]
    return result


def convert_station_type(s):
    """开关站类型 中文/数字 统一映射
    0=开关站, 1=箱式变电站
    """
    if pd.isnull(s):
        return u""
    s = full_clean(s)
    if s in (u"开关站",):
        return u"0"
    elif s in (u"箱式变电站",):
        return u"1"
    # 数字标准化：去小数点
    if s in (u"0", u"0.0"):
        return u"0"
    elif s in (u"1", u"1.0"):
        return u"1"
    return s


# ========== 批量任务配置区 ==========
# ignore_only_row: True=忽略独有行，只对比共同数据；False=严格校验行数一致
TASKS = [
    # 任务1：配网馈线表实时库和达梦对比
    {
        "name": u"配网馈线表实时库和达梦对比",
        "file1": u"C:\\Users\\13303\\Desktop\\kx610.xls",
        "file2": u"C:\\Users\\13303\\Desktop\\kx610_dm.xls",
        "col_map": {
            u"馈线ID号": u"ID",
            u"馈线名称": u"NAME",
            # u"图形名": u"GRAPH_NAME",
        },
        "transform_funcs": {
            u"馈线ID号": lambda x: trim_str_length(x, 19),
            u"馈线名称": full_clean,
            # u"图形名": full_clean,
        },
        "key_col": u"馈线ID号",
        "ignore_only_row": True   # 忽略独有行
    },
    # 任务2：开关站实时库商用库对比
    {
        "name": u"开关站实时库商用库对比",
        "file1": u"C:\\Users\\13303\\Desktop\\kgz610.xls",
        "file2": u"C:\\Users\\13303\\Desktop\\kgz610_dm.xls",
        "col_map": {
            u"开关站ID号": u"ID",
            u"开关站名称": u"NAME",
            u"所属馈线": u"feeder_name",
            # u"开关站类型": u"COMBINED_STATE",
        },
        "transform_funcs": {
            u"开关站ID号": lambda x: trim_str_length(x, 19),
            u"开关站名称": full_clean,
            u"所属馈线": full_clean,
            # u"开关站类型": convert_station_type,
        },
        "key_col": u"开关站ID号",
        "ignore_only_row": True   # 忽略独有行
    },
]


# ========== 文件读取 ==========
def load_excel(file_path):
    """
    读取 Excel/TSV 文件，自动检测编码
    返回 DataFrame 或 None
    """
    if not os.path.exists(file_path):
        print u"错误：文件不存在 %s" % file_path
        return None

    suffix = os.path.splitext(file_path)[1].lower()
    basename = os.path.basename(file_path)

    # ---- 优先按 Excel 读取 ----
    try:
        if suffix == u".xls":
            df = pd.read_excel(file_path, engine=u"xlrd")
        elif suffix == u".xlsx":
            df = pd.read_excel(file_path, engine=u"openpyxl")
        else:
            df = pd.read_excel(file_path)

        df.columns = [unicode(c).strip() for c in df.columns]
        print u"【调试】%s 实际表头：%s" % (basename, list(df.columns))
        return df
    except Exception, e1:
        # 兼容伪Excel（TSV格式），多编码 + 制表符分隔
        enc_list = [u"utf-8", u"gb18030", u"gbk"]
        for enc in enc_list:
            try:
                df = pd.read_csv(file_path, encoding=enc, sep=u"\t")
                df.columns = [unicode(c).strip() for c in df.columns]
                print u"【调试】使用编码 %s 读取成功" % enc
                print u"【调试】%s 实际表头：%s" % (basename, list(df.columns))
                return df
            except Exception:
                continue
        # 所有编码都失败
        print u"读取失败，所有编码均无法解析：%s" % file_path
        traceback.print_exc()
        return None


# ========== 单任务对比逻辑 ==========
def compare_single_task(task):
    """执行单个对比任务，返回 True/False"""
    print
    print u"=" * 70
    print u"【任务：%s】" % task[u"name"]
    print u"对比文件：%s  vs  %s" % (task[u"file1"], task[u"file2"])
    print u"=" * 70

    try:
        df1 = load_excel(task[u"file1"])
        df2 = load_excel(task[u"file2"])
        if df1 is None or df2 is None:
            return False

        # 校验列是否存在
        map_keys = list(task[u"col_map"].keys())
        map_vals = list(task[u"col_map"].values())
        missing1 = [c for c in map_keys if c not in df1.columns]
        missing2 = [c for c in map_vals if c not in df2.columns]
        if missing1 or missing2:
            if missing1:
                print u"❌ 文件1缺失列：%s" % missing1
            if missing2:
                print u"❌ 文件2缺失列：%s" % missing2
            return False

        # 筛选列 + 统一列名（统一为文件1的列名，方便 transform 查找）
        df1 = df1[map_keys]
        # Python 2.5 无字典推导式，用 generator + dict()
        rev_map = dict((v, k) for k, v in task[u"col_map"].items())
        df2 = df2[map_vals].rename(columns=rev_map)

        # 应用清洗/转换函数
        for col, func in task[u"transform_funcs"].items():
            df1[col] = df1[col].apply(func)
            df2[col] = df2[col].apply(func)

        # 设主键索引并排序
        key_col = task[u"key_col"]
        df1 = df1.set_index(key_col).sort_index()
        df2 = df2.set_index(key_col).sort_index()

        # ---- 独有行检测 ----
        only1 = sorted(set(df1.index) - set(df2.index))
        only2 = sorted(set(df2.index) - set(df1.index))

        if len(only1) > 0:
            print
            print u"❌ 仅在第一个文件的主键(%d条)：" % len(only1)
            for k in only1[:10]:
                print u"  %s" % k
            if len(only1) > 10:
                print u"  ... 省略剩余 %d 条" % (len(only1) - 10)

        if len(only2) > 0:
            print
            print u"❌ 仅在第二个文件的主键(%d条)：" % len(only2)
            for k in only2[:10]:
                print u"  %s" % k
            if len(only2) > 10:
                print u"  ... 省略剩余 %d 条" % (len(only2) - 10)

        # ---- 共同行逐字段对比（手动实现，替代 df.compare()） ----
        # df.compare() 是 pandas 1.1.0 (2020) 才引入的，老 pandas 不存在。
        # 这里的实现与 compare(keep_shape=True, keep_equal=True) 语义一致，
        # 并在输出格式上更友好：逐行逐列展示"文件1的值 ←→ 文件2的值"。
        common_idx = df1.index.intersection(df2.index)
        df_common1 = df1.loc[common_idx]
        df_common2 = df2.loc[common_idx]

        # 差异掩码：值不同 且 非"双方均为空"
        # NaN != NaN 在 pandas 中为 True，所以要排除 isnull1 & isnull2
        diff_mask = (df_common1 != df_common2) & \
                    ~(df_common1.isnull() & df_common2.isnull())
        diff_count = diff_mask.any(axis=1).sum()

        if diff_count == 0:
            print
            print u"✅ 共同行【所有字段完全一致】"
        else:
            print
            print u"⚠️ 共同行存在 %d 行数据差异：" % diff_count
            print u"-" * 70
            # 只遍历有差异的行（通常量很少，性能 OK）
            diff_row_mask = diff_mask.any(axis=1)
            diff_indices = diff_row_mask[diff_row_mask].index
            for idx in diff_indices:
                print u"  [主键] %s" % safe_unicode(idx)
                row_mask = diff_mask.loc[idx]
                diff_cols = [c for c in row_mask.index if row_mask[c]]
                for col in diff_cols:
                    v1 = df_common1.loc[idx, col]
                    v2 = df_common2.loc[idx, col]
                    # NaN 统一显示为 "(空)"
                    s1 = u"(空)" if pd.isnull(v1) else safe_unicode(v1)
                    s2 = u"(空)" if pd.isnull(v2) else safe_unicode(v2)
                    print u"    %s: [文件1] %s  ←→  [文件2] %s" % (col, s1, s2)
                print

        # ---- 统计信息 ----
        print
        print u"【统计】"
        print u"  文件1总行数：%d" % len(df1)
        print u"  文件2总行数：%d" % len(df2)
        print u"  共同行数：%d" % len(common_idx)
        print u"  独有行数：文件1:%d | 文件2:%d" % (len(only1), len(only2))
        print u"  数据差异行数：%d" % diff_count

        # ---- 判定任务是否通过 ----
        ignore_only = task.get(u"ignore_only_row", False)
        if ignore_only:
            return diff_count == 0
        else:
            return (len(only1) == 0 and len(only2) == 0 and diff_count == 0)

    except Exception, e:
        print u"❌ 任务执行异常：%s" % safe_unicode(e)
        traceback.print_exc()
        return False


def main():
    print u"=" * 70
    print u"批量Excel对比工具（深度清洗+区分独有行）Python 2.5 兼容版"
    print u"共加载 %d 个对比任务" % len(TASKS)
    print u"=" * 70

    success_count = 0
    for idx, task in enumerate(TASKS):
        print
        print u"[%d/%d] 开始执行：%s" % (idx + 1, len(TASKS), task[u"name"])
        ok = compare_single_task(task)
        if ok:
            success_count += 1

    print
    print u"=" * 70
    print u"执行汇总：成功 %d / 总计 %d" % (success_count, len(TASKS))
    print u"=" * 70

    # 返回退出码：全部成功=0，否则=1（方便 CI / 脚本判断）
    sys.exit(0 if success_count == len(TASKS) else 1)


if __name__ == u"__main__":
    main()
