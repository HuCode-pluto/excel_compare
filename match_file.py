import pandas as pd
import re
import time

# ===================== 多任务配置区 =====================
# 每个任务是一个字典，可添加任意多个任务，脚本会逐个自动执行
# 新增可选参数：
#   keep_duplicates: 文件2匹配键重复时的处理方式，默认 "first"，可选 "last"
#   output_unmatched: 是否将未匹配行单独输出到文件，默认 False
#   verbose: 是否打印匹配样例，默认 True
TASKS = [
    # ========== 任务1 示例 ==========
    {
        "file1": "file1.xlsx",
        "file2": "file2.xlsx",
        "output": "result1.xlsx",
        "has_header": False,
        "compare_cols": [
            (0, ["strip"], 0, ["strip"]),
            (1, ["chinese"], 1, ["strip", "chinese"]),
        ],
        "return_cols": [2],
        "keep_duplicates": "first",      # 重复时保留第一条
        "output_unmatched": False,       # 不单独输出未匹配行
        "verbose": True                  # 打印匹配样例
    },
    # ========== 任务2 示例 ==========
    # {
    #     "file1": "task2_file1.xlsx",
    #     "file2": "task2_file2.xlsx",
    #     "output": "result2.xlsx",
    #     "has_header": True,
    #     "compare_cols": [
    #         ("订单号", ["strip", "number"], "订单编号", ["strip", "number"]),
    #         ("客户名", ["strip"], "客户姓名", ["strip", "chinese"]),
    #     ],
    #     "return_cols": ["金额", "备注"],
    #     "keep_duplicates": "last",
    #     "output_unmatched": True,       # 输出“未匹配.xlsx”
    #     "verbose": True
    # },
]

# 支持的预处理规则列表（可按顺序组合使用）：
#   none        不处理
#   strip       去除首尾空格
#   lower       转小写
#   upper       转大写
#   chinese     只保留中文字符
#   number      只保留数字
#   left:n      从左边截取n个字符，例：left:3
#   right:n     从右边截取n个字符，例：right:4
#   mid:start:n 从第start位开始截取n个字符（start从0计数），例：mid:2:3
# ========================================================

# ---------- 预编译正则，提升性能 ----------
_CHINESE_RE = re.compile(r"[\u4e00-\u9fa5]")
_NUMBER_RE = re.compile(r"\d")


def process_text(text, rule_list):
    """按顺序执行多个预处理规则（逐行处理）"""
    if pd.isna(text):
        return ""
    s = str(text)

    for rule in rule_list:
        if rule == "none":
            continue
        elif rule == "strip":
            s = s.strip()
        elif rule == "lower":
            s = s.lower()
        elif rule == "upper":
            s = s.upper()
        elif rule == "chinese":
            s = "".join(_CHINESE_RE.findall(s))
        elif rule == "number":
            s = "".join(_NUMBER_RE.findall(s))
        elif rule.startswith("left:"):
            try:
                n = int(rule.split(":")[1])
            except (IndexError, ValueError):
                raise ValueError(f"规则格式错误: {rule}，正确示例 left:3")
            s = s[:n]
        elif rule.startswith("right:"):
            try:
                n = int(rule.split(":")[1])
            except (IndexError, ValueError):
                raise ValueError(f"规则格式错误: {rule}，正确示例 right:4")
            s = s[-n:] if n <= len(s) else s
        elif rule.startswith("mid:"):
            parts = rule.split(":")
            if len(parts) != 3:
                raise ValueError(f"规则格式错误: {rule}，正确示例 mid:2:3")
            try:
                start = int(parts[1])
                length = int(parts[2])
            except ValueError:
                raise ValueError(f"规则参数错误: {rule}，起始和长度必须为整数")
            s = s[start:start + length]
        # 未知规则跳过不处理
    return s


def _can_vectorize(rule_list):
    """判断规则列表能否完全用 Pandas 向量化操作完成"""
    allowed = {"none", "strip", "lower", "upper"}
    for rule in rule_list:
        if rule not in allowed:
            return False
    return True


def _vectorized_process(series, rule_list):
    """使用向量化方法处理 Series，仅处理简单规则组合"""
    s = series.astype(str)
    for rule in rule_list:
        if rule == "none":
            continue
        elif rule == "strip":
            s = s.str.strip()
        elif rule == "lower":
            s = s.str.lower()
        elif rule == "upper":
            s = s.str.upper()
        # 若包含 chinese/number/截取等规则，不会进入此函数（由 _can_vectorize 保证）
    return s


def run_single_task(task):
    """执行单个对比任务（优化版）"""
    start_time = time.time()

    # 读取任务参数，处理默认值
    file1_path = task["file1"]
    file2_path = task["file2"]
    output_path = task["output"]
    has_header = task["has_header"]
    compare_cols = task["compare_cols"]
    return_cols = task["return_cols"]
    keep_dup = task.get("keep_duplicates", "first")
    output_unmatched = task.get("output_unmatched", False)
    verbose = task.get("verbose", True)

    # 1. 读取文件，避免空单元格变成字符串 "nan"
    read_param = {"dtype": str, "keep_default_na": False, "na_values": []}
    if not has_header:
        read_param["header"] = None

    df1 = pd.read_excel(file1_path, **read_param)
    df2 = pd.read_excel(file2_path, **read_param)

    original_col_count = len(df1.columns)
    print(f"\n{'='*50}")
    print(f"处理任务：{file1_path}  ↔  {file2_path}")
    print(f"文件1原有 {original_col_count} 列，返回列将从第 {original_col_count+1} 列开始追加")
    print(f"配置对比列数：{len(compare_cols)} 列，待返回列数：{len(return_cols)} 列")

    # 2. 列存在性校验
    for idx, (col1, rules1, col2, rules2) in enumerate(compare_cols):
        if col1 not in df1.columns:
            raise KeyError(f"文件1中不存在列: {col1}（可用列：{list(df1.columns)}）")
        if col2 not in df2.columns:
            raise KeyError(f"文件2中不存在列: {col2}（可用列：{list(df2.columns)}）")
    for col in return_cols:
        if col not in df2.columns:
            raise KeyError(f"文件2中不存在待返回列: {col}（可用列：{list(df2.columns)}）")

    # 3. 生成临时匹配键（根据规则自动选择向量化或逐行处理）
    key_cols = []
    for idx, (col1, rules1, col2, rules2) in enumerate(compare_cols):
        key_name = f"_key_{idx}"
        key_cols.append(key_name)

        # 文件1 键列
        if _can_vectorize(rules1):
            df1[key_name] = _vectorized_process(df1[col1], rules1)
        else:
            df1[key_name] = df1[col1].apply(lambda x: process_text(x, rules1))

        # 文件2 键列
        if _can_vectorize(rules2):
            df2[key_name] = _vectorized_process(df2[col2], rules2)
        else:
            df2[key_name] = df2[col2].apply(lambda x: process_text(x, rules2))

    # 4. 文件2按所有对比键去重，先统计重复情况
    dup_mask = df2.duplicated(subset=key_cols, keep=False)
    dup_count = dup_mask.sum()
    if dup_count > 0:
        print(f"⚠️ 文件2中匹配键存在 {dup_count} 行重复，将按照 keep='{keep_dup}' 保留")
    df2_unique = df2.drop_duplicates(subset=key_cols, keep=keep_dup)

    # 5. 提取文件2的匹配键 + 待返回列
    df2_merge = df2_unique[key_cols + return_cols].copy()

    # 6. 多列联合左连接
    result = df1.merge(df2_merge, on=key_cols, how="left")

    # 7. 删除临时键列
    result.drop(columns=key_cols, inplace=True)

    # 8. 有表头模式下：列名冲突自动加后缀
    if has_header:
        rename_map = {}
        for col in return_cols:
            if col in df1.columns:
                rename_map[col] = f"{col}_来自文件2"
        if rename_map:
            result.rename(columns=rename_map, inplace=True)
            print(f"检测到列名重复，已自动重命名：{list(rename_map.values())}")

    # 9. 匹配统计
    first_return_col = result.columns[original_col_count]
    matched = result[first_return_col].notna().sum()
    total = len(result)
    match_rate = matched / total * 100 if total > 0 else 0
    elapsed = time.time() - start_time
    print(f"总行数：{total}")
    print(f"全部列匹配成功：{matched} 行 ({match_rate:.1f}%)")
    print(f"未匹配：{total - matched} 行")
    print(f"任务耗时：{elapsed:.2f} 秒")

    # 10. 打印匹配样例（前3条匹配成功的行）
    if verbose and matched > 0:
        print("\n前3条匹配成功样例：")
        matched_rows = result[result[first_return_col].notna()].head(3)
        # 仅展示原始列 + 首个返回列
        show_cols = list(result.columns[:original_col_count]) + [first_return_col]
        try:
            print(matched_rows[show_cols].to_string(index=False))
        except Exception:
            print("（样例展示失败）")

    # 11. 保存结果
    result.to_excel(output_path, index=False, header=has_header)
    print(f"结果已保存到：{output_path}")

    # 12. 可选：输出未匹配行
    if output_unmatched:
        unmatched = result[result[first_return_col].isna()]
        if not unmatched.empty:
            unmatched_path = output_path.replace(".xlsx", "_未匹配.xlsx")
            unmatched.to_excel(unmatched_path, index=False, header=has_header)
            print(f"未匹配行已单独保存到：{unmatched_path}")
        else:
            print("所有行均已匹配，无需生成未匹配文件。")


def main():
    total_tasks = len(TASKS)
    print(f"共配置 {total_tasks} 个任务，开始批量处理...\n")

    success = 0
    for i, task in enumerate(TASKS, 1):
        try:
            run_single_task(task)
            success += 1
        except Exception as e:
            print(f"\n❌ 任务 {i} 执行失败：{e}")

    print(f"\n{'='*50}")
    print(f"全部任务处理完成：成功 {success}/{total_tasks}")


if __name__ == "__main__":
    main()