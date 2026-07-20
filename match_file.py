import pandas as pd
import re
import time

# ===================== 多任务配置区 =====================
# 每个任务是一个字典，可添加任意多个任务，脚本会逐个自动执行
# 新增可选参数：
#   keep_duplicates: 文件2匹配键重复时的处理方式，默认 "first"，可选 "last"
#   output_unmatched: 是否将未匹配行单独输出到文件，默认 False
#   verbose: 是否打印匹配样例，默认 True
#   secondary_compare_cols: 二次匹配规则（可选），格式同compare_cols，仅对第一次未匹配行生效

# 支持的预处理规则列表（可按顺序组合使用）：
#   none              不处理
#   strip             去除首尾空格
#   lower             转小写
#   upper             转大写
#   chinese           只保留中文字符
#   number            只保留数字
#   left:n            从左边截取n个字符，例：left:3
#   right:n           从右边截取n个字符，例：right:4
#   mid:start:n       从第start位开始截取n个字符（start从0计数），例：mid:2:3
#   remove_left:列名   如果当前列以指定列内容开头，则删除该前缀（精确匹配）
#   remove_right:列名  如果当前列以指定列内容结尾，则删除该后缀
#   trim_left:列名     无条件从左侧截去指定列值长度的字符
#   trim_right:列名    无条件从右侧截去指定列值长度的字符
#   remove:字符串      删除当前列中所有出现的指定字符串（区分大小写）
#   symbols           删除所有非中文、非字母、非数字的符号（包括空格、标点）
#   cut_left:n        从左侧删除n个字符（固定长度），例：cut_left:2
#   cut_right:n       从右侧删除n个字符（固定长度），例：cut_right:3
#   remove_suffix:后缀1|后缀2|...  若字符串以任意指定后缀结尾则删除该后缀（多后缀用|分隔，自动优先匹配长后缀）
#   keep_from:字符串   保留指定字符串及之后的所有内容，删除其前面的字符（匹配第一个出现的位置，无匹配则不处理）
#   remove_left4_if_digits   若前4位全部是数字则删除前4位，否则不处理
# ========================================================

TASKS = [
    # ========== 刀闸匹配（含二次匹配） ==========
    {
        # 文件1路径（主表）
        "file1": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260629\\fx_dz_629.xlsx",
        # 文件2路径（从表）
        "file2": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260629\\刀闸.xlsx",
        # 输出文件路径
        "output": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260629\\刀闸匹配结果.xlsx",
        # Excel是否有表头：True=有表头，列名使用字符串；False=无表头，列名使用数字索引
        "has_header": True,
        # 第一次对比列配置：每个元组为(文件1列, 文件1预处理规则列表, 文件2列, 文件2预处理规则列表)
        "compare_cols": [
            ('刀闸名称', ["symbols", "remove_left:站房名称", "left:8", "remove_suffix:2|手|手车2|手车"], '名称',["symbols", "left:8", "remove_suffix:2"]),
            ('站房名称', ["symbols", "chinese"], '所属站房', ["symbols", "chinese"]),
        ],
        # 二次匹配对比列配置（可选）：仅对第一次未匹配行生效
        "secondary_compare_cols": [
            # 示例：放宽匹配规则，仅匹配站房名称+刀闸名称前6位
            ('刀闸名称', ["symbols", "remove_left:站房名称", "left:2"], '名称', ["symbols", "left:2"]),
            ('站房名称', ["symbols", "chinese"], '所属站房', ["symbols", "chinese"]),
        ],
        # 待返回列：匹配成功后从文件2提取这些列追加到文件1末尾
        "return_cols": ["ID", "名称", "所属站房"],
        # 文件2匹配键重复时保留第一条
        "keep_duplicates": "first",
        # 单独输出最终未匹配行（两次匹配都失败的）
        "output_unmatched": True,
        # 打印匹配样例
        "verbose": True
    },
# ========== 断路器/负荷开关匹配（含二次匹配） ==========
    {
        # 文件1路径（主表）
        "file1": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260629\\fx_kg_二遥_629.xlsx",
        # 文件2路径（从表）
        "file2": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260629\\断路器.xlsx",
        # 输出文件路径
        "output": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260629\\开关匹配结果.xlsx",
        # Excel是否有表头：True=有表头，列名使用字符串；False=无表头，列名使用数字索引
        "has_header": True,
        # 第一次对比列配置：每个元组为(文件1列, 文件1预处理规则列表, 文件2列, 文件2预处理规则列表)
        "compare_cols": [
            ('开关名称', ["symbols", "remove_left:站房名称", "left:8", "remove_suffix:2|1"], '名称',["symbols","remove_left4_if_digits","left:8", "remove_suffix:2|1"]),
            ('站房名称', ["symbols", "chinese"], '所属站房', ["symbols", "chinese"]),
        ],
        # 二次匹配对比列配置（可选）：仅对第一次未匹配行生效
        "secondary_compare_cols": [
            # 示例：放宽匹配规则，仅匹配站房名称+刀闸名称前2位
            ('开关名称', ["symbols", "remove_left:站房名称", "left:2"], '名称', ["symbols", "left:2"]),
            ('站房名称', ["symbols", "chinese"], '所属站房', ["symbols", "chinese"]),
        ],
        # 待返回列：匹配成功后从文件2提取这些列追加到文件1末尾
        "return_cols": ["ID", "名称", "所属站房"],
        # 文件2匹配键重复时保留第一条
        "keep_duplicates": "first",
        # 单独输出最终未匹配行（两次匹配都失败的）
        "output_unmatched": True,
        # 打印匹配样例
        "verbose": True
    },
# ========== 接地刀闸开关匹配（含二次匹配） ==========
    {
        # 文件1路径（主表）
        "file1": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260629\\fx_jddz_629.xlsx",
        # 文件2路径（从表）
        "file2": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260629\\接地刀闸.xlsx",
        # 输出文件路径
        "output": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260629\\接地刀闸匹配结果.xlsx",
        # Excel是否有表头：True=有表头，列名使用字符串；False=无表头，列名使用数字索引
        "has_header": True,
        # 第一次对比列配置：每个元组为(文件1列, 文件1预处理规则列表, 文件2列, 文件2预处理规则列表)
        "compare_cols": [
            ('接地刀名称', ["symbols", "remove_left:站房名称", "left:8", "remove_suffix:2|1"], '名称', ["symbols", "left:8", "remove_suffix:2|1"]),
            ('站房名称', ["symbols", "chinese"], '所属站房', ["symbols", "chinese"]),
        ],
        # 二次匹配对比列配置（可选）：仅对第一次未匹配行生效
        "secondary_compare_cols": [
            # 示例：放宽匹配规则，仅匹配站房名称+刀闸名称前6位
            ('接地刀名称', ["symbols", "remove_left:站房名称", "left:2"], '名称', ["symbols", "left:2"]),
            ('站房名称', ["symbols", "chinese"], '所属站房', ["symbols", "chinese"]),
        ],
        # 待返回列：匹配成功后从文件2提取这些列追加到文件1末尾
        "return_cols": ["ID", "名称", "所属站房"],
        # 文件2匹配键重复时保留第一条
        "keep_duplicates": "first",
        # 单独输出最终未匹配行（两次匹配都失败的）
        "output_unmatched": True,
        # 打印匹配样例
        "verbose": True
    },
# ========== 母线匹配（含二次匹配） ==========
    {
        # 文件1路径（主表）
        "file1": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260629\\fx_mx_629.xlsx",
        # 文件2路径（从表）
        "file2": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260629\\母线段.xlsx",
        # 输出文件路径
        "output": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260629\\母线匹配结果.xlsx",
        # Excel是否有表头：True=有表头，列名使用字符串；False=无表头，列名使用数字索引
        "has_header": True,
        # 第一次对比列配置：每个元组为(文件1列, 文件1预处理规则列表, 文件2列, 文件2预处理规则列表)
        "compare_cols": [
            ('母线名称', ["symbols", "remove_left:站房名称", "left:8", "remove_suffix:2|1"], '设备名称', ["symbols", "left:8", "remove_suffix:2|1"]),
            ('站房名称', ["symbols", "chinese"], '所属站房', ["symbols", "chinese"]),
        ],
        # 二次匹配对比列配置（可选）：仅对第一次未匹配行生效
        # "secondary_compare_cols": [
        #     # 示例：放宽匹配规则，仅匹配站房名称+刀闸名称前6位
        #     ('母线名称', ["symbols", "remove_left:站房名称", "left:2"], '名称', ["symbols", "left:2"]),
        #     ('站房名称', ["symbols", "chinese"], '所属站房', ["symbols", "chinese"]),
        # ],
        # 待返回列：匹配成功后从文件2提取这些列追加到文件1末尾
        "return_cols": ["ID", "设备名称", "所属站房"],
        # 文件2匹配键重复时保留第一条
        "keep_duplicates": "first",
        # 单独输出最终未匹配行（两次匹配都失败的）
        "output_unmatched": True,
        # 打印匹配样例
        "verbose": True
    },
]

# ---------- 预编译正则表达式，提升性能 ----------
# 提取中文的正则（Unicode范围 \u4e00-\u9fa5）
_CHINESE_RE = re.compile(r"[\u4e00-\u9fa5]")
# 提取数字的正则（0-9）
_NUMBER_RE = re.compile(r"\d")
_SYMBOLS_RE = re.compile(r"[^\u4e00-\u9fa5a-zA-Z0-9]")  # 匹配非中文、非字母、非数字的符号

# 需要行上下文（同一行其他列的值）的规则前缀集合
_CONTEXT_RULE_PREFIXES = {"remove_left:", "remove_right:", "trim_left:", "trim_right:"}


def _requires_context(rule_list):
    """判断规则列表中是否包含需要行上下文的跨列规则"""
    for rule in rule_list:
        for prefix in _CONTEXT_RULE_PREFIXES:
            if rule.startswith(prefix):
                return True
    return False


def process_text(text, rule_list, row=None):
    """
    按顺序执行多个预处理规则，支持引用同一行其他列的值。
    参数：
        text: 原始文本值（可能为 NaN 或任意类型）
        rule_list: 规则字符串列表
        row: 可选，当规则涉及跨列引用时必须传入该行的 Series，否则为 None
    返回：清洗后的字符串
    """
    # 空值保护：NaN 或 None 统一转为空字符串
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

        # ---------- 新增规则：去除所有符号 ----------
        elif rule == "symbols":
            # 使用预编译的正则删除所有非中文、非字母、非数字的字符（包括空格、标点等）
            s = _SYMBOLS_RE.sub("", s)

        # ---------- 新增规则：去除指定字符串 ----------
        elif rule.startswith("remove:") and not rule.startswith("remove_left:") and not rule.startswith(
                "remove_right:"):
            # 格式 remove:要删除的文本，第一个冒号后的所有内容视为目标字符串
            target = rule.split(":", 1)[1]
            s = s.replace(target, "")
        # ---------- 新增：多后缀删除（固定字符串，无需行上下文） ----------
        elif rule.startswith("remove_suffix:"):
            suffix_str = rule.split(":", 1)[1]
            # 用竖线分割多个后缀，过滤空串
            suffixes = [suf for suf in suffix_str.split("|") if suf]
            if not suffixes:
                continue
            # 按后缀长度降序排序：优先匹配长后缀，避免短后缀提前截断
            suffixes.sort(key=lambda x: len(x), reverse=True)
            for suf in suffixes:
                if s.endswith(suf):
                    s = s[:-len(suf)]
                    break  # 只删除第一个匹配的后缀，不循环删除

        # ---------- 新增：保留指定字符串及之后内容，删除前面内容 ----------
        # elif rule.startswith("keep_from:"):
        #     # 解析目标字符串（第一个冒号后的所有内容都视为目标）
        #     target = rule.split(":", 1)[1]
        #     # 空目标直接跳过
        #     if not target:
        #         continue
        #     # 查找目标字符串第一次出现的位置
        #     pos = s.find(target)
        #     # 找到则从该位置开始截取（包含目标字符串本身）
        #     if pos != -1:
        #         s = s[pos:]
        #     # 未找到则不处理，保持原字符串不变
        # ---------- 多目标：保留最先出现的指定字符串及之后内容，删除前面内容 ----------
        elif rule.startswith("keep_from:"):
            # 解析规则，用竖线分割多个候选目标字符串，过滤空串
            targets_str = rule.split(":", 1)[1]
            targets = [t for t in targets_str.split("|") if t]
            if not targets:
                continue

            # 遍历所有目标，找到在字符串中最先出现的位置（最小有效索引）
            min_pos = None
            for target in targets:
                pos = s.find(target)
                if pos != -1:
                    # 记录最靠左的出现位置
                    if min_pos is None or pos < min_pos:
                        min_pos = pos

            # 找到任意目标则从最靠前的位置截取（包含目标字符串本身）
            if min_pos is not None:
                s = s[min_pos:]
            # 所有目标都未匹配则不处理，保持原字符串
        # ---------- 字符串截取规则 ----------
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

        elif rule.startswith("cut_left:"):
            try:
                n = int(rule.split(":")[1])
            except (IndexError, ValueError):
                raise ValueError(f"规则格式错误: {rule}，正确示例 cut_left:2")
            s = s[n:] if n <= len(s) else ""

        elif rule.startswith("cut_right:"):
            try:
                n = int(rule.split(":")[1])
            except (IndexError, ValueError):
                raise ValueError(f"规则格式错误: {rule}，正确示例 cut_right:3")
            s = s[:-n] if n <= len(s) else ""

        # ---------- 新增：前4位全为数字则删除前4位 ----------
        elif rule == "remove_left4_if_digits":
            # 长度至少4位，且前4位全部为数字时才删除
            if len(s) >= 4 and s[:4].isdigit():
                s = s[4:]
            # 不满足条件则保持原字符串不变

        # ---------- 跨列引用规则（需要行上下文） ----------
        elif rule.startswith("remove_left:"):
            ref_col = rule.split(":", 1)[1]
            if row is None:
                raise ValueError(f"规则 {rule} 需要行上下文，但未提供 row 参数。")
            ref_val = str(row[ref_col]) if pd.notna(row[ref_col]) else ""
            if ref_val and s.startswith(ref_val):
                s = s[len(ref_val):]
        elif rule.startswith("remove_right:"):
            ref_col = rule.split(":", 1)[1]
            if row is None:
                raise ValueError(f"规则 {rule} 需要行上下文，但未提供 row 参数。")
            ref_val = str(row[ref_col]) if pd.notna(row[ref_col]) else ""
            if ref_val and s.endswith(ref_val):
                s = s[:-len(ref_val)]
        elif rule.startswith("trim_left:"):
            ref_col = rule.split(":", 1)[1]
            if row is None:
                raise ValueError(f"规则 {rule} 需要行上下文，但未提供 row 参数。")
            ref_val = str(row[ref_col]) if pd.notna(row[ref_col]) else ""
            n = len(ref_val)  # 截取长度 = 指定列字符数
            # 边界保护：截取长度超过字符串总长时返回空，避免索引异常
            s = s[n:] if n <= len(s) else ""
        elif rule.startswith("trim_right:"):
            ref_col = rule.split(":", 1)[1]
            if row is None:
                raise ValueError(f"规则 {rule} 需要行上下文，但未提供 row 参数。")
            ref_val = str(row[ref_col]) if pd.notna(row[ref_col]) else ""
            n = len(ref_val)  # 截取长度 = 指定列字符数
            s = s[:-n] if n <= len(s) else ""

        # 未知规则跳过（保持扩展性）
    return s


def _can_vectorize(rule_list):
    """
    判断规则列表是否全部为简单操作（none, strip, lower, upper）。
    这些操作可以用 pandas 的向量化方法直接处理，无需逐行 apply。
    返回 True 表示可以向量化，False 表示需要逐行处理。
    """
    allowed = {"none", "strip", "lower", "upper"}
    for rule in rule_list:
        if rule not in allowed:
            return False
    return True


def _vectorized_process(series, rule_list):
    """
    使用 pandas 向量化方法处理整列数据（Series）。
    仅处理 none、strip、lower、upper 这些简单规则，
    因为调用前已通过 _can_vectorize 保证规则列表只包含这些操作。
    """
    # 将整列转换为字符串类型（统一处理）
    s = series.astype(str)
    for rule in rule_list:
        if rule == "none":
            continue
        elif rule == "strip":
            s = s.str.strip()  # 对整个 Series 去除首尾空格
        elif rule == "lower":
            s = s.str.lower()  # 整体小写化
        elif rule == "upper":
            s = s.str.upper()  # 整体大写化
        # 注意：这里不包含 chinese、number 等复杂规则，由调用逻辑保证
    return s


def _match_data(df1, df2, compare_cols, return_cols, keep_dup, has_header):
    """
    通用匹配函数：根据指定的对比列规则执行匹配，返回匹配结果
    参数：
        df1: 文件1的DataFrame
        df2: 文件2的DataFrame
        compare_cols: 对比列配置
        return_cols: 待返回列
        keep_dup: 去重策略
        has_header: 是否有表头
    返回：
        matched_result: 匹配后的DataFrame
        key_cols: 生成的临时键列列表（用于后续删除）
    """
    # 生成临时匹配键
    key_cols = []
    for idx, (col1, rules1, col2, rules2) in enumerate(compare_cols):
        key_name = f"_key_{idx}"
        key_cols.append(key_name)

        # 处理文件1的键列
        if _can_vectorize(rules1):
            df1[key_name] = _vectorized_process(df1[col1], rules1)
        else:
            if _requires_context(rules1):
                df1[key_name] = df1.apply(lambda row: process_text(row[col1], rules1, row), axis=1)
            else:
                df1[key_name] = df1[col1].apply(lambda x: process_text(x, rules1))

        # 处理文件2的键列
        if _can_vectorize(rules2):
            df2[key_name] = _vectorized_process(df2[col2], rules2)
        else:
            if _requires_context(rules2):
                df2[key_name] = df2.apply(lambda row: process_text(row[col2], rules2, row), axis=1)
            else:
                df2[key_name] = df2[col2].apply(lambda x: process_text(x, rules2))

    # 文件2匹配键去重
    dup_mask = df2.duplicated(subset=key_cols, keep=False)
    dup_count = dup_mask.sum()
    if dup_count > 0:
        print(f"⚠️ 文件2中匹配键存在 {dup_count} 行重复，将按照 keep='{keep_dup}' 保留")
    df2_unique = df2.drop_duplicates(subset=key_cols, keep=keep_dup)

    # 提取匹配所需列
    df2_merge = df2_unique[key_cols + return_cols].copy()

    # 多列联合左连接
    matched_result = df1.merge(df2_merge, on=key_cols, how="left")

    # 处理列名冲突（有表头模式）
    if has_header:
        rename_map = {}
        for col in return_cols:
            if col in df1.columns:
                rename_map[col] = f"{col}_来自文件2"
        if rename_map:
            matched_result.rename(columns=rename_map, inplace=True)

    return matched_result, key_cols


def run_single_task(task):
    """
    执行单个对比任务（支持二次匹配）：
    1. 读取文件1和文件2
    2. 校验列存在性
    3. 第一次匹配：按主规则匹配所有行
    4. 分离第一次匹配成功/未匹配行
    5. 第二次匹配：对未匹配行按二次规则匹配（配置了的话）
    6. 合并两次匹配结果（第一次结果优先，不覆盖）
    7. 统计匹配率、输出结果、打印样例
    """
    # 记录任务开始时间，用于统计耗时
    start_time = time.time()

    # ----- 1. 读取任务参数，设置默认值 -----
    file1_path = task["file1"]
    file2_path = task["file2"]
    output_path = task["output"]
    has_header = task["has_header"]
    compare_cols = task["compare_cols"]
    # 读取二次匹配规则（可选，无则为None）
    secondary_compare_cols = task.get("secondary_compare_cols")
    return_cols = task["return_cols"]
    keep_dup = task.get("keep_duplicates", "first")
    output_unmatched = task.get("output_unmatched", False)
    verbose = task.get("verbose", True)

    # ----- 2. 读取 Excel 文件 -----
    # 统一以字符串类型读取，避免数字/日期等自动转换导致的格式差异
    read_param = {"dtype": str, "keep_default_na": False, "na_values": []}
    if not has_header:
        read_param["header"] = None

    df1 = pd.read_excel(file1_path, **read_param)
    df2 = pd.read_excel(file2_path, **read_param)
    # 复制原始df2用于二次匹配（避免键列污染）
    df2_secondary = df2.copy()

    # 记录文件1原始列数
    original_col_count = len(df1.columns)
    print(f"\n{'=' * 50}")
    print(f"处理任务：{file1_path}  ↔  {file2_path}")
    print(f"文件1原有 {original_col_count} 列，返回列将从第 {original_col_count + 1} 列开始追加")
    print(f"配置对比列数：{len(compare_cols)} 列，待返回列数：{len(return_cols)} 列")
    if secondary_compare_cols:
        print(f"配置二次匹配对比列数：{len(secondary_compare_cols)} 列")

    # ----- 3. 列存在性校验 -----
    # 校验第一次匹配列
    for idx, (col1, rules1, col2, rules2) in enumerate(compare_cols):
        if col1 not in df1.columns:
            raise KeyError(f"文件1中不存在列: {col1}（可用列：{list(df1.columns)}）")
        if col2 not in df2.columns:
            raise KeyError(f"文件2中不存在列: {col2}（可用列：{list(df2.columns)}）")
    # 校验二次匹配列（如果配置了）
    if secondary_compare_cols:
        for idx, (col1, rules1, col2, rules2) in enumerate(secondary_compare_cols):
            if col1 not in df1.columns:
                raise KeyError(f"文件1中不存在二次匹配列: {col1}（可用列：{list(df1.columns)}）")
            if col2 not in df2_secondary.columns:
                raise KeyError(f"文件2中不存在二次匹配列: {col2}（可用列：{list(df2_secondary.columns)}）")
    # 校验返回列
    for col in return_cols:
        if col not in df2.columns:
            raise KeyError(f"文件2中不存在待返回列: {col}（可用列：{list(df2.columns)}）")

    # ----- 4. 第一次匹配：执行主规则匹配 -----
    print("\n【第一次匹配】")
    # 执行第一次匹配
    first_match_result, first_key_cols = _match_data(
        df1.copy(), df2, compare_cols, return_cols, keep_dup, has_header
    )
    # 删除临时键列
    first_match_result.drop(columns=first_key_cols, inplace=True)

    # 统计第一次匹配结果
    first_return_col = first_match_result.columns[original_col_count]
    first_matched = first_match_result[first_return_col].notna().sum()
    total = len(first_match_result)
    first_unmatched = total - first_matched

    # 打印第一次匹配统计
    print(f"总行数：{total}")
    print(f"第一次匹配成功：{first_matched} 行 ({first_matched / total * 100:.1f}%)")
    print(f"第一次未匹配：{first_unmatched} 行")

    # ----- 5. 分离第一次匹配成功/未匹配行 -----
    # 第一次匹配成功的行
    first_matched_rows = first_match_result[first_match_result[first_return_col].notna()].copy()
    # 第一次未匹配的行（保留原始结构，用于二次匹配）
    first_unmatched_rows = first_match_result[first_match_result[first_return_col].isna()].copy()
    # 提取未匹配行的原始数据（仅保留文件1原始列）用于二次匹配
    unmatched_original = first_unmatched_rows[df1.columns].copy()

    # ----- 6. 第二次匹配：对未匹配行执行二次规则（如果配置了） -----
    second_matched = 0
    second_matched_rows = pd.DataFrame()
    if secondary_compare_cols and not unmatched_original.empty:
        print("\n【第二次匹配】")
        print(f"待匹配行数：{len(unmatched_original)}")

        # 执行二次匹配
        second_match_result, second_key_cols = _match_data(
            unmatched_original, df2_secondary, secondary_compare_cols, return_cols, keep_dup, has_header
        )
        # 删除临时键列
        second_match_result.drop(columns=second_key_cols, inplace=True)

        # 统计二次匹配结果
        second_matched = second_match_result[first_return_col].notna().sum()
        second_unmatched = len(second_match_result) - second_matched

        # 提取二次匹配成功的行
        second_matched_rows = second_match_result[second_match_result[first_return_col].notna()].copy()

        print(f"第二次匹配成功：{second_matched} 行 ({second_matched / len(unmatched_original) * 100:.1f}%)")
        print(f"第二次未匹配：{second_unmatched} 行")

    # ----- 7. 合并两次匹配结果 -----
    # 合并规则：第一次匹配成功行 + 第二次匹配成功行
    # 保证第一次结果不被覆盖，仅补充第二次匹配成功的行
    final_result = pd.concat([first_matched_rows, second_matched_rows], ignore_index=True)

    # 补充最终未匹配的行（两次都没匹配上的）
    if secondary_compare_cols and not unmatched_original.empty:
        # 提取二次匹配未成功的行
        second_unmatched_rows = second_match_result[second_match_result[first_return_col].isna()].copy()
        final_result = pd.concat([final_result, second_unmatched_rows], ignore_index=True)
    elif not first_matched_rows.empty:
        # 无二次匹配时，补充第一次未匹配行
        final_result = pd.concat([final_result, first_unmatched_rows], ignore_index=True)

    # 恢复原始行顺序（按索引排序）
    final_result = final_result.sort_index().reset_index(drop=True)

    # ----- 8. 总匹配统计 -----
    total_matched = first_matched + second_matched
    total_match_rate = total_matched / total * 100 if total > 0 else 0
    final_unmatched = total - total_matched
    elapsed = time.time() - start_time

    print("\n【匹配汇总】")
    print(f"总匹配成功：{total_matched} 行 ({total_match_rate:.1f}%)")
    print(f"最终未匹配：{final_unmatched} 行")
    print(f"任务总耗时：{elapsed:.2f} 秒")

    # ----- 9. 打印匹配样例 -----
    if verbose and total_matched > 0:
        print("\n前3条匹配成功样例：")
        show_cols = list(final_result.columns[:original_col_count]) + [first_return_col]
        try:
            # 取前3条匹配成功的行展示
            sample_rows = final_result[final_result[first_return_col].notna()].head(3)
            print(sample_rows[show_cols].to_string(index=False))
        except Exception:
            print("（样例展示失败）")

    # ----- 10. 保存最终结果文件 -----
    final_result.to_excel(output_path, index=False, header=has_header)
    print(f"\n最终结果已保存到：{output_path}")

    # ----- 11. 输出最终未匹配行（两次都失败的）-----
    # if output_unmatched:
    #     final_unmatched_rows = final_result[final_result[first_return_col].isna()]
    #     if not final_unmatched_rows.empty:
    #         unmatched_path = output_path.replace(".xlsx", "_未匹配.xlsx")
    #         final_unmatched_rows.to_excel(unmatched_path, index=False, header=has_header)
    #         print(f"最终未匹配行已保存到：{unmatched_path}")
    #     else:
    #         print("所有行均已匹配，无需生成未匹配文件。")


def main():
    """主函数：遍历所有配置任务，依次执行，并捕获异常保证任务独立性"""
    total_tasks = len(TASKS)
    print(f"共配置 {total_tasks} 个任务，开始批量处理...\n")

    success = 0
    for i, task in enumerate(TASKS, 1):
        try:
            run_single_task(task)
            success += 1
        except Exception as e:
            # 单个任务失败不影响后续任务，打印错误信息并继续
            print(f"\n❌ 任务 {i} 执行失败：{e}")

    print(f"\n{'=' * 50}")
    print(f"全部任务处理完成：成功 {success}/{total_tasks}")


# 程序入口
if __name__ == "__main__":
    main()