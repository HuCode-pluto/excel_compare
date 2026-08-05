import pandas as pd
import re
import time
import os




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
# ========== 开关站匹配（含二次匹配） ==========
    {
        # 文件1路径（主表）
        "file1": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260804\\kgz_804.xls",
        # 文件2路径（从表）
        "file2": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260804\\配电站房.xlsx",
        # 文件3路径
        "file3": "C:\\Users\\13303\\Desktop\\work\\fxdata\\示例\\2工具匹配成功后生成\\调控云设备模型20260804\\奉贤地调设备匹配情况总表-0513 - 匹配结果.xlsx",
        "sheet3": "站房",   # 可选，不指定则使用第一个Sheet

        # 输出文件路径
        "output": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260804\\开关站匹配结果.xlsx",
        # Excel是否有表头：True=有表头，列名使用字符串；False=无表头，列名使用数字索引
        "has_header": True,
        # 第一次对比列配置：每个元组为(文件1列, 文件1预处理规则列表, 文件2列, 文件2预处理规则列表)
        "compare_cols": [
            ('站房NAME', ["symbols"], '站房名称',["symbols"]),
        ],
        # 待返回列：匹配成功后从文件2提取这些列追加到文件1末尾
        "return_cols": ["ID", "站房名称", ],
        # 文件2匹配键重复时保留第一条
        "keep_duplicates": "first",
        # 单独输出最终未匹配行（两次匹配都失败的）
        "output_unmatched": False,
        # 打印匹配样例
        "verbose": True,

        "three_compare_cols": [
            ('站房NAME', ["symbols"], '地调站房名称', ["symbols"]),
        ],
        "three_return_cols": ["调控云ID", "调控云名称"],
        "three_keep_duplicates": "first"   # 可选，默认与keep_duplicates一致
    },
    # ========== 刀闸匹配（含二次匹配） ==========
    {
        # 文件1路径（主表）
        "file1": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260804\\dz_804.xls",
        # 文件2路径（从表）
        "file2": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260804\\刀闸.xlsx",
        # 文件3路径
        "file3": "C:\\Users\\13303\\Desktop\\work\\fxdata\\示例\\2工具匹配成功后生成\\调控云设备模型20260804\\奉贤地调设备匹配情况总表-0513 - 匹配结果.xlsx",
        "sheet3": "刀闸",  # 可选，不指定则使用第一个Sheet
        # 输出文件路径
        "output": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260804\\刀闸匹配结果.xlsx",
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
        "output_unmatched": False,
        # 打印匹配样例
        "verbose": True,

        "three_compare_cols": [
            ('刀闸名称', ["symbols"], '地调设备名称', ["symbols"]),
            ('站房名称', ["symbols", "chinese"], '站房名称1', ["symbols", "chinese"]),
        ],
        "three_return_cols": ["调控云设备ID", "调控云设备名称","站房名称1"],
        "three_keep_duplicates": "first"   # 可选，默认与keep_duplicates一致

    },
# ========== 断路器/负荷开关匹配（含二次匹配） ==========
    {
        # 文件1路径（主表）
        "file1": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260804\\kg_804.xls",
        # 文件2路径（从表）
        "file2": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260804\\断路器.xlsx",
        # 文件3路径
        "file3": "C:\\Users\\13303\\Desktop\\work\\fxdata\\示例\\2工具匹配成功后生成\\调控云设备模型20260804\\奉贤地调设备匹配情况总表-0513 - 匹配结果.xlsx",
        "sheet3": "开关",  # 可选，不指定则使用第一个Sheet
        # 输出文件路径
        "output": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260804\\开关匹配结果.xlsx",
        # Excel是否有表头：True=有表头，列名使用字符串；False=无表头，列名使用数字索引
        "has_header": True,
        # 第一次对比列配置：每个元组为(文件1列, 文件1预处理规则列表, 文件2列, 文件2预处理规则列表)
        "compare_cols": [
            ('开关名称', ["symbols", "remove:自建", "remove_left:站房名称", "left:8","remove_suffix:2|1"], '名称',["symbols","remove_left4_if_digits","left:8", "remove_suffix:2|1"]),
            ('站房名称', ["symbols", "chinese"], '所属站房', ["symbols", "chinese"]),
        ],
        # 二次匹配对比列配置（可选）：仅对第一次未匹配行生效
        "secondary_compare_cols": [
            # 示例：放宽匹配规则，仅匹配站房名称+刀闸名称前2位
            ('开关名称', ["symbols", "remove_left:站房名称", "left:2"], '名称', ["symbols", "remove_left4_if_digits","left:2"]),
            ('站房名称', ["symbols", "chinese"], '所属站房', ["symbols", "chinese"]),
        ],
        # 待返回列：匹配成功后从文件2提取这些列追加到文件1末尾
        "return_cols": ["ID", "名称", "所属站房"],
        # 文件2匹配键重复时保留第一条
        "keep_duplicates": "first",
        # 单独输出最终未匹配行（两次匹配都失败的）
        "output_unmatched": False,
        # 打印匹配样例
        "verbose": True,

        "three_compare_cols": [
            ('开关名称', ["symbols"], '地调设备名称', ["symbols"]),
            ('站房名称', ["symbols", "chinese"], '站房名称橙色站房未匹配', ["symbols", "chinese"]),
        ],
        "three_return_cols": ["调控云ID", "调控云设备名称", "站房名称橙色站房未匹配"],
        "three_keep_duplicates": "first"  # 可选，默认与keep_duplicates一致

    },
# ========== 接地刀闸开关匹配（含二次匹配） ==========
    {
        # 文件1路径（主表）
        "file1": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260804\\jddz_804.xls",
        # 文件2路径（从表）
        "file2": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260804\\接地刀闸.xlsx",
        # 文件3路径
        "file3": "C:\\Users\\13303\\Desktop\\work\\fxdata\\示例\\2工具匹配成功后生成\\调控云设备模型20260804\\奉贤地调设备匹配情况总表-0513 - 匹配结果.xlsx",
        "sheet3": "接地刀闸",  # 可选，不指定则使用第一个Sheet
        # 输出文件路径
        "output": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260804\\接地刀闸匹配结果.xlsx",
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
        "output_unmatched": False,
        # 打印匹配样例
        "verbose": True,

        "three_compare_cols": [
            ('接地刀名称', ["symbols"], '地调设备名称', ["symbols"]),
            ('站房名称', ["symbols", "chinese"], '地调站房名称', ["symbols", "chinese"]),
        ],
        "three_return_cols": ["调控云设备ID", "调控云设备名称", "地调站房名称"],
        "three_keep_duplicates": "first"  # 可选，默认与keep_duplicates一致

    },
# ========== 母线匹配（含二次匹配） ==========
    {
        # 文件1路径（主表）
        "file1": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260804\\mx_804.xls",
        # 文件2路径（从表）
        "file2": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260804\\母线段.xlsx",
        # 文件3路径
        "file3": "C:\\Users\\13303\\Desktop\\work\\fxdata\\示例\\2工具匹配成功后生成\\调控云设备模型20260804\\奉贤地调设备匹配情况总表-0513 - 匹配结果.xlsx",
        "sheet3": "母线",  # 可选，不指定则使用第一个Sheet
        # 输出文件路径
        "output": "C:\\Users\\13303\\Desktop\\work\\fxdata\示例\\2工具匹配成功后生成\\调控云设备模型20260804\\母线匹配结果.xlsx",
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
        "output_unmatched": False,
        # 打印匹配样例
        "verbose": True,

        "three_compare_cols": [
            ('母线名称', ["symbols"], '地调设备名称', ["symbols"]),
            ('站房名称', ["symbols", "chinese"], '站房名称橙色站房未匹配', ["symbols", "chinese"]),
        ],
        "three_return_cols": ["调控云设备ID", "调控云设备名称", "站房名称橙色站房未匹配"],
        "three_keep_duplicates": "first"  # 可选，默认与keep_duplicates一致

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
    判断规则列表是否全部可以向量化处理（无需行上下文）。
    支持更多规则：none, strip, lower, upper, chinese, number, symbols,
                 left:n, right:n, mid:s:n, cut_left:n, cut_right:n,
                 remove:text, remove_left4_if_digits
    返回 True 表示可以向量化，False 表示需要逐行处理。
    """
    for rule in rule_list:
        if _requires_context_for_rule(rule):
            return False
        # 检查是否是已知可向量化的规则
        if not _is_vectorizable_rule(rule):
            return False
    return True


def _requires_context_for_rule(rule):
    """判断单个规则是否需要行上下文（跨列引用）"""
    if rule == "none" or rule is None:
        return False
    for prefix in _CONTEXT_RULE_PREFIXES:
        if rule.startswith(prefix):
            return True
    return False


def _is_vectorizable_rule(rule):
    """判断单个规则是否可向量化"""
    if rule in {"none", "strip", "lower", "upper", "chinese", "number",
                "symbols", "remove_left4_if_digits"}:
        return True
    if rule.startswith(("left:", "right:", "mid:", "cut_left:", "cut_right:", "remove_suffix:")):
        return True
    if rule.startswith("remove:") and not rule.startswith("remove_left:") and not rule.startswith("remove_right:"):
        return True
    if rule.startswith("keep_from:"):
        # keep_from 涉及多目标+find，暂不向量化但也不需要行上下文
        return False
    return False


def _vectorized_process(series, rule_list):
    """
    使用 pandas 向量化方法处理整列数据（Series），大幅提升性能。
    支持 strip, lower, upper, chinese, number, symbols,
         left:n, right:n, mid:s:n, cut_left:n, cut_right:n,
         remove:text, remove_left4_if_digits 等规则。
    """
    # 统一转为字符串
    s = series.fillna("").astype(str)

    for rule in rule_list:
        if rule == "none":
            continue
        elif rule == "strip":
            s = s.str.strip()
        elif rule == "lower":
            s = s.str.lower()
        elif rule == "upper":
            s = s.str.upper()
        elif rule == "chinese":
            # 只保留中文字符
            s = s.str.replace(r"[^一-龥]", "", regex=True)
        elif rule == "number":
            # 只保留数字
            s = s.str.replace(r"[^0-9]", "", regex=True)
        elif rule == "symbols":
            # 删除所有非中文、非字母、非数字的符号
            s = s.str.replace(r"[^一-龥a-zA-Z0-9]", "", regex=True)
        elif rule.startswith("left:"):
            n = int(rule.split(":")[1])
            s = s.str[:n]
        elif rule.startswith("right:"):
            n = int(rule.split(":")[1])
            s = s.str[-n:]
        elif rule.startswith("mid:"):
            parts = rule.split(":")
            start = int(parts[1])
            length = int(parts[2])
            s = s.str[start:start + length]
        elif rule.startswith("cut_left:"):
            n = int(rule.split(":")[1])
            s = s.str[n:]
        elif rule.startswith("cut_right:"):
            n = int(rule.split(":")[1])
            s = s.str[:-n]
        elif rule.startswith("remove:") and not rule.startswith("remove_left:") \
                and not rule.startswith("remove_right:") and not rule.startswith("remove_suffix:"):
            target = rule.split(":", 1)[1]
            s = s.str.replace(target, "", regex=False)
        elif rule.startswith("remove_suffix:"):
            # 多后缀删除：构建正则 (suffix1|suffix2|...)$ 按长度降序匹配
            suffix_str = rule.split(":", 1)[1]
            suffixes = [suf for suf in suffix_str.split("|") if suf]
            if suffixes:
                # 按长度降序排序
                suffixes.sort(key=lambda x: len(x), reverse=True)
                # 转义特殊正则字符后拼接
                escaped = [re.escape(suf) for suf in suffixes]
                pattern = "(" + "|".join(escaped) + ")$"
                s = s.str.replace(pattern, "", regex=True)
        elif rule == "remove_left4_if_digits":
            # 前4位全是数字则删除
            mask = s.str.len() >= 4
            is_digit_prefix = s.str[:4].str.match(r"^\d{4}$")
            combined_mask = mask & is_digit_prefix
            s = s.where(~combined_mask, s.str[4:])
        # 不支持向量化的规则（如 remove_suffix, keep_from 等）不会走到这里，
        # 因为 _can_vectorize 已过滤

    return s


def _match_data(df1, df2, compare_cols, return_cols, keep_dup, has_header, suffix="_来自文件2", force_suffix=False):
    """
    通用匹配函数：根据指定的对比列规则执行匹配，返回匹配结果
    参数：
        df1: 文件1的DataFrame
        df2: 文件2的DataFrame
        compare_cols: 对比列配置
        return_cols: 待返回列
        keep_dup: 去重策略
        has_header: 是否有表头
        suffix: 列名冲突时追加的后缀（默认 "_来自文件2"）
    返回：
        matched_result: 匹配后的DataFrame
        key_cols: 生成的临时键列列表（用于后续删除）
    """

    # 不对原始 df1/df2 做修改，在副本上生成临时匹配键
    df1 = df1.copy()
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
    # if has_header:
    #     rename_map = {}
    #     for col in return_cols:
    #         if col in df1.columns:
    #             rename_map[col] = f"{col}_来自文件2"
    #     if rename_map:
    #         matched_result.rename(columns=rename_map, inplace=True)
    if has_header:
        rename_map = {}
        for col in return_cols:
            # 如果强制添加后缀，或者列名在df1中已存在，则重命名
            if force_suffix or (col in df1.columns):
                rename_map[col] = f"{col}{suffix}"
        if rename_map:
            matched_result.rename(columns=rename_map, inplace=True)

    return matched_result, key_cols


def _safe_pct(numerator, denominator):
    """安全计算百分比，避免除零返回 NaN"""
    if denominator == 0:
        return 0.0
    return numerator / denominator * 100


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

    # ===== 新增：读取第三次匹配参数（可选） =====
    file3_path = task.get("file3")
    sheet3 = task.get("sheet3")  # 可选，默认为None
    three_compare_cols = task.get("three_compare_cols")
    three_return_cols = task.get("three_return_cols")
    three_keep_dup = task.get("three_keep_duplicates", keep_dup)  # 默认与主配置一致

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
    df1['_row_id'] = df1.index  # 用于后续匹配后排序和合并
    print(f"\n{'=' * 50}")
    print(f"处理任务：{file1_path}  ↔  {file2_path}")
    print(f"文件1原有 {original_col_count} 列，返回列将从第 {original_col_count + 1} 列开始追加")
    print(f"配置对比列数：{len(compare_cols)} 列，待返回列数：{len(return_cols)} 列")
    if secondary_compare_cols:
        print(f"配置二次匹配对比列数：{len(secondary_compare_cols)} 列")

    # ===== 新增：打印三次匹配信息 =====
    if file3_path and three_compare_cols:
        print(f"配置第三次匹配（文件3：{file3_path}，Sheet：{sheet3 or '默认'}）")


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

    # ===== 新增：校验第三次匹配列 + 预加载文件3（避免重复读取）=====
    df3_preloaded = None  # 缓存文件3，避免重复读取
    if file3_path and three_compare_cols:
        # 读取文件3一次，用于校验列并缓存供后续匹配使用
        read_param3 = {"dtype": str, "keep_default_na": False, "na_values": []}
        if not has_header:
            read_param3["header"] = None
        if sheet3:
            df3_preloaded = pd.read_excel(file3_path, sheet_name=sheet3, **read_param3)
        else:
            df3_preloaded = pd.read_excel(file3_path, **read_param3)
        for idx, (col1, rules1, col2, rules2) in enumerate(three_compare_cols):
            if col1 not in df1.columns:
                raise KeyError(f"文件1中不存在第三次匹配列: {col1}（可用列：{list(df1.columns)}）")
            if col2 not in df3_preloaded.columns:
                raise KeyError(f"文件3中不存在列: {col2}（可用列：{list(df3_preloaded.columns)}）")
        for col in three_return_cols:
            if col not in df3_preloaded.columns:
                raise KeyError(f"文件3中不存在待返回列: {col}（可用列：{list(df3_preloaded.columns)}）")

    # ----- 4. 第一次匹配：执行主规则匹配 -----
    print("\n【第一次匹配】")
    t1 = time.time()
    # 执行第一次匹配（_match_data 内部在副本上操作，不修改原 df）
    first_match_result, first_key_cols = _match_data(
        df1, df2, compare_cols, return_cols, keep_dup, has_header, suffix="_来自文件2"
    )
    # 删除临时键列
    first_match_result.drop(columns=first_key_cols, inplace=True)
    print(f"  第一次匹配耗时：{time.time() - t1:.2f} 秒")

    # 统计第一次匹配结果
    # first_return_col = first_match_result.columns[original_col_count]
    first_return_col = first_match_result.columns[original_col_count + 1]
    first_matched = first_match_result[first_return_col].notna().sum()
    total = len(first_match_result)
    first_unmatched = total - first_matched

    # 打印第一次匹配统计
    print(f"总行数：{total}")
    print(f"第一次匹配成功：{first_matched} 行 ({_safe_pct(first_matched, total):.1f}%)")
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

        print(f"第二次匹配成功：{second_matched} 行 ({_safe_pct(second_matched, len(unmatched_original)):.1f}%)")
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
    # final_result = final_result.sort_index().reset_index(drop=True)
    # 恢复原始行顺序（按_row_id排序）
    final_result = final_result.sort_values('_row_id').reset_index(drop=True)
    # ----- 8. 第三次匹配（如果配置了） -----
    third_matched = 0
    third_to_return_mapping = {}  # {third_match中的列名: final_result中的return列名}
    if file3_path and three_compare_cols:
        # 提取当前仍未匹配的行（前两次都未匹配，即 first_return_col 为空）
        unmatched_mask = final_result[first_return_col].isna()
        if unmatched_mask.any():
            print("\n【第三次匹配】")
            unmatched_rows = final_result[unmatched_mask].copy()
            print(f"待匹配行数：{len(unmatched_rows)}")

            # 直接使用已预加载的文件3（避免重复读取）
            df3 = df3_preloaded

            # 执行第三次匹配（force_suffix=True 确保列名带后缀防冲突）
            third_match_result, third_key_cols = _match_data(
                unmatched_rows[df1.columns].copy(),
                df3,
                three_compare_cols,
                three_return_cols,
                three_keep_dup,
                has_header,
                suffix="_来自文件3",
                force_suffix=True
            )
            third_match_result.drop(columns=third_key_cols, inplace=True)

            # 统计第三次匹配结果
            third_return_col_name = third_match_result.columns[original_col_count + 1]
            third_matched = third_match_result[third_return_col_name].notna().sum()
            third_unmatched = len(third_match_result) - third_matched
            print(f"第三次匹配成功：{third_matched} 行 ({_safe_pct(third_matched, len(unmatched_rows)):.1f}%)")
            print(f"第三次未匹配：{third_unmatched} 行")

            # ===== 构建 three_return_cols → return_cols 的映射 =====
            # 位置对应：three_return_cols[i] 的值填入 return_cols[i]
            for i, three_col in enumerate(three_return_cols):
                if i >= len(return_cols):
                    print(f"⚠️ 三次匹配返回列 '{three_col}' 超出 return_cols 范围，跳过")
                    continue
                # 三次匹配结果中的实际列名（force_suffix=True 总是加后缀）
                three_actual = f"{three_col}_来自文件3"
                if three_actual not in third_match_result.columns:
                    # 兜底：可能没加后缀（比如列名本来就不冲突）
                    three_actual = three_col
                # final_result 中 return_cols 的实际列名（可能加了 _来自文件2 后缀）
                return_original = return_cols[i]
                if return_original in final_result.columns:
                    return_actual = return_original
                elif f"{return_original}_来自文件2" in final_result.columns:
                    return_actual = f"{return_original}_来自文件2"
                else:
                    print(f"⚠️ 找不到 return_cols 列 '{return_original}'，跳过映射")
                    continue
                third_to_return_mapping[three_actual] = return_actual

            if third_to_return_mapping:
                print(f"  三次匹配列映射：{dict((v, k) for k, v in third_to_return_mapping.items())}")

            # 将匹配成功的行更新到 final_result 的对应 return_cols
            if third_matched > 0:
                third_matched_rows = third_match_result[third_match_result[third_return_col_name].notna()]
                final_result.set_index('_row_id', inplace=True)
                third_matched_rows.set_index('_row_id', inplace=True)
                for three_col, return_col in third_to_return_mapping.items():
                    # 用三次匹配的值覆盖空白的 return_col（只更新，不覆盖已有值）
                    mask = final_result[return_col].isna()
                    final_result.loc[mask, return_col] = third_matched_rows.loc[mask, three_col]
                final_result.reset_index(drop=False, inplace=True)
        else:
            print("\n【第三次匹配】所有行已在前两次匹配成功，跳过。")

    # ----- 8. 总匹配统计 -----
    total_matched = first_matched + second_matched + third_matched
    total_match_rate = _safe_pct(total_matched, total)
    final_unmatched = total - total_matched
    elapsed = time.time() - start_time

    print("\n【匹配汇总】")
    print(f"总匹配成功：{total_matched} 行 ({total_match_rate:.1f}%)")
    print(f"最终未匹配：{final_unmatched} 行")
    print(f"任务总耗时：{elapsed:.2f} 秒")

    # ----- 9. 打印匹配样例 -----
    if verbose and total_matched > 0:
        print("\n【前3条匹配成功样例】")
        # 构造显示列：原始列（不含_row_id）+ 返回列
        display_cols = [col for col in df1.columns if col != '_row_id']
        # 添加返回列（文件2的第一个返回列，以及文件3的新增列的前两个）
        if first_return_col:
            display_cols.append(first_return_col)
        # 添加其他 return_cols（返回列的实际名称）
        for col in return_cols:
            actual = col
            if actual not in display_cols and actual in final_result.columns:
                display_cols.append(actual)
            # 也检查带后缀的版本
            for suffix in ["_来自文件2", "_来自文件3"]:
                candidate = f"{col}{suffix}"
                if candidate not in display_cols and candidate in final_result.columns:
                    display_cols.append(candidate)
        try:
            # 取匹配成功的前3行
            success_mask = final_result[first_return_col].notna()
            sample_rows = final_result[success_mask].head(3).copy()
            # 将 NaN 替换为 "未匹配" 提升可读性
            for col in display_cols:
                if col in sample_rows.columns:
                    sample_rows[col] = sample_rows[col].fillna("未匹配")
            print(sample_rows[display_cols].to_string(index=False))
        except Exception:
            print("（样例展示失败）")

    # 打印未匹配样例（便于排查匹配失败原因）
    if verbose and final_unmatched > 0:
        print(f"\n【前3条未匹配样例】（共 {final_unmatched} 行未匹配）")
        display_cols = [col for col in df1.columns if col != '_row_id']
        try:
            unmatched_sample = final_result[final_result[first_return_col].isna()].head(3).copy()
            for col in display_cols:
                if col in unmatched_sample.columns:
                    unmatched_sample[col] = unmatched_sample[col].fillna("(空)")
            print(unmatched_sample[display_cols].to_string(index=False))
        except Exception:
            print("（样例展示失败）")

    # ----- 10. 保存最终结果文件（NaN替换为"未匹配"便于阅读）-----
    # 文件名追加匹配率：XXX（90.0%）.xlsx
    output_result = final_result.copy()
    for col in return_cols:
        # 处理可能被重命名的返回列（原始名和加后缀的版本）
        for actual_col in output_result.columns:
            if actual_col == col or actual_col.startswith(f"{col}_来自"):
                output_result[actual_col] = output_result[actual_col].fillna("未匹配")

    output_result.drop(columns=['_row_id'], inplace=True)
    # 生成带匹配率的文件名
    base, ext = os.path.splitext(output_path)
    rated_path = f"{base}（{total_match_rate:.1f}%）{ext}"
    output_result.to_excel(rated_path, index=False, header=has_header)
    print(f"\n最终结果已保存到：{rated_path}（未匹配行显示为'未匹配'）")

    # ----- 11. 单独输出未匹配行（如果配置了）-----
    if output_unmatched and final_unmatched > 0:
        # 三次匹配都失败的行：first_return_col 为空
        final_unmatched_mask = final_result[first_return_col].isna()
        final_unmatched_rows = final_result[final_unmatched_mask].copy()
        if not final_unmatched_rows.empty:
            final_unmatched_rows.drop(columns=['_row_id'], inplace=True)
            # NaN替换为"(空)"
            final_unmatched_rows = final_unmatched_rows.fillna("(空)")
            base, ext = os.path.splitext(output_path)
            unmatched_path = f"{base}（{total_match_rate:.1f}%）_未匹配{ext}"
            final_unmatched_rows.to_excel(unmatched_path, index=False, header=has_header)
            print(f"未匹配行已单独保存到：{unmatched_path}")
    elif output_unmatched and final_unmatched == 0:
        print("所有行均已匹配，无需生成未匹配文件。")


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