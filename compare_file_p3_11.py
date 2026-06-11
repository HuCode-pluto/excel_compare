# -*- coding: utf-8 -*-
import pandas as pd
import sys
import os
import traceback
from datetime import datetime


# 自定义输出类：同时输出到终端和日志文件
class TeeOutput:
    def __init__(self, terminal, log):
        self.terminal = terminal
        self.log = log
    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)
        self.log.flush()
    def flush(self):
        self.terminal.flush()
        self.log.flush()


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


def trim_str_length(s, max_length):
    if pd.isna(s):
        return ""
    s = full_clean(s)
    result = str(s).strip()
    if max_length is not None:
        result = result[:max_length]
    return result


def convert_station_type(s):
    """中文/数字 统一映射"""
    if pd.isna(s):
        return ""
    s = full_clean(s)
    mapping = {
        "杆变": "0", "箱式变": "1", "分布式电源": "2",
        "非接地刀闸": "2", "配网小车刀闸": "100",
        "非常重要": "1", "比较重要": "2", "一般": "3",
        "横向": "0", "竖向": "1",
        "支撑杆": "0", "断联杆": "1", "耐张杆": "2",
        "跨越杆": "3", "辅杆": "4", "转角杆": "5",
        "黑图": "0", "红图": "1", "黄图": "2",
        "接地刀闸": "3", "光纤纵差联络开关": "15",
        "出线开关": "16", "负荷开关": "4", "电缆头": "7",
        "线路开关": "8", "母联开关": "9", "常规联络开关": "10",
        "分布式电源控制开关": "11", "准同期开关": "12",
        "专用联络开关": "13", "电压型开关": "14", "小车开关": "99",
        "分段开关": "0", "联络开关": "1",
        "电磁型傻瓜开关": "0", "电压型智能开关": "1",
        "不可拆搭": "0", "可拆搭": "1",
        "架空线": "0", "电缆": "1", "专供": "2", "大用户": "3",
        "未变化": "0", "更新": "1", "删除": "2", "新增": "3",
        "投运": "0", "未投运": "1", "待退役": "2",
        "配网开关": "13502", "配网变压器": "13505",
        "配网母线": "13506", "配网刀闸": "13513",
        "配网接地刀闸": "13514", "配网故障指示器": "13523",
        "配网开关站": "13501", "配网终端": "13510",
        "正交变换法": "0", "正则化法": "1", "混合法": "2",
        "风险主题": "1", "运行方式": "2",
        "备用主题3": "3", "备用主题4": "4",
        "开关站": "0", "环网柜": "1", "街坊站": "2",
        "光伏站": "3", "箱式变电站": "4", "配电站": "5",
        "用户站": "6", "电缆分支箱": "7", "低压配变箱": "8",
        "低压电缆分支箱": "9", "低压电缆终端箱": "10",
        "未定义类别": "0", "DTU": "1", "FTU": "2",
        "TTU": "3", "配网子站": "4", "故障指示器": "5",
        "DTU－开关站": "6", "其它": "0",
        "事故总": "1", "预告信号": "2", "动作信号": "3",
        "故障信号": "4", "通讯信号": "5", "控制信号": "6",
        "设备信号": "7", "重合闸信号": "9", "备自投信号": "10",
        "机组人工停机信号": "11", "厂站VQC信号": "12",
        "纵差保护": "13", "短路故障": "19", "接地故障": "20",
        "电池活化": "21", "纵差后备保护": "22",
        "Uab": "1", "Ubc": "2", "Uca": "3",
        "AVC控制对象厂站可调电压上限": "108",
        "AVC控制对象厂站可调电压下限": "109",
    }
    return mapping.get(s, s)


def yn_station_type(s):
    """是/否 转换成 1/0"""
    if pd.isna(s):
        return ""
    s = full_clean(s)
    if s in ("是",):
        return "1"
    elif s in ("否",):
        return "0"
    elif s in ("10kV",):
        return "112871465677750278"
    return s


# ========== 全局配置区 ==========
OUTPUT_DIR = "C:\\Users\\13303\\Desktop\\excel_compare\\compare_result"


# ========== 批量任务配置区 ==========
# ignore_only_row: True=忽略独有行，只对比共同数据；False=严格校验行数一致
TASKS = [
    # 任务1：配网馈线表实时库和达梦对比
    {
        "name": "配网馈线表实时库和达梦对比",
        "file1": "C:\\Users\\13303\\Desktop\\kxb.xls",
        "file2": "C:\\Users\\13303\\Desktop\\kxb_dm.xls",
        "col_map": {
            "馈线ID号": "ID",
            "馈线名称": "NAME",
        },
        "transform_funcs": {
            "馈线ID号": lambda x: trim_str_length(x, 19),
            "馈线名称": full_clean,
        },
        "key_col": "馈线ID号",
        "ignore_only_row": False
    },
    # 任务2：开关站实时库商用库对比
    {
        "name": "开关站实时库商用库对比",
        "file1": "C:\\Users\\13303\\Desktop\\kgzb.xls",
        "file2": "C:\\Users\\13303\\Desktop\\kgzb_dm.xls",
        "col_map": {
            "开关站ID号": "ID",
            "开关站名称": "NAME",
            "所属馈线": "feeder_name",
            "开关站类型": "COMBINED_TYPE",
        },
        "transform_funcs": {
            "开关站ID号": lambda x: trim_str_length(x, 19),
            "开关站名称": full_clean,
            "所属馈线": full_clean,
            "开关站类型": convert_station_type,
        },
        "key_col": "开关站ID号",
        "ignore_only_row": False
    },
    # 任务3：开关实时库商用库对比
    {
        "name": "开关实时库商用库对比",
        "file1": "C:\\Users\\13303\\Desktop\\kgb.xls",
        "file2": "C:\\Users\\13303\\Desktop\\kgb_dm.xls",
        "col_map": {
            "开关ID号": "ID",
            "开关名称": "NAME",
            "所属开关站": "combined_name",
            "所属馈线": "feeder_name",
            "所属组合设备": "composite_switch_name",
            "开关类型": "BRK_TYPE",
            "开关联络类型": "BRK_CONNECT_TYPE",
        },
        "transform_funcs": {
            "开关ID号": lambda x: trim_str_length(x, 19),
            "开关名称": full_clean,
            "所属开关站": full_clean,
            "所属馈线": full_clean,
            "所属组合设备": full_clean,
            "开关类型": convert_station_type,
            "开关联络类型": convert_station_type,
        },
        "key_col": "开关ID号",
        "ignore_only_row": False
    },
    # 任务4：组合开关实时库商用库对比
    {
        "name": "组合开关实时库商用库对比",
        "file1": "C:\\Users\\13303\\Desktop\\zhkgb.xls",
        "file2": "C:\\Users\\13303\\Desktop\\zhkgb_dm.xls",
        "col_map": {
            "组合开关名称": "NAME",
            "所属开关站": "combined_name",
            "所属馈线": "feeder_name",
        },
        "transform_funcs": {
            "组合开关名称": full_clean,
            "所属开关站": full_clean,
            "所属馈线": full_clean,
        },
        "key_col": "组合开关名称",
        "ignore_only_row": False
    },
    # 任务5：刀闸实时库商用库对比
    {
        "name": "刀闸实时库商用库对比",
        "file1": "C:\\Users\\13303\\Desktop\\dzb.xls",
        "file2": "C:\\Users\\13303\\Desktop\\dzb_dm.xls",
        "col_map": {
            "刀闸ID号": "ID",
            "刀闸名称": "NAME",
            "所属开关站": "combined_name",
            "所属馈线": "feeder_name",
            "所属组合设备": "composite_switch_name",
            "刀闸类型": "DISCR_TYPE",
        },
        "transform_funcs": {
            "刀闸ID号": lambda x: trim_str_length(x, 19),
            "刀闸名称": full_clean,
            "所属开关站": full_clean,
            "所属馈线": full_clean,
            "所属组合设备": full_clean,
            "刀闸类型": convert_station_type,
        },
        "key_col": "刀闸ID号",
        "ignore_only_row": False
    },
    # 任务6：接地刀闸实时库商用库对比
    {
        "name": "接地刀闸实时库商用库对比",
        "file1": "C:\\Users\\13303\\Desktop\\jddzb.xls",
        "file2": "C:\\Users\\13303\\Desktop\\jddzb_dm.xls",
        "col_map": {
            "接地刀闸ID号": "ID",
            "接地刀闸名称": "NAME",
            "所属开关站": "combined_name",
            "所属馈线": "feeder_name",
            "所属组合设备": "composite_switch_name",
        },
        "transform_funcs": {
            "接地刀闸ID号": lambda x: trim_str_length(x, 19),
            "接地刀闸名称": full_clean,
            "所属开关站": full_clean,
            "所属馈线": full_clean,
            "所属组合设备": full_clean,
        },
        "key_col": "接地刀闸ID号",
        "ignore_only_row": False
    },
    # 任务7：母线表实时库商用库对比
    {
        "name": "母线表实时库商用库对比",
        "file1": "C:\\Users\\13303\\Desktop\\mxb.xls",
        "file2": "C:\\Users\\13303\\Desktop\\mxb_dm.xls",
        "col_map": {
            "母线ID号": "ID",
            "母线名称": "NAME",
            "所属开关站": "combined_name",
            "所属馈线": "feeder_name",
        },
        "transform_funcs": {
            "母线ID号": lambda x: trim_str_length(x, 19),
            "母线名称": full_clean,
            "所属开关站": full_clean,
            "所属馈线": full_clean,
        },
        "key_col": "母线ID号",
        "ignore_only_row": False
    },
    # 任务8：终端信息表实时库商用库对比
    {
        "name": "终端信息表实时库商用库对比",
        "file1": "C:\\Users\\13303\\Desktop\\zdxxb.xls",
        "file2": "C:\\Users\\13303\\Desktop\\zdxxb_dm.xls",
        "col_map": {
            "终端ID": "ID",
            "终端名称": "NAME",
            "所属开关站": "combined_name",
            "所属馈线": "feeder_name",
            "终端编号": "code",
            "是否通讯状态统计": "if_stat_static",
            "终端类别": "term_type",
            "运行定制区号": "cur_fixed_area",
        },
        "transform_funcs": {
            "终端ID": lambda x: trim_str_length(x, 19),
            "终端名称": full_clean,
            "所属开关站": full_clean,
            "终端编号": full_clean,
            "是否通讯状态统计": yn_station_type,
            "终端类别": convert_station_type,
            "运行定制区号": full_clean,
            "所属馈线": full_clean,
        },
        "key_col": "终端ID",
        "ignore_only_row": False
    },
    # 任务9：保护节点表实时库商用库对比
    {
        "name": "保护节点表实时库商用库对比",
        "file1": "C:\\Users\\13303\\Desktop\\bhjdb.xls",
        "file2": "C:\\Users\\13303\\Desktop\\bhjdb_dm.xls",
        "col_map": {
            "标识": "ID",
            "中文名称": "NAME",
            "所属开关站": "combined_name",
            "所属馈线": "feeder_name",
            "电压类型ID": "BV_ID",
            "类型": "PNT_TYPE",
            "开关数目": "BRK_NUM",
            "相应开关1": "RELY_BRK1",
            "相应开关2": "RELY_BRK2",
        },
        "transform_funcs": {
            "标识": lambda x: trim_str_length(x, 19),
            "中文名称": full_clean,
            "所属开关站": full_clean,
            "电压类型ID": yn_station_type,
            "类型": convert_station_type,
            "开关数目": full_clean,
            "相应开关1": full_clean,
            "相应开关2": full_clean,
        },
        "key_col": "标识",
        "ignore_only_row": False
    },
    # 任务10：测点遥测表实时库商用库对比
    {
        "name": "测点遥测表实时库商用库对比",
        "file1": "C:\\Users\\13303\\Desktop\\cdycb.xls",
        "file2": "C:\\Users\\13303\\Desktop\\cdycb_dm.xls",
        "col_map": {
            "中文名称": "NAME",
            "所属开关站": "combined_name",
            "所属馈线": "feeder_name",
            "电压类型ID": "BV_ID",
            "类型": "PNT_TYPE",
            "开关数目": "BRK_NUM",
            "相应开关1": "RELY_BRK1",
            "相应开关2": "RELY_BRK2",
        },
        "transform_funcs": {
            "中文名称": full_clean,
            "所属开关站": full_clean,
            "电压类型ID": yn_station_type,
            "类型": convert_station_type,
            "开关数目": full_clean,
            "相应开关1": full_clean,
            "相应开关2": full_clean,
        },
        "key_col": "标识",
        "ignore_only_row": False
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
            if col in df1.columns:
                df1[col] = df1[col].apply(func)
            if col in df2.columns:
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

        diff_df = df_common1.compare(df_common2, keep_shape=True, keep_equal=True)
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
