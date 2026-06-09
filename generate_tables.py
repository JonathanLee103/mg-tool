import pandas as pd
import numpy as np
import os
from datetime import datetime


def generate(work_dir=".", log_cb=None, progress_cb=None):
    """执行订单表和库存表生成。

    Args:
        work_dir: 工作目录路径，源文件从此目录读取，输出也写入此目录
        log_cb:   日志回调 log_cb(msg: str)
        progress_cb: 进度回调 progress_cb(pct: int, status: str)
    """
    if log_cb is None:
        log_cb = print
    if progress_cb is None:
        progress_cb = lambda pct, status: None

    def log(msg):
        log_cb(msg)

    def progress(pct, status):
        progress_cb(pct, status)

    # ── 1. 加载数据 ──────────────────────────────────────────

    log("加载源文件...")
    progress(5, "加载采购订单数据...")

    df_po = pd.read_excel(
        os.path.join(work_dir, "系统导出_采购订单数据.xlsx"),
        sheet_name="数据",
        skiprows=1,
    )
    log(f"  采购订单数据: {df_po.shape[0]} 行, {df_po.shape[1]} 列")

    progress(10, "加载订单汇总...")
    df_summary = pd.read_excel(
        os.path.join(work_dir, "系统导出_订单汇总.xls"), sheet_name="Sheet0"
    )
    log(f"  订单汇总: {df_summary.shape[0]} 行, {df_summary.shape[1]} 列")

    progress(15, "加载业务库存数据...")
    df_basic = pd.read_excel(os.path.join(work_dir, "系统导出_业务库存数据.xlsx"))
    log(f"  业务库存数据: {df_basic.shape[0]} 行, {df_basic.shape[1]} 列")

    progress(20, "加载已发货库存明细...")
    df_shipped = pd.read_excel(
        os.path.join(work_dir, "系统导出_已发货库存明细.xlsx")
    )
    log(f"  已发货库存明细: {df_shipped.shape[0]} 行, {df_shipped.shape[1]} 列")

    # ── 2. 库存 UNION ────────────────────────────────────────

    log("\n合并库存数据...")
    progress(25, "合并库存数据...")

    df_basic = df_basic.rename(columns={"件数(张数,根数,支数)": "件数"})

    # 统一列名：已发货的 渠道规格描述 → 渠道规格（与业务库存一致）
    if "渠道规格描述" in df_shipped.columns:
        df_shipped = df_shipped.rename(columns={"渠道规格描述": "渠道规格"})

    # 已发货补充缺失列
    for col in ["库龄", "原料最初入库日期(年月日)", "业务入库日期(年月日)"]:
        if col not in df_shipped.columns:
            df_shipped[col] = np.nan

    # 标记数据来源
    df_basic["_source"] = "basic"
    df_shipped["_source"] = "shipped"

    common_cols = list(set(df_basic.columns) & set(df_shipped.columns))
    log(f"  共同列: {len(common_cols)}")

    df_union = pd.concat(
        [df_basic[common_cols], df_shipped[common_cols]], ignore_index=True
    )
    log(f"  UNION 结果: {df_union.shape[0]} 行")

    # 识别捆包号重叠（同时在两个表中存在）
    bundle_source_count = df_union.groupby("捆包号")["_source"].transform("nunique")
    bundle_dup = bundle_source_count > 1
    dup_count = bundle_dup.sum()
    if dup_count > 0:
        log(f"  捆包号重叠: {dup_count} 行")

    mask_basic = df_union["_source"] == "basic"
    mask_shipped = df_union["_source"] == "shipped"
    mask_basic_only = mask_basic & ~bundle_dup
    mask_shipped_only = mask_shipped & ~bundle_dup

    # 规格: basic(含重叠) → 渠道规格; shipped_only → 规格
    df_union["_规格"] = ""
    df_union.loc[mask_basic, "_规格"] = (
        df_union.loc[mask_basic, "渠道规格"].fillna("").astype(str)
    )
    df_union.loc[mask_shipped_only, "_规格"] = (
        df_union.loc[mask_shipped_only, "规格"].fillna("").astype(str)
    )

    # 首次入库时间: basic(含重叠) → 原料最初入库日期(年月日); shipped_only → 制造出厂发货时间
    df_union["_首次入库时间"] = ""
    df_union.loc[mask_basic, "_首次入库时间"] = (
        df_union.loc[mask_basic, "原料最初入库日期(年月日)"]
        .fillna("").astype(str)
    )
    df_union.loc[mask_shipped_only, "_首次入库时间"] = (
        df_union.loc[mask_shipped_only, "制造出厂发货时间"]
        .fillna("").astype(str)
    )

    # 最近入库日期: basic_only → 业务入库日期(年月日); shipped(含重叠) → 制造出厂发货时间
    df_union["_最近入库日期"] = ""
    df_union.loc[mask_basic_only, "_最近入库日期"] = (
        df_union.loc[mask_basic_only, "业务入库日期(年月日)"]
        .fillna("").astype(str)
    )
    df_union.loc[mask_shipped, "_最近入库日期"] = (
        df_union.loc[mask_shipped, "制造出厂发货时间"]
        .fillna("").astype(str)
    )

    # 重叠行的补充处理: 从对应来源取值
    if dup_count > 0:
        basic_dup_map = df_union[mask_basic & bundle_dup].set_index("捆包号")
        shipped_dup_map = df_union[mask_shipped & bundle_dup].set_index("捆包号")
        mask_sd = mask_shipped & bundle_dup
        mask_bd = mask_basic & bundle_dup
        # shipped重叠行: 规格→渠道规格(basic), 首次→原料最初入库日期(basic)
        df_union.loc[mask_sd, "_规格"] = (
            df_union.loc[mask_sd, "捆包号"]
            .map(basic_dup_map["渠道规格"]).fillna("").astype(str)
        )
        df_union.loc[mask_sd, "_首次入库时间"] = (
            df_union.loc[mask_sd, "捆包号"]
            .map(basic_dup_map["原料最初入库日期(年月日)"]).fillna("").astype(str)
        )
        # basic重叠行: 最近入库日期→制造出厂发货时间(shipped)
        df_union.loc[mask_bd, "_最近入库日期"] = (
            df_union.loc[mask_bd, "捆包号"]
            .map(shipped_dup_map["制造出厂发货时间"]).fillna("").astype(str)
        )

    # 日期格式化: YYYYMMDDHHMMSS → YYYY-MM-DD, YYYYMM → YYYY-MM
    def _fmt_date(val):
        if pd.isna(val) or val == "":
            return val
        s = str(val).strip()
        # 处理 float 转字符串带来的 .0 后缀
        if s.endswith(".0"):
            s = s[:-2]
        # 跳过已格式化的日期
        if "-" in s:
            return s
        if len(s) >= 14:  # YYYYMMDDHHMMSS → YYYY-MM-DD
            return f"{s[:4]}-{s[4:6]}-{s[6:8]}"
        if len(s) == 8:   # YYYYMMDD → YYYY-MM-DD
            return f"{s[:4]}-{s[6:8]}-{s[8:10]}"
        if len(s) == 6:   # YYYYMM → YYYY-MM
            return f"{s[:4]}-{s[4:6]}"
        return s

    df_union["_首次入库时间"] = df_union["_首次入库时间"].apply(_fmt_date)
    df_union["_最近入库日期"] = df_union["_最近入库日期"].apply(_fmt_date)

    progress(30, "格式化日期列...")
    if df_union["业务入库日期"].dtype in ("float64", "int64"):
        df_union["业务入库日期"] = df_union["业务入库日期"].apply(
            lambda x: str(int(x)) if pd.notna(x) else ""
        )
    if df_union["采购交货期"].dtype in ("float64", "int64"):
        df_union["采购交货期"] = df_union["采购交货期"].apply(
            lambda x: _fmt_date(x) if pd.notna(x) else ""
        )

    # ── 3. 构建订单表 ────────────────────────────────────────

    log("\n构建订单表...")
    progress(35, "构建订单表...")

    df_order = df_po.merge(
        df_summary[["产销合同号", "客户零件号", "后处理方式", "精整分流"]],
        left_on="供应商订单子项号",
        right_on="产销合同号",
        how="left",
    )

    spec_parts = df_order["规格描述"].astype(str).str.split("*", n=2, expand=True)

    df_order_out = pd.DataFrame(
        {
            "供应商合同号子项号": df_order["供应商订单子项号"],
            "客户零件号": df_order["客户零件号"],
            "订货规格": df_order["规格描述"],
            "订货牌号": df_order["牌号"],
            "合同月份": df_order["采购交货月"].apply(
                lambda x: _fmt_date(x) if pd.notna(x) else ""
            ),
            "产销合同号": df_order["供应商订单子项号"],
            "厚度/直径": spec_parts[0],
            "宽度": spec_parts[1],
            "长度": "C",
            "订货重量": df_order["采购订货重量"],
            "表面质量": df_order["表面质量"],
            "最终用户名称": df_order["最终用户名称"],
            "终到站港描述": df_order["股份终到站港名称"],
            "运输方式": df_order["运输方式名称"],
            "后处理方式": df_order["后处理方式"],
            "精整分流": df_order["精整分流"],
            "订单量": df_order["采购订货重量"],
            "总明细量": df_order["入库重量"],
            "未交付量": df_order["未交付重量"],
        }
    )

    if "产销合同号_dd" in df_order.columns:
        df_order_out["产销合同号"] = df_order["产销合同号_dd"].fillna(
            df_order_out["产销合同号"]
        )

    log(f"  订单表: {df_order_out.shape[0]} 行, {df_order_out.shape[1]} 列")

    progress(50, "构建库存表...")

    # ── 4. 构建库存表 ────────────────────────────────────────

    log("\n构建库存表...")

    df_inv = df_union.merge(
        df_po[
            [
                "供应商订单子项号",
                "规格描述",
                "表面质量",
                "采购订货重量",
                "股份终到站港名称",
            ]
        ],
        left_on="钢厂订单号",
        right_on="供应商订单子项号",
        how="left",
        suffixes=("", "_采购"),
    )

    df_inv = df_inv.merge(
        df_summary[["SAP", "后处理方式", "精整分流", "订货规格", "宽度", "客户零件号"]],
        left_on="钢厂订单号",
        right_on="SAP",
        how="left",
        suffixes=("", "_汇总"),
    )

    progress(60, "计算派生字段...")

    inv_spec_parts = df_inv["_规格"].astype(str).str.split("*", n=2, expand=True)
    df_inv["_厚度"] = inv_spec_parts[0]
    df_inv["_宽度"] = inv_spec_parts[1]

    # 订货规格：通过 钢厂订单号=SAP 直接从订单汇总获取
    df_inv["_订货规格"] = df_inv["订货规格"].fillna("").astype(str)

    # 套材判断：钢厂订单号=SAP，比较规格与订单汇总.宽度
    import re

    def _is_taocai(spec_str, summary_width):
        if pd.isna(spec_str):
            return False
        spec_str = str(spec_str).strip()
        parts = spec_str.split("*")
        if len(parts) < 2:
            return False
        nums = re.findall(r"\d+\.?\d*", parts[1])
        if not nums or pd.isna(summary_width):
            return False
        try:
            return float(nums[0]) != float(summary_width)
        except (ValueError, TypeError):
            return False

    def _taocai_for_row(r):
        if str(r.get("存货性质", "")).strip() == "成品":
            return "套材"
        return "套材" if _is_taocai(r["_规格"], r.get("宽度_汇总")) else "非套材"

    df_inv["_套材"] = df_inv.apply(_taocai_for_row, axis=1)

    contract_month = df_inv["采购交货期"].fillna("")

    def _calc_age(row):
        shipped_time = row.get("制造出厂发货时间", "")
        if pd.isna(shipped_time) or str(shipped_time).strip() == "":
            return 0
        try:
            dt = pd.to_datetime(_fmt_date(shipped_time)).date()
            return (datetime.now().date() - dt).days
        except Exception:
            return 0

    inventory_age = df_inv.apply(_calc_age, axis=1)

    progress(70, "生成库存表...")

    df_inv_out = pd.DataFrame(
        {
            "资源号": df_inv["父捆包号材料管理号"],
            "捆包号": df_inv["捆包号"],
            "物料号": df_inv["物料号"],
            "母卷号": df_inv["母捆包号"],
            "规格": df_inv["_规格"],
            "牌号": df_inv["牌号"],
            "入库重量": df_inv["净重(吨)"],
            "入库数量": df_inv["件数"],
            "供应商合同号子项号": df_inv["钢厂订单号"],
            "客户名称": df_inv["客户名称"],
            "仓库名称": df_inv["仓库名称"],
            "捆包状态": df_inv["实物库存状态"],
            "封锁类型": df_inv["封锁类型"],
            "库龄": inventory_age,
            "品种代码": df_inv["品种附属码"],
            "首次入库时间": df_inv["_首次入库时间"],
            "最近入库日期": df_inv["_最近入库日期"],
            "客户零件号": df_inv["客户零件号_汇总"].fillna("").astype(str),
            "订货规格": df_inv["_订货规格"],
            "合同月份": contract_month,
            "厚度/直径": df_inv["_厚度"],
            "宽度": df_inv["_宽度"],
            "长度": "C",
            "订货重量": df_inv["采购订货重量"],
            "表面质量": df_inv["表面质量"],
            "最终用户名称": df_inv["最终用户名称"],
            "终到站港描述": df_inv["股份终到站港名称"],
            "后处理方式": df_inv["后处理方式"],
            "精整分流": df_inv["精整分流"],
            "物资类别": df_inv["存货性质"],
            "套材": df_inv["_套材"],
            "业务类型": df_inv["贸易方式"],
        }
    )

    log(f"  库存表: {df_inv_out.shape[0]} 行, {df_inv_out.shape[1]} 列")

    # ── 5. 输出 Excel ────────────────────────────────────────

    progress(80, "写入 Excel 文件...")
    output_path = os.path.join(work_dir, "output_订单表_库存表.xlsx")
    log(f"\n写入 {output_path}...")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_order_out.to_excel(writer, sheet_name="订单表", index=False)
        df_inv_out.to_excel(writer, sheet_name="库存表", index=False)

    progress(90, "格式化 Excel...")
    from openpyxl import load_workbook

    wb = load_workbook(output_path)

    ws_order = wb["订单表"]
    for row_idx in range(2, ws_order.max_row + 1):
        cell = ws_order.cell(row=row_idx, column=5)
        if cell.value is not None:
            cell.number_format = "@"

    ws_inv = wb["库存表"]
    for row_idx in range(2, ws_inv.max_row + 1):
        for col_idx in (17, 20):
            cell = ws_inv.cell(row=row_idx, column=col_idx)
            if cell.value is not None:
                cell.number_format = "@"

    wb.save(output_path)

    # ── 6. 验证 ──────────────────────────────────────────────

    progress(95, "验证结果...")

    log("\n=== 验证结果 ===")
    log(f"订单表: {df_order_out.shape[0]} 行 (期望: {df_po.shape[0]})")
    log(
        f"库存表: {df_inv_out.shape[0]} 行 (期望: {df_basic.shape[0] + df_shipped.shape[0]})"
    )

    po_match = df_order["客户零件号"].notna().sum()
    log(f"订单表 采购→订单汇总 匹配: {po_match}/{len(df_order)}")

    inv_po_match = df_inv["供应商订单子项号"].notna().sum()
    log(f"库存表 库存→采购 匹配: {inv_po_match}/{len(df_inv)}")

    inv_sum_match = df_inv["后处理方式"].notna().sum()
    log(f"库存表 库存→订单汇总 匹配: {inv_sum_match}/{len(df_inv)}")

    taocai_count = (df_inv_out["套材"] == "套材").sum()
    log(f"库存表 套材数: {taocai_count}")

    progress(100, "完成!")
    log("\n完成!")

    return {
        "order_rows": df_order_out.shape[0],
        "inventory_rows": df_inv_out.shape[0],
        "po_match": f"{po_match}/{len(df_order)}",
        "inv_po_match": f"{inv_po_match}/{len(df_inv)}",
        "inv_sum_match": f"{inv_sum_match}/{len(df_inv)}",
        "taocai_count": taocai_count,
        "output_path": output_path,
    }


FILES_EXPECTED = [
    "系统导出_采购订单数据.xlsx",
    "系统导出_订单汇总.xls",
    "系统导出_业务库存数据.xlsx",
    "系统导出_已发货库存明细.xlsx",
]


def detect_files(work_dir):
    """检测工作目录中是否存在所有必需的源文件。

    Returns:
        dict: 文件名 → 是否存在 (bool)
    """
    return {f: os.path.isfile(os.path.join(work_dir, f)) for f in FILES_EXPECTED}


# ── CLI 入口 ─────────────────────────────────────────────────
if __name__ == "__main__":
    generate()
