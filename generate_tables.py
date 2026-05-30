import pandas as pd
import numpy as np
import os


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

    for col in ["库龄", "原料最初入库日期(年月日)"]:
        if col not in df_shipped.columns:
            df_shipped[col] = np.nan

    if "业务入库日期(年月日)" in df_basic.columns:
        df_basic = df_basic.drop(columns=["业务入库日期(年月日)"])

    common_cols = list(set(df_basic.columns) & set(df_shipped.columns))
    log(f"  共同列: {len(common_cols)}")

    df_union = pd.concat(
        [df_basic[common_cols], df_shipped[common_cols]], ignore_index=True
    )
    log(f"  UNION 结果: {df_union.shape[0]} 行")

    progress(30, "格式化日期列...")
    if df_union["业务入库日期"].dtype in ("float64", "int64"):
        df_union["业务入库日期"] = df_union["业务入库日期"].apply(
            lambda x: str(int(x)) if pd.notna(x) else ""
        )
    if df_union["采购交货期"].dtype in ("float64", "int64"):
        df_union["采购交货期"] = df_union["采购交货期"].apply(
            lambda x: str(int(x)) if pd.notna(x) else ""
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
            "合同月份": df_order["采购交货月"],
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
        df_summary[["产销合同号", "后处理方式", "精整分流"]],
        left_on="钢厂订单号",
        right_on="产销合同号",
        how="left",
        suffixes=("", "_汇总"),
    )

    progress(60, "计算派生字段...")

    inv_spec_parts = df_inv["规格"].astype(str).str.split("*", n=2, expand=True)
    df_inv["_厚度"] = inv_spec_parts[0]
    df_inv["_宽度"] = inv_spec_parts[1]

    df_inv["_订货规格"] = (
        df_inv["牌号"].fillna("").astype(str)
        + df_inv["规格描述"].fillna("").astype(str)
        + df_inv["表面质量"].fillna("").astype(str)
    )

    def _normalize_spec(s):
        if pd.isna(s):
            return ""
        parts = []
        for seg in str(s).split("*"):
            seg = seg.strip()
            try:
                v = float(seg)
                normalized = str(v)
                if normalized.startswith("."):
                    normalized = "0" + normalized
            except ValueError:
                normalized = seg
            parts.append(normalized)
        return "*".join(parts)

    df_inv["_套材"] = df_inv.apply(
        lambda r: (
            _normalize_spec(r.get("规格描述"))
            != _normalize_spec(r["规格"])
        )
        if pd.notna(r.get("规格描述"))
        else False,
        axis=1,
    )

    contract_month = df_inv["采购交货期"].fillna("")

    progress(70, "生成库存表...")

    df_inv_out = pd.DataFrame(
        {
            "资源号": df_inv["父捆包号材料管理号"],
            "捆包号": df_inv["捆包号"],
            "物料号": df_inv["物料号"],
            "母卷号": df_inv["母捆包号"],
            "规格": df_inv["规格"],
            "牌号": df_inv["牌号"],
            "入库重量": df_inv["净重(吨)"],
            "入库数量": df_inv["件数"],
            "供应商合同号子项号": df_inv["钢厂订单号"],
            "客户名称": df_inv["客户名称"],
            "仓库名称": df_inv["仓库名称"],
            "捆包状态": df_inv["实物库存状态"],
            "封锁类型": df_inv["封锁类型"],
            "库龄": df_inv["库龄"],
            "品种代码": df_inv["品种附属码"],
            "首次入库时间": df_inv["原料最初入库日期(年月日)"],
            "最近入库日期": df_inv["业务入库日期"],
            "客户零件号": df_inv["客户零件号"],
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
            cell.value = str(int(cell.value))
            cell.number_format = "@"

    ws_inv = wb["库存表"]
    for row_idx in range(2, ws_inv.max_row + 1):
        for col_idx in (17, 20):
            cell = ws_inv.cell(row=row_idx, column=col_idx)
            if cell.value is not None:
                cell.value = str(int(cell.value))
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

    taocai_count = df_inv_out["套材"].sum()
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
