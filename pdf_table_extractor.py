"""
PDF 表格提取脚本（前 4 页）

功能说明：
1) 只处理 PDF 前 4 页，不处理第 5 页及之后页面。
2) 页 1-2 作为第一张表，页 3-4 作为第二张表。
3) 首选 pdfplumber + pandas；失败时回退到 camelot；再失败则按行文本正则拆分。
4) 清洗数据：删除重复表头、删除无效说明行、去空行、数值列转数值、保留样本编号列。
5) 导出到同一个 Excel：sheet table_1 / table_2。

依赖：
- 必需: pip install pdfplumber pandas openpyxl
- 可选: pip install camelot-py[base]
"""

from __future__ import annotations

import argparse
import importlib
import os
import re
from dataclasses import dataclass
from typing import Iterable, List, Optional

import pandas as pd
import pdfplumber

# ============================================================
# 用户配置（请按需修改）
# ============================================================
PDF_INPUT_PATH = r"D:\A_Projects_MQ\20260410-python-PDF_excel\150 2.pdf"
EXCEL_OUTPUT_PATH = r"D:\A_Projects_MQ\20260410-python-PDF_excel\output.xlsx"

TABLE1_PAGE_RANGE = (1, 2)  # 页 1-2 -> table_1
TABLE2_PAGE_RANGE = (3, 4)  # 页 3-4 -> table_2
# ============================================================

SAMPLE_ID_PATTERN = re.compile(r"^\d{1,3}-\d{1,3}$")
NUM_PATTERN = re.compile(r"[-+]?\d+(?:\.\d+)?")


@dataclass
class ExtractConfig:
    page_range: tuple[int, int]
    table_name: str


def normalize_cell(value: object) -> str:
    """标准化单元格文本。"""
    if value is None:
        return ""
    text = str(value)
    text = text.replace("\u3000", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def is_empty_row(row: Iterable[object]) -> bool:
    normalized = [normalize_cell(v) for v in row]
    return all(v == "" for v in normalized)


def pad_or_trim_row(row: List[str], width: int) -> List[str]:
    if len(row) == width:
        return row
    if len(row) < width:
        return row + [""] * (width - len(row))
    return row[:width]


def looks_like_header(row: List[str]) -> bool:
    """
    粗略判断是否是表头：非空文本比例较高且非纯数字字段较多。
    """
    if not row:
        return False

    non_empty = [c for c in row if c != ""]
    if not non_empty:
        return False

    text_like = 0
    for c in non_empty:
        if not re.fullmatch(r"[-+]?\d+(?:\.\d+)?", c):
            text_like += 1

    return text_like >= max(2, int(len(non_empty) * 0.4))


def row_similarity_ratio(a: List[str], b: List[str]) -> float:
    if not a or not b or len(a) != len(b):
        return 0.0
    same = sum(1 for x, y in zip(a, b) if x == y and x != "")
    return same / max(1, len(a))


def remove_duplicate_headers(df: pd.DataFrame) -> pd.DataFrame:
    """
    删除重复表头行：
    - 与当前列名高度一致的行
    - 与首行高度一致的重复标题行
    """
    if df.empty:
        return df

    current_header = [normalize_cell(c) for c in df.columns]

    first_row = [normalize_cell(v) for v in df.iloc[0].tolist()] if len(df) > 0 else []
    duplicate_indexes = []

    for idx, row in df.iterrows():
        row_list = [normalize_cell(v) for v in row.tolist()]
        sim_with_cols = row_similarity_ratio(row_list, current_header)
        sim_with_first = row_similarity_ratio(row_list, first_row) if first_row else 0.0

        # 高相似度判定为重复表头
        if sim_with_cols >= 0.75 or (idx > 0 and sim_with_first >= 0.9):
            duplicate_indexes.append(idx)

    if duplicate_indexes:
        df = df.drop(index=duplicate_indexes)
    return df.reset_index(drop=True)


def remove_invalid_rows(df: pd.DataFrame) -> pd.DataFrame:
    """
    删除无效说明行，如 "--- --- 450.00 450.00" 这类。
    规则：
    - 短横线字段占比高
    - 有效文本极少且主要为分隔符
    """
    if df.empty:
        return df

    keep_mask: List[bool] = []

    for _, row in df.iterrows():
        cells = [normalize_cell(v) for v in row.tolist()]

        if all(c == "" for c in cells):
            keep_mask.append(False)
            continue

        dash_like = sum(1 for c in cells if re.fullmatch(r"[-_]{2,}", c) is not None)
        non_empty = [c for c in cells if c != ""]

        # 像 "--- --- 450.00 450.00" 这样的混合说明行，通常前半是分隔符。
        leading_dash = 0
        for c in cells:
            if re.fullmatch(r"[-_]{2,}", c):
                leading_dash += 1
            elif c == "":
                continue
            else:
                break

        drop = False
        if non_empty and dash_like / len(non_empty) >= 0.6:
            drop = True
        if leading_dash >= 2 and len(non_empty) <= max(4, len(cells) // 2 + 1):
            drop = True

        keep_mask.append(not drop)

    return df.loc[keep_mask].reset_index(drop=True)


def drop_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    non_empty_mask = []
    for _, row in df.iterrows():
        non_empty_mask.append(not is_empty_row(row.tolist()))
    return df.loc[non_empty_mask].reset_index(drop=True)


def detect_id_columns(columns: Iterable[object]) -> set[str]:
    """识别样本编号列，保留为文本。"""
    id_pattern = re.compile(r"(样本|编号|sample|id|no\.?|序号)", re.IGNORECASE)
    result = set()
    for col in columns:
        col_name = normalize_cell(col)
        if id_pattern.search(col_name):
            result.add(col_name)
    return result


def to_numeric_if_possible(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.strip()
        .replace({"": pd.NA, "None": pd.NA, "nan": pd.NA})
    )
    converted = pd.to_numeric(cleaned, errors="coerce")

    non_null_src = cleaned.notna().sum()
    if non_null_src == 0:
        return series

    success_ratio = converted.notna().sum() / non_null_src
    return converted if success_ratio >= 0.6 else series


def clean_dataframe(df: pd.DataFrame, table_name: str) -> pd.DataFrame:
    """统一清洗流程。"""
    if df.empty:
        return df

    before_rows = len(df)

    # 规范列名
    df = df.copy()
    df.columns = [normalize_cell(c) if normalize_cell(c) else f"col_{i+1}" for i, c in enumerate(df.columns)]

    df = remove_duplicate_headers(df)
    df = remove_invalid_rows(df)
    df = drop_empty_rows(df)

    # 去除全空列
    df = df.loc[:, ~(df.apply(lambda col: col.astype(str).str.strip().isin(["", "nan", "None"]).all()))]

    # 数值转换（排除样本编号列）
    id_cols = detect_id_columns(df.columns)
    for col in df.columns:
        if normalize_cell(col) in id_cols:
            continue
        df[col] = to_numeric_if_possible(df[col])

    after_rows = len(df)
    print(f"[{table_name}] 清洗完成: {before_rows} -> {after_rows} 行")
    return df


def build_dataframe_from_rows(rows: List[List[str]]) -> pd.DataFrame:
    """将不规则二维列表构建为 DataFrame，并尽量识别表头。"""
    if not rows:
        return pd.DataFrame()

    max_cols = max(len(r) for r in rows)
    norm_rows = [pad_or_trim_row([normalize_cell(c) for c in r], max_cols) for r in rows]

    # 识别首个可用表头
    header_idx = None
    for i, row in enumerate(norm_rows[:6]):
        if looks_like_header(row):
            header_idx = i
            break

    if header_idx is None:
        columns = [f"col_{i+1}" for i in range(max_cols)]
        data_rows = norm_rows
    else:
        raw_cols = norm_rows[header_idx]
        columns = []
        used = {}
        for i, c in enumerate(raw_cols):
            name = c if c else f"col_{i+1}"
            if name in used:
                used[name] += 1
                name = f"{name}_{used[name]}"
            else:
                used[name] = 1
            columns.append(name)
        data_rows = norm_rows[header_idx + 1 :]

    if not data_rows:
        return pd.DataFrame(columns=columns)

    return pd.DataFrame(data_rows, columns=columns)


def extract_text_rows_with_regex(text: str) -> List[List[str]]:
    """
    按行提取并用正则拆字段的兜底方案。
    拆分策略：
    - 优先按 2 个及以上空白切分
    - 切分不足时，提取数值并保留前导文本
    """
    rows: List[List[str]] = []

    for line in text.splitlines():
        s = re.sub(r"\s+", " ", line).strip()
        if not s:
            continue

        # 跳过明显页眉页脚
        if re.search(r"^(第\s*\d+\s*页|page\s*\d+)", s, flags=re.IGNORECASE):
            continue

        parts = re.split(r"\s{2,}", line.strip())
        parts = [normalize_cell(p) for p in parts if normalize_cell(p)]

        if len(parts) >= 2:
            rows.append(parts)
            continue

        # 正则兜底：提取文本前缀 + 全部数值
        nums = re.findall(r"[-+]?\d+(?:\.\d+)?", s)
        text_prefix = re.sub(r"[-+]?\d+(?:\.\d+)?", " ", s)
        text_prefix = re.sub(r"\s+", " ", text_prefix).strip()

        merged: List[str] = []
        if text_prefix:
            merged.append(text_prefix)
        merged.extend(nums)

        if len(merged) >= 2:
            rows.append(merged)

    return rows


def extract_with_pdfplumber(pdf_path: str, page_range: tuple[int, int]) -> Optional[pd.DataFrame]:
    """首选方案：pdfplumber 提取表格。"""
    start, end = page_range
    all_rows: List[List[str]] = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            if start < 1 or end > total_pages:
                print(f"[pdfplumber] 页码越界: {page_range}, PDF 总页数={total_pages}")
                return None

            for page_no in range(start, end + 1):
                page = pdf.pages[page_no - 1]
                tables = page.extract_tables(
                    table_settings={
                        "vertical_strategy": "lines",
                        "horizontal_strategy": "lines",
                        "intersection_tolerance": 8,
                        "snap_tolerance": 5,
                        "join_tolerance": 5,
                    }
                )

                got_table = False
                if tables:
                    for tb in tables:
                        if not tb:
                            continue
                        got_table = True
                        for row in tb:
                            row_norm = [normalize_cell(c) for c in (row or [])]
                            if row_norm and not all(v == "" for v in row_norm):
                                all_rows.append(row_norm)

                if not got_table:
                    # 该页没识别到结构化表格，立即触发按行文本兜底（页级）
                    text = page.extract_text() or ""
                    all_rows.extend(extract_text_rows_with_regex(text))

        if not all_rows:
            return None

        df = build_dataframe_from_rows(all_rows)
        print(f"[pdfplumber] 提取到 {len(df)} 行, {len(df.columns)} 列")
        return df
    except Exception as exc:
        print(f"[pdfplumber] 提取失败: {exc}")
        return None


def extract_with_camelot(pdf_path: str, page_range: tuple[int, int]) -> Optional[pd.DataFrame]:
    """备选方案：camelot（动态导入，避免环境未安装时报错）。"""
    try:
        camelot = importlib.import_module("camelot")
    except Exception:
        print("[camelot] 未安装或不可用，跳过")
        return None

    start, end = page_range
    pages = f"{start}-{end}"

    try:
        tables = camelot.read_pdf(pdf_path, pages=pages, flavor="stream")
        if not tables or len(tables) == 0:
            return None

        chunks: List[pd.DataFrame] = []
        for t in tables:
            cdf = t.df.copy()
            cdf = cdf.replace("\n", " ", regex=True)
            chunks.append(cdf)

        df = pd.concat(chunks, axis=0, ignore_index=True)
        # 将 Camelot 默认列名改成简单列名，后续统一清洗
        df.columns = [f"col_{i+1}" for i in range(len(df.columns))]
        print(f"[camelot] 提取到 {len(df)} 行, {len(df.columns)} 列")
        return df
    except Exception as exc:
        print(f"[camelot] 提取失败: {exc}")
        return None


def extract_with_regex_fallback(pdf_path: str, page_range: tuple[int, int]) -> Optional[pd.DataFrame]:
    """终极兜底：整页文本按行 + 正则拆分字段。"""
    start, end = page_range
    all_rows: List[List[str]] = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_no in range(start, end + 1):
                text = pdf.pages[page_no - 1].extract_text() or ""
                all_rows.extend(extract_text_rows_with_regex(text))

        if not all_rows:
            return None

        df = build_dataframe_from_rows(all_rows)
        print(f"[regex] 提取到 {len(df)} 行, {len(df.columns)} 列")
        return df
    except Exception as exc:
        print(f"[regex] 提取失败: {exc}")
        return None


def collect_row_blobs(df: pd.DataFrame) -> List[str]:
    """把每行非空单元格拼接成文本，用于二次解析。"""
    blobs: List[str] = []
    for _, row in df.iterrows():
        parts = [normalize_cell(x) for x in row.tolist()]
        parts = [p for p in parts if p]
        if parts:
            blobs.append(" ".join(parts))
    return blobs


def is_number_token(token: str) -> bool:
    return NUM_PATTERN.fullmatch(token) is not None


def parse_table1_blob(blob: str) -> List[tuple[str, str, str, str, str]]:
    """
    解析 table_1 长文本行。
    支持两种模式：
    1) 交错模式: 样本编号 + 3个数值
    2) 块模式: 样本编号批量 + 3段数值批量
    """
    tokens = blob.split(" ")
    records: List[tuple[str, str, str, str, str]] = []

    bezeichnung = ""
    first_id_idx = None
    for i, tk in enumerate(tokens):
        if SAMPLE_ID_PATTERN.fullmatch(tk):
            first_id_idx = i
            break

    if first_id_idx is not None and first_id_idx > 0:
        for tk in tokens[:first_id_idx]:
            if tk and not is_number_token(tk) and tk not in {"---", "Bezeichnung", "Probe", "N/mm", "F_min", "min", "max", "标注"}:
                bezeichnung = tk
                break

    # 交错模式
    i = 0
    n = len(tokens)
    while i < n:
        tk = tokens[i]
        if SAMPLE_ID_PATTERN.fullmatch(tk):
            sid = tk
            nums: List[str] = []
            j = i + 1
            while j < n and len(nums) < 3:
                t2 = tokens[j]
                if SAMPLE_ID_PATTERN.fullmatch(t2):
                    break
                if is_number_token(t2):
                    nums.append(t2)
                j += 1
            if len(nums) == 3:
                records.append((bezeichnung, sid, nums[0], nums[1], nums[2]))
                i = j
                continue
        i += 1

    # 块模式
    if len(records) < 5:
        id_tokens = [t for t in tokens if SAMPLE_ID_PATTERN.fullmatch(t)]
        if id_tokens:
            seen = set()
            ids: List[str] = []
            for sid in id_tokens:
                if sid not in seen:
                    ids.append(sid)
                    seen.add(sid)

            num_tokens = [t for t in tokens if is_number_token(t)]
            n_ids = len(ids)
            needed = 3 * n_ids
            if len(num_tokens) >= needed:
                num_tokens = num_tokens[:needed]
                col1 = num_tokens[0:n_ids]
                col2 = num_tokens[n_ids : 2 * n_ids]
                col3 = num_tokens[2 * n_ids : 3 * n_ids]
                block_records = [
                    (bezeichnung, sid, v1, v2, v3)
                    for sid, v1, v2, v3 in zip(ids, col1, col2, col3)
                ]
                if len(block_records) > len(records):
                    records = block_records

    return records


def sort_sample_id_series(series: pd.Series) -> pd.Series:
    def sort_key(s: str) -> tuple[int, int]:
        left, right = str(s).split("-")
        return int(left), int(right)

    return series.map(sort_key)


def refine_table1(df: pd.DataFrame) -> pd.DataFrame:
    """将 table_1 二次结构化，拆分挤压在单元格内的记录。"""
    if df.empty:
        return df

    blobs = collect_row_blobs(df)
    rows: List[tuple[str, str, str, str, str]] = []
    for blob in blobs:
        rows.extend(parse_table1_blob(blob))

    if not rows:
        return df

    out = pd.DataFrame(
        rows,
        columns=[
            "Bezeichnung",
            "Probe",
            "Minimale Steifigkeit auf N/mm",
            "Minimale Steifigkeit ab N/mm",
            "F_min N",
        ],
    )

    out = out[out["Probe"].astype(str).str.match(SAMPLE_ID_PATTERN)]
    out = out.drop_duplicates(subset=["Probe"], keep="first")

    out["Minimale Steifigkeit auf N/mm"] = pd.to_numeric(out["Minimale Steifigkeit auf N/mm"], errors="coerce")
    out["Minimale Steifigkeit ab N/mm"] = pd.to_numeric(out["Minimale Steifigkeit ab N/mm"], errors="coerce")
    out["F_min N"] = pd.to_numeric(out["F_min N"], errors="coerce")

    out = out.sort_values(by="Probe", key=sort_sample_id_series).reset_index(drop=True)
    return out


def refine_table2(df: pd.DataFrame) -> pd.DataFrame:
    """将 table_2 二次结构化，把长文本拆成 3 列数值。"""
    if df.empty:
        return df

    blobs = collect_row_blobs(df)
    def build_output(rows: List[tuple[str, str, str, str]]) -> pd.DataFrame:
        out_local = pd.DataFrame(
            rows,
            columns=["样本编号", "Mitl_Steigung_Taster N/mm", "Dehn. Fmax Taster mm", "F_max N"],
        )
        out_local["Mitl_Steigung_Taster N/mm"] = pd.to_numeric(out_local["Mitl_Steigung_Taster N/mm"], errors="coerce")
        out_local["Dehn. Fmax Taster mm"] = pd.to_numeric(out_local["Dehn. Fmax Taster mm"], errors="coerce")
        out_local["F_max N"] = pd.to_numeric(out_local["F_max N"], errors="coerce")
        out_local = out_local.dropna(
            how="all", subset=["Mitl_Steigung_Taster N/mm", "Dehn. Fmax Taster mm", "F_max N"]
        )
        return out_local.reset_index(drop=True)

    def score_candidate(out_local: pd.DataFrame) -> int:
        if out_local.empty:
            return -1
        c1 = out_local["Mitl_Steigung_Taster N/mm"]
        c2 = out_local["Dehn. Fmax Taster mm"]
        c3 = out_local["F_max N"]

        score = 0
        if c1.notna().mean() >= 0.9:
            score += 1
        if c2.notna().mean() >= 0.9:
            score += 1
        if c3.notna().mean() >= 0.9:
            score += 1

        # table_2 的典型值域：Steigung 较大、Dehn 较小、F_max 约数百
        if c1.median(skipna=True) > 1000:
            score += 2
        if 0 < c2.median(skipna=True) < 10:
            score += 3
        if c2.median(skipna=True) >= 100:
            score -= 6
        if c3.median(skipna=True) > 100:
            score += 2

        return score

    candidates: List[tuple[int, int, pd.DataFrame]] = []

    # 关键修复：每个 blob 独立解析，避免跨 blob 拼接造成列错位。
    for b in blobs:
        nums = NUM_PATTERN.findall(b)
        if len(nums) < 9:
            continue

        # 模式1：交错模式（每 3 个数值为一条记录）
        inter_rows: List[tuple[str, str, str, str]] = []
        usable = (len(nums) // 3) * 3
        if usable >= 9:
            for i in range(0, usable, 3):
                inter_rows.append((str(i // 3 + 1), nums[i], nums[i + 1], nums[i + 2]))
            out_inter = build_output(inter_rows)
            candidates.append((score_candidate(out_inter), len(out_inter), out_inter))

        # 模式2：块模式（前1/3, 中1/3, 后1/3 分别为三列）
        if usable >= 9:
            n = usable // 3
            seq1 = nums[0:n]
            seq2 = nums[n : 2 * n]
            seq3 = nums[2 * n : 3 * n]
            block_rows = [(str(i + 1), seq1[i], seq2[i], seq3[i]) for i in range(n)]
            out_block = build_output(block_rows)
            candidates.append((score_candidate(out_block), len(out_block), out_block))

    if not candidates:
        return df

    # 先过滤低质量候选，再合并去重，保留所有有效段。
    def looks_like_table2(out_local: pd.DataFrame) -> bool:
        if out_local.empty or len(out_local) < 5:
            return False
        c1 = out_local["Mitl_Steigung_Taster N/mm"]
        c2 = out_local["Dehn. Fmax Taster mm"]
        c3 = out_local["F_max N"]

        c1_med = c1.median(skipna=True)
        c2_med = c2.median(skipna=True)
        c3_med = c3.median(skipna=True)
        c2_q90 = c2.quantile(0.9)

        return (
            c1_med > 100
            and 0 < c2_med < 10
            and c2_q90 < 10
            and c3_med > 100
        )

    good = [c for c in candidates if c[0] >= 6 and looks_like_table2(c[2])]
    if not good:
        good = sorted(candidates, key=lambda x: (x[0], x[1]), reverse=True)[:1]

    merged = pd.concat([c[2] for c in good], axis=0, ignore_index=True)
    merged = merged.drop_duplicates(
        subset=["Mitl_Steigung_Taster N/mm", "Dehn. Fmax Taster mm", "F_max N"], keep="first"
    ).reset_index(drop=True)

    merged["样本编号"] = [str(i + 1) for i in range(len(merged))]
    merged = merged[["样本编号", "Mitl_Steigung_Taster N/mm", "Dehn. Fmax Taster mm", "F_max N"]]
    return merged


def finalize_tables(table1: pd.DataFrame, table2: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """把初提取结果做二次修复，确保输出最终可用 Excel。"""
    fixed1 = refine_table1(table1)
    fixed2 = refine_table2(table2)
    print(f"[finalize] table_1: {len(table1)} -> {len(fixed1)} 行")
    print(f"[finalize] table_2: {len(table2)} -> {len(fixed2)} 行")
    return fixed1, fixed2


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="提取 PDF 前4页两张表并导出最终 Excel")
    parser.add_argument("--pdf", default=PDF_INPUT_PATH, help="输入 PDF 路径")
    parser.add_argument("--out", default=EXCEL_OUTPUT_PATH, help="输出 Excel 路径")
    return parser.parse_args()


def extract_single_table(pdf_path: str, config: ExtractConfig) -> pd.DataFrame:
    """按优先级提取并清洗单张表。"""
    print("\n" + "=" * 64)
    print(f"开始提取 {config.table_name}，页码: {config.page_range[0]}-{config.page_range[1]}")
    print("=" * 64)

    candidates = [
        ("pdfplumber", extract_with_pdfplumber),
        ("camelot", extract_with_camelot),
        ("regex", extract_with_regex_fallback),
    ]

    for name, fn in candidates:
        print(f"尝试方案: {name}")
        df = fn(pdf_path, config.page_range)
        if df is not None and not df.empty:
            df = clean_dataframe(df, config.table_name)
            if not df.empty:
                print(f"{config.table_name} 提取成功，方案={name}")
                return df

    print(f"{config.table_name} 提取失败，返回空表")
    return pd.DataFrame()


def validate_page_ranges(pdf_path: str, ranges: list[tuple[int, int]]) -> None:
    """校验页码范围，确保只处理前 4 页。"""
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)

    for start, end in ranges:
        if start < 1 or end < start:
            raise ValueError(f"无效页码范围: {start}-{end}")
        if end > total:
            raise ValueError(f"页码范围 {start}-{end} 超出 PDF 总页数 {total}")
        if end > 4:
            raise ValueError(f"禁止处理第 5 页及之后页面，当前范围: {start}-{end}")


def export_to_excel(path: str, table1: pd.DataFrame, table2: pd.DataFrame) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        table1.to_excel(writer, sheet_name="table_1", index=False)
        table2.to_excel(writer, sheet_name="table_2", index=False)


def convert_pdf_to_excel(pdf_path: str, excel_output_path: str) -> dict:
    """对外复用入口：输入 PDF，输出最终清洗后的 Excel。"""
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF 文件不存在: {pdf_path}")

    out_dir = os.path.dirname(excel_output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    validate_page_ranges(pdf_path, [TABLE1_PAGE_RANGE, TABLE2_PAGE_RANGE])

    cfg1 = ExtractConfig(page_range=TABLE1_PAGE_RANGE, table_name="table_1")
    cfg2 = ExtractConfig(page_range=TABLE2_PAGE_RANGE, table_name="table_2")

    df1 = extract_single_table(pdf_path, cfg1)
    df2 = extract_single_table(pdf_path, cfg2)
    df1, df2 = finalize_tables(df1, df2)

    export_to_excel(excel_output_path, df1, df2)

    return {
        "output": excel_output_path,
        "table_1_rows": len(df1),
        "table_1_cols": len(df1.columns),
        "table_2_rows": len(df2),
        "table_2_cols": len(df2.columns),
    }


def main() -> None:
    args = parse_args()
    pdf_path = args.pdf
    excel_output_path = args.out

    result = convert_pdf_to_excel(pdf_path, excel_output_path)

    print("\n导出完成")
    print(f"输出文件: {result['output']}")
    print(f"sheet table_1: {result['table_1_rows']} 行 x {result['table_1_cols']} 列")
    print(f"sheet table_2: {result['table_2_rows']} 行 x {result['table_2_cols']} 列")


if __name__ == "__main__":
    main()
