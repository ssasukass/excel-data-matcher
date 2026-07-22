import sys, os
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
_lib = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_lib')
if os.path.isdir(_lib) and _lib not in sys.path: sys.path.insert(0, _lib)

"""
Excel 万能匹配填充工具
按指定列作为索引，从源表匹配数据填充到目标表。

用法示例：
  # 基本：按姓名匹配，填入学号和班级
  python match_excel.py -s 源表.xlsx -t 目标.xlsx -k 姓名 -f 学号 -f 班级

  # 多列索引 + 多列填充
  python match_excel.py -s 源.xlsx -t 目标.xlsx -k 姓名 -k 专业 -f 学号 -f 班级 -f 学院

  # 列名不同：源表英文名，目标表中文名
  python match_excel.py -s src.xlsx -t tgt.xlsx -k name=姓名 -f student_id=学号 -f class_name=班级

  # 模糊匹配（适用于单列索引）
  python match_excel.py -s 源.xlsx -t 目标.xlsx -k 姓名 -f 学号 --fuzzy --threshold 0.85

  # 完整列名映射
  python match_excel.py -s src.xlsx -t tgt.xlsx -k ID=学号 -k Name=姓名 -f Class=班级
"""

import argparse, sys, shutil
from pathlib import Path
from difflib import SequenceMatcher
from collections import Counter

# ── 解析参数 ──────────────────────────────────────────────
def parse_args():
    p = argparse.ArgumentParser(
        description='按指定索引列从源表匹配数据并填充到目标表',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    p.add_argument('-s', '--source', required=True, help='源表路径（含完整数据）')
    p.add_argument('-t', '--target', required=True, help='目标表路径（待填充）')
    p.add_argument('-o', '--output', default=None, help='输出文件路径（默认：已填充_文件名.xlsx）')
    p.add_argument('-k', '--key', action='append', default=[],
                   help='索引列，可多次指定。格式：COL 或 SRC_COL:TGT_COL（列名不同时）')
    p.add_argument('-f', '--fill', action='append', default=[],
                   help='需填充的列，可多次指定。格式：COL 或 SRC_COL:TGT_COL')
    p.add_argument('--source-sheet', default=None, help='源表 sheet 名称（默认自动检测）')
    p.add_argument('--target-sheet', default=None, help='目标表 sheet 名称（默认自动检测）')
    p.add_argument('--fuzzy', action='store_true', help='启用模糊匹配（单列索引时有效）')
    p.add_argument('--threshold', type=float, default=0.85, help='模糊匹配相似度阈值（默认 0.85）')
    p.add_argument('--source-header-row', type=int, default=0, help='源表表头行号（0起始，默认0即第一行）')
    p.add_argument('--target-header-row', type=int, default=0, help='目标表表头行号（0起始，默认0即第一行）')
    p.add_argument('--no-backup', action='store_true', help='不创建备份文件')
    p.add_argument('--append-missing', action='store_true', help='将源表有但目标表没有的记录追加到输出')
    return p.parse_args()


def parse_col_pair(s):
    """解析列名映射：'COL' -> (col, col)  'SRC:TGT' -> (src, tgt)"""
    if ':' in s and s.count(':') == 1:
        src, tgt = s.split(':', 1)
        return src.strip(), tgt.strip()
    if '：' in s and s.count('：') == 1:
        src, tgt = s.split('：', 1)
        return src.strip(), tgt.strip()
    return s.strip(), s.strip()


def detect_sheet(fp):
    import pandas as pd
    xls = pd.ExcelFile(fp)
    for n in xls.sheet_names:
        if not pd.read_excel(xls, sheet_name=n, nrows=1).empty:
            return n
    return xls.sheet_names[0]


def find_best_match(values, source_map, threshold=0.85):
    """在源表索引中找最佳模糊匹配"""
    is_comp = chr(9) in values
    best_s, best_m = 0, None
    for skey, m in source_map.items():
        if is_comp:
            tp = values.split(chr(9))
            sp = skey.split(chr(9))
            if len(tp) != len(sp): continue
            scores = [SequenceMatcher(None, tv, sv).ratio() for tv, sv in zip(tp, sp)]
            s = sum(scores) / len(scores)
        else:
            s = SequenceMatcher(None, values, skey).ratio()
        if s > best_s:
            best_s, best_m = s, skey
    return (best_m, best_s) if best_s >= threshold else (None, 0)

def main():
    a = parse_args()

    # ── 解析索引列和填充列 ──
    key_pairs = [parse_col_pair(k) for k in a.key]
    fill_pairs = [parse_col_pair(f) for f in a.fill]

    if not key_pairs:
        print('[错误] 请至少指定一个索引列，如 -k 姓名')
        sys.exit(1)
    if not fill_pairs:
        print('[错误] 请至少指定一个填充列，如 -f 学号')
        sys.exit(1)

    if a.fuzzy and len(key_pairs) > 1:
        print('[信息] 多列索引模糊匹配：各字段分别计算相似度取平均')

    # ── 文件检查 ──
    sp, tp = Path(a.source), Path(a.target)
    if not sp.exists(): print(f'[错误] 源文件不存在: {sp}'); sys.exit(1)
    if not tp.exists(): print(f'[错误] 目标文件不存在: {tp}'); sys.exit(1)

    # ── 读取数据 ──
    import pandas as pd
    ss = a.source_sheet or detect_sheet(sp)
    ts = a.target_sheet or detect_sheet(tp)
    if not ss: print('[错误] 无法检测源表 sheet'); sys.exit(1)
    if not ts: print('[错误] 无法检测目标表 sheet'); sys.exit(1)

    print(f'源表: {sp} [sheet: {ss}]')
    print(f'目标表: {tp} [sheet: {ts}]')
    print(f'索引列: {[p[0] for p in key_pairs]}', end='')
    if any(s != t for s, t in key_pairs):
        print(f' 映射: {key_pairs}', end='')
    print()
    print(f'填充列: {[p[0] for p in fill_pairs]}', end='')
    if any(s != t for s, t in fill_pairs):
        print(f' 映射: {fill_pairs}', end='')
    print('\n' + '-' * 50)

    ds = pd.read_excel(sp, sheet_name=ss, header=a.source_header_row, dtype=str).dropna(how='all').reset_index(drop=True)
    dt = pd.read_excel(tp, sheet_name=ts, header=a.target_header_row, dtype=str).dropna(how='all').reset_index(drop=True)
    ds.columns = ds.columns.str.strip()
    dt.columns = dt.columns.str.strip()

    print(f'源表: {len(ds)} 行  目标表: {len(dt)} 行')
    print('-' * 50)

    # ── 列存在性检查 ──
    src_key_cols = [p[0] for p in key_pairs]
    tgt_key_cols = [p[1] for p in key_pairs]
    src_fill_cols = [p[0] for p in fill_pairs]
    tgt_fill_cols = [p[1] for p in fill_pairs]

    missing = []
    for c in src_key_cols + src_fill_cols:
        if c not in ds.columns: missing.append(f'源表缺少: {c}')
    for c in tgt_key_cols:
        if c not in dt.columns: missing.append(f'目标表缺少: {c}')
    if missing:
        for m in missing: print(f'  [错误] {m}')
        print(f'  源表列: {list(ds.columns)}')
        print(f'  目标表列: {list(dt.columns)}')
        sys.exit(1)

    # 填充列在目标表中不存在时，自动创建
    for c in tgt_fill_cols:
        if c not in dt.columns:
            dt[c] = ''
            print(f'  [提示] 目标表无 "{c}" 列，已自动创建')

    # ── 数据清洗 ──
    for c in src_key_cols:
        ds[c] = ds[c].astype(str).str.strip()
    for c in tgt_key_cols:
        dt[c] = dt[c].astype(str).str.strip()

    # ── 构建源表索引 ──
    source_map = {}
    dups = []
    for _, r in ds.iterrows():
        # 构造复合键
        key_parts = []
        valid = True
        for c in src_key_cols:
            v = str(r[c]).strip()
            if v.lower() in ('', 'nan', 'none'):
                valid = False
                break
            key_parts.append(v)
        if not valid:
            continue
        key = '\t'.join(key_parts)  # 用制表符连接多键
        
        if key in source_map:
            dups.append(key)
        # 存储填充列的值
        vals = {}
        for sf, tf in zip(src_fill_cols, tgt_fill_cols):
            vals[tf] = r.get(sf, '')
        vals["_raw"] = r.to_dict()
        source_map[key] = vals

    if dups:
        dc = Counter(dups)
        print(f'[警告] 源表 {len(dc)} 个重复键值（将使用最后一条记录）:')
        for k, c in list(dc.items())[:5]:
            print(f'  "{k.replace(chr(9), " / ")}" x{c+1}')
        if len(dc) > 5: print(f'  ... 共 {len(dc)} 个')


    # ── 按姓名建立索引（优化模糊匹配速度） ──
    from collections import defaultdict
    name_index = defaultdict(list)
    for key in source_map:
        parts = key.split(chr(9))
        name_index[parts[0]].append(key)

    # ── 备份 ──
    if not a.no_backup:
        bp = tp.parent / f'备份_{tp.stem}{tp.suffix}'
        try:
            shutil.copy2(tp, bp)
            print(f'[备份] {bp}')
        except Exception as e:
            print(f'[警告] 备份失败: {e}')

    # ── 执行匹配 ──
    matched = 0
    matched_source_keys = set()
    unmatched = []
    fuzzy_matched = []

    for idx, row in dt.iterrows():
        key_parts = []
        valid = True
        for c in tgt_key_cols:
            v = str(row[c]).strip()
            if v.lower() in ('', 'nan', 'none'):
                valid = False
                break
            key_parts.append(v)
        if not valid:
            continue

        key = '\t'.join(key_parts)

        # 精确匹配
        if key in source_map:
            for tf in tgt_fill_cols:
                dt.at[idx, tf] = source_map[key].get(tf, '')
            matched_source_keys.add(key)
            matched += 1
            continue

        # 模糊匹配
        if a.fuzzy:
            # 复合键优先按姓名过滤（大幅提速）
            if chr(9) in key and len(key_parts) > 1:
                candidates = {k: source_map[k] for k in name_index.get(key_parts[0], [])}
                bm, sc = find_best_match(key, candidates, a.threshold) if candidates else (None, 0)
                # ↑ 按姓名精确过滤后只比院系字段，不再回退全表扫描（百倍提速）
            else:
                bm, sc = find_best_match(key, source_map, a.threshold)
            if bm:
                for tf in tgt_fill_cols:
                    dt.at[idx, tf] = source_map[bm].get(tf, '')
                matched_source_keys.add(bm)
                fuzzy_matched.append((key_parts[0], bm.split(chr(9))[0], round(sc, 4)))
                matched += 1
                continue
        unmatched.append(key_parts)

    # ── 追加缺失记录 ──
    if a.append_missing:
        new_rows = []
        for src_key, vals in source_map.items():
            if src_key not in matched_source_keys:
                row_data = {}
                src_parts = src_key.split(chr(9))
                for si, (sc, tc) in enumerate(zip(src_key_cols, tgt_key_cols)):
                    if si < len(src_parts):
                        row_data[tc] = src_parts[si]
                for tc in tgt_fill_cols:
                    row_data[tc] = vals.get(tc, '')
                # Fill matching source→target columns
                if "_raw" in vals:
                    for sc in ds.columns:
                        if sc in dt.columns and sc not in row_data:
                            row_data[sc] = vals["_raw"].get(sc, "")
                        for _sf, _tf in key_pairs + fill_pairs:
                            if sc == _sf and _tf not in row_data:
                                row_data[_tf] = vals["_raw"].get(_sf, "")
                    del vals["_raw"]
                new_rows.append(row_data)
        
        if new_rows:
            import pandas as _pd
            new_df = _pd.DataFrame(new_rows)
            # Only keep columns that exist in dt
            for col in new_df.columns:
                if col not in dt.columns:
                    dt[col] = ''
            dt = _pd.concat([dt, new_df], ignore_index=True)
            print(f'[追加] 将 {len(new_rows)} 条缺失记录追加到目标表')

    # ── 输出结果 ──
    print('=' * 50)
    print('匹配完成！')
    print(f'  成功: {matched} 人')
    if fuzzy_matched:
        print(f'  其中模糊匹配: {len(fuzzy_matched)} 人')
        for orig, m, sc in fuzzy_matched[:15]:
            print(f'    "{orig}" -> "{m}" ({sc})')
        if len(fuzzy_matched) > 15:
            print(f'    ... 共 {len(fuzzy_matched)} 个')
    print(f'  未匹配: {len(unmatched)} 人')

    if unmatched:
        print(f'\n未匹配的索引值（前30）:')
        for parts in unmatched[:30]:
            print(f'  {" / ".join(parts)}')
        if len(unmatched) > 30:
            print(f'  ... 共 {len(unmatched)} 个')

    # ── 保存 ──
    out = Path(a.output) if a.output else (tp.parent / f'已填充_{tp.name}')
    try:
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            dt.to_excel(w, sheet_name=ts, index=False)

        print(f"[完成] {out.resolve()}")
    except PermissionError:
        import tempfile
        fallback = Path(tempfile.gettempdir()) / f"已填充_{tp.stem}{tp.suffix}"
        print(f'[完成] (临时目录) {fallback.resolve()}')

        with pd.ExcelWriter(fallback, engine='openpyxl') as w:

            dt.to_excel(w, sheet_name=ts, index=False)

        print(f'[完成] (临时目录) {fallback.resolve()}')

    except Exception as e:
        print(f"[错误] 保存失败: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()




