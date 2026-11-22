# -*- coding: utf-8 -*-
"""
DrawingAnalysisMiya Excelブック生成・印刷レイアウト設定・新規Excel起動モジュール
- 印刷設定済み（A4横・横1ページ自動収縮・印刷タイトル行繰り返し・ヘッダー/フッター・折り返し自動行高・SUM集計行）のブックを構築
- Excelテンプレート (.xltx) + `excel.exe /t` / Windows Shell `new` により「新規の未保存Excelブック」として開き、
  ユーザーが必要に応じてそのまま印刷または任意の場所へ保存できるようにする
"""

import datetime
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from api_client import APP_NAME, APP_VERSION

TEMP_EXCEL_DIR = os.path.join(tempfile.gettempdir(), f"{APP_NAME}_Excel")
TEMP_FILE_MAX_AGE_SECONDS = 24 * 60 * 60

# 数値フォーマット定義（列名に応じて自動適用）
_INT_COLUMNS = {"数量", "切断本数(枚)"}
_DEC0_COLUMNS = {"切断長さ (mm)", "合計長さ (mm)", "合計切断長 (mm)"}
_DEC1_COLUMNS = {"1個あたり体積 (mm³)", "体積 (mm³)"}
_DEC2_COLUMNS = {"比重 (g/cm³)", "小計重量 (kg)"}
_DEC3_COLUMNS = {"1個あたり重量 (kg)", "単重 (kg)"}

# シート別に末尾合計行 (=SUM) を付与する列
_SHEET_SUM_COLUMNS = {
    "体積・重量解析": {"小計重量 (kg)"},
    "部材表_切断寸法": {"数量", "合計長さ (mm)", "小計重量 (kg)"},
    "使用購入品_手配": {"数量"},
    "使用機械加工品": {"数量"},
}

# 各列の最大表示幅（長文列がA4横印刷時の縮小率を過度に下げないよう調整）
_COLUMN_MAX_WIDTH = {
    "備考 (計算式等)": 40,
    "加工工程・内容": 38,
    "補足 / 根拠": 46,
    "値 / 内容": 40,
    "備考": 32,
}


def _display_width(text):
    """日本語全角文字を考慮した概算表示文字幅を返す"""
    s = str(text or "")
    lines = s.splitlines() or [""]
    return max(sum(2.0 if ord(ch) > 255 else 1.05 for ch in ln) for ln in lines)


def _estimate_wrapped_lines(text, col_width):
    """指定列幅に対する折り返し行数を概算する"""
    s = str(text or "")
    if not s:
        return 1
    usable = max(6.0, float(col_width) - 2.5)
    total_lines = 0
    for ln in s.splitlines():
        if not ln:
            total_lines += 1
        else:
            w = sum(2.0 if ord(ch) > 255 else 1.05 for ch in ln)
            total_lines += max(1, math.ceil(w / usable))
    return max(1, total_lines)


def _clean_cell_value(val):
    if val is None:
        return ""
    if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
        return ""
    return val


def _extract_columns_and_rows(df):
    """pandas.DataFrame または軽量テストDataFrameから (columns, rows) を抽出する"""
    cols = list(getattr(df, "columns", []))
    if hasattr(df, "to_dict"):
        records = df.to_dict(orient="records")
        rows = [[_clean_cell_value(rec.get(c, "")) for c in cols] for rec in records]
    elif hasattr(df, "rows"):
        rows = [[_clean_cell_value(rec.get(c, "")) for c in cols] for rec in df.rows]
    else:
        rows = []
    return cols, rows


def _sanitize_header_footer_text(text):
    """Excelヘッダー/フッターで特殊記号扱いになる '&' をエスケープする"""
    return str(text or "").replace("&", "&&")


def _sanitize_filename_stem(name, max_len=36):
    stem = os.path.splitext(os.path.basename(str(name or "")))[0]
    cleaned = re.sub(r'[\\/:*?"<>|\s]+', "_", stem).strip("_.")
    if not cleaned:
        cleaned = "図面解析結果"
    return cleaned[:max_len]


def _get_number_format(col_name, val):
    if not isinstance(val, (int, float)) or isinstance(val, bool):
        return None
    if col_name in _INT_COLUMNS:
        return "#,##0"
    if col_name in _DEC0_COLUMNS:
        return "#,##0"
    if col_name in _DEC1_COLUMNS:
        return "#,##0.0"
    if col_name in _DEC3_COLUMNS:
        return "#,##0.000"
    if col_name in _DEC2_COLUMNS:
        return "0.00" if "比重" in col_name else "#,##0.00"
    return "#,##0.##"


def _apply_print_setup(ws, num_cols, last_row, drawing_label):
    """そのまま Ctrl+P で美しくA4横印刷できるページ設定をワークシートへ適用する"""
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE

    # 横幅は必ず1ページに収め、縦は行数に応じて自動分割
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    # 2ページ目以降にもレポートタイトル＋表見出し行(1〜4行目)を自動繰り返し印刷
    ws.print_title_rows = "1:4"
    ws.print_area = f"A1:{get_column_letter(max(1, num_cols))}{max(1, last_row)}"

    # 余白（インチ単位・実務帳票向けにすっきり収まる設定）
    ws.page_margins.left = 0.35
    ws.page_margins.right = 0.35
    ws.page_margins.top = 0.55
    ws.page_margins.bottom = 0.55
    ws.page_margins.header = 0.25
    ws.page_margins.footer = 0.25

    # 水平中央寄せ・枠線は明示的罫線で描画するため印刷グリッド線はオフ
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = False
    ws.print_options.gridLines = False

    # 印刷ヘッダー / フッター
    safe_drawing = _sanitize_header_footer_text(drawing_label)
    ws.oddHeader.left.text = safe_drawing
    ws.oddHeader.left.size = 8
    ws.oddHeader.left.font = "Meiryo UI"

    ws.oddHeader.right.text = "印刷日: &D &T"
    ws.oddHeader.right.size = 8
    ws.oddHeader.right.font = "Meiryo UI"

    ws.oddFooter.left.text = _sanitize_header_footer_text(f"{APP_NAME} {APP_VERSION}")
    ws.oddFooter.left.size = 8
    ws.oddFooter.left.font = "Meiryo UI"

    ws.oddFooter.center.text = "&P / &N ページ"
    ws.oddFooter.center.size = 9
    ws.oddFooter.center.font = "Meiryo UI"

    ws.oddFooter.right.text = "&A"
    ws.oddFooter.right.size = 8
    ws.oddFooter.right.font = "Meiryo UI"


def build_excel_workbook(tables, source_file_path="", page_note="", as_template=True):
    """
    build_export_tables の出力から、印刷設定済み・書式付き openpyxl.Workbook を構築する。
    as_template=True の場合は .xltx (Excelテンプレート) として保存可能な状態にする。
    """
    wb = Workbook()
    # デフォルトの空シートを削除
    if wb.active is not None:
        wb.remove(wb.active)

    drawing_name = os.path.basename(source_file_path) if source_file_path else "解析図面"
    drawing_label = f"{drawing_name}{page_note}"
    created_str = datetime.datetime.now().strftime("%Y/%m/%d %H:%M")

    # 共通スタイル（カラー印刷でもモノクロ印刷でも見やすい配色）
    title_fill = PatternFill(start_color="EAF2FD", end_color="EAF2FD", fill_type="solid")
    title_font = Font(name="Meiryo UI", size=13, bold=True, color="1B365D")
    meta_font = Font(name="Meiryo UI", size=9, color="495057")
    header_fill = PatternFill(start_color="D9E8FB", end_color="D9E8FB", fill_type="solid")
    header_font = Font(name="Meiryo UI", size=10, bold=True, color="1B365D")
    body_font = Font(name="Meiryo UI", size=9, color="212529")
    body_bold_font = Font(name="Meiryo UI", size=9, bold=True, color="1B365D")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    highlight_fill = PatternFill(start_color="FFF8E1", end_color="FFF8E1", fill_type="solid")
    total_fill = PatternFill(start_color="E9F5EC", end_color="E9F5EC", fill_type="solid")
    total_font = Font(name="Meiryo UI", size=10, bold=True, color="145A32")

    thin_side = Side(style="thin", color="CBD5E1")
    med_navy = Side(style="medium", color="1B365D")
    double_bottom = Side(style="double", color="1B365D")
    accent_bottom = Side(style="medium", color="0D6EFD")

    title_border = Border(bottom=accent_bottom)
    header_border = Border(left=thin_side, right=thin_side, top=med_navy, bottom=med_navy)
    body_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    total_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_bottom)

    for sheet_name, _, df in tables:
        cols, rows = _extract_columns_and_rows(df)
        if not cols:
            continue

        ws = wb.create_sheet(title=sheet_name[:31])
        num_cols = len(cols)
        last_col_letter = get_column_letter(num_cols)

        # --- Row 1: レポートタイトルバナー ---
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
        title_cell = ws.cell(row=1, column=1, value=f"2D機械図面 解析レポート － {sheet_name}")
        title_cell.font = title_font
        title_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for c_idx in range(1, num_cols + 1):
            cell = ws.cell(row=1, column=c_idx)
            cell.fill = title_fill
            cell.border = title_border
        ws.row_dimensions[1].height = 28

        # --- Row 2: 図面・作成メタ情報 ---
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_cols)
        meta_cell = ws.cell(
            row=2,
            column=1,
            value=f"対象図面: {drawing_label}　｜　作成日時: {created_str}　｜　{APP_NAME} {APP_VERSION}",
        )
        meta_cell.font = meta_font
        meta_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[2].height = 18

        # --- Row 3: 余白行 ---
        ws.row_dimensions[3].height = 6

        # --- Row 4: 表見出し ---
        for c_idx, col_name in enumerate(cols, 1):
            h_cell = ws.cell(row=4, column=c_idx, value=str(col_name))
            h_cell.fill = header_fill
            h_cell.font = header_font
            h_cell.border = header_border
            h_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[4].height = 26

        # --- 列幅を先に計算（データ行の折り返し行高計算に使う） ---
        col_widths = []
        for c_idx, col_name in enumerate(cols):
            max_w = _display_width(col_name) + 3.0
            for r in rows:
                val = r[c_idx] if c_idx < len(r) else ""
                max_w = max(max_w, _display_width(val) + 2.5)
            cap = _COLUMN_MAX_WIDTH.get(col_name, 34)
            final_w = min(max(max_w, 8.0), float(cap))
            col_widths.append(final_w)
            ws.column_dimensions[get_column_letter(c_idx + 1)].width = final_w

        # --- Row 5〜: データ本体 ---
        data_start_row = 5
        for r_offset, row_vals in enumerate(rows):
            excel_row = data_start_row + r_offset
            is_zebra = (r_offset % 2) == 1
            is_summary_key_row = (
                sheet_name == "見積り・搬送サマリー"
                and len(row_vals) >= 2
                and str(row_vals[1])
                in (
                    "推奨トラック車種",
                    "概算材料費 合計",
                    "最大長さ (L) × 最大幅 (W) × 最大高さ (H)",
                    "総重量",
                )
            )

            max_lines = 1
            for c_idx, val in enumerate(row_vals, 1):
                col_name = cols[c_idx - 1]
                cell = ws.cell(row=excel_row, column=c_idx, value=val)
                cell.border = body_border

                if is_summary_key_row:
                    cell.fill = highlight_fill
                    cell.font = body_bold_font
                else:
                    if is_zebra:
                        cell.fill = zebra_fill
                    cell.font = body_font

                num_fmt = _get_number_format(col_name, val)
                if num_fmt:
                    cell.number_format = num_fmt
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    horiz = "center" if col_name in ("品番", "単位") else "left"
                    cell.alignment = Alignment(horizontal=horiz, vertical="center", wrap_text=True)

                est = _estimate_wrapped_lines(val, col_widths[c_idx - 1])
                if est > max_lines:
                    max_lines = est

            ws.row_dimensions[excel_row].height = max(20, max_lines * 14 + 6)

        last_row = data_start_row + len(rows) - 1 if rows else 4

        # --- 集計行 (=SUM) の追加 ---
        sum_cols = _SHEET_SUM_COLUMNS.get(sheet_name)
        if sum_cols and rows:
            total_row = last_row + 1
            ws.cell(row=total_row, column=1, value="合計")
            for c_idx, col_name in enumerate(cols, 1):
                cell = ws.cell(row=total_row, column=c_idx)
                cell.fill = total_fill
                cell.font = total_font
                cell.border = total_border
                if col_name in sum_cols:
                    col_l = get_column_letter(c_idx)
                    cell.value = f"=SUM({col_l}{data_start_row}:{col_l}{last_row})"
                    cell.number_format = _get_number_format(col_name, 1.0) or "#,##0.##"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif c_idx == 1:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[total_row].height = 24
            last_row = total_row

        # 見出し行固定 (A5) & 印刷設定適用
        ws.freeze_panes = "A5"
        _apply_print_setup(ws, num_cols, last_row, drawing_label)

    # 最初のシートをアクティブにする（全シート空の異常をガード）
    if not wb.sheetnames:
        raise ValueError("出力可能なシートデータがありません。")
    wb.active = 0

    wb.template = bool(as_template)
    return wb


def _cleanup_old_temp_files(temp_dir=TEMP_EXCEL_DIR, max_age_sec=TEMP_FILE_MAX_AGE_SECONDS):
    """一定時間以上経過した一時テンプレートファイルを静かに整理する"""
    if not os.path.isdir(temp_dir):
        return
    now = time.time()
    try:
        for fname in os.listdir(temp_dir):
            if not fname.lower().endswith((".xltx", ".xlsx")):
                continue
            fpath = os.path.join(temp_dir, fname)
            try:
                if os.path.isfile(fpath) and (now - os.path.getmtime(fpath)) > max_age_sec:
                    os.remove(fpath)
            except Exception:
                pass
    except Exception:
        pass


def _find_excel_executable():
    """Windows 環境で Microsoft Excel (EXCEL.EXE) の実行パスを検出する"""
    if sys.platform == "win32":
        try:
            import winreg

            reg_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"
            for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
                try:
                    with winreg.OpenKey(hive, reg_path) as key:
                        exe_path, _ = winreg.QueryValueEx(key, "")
                        if exe_path and os.path.isfile(exe_path):
                            return exe_path
                except OSError:
                    continue
        except Exception:
            pass

        common_candidates = [
            os.path.join(
                os.environ.get("ProgramFiles", r"C:\Program Files"),
                "Microsoft Office",
                "root",
                "Office16",
                "EXCEL.EXE",
            ),
            os.path.join(
                os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
                "Microsoft Office",
                "root",
                "Office16",
                "EXCEL.EXE",
            ),
        ]
        for cand in common_candidates:
            if os.path.isfile(cand):
                return cand

    return shutil.which("excel.exe") or shutil.which("EXCEL.EXE")


def launch_template_in_new_excel(template_path):
    """
    テンプレートファイル (.xltx) を「新規の未保存ブック」として Excel で開く。
    1. Microsoft Excel が見つかる場合: `EXCEL.EXE /t <path>` で確実に新規未保存ブックとして起動
    2. それ以外: Windows Shell の `"new"` 動詞 (`os.startfile(path, "new")`) → フォールバック `os.startfile(path)`
    """
    excel_exe = _find_excel_executable()
    if excel_exe:
        subprocess.Popen([excel_exe, "/t", template_path])
        return "excel_exe"

    if hasattr(os, "startfile"):
        try:
            os.startfile(template_path, "new")
            return "shell_new"
        except OSError:
            os.startfile(template_path)
            return "shell_open"

    raise RuntimeError("この環境ではExcelを自動起動できませんでした。")


def open_tables_in_new_excel(
    tables, source_file_path="", page_note="", temp_dir=TEMP_EXCEL_DIR
):
    """
    解析テーブルから印刷設定済みテンプレート (.xltx) を一時生成し、新規のExcelブックとして開く。
    戻り値: (template_path, launch_method)
    """
    if not tables:
        raise ValueError("Excelで開くデータがありません。")

    os.makedirs(temp_dir, exist_ok=True)
    _cleanup_old_temp_files(temp_dir)

    stem = _sanitize_filename_stem(source_file_path)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    template_path = os.path.join(temp_dir, f"図面解析_{stem}_{ts}.xltx")

    wb = build_excel_workbook(
        tables,
        source_file_path=source_file_path,
        page_note=page_note,
        as_template=True,
    )
    wb.save(template_path)
    method = launch_template_in_new_excel(template_path)
    return template_path, method


def save_tables_as_xlsx(tables, file_path, source_file_path="", page_note=""):
    """フォールバック用: 印刷設定済みの標準Excelブック (.xlsx) として指定パスへ保存する"""
    if not tables:
        raise ValueError("保存するデータがありません。")
    wb = build_excel_workbook(
        tables,
        source_file_path=source_file_path,
        page_note=page_note,
        as_template=False,
    )
    wb.save(file_path)
    return file_path
