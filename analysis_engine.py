# -*- coding: utf-8 -*-
"""
2D機械図面 解析エンジンモジュール
- 体積・重量解析（数量反映・全体外形寸法読み取り・3ステップマルチターン解析）
- 部材表・切断寸法解析（鋼材種別・断面規格・切断寸法・端部切断・数量・合計長さ・定尺目安）
- 使用購入品（手配用：メーカー・型式・数量・単位・手配区分）抽出
- 使用機械加工品（製作・外注手配用：素材寸法・加工工程・公差・熱処理/表面処理）抽出
- 見積り・搬送サマリー（材質別重量集計・概算材料費・推奨トラック選定・鋼材定尺発注集計）
- 高速PDF Visual Grounding 前処理 (PyMuPDF優先 + 適応的DPI + マルチページ対応)
- ASTベース安全数式評価 / JSON修復 / 並列実行 / キャンセル制御 / エクスポート用テーブル生成
"""

import ast
import concurrent.futures
import io
import json
import math
import operator
import os
import re
import warnings
from typing import List

import cv2
import fitz  # PyMuPDF
import numpy as np
import pandas as pd
import pdfplumber
from PIL import Image
from pydantic import BaseModel, Field

warnings.filterwarnings("ignore", category=FutureWarning)
import google.generativeai as genai

from api_client import (
    AnalysisCancelledError,
    DEFAULT_AZURE_API_VERSION,
    DEFAULT_GEMINI_MODEL_STEP1,
    DEFAULT_GEMINI_MODEL_STEP2,
    DEFAULT_GEMINI_MODEL_STEP3,
    DiagnosedAPIError,
    RateLimiter,
    encode_pil_image_data_url,
    generate_structured_json,
    get_recommended_gemini_settings,
    parse_azure_endpoint_info,
    safe_float as safe_float_param,
    safe_int as safe_int_param,
)


# ==============================
# 材質別デフォルト目安単価 (円/kg) および標準比重
# ==============================
DEFAULT_MATERIAL_UNIT_PRICES = {
    "SS400": 200.0,
    "SPHC": 200.0,
    "SPCC": 210.0,
    "STKR": 220.0,
    "STK": 220.0,
    "SGP": 220.0,
    "S45C": 280.0,
    "S50C": 300.0,
    "SCM435": 450.0,
    "SCM440": 450.0,
    "SKD11": 1200.0,
    "SUS304": 750.0,
    "SUS316": 1050.0,
    "SUS430": 600.0,
    "A5052": 900.0,
    "A6061": 950.0,
    "A6063": 880.0,
    "AL": 900.0,
    "FC250": 350.0,
    "FCD450": 400.0,
    "C3604": 1500.0,
    "MCナイロン": 2200.0,
    "POM": 1800.0,
}

DEFAULT_MATERIAL_DENSITIES = {
    "SS400": 7.85,
    "SPHC": 7.85,
    "SPCC": 7.85,
    "STKR": 7.85,
    "STK": 7.85,
    "SGP": 7.85,
    "S45C": 7.85,
    "S50C": 7.85,
    "SCM": 7.85,
    "SKD": 7.85,
    "SUS304": 7.93,
    "SUS316": 7.98,
    "SUS430": 7.70,
    "SUS": 7.93,
    "A5052": 2.68,
    "A6061": 2.70,
    "A6063": 2.70,
    "AL": 2.70,
    "FC250": 7.20,
    "FCD": 7.10,
    "C3604": 8.50,
    "MCナイロン": 1.16,
    "POM": 1.41,
}


# ==============================
# 構造化データ（Structured Output）のスキーマ定義
# ==============================
class Element(BaseModel):
    element_name: str = Field(
        description="構成要素の名称（例: ベースプレート、補強リブ、円柱ボス、中心貫通穴、ザグリ穴）"
    )
    dimensions: str = Field(
        description="三面図（正面図、平面図、側面図など）から対応する寸法を特定し、統合した寸法情報（例: W150 x H100 x D10, 径D20 x 深さ15）。絶対に空欄にしないこと。"
    )
    calculation_formula: str = Field(
        description="三面図から復元した1個あたり要素形状の体積(mm³)を算出する数式。文字、単位、変数、スペースを含めず、算術評価可能な文字列（例: '150*100*10'、円柱や穴などの空間を引き算する場合は '-(10**2)*3.14159265*15' とする）。"
    )
    notes: str = Field(
        description="立体の解釈に関する特記事項。ない場合は必ず空文字（''）にすること。記述する場合も極めて簡潔（15文字以内）にまとめること。"
    )


class PartBasic(BaseModel):
    part_number: str = Field(
        description="部品番号（バルーン記号、部品表の番号、または図面内で指定されている番号）。特定できない場合は空文字にする"
    )
    part_name: str = Field(
        description="部品名（タイトルブロックや部品表に記載されている日本語または英語の名称）。特定できない場合は '不明な部品' とする"
    )
    material: str = Field(
        description="部品の材質（例: SS400, SUS304, S45C, A5052, FC250 など）。特定できない場合は空文字にする"
    )
    density_g_cm3: float = Field(
        description="材質に対応する一般的な比重（g/cm³）。図面に記載がない場合も一般的な値（SS400なら7.85等）を推測して割り当てること。全く不明な場合は 0.0 とする"
    )
    quantity: int = Field(
        default=1,
        description="この部品の図面全体での必要数量（個数・本数）。部品表や注記（例: 2ヶ、×4、数量2 等）から正確に読み取ること。記載がない場合は 1 とする。",
    )


class PartListResult(BaseModel):
    overall_length_mm: float = Field(
        default=0.0,
        description="図面に描かれた製品・ユニット全体の最大長さ L (mm)。三面図の総寸法から読み取ること。不明な場合は 0.0。",
    )
    overall_width_mm: float = Field(
        default=0.0,
        description="図面に描かれた製品・ユニット全体の最大幅 W (mm)。不明な場合は 0.0。",
    )
    overall_height_mm: float = Field(
        default=0.0,
        description="図面に描かれた製品・ユニット全体の最大高さ H (mm)。不明な場合は 0.0。",
    )
    transport_notes: str = Field(
        default="",
        description="搬送・荷役時の注意点（長尺物、偏荷重、突起物、分割輸送の可否など）。特になければ空文字。",
    )
    parts: List[PartBasic] = Field(description="図面から抽出されたすべての部品の基本情報のリスト")


class PartElementsResult(BaseModel):
    elements: List[Element] = Field(
        description="指定された部品（1個あたり）を構成する、体積計算が可能な要素形状のリスト"
    )


class CompletePart(BaseModel):
    part_number: str = Field(description="部品番号")
    part_name: str = Field(description="部品名")
    material: str = Field(description="材質")
    density_g_cm3: float = Field(description="比重")
    quantity: int = Field(default=1, description="必要数量（個数・本数）")
    elements: List[Element] = Field(description="この部品（1個あたり）を構成する要素のリスト")


class MissingPartsResult(BaseModel):
    missing_parts: List[CompletePart] = Field(
        description="検証フェーズで新たに発見された、抽出漏れの部品とその要素のリスト。見落としがない場合は空リストにする。"
    )


# --- 部材表（切断寸法）・使用購入品（手配）・使用機械加工品 スキーマ ---
class MaterialItem(BaseModel):
    part_number: str = Field(
        description="対応する部品番号・符号・バルーン番号（特定できない場合は空文字）"
    )
    part_name: str = Field(
        description="使用する部品名または部位（例: ベースプレート、フレーム縦柱、補強リブ）"
    )
    material_type: str = Field(
        description="鋼材・部材種別（例: 鋼板(PL), 平鋼(FB), 等辺山形鋼(L鋼), 溝形鋼(C鋼), H形鋼(H), 角パイプ(STKR), 丸パイプ(STK/SGP), 丸棒(RB), 樹脂板 など）"
    )
    material_spec: str = Field(
        description="材質規格（例: SS400, SUS304, S45C, A5052, SPHC, MCナイロン など）"
    )
    section_size: str = Field(
        description="断面規格または板厚（例: PL10, FB-6x50, L-50x50x6, C-100x50x5x7.5, □-50x50x2.3, φ30, 50A など）"
    )
    cut_dimensions: str = Field(
        description="手配・切断用の切断寸法（板材は 幅x長さ 例: '150x300'、形鋼・棒・パイプは切断長 例: 'L=1250'）。必ず具体的なmm寸法を記載すること。"
    )
    cut_length_mm: float = Field(
        default=0.0,
        description="形鋼・パイプ・丸棒・平鋼などの長尺部材における1本あたりの切断長さ(mm)の数値（例: 1250.0）。板材（鋼板PL・樹脂板等）の場合は必ず 0.0 とする。",
    )
    cut_angle_end: str = Field(
        default="90°直角",
        description="端部の切断形状（例: '90°直角', '両端45°留め切り', '片側45°', 'レーザー切断', 'シャーリング'）",
    )
    quantity: int = Field(
        default=1,
        description="必要な切断本数または枚数（整数。図面の数量・×個数から正確に読み取ること）",
    )
    unit_volume_formula: str = Field(
        default="",
        description="この部材1本/1枚あたりの材料体積(mm³)の算術式（例: 鋼板PL10x150x300なら '10*150*300'、角パイプ50x50xt2.3xL1200なら '((50*50)-(45.4*45.4))*1200'、L-50x50x6xL1000なら '((50*6)+(44*6))*1000'）。純粋な数値と +-*/()** のみ。",
    )
    density_g_cm3: float = Field(
        default=0.0,
        description="材質に対応する比重(g/cm³)。SS400なら7.85、SUS304なら7.93、アルミなら2.70。",
    )
    surface_treatment: str = Field(
        default="",
        description="表面処理・塗装・メッキ指示（例: 塗装、溶融亜鉛メッキ、酸洗、無処理 等。記載なければ空文字）",
    )
    notes: str = Field(
        default="",
        description="開先加工・曲げ・穴あけ・手配上の注意点など（なければ空文字）",
    )


class MaterialListResult(BaseModel):
    overall_length_mm: float = Field(default=0.0, description="図面全体の最大外形長さ L (mm)")
    overall_width_mm: float = Field(default=0.0, description="図面全体の最大外形幅 W (mm)")
    overall_height_mm: float = Field(default=0.0, description="図面全体の最大外形高さ H (mm)")
    transport_notes: str = Field(default="", description="搬送・荷役時の注意点")
    materials: List[MaterialItem] = Field(
        description="図面内で使用されている鋼材・素材・板材・形鋼等の切断寸法付き部材表"
    )


class PurchasedPartItem(BaseModel):
    part_number: str = Field(
        description="部品番号・バルーン番号（特定できない場合は空文字）"
    )
    part_name: str = Field(
        description="購入品の名称（例: 六角穴付ボルト、平座金、深溝玉軸受、Oリング、エアシリンダ、リニアガイド など）"
    )
    manufacturer: str = Field(
        description="メーカー名または規格区分（例: MISUMI, THK, NSK, SMC, OMRON, JIS規格品 等。記載がない場合は空文字）"
    )
    model_number: str = Field(
        description="手配に必要な型式・呼び径・サイズ規格（例: M8x25, 6204ZZ, P-20, CDQ2B32-50DZ など）"
    )
    material_or_spec: str = Field(
        description="材質・仕様・表面処理（例: SCM435 黒染め, SUS304, NBR, DC24V 等。なければ空文字）"
    )
    quantity: int = Field(
        default=1,
        description="手配必要数量（整数。部品表や図面注記の『4-M8』等から総必要数を正確に読み取ること）",
    )
    unit: str = Field(
        default="個",
        description="手配単位（例: '個', '本', '組', 'セット', '枚', 'm'）",
    )
    procurement_category: str = Field(
        default="標準市販品",
        description="手配区分（例: 'JIS規格品', '標準市販品', 'メーカー取寄品', '電気・空圧機器'）",
    )
    notes: str = Field(
        default="",
        description="取り付け箇所や手配時の注意点（なければ空文字）",
    )


class PurchasedPartListResult(BaseModel):
    purchased_parts: List[PurchasedPartItem] = Field(
        description="図面内で使用されている市販品・標準規格品・メーカー購入品の手配一覧"
    )


class MachinedPartItem(BaseModel):
    part_number: str = Field(
        description="部品番号・バルーン番号（特定できない場合は空文字）"
    )
    part_name: str = Field(
        description="機械加工品の部品名（例: シャフト、ハウジング、ベースブロック、ブラケット など）"
    )
    material: str = Field(
        description="材質（例: S45C, SS400, SUS304, A5052, FC250 など）"
    )
    blank_or_size: str = Field(
        description="概略外形寸法または必要素材サイズ（例: 150 x 100 x 20, φ50 x L120 など）"
    )
    machining_processes: str = Field(
        description="図面形状・注記から読み取れる加工工程・加工内容（例: マシニング・4-M8タップ・φ10ザグリ穴、旋盤・外径研磨・キー溝加工、溶接製缶後フライス仕上 など）"
    )
    precision_surface: str = Field(
        description="寸法公差・はめあい・幾何公差・表面粗さ（例: φ20H7, Ra1.6, 平行度0.02 等。なければ空文字）"
    )
    heat_surface_treatment: str = Field(
        default="",
        description="熱処理・表面処理（例: 焼入れHRC55, 無電解ニッケルメッキ, 黒染め, アルマイト, 塗装 等。なければ空文字）",
    )
    quantity: int = Field(
        default=1,
        description="製作必要数量（整数）",
    )
    notes: str = Field(
        default="",
        description="加工上の注意点や特記事項（なければ空文字）",
    )


class MachinedPartListResult(BaseModel):
    machined_parts: List[MachinedPartItem] = Field(
        description="図面内で切削・旋盤・フライス・研削・穴あけ等の機械加工を要する製作部品の一覧"
    )


class PartsBreakdownResult(BaseModel):
    overall_length_mm: float = Field(default=0.0, description="図面全体の最大外形長さ L (mm)")
    overall_width_mm: float = Field(default=0.0, description="図面全体の最大外形幅 W (mm)")
    overall_height_mm: float = Field(default=0.0, description="図面全体の最大外形高さ H (mm)")
    transport_notes: str = Field(default="", description="搬送・荷役時の注意点")
    materials: List[MaterialItem] = Field(
        description="使用部材・切断寸法（部材表）の一覧。解析対象外の場合は空リストにする。"
    )
    purchased_parts: List[PurchasedPartItem] = Field(
        description="使用購入品（手配用）の一覧。解析対象外の場合は空リストにする。"
    )
    machined_parts: List[MachinedPartItem] = Field(
        description="使用機械加工品の一覧。解析対象外の場合は空リストにする。"
    )


# ==============================
# 安全なASTベース算術評価器 & 共通ユーティリティ
# ==============================
_BIN_OPS = {
    ast.Add: operator.add,
    ast.Sub: operator.sub,
    ast.Mult: operator.mul,
    ast.Div: operator.truediv,
    ast.Pow: operator.pow,
}
_UNARY_OPS = {
    ast.UAdd: operator.pos,
    ast.USub: operator.neg,
}


def _eval_ast_node(node):
    if isinstance(node, ast.Expression):
        return _eval_ast_node(node.body)
    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
        return float(node.value)
    if isinstance(node, ast.UnaryOp) and type(node.op) in _UNARY_OPS:
        return _UNARY_OPS[type(node.op)](_eval_ast_node(node.operand))
    if isinstance(node, ast.BinOp) and type(node.op) in _BIN_OPS:
        left = _eval_ast_node(node.left)
        right = _eval_ast_node(node.right)
        if isinstance(node.op, ast.Pow):
            if abs(right) > 12:
                raise ValueError(f"Exponent out of safe bounds: {right}")
        if isinstance(node.op, ast.Div) and right == 0:
            raise ZeroDivisionError("Division by zero")
        return _BIN_OPS[type(node.op)](left, right)
    raise ValueError(f"Unsupported AST node: {type(node).__name__}")


def evaluate_formula(formula_str):
    """
    ASTベースで安全に算術数式文字列を評価する（eval不使用）。
    全角記号・×・÷・π・pi・^ 等の表記揺れを自動正規化し、ゼロ除算や過大なべき乗を安全にガードする。
    """
    if not formula_str:
        return 0.0
    try:
        s = str(formula_str).strip()
        s = s.replace("×", "*").replace("÷", "/").replace("−", "-").replace("ー", "-")
        s = s.replace("π", "3.14159265")
        s = re.sub(r"\bpi\b", "3.14159265", s, flags=re.IGNORECASE)
        s = s.replace("^", "**")
        s = s.translate(
            str.maketrans("０１２３４５６７８９＋－＊／（）．", "0123456789+-*/().")
        )
        allowed_chars = set("0123456789+-*/(). ")
        sanitized = "".join(c for c in s if c in allowed_chars)
        if not sanitized.strip():
            return 0.0
        tree = ast.parse(sanitized, mode="eval")
        val = float(_eval_ast_node(tree))
        if math.isnan(val) or math.isinf(val):
            return 0.0
        return val
    except Exception as e:
        print(f"Formula evaluation error: {e} for formula: '{formula_str}'")
        return 0.0


def repair_and_parse_json(json_str):
    """途切れたJSON文字列を修復してパースを試みるベストエフォート機能"""
    for i in range(len(json_str), max(0, len(json_str) - 200), -1):
        temp_str = json_str[:i]

        stack = []
        in_string = False
        escape = False
        for char in temp_str:
            if escape:
                escape = False
                continue
            if char == "\\":
                escape = True
                continue
            if char == '"':
                in_string = not in_string
                continue
            if not in_string:
                if char in "{[":
                    stack.append(char)
                elif char == "}":
                    if stack and stack[-1] == "{":
                        stack.pop()
                elif char == "]":
                    if stack and stack[-1] == "[":
                        stack.pop()

        repaired = temp_str
        if in_string:
            repaired += '"'

        while stack:
            char = stack.pop()
            if char == "{":
                repaired += "}"
            elif char == "[":
                repaired += "]"

        try:
            return json.loads(repaired)
        except json.JSONDecodeError:
            continue

    return None


def parse_json_response(result_text, status_callback=None):
    """モデルからのテキストレスポンスからJSONを抽出し、パースするヘルパー"""
    if not result_text:
        return None
    start = result_text.find("{")
    end = result_text.rfind("}")

    if start != -1:
        if end != -1 and end > start:
            json_str = result_text[start : end + 1]
        else:
            json_str = result_text[start:]
    else:
        json_str = result_text

    try:
        return json.loads(json_str)
    except json.JSONDecodeError:
        if status_callback:
            status_callback("JSONの自己修復を試みています...")
        return repair_and_parse_json(json_str)


def safe_float(val):
    """AIが数値を文字列として返却した場合の型エラーを防ぐヘルパー"""
    if val is None or val == "":
        return 0.0
    try:
        if isinstance(val, str):
            val = val.replace(",", "")
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def safe_quantity(val):
    """数量を安全に正の整数(最低1)へ正規化する"""
    if val is None or val == "":
        return 1
    try:
        if isinstance(val, str):
            m = re.search(r"\d+", val.replace(",", ""))
            if not m:
                return 1
            q = int(m.group(0))
        else:
            q = int(round(float(val)))
        return max(1, q)
    except Exception:
        return 1


def resolve_material_density(material_name, fallback_density=7.85):
    """材質名から代表的な比重(g/cm³)を推定する"""
    mat = str(material_name or "").upper().strip()
    if mat:
        for key, den in DEFAULT_MATERIAL_DENSITIES.items():
            if key.upper() in mat:
                return den
    return safe_float(fallback_density) or 7.85


def resolve_material_unit_price(material_name, custom_prices=None, default_price=250.0):
    """材質名から概算材料単価(円/kg)を解決する"""
    mat = str(material_name or "").strip()
    mat_upper = mat.upper()

    if isinstance(custom_prices, dict):
        if mat in custom_prices:
            return safe_float(custom_prices[mat])
        for k, v in custom_prices.items():
            if k.upper() == mat_upper and mat_upper:
                return safe_float(v)
        for k, v in custom_prices.items():
            if k and k.upper() in mat_upper:
                return safe_float(v)

    if mat_upper:
        for k, v in DEFAULT_MATERIAL_UNIT_PRICES.items():
            if k.upper() in mat_upper:
                return v

    return safe_float(default_price) or 250.0


def calculate_volume_and_weight(parts, default_density=7.85):
    """
    体積・重量（1個あたり および 数量掛け小計）をPython側で計算して付与し、
    全体合計 (total_volume_mm3, total_weight_kg) を返す。
    """
    total_overall_volume = 0.0
    total_overall_weight = 0.0

    for part in parts:
        p_mat = str(part.get("material") or "")
        p_den = safe_float(part.get("density_g_cm3"))
        if p_den <= 0.0:
            p_den = resolve_material_density(p_mat, default_density)
            part["density_g_cm3"] = p_den

        qty = safe_quantity(part.get("quantity", 1))
        part["quantity"] = qty

        unit_volume = 0.0
        for elem in part.get("elements", []):
            e_calc = str(elem.get("calculation_formula") or "")
            e_vol = evaluate_formula(e_calc)
            unit_volume += e_vol
            elem["calculated_volume_mm3"] = e_vol
            elem["calculated_weight_kg"] = e_vol * p_den / 1000000.0

        unit_weight = unit_volume * p_den / 1000000.0
        subtotal_volume = unit_volume * qty
        subtotal_weight = unit_weight * qty

        part["unit_volume_mm3"] = unit_volume
        part["unit_weight_kg"] = unit_weight
        part["calculated_volume_mm3"] = subtotal_volume
        part["calculated_weight_kg"] = subtotal_weight

        total_overall_volume += subtotal_volume
        total_overall_weight += subtotal_weight

    return total_overall_volume, total_overall_weight


def _is_plate_material(material_type, section_size):
    """部材種別・断面規格から板材（鋼板PL・樹脂板等）か長尺鋼材かを判定する"""
    m_type = str(material_type or "").upper()
    sec = str(section_size or "").upper()
    combined = f"{m_type} {sec}"
    # 平鋼(FB)は長尺扱いとして除外
    if any(k in combined for k in ("平鋼", "フラットバー", "FB-", "FB ")):
        return False
    if re.search(r"\bPL\s*\d", sec) or sec.startswith("PL"):
        return True
    return any(k in m_type for k in ("鋼板", "プレート", "板", "PL", "PLATE", "SHEET"))


def _extract_max_dimension_mm(dim_str):
    """切断寸法文字列（例: '150x300', 'L=1250', 't10x1200x2400'）から最大寸法(mm)を抽出する"""
    nums = [safe_float(x) for x in re.findall(r"\d+(?:\.\d+)?", str(dim_str or ""))]
    return max(nums) if nums else 0.0


def calculate_materials_metrics(materials, default_density=7.85):
    """部材表（切断寸法）の合計長さ(mm)・単重(kg)・小計重量(kg)をPython側で計算・補完する"""
    total_mat_weight = 0.0
    for item in materials:
        qty = safe_quantity(item.get("quantity", 1))
        item["quantity"] = qty

        is_plate = _is_plate_material(item.get("material_type"), item.get("section_size"))
        item["is_plate"] = is_plate

        cut_len = safe_float(item.get("cut_length_mm", 0.0))
        # 板材は定尺バー集計から除外する
        if is_plate:
            cut_len = 0.0
        elif cut_len <= 0.0:
            # cut_length_mm が未設定でも cut_dimensions に "L=1250" 等があれば補完
            dim_str = str(item.get("cut_dimensions") or "")
            m_len = re.search(r"L\s*[=:]?\s*(\d+(?:\.\d+)?)", dim_str, re.IGNORECASE)
            if m_len:
                cut_len = safe_float(m_len.group(1))
        item["cut_length_mm"] = cut_len
        item["total_length_mm"] = cut_len * qty if cut_len > 0.0 else 0.0

        mat_spec = str(item.get("material_spec") or "")
        den = safe_float(item.get("density_g_cm3", 0.0))
        if den <= 0.0:
            den = resolve_material_density(mat_spec, default_density)
        item["density_g_cm3"] = den

        formula = str(item.get("unit_volume_formula") or "")
        unit_vol = evaluate_formula(formula) if formula else 0.0
        unit_wt = unit_vol * den / 1000000.0 if unit_vol > 0.0 else 0.0
        subtotal_wt = unit_wt * qty

        item["unit_weight_kg"] = unit_wt
        item["total_weight_kg"] = subtotal_wt
        total_mat_weight += subtotal_wt

    return total_mat_weight


# ==============================
# 搬送・トラック選定 & 見積り・定尺集計エンジン
# ==============================
_TRUCK_SPECS = [
    ("軽トラック (平ボディ)", 350.0, 1900.0, 1350.0),
    ("1t〜1.5t 小型トラック (平ボディ)", 1250.0, 2800.0, 1550.0),
    ("2t 標準トラック (平ボディ)", 2000.0, 3100.0, 1600.0),
    ("2t ロングトラック (平ボディ)", 2000.0, 4300.0, 1750.0),
    ("4t 標準トラック (平ボディ)", 4000.0, 6200.0, 2150.0),
    ("4t ワイド・ロングトラック", 4000.0, 7200.0, 2350.0),
    ("10t 大型トラック (平ボディ)", 10000.0, 9600.0, 2400.0),
]


def recommend_truck(
    total_weight_kg,
    length_mm=0.0,
    width_mm=0.0,
    height_mm=0.0,
    max_member_length_mm=0.0,
    max_unit_weight_kg=0.0,
):
    """
    総重量・全体外形寸法(L×W×H)・最大長尺部材長から、日本の実務基準に基づく
    推奨トラック車種・荷役アドバイス・判定根拠を算出して辞書で返す。
    """
    wt = max(0.0, safe_float(total_weight_kg))
    # L/W は大きい方を荷台長方向、小さい方を荷台幅方向に配置可能として評価
    dim_l = max(safe_float(length_mm), safe_float(max_member_length_mm))
    dim_w = safe_float(width_mm)
    dim_h = safe_float(height_mm)
    if dim_w > dim_l:
        dim_l, dim_w = dim_w, dim_l

    # 梱包・パレット・治具分の安全余裕 (約10%増) を考慮した参考積載重量
    design_wt = wt * 1.10

    if wt <= 0.0 and dim_l <= 0.0 and dim_w <= 0.0 and dim_h <= 0.0:
        return {
            "recommended_truck": "判定データなし（体積または部材表を解析してください）",
            "bed_spec": "---",
            "handling_advice": "重量解析または部材表解析を実行すると荷役アドバイスが表示されます",
            "reasons": "重量・外形寸法データが未抽出です",
            "design_weight_kg": 0.0,
            "effective_length_mm": 0.0,
            "effective_width_mm": 0.0,
            "effective_height_mm": 0.0,
        }

    selected_name = "大型低床トレーラー / 特殊輸送・分割輸送を要検討"
    bed_spec = "標準10t車の上限超過"
    for name, max_kg, max_l, max_w in _TRUCK_SPECS:
        if design_wt <= max_kg and dim_l <= max_l and dim_w <= max_w:
            selected_name = name
            bed_spec = f"積載目安 〜{int(max_kg):,}kg / 荷台目安 L{int(max_l):,}×W{int(max_w):,}mm"
            break

    reasons = []
    if wt > 0:
        reasons.append(f"総重量 {wt:,.1f} kg（梱包余裕10%見込み: {design_wt:,.1f} kg）")
    else:
        reasons.append("総重量データなし（寸法ベースまたは目安表示）")
    if dim_l > 0 or dim_w > 0 or dim_h > 0:
        reasons.append(f"最大外形/長尺目安 L{dim_l:,.0f} × W{dim_w:,.0f} × H{dim_h:,.0f} mm")
    if dim_h > 2400:
        reasons.append("高さが 2,400mm を超えるため低床車・横倒し積載・分割輸送を要確認")

    # 荷役（積み降ろし）アドバイス
    if wt >= 1000 or max_unit_weight_kg >= 500:
        handling = "重量物のため、ユニック車（作業半径・吊り荷重要確認）または大型フォークリフト・天井クレーンでの荷役を推奨"
    elif wt >= 150 or max_unit_weight_kg >= 50:
        handling = "手降ろし困難な重量のため、ユニック車（クレーン付トラック）またはフォークリフトの手配を推奨"
    elif wt > 0:
        handling = "2名作業またはハンドリフト・小型フォークリフトで荷役可能目安"
    else:
        handling = "重量解析または部材表解析を実行すると荷役アドバイスが表示されます"

    return {
        "recommended_truck": selected_name,
        "bed_spec": bed_spec,
        "handling_advice": handling,
        "reasons": " ／ ".join(reasons),
        "design_weight_kg": design_wt,
        "effective_length_mm": dim_l,
        "effective_width_mm": dim_w,
        "effective_height_mm": dim_h,
    }


def build_estimation_and_logistics_summary(
    result_data, custom_prices=None, default_unit_price=250.0, default_density=7.85
):
    """
    解析結果から業務判断用の統合サマリーを作成する:
    - 材質別重量集計 ＆ 概算材料費（体積解析があれば体積ベース優先、なければ部材表ベース）
    - 推奨トラック車種 ＆ 荷役アドバイス
    - 鋼材・断面規格ごとの定尺発注集計（ロス10%見込み、5.5m / 6.0m 定尺目安本数）
    """
    if not result_data:
        return {}

    parts = result_data.get("parts", [])
    materials = result_data.get("materials", [])
    purchased = result_data.get("purchased_parts", [])
    machined = result_data.get("machined_parts", [])
    overall = result_data.get("overall_dimensions", {}) or {}

    # 1. 材質別重量集計（体積解析結果があればそちらを優先し、なければ部材表の計算重量を使用）
    mat_weight_map = {}  # {material_label: {"density": float, "weight_kg": float}}
    max_unit_weight = 0.0
    weight_source = "なし"

    if materials:
        calculate_materials_metrics(materials, default_density=default_density)

    if parts:
        weight_source = "体積・重量解析"
        calculate_volume_and_weight(parts, default_density=default_density)
        for p in parts:
            mat_name = str(p.get("material") or "").strip() or "材質不明"
            den = safe_float(p.get("density_g_cm3"))
            sub_wt = max(0.0, safe_float(p.get("calculated_weight_kg")))
            u_wt = max(0.0, safe_float(p.get("unit_weight_kg")))
            max_unit_weight = max(max_unit_weight, u_wt)
            entry = mat_weight_map.setdefault(mat_name, {"density": den, "weight_kg": 0.0})
            entry["weight_kg"] += sub_wt
            if entry["density"] <= 0 and den > 0:
                entry["density"] = den
    elif materials:
        weight_source = "部材表 (切断寸法)"
        for m in materials:
            mat_name = str(m.get("material_spec") or "").strip() or "材質不明"
            den = safe_float(m.get("density_g_cm3"))
            sub_wt = max(0.0, safe_float(m.get("total_weight_kg")))
            u_wt = max(0.0, safe_float(m.get("unit_weight_kg")))
            max_unit_weight = max(max_unit_weight, u_wt)
            entry = mat_weight_map.setdefault(mat_name, {"density": den, "weight_kg": 0.0})
            entry["weight_kg"] += sub_wt
            if entry["density"] <= 0 and den > 0:
                entry["density"] = den

    material_cost_rows = []
    total_weight_kg = 0.0
    total_material_cost_jpy = 0.0
    for mat_name, info in sorted(
        mat_weight_map.items(), key=lambda kv: kv[1]["weight_kg"], reverse=True
    ):
        wt = info["weight_kg"]
        unit_price = resolve_material_unit_price(mat_name, custom_prices, default_unit_price)
        cost = wt * unit_price
        total_weight_kg += wt
        total_material_cost_jpy += cost
        material_cost_rows.append(
            {
                "material": mat_name,
                "density_g_cm3": info["density"],
                "weight_kg": wt,
                "unit_price_jpy": unit_price,
                "estimated_cost_jpy": cost,
            }
        )

    # 2. 鋼材・断面規格ごとの定尺発注集計（部材表の名寄せ）
    stock_map = {}
    max_member_len = 0.0
    for item in materials:
        m_type = str(item.get("material_type") or "").strip() or "その他部材"
        m_spec = str(item.get("material_spec") or "").strip() or "材質不明"
        sec = str(item.get("section_size") or "").strip() or str(item.get("cut_dimensions") or "-")
        qty = safe_quantity(item.get("quantity", 1))
        cut_l = safe_float(item.get("cut_length_mm", 0.0))
        tot_l = safe_float(item.get("total_length_mm", 0.0))
        tot_wt = safe_float(item.get("total_weight_kg", 0.0))
        is_plate = bool(item.get("is_plate", _is_plate_material(m_type, sec)))
        dim_max = _extract_max_dimension_mm(item.get("cut_dimensions"))
        max_member_len = max(max_member_len, cut_l, dim_max)

        key = (m_type, m_spec, sec)
        agg = stock_map.setdefault(
            key,
            {
                "material_type": m_type,
                "material_spec": m_spec,
                "section_size": sec,
                "is_plate": is_plate,
                "total_pieces": 0,
                "total_length_mm": 0.0,
                "total_weight_kg": 0.0,
            },
        )
        agg["total_pieces"] += qty
        if not is_plate:
            agg["total_length_mm"] += tot_l
        agg["total_weight_kg"] += tot_wt

    stock_summary_rows = []
    for agg in sorted(
        stock_map.values(),
        key=lambda x: (x["material_type"], x["material_spec"], x["section_size"]),
    ):
        tot_l = agg["total_length_mm"]
        if not agg.get("is_plate") and tot_l > 0:
            # 切断しろ・端材ロス10%見込みでの5.5m / 6.0m定尺目安本数
            req_l = tot_l * 1.10
            bars_55 = math.ceil(req_l / 5500.0)
            bars_60 = math.ceil(req_l / 6000.0)
            stock_guide = f"5.5m×{bars_55}本 ／ 6.0m×{bars_60}本 (ロス10%込)"
        else:
            stock_guide = "板材・定尺対象外（切板・定尺板手配）"
        agg["stock_guide"] = stock_guide
        stock_summary_rows.append(agg)

    # 3. トラック選定
    truck_info = recommend_truck(
        total_weight_kg=total_weight_kg,
        length_mm=overall.get("length_mm", 0.0),
        width_mm=overall.get("width_mm", 0.0),
        height_mm=overall.get("height_mm", 0.0),
        max_member_length_mm=max_member_len,
        max_unit_weight_kg=max_unit_weight,
    )

    return {
        "weight_source": weight_source,
        "total_weight_kg": total_weight_kg,
        "total_material_cost_jpy": total_material_cost_jpy,
        "overall_dimensions": {
            "length_mm": safe_float(overall.get("length_mm", 0.0)),
            "width_mm": safe_float(overall.get("width_mm", 0.0)),
            "height_mm": safe_float(overall.get("height_mm", 0.0)),
            "transport_notes": str(overall.get("transport_notes") or ""),
        },
        "truck": truck_info,
        "material_cost_rows": material_cost_rows,
        "stock_summary_rows": stock_summary_rows,
        "purchased_count": len(purchased),
        "purchased_total_qty": sum(safe_quantity(x.get("quantity", 1)) for x in purchased),
        "machined_count": len(machined),
        "machined_total_qty": sum(safe_quantity(x.get("quantity", 1)) for x in machined),
    }


# ==============================
# PDF / 画像読み込み & 高速前処理
# ==============================
def get_pdf_page_count(file_path):
    """PDF のページ数を返す（画像ファイルの場合は 1）"""
    ext = os.path.splitext(file_path)[1].lower()
    if ext != ".pdf":
        return 1
    try:
        with fitz.open(file_path) as doc:
            return max(1, len(doc))
    except Exception:
        return 1


def load_preview_image(file_path, page_index=0, max_dim=1400):
    """ファイル選択・ページ切替時に即座に表示するための軽量プレビュー画像を生成する"""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".pdf":
        with fitz.open(file_path) as doc:
            idx = max(0, min(int(page_index), len(doc) - 1))
            page = doc.load_page(idx)
            long_side_pt = max(page.rect.width, page.rect.height, 1.0)
            dpi = max(96, min(200, int(max_dim * 72.0 / long_side_pt)))
            pix = page.get_pixmap(dpi=dpi)
        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    else:
        img = Image.open(file_path).convert("RGB")
        if max(img.width, img.height) > max_dim:
            img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
    return img


def prepare_drawing_inputs(file_path, page_index=0, status_callback=None):
    """
    PDFまたは画像ファイルから Visual Grounding 画像・API送信用画像・抽出テキストを作成する。
    - 大型CAD図面(A0/A1等)でもメモリ消費と処理時間を抑えるため、長辺上限 3200px の適応的DPIを採用
    - テキスト座標抽出に PyMuPDF (C実装) を優先使用し、従来の pdfplumber 単体より大幅に高速化
    """
    ext = os.path.splitext(file_path)[1].lower()
    text_pdf = ""

    if ext == ".pdf":
        with fitz.open(file_path) as doc:
            idx = max(0, min(int(page_index), len(doc) - 1))
            page = doc.load_page(idx)
            long_side_pt = max(page.rect.width, page.rect.height, 1.0)
            render_dpi = max(150, min(350, int(3200 * 72.0 / long_side_pt)))
            pix = page.get_pixmap(dpi=render_dpi)
            fitz_words = page.get_text("words")

        raw_image = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
        annotated_img_cv = cv2.cvtColor(np.array(raw_image), cv2.COLOR_RGB2BGR)
        scale_factor = render_dpi / 72.0

        if status_callback:
            status_callback("ベクターデータを解析・グラウンディング画像生成中...")

        words_normalized = []
        if fitz_words:
            for w in fitz_words:
                txt = str(w[4]).strip()
                if txt:
                    words_normalized.append((float(w[0]), float(w[1]), float(w[2]), float(w[3]), txt))
        else:
            try:
                with pdfplumber.open(file_path) as pdf:
                    if idx < len(pdf.pages):
                        for w in pdf.pages[idx].extract_words():
                            txt = str(w.get("text", "")).strip()
                            if txt:
                                words_normalized.append(
                                    (
                                        float(w["x0"]),
                                        float(w["top"]),
                                        float(w["x1"]),
                                        float(w["bottom"]),
                                        txt,
                                    )
                                )
            except Exception:
                pass

        img_h, img_w = annotated_img_cv.shape[:2]
        base_dim = max(img_h, img_w)
        line_thickness = max(1, int(base_dim / 1500))
        font_scale = max(0.4, base_dim / 2500)
        text_thickness = max(1, int(font_scale * 1.5))

        text_data_list = []
        for w_idx, (wx0, wy0, wx1, wy1, text) in enumerate(words_normalized):
            text_data_list.append(f"[ID:{w_idx}] X:{wx0:.1f}, Y:{wy0:.1f} TEXT: {text}")
            x0 = int(wx0 * scale_factor)
            y0 = int(wy0 * scale_factor)
            x1 = int(wx1 * scale_factor)
            y1 = int(wy1 * scale_factor)

            cv2.rectangle(annotated_img_cv, (x0, y0), (x1, y1), (0, 0, 255), line_thickness)
            cv2.putText(
                annotated_img_cv,
                str(w_idx),
                (x0, max(line_thickness + 2, y0 - 3)),
                cv2.FONT_HERSHEY_SIMPLEX,
                font_scale,
                (255, 0, 0),
                text_thickness,
            )

        text_pdf = "\n".join(text_data_list)
        processed_image_pil = Image.fromarray(cv2.cvtColor(annotated_img_cv, cv2.COLOR_BGR2RGB))
    else:
        processed_image_pil = Image.open(file_path).convert("RGB")

    api_image = processed_image_pil.copy()
    api_image.thumbnail((2000, 2000), Image.Resampling.LANCZOS)
    return processed_image_pil, api_image, text_pdf


# ==============================
# エンジン設定解決ヘルパー
# ==============================
def _resolve_runtime_config(settings):
    engine = settings.get("engine", "gemini")
    if engine == "azure":
        base_endpoint, dep1, api_version = parse_azure_endpoint_info(
            settings.get("azure_endpoint", ""),
            settings.get("azure_deployment_step1", ""),
            settings.get("azure_api_version", DEFAULT_AZURE_API_VERSION),
        )
        dep2 = settings.get("azure_deployment_step2", "").strip() or dep1
        dep3 = settings.get("azure_deployment_step3", "").strip() or dep1
        rpm = safe_int_param(settings.get("azure_rpm", 60), 60, 1, 2000)
        threads = safe_int_param(settings.get("azure_threads", 3), 3, 1, 20)
        return {
            "provider": "azure",
            "engine_label": "Azure OpenAI",
            "api_key": settings.get("azure_api_key", "").strip(),
            "base_endpoint": base_endpoint,
            "api_version": api_version,
            "target_step1": dep1,
            "target_step2": dep2,
            "target_step3": dep3,
            "rpm_step1": rpm,
            "rpm_step2": rpm,
            "rpm_step3": rpm,
            "threads_step2": threads,
            "threads_step3": threads,
            "temperature": safe_float_param(settings.get("azure_temp", 0.0), 0.0, 0.0, 2.0),
            "max_tokens": safe_int_param(settings.get("azure_tokens", 16384), 16384, 1024, 128000),
            "disable_safety": True,
            "custom_prompts": "\n".join(settings.get("azure_prompts", [])),
            "rate_limiter": RateLimiter(min_interval=0.5),
            "image_data_url": None,
        }

    plan = settings.get("plan", "free")
    is_free = plan == "free"
    prefix = f"{plan}_"
    api_key = settings.get(f"{prefix}key", "").strip()
    target_step1 = settings.get(f"{prefix}model_step1", DEFAULT_GEMINI_MODEL_STEP1)
    target_step2 = settings.get(f"{prefix}model_step2", DEFAULT_GEMINI_MODEL_STEP2)
    target_step3 = settings.get(f"{prefix}model_step3", DEFAULT_GEMINI_MODEL_STEP3)
    model_perf_dict = settings.get(f"{prefix}model_perf", {}) or {}
    model_meta = settings.get("gemini_model_meta", {}) or {}

    def _get_perf(m_id):
        perf = model_perf_dict.get(m_id, {})
        rec = get_recommended_gemini_settings(m_id, is_free=is_free, model_meta=model_meta)
        return (
            safe_int_param(perf.get("rpm", rec["rpm"]), rec["rpm"], 1, 2000),
            safe_int_param(perf.get("threads", rec["threads"]), rec["threads"], 1, 20),
        )

    rpm1, _ = _get_perf(target_step1)
    rpm2, thr2 = _get_perf(target_step2)
    rpm3, thr3 = _get_perf(target_step3)
    genai.configure(api_key=api_key)

    return {
        "provider": "gemini",
        "engine_label": "Gemini",
        "api_key": api_key,
        "base_endpoint": "",
        "api_version": DEFAULT_AZURE_API_VERSION,
        "target_step1": target_step1,
        "target_step2": target_step2,
        "target_step3": target_step3,
        "rpm_step1": rpm1,
        "rpm_step2": rpm2,
        "rpm_step3": rpm3,
        "threads_step2": thr2,
        "threads_step3": thr3,
        "temperature": safe_float_param(settings.get(f"{prefix}temp", 0.0), 0.0, 0.0, 2.0),
        "max_tokens": safe_int_param(settings.get(f"{prefix}tokens", 65535), 65535, 1024, 65535),
        "disable_safety": bool(settings.get(f"{prefix}safety", True)),
        "custom_prompts": "\n".join(settings.get(f"{prefix}prompts", [])),
        "rate_limiter": RateLimiter(min_interval=1.5),
        "image_data_url": None,
    }


def _check_cancel(cancel_event):
    if cancel_event is not None and cancel_event.is_set():
        raise AnalysisCancelledError("ユーザーによって解析が中止されました。")


def _call_model(
    cfg, prompt, api_image, response_schema, step_num, status_callback=None, cancel_event=None
):
    _check_cancel(cancel_event)
    target = cfg[f"target_step{step_num}"]
    rpm = cfg[f"rpm_step{step_num}"]
    return generate_structured_json(
        provider=cfg["provider"],
        prompt=prompt,
        pil_img=api_image,
        response_schema=response_schema,
        api_key=cfg["api_key"],
        model_name=target,
        disable_safety=cfg["disable_safety"],
        azure_endpoint=cfg["base_endpoint"],
        azure_deployment=target,
        azure_api_version=cfg["api_version"],
        temperature=cfg["temperature"],
        max_tokens=cfg["max_tokens"],
        rpm=rpm,
        rate_limiter=cfg["rate_limiter"],
        status_callback=status_callback,
        cancel_event=cancel_event,
        image_data_url=cfg.get("image_data_url"),
    )


# ==============================
# 解析パイプライン本体
# ==============================
def _run_volume_step1_and_step2(
    cfg, api_image, text_pdf, density, status_callback=None, cancel_event=None
):
    """体積・重量解析の Step1 (部品一覧・数量・全体外形寸法) と Step2 (並列要素抽出) を実行する"""
    def _status(msg):
        if status_callback:
            status_callback(msg)

    custom_prompts = cfg["custom_prompts"]
    engine_label = cfg["engine_label"]

    _check_cancel(cancel_event)
    _status(f"[体積解析 1/3] 部品一覧・数量・全体外形寸法を抽出中 ({engine_label} / {cfg['target_step1']})...")
    step1_prompt = f"""
    あなたは2D機械図面解析・見積り・搬送計画のエキスパートです。
    提供された図面画像とアノテーションIDテキストから、以下をすべて正確に抽出してください。
    1. 製品・ユニット全体の最大外形寸法（長さ overall_length_mm, 幅 overall_width_mm, 高さ overall_height_mm）と搬送上の注意点（transport_notes）。
    2. 図面内の「すべての部品の一覧」（品番、部品名、材質、比重、および図面内での必要数量 quantity）。
    
    【重要な前提条件（Visual Grounding & 数量読取）】
    - 画像には、テキストが抽出された正確な位置を示す「赤い枠」と「青い文字のID番号」が描画されています。
    - 部品表（BOM）の「数量」「個数」「QTY」列や、図中の「2ヶ」「×4」などの注記から、各部品の正確な数量（quantity）を必ず読み取ってください。
    - デフォルト比重: {density} g/cm³ （材質記載がない場合に適用）
    
    テキストデータ:
    {text_pdf}
    
    ユーザーからの追加指示: {custom_prompts}
    """

    resp1_text = _call_model(
        cfg, step1_prompt, api_image, PartListResult, 1, status_callback, cancel_event
    )
    part_list_data = parse_json_response(resp1_text, status_callback)
    if not part_list_data:
        raise Exception("体積解析(ステップ1): JSONの解析に失敗しました。")

    overall_dims = {
        "length_mm": safe_float(part_list_data.get("overall_length_mm", 0.0)),
        "width_mm": safe_float(part_list_data.get("overall_width_mm", 0.0)),
        "height_mm": safe_float(part_list_data.get("overall_height_mm", 0.0)),
        "transport_notes": str(part_list_data.get("transport_notes") or ""),
    }

    extracted_parts_basic = part_list_data.get("parts", [])
    final_parts = []
    fatal_error_box = {"err": None}

    def process_part(idx, p_basic):
        if fatal_error_box["err"] is not None:
            return None
        _check_cancel(cancel_event)

        p_name = p_basic.get("part_name", "不明")
        p_num = p_basic.get("part_number", "")
        _status(
            f"[体積解析 2/3] 部品詳細(三面図)を解析中... ({idx + 1}/{len(extracted_parts_basic)}) {p_name}"
        )

        step2_prompt = f"""
        あなたは機械図面の解析エキスパートです。
        提供された図面画像とテキストデータから、指定された特定の部品「1個あたり」の構成要素と三面図から導き出される3次元寸法を漏れなく抽出し、体積計算式を作成してください。
        （※数量掛けはプログラム側で行うため、ここでは必ず1個あたりの体積計算式を出力してください）
        
        【対象部品】
        品番: {p_num}
        品名: {p_name}
        
        【三面図からの立体形状復元と寸法マッピングルール（極めて重要）】
        1. 対象部品を構成する要素（本体ベース、補強リブ、ボス、貫通穴、ザグリ穴、切り欠きなど）に分解する。
        2. 図面の「線種」を正確に読み取り、立体構造に反映すること。
           - 実線（太線）: 外形線（実際の材料が存在する輪郭）。
           - 破線（隠れ線）: 内部空洞や裏側の形状（穴、ザグリ、スロットなど、体積を『引き算（マイナス）』すべき領域）。
           - 一点鎖線（中心線）: 円柱・円形穴の中心、対称軸。
        3. 正面図・平面図・側面図の対応関係から統合寸法（W x H x D）を正確に特定すること。
        4. 「calculation_formula」は純粋な算術式（数値、小数点、`+,-,*,/`、括弧 `()`、べき乗 `**` のみ）とし、円周率は `3.14159265`、穴・切り欠き等の除去部は先頭にマイナスをつけた式にすること。
        
        テキストデータ:
        {text_pdf}
        
        ユーザーからの追加指示: {custom_prompts}
        """

        elements = []
        try:
            resp2_text = _call_model(
                cfg, step2_prompt, api_image, PartElementsResult, 2, status_callback, cancel_event
            )
            elements_data = parse_json_response(resp2_text, status_callback)
            elements = elements_data.get("elements", []) if elements_data else []
        except AnalysisCancelledError:
            raise
        except DiagnosedAPIError as api_err:
            if api_err.fatal:
                fatal_error_box["err"] = api_err
                raise api_err
            print(f"API Error on part {p_name}: {api_err}")
        except Exception as e:
            print(f"API Error on part {p_name}: {e}")

        return {
            "part_number": p_basic.get("part_number", ""),
            "part_name": p_basic.get("part_name", ""),
            "material": p_basic.get("material", ""),
            "density_g_cm3": p_basic.get("density_g_cm3", 0.0),
            "quantity": safe_quantity(p_basic.get("quantity", 1)),
            "elements": elements,
        }

    max_workers = max(1, cfg["threads_step2"])
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [
            executor.submit(process_part, idx, p)
            for idx, p in enumerate(extracted_parts_basic)
        ]
        for future in futures:
            _check_cancel(cancel_event)
            result = future.result()
            if result is not None:
                final_parts.append(result)

    if fatal_error_box["err"] is not None:
        raise fatal_error_box["err"]

    return final_parts, overall_dims


def _run_volume_step3_verification(
    cfg, api_image, text_pdf, density, existing_parts, status_callback=None, cancel_event=None
):
    """体積・重量解析の Step3 (抽出漏れ最終検証) を実行し、追加すべき不足部品リストを返す"""
    _check_cancel(cancel_event)
    if status_callback:
        status_callback(
            f"[体積解析 3/3] 三面図からの抽出漏れがないか再検証中 ({cfg['engine_label']} / {cfg['target_step3']})..."
        )
    extracted_summary = "\n".join(
        [
            f"- 品番:{p['part_number']} 品名:{p['part_name']} 数量:{p.get('quantity', 1)}"
            for p in existing_parts
        ]
    )
    step3_prompt = f"""
    あなたは機械図面の解析エキスパートです。最終チェックを行います。
    提供された図面画像およびアノテーションIDテキストと、これまでに抽出した以下の部品リストを比較し、「まだ抽出されていない（見落としている）部品」が存在しないか厳密に検証してください。
    
    【すでに抽出済みの部品リスト】
    {extracted_summary}
    
    見落としている部品や非表示線（破線等）で示されたサブアセンブリ等があれば、その部品情報（品番、品名、材質、比重、数量 quantity）と、1個あたりの構成要素（寸法、計算式等）を抽出してください。
    見落としが一切無い場合は、「missing_parts」を空リスト（[]）として返却してください。
    
    【前提条件・数式ルール】
    - デフォルト比重: {density} g/cm³ （材質記載がない場合に適用）
    - 数式（calculation_formula）は文字や単位を含めず、純粋な数値と `+,-,*,/,(),**` のみで構築してください。円周率は `3.14159265` を使用すること。
    
    テキストデータ:
    {text_pdf}
    
    ユーザーからの追加指示: {cfg['custom_prompts']}
    """

    resp3_text = _call_model(
        cfg, step3_prompt, api_image, MissingPartsResult, 3, status_callback, cancel_event
    )
    verification_data = parse_json_response(resp3_text, status_callback)
    if verification_data and "missing_parts" in verification_data:
        return verification_data["missing_parts"] or []
    return []


def _run_parts_breakdown_analysis(
    cfg, api_image, text_pdf, targets, known_parts=None, status_callback=None, cancel_event=None
):
    """選択された分類項目（部材表・切断寸法 / 使用購入品 / 使用機械加工品）を高精度抽出する"""
    want_mat = bool(targets.get("materials"))
    want_pur = bool(targets.get("purchased_parts"))
    want_mac = bool(targets.get("machined_parts"))

    selected_labels = []
    if want_mat:
        selected_labels.append("部材表(切断寸法)")
    if want_pur:
        selected_labels.append("使用購入品(手配)")
    if want_mac:
        selected_labels.append("使用機械加工品")

    if not selected_labels:
        return [], [], [], {}

    _check_cancel(cancel_event)
    if status_callback:
        status_callback(
            f"[手配・部材表解析] {'・'.join(selected_labels)} を抽出中 ({cfg['engine_label']} / {cfg['target_step3']})..."
        )

    selected_count = sum([want_mat, want_pur, want_mac])
    if selected_count == 1:
        if want_mat:
            schema = MaterialListResult
        elif want_pur:
            schema = PurchasedPartListResult
        else:
            schema = MachinedPartListResult
    else:
        schema = PartsBreakdownResult

    instructions = []
    if want_mat:
        instructions.append(
            "【1. 部材表・切断寸法（鋼材・板材・形鋼・パイプ・棒材・樹脂等の木取り・切断リスト）】\n"
            "- 工場や鋼材業者へそのまま手配・切断指示できる「部材表」を作成してください。\n"
            "- 各部材について、鋼材種別（鋼板PL, 平鋼FB, 等辺山形鋼L鋼, 溝形鋼C鋼, H形鋼, 角パイプSTKR, 丸パイプ, 丸棒等）、材質（SS400, SUS304等）、断面規格・板厚（例: PL10, L-50x50x6, □-50x50x2.3）、具体的な切断寸法（例: '150x300' や 'L=1250'）、長尺部材の1本あたり切断長さ数値 cut_length_mm(mm)、端部切断形状（90°直角 / 両端45°留め切り等）、必要数量 quantity（整数）、1本/1枚あたりの体積計算式 unit_volume_formula(mm³)、比重 density_g_cm3、表面処理を漏れなく記載してください。\n"
            "- さらに図面全体の最大外形寸法（overall_length_mm, overall_width_mm, overall_height_mm）と搬送上の注意点（transport_notes）も読み取ってください。"
        )
    if want_pur:
        instructions.append(
            "【2. 使用購入品（発注・手配用リスト）】\n"
            "- ボルト、ナット、座金、ピン、ベアリング、オイルシール、Oリング、ばね、シリンダ、モーター、リニアガイド、継手、電気部品などの市販品・JIS規格品をすべて抽出してください。\n"
            "- 購買手配にそのまま使えるよう、メーカー名、正確な型式・呼び径、材質・仕様、総必要数量 quantity（整数）、手配単位 unit（個/本/組/m等）、手配区分 procurement_category（JIS規格品/標準市販品/メーカー取寄品等）を漏れなく読み取ってください。"
        )
    if want_mac:
        instructions.append(
            "【3. 使用機械加工品（製作・外注手配リスト）】\n"
            "- 市販の購入品ではなく、素材から切削・旋盤・フライス・マシニング・穴あけ・タップ・ザグリ・研削・ワイヤーカット・溶接後機械加工などを行って製作する部品を抽出してください。\n"
            "- 材質、素材・概略外形寸法、加工工程・内容、公差・はめあい・表面粗さ、熱処理・表面処理（heat_surface_treatment）、製作数量 quantity（整数）を詳しく記載してください。"
        )

    exclude_notes = []
    if selected_count > 1:
        if not want_mat:
            exclude_notes.append("- `materials` は対象外のため必ず空リスト `[]` にしてください。")
        if not want_pur:
            exclude_notes.append("- `purchased_parts` は対象外のため必ず空リスト `[]` にしてください。")
        if not want_mac:
            exclude_notes.append("- `machined_parts` は対象外のため必ず空リスト `[]` にしてください。")

    known_parts_block = ""
    if known_parts:
        lines = [
            f"- 品番:{p.get('part_number', '')} 品名:{p.get('part_name', '')} 材質:{p.get('material', '')} 数量:{p.get('quantity', 1)}"
            for p in known_parts
        ]
        known_parts_block = (
            "\n【参考: すでに認識された部品一覧（品番・品名・数量の整合性を可能な限り揃えてください）】\n"
            + "\n".join(lines)
            + "\n"
        )

    prompt = f"""
    あなたは機械設計・製缶・生産技術・部品購買のエキスパートです。
    提供された2D機械図面画像（赤枠と青ID番号によるVisual Grounding付き）およびテキストデータを精査し、手配・見積り・切断指示にすぐ活用できる精度で指定カテゴリの情報を抽出してください。

    【抽出対象カテゴリとルール】
    {chr(10).join(instructions)}

    【分類の整合性ルール】
    - ボルト・ベアリング・シリンダ等の市販品・JIS規格品は「使用購入品」に分類し、「使用機械加工品」には重複計上しないでください。
    - 図面の部品表（BOM）、表題欄、注記、バルーン記号、寸法・切断長・公差・表面処理記号をすべて確認してください。
    - 数値や記号は半角英数字で見やすく統一してください。
    {chr(10).join(exclude_notes)}
    {known_parts_block}
    テキストデータ:
    {text_pdf}

    ユーザーからの追加指示: {cfg['custom_prompts']}
    """

    resp_text = _call_model(cfg, prompt, api_image, schema, 3, status_callback, cancel_event)
    parsed = parse_json_response(resp_text, status_callback)
    if not parsed:
        raise Exception(f"手配・部材表解析（{'・'.join(selected_labels)}）: JSONの解析に失敗しました。")

    materials = parsed.get("materials", []) if want_mat else []
    purchased = parsed.get("purchased_parts", []) if want_pur else []
    machined = parsed.get("machined_parts", []) if want_mac else []
    overall_dims = {
        "length_mm": safe_float(parsed.get("overall_length_mm", 0.0)),
        "width_mm": safe_float(parsed.get("overall_width_mm", 0.0)),
        "height_mm": safe_float(parsed.get("overall_height_mm", 0.0)),
        "transport_notes": str(parsed.get("transport_notes") or ""),
    }
    return materials, purchased, machined, overall_dims


def _merge_overall_dims(dim_a, dim_b):
    a = dim_a or {}
    b = dim_b or {}
    seen = []
    for raw in (a.get("transport_notes"), b.get("transport_notes")):
        s = str(raw or "").strip()
        if s and s not in seen:
            seen.append(s)
    return {
        "length_mm": max(safe_float(a.get("length_mm", 0.0)), safe_float(b.get("length_mm", 0.0))),
        "width_mm": max(safe_float(a.get("width_mm", 0.0)), safe_float(b.get("width_mm", 0.0))),
        "height_mm": max(safe_float(a.get("height_mm", 0.0)), safe_float(b.get("height_mm", 0.0))),
        "transport_notes": " ／ ".join(seen),
    }


def run_drawing_analysis(
    file_path,
    density,
    settings,
    targets,
    status_callback=None,
    preview_callback=None,
    page_index=0,
    cancel_event=None,
):
    """
    選択された解析項目 (targets) に基づき図面解析を実行して統合結果辞書を返す。
    """
    _check_cancel(cancel_event)
    cfg = _resolve_runtime_config(settings)
    processed_image_pil, api_image, text_pdf = prepare_drawing_inputs(
        file_path, page_index=page_index, status_callback=status_callback
    )

    if preview_callback:
        preview_callback(processed_image_pil)

    _check_cancel(cancel_event)

    if cfg["provider"] == "azure":
        cfg["image_data_url"] = encode_pil_image_data_url(api_image)

    want_volume = bool(targets.get("volume"))
    want_breakdown = any(
        targets.get(k) for k in ("materials", "purchased_parts", "machined_parts")
    )

    default_den = safe_float(density) or 7.85
    final_parts = []
    materials, purchased_parts, machined_parts = [], [], []
    overall_vol, overall_brk = {}, {}

    if want_volume and want_breakdown:
        final_parts, overall_vol = _run_volume_step1_and_step2(
            cfg, api_image, text_pdf, density, status_callback, cancel_event
        )
        _check_cancel(cancel_event)

        if cfg["threads_step3"] >= 2:
            if status_callback:
                status_callback("体積最終検証と部材表・手配リスト解析を並列実行中...")
            with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
                fut_verify = executor.submit(
                    _run_volume_step3_verification,
                    cfg,
                    api_image,
                    text_pdf,
                    density,
                    final_parts,
                    status_callback,
                    cancel_event,
                )
                fut_breakdown = executor.submit(
                    _run_parts_breakdown_analysis,
                    cfg,
                    api_image,
                    text_pdf,
                    targets,
                    final_parts if final_parts else None,
                    status_callback,
                    cancel_event,
                )
                missing_parts = fut_verify.result()
                materials, purchased_parts, machined_parts, overall_brk = fut_breakdown.result()
        else:
            missing_parts = _run_volume_step3_verification(
                cfg, api_image, text_pdf, density, final_parts, status_callback, cancel_event
            )
            materials, purchased_parts, machined_parts, overall_brk = _run_parts_breakdown_analysis(
                cfg,
                api_image,
                text_pdf,
                targets,
                final_parts if final_parts else None,
                status_callback,
                cancel_event,
            )

        if missing_parts:
            final_parts.extend(missing_parts)
        calculate_volume_and_weight(final_parts, default_density=default_den)

    elif want_volume:
        final_parts, overall_vol = _run_volume_step1_and_step2(
            cfg, api_image, text_pdf, density, status_callback, cancel_event
        )
        missing_parts = _run_volume_step3_verification(
            cfg, api_image, text_pdf, density, final_parts, status_callback, cancel_event
        )
        if missing_parts:
            final_parts.extend(missing_parts)
        calculate_volume_and_weight(final_parts, default_density=default_den)

    elif want_breakdown:
        materials, purchased_parts, machined_parts, overall_brk = _run_parts_breakdown_analysis(
            cfg, api_image, text_pdf, targets, None, status_callback, cancel_event
        )

    if materials:
        calculate_materials_metrics(materials, default_density=default_den)
    for item in purchased_parts:
        item["quantity"] = safe_quantity(item.get("quantity", 1))
    for item in machined_parts:
        item["quantity"] = safe_quantity(item.get("quantity", 1))

    overall_dims = _merge_overall_dims(overall_vol, overall_brk)

    return {
        "targets": dict(targets),
        "default_density": default_den,
        "overall_dimensions": overall_dims,
        "parts": final_parts,
        "materials": materials,
        "purchased_parts": purchased_parts,
        "machined_parts": machined_parts,
    }


# ==============================
# Excel / CSV 出力用テーブル生成
# ==============================
def build_export_tables(result_data, custom_prices=None, default_unit_price=250.0):
    """
    解析結果辞書からエクスポート可能なテーブルのリスト
    [(sheet_name, file_suffix, pandas.DataFrame), ...] を生成する。
    """
    if not result_data:
        return []

    tables = []
    default_den = safe_float(result_data.get("default_density", 7.85)) or 7.85
    summary = build_estimation_and_logistics_summary(
        result_data,
        custom_prices=custom_prices,
        default_unit_price=default_unit_price,
        default_density=default_den,
    )

    # 0. 見積り・搬送サマリー（重量・金額・トラック・定尺集計が存在する場合）
    summary_rows = []
    if summary and (
        summary.get("material_cost_rows")
        or summary.get("stock_summary_rows")
        or summary.get("total_weight_kg", 0) > 0
    ):
        truck = summary["truck"]
        od = summary["overall_dimensions"]
        summary_rows.append(
            {
                "区分": "搬送・トラック選定",
                "項目": "推奨トラック車種",
                "値 / 内容": truck["recommended_truck"],
                "補足 / 根拠": f"{truck['bed_spec']} ／ {truck['reasons']}",
            }
        )
        summary_rows.append(
            {
                "区分": "搬送・トラック選定",
                "項目": "荷役・積降ろし目安",
                "値 / 内容": truck["handling_advice"],
                "補足 / 根拠": od.get("transport_notes", ""),
            }
        )
        summary_rows.append(
            {
                "区分": "全体寸法・総重量",
                "項目": "製品全体外形 / 総重量",
                "値 / 内容": f"L{od['length_mm']:,.0f} × W{od['width_mm']:,.0f} × H{od['height_mm']:,.0f} mm ／ {summary['total_weight_kg']:,.2f} kg",
                "補足 / 根拠": f"重量算出元: {summary['weight_source']}",
            }
        )
        summary_rows.append(
            {
                "区分": "概算材料費",
                "項目": "概算材料費 合計",
                "値 / 内容": f"{int(round(summary['total_material_cost_jpy'])):,} 円",
                "補足 / 根拠": "材質別重量 × 適用単価(円/kg)の合計目安",
            }
        )
        for mc in summary.get("material_cost_rows", []):
            summary_rows.append(
                {
                    "区分": "材質別集計",
                    "項目": mc["material"],
                    "値 / 内容": f"{mc['weight_kg']:,.2f} kg ／ {int(round(mc['estimated_cost_jpy'])):,} 円",
                    "補足 / 根拠": f"比重 {mc['density_g_cm3']:.2f} ／ 単価 {mc['unit_price_jpy']:,.0f} 円/kg",
                }
            )
        for st_row in summary.get("stock_summary_rows", []):
            summary_rows.append(
                {
                    "区分": "鋼材・定尺発注集計",
                    "項目": f"{st_row['material_type']} / {st_row['material_spec']} / {st_row['section_size']}",
                    "値 / 内容": f"計{st_row['total_pieces']}本(枚) ／ 総長 {st_row['total_length_mm']:,.0f} mm ／ {st_row['total_weight_kg']:,.2f} kg",
                    "補足 / 根拠": st_row["stock_guide"],
                }
            )
    if summary_rows:
        tables.append(("見積り・搬送サマリー", "見積り搬送サマリー", pd.DataFrame(summary_rows)))

    # 1. 体積・重量解析
    vol_rows = []
    for part in result_data.get("parts", []):
        p_num = str(part.get("part_number") or "")
        p_name = str(part.get("part_name") or "不明な部品")
        p_mat = str(part.get("material") or "")
        p_den = safe_float(part.get("density_g_cm3"))
        p_qty = safe_quantity(part.get("quantity", 1))
        p_unit_wt = safe_float(part.get("unit_weight_kg", 0.0))
        p_sub_wt = safe_float(part.get("calculated_weight_kg", 0.0))

        elements = part.get("elements", [])
        if elements:
            for elem in elements:
                e_name = str(elem.get("element_name") or "")
                e_dim = str(elem.get("dimensions") or "")
                e_calc = str(elem.get("calculation_formula") or "")
                e_vol = elem.get("calculated_volume_mm3", 0.0)
                e_weight = elem.get("calculated_weight_kg", 0.0)
                e_notes = str(elem.get("notes") or "")
                combined_notes = f"[式: {e_calc}] {e_notes}".strip() if e_calc else e_notes

                vol_rows.append(
                    {
                        "品番": p_num,
                        "部品名": p_name,
                        "材質": p_mat,
                        "比重 (g/cm³)": p_den,
                        "数量": p_qty,
                        "要素名": e_name,
                        "読み取り寸法": e_dim,
                        "1個あたり体積 (mm³)": e_vol,
                        "1個あたり重量 (kg)": e_weight,
                        "小計重量 (kg)": e_weight * p_qty,
                        "備考 (計算式等)": combined_notes,
                    }
                )
        else:
            vol_rows.append(
                {
                    "品番": p_num,
                    "部品名": p_name,
                    "材質": p_mat,
                    "比重 (g/cm³)": p_den,
                    "数量": p_qty,
                    "要素名": "",
                    "読み取り寸法": "",
                    "1個あたり体積 (mm³)": part.get("unit_volume_mm3", 0.0),
                    "1個あたり重量 (kg)": p_unit_wt,
                    "小計重量 (kg)": p_sub_wt,
                    "備考 (計算式等)": "",
                }
            )
    if vol_rows:
        tables.append(("体積・重量解析", "体積重量", pd.DataFrame(vol_rows)))

    # 2. 部材表 (切断寸法)
    mat_rows = []
    for item in result_data.get("materials", []):
        mat_rows.append(
            {
                "品番": str(item.get("part_number") or ""),
                "部品名・部位": str(item.get("part_name") or ""),
                "部材種別": str(item.get("material_type") or ""),
                "材質": str(item.get("material_spec") or ""),
                "断面規格・板厚": str(item.get("section_size") or ""),
                "切断寸法 (mm)": str(item.get("cut_dimensions") or ""),
                "切断長さ (mm)": safe_float(item.get("cut_length_mm", 0.0)),
                "端部切断": str(item.get("cut_angle_end") or ""),
                "数量": safe_quantity(item.get("quantity", 1)),
                "合計長さ (mm)": safe_float(item.get("total_length_mm", 0.0)),
                "小計重量 (kg)": round(safe_float(item.get("total_weight_kg", 0.0)), 3),
                "表面処理": str(item.get("surface_treatment") or ""),
                "備考": str(item.get("notes") or ""),
            }
        )
    if mat_rows:
        tables.append(("部材表_切断寸法", "部材表", pd.DataFrame(mat_rows)))

    # 3. 使用購入品 (手配)
    pur_rows = []
    for item in result_data.get("purchased_parts", []):
        pur_rows.append(
            {
                "品番": str(item.get("part_number") or ""),
                "品名": str(item.get("part_name") or ""),
                "メーカー": str(item.get("manufacturer") or ""),
                "型式・規格": str(item.get("model_number") or ""),
                "材質・仕様": str(item.get("material_or_spec") or ""),
                "数量": safe_quantity(item.get("quantity", 1)),
                "単位": str(item.get("unit") or "個"),
                "手配区分": str(item.get("procurement_category") or ""),
                "備考": str(item.get("notes") or ""),
            }
        )
    if pur_rows:
        tables.append(("使用購入品_手配", "使用購入品", pd.DataFrame(pur_rows)))

    # 4. 使用機械加工品
    mac_rows = []
    for item in result_data.get("machined_parts", []):
        mac_rows.append(
            {
                "品番": str(item.get("part_number") or ""),
                "部品名": str(item.get("part_name") or ""),
                "材質": str(item.get("material") or ""),
                "素材・概略寸法": str(item.get("blank_or_size") or ""),
                "加工工程・内容": str(item.get("machining_processes") or ""),
                "公差・粗さ": str(item.get("precision_surface") or ""),
                "熱処理・表面処理": str(item.get("heat_surface_treatment") or ""),
                "数量": safe_quantity(item.get("quantity", 1)),
                "備考": str(item.get("notes") or ""),
            }
        )
    if mac_rows:
        tables.append(("使用機械加工品", "使用機械加工品", pd.DataFrame(mac_rows)))

    return tables
