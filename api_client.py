# -*- coding: utf-8 -*-
"""
AI API 通信・モデル管理・レート制御・エラー診断モジュール
Gemini API (無料枠 / 課金枠) および Azure OpenAI Service に対応
"""

import base64
import io
import json
import random
import re
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import warnings

warnings.filterwarnings("ignore", category=FutureWarning)
import google.generativeai as genai

try:
    from openai import AzureOpenAI, OpenAI
except ImportError:
    AzureOpenAI = None
    OpenAI = None


# ==============================
# 定数定義
# ==============================
APP_NAME = "DrawingAnalysisMiya"
APP_VERSION = "v2.0.0"

DEFAULT_AZURE_API_VERSION = "v1"
AZURE_API_VERSION_CANDIDATES = [
    "v1",
    "2025-04-01-preview",
    "2025-01-01-preview",
    "2024-12-01-preview",
    "2024-10-21",
    "2024-08-01-preview",
    "2024-06-01",
]

DEFAULT_GEMINI_MODELS = [
    ("Gemini 4 Hydrogen", "gemini-4-hydrogen"),
    ("Gemini 4 Argon", "gemini-4-argon"),
    ("Gemini 4 Neon", "gemini-4-neon"),
    ("Gemini 4 Carbon", "gemini-4-carbon"),
    ("Gemini 3.8 Flash", "gemini-3.8-flash"),
    ("Gemini 3.5 Flash-Lite", "gemini-3.5-flash-lite"),
]

DEFAULT_AZURE_MODELS = [
    "gpt-5.2",
    "gpt-5.1",
    "gpt-5",
    "gpt-5-mini",
    "gpt-5-nano",
    "gpt-5-chat",
    "o4-mini",
]

DEFAULT_GEMINI_MODEL_STEP1 = "gemini-4-hydrogen"
DEFAULT_GEMINI_MODEL_STEP2 = "gemini-4-hydrogen"
DEFAULT_GEMINI_MODEL_STEP3 = "gemini-4-neon"

DEFAULT_AZURE_DEPLOYMENT_STEP1 = "gpt-5-mini"
DEFAULT_AZURE_DEPLOYMENT_STEP2 = "gpt-5-mini"
DEFAULT_AZURE_DEPLOYMENT_STEP3 = "gpt-5"


# ==============================
# 入力値バリデーション・URL正規化
# ==============================
def safe_int(val, default, min_val=None, max_val=None):
    """任意の値（文字列・数値・Tk変数）を安全に整数へ変換して範囲内に収める"""
    try:
        if hasattr(val, "get"):
            val = val.get()
        num = int(float(str(val).strip()))
    except Exception:
        num = default
    if min_val is not None:
        num = max(min_val, num)
    if max_val is not None:
        num = min(max_val, num)
    return num


def safe_float(val, default, min_val=None, max_val=None):
    """任意の値（文字列・数値・Tk変数）を安全に小数へ変換して範囲内に収める"""
    try:
        if hasattr(val, "get"):
            val = val.get()
        num = float(str(val).strip())
    except Exception:
        num = default
    if min_val is not None:
        num = max(min_val, num)
    if max_val is not None:
        num = min(max_val, num)
    return num


def parse_azure_endpoint_info(endpoint_str, deployment_str="", api_version_str=""):
    """
    ユーザーが入力した Azure エンドポイント URL を解析・正規化する。
    ベース URL（https://xxx.openai.azure.com/）だけでなく、
    誤ってターゲット URI 全体
    （https://xxx.openai.azure.com/openai/deployments/my-dep/chat/completions?api-version=...）
    が貼り付けられた場合でも、自動的にベース URL・デプロイ名・API バージョンを抽出して補完する。
    """
    raw_ep = (endpoint_str or "").strip()
    dep = (deployment_str or "").strip()
    ver = (api_version_str or "").strip() or DEFAULT_AZURE_API_VERSION

    if not raw_ep:
        return "", dep, ver

    if not raw_ep.startswith(("http://", "https://")):
        raw_ep = "https://" + raw_ep

    try:
        parsed = urllib.parse.urlparse(raw_ep)
        if not parsed.netloc:
            return raw_ep, dep, ver
        base_endpoint = f"{parsed.scheme}://{parsed.netloc}/"

        # パスに /deployments/{deployment_name} が含まれている場合、デプロイ名を抽出
        m_dep = re.search(r"/deployments/([^/]+)", parsed.path)
        if m_dep and not dep:
            dep = urllib.parse.unquote(m_dep.group(1))

        # クエリに api-version が含まれている場合、抽出
        qs = urllib.parse.parse_qs(parsed.query)
        if "api-version" in qs and qs["api-version"]:
            if not api_version_str or api_version_str.strip() == "":
                ver = qs["api-version"][0]
        elif "/openai/v1" in parsed.path.lower():
            if not api_version_str or api_version_str.strip() == "":
                ver = "v1"

        return base_endpoint, dep, ver
    except Exception:
        return raw_ep, dep, ver


# ==============================
# 推奨設定・モデルフィルタリング
# ==============================
def get_recommended_gemini_settings(model_id, is_free=True, model_meta=None):
    """Gemini モデル・プラン・API取得メタ情報に応じた最適化済み推奨設定を返す"""
    mid = (model_id or "").lower()
    if any(k in mid for k in ["carbon", "pro", "ultra"]):
        rpm = 1 if is_free else 100
        threads = 1 if is_free else 3
    elif "neon" in mid:
        rpm = 2 if is_free else 150
        threads = 1 if is_free else 5
    elif "argon" in mid:
        rpm = 5 if is_free else 200
        threads = 1 if is_free else 5
    elif "lite" in mid:
        rpm = 15 if is_free else 400
        threads = 1 if is_free else 8
    else:
        rpm = 12 if is_free else 300
        threads = 1 if is_free else 5

    max_tokens = 65535
    if isinstance(model_meta, dict):
        meta = model_meta.get(model_id) or model_meta.get(mid)
        if isinstance(meta, dict) and meta.get("output_token_limit"):
            try:
                out_lim = int(meta["output_token_limit"])
                if out_lim >= 1024:
                    max_tokens = min(out_lim, 65535)
            except Exception:
                pass

    return {
        "rpm": rpm,
        "threads": threads,
        "temperature": 0.0,
        "max_tokens": max_tokens,
        "disable_safety": True,
    }


def get_optimized_gemini_params_for_steps(step_models, is_free=True, model_meta=None):
    """Step1〜3で選択中の全モデルに対して安全かつ最適な抽出パラメータを算出する"""
    valid_models = [m.strip() for m in (step_models or []) if str(m).strip()]
    if not valid_models:
        valid_models = [DEFAULT_GEMINI_MODEL_STEP1]

    token_limits = []
    for mid in valid_models:
        rec = get_recommended_gemini_settings(mid, is_free=is_free, model_meta=model_meta)
        token_limits.append(rec["max_tokens"])

    return {
        "temperature": 0.0,
        "max_tokens": min(token_limits) if token_limits else 65535,
        "disable_safety": True,
    }


def pick_recommended_gemini_step_models(models_list):
    """取得済みモデル一覧から Step1/2/3 に最適な推奨モデルIDを選定する"""
    visible = filter_mainstream_gemini_models(models_list) if models_list else list(DEFAULT_GEMINI_MODELS)
    ids = [m[1] for m in visible] if visible else [m[1] for m in DEFAULT_GEMINI_MODELS]

    def _find_first(predicate):
        for mid in ids:
            if predicate(mid.lower()):
                return mid
        return None

    # Step1/2: 高速・バランス重視 (hydrogen / flash)
    fast_model = (
        _find_first(lambda m: "hydrogen" in m)
        or _find_first(lambda m: "flash" in m and "lite" not in m)
        or _find_first(lambda m: "argon" in m)
        or (ids[0] if ids else DEFAULT_GEMINI_MODEL_STEP1)
    )
    # Step3: 高度推論・見落とし検証重視 (neon / pro / carbon / argon)
    verify_model = (
        _find_first(lambda m: "neon" in m)
        or _find_first(lambda m: "pro" in m)
        or _find_first(lambda m: "carbon" in m)
        or _find_first(lambda m: "argon" in m)
        or fast_model
    )
    return fast_model, fast_model, verify_model


def get_default_gemini_perf_map(models_list=None, is_free=True, model_meta=None):
    """モデル一覧に対するデフォルトのモデル別パフォーマンス辞書を生成する"""
    target_models = models_list if models_list else DEFAULT_GEMINI_MODELS
    perf = {}
    for _, m_id in target_models:
        rec = get_recommended_gemini_settings(m_id, is_free=is_free, model_meta=model_meta)
        perf[m_id] = {"rpm": rec["rpm"], "threads": rec["threads"]}
    return perf


def get_recommended_azure_settings(deployment_name, detected_api_version=None, model_meta=None):
    """Azure OpenAI デプロイ名（モデル種別）とメタ情報に応じた最適化済み推奨設定を返す"""
    dep = (deployment_name or "").lower()
    api_ver = (detected_api_version or "").strip() or DEFAULT_AZURE_API_VERSION
    if any(k in dep for k in ["mini", "nano"]) and not any(
        dep.startswith(p) for p in ["o1", "o3", "o4"]
    ):
        rpm = 120
        threads = 5
        max_tokens = 16384
    elif (
        any(dep.startswith(p) or f"-{p}" in dep for p in ["gpt-6", "o4", "o3", "o1"])
        or "pro" in dep
        or "sol" in dep
    ):
        rpm = 30
        threads = 2
        max_tokens = 16384
    elif dep.startswith("gpt-35") or dep == "gpt-4":
        rpm = 60
        threads = 3
        max_tokens = 4096
    else:
        rpm = 60
        threads = 3
        max_tokens = 16384

    if isinstance(model_meta, dict):
        meta = model_meta.get(deployment_name) or model_meta.get(dep)
        if isinstance(meta, dict) and meta.get("max_output_tokens"):
            try:
                out_lim = int(meta["max_output_tokens"])
                if out_lim >= 1024:
                    max_tokens = min(out_lim, 128000)
            except Exception:
                pass

    return {
        "api_version": api_ver,
        "rpm": rpm,
        "threads": threads,
        "temperature": 0.0,
        "max_tokens": max_tokens,
    }


def get_optimized_azure_settings_for_steps(
    dep1, dep2="", dep3="", detected_api_version=None, model_meta=None
):
    """Step1〜3のデプロイ構成全体に対して安全かつ最適な Azure 設定を算出する"""
    deps = [d.strip() for d in (dep1, dep2, dep3) if str(d).strip()]
    if not deps:
        return get_recommended_azure_settings("", detected_api_version, model_meta)

    recs = [
        get_recommended_azure_settings(d, detected_api_version, model_meta)
        for d in deps
    ]
    return {
        "api_version": recs[0]["api_version"],
        "rpm": min(r["rpm"] for r in recs),
        "threads": min(r["threads"] for r in recs),
        "temperature": 0.0,
        "max_tokens": min(r["max_tokens"] for r in recs),
    }


def filter_mainstream_gemini_models(models_list):
    """
    Gemini のモデル一覧 [(display_name, model_id), ...] から
    現在（最新世代）の主力モデルのみを動的に抽出する。
    """
    if not models_list:
        return []

    def _extract_ver(mid):
        m = re.search(r"gemini-(\d+(?:\.\d+)?)", mid.lower())
        return float(m.group(1)) if m else 0.0

    clean_candidates = []
    for item in models_list:
        mid = item[1].lower()
        if any(
            k in mid
            for k in ["-exp", "experimental", "thinking-exp", "-latest", "tuning", "robotics"]
        ):
            continue
        if re.search(r"-(00\d|\d{2}-\d{2}|20\d{2})", mid):
            continue
        clean_candidates.append(item)

    if not clean_candidates:
        return models_list[:8]

    max_ver = max((_extract_ver(item[1]) for item in clean_candidates), default=0.0)

    mainstream = []
    for item in clean_candidates:
        mid = item[1].lower()
        ver = _extract_ver(mid)
        if max_ver >= 4.0:
            if ver >= 4.0 or (ver >= 3.5 and "flash" in mid):
                mainstream.append(item)
        elif max_ver >= 3.0:
            if ver >= 3.0:
                mainstream.append(item)
        else:
            if ver >= max_ver - 0.5:
                mainstream.append(item)

    return mainstream if mainstream else clean_candidates[:8]


def filter_mainstream_azure_models(models_list):
    """
    Azure OpenAI のモデル一覧 [model_id, ...] から現在（最新世代）の主力モデルを抽出する。
    """
    if not models_list:
        return []

    current_prefixes = ("gpt-6", "gpt-5", "o4-mini", "o4")
    legacy_prefixes = ("gpt-4.5", "gpt-4.1", "gpt-4o", "o3-mini", "o3")
    date_pattern = re.compile(r"-(20\d{2}-\d{2}-\d{2}|\d{4})$")
    exclude_sub = ("preview-20", "search", "transcribe", "diarize", "tts", "audio", "codex")

    def _pick_by_prefixes(prefixes):
        clean = []
        for mid in models_list:
            m = mid.lower()
            if not any(m == p or m.startswith(p + "-") or m.startswith(p + ".") for p in prefixes):
                continue
            if date_pattern.search(m):
                continue
            if any(k in m for k in exclude_sub):
                continue
            clean.append(mid)
        if clean:
            return clean

        family_latest = {}
        for mid in models_list:
            m = mid.lower()
            if not any(m == p or m.startswith(p + "-") or m.startswith(p + ".") for p in prefixes):
                continue
            if any(k in m for k in exclude_sub):
                continue
            fam = date_pattern.sub("", m)
            if fam not in family_latest or mid > family_latest[fam]:
                family_latest[fam] = mid
        if family_latest:
            return [mid for mid in models_list if mid in family_latest.values()]
        return []

    current_main = _pick_by_prefixes(current_prefixes)
    if current_main:
        return current_main

    legacy_main = _pick_by_prefixes(legacy_prefixes)
    if legacy_main:
        return legacy_main

    return models_list[:8]


# ==============================
# エラー診断
# ==============================
def diagnose_api_error(provider, error, model_or_deployment="", endpoint=""):
    """
    Gemini API / Azure OpenAI Service で発生したエラーを解析し、
    利用者が原因と対処法をひと目で理解できる構造化情報(dict)を返す。
    """
    err_str = str(error)
    err_lower = err_str.lower()
    target_name = model_or_deployment or "未指定"

    network_keywords = [
        "getaddrinfo failed",
        "name resolution",
        "nameresolutionerror",
        "connectionerror",
        "connecttimeout",
        "readtimeout",
        "timed out",
        "urlopen error",
        "ssl:",
        "certificate",
        "proxy",
        "winerror 10060",
        "winerror 10061",
        "remote disconnected",
        "connection refused",
    ]
    if any(k in err_lower for k in network_keywords):
        actions = [
            "インターネット接続が有効か確認してください。",
            "社内ネットワークやプロキシ環境をお使いの場合、AIサービスへの通信が制限されていないか確認してください。",
        ]
        if provider == "azure" and endpoint:
            actions.insert(
                0,
                f"エンドポイントURL「{endpoint}」にスペルミスや余分な文字がないか確認してください。",
            )
        return {
            "kind": "network",
            "title": "ネットワーク / URL接続エラー",
            "summary": "AIサーバーへの接続が確立できませんでした（通信経路またはURLの問題）。",
            "causes": [
                (
                    "入力したエンドポイントURLのホスト名が間違っているか、存在しません。"
                    if provider == "azure"
                    else "APIサーバーへ到達できませんでした。"
                ),
                "インターネットに接続されていない、または社内プロキシ・ファイアウォールにブロックされています。",
                "通信がタイムアウトしました。",
            ],
            "actions": actions,
            "raw": err_str,
        }

    if provider == "azure":
        if "instant inference quota" in err_lower or "still being initialized" in err_lower:
            return {
                "kind": "init_quota",
                "title": "Azure クォータ初期化中 (HTTP 403)",
                "summary": "Azure側で新しく作成したサブスクリプション／リソースの推論利用枠（クォータ）がまだ準備中です。",
                "causes": [
                    "APIやサブスクリプションを作成した直後のため、Azure側のバックエンドでモデル利用枠（TPM）の割り当て処理が進行中です。",
                    "アプリの設定ミスやリクエストの送りすぎ（429エラー）ではなく、Azure側の利用準備待ち状態です。",
                ],
                "actions": [
                    "作成直後の場合は、数分〜数十分ほど待ってから再度「🔌 接続テスト」を実行してください。",
                    "数時間経っても解消しない場合は、Azure AI Foundry の「管理センター ＞ クォータ」で対象モデルの利用枠が 0 のまま止まっていないか確認してください。",
                    "無料試用版や学生向けサブスクリプションなどでは標準クォータが自動付与されない場合があるため、Azure Portal でサブスクリプション種別もご確認ください。",
                ],
                "raw": err_str,
            }

        if (
            "401" in err_str
            or "unauthorized" in err_lower
            or "invalid subscription key" in err_lower
        ):
            return {
                "kind": "auth",
                "title": "認証エラー (HTTP 401)",
                "summary": "Azure OpenAI の APIキーまたはエンドポイントの組み合わせが正しくありません。",
                "causes": [
                    "APIキーに入力ミスがある、またはコピー時に文字が欠けている・余分な空白が含まれています。",
                    "別の Azure リソースの APIキー と エンドポイントURL を組み合わせて入力しています。",
                    "Azure Portal 上で APIキー が再生成され、古いキーが無効になっています。",
                ],
                "actions": [
                    "Azure Portal / Azure AI Foundry の「キーとエンドポイント」画面を開き、同じリソースの「KEY 1（または KEY 2）」と「エンドポイント」をセットで再コピーしてください。",
                    "「確認」ボタンで入力したキーの前後に余計な文字がないか確認してください。",
                ],
                "raw": err_str,
            }

        if (
            "404" in err_str
            or "deploymentnotfound" in err_lower
            or "resource not found" in err_lower
            or "not found" in err_lower
        ):
            return {
                "kind": "not_found",
                "title": "デプロイ名 / URL が見つかりません (HTTP 404)",
                "summary": f"指定されたデプロイ名「{target_name}」またはエンドポイントURLが存在しません。",
                "causes": [
                    f"「デプロイ名」に入力した「{target_name}」という名前のデプロイが Azure 側に存在しません（モデル名そのものではなく、自分で付けた『デプロイ名』を入力する必要があります）。",
                    "モデルを作成した直後で、デプロイのプロビジョニングがまだ完了していません。",
                    "エンドポイントURLの形式やリージョン、または APIバージョンが一致していません。",
                ],
                "actions": [
                    "Azure AI Foundry の「デプロイ（Deployments）」画面を開き、デプロイ状態が「成功」になっている「デプロイ名」を正確にコピーして貼り付けてください。",
                    "「キーとエンドポイント」画面に「Azure OpenAI エンドポイント（https://〜.openai.azure.com/）」が表示されている場合は、そちらのURLをお試しください。",
                    "モデルをデプロイした直後の場合は、2〜3分待ってから再テストしてください。",
                ],
                "raw": err_str,
            }

        if (
            "403" in err_str
            or "forbidden" in err_lower
            or "access denied" in err_lower
            or "virtual network" in err_lower
        ):
            return {
                "kind": "forbidden",
                "title": "アクセス拒否エラー (HTTP 403)",
                "summary": "Azure リソースへのアクセスが権限設定またはファイアウォールにより拒否されました。",
                "causes": [
                    "Azure リソース側のネットワーク設定（IP制限 / プライベートエンドポイント / VNet）により、現在のPCからのアクセスが許可されていません。",
                    "リソース側で「ローカル認証（APIキーによる認証）」が無効化され、Microsoft Entra ID 認証のみ許可されています。",
                    "利用中のサブスクリプションが無効化または凍結されています。",
                ],
                "actions": [
                    "Azure Portal で対象リソースの「ネットワーク」設定を開き、接続元のIPアドレスが許可されているか確認してください。",
                    "対象リソースで APIキー認証（ローカル認証）が有効になっているか管理者に確認してください。",
                ],
                "raw": err_str,
            }

        if (
            "429" in err_str
            or "rate limit" in err_lower
            or "too many requests" in err_lower
            or "insufficient_quota" in err_lower
            or "quota" in err_lower
        ):
            m = re.search(r"(?:retry after|try again in)\s*(\d+)", err_lower)
            wait_note = f"（約 {m.group(1)} 秒後に再試行するよう案内されています）" if m else ""
            return {
                "kind": "rate_limit",
                "title": "レート制限 / 利用枠超過 (HTTP 429)",
                "summary": f"Azure OpenAI の利用枠（TPM/RPM）上限に達しました。{wait_note}",
                "causes": [
                    "短時間に多くのリクエストを送信したため、デプロイに設定された1分間あたりの上限（TPM/RPM）を超過しました。",
                    "デプロイのクォータ割り当てが非常に小さい設定になっています。",
                ],
                "actions": [
                    "数十秒〜数分ほど待ってから再度お試しください。",
                    "設定画面の「RPM」を小さくし、「スレッド（並列数）」を 1〜2 に下げてください。",
                    "必要に応じて Azure AI Foundry の「デプロイ」設定から割り当てクォータ（TPM）を増やしてください。",
                ],
                "raw": err_str,
            }

        if (
            "400" in err_str
            or "invalid_request_error" in err_lower
            or "unsupported" in err_lower
            or "api-version" in err_lower
        ):
            return {
                "kind": "bad_request",
                "title": "設定 / パラメータエラー (HTTP 400)",
                "summary": "APIバージョンまたはモデルの設定がサーバーに受け付けられませんでした。",
                "causes": [
                    "選択した「APIバージョン」が、お使いのモデルやエンドポイントに対応していません。",
                    "指定したデプロイが画像入力（Vision）やチャット補完形式に対応していないモデルです。",
                ],
                "actions": [
                    "「APIバージョン」を最新の「v1」、または「2025-04-01-preview」「2024-10-21」等に切り替えて再テストしてください。",
                    "デプロイしているモデルが gpt-5 / gpt-5-mini / gpt-6 等のチャット・画像対応モデルか確認してください。",
                ],
                "raw": err_str,
            }

        return {
            "kind": "unknown",
            "title": "Azure OpenAI 通信エラー",
            "summary": "Azure OpenAI との通信中に予期しないエラーが発生しました。",
            "causes": [
                "エンドポイントURL・APIキー・デプロイ名・APIバージョンのいずれかに誤りがあるか、Azure側で一時的なエラーが発生しています。"
            ],
            "actions": [
                "下の「詳細情報（エラー原文）」の内容を確認し、入力内容（URL・キー・デプロイ名）を見直してください。"
            ],
            "raw": err_str,
        }

    # --- Gemini API 用エラー診断 ---
    if (
        "api_key_invalid" in err_lower
        or "api key not valid" in err_lower
        or "invalid api key" in err_lower
        or ("400" in err_str and "key" in err_lower)
        or "401" in err_str
    ):
        return {
            "kind": "auth",
            "title": "Gemini APIキーエラー",
            "summary": "入力された Gemini API キーが正しく認識されませんでした。",
            "causes": [
                "APIキーに誤りがある、またはコピー時に前後の文字が欠けている・余分な空白が混入しています。",
                "Google AI Studio 側でキーが削除または無効化されています。",
            ],
            "actions": [
                "Google AI Studio（https://aistudio.google.com/app/apikey）で発行された APIキー を再度コピーし、右クリック→「貼り付け」で入力してください。"
            ],
            "raw": err_str,
        }

    if "403" in err_str or "permission_denied" in err_lower or "permission denied" in err_lower:
        return {
            "kind": "forbidden",
            "title": "アクセス権限エラー (HTTP 403)",
            "summary": "このAPIキーでは対象モデルまたはAPIへのアクセスが許可されていません。",
            "causes": [
                "Google Cloud プロジェクト側で Generative Language API が有効化されていないか、APIキーの制限設定により拒否されています。",
                "入力したキーが Gemini API 用ではなく、別のGoogleサービスのキーになっています。",
            ],
            "actions": [
                "Google AI Studio で新しいプロジェクトの APIキー を発行してお試しください。"
            ],
            "raw": err_str,
        }

    if "404" in err_str or "not found" in err_lower or "is not supported" in err_lower:
        return {
            "kind": "not_found",
            "title": "モデル利用不可 (HTTP 404)",
            "summary": f"モデル「{target_name}」が存在しないか、現在利用できません。",
            "causes": [
                f"選択中のモデル「{target_name}」が廃止されたか、モデルIDの表記が古くなっています。",
                "お使いのアカウント（無料枠等）ではこのモデルが公開されていません。",
            ],
            "actions": [
                "設定画面の「🌐 最新モデルを更新」ボタンを押して、最新の利用可能なモデル一覧を取得してください。",
                "「Gemini 4 Hydrogen」や「Gemini 3.8 Flash」などの主力モデルに切り替えて再テストしてください。",
            ],
            "raw": err_str,
        }

    if "429" in err_str or "quota" in err_lower or "resource_exhausted" in err_lower:
        is_zero_limit = "limit: 0" in err_lower or "limit:0" in err_lower
        m = re.search(r"retry in ([\d\.]+)s", err_lower) or re.search(
            r"seconds:\s*(\d+)", err_lower
        )
        if is_zero_limit:
            return {
                "kind": "rate_limit",
                "title": "モデル利用枠なし (HTTP 429)",
                "summary": f"現在のプランではモデル「{target_name}」の利用枠が 0 に設定されているため利用できません。",
                "causes": [
                    f"無料枠（Free Tier）では「{target_name}」が利用対象外になっている、またはプロジェクトのクォータが 0 になっています。"
                ],
                "actions": [
                    "「Gemini 4 Hydrogen」「Gemini 3.8 Flash」「Gemini 3.5 Flash-Lite」等の無料枠対応モデルを選択してください。",
                    "課金枠（Paid Tier）をご利用の場合は、Google AI Studio で課金設定が有効になっているか確認してください。",
                ],
                "raw": err_str,
            }
        if "per day" in err_lower or "perday" in err_lower:
            return {
                "kind": "rate_limit",
                "title": "1日の利用上限に達しました (HTTP 429)",
                "summary": "Gemini API の1日あたりの無料利用上限（RPD）に到達しました。",
                "causes": [
                    "本日分のAPIリクエスト回数が上限に達したため、日付が変わる（またはリセット時刻になる）まで制限されています。"
                ],
                "actions": [
                    "しばらく時間をおいて（または翌日以降に）再度お試しください。",
                    "別のモデル（例: Flash-Lite 等）に切り替えるか、課金枠（Paid Tier）のご利用をご検討ください。",
                ],
                "raw": err_str,
            }
        wait_note = f"（約 {int(float(m.group(1)))} 秒後に回復見込み）" if m else ""
        return {
            "kind": "rate_limit",
            "title": "API利用枠超過 (HTTP 429)",
            "summary": f"短時間にリクエストが集中し、利用制限に達しました。{wait_note}",
            "causes": [
                "1分間あたりのリクエスト上限（RPM）またはトークン上限（TPM）を超過しました。",
                "無料枠のバースト制限に引っかかっています。",
            ],
            "actions": [
                "数十秒ほど待ってから再試行してください。",
                "「RPM」を 10〜12 以下、「スレッド」を 1 に下げると安定します。",
            ],
            "raw": err_str,
        }

    return {
        "kind": "unknown",
        "title": "Gemini API 通信エラー",
        "summary": "Gemini API との通信中に予期しないエラーが発生しました。",
        "causes": [
            "APIキーの設定、選択モデル、またはネットワーク環境に問題がある可能性があります。"
        ],
        "actions": [
            "下の「詳細情報（エラー原文）」を確認し、APIキーとモデル設定を見直してください。"
        ],
        "raw": err_str,
    }


def format_api_error_message(diag):
    """diagnose_api_error の結果を読みやすいテキスト形式に整形する"""
    lines = [diag.get("summary", ""), ""]
    if diag.get("causes"):
        lines.append("【考えられる原因】")
        for c in diag["causes"]:
            lines.append(f"・{c}")
        lines.append("")
    if diag.get("actions"):
        lines.append("【対処方法】")
        for a in diag["actions"]:
            lines.append(f"・{a}")
        lines.append("")
    lines.append(f"【詳細（エラー原文）】\n{diag.get('raw', '')}")
    return "\n".join(lines)


class DiagnosedAPIError(Exception):
    """診断情報 (diag dict) を添付した API 例外"""

    def __init__(self, diag, fatal=True):
        self.diag = diag
        self.fatal = fatal
        super().__init__(format_api_error_message(diag))


class AnalysisCancelledError(Exception):
    """ユーザー操作により解析が中止されたことを表す例外"""


def _sleep_with_cancel(seconds, cancel_event=None):
    """キャンセルイベントを監視しながら指定秒数待機する"""
    if seconds <= 0:
        return
    if cancel_event is None:
        time.sleep(seconds)
        return
    if cancel_event.wait(timeout=seconds):
        raise AnalysisCancelledError("ユーザーによって解析が中止されました。")


# ==============================
# レート制限（RPMスロットリング）
# ==============================
class RateLimiter:
    """スレッドセーフな RPM レートリミッター（キャンセル対応）"""

    def __init__(self, min_interval=1.0):
        self._lock = threading.Lock()
        self._timestamps = []
        self._min_interval = min_interval

    def wait_if_needed(self, max_rpm, status_callback=None, cancel_event=None):
        if cancel_event is not None and cancel_event.is_set():
            raise AnalysisCancelledError("ユーザーによって解析が中止されました。")

        max_rpm = max(1, int(max_rpm or 1))
        total_wait = 0.0

        with self._lock:
            now = time.time()
            self._timestamps = [t for t in self._timestamps if now - t < 60.0]

            if len(self._timestamps) >= max_rpm:
                expected_run_time = self._timestamps[0] + 60.0
                total_wait = max(0.0, expected_run_time - now)
            else:
                if self._timestamps:
                    last_planned = max(self._timestamps)
                    if now < last_planned + self._min_interval:
                        expected_run_time = last_planned + self._min_interval
                        total_wait = max(0.0, expected_run_time - now)
                    else:
                        expected_run_time = now
                        total_wait = 0.0
                else:
                    expected_run_time = now
                    total_wait = 0.0

            self._timestamps.append(expected_run_time)
            self._timestamps.sort()

        if total_wait > 0:
            if total_wait > 2.0 and status_callback:
                status_callback(f"通信間隔調整中... (約 {int(total_wait)}秒待機)")
            _sleep_with_cancel(total_wait, cancel_event)


# ==============================
# 最新モデル一覧取得機能
# ==============================
def fetch_gemini_models_with_meta(api_key):
    """
    Gemini API から最新の利用可能な生成モデル一覧と各モデルのメタ情報
    (output_token_limit / input_token_limit) を取得し、(models_list, model_meta) を返す。
    """
    if not api_key or not api_key.strip():
        raise Exception("APIキーが入力されていません。")

    genai.configure(api_key=api_key.strip())
    new_models = []
    model_meta = {}
    exclude_keywords = ["tts", "audio", "image", "vision", "embedding", "aqa", "learnlm"]

    for m in genai.list_models():
        methods = getattr(m, "supported_generation_methods", []) or []
        if "generateContent" in methods and "gemini" in m.name.lower():
            model_id = m.name.replace("models/", "")
            if any(k in model_id.lower() for k in exclude_keywords):
                continue
            display_name = getattr(m, "display_name", model_id) or model_id
            new_models.append((f"{display_name} ({model_id})", model_id))
            out_lim = getattr(m, "output_token_limit", None)
            in_lim = getattr(m, "input_token_limit", None)
            meta_entry = {}
            if out_lim:
                try:
                    meta_entry["output_token_limit"] = int(out_lim)
                except Exception:
                    pass
            if in_lim:
                try:
                    meta_entry["input_token_limit"] = int(in_lim)
                except Exception:
                    pass
            if meta_entry:
                model_meta[model_id] = meta_entry

    def _gemini_sort_key(item):
        mid = item[1].lower()
        ver_match = re.search(r"gemini-(\d+(?:\.\d+)?)", mid)
        ver_num = float(ver_match.group(1)) if ver_match else 0.0
        is_preview = 1 if ("preview" in mid or "exp" in mid) else 0
        is_flash = 1 if "flash" in mid else 0
        return (-ver_num, is_preview, -is_flash, mid)

    new_models.sort(key=_gemini_sort_key)
    return new_models, model_meta


def fetch_gemini_models(api_key):
    """後方互換ラッパー: Gemini モデル一覧のみを返す"""
    models, _ = fetch_gemini_models_with_meta(api_key)
    return models


def fetch_azure_models_detailed(api_key, endpoint, api_version=DEFAULT_AZURE_API_VERSION):
    """
    Azure OpenAI Service / Azure AI Foundry から利用可能な最新モデル一覧を取得し、
    通信に成功した最新の推奨APIバージョンおよび各モデルのメタ情報も判定して
    (model_ids, best_api_version, model_meta) を返す。
    """
    base_endpoint, _, resolved_ver = parse_azure_endpoint_info(endpoint, "", api_version)
    if not base_endpoint:
        raise Exception("エンドポイントURLが正しく入力されていません。")
    if not api_key or not api_key.strip():
        raise Exception("APIキーが入力されていません。")

    api_key = api_key.strip()
    # 最新の v1 API (/openai/v1/models) を最優先で試行し、未対応環境のみ従来の日付バージョンへフォールバック
    candidate_requests = [
        (f"{base_endpoint.rstrip('/')}/openai/v1/models", "v1"),
    ]
    versions_to_try = []
    for v in list(AZURE_API_VERSION_CANDIDATES) + [resolved_ver]:
        if v and v.lower() != "v1" and v not in versions_to_try:
            versions_to_try.append(v)

    for v in versions_to_try:
        candidate_requests.append(
            (f"{base_endpoint.rstrip('/')}/openai/models?api-version={urllib.parse.quote(v)}", v)
        )

    raw_data = None
    best_api_version = resolved_ver or DEFAULT_AZURE_API_VERSION
    last_err = None
    for url, ver_label in candidate_requests:
        req = urllib.request.Request(
            url,
            headers={"Content-Type": "application/json", "api-key": api_key},
            method="GET",
        )
        try:
            with urllib.request.urlopen(req, timeout=30.0) as resp:
                raw_data = json.loads(resp.read().decode("utf-8"))
                best_api_version = ver_label
                break
        except urllib.error.HTTPError as http_err:
            err_body = ""
            try:
                err_body = http_err.read().decode("utf-8", errors="replace")
            except Exception:
                pass
            last_err = Exception(f"HTTP {http_err.code}: {err_body or http_err.reason}")
            if http_err.code in (401, 403):
                raise last_err
        except Exception as e:
            last_err = e

    if raw_data is None:
        raise last_err if last_err else Exception("モデル一覧の取得に失敗しました。")

    items = (
        raw_data.get("data", [])
        if isinstance(raw_data, dict)
        else (raw_data if isinstance(raw_data, list) else [])
    )
    exclude_keywords = [
        "embedding",
        "whisper",
        "tts",
        "dall-e",
        "sora",
        "audio",
        "realtime",
        "transcribe",
        "moderation",
        "babbage",
        "davinci",
        "ada",
        "curie",
        "text-",
        "codex",
        "instruct",
    ]

    seen = set()
    model_ids = []
    model_meta = {}
    for item in items:
        if not isinstance(item, dict):
            continue
        mid = str(item.get("id") or item.get("model") or "").strip()
        if not mid or mid in seen:
            continue
        status = str(item.get("status", "")).lower()
        if status and status not in ("succeeded", "active", "available"):
            continue
        caps = item.get("capabilities")
        if isinstance(caps, dict) and caps.get("chat_completion") is False:
            continue
        mid_lower = mid.lower()
        if any(k in mid_lower for k in exclude_keywords):
            continue
        seen.add(mid)
        model_ids.append(mid)

        # モデルにトークン上限メタ情報が含まれていれば保持
        max_out = (
            item.get("max_output_tokens")
            or item.get("max_completion_tokens")
            or (caps.get("max_output_tokens") if isinstance(caps, dict) else None)
        )
        if max_out:
            try:
                model_meta[mid] = {"max_output_tokens": int(max_out)}
            except Exception:
                pass

    def _azure_sort_key(mid):
        m = mid.lower()
        ver_m = re.match(r"^gpt-(\d+(?:\.\d+)?)", m)
        if ver_m:
            ver_val = float(ver_m.group(1))
            if ver_val >= 6.0:
                rank = 0
            elif ver_val >= 5.0:
                rank = 1
            elif ver_val >= 4.5:
                rank = 3
            elif ver_val >= 4.1:
                rank = 4
            else:
                rank = 6
            sub_ver = -ver_val
        elif m.startswith("o4"):
            rank, sub_ver = 2, 0.0
        elif m.startswith("gpt-4o"):
            rank, sub_ver = 5, 0.0
        elif m.startswith("o3"):
            rank, sub_ver = 7, 0.0
        elif m.startswith("o1"):
            rank, sub_ver = 8, 0.0
        else:
            rank, sub_ver = 9, 0.0
        has_date = 1 if re.search(r"-20\d{2}-", m) else 0
        return (rank, sub_ver, has_date, m)

    model_ids.sort(key=_azure_sort_key)
    return model_ids, best_api_version, model_meta


def fetch_azure_models_with_meta(api_key, endpoint, api_version=DEFAULT_AZURE_API_VERSION):
    """後方互換ラッパー: (model_ids, best_api_version) を返す"""
    model_ids, best_ver, _ = fetch_azure_models_detailed(api_key, endpoint, api_version)
    return model_ids, best_ver


# ==============================
# 通信・接続テスト・構造化出力
# ==============================
_AZURE_CLIENT_CACHE = {}
_AZURE_CLIENT_CACHE_LOCK = threading.Lock()


def _get_azure_sdk_client(use_v1_api, api_key, base_endpoint, resolved_ver):
    """スレッドセーフに OpenAI / AzureOpenAI クライアントを再利用しコネクションプールを活かす"""
    cache_key = (use_v1_api, api_key, base_endpoint, resolved_ver)
    with _AZURE_CLIENT_CACHE_LOCK:
        client = _AZURE_CLIENT_CACHE.get(cache_key)
        if client is not None:
            return client
        if use_v1_api and OpenAI is not None:
            client = OpenAI(
                api_key=api_key,
                base_url=f"{base_endpoint.rstrip('/')}/openai/v1/",
                max_retries=0,
                timeout=180.0,
            )
        elif not use_v1_api and AzureOpenAI is not None:
            client = AzureOpenAI(
                api_key=api_key,
                api_version=resolved_ver,
                azure_endpoint=base_endpoint,
                max_retries=0,
                timeout=180.0,
            )
        else:
            return None
        # キャッシュ肥大防止
        if len(_AZURE_CLIENT_CACHE) > 8:
            _AZURE_CLIENT_CACHE.clear()
        _AZURE_CLIENT_CACHE[cache_key] = client
        return client


def encode_pil_image_data_url(pil_img):
    """
    PIL画像をAPI送信用 data URL へエンコードする。
    PNG と 高品質JPEG(92) を比較してサイズが小さい方を採用し、転送量を削減する。
    """
    png_buf = io.BytesIO()
    pil_img.save(png_buf, format="PNG", optimize=True)
    png_bytes = png_buf.getvalue()

    try:
        jpg_buf = io.BytesIO()
        pil_img.convert("RGB").save(jpg_buf, format="JPEG", quality=92, optimize=True)
        jpg_bytes = jpg_buf.getvalue()
        if len(jpg_bytes) < len(png_bytes):
            return f"data:image/jpeg;base64,{base64.b64encode(jpg_bytes).decode('utf-8')}"
    except Exception:
        pass

    return f"data:image/png;base64,{base64.b64encode(png_bytes).decode('utf-8')}"


def _encode_pil_image_to_base64(pil_img):
    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode("utf-8")


def test_gemini_connection(api_key, model_name):
    """Gemini API の接続テストを実行する"""
    if not api_key or not api_key.strip():
        raise Exception("APIキーが入力されていません。")
    genai.configure(api_key=api_key.strip())
    model = genai.GenerativeModel(model_name)
    response = model.generate_content(f"Connection Test: {time.time()}")
    return getattr(response, "text", "OK")


def call_azure_openai_api(
    api_key,
    endpoint,
    deployment,
    api_version,
    prompt,
    pil_img=None,
    expect_json=False,
    temperature=0.0,
    max_tokens=8192,
    system_prompt=None,
    image_data_url=None,
):
    """
    Azure OpenAI Service に対してリクエストを送信し、応答テキストを返す。
    openai パッケージが導入されている場合は公式 SDK (OpenAI / AzureOpenAI) を再利用し、
    未導入環境では標準ライブラリ (urllib.request) による REST API 通信へ自動フォールバックする。
    推論モデル (o1/o3/o4/gpt-5/gpt-6 等) と標準モデルのパラメータ差異を自動吸収する。
    """
    base_endpoint, resolved_dep, resolved_ver = parse_azure_endpoint_info(
        endpoint, deployment, api_version
    )
    if not base_endpoint:
        raise Exception("エンドポイントURLが正しく入力されていません。")
    if not resolved_dep:
        raise Exception("デプロイ名が入力されていません。")
    if not api_key:
        raise Exception("APIキーが入力されていません。")

    resolved_data_url = image_data_url
    if resolved_data_url is None and pil_img is not None:
        resolved_data_url = encode_pil_image_data_url(pil_img)

    if resolved_data_url is not None:
        user_content = [
            {"type": "text", "text": prompt},
            {
                "type": "image_url",
                "image_url": {
                    "url": resolved_data_url,
                    "detail": "high",
                },
            },
        ]
    else:
        user_content = prompt

    sys_msg = (
        system_prompt
        or "あなたは高精度な2D機械図面解析・立体体積計算アシスタントです。指示されたJSON出力形式を厳守して回答してください。"
    )
    messages = [
        {"role": "system", "content": sys_msg},
        {"role": "user", "content": user_content},
    ]

    dep_lower = resolved_dep.lower()
    is_reasoning_model = any(
        dep_lower.startswith(prefix) or f"-{prefix}" in dep_lower
        for prefix in ["o1", "o3", "o4", "gpt-5", "gpt-6"]
    )

    use_max_completion_tokens = is_reasoning_model
    use_temperature = not is_reasoning_model
    use_json_format = expect_json

    for _ in range(4):
        kwargs = {
            "model": resolved_dep,
            "messages": messages,
        }
        if use_max_completion_tokens:
            kwargs["max_completion_tokens"] = max_tokens
        else:
            kwargs["max_tokens"] = max_tokens

        if use_temperature:
            kwargs["temperature"] = temperature

        if use_json_format:
            kwargs["response_format"] = {"type": "json_object"}

        use_v1_api = resolved_ver.lower() == "v1"

        try:
            sdk_client = _get_azure_sdk_client(use_v1_api, api_key, base_endpoint, resolved_ver)
            if sdk_client is not None:
                response = sdk_client.chat.completions.create(**kwargs)
                if not response.choices or response.choices[0].message.content is None:
                    raise Exception("コンテンツフィルタ等により空の応答が返されました。")
                return response.choices[0].message.content.strip()
            else:
                if use_v1_api:
                    url = f"{base_endpoint.rstrip('/')}/openai/v1/chat/completions"
                    payload = dict(kwargs)
                else:
                    url = (
                        f"{base_endpoint.rstrip('/')}/openai/deployments/"
                        f"{urllib.parse.quote(resolved_dep)}/chat/completions"
                        f"?api-version={urllib.parse.quote(resolved_ver)}"
                    )
                    payload = {k: v for k, v in kwargs.items() if k != "model"}
                req_data = json.dumps(payload).encode("utf-8")
                req = urllib.request.Request(
                    url,
                    data=req_data,
                    headers={"Content-Type": "application/json", "api-key": api_key},
                    method="POST",
                )
                try:
                    with urllib.request.urlopen(req, timeout=180.0) as resp:
                        resp_body = json.loads(resp.read().decode("utf-8"))
                except urllib.error.HTTPError as http_err:
                    err_body = ""
                    try:
                        err_body = http_err.read().decode("utf-8", errors="replace")
                    except Exception:
                        pass
                    raise Exception(f"HTTP {http_err.code}: {err_body or http_err.reason}")

                choices = resp_body.get("choices", [])
                if not choices or choices[0].get("message", {}).get("content") is None:
                    raise Exception("コンテンツフィルタ等により空の応答が返されました。")
                return choices[0]["message"]["content"].strip()

        except Exception as e:
            err_str = str(e)
            err_lower = err_str.lower()
            if (
                "400" in err_str
                or "unsupported" in err_lower
                or "invalid_request_error" in err_lower
            ):
                changed = False
                if "max_tokens" in err_lower and not use_max_completion_tokens:
                    use_max_completion_tokens = True
                    changed = True
                elif "max_completion_tokens" in err_lower and use_max_completion_tokens:
                    use_max_completion_tokens = False
                    changed = True
                if "temperature" in err_lower and use_temperature:
                    use_temperature = False
                    changed = True
                if "response_format" in err_lower and use_json_format:
                    use_json_format = False
                    changed = True
                if changed:
                    continue
            raise e


def test_azure_connection(api_key, endpoint, deployment, api_version):
    """設定ダイアログの「接続テスト」ボタンから呼ばれる Azure OpenAI 接続確認関数"""
    return call_azure_openai_api(
        api_key=api_key,
        endpoint=endpoint,
        deployment=deployment,
        api_version=api_version,
        prompt=f"Connection Test ({time.time()}): OK とだけ返答してください。",
        pil_img=None,
        expect_json=False,
        temperature=0.0,
        max_tokens=32,
        system_prompt="You are a helpful assistant. Reply briefly.",
    )


def generate_structured_json(
    provider,
    prompt,
    pil_img,
    response_schema,
    *,
    # Gemini 用パラメータ
    api_key="",
    model_name="",
    disable_safety=True,
    # Azure OpenAI 用パラメータ
    azure_endpoint="",
    azure_deployment="",
    azure_api_version=DEFAULT_AZURE_API_VERSION,
    # 共通パラメータ
    temperature=0.0,
    max_tokens=8192,
    rpm=12,
    rate_limiter=None,
    status_callback=None,
    max_retries=4,
    cancel_event=None,
    image_data_url=None,
):
    """
    Gemini API / Azure OpenAI Service 共通の、レート制御・自動リトライ・キャンセル・エラー診断付き
    構造化 JSON 生成呼び出し。応答テキストを返す。
    """
    safety_settings = None
    if provider == "gemini" and disable_safety:
        safety_settings = [
            {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
        ]

    # Azure 用に Pydantic スキーマを JSON 出力仕様として明示追加
    azure_prompt = prompt
    if provider == "azure" and response_schema is not None:
        try:
            schema_json = json.dumps(
                response_schema.model_json_schema(), ensure_ascii=False, indent=2
            )
            azure_prompt = (
                f"{prompt}\n\n"
                f"【出力JSON形式（絶対厳守）】\n"
                f"余計な解説文やマークダウン記法を含めず、以下のJSON Schemaに完全に準拠した純粋なJSONオブジェクトのみを出力してください。\n"
                f"{schema_json}"
            )
        except Exception:
            pass

    target_name = azure_deployment if provider == "azure" else model_name
    last_diag = None

    for attempt in range(max_retries):
        if cancel_event is not None and cancel_event.is_set():
            raise AnalysisCancelledError("ユーザーによって解析が中止されました。")

        if rate_limiter is not None:
            rate_limiter.wait_if_needed(
                rpm, status_callback=status_callback, cancel_event=cancel_event
            )

        required_sleep = 0.0
        try:
            if provider == "azure":
                return call_azure_openai_api(
                    api_key=api_key,
                    endpoint=azure_endpoint,
                    deployment=azure_deployment,
                    api_version=azure_api_version,
                    prompt=azure_prompt,
                    pil_img=pil_img,
                    expect_json=True,
                    temperature=temperature,
                    max_tokens=max_tokens,
                    image_data_url=image_data_url,
                )
            else:
                tokens_clipped = min(int(max_tokens or 65535), 65535)
                gen_config = genai.types.GenerationConfig(
                    response_mime_type="application/json",
                    response_schema=response_schema,
                    temperature=temperature,
                    max_output_tokens=tokens_clipped,
                )
                model_instance = genai.GenerativeModel(model_name)
                response = model_instance.generate_content(
                    [prompt, pil_img],
                    generation_config=gen_config,
                    safety_settings=safety_settings,
                )
                if not getattr(response, "candidates", None):
                    raise Exception("AIからの応答が空でした（安全フィルタ等によりブロックされた可能性があります）。")
                if not getattr(response, "parts", None):
                    raise Exception("AIからの応答テキストが空でした（安全フィルタ等によりブロックされた可能性があります）。")
                return response.text

        except AnalysisCancelledError:
            raise
        except Exception as api_err:
            if cancel_event is not None and cancel_event.is_set():
                raise AnalysisCancelledError("ユーザーによって解析が中止されました。")

            err_str = str(api_err)
            err_lower = err_str.lower()
            diag = diagnose_api_error(
                provider,
                api_err,
                model_or_deployment=target_name,
                endpoint=azure_endpoint if provider == "azure" else "",
            )
            last_diag = diag

            # 致命的な設定・権限・存在しないモデル等のエラーはリトライせず即時通知
            fatal_kinds = (
                ["init_quota", "auth", "forbidden", "not_found", "bad_request"]
                if provider == "azure"
                else ["auth", "forbidden", "not_found"]
            )
            if diag["kind"] in fatal_kinds:
                raise DiagnosedAPIError(diag)

            if diag["kind"] == "rate_limit":
                if provider == "gemini" and (
                    "limit: 0" in err_lower
                    or "limit:0" in err_lower
                    or "per day" in err_lower
                    or "perday" in err_lower
                ):
                    raise DiagnosedAPIError(diag)

                m = (
                    re.search(r"retry in ([\d\.]+)s", err_lower)
                    or re.search(r"seconds:\s*(\d+)", err_lower)
                    or re.search(r"(?:retry after|try again in)\s*(\d+)", err_lower)
                )
                if m:
                    wait_sec = float(m.group(1))
                    limit_threshold = 30.0 if provider == "azure" else 20.0
                    if wait_sec > limit_threshold:
                        raise DiagnosedAPIError(diag)
                    required_sleep = wait_sec + 2.0
                else:
                    required_sleep = 4.0 * (2 ** attempt) + random.uniform(1.0, 2.5)
            else:
                required_sleep = 2.0 * (2 ** attempt) + random.uniform(0.5, 1.5)

            if attempt == max_retries - 1:
                raise DiagnosedAPIError(diag, fatal=False)

            if status_callback:
                status_callback(
                    f"一時的な通信エラーのため待機後に再試行します... ({attempt + 1}/{max_retries - 1}, 約{int(required_sleep)}秒)"
                )
            _sleep_with_cancel(min(required_sleep, 45.0), cancel_event)

    if last_diag:
        raise DiagnosedAPIError(last_diag, fatal=False)
    raise Exception("AIとの通信に失敗しました。")
