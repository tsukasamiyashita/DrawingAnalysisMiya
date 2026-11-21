# -*- coding: utf-8 -*-
"""
AI詳細設定ダイアログ & エラー診断ダイアログモジュール
Gemini API (無料枠 / 課金枠) および Azure OpenAI Service に対応
常時ライトモード固定
"""

import copy
import threading
import tkinter as tk
from tkinter import Menu, messagebox
import tkinter.scrolledtext as st
import webbrowser
import customtkinter as ctk

from api_client import (
    AZURE_API_VERSION_CANDIDATES,
    DEFAULT_AZURE_API_VERSION,
    DEFAULT_AZURE_DEPLOYMENT_STEP1,
    DEFAULT_AZURE_DEPLOYMENT_STEP2,
    DEFAULT_AZURE_DEPLOYMENT_STEP3,
    DEFAULT_AZURE_MODELS,
    DEFAULT_GEMINI_MODELS,
    DEFAULT_GEMINI_MODEL_STEP1,
    DEFAULT_GEMINI_MODEL_STEP2,
    DEFAULT_GEMINI_MODEL_STEP3,
    diagnose_api_error,
    fetch_azure_models_detailed,
    fetch_gemini_models_with_meta,
    filter_mainstream_azure_models,
    filter_mainstream_gemini_models,
    format_api_error_message,
    get_optimized_azure_settings_for_steps,
    get_optimized_gemini_params_for_steps,
    get_recommended_azure_settings,
    get_recommended_gemini_settings,
    parse_azure_endpoint_info,
    pick_recommended_gemini_step_models,
    safe_float,
    safe_int,
    test_azure_connection,
    test_gemini_connection,
)


# ==============================
# ライトモード固定カラーパレット
# ==============================
BG_COLOR = "#F0F4F8"
CARD_BG = "#FFFFFF"
PRIMARY = "#0D6EFD"
PRIMARY_HOVER = "#0B5ED7"
TEXT_COLOR = "#212529"
MUTED_TEXT = "#6C757D"
BORDER_COLOR = "#DEE2E6"
HEADER_BG = "#E9ECEF"
AZURE_HEADER_BG = "#E7F1FF"
COLOR_SUCCESS = "#198754"
COLOR_SUCCESS_HOVER = "#157347"
COLOR_WARNING = "#FFC107"
COLOR_WARNING_HOVER = "#E0A800"
COLOR_DANGER = "#DC3545"
COLOR_DANGER_HOVER = "#B02A37"
COLOR_NEUTRAL = "#6C757D"
COLOR_NEUTRAL_HOVER = "#5C636A"

TAB_GEMINI_FREE = "🟢 Gemini 無料枠 (Free)"
TAB_GEMINI_PAID = "🔵 Gemini 課金枠 (Paid)"
TAB_AZURE = "🔷 Azure OpenAI Service"


# ==============================
# 共通UIヘルパー
# ==============================
def center_window(win, parent, w, h):
    """ウィンドウを親ウィンドウの中央かつ画面内に収まる位置へ配置する"""
    try:
        screen_w = win.winfo_screenwidth()
        screen_h = win.winfo_screenheight()
        pw = parent.winfo_width() if parent and parent.winfo_width() > 100 else 1100
        ph = parent.winfo_height() if parent and parent.winfo_height() > 100 else 750
        px = parent.winfo_rootx() if parent else 0
        py = parent.winfo_rooty() if parent else 0
        raw_x = px + (pw // 2) - (w // 2)
        raw_y = py + (ph // 2) - (h // 2)
        x = max(10, min(raw_x, screen_w - w - 10))
        y = max(10, min(raw_y, screen_h - h - 40))
        win.geometry(f"{w}x{h}+{x}+{y}")
    except Exception:
        win.geometry(f"{w}x{h}")


def bind_context_menu(widget):
    """CTkEntry / CTkComboBox / tk.Entry に右クリックメニュー（貼り付け等）を追加する"""
    target = getattr(widget, "_entry", widget)

    def _show_menu(event):
        menu = Menu(target, tearoff=0, bg="white", fg=TEXT_COLOR)

        def _paste():
            try:
                text = target.clipboard_get()
                try:
                    target.delete("sel.first", "sel.last")
                except tk.TclError:
                    pass
                target.insert(tk.INSERT, text)
            except tk.TclError:
                pass

        menu.add_command(label="切り取り", command=lambda: target.event_generate("<<Cut>>"))
        menu.add_command(label="コピー", command=lambda: target.event_generate("<<Copy>>"))
        menu.add_command(label="貼り付け", command=_paste)
        menu.add_separator()
        menu.add_command(label="すべて選択", command=lambda: target.select_range(0, tk.END))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    try:
        target.bind("<Button-3>", _show_menu)
    except Exception:
        pass


def show_api_error_dialog(parent, diag):
    """
    diagnose_api_error() の診断結果を、原因・対処法・エラー原文に分けて
    初心者にもひと目で分かりやすく表示するエラーダイアログ（ライトモード固定）。
    """
    err_win = tk.Toplevel(parent)
    err_win.title(f"接続・通信エラー - {diag.get('title', 'エラー')}")
    screen_w = err_win.winfo_screenwidth()
    screen_h = err_win.winfo_screenheight()
    win_w = min(650, screen_w - 40)
    win_h = min(560, screen_h - 80)
    err_win.configure(bg=BG_COLOR)
    center_window(err_win, parent, win_w, win_h)
    err_win.transient(parent)
    err_win.grab_set()

    def _close_err():
        err_win.destroy()
        try:
            if parent and parent.winfo_exists() and isinstance(parent, (tk.Toplevel, ctk.CTkToplevel)):
                parent.grab_set()
        except Exception:
            pass

    err_win.protocol("WM_DELETE_WINDOW", _close_err)

    kind = diag.get("kind", "unknown")
    is_wait_type = kind in ["init_quota", "rate_limit"]
    accent_col = COLOR_WARNING_HOVER if is_wait_type else COLOR_DANGER
    badge_bg = "#FFF3CD" if is_wait_type else "#F8D7DA"
    badge_fg = "#856404" if is_wait_type else "#842029"
    badge_icon = "⏳" if is_wait_type else "❌"

    # 上部タイトルバナー
    banner = tk.Frame(
        err_win,
        bg=CARD_BG,
        padx=16,
        pady=12,
        highlightbackground=BORDER_COLOR,
        highlightthickness=1,
    )
    banner.pack(fill=tk.X, side=tk.TOP)

    title_row = tk.Frame(banner, bg=CARD_BG)
    title_row.pack(fill=tk.X)
    tk.Label(
        title_row,
        text=f" {badge_icon} {diag.get('title', 'エラー')} ",
        font=("Segoe UI", 11, "bold"),
        bg=badge_bg,
        fg=badge_fg,
        padx=8,
        pady=3,
    ).pack(side=tk.LEFT)

    tk.Label(
        banner,
        text=diag.get("summary", ""),
        font=("Meiryo UI", 10, "bold"),
        bg=CARD_BG,
        fg=TEXT_COLOR,
        justify="left",
        anchor="w",
        wraplength=win_w - 48,
    ).pack(fill=tk.X, pady=(8, 0))

    # 下部ボタンバー
    btn_bar = tk.Frame(err_win, bg=BG_COLOR, padx=16, pady=10)
    btn_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def _copy_all():
        try:
            text = f"[{diag.get('title', '')}]\n" + format_api_error_message(diag)
            err_win.clipboard_clear()
            err_win.clipboard_append(text)
            btn_copy.configure(text="✅ コピーしました")
            err_win.after(
                2000,
                lambda: btn_copy.winfo_exists() and btn_copy.configure(text="📋 エラー内容をコピー"),
            )
        except Exception:
            pass

    btn_copy = ctk.CTkButton(
        btn_bar,
        text="📋 エラー内容をコピー",
        command=_copy_all,
        fg_color="#E9ECEF",
        hover_color="#DEE2E6",
        text_color=TEXT_COLOR,
        border_width=1,
        border_color=BORDER_COLOR,
        width=170,
    )
    btn_copy.pack(side=tk.LEFT)
    ctk.CTkButton(
        btn_bar,
        text="OK",
        command=_close_err,
        fg_color=PRIMARY,
        hover_color=PRIMARY_HOVER,
        text_color="white",
        width=110,
    ).pack(side=tk.RIGHT)

    # 本文エリア
    body = tk.Frame(err_win, bg=BG_COLOR, padx=14, pady=8)
    body.pack(fill=tk.BOTH, expand=True)

    def _add_section_card(title, items, bar_color):
        if not items:
            return
        card_outer = tk.Frame(body, bg=BORDER_COLOR, padx=1, pady=1)
        card_outer.pack(fill=tk.X, pady=(0, 8))
        card_row = tk.Frame(card_outer, bg=CARD_BG)
        card_row.pack(fill=tk.BOTH, expand=True)
        tk.Frame(card_row, bg=bar_color, width=5).pack(side=tk.LEFT, fill=tk.Y)
        content = tk.Frame(card_row, bg=CARD_BG, padx=12, pady=8)
        content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tk.Label(
            content,
            text=title,
            font=("Segoe UI", 9, "bold"),
            bg=CARD_BG,
            fg=bar_color,
            anchor="w",
        ).pack(fill=tk.X, pady=(0, 4))
        for item in items:
            tk.Label(
                content,
                text=f"・{item}",
                font=("Meiryo UI", 9),
                bg=CARD_BG,
                fg=TEXT_COLOR,
                justify="left",
                anchor="w",
                wraplength=win_w - 75,
            ).pack(fill=tk.X, pady=1)

    _add_section_card("🔍 考えられる原因", diag.get("causes", []), accent_col)
    _add_section_card("💡 対処方法", diag.get("actions", []), PRIMARY)

    # エラー原文エリア
    tk.Label(
        body,
        text="▼ 詳細情報（システムからのエラー原文）",
        font=("Segoe UI", 8, "bold"),
        bg=BG_COLOR,
        fg=MUTED_TEXT,
        anchor="w",
    ).pack(fill=tk.X, pady=(2, 2))
    raw_box = st.ScrolledText(
        body,
        wrap=tk.WORD,
        font=("Consolas", 8),
        bg="#F8FAFC",
        fg=MUTED_TEXT,
        relief=tk.SOLID,
        borderwidth=1,
        height=6,
        padx=8,
        pady=6,
    )
    raw_box.pack(fill=tk.BOTH, expand=True)
    raw_box.insert(tk.END, diag.get("raw", ""))
    raw_box.configure(state=tk.DISABLED)


# ==============================
# チェックボックスリストコンポーネント
# ==============================
class CTkScrollableCheckboxList(ctk.CTkScrollableFrame):
    def __init__(self, master, **kwargs):
        kwargs.setdefault("fg_color", CARD_BG)
        kwargs.setdefault("border_color", BORDER_COLOR)
        super().__init__(master, **kwargs)
        self.items = []

    def set_items(self, item_texts):
        for item in self.items:
            item["cb"].destroy()
        self.items.clear()
        for text in item_texts:
            if str(text).strip():
                self.add_item(str(text).strip())

    def add_item(self, text):
        var = ctk.StringVar(value="")
        cb = ctk.CTkCheckBox(
            self,
            text=text,
            variable=var,
            onvalue=text,
            offvalue="",
            text_color=TEXT_COLOR,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            border_color="#ADB5BD",
            checkbox_width=18,
            checkbox_height=18,
        )
        cb.pack(anchor="w", padx=6, pady=3, fill="x")
        self.items.append({"text": text, "var": var, "cb": cb})

    def get_all_items(self):
        return [item["text"] for item in self.items]

    def get_selected_items(self):
        return [item["text"] for item in self.items if item["var"].get() != ""]

    def remove_selected(self):
        new_items = []
        for item in self.items:
            if item["var"].get() != "":
                item["cb"].destroy()
            else:
                new_items.append(item)
        self.items = new_items


# ==============================
# API詳細設定ダイアログ
# ==============================
class APISettingsDialog(ctk.CTkToplevel):
    def __init__(self, parent, current_settings, on_save_callback):
        super().__init__(parent)
        self.title("⚙️ AI詳細設定 (Gemini API / Azure OpenAI Service)")
        self.configure(fg_color=BG_COLOR)

        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        dialog_w = min(1200, screen_w - 40)
        dialog_h = min(820, screen_h - 80)
        center_window(self, parent, dialog_w, dialog_h)
        self.minsize(920, 640)
        self.transient(parent)
        self.grab_set()

        self.settings = copy.deepcopy(current_settings)
        self.on_save_callback = on_save_callback

        # Gemini モデルリスト初期化
        saved_models = self.settings.get("models_list", [])
        if not saved_models or any(
            not isinstance(item, (list, tuple)) or len(item) < 2 or "1.5" in str(item[1])
            for item in saved_models
        ):
            self.models_list = [list(m) for m in DEFAULT_GEMINI_MODELS]
        else:
            self.models_list = [[str(m[0]), str(m[1])] for m in saved_models]

        # Azure モデルリスト初期化
        saved_az_models = self.settings.get("azure_models_list", [])
        if isinstance(saved_az_models, list) and saved_az_models:
            self.azure_models_list = [str(m) for m in saved_az_models if str(m).strip()]
        else:
            self.azure_models_list = list(DEFAULT_AZURE_MODELS)

        # フィルタ設定
        self.gemini_mainstream_only_var = ctk.BooleanVar(
            value=bool(self.settings.get("gemini_mainstream_only", True))
        )
        self.azure_mainstream_only_var = ctk.BooleanVar(
            value=bool(self.settings.get("azure_mainstream_only", True))
        )
        self.detected_azure_api_ver = (
            self.settings.get("azure_api_version", DEFAULT_AZURE_API_VERSION)
            or DEFAULT_AZURE_API_VERSION
        )
        self.gemini_model_meta = copy.deepcopy(self.settings.get("gemini_model_meta", {}) or {})
        self.azure_model_meta = copy.deepcopy(self.settings.get("azure_model_meta", {}) or {})

        # 実行プロファイル選択（gemini_free / gemini_paid / azure）
        engine = self.settings.get("engine", "gemini")
        plan = self.settings.get("plan", "free")
        initial_profile = (
            "azure"
            if engine == "azure"
            else ("gemini_paid" if plan == "paid" else "gemini_free")
        )
        self.active_profile_var = ctk.StringVar(value=initial_profile)
        self._last_gemini_plan = plan if plan in ("free", "paid") else "free"

        # Gemini タブ変数（数値は StringVar で保持し空欄入力時の TclError を防止）
        def _normalize_prompts(val):
            if isinstance(val, list):
                return [str(x).strip() for x in val if str(x).strip()]
            if isinstance(val, str):
                return [x.strip() for x in val.splitlines() if x.strip()]
            return []

        self.gemini_vars = {}
        for p_type in ("free", "paid"):
            self.gemini_vars[p_type] = {
                "key": ctk.StringVar(value=str(self.settings.get(f"{p_type}_key", ""))),
                "model_step1": ctk.StringVar(
                    value=str(
                        self.settings.get(f"{p_type}_model_step1", DEFAULT_GEMINI_MODEL_STEP1)
                    )
                ),
                "model_step2": ctk.StringVar(
                    value=str(
                        self.settings.get(f"{p_type}_model_step2", DEFAULT_GEMINI_MODEL_STEP2)
                    )
                ),
                "model_step3": ctk.StringVar(
                    value=str(
                        self.settings.get(f"{p_type}_model_step3", DEFAULT_GEMINI_MODEL_STEP3)
                    )
                ),
                "perf_vars": {},  # {model_id: {"rpm": StringVar, "threads": StringVar}}
                "temp": ctk.StringVar(value=str(self.settings.get(f"{p_type}_temp", 0.0))),
                "tokens": ctk.StringVar(value=str(self.settings.get(f"{p_type}_tokens", 65535))),
                "safety": ctk.BooleanVar(value=bool(self.settings.get(f"{p_type}_safety", True))),
                "prompts": _normalize_prompts(self.settings.get(f"{p_type}_prompts", [])),
            }

        # Azure OpenAI タブ変数
        az_dep1 = str(
            self.settings.get(
                "azure_deployment_step1",
                self.settings.get("azure_deployment", DEFAULT_AZURE_DEPLOYMENT_STEP1),
            )
        )
        az_dep2 = str(self.settings.get("azure_deployment_step2", az_dep1 or DEFAULT_AZURE_DEPLOYMENT_STEP2))
        az_dep3 = str(self.settings.get("azure_deployment_step3", DEFAULT_AZURE_DEPLOYMENT_STEP3))
        self.azure_vars = {
            "endpoint": ctk.StringVar(value=str(self.settings.get("azure_endpoint", ""))),
            "key": ctk.StringVar(value=str(self.settings.get("azure_api_key", ""))),
            "deployment_step1": ctk.StringVar(value=az_dep1),
            "deployment_step2": ctk.StringVar(value=az_dep2),
            "deployment_step3": ctk.StringVar(value=az_dep3),
            "api_version": ctk.StringVar(
                value=str(
                    self.settings.get("azure_api_version", DEFAULT_AZURE_API_VERSION)
                    or DEFAULT_AZURE_API_VERSION
                )
            ),
            "rpm": ctk.StringVar(value=str(self.settings.get("azure_rpm", 60))),
            "threads": ctk.StringVar(value=str(self.settings.get("azure_threads", 3))),
            "temp": ctk.StringVar(value=str(self.settings.get("azure_temp", 0.0))),
            "tokens": ctk.StringVar(value=str(self.settings.get("azure_tokens", 8192))),
            "prompts": _normalize_prompts(self.settings.get("azure_prompts", [])),
        }

        self.saved_prompts = _normalize_prompts(self.settings.get("saved_prompts", []))
        self.fav_lists = []
        self.gemini_tab_controls = {}  # {plan_type: {"combos": [cb1, cb2, cb3], "render_perf": fn}}
        self.gemini_count_labels = []
        self.azure_combos = []
        self.azure_count_label = None
        self.tab_reset_handlers = {}

        self._setup_ui()
        self._initial_snapshot = self._collect_settings(normalize_azure=False)

        self.protocol("WM_DELETE_WINDOW", self.cancel_and_close)
        self.after(60, self._bring_to_front)

    def _bring_to_front(self):
        try:
            self.lift()
            self.focus_force()
        except Exception:
            pass

    def _get_visible_gemini_models(self):
        if self.gemini_mainstream_only_var.get():
            return filter_mainstream_gemini_models(self.models_list)
        return self.models_list

    def _get_visible_azure_models(self):
        if self.azure_mainstream_only_var.get():
            return filter_mainstream_azure_models(self.azure_models_list)
        return self.azure_models_list

    def _model_id_to_display(self, model_id):
        for disp, mid in self.models_list:
            if mid == model_id:
                return disp
        return model_id

    def _display_to_model_id(self, display_val):
        val = str(display_val).strip()
        for disp, mid in self.models_list:
            if disp == val or mid == val:
                return mid
        return val

    def _refresh_gemini_combos(self):
        visible = self._get_visible_gemini_models()
        display_names = [m[0] for m in visible]
        count_text = f"({len(visible)} / {len(self.models_list)}件)"
        for lbl in self.gemini_count_labels:
            if lbl.winfo_exists():
                lbl.configure(text=count_text)
        for p_type, ctrl in self.gemini_tab_controls.items():
            for idx, cb in enumerate(ctrl["combos"]):
                step_key = f"model_step{idx + 1}"
                current_id = self.gemini_vars[p_type][step_key].get()
                matched_display = self._model_id_to_display(current_id)
                cb.configure(values=display_names)
                cb.set(matched_display)
            ctrl["render_perf"]()

    def _refresh_azure_combos(self):
        visible = self._get_visible_azure_models()
        for idx, cb in enumerate(self.azure_combos):
            step_key = f"deployment_step{idx + 1}"
            current_val = self.azure_vars[step_key].get()
            cb.configure(values=visible)
            cb.set(current_val)
        if self.azure_count_label and self.azure_count_label.winfo_exists():
            self.azure_count_label.configure(
                text=f"({len(visible)} / {len(self.azure_models_list)}件)"
            )

    def _update_all_fav_lists(self):
        for f_list in self.fav_lists:
            if f_list.winfo_exists():
                f_list.set_items(self.saved_prompts)

    # ==============================
    # UI 構築
    # ==============================
    def _setup_ui(self):
        # --- 下部ボタン領域 ---
        btn_frame = ctk.CTkFrame(
            self,
            fg_color=CARD_BG,
            corner_radius=0,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        btn_frame.pack(side="bottom", fill="x")

        btn_inner = ctk.CTkFrame(btn_frame, fg_color="transparent")
        btn_inner.pack(fill="x", padx=16, pady=10)

        ctk.CTkButton(
            btn_inner,
            text="キャンセル",
            fg_color="#E9ECEF",
            hover_color="#DEE2E6",
            text_color=TEXT_COLOR,
            border_width=1,
            border_color=BORDER_COLOR,
            command=self.cancel_and_close,
            width=120,
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            btn_inner,
            text="設定を適用して保存・閉じる",
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="white",
            font=ctk.CTkFont(weight="bold"),
            command=self.save_and_close,
            width=220,
        ).pack(side="left", padx=(0, 10))

        ctk.CTkLabel(
            btn_inner,
            text="※適用すると設定ファイルへ自動保存されます",
            text_color=MUTED_TEXT,
            font=ctk.CTkFont(size=11),
        ).pack(side="left")

        ctk.CTkButton(
            btn_inner,
            text="🔄 このタブを推奨値に戻す",
            fg_color="#E9ECEF",
            hover_color="#DEE2E6",
            text_color=TEXT_COLOR,
            border_width=1,
            border_color=BORDER_COLOR,
            command=self.reset_current_tab_all,
            width=190,
        ).pack(side="right")

        # --- スクロール可能なメイン領域 ---
        self.scroll_frame = ctk.CTkScrollableFrame(self, fg_color=BG_COLOR, corner_radius=0)
        self.scroll_frame.pack(side="top", fill="both", expand=True, padx=6, pady=6)

        ctk.CTkLabel(
            self.scroll_frame,
            text="AI エンジン詳細設定 (Gemini API / Azure OpenAI Service)",
            font=ctk.CTkFont(size=19, weight="bold"),
            text_color=PRIMARY,
        ).pack(pady=(8, 6))

        # --- 実際に使用するAIエンジン / プランの選択 ---
        plan_frame = ctk.CTkFrame(
            self.scroll_frame,
            fg_color=CARD_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        plan_frame.pack(fill="x", padx=12, pady=(0, 8))

        ctk.CTkLabel(
            plan_frame,
            text=" 実際に使用するAIエンジン / プランの選択 ",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=PRIMARY,
        ).pack(anchor="w", padx=12, pady=(8, 2))

        plan_inner = ctk.CTkFrame(plan_frame, fg_color="transparent")
        plan_inner.pack(fill="x", padx=12, pady=(2, 10))

        ctk.CTkLabel(
            plan_inner,
            text="解析時に使用するAI設定を選んでください（下のタブとは連動しません）:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(side="left", padx=(0, 14))

        for label_text, val in [
            ("Gemini API (無料枠)", "gemini_free"),
            ("Gemini API (課金枠)", "gemini_paid"),
            ("Azure OpenAI Service", "azure"),
        ]:
            ctk.CTkRadioButton(
                plan_inner,
                text=label_text,
                variable=self.active_profile_var,
                value=val,
                text_color=TEXT_COLOR,
                fg_color=PRIMARY,
                hover_color=PRIMARY_HOVER,
                border_color="#ADB5BD",
                font=ctk.CTkFont(size=12, weight="bold"),
            ).pack(side="left", padx=(0, 16))

        # --- タブ (3タブ構成) ---
        self.tabview = ctk.CTkTabview(
            self.scroll_frame,
            fg_color=BG_COLOR,
            segmented_button_fg_color="#DEE2E6",
            segmented_button_selected_color=PRIMARY,
            segmented_button_selected_hover_color=PRIMARY_HOVER,
            segmented_button_unselected_color="#6C757D",
            segmented_button_unselected_hover_color="#5C636A",
            text_color="white",
            text_color_disabled="#ADB5BD",
            border_width=1,
            border_color=BORDER_COLOR,
        )
        self.tabview.pack(fill="both", expand=True, padx=12, pady=(2, 10))

        tab_free = self.tabview.add(TAB_GEMINI_FREE)
        tab_paid = self.tabview.add(TAB_GEMINI_PAID)
        tab_azure = self.tabview.add(TAB_AZURE)

        self._build_gemini_tab(tab_free, "free")
        self._build_gemini_tab(tab_paid, "paid")
        self._build_azure_tab(tab_azure)

        self._update_all_fav_lists()

        # 現在アクティブなプロファイルのタブを初期選択
        prof = self.active_profile_var.get()
        if prof == "azure":
            self.tabview.set(TAB_AZURE)
        elif prof == "gemini_paid":
            self.tabview.set(TAB_GEMINI_PAID)
        else:
            self.tabview.set(TAB_GEMINI_FREE)

    # ==============================
    # 共通: カスタムプロンプトセクション
    # ==============================
    def _build_prompt_section(self, parent_tab, section_title, get_prompts_cb, set_prompts_cb):
        prompt_frame = ctk.CTkFrame(
            parent_tab,
            fg_color=CARD_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        prompt_frame.pack(fill="both", expand=True, padx=6, pady=6)

        ctk.CTkLabel(
            prompt_frame,
            text=section_title,
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=PRIMARY,
        ).pack(anchor="w", padx=12, pady=(8, 4))

        input_inner = ctk.CTkFrame(prompt_frame, fg_color="transparent")
        input_inner.pack(fill="x", padx=12, pady=(2, 8))

        entry_new_prompt = ctk.CTkEntry(
            input_inner,
            placeholder_text="新しい追加指示を入力して Enter または「＋ 指示を追加」...",
            fg_color="white",
            border_color="#CED4DA",
            text_color=TEXT_COLOR,
        )
        entry_new_prompt.pack(side="left", fill="x", expand=True, padx=(0, 8))
        bind_context_menu(entry_new_prompt)

        lists_frame = ctk.CTkFrame(prompt_frame, fg_color="transparent")
        lists_frame.pack(fill="both", expand=True, padx=12, pady=(0, 10))
        lists_frame.grid_columnconfigure(0, weight=1)
        lists_frame.grid_columnconfigure(1, weight=1)

        # 左側: 現在使用する指示
        left_frame = ctk.CTkFrame(lists_frame, fg_color="transparent")
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        ctk.CTkLabel(
            left_frame,
            text="▼ 現在の解析に使用する指示",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(anchor="w")

        current_list = CTkScrollableCheckboxList(left_frame, height=130, border_width=1)
        current_list.pack(fill="both", expand=True, pady=(4, 6))

        def sync_current_to_var():
            set_prompts_cb(current_list.get_all_items())

        def add_current_prompt(e=None):
            text = entry_new_prompt.get().strip()
            if text:
                current_list.add_item(text)
                entry_new_prompt.delete(0, "end")
                sync_current_to_var()

        entry_new_prompt.bind("<Return>", add_current_prompt)
        ctk.CTkButton(
            input_inner,
            text="＋ 指示を追加",
            command=add_current_prompt,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="white",
            width=120,
        ).pack(side="left")

        left_actions = ctk.CTkFrame(left_frame, fg_color="transparent")
        left_actions.pack(fill="x")

        def delete_current_selected():
            current_list.remove_selected()
            sync_current_to_var()

        def save_selected_to_fav():
            sel = current_list.get_selected_items()
            if not sel:
                messagebox.showinfo("情報", "お気に入りに保存したい指示にチェックを入れてください。", parent=self)
                return
            added = 0
            for text in sel:
                if text not in self.saved_prompts:
                    self.saved_prompts.append(text)
                    added += 1
            if added > 0:
                self._update_all_fav_lists()
                messagebox.showinfo("保存完了", f"{added}件の指示をお気に入りに保存しました。", parent=self)
            else:
                messagebox.showinfo("情報", "選択された指示は既にお気に入りに保存されています。", parent=self)

        ctk.CTkButton(
            left_actions,
            text="🗑 選択を削除",
            command=delete_current_selected,
            fg_color="#E9ECEF",
            hover_color="#DEE2E6",
            text_color=TEXT_COLOR,
            border_width=1,
            border_color=BORDER_COLOR,
            width=110,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            left_actions,
            text="⭐ 選択をお気に入りに保存",
            command=save_selected_to_fav,
            fg_color="#FFF3CD",
            hover_color="#FFE69C",
            text_color="#664D03",
            border_width=1,
            border_color="#FFDA6A",
            width=175,
        ).pack(side="left")

        # 右側: お気に入り
        right_frame = ctk.CTkFrame(lists_frame, fg_color="transparent")
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        ctk.CTkLabel(
            right_frame,
            text="⭐ お気に入り (よく使う指示・全タブ共通)",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#B26A00",
        ).pack(anchor="w")

        fav_list = CTkScrollableCheckboxList(right_frame, height=130, border_width=1)
        fav_list.pack(fill="both", expand=True, pady=(4, 6))
        self.fav_lists.append(fav_list)

        right_actions = ctk.CTkFrame(right_frame, fg_color="transparent")
        right_actions.pack(fill="x")

        def add_fav_to_current():
            sel = fav_list.get_selected_items()
            if not sel:
                return
            for text in sel:
                current_list.add_item(text)
            sync_current_to_var()

        def delete_fav_selected():
            sel = fav_list.get_selected_items()
            if not sel:
                return
            if messagebox.askyesno("確認", "選択したお気に入りを削除しますか？", parent=self):
                for text in sel:
                    if text in self.saved_prompts:
                        self.saved_prompts.remove(text)
                self._update_all_fav_lists()

        ctk.CTkButton(
            right_actions,
            text="◀ 選択を左に追加",
            command=add_fav_to_current,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="white",
            width=135,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            right_actions,
            text="🗑 選択を削除",
            command=delete_fav_selected,
            fg_color="#E9ECEF",
            hover_color="#DEE2E6",
            text_color=TEXT_COLOR,
            border_width=1,
            border_color=BORDER_COLOR,
            width=110,
        ).pack(side="left")

        current_list.set_items(get_prompts_cb())

    # ==============================
    # Gemini タブ構築 (Free / Paid)
    # ==============================
    def _build_gemini_tab(self, parent_tab, plan_type):
        vars_dict = self.gemini_vars[plan_type]
        is_free = plan_type == "free"
        tab_key = TAB_GEMINI_FREE if is_free else TAB_GEMINI_PAID

        # --- ① APIキー ---
        key_frame = ctk.CTkFrame(
            parent_tab,
            fg_color=CARD_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        key_frame.pack(fill="x", padx=6, pady=5)

        ctk.CTkLabel(
            key_frame,
            text=" ① APIキー ",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=PRIMARY,
        ).pack(anchor="w", padx=12, pady=(8, 2))

        key_inner = ctk.CTkFrame(key_frame, fg_color="transparent")
        key_inner.pack(fill="x", padx=12, pady=(2, 10))

        ctk.CTkLabel(
            key_inner,
            text=f"{'無料枠 (Free)' if is_free else '課金枠 (Paid)'} 用のAPIキー:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(side="left", padx=(0, 8))

        entry_key = ctk.CTkEntry(
            key_inner,
            textvariable=vars_dict["key"],
            show="*",
            placeholder_text="AIza...",
            fg_color="white",
            border_color="#CED4DA",
            text_color=TEXT_COLOR,
        )
        entry_key.pack(side="left", fill="x", expand=True, padx=(0, 6))
        bind_context_menu(entry_key)

        btn_toggle = ctk.CTkButton(
            key_inner,
            text="確認",
            width=64,
            fg_color="#E9ECEF",
            hover_color="#DEE2E6",
            text_color=TEXT_COLOR,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        btn_toggle.pack(side="left", padx=(0, 6))

        def toggle_key(e=entry_key, b=btn_toggle):
            if e.cget("show") == "*":
                e.configure(show="")
                b.configure(text="隠す")
            else:
                e.configure(show="*")
                b.configure(text="確認")

        btn_toggle.configure(command=toggle_key)

        def run_gemini_test(k_var, get_model_id_fn, btn, default_text="テスト"):
            key = k_var.get().strip()
            if not key:
                messagebox.showwarning("警告", "APIキーが入力されていません。", parent=self)
                return
            model_name = get_model_id_fn()
            if not model_name:
                messagebox.showwarning("警告", "テストするモデルが選択されていません。", parent=self)
                return

            btn.configure(state="disabled", text="確認中...")

            def _worker():
                try:
                    test_gemini_connection(key, model_name)
                    self.after(
                        0,
                        lambda: messagebox.showinfo(
                            "テスト成功",
                            f"APIキーは正しく認識されました。\nAIモデル「{model_name}」による通信は正常です！",
                            parent=self,
                        ),
                    )
                except Exception as e:
                    diag = diagnose_api_error("gemini", e, model_or_deployment=model_name)
                    self.after(0, lambda d=diag: show_api_error_dialog(self, d))
                finally:
                    self.after(
                        0,
                        lambda: btn.winfo_exists() and btn.configure(state="normal", text=default_text),
                    )

            threading.Thread(target=_worker, daemon=True).start()

        btn_test_main = ctk.CTkButton(
            key_inner,
            text="🔌 接続テスト",
            width=110,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="white",
        )
        btn_test_main.configure(
            command=lambda: run_gemini_test(
                vars_dict["key"],
                lambda: self._display_to_model_id(current_plan_combos[0].get()),
                btn_test_main,
                "🔌 接続テスト",
            )
        )
        btn_test_main.pack(side="left")

        # --- 中段左右分割 ---
        middle_frame = ctk.CTkFrame(parent_tab, fg_color="transparent")
        middle_frame.pack(fill="x", padx=6, pady=5)
        middle_frame.grid_columnconfigure(0, weight=3)
        middle_frame.grid_columnconfigure(1, weight=2)

        # --- ② 使用モデル設定 (ステップ別) ---
        model_frame = ctk.CTkFrame(
            middle_frame,
            fg_color=CARD_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        model_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        ctk.CTkLabel(
            model_frame,
            text=" ② 使用モデル設定 (ステップ別) ",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=PRIMARY,
        ).pack(anchor="w", padx=12, pady=(8, 4))

        model_inner = ctk.CTkFrame(model_frame, fg_color="transparent")
        model_inner.pack(fill="x", padx=12, pady=(2, 4))

        current_plan_combos = []
        step_labels = ["Step1 (一覧抽出):", "Step2 (要素抽出):", "Step3 (最終検証):"]
        var_keys = ["model_step1", "model_step2", "model_step3"]
        visible_models = self._get_visible_gemini_models()
        display_names = [m[0] for m in visible_models]

        def _on_step_model_change(selected_display, v_key):
            mid = self._display_to_model_id(selected_display)
            vars_dict[v_key].set(mid)
            if mid not in vars_dict["perf_vars"]:
                rec = get_recommended_gemini_settings(
                    mid, is_free=is_free, model_meta=self.gemini_model_meta
                )
                vars_dict["perf_vars"][mid] = {
                    "rpm": ctk.StringVar(value=str(rec["rpm"])),
                    "threads": ctk.StringVar(value=str(rec["threads"])),
                }
            # 選択中ステップのトークン上限に合わせて最大トークンを安全側へ自動最適化
            opt_params = get_optimized_gemini_params_for_steps(
                [
                    vars_dict["model_step1"].get(),
                    vars_dict["model_step2"].get(),
                    vars_dict["model_step3"].get(),
                ],
                is_free=is_free,
                model_meta=self.gemini_model_meta,
            )
            cur_tok = safe_int(vars_dict["tokens"], opt_params["max_tokens"], 1024, 2097152)
            if cur_tok > opt_params["max_tokens"]:
                vars_dict["tokens"].set(str(opt_params["max_tokens"]))
            render_perf_settings()

        for label_text, v_key in zip(step_labels, var_keys):
            step_row = ctk.CTkFrame(model_inner, fg_color="transparent")
            step_row.pack(fill="x", pady=3)

            ctk.CTkLabel(
                step_row,
                text=label_text,
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=TEXT_COLOR,
                width=120,
                anchor="w",
            ).pack(side="left")

            cb = ctk.CTkComboBox(
                step_row,
                values=display_names,
                fg_color="white",
                border_color="#CED4DA",
                button_color="#E9ECEF",
                button_hover_color="#DEE2E6",
                dropdown_fg_color="white",
                dropdown_hover_color="#E7F1FF",
                dropdown_text_color=TEXT_COLOR,
                text_color=TEXT_COLOR,
                command=lambda val, k=v_key: _on_step_model_change(val, k),
            )
            cb.pack(side="left", fill="x", expand=True, padx=(0, 6))
            bind_context_menu(cb)

            curr_id = vars_dict[v_key].get()
            cb.set(self._model_id_to_display(curr_id))

            # 手入力後にフォーカスが外れた場合も同期
            try:
                cb._entry.bind(
                    "<FocusOut>",
                    lambda e, c=cb, k=v_key: _on_step_model_change(c.get(), k),
                )
            except Exception:
                pass

            btn_step_test = ctk.CTkButton(
                step_row,
                text="テスト",
                width=58,
                fg_color="#E9ECEF",
                hover_color="#DEE2E6",
                text_color=TEXT_COLOR,
                border_width=1,
                border_color=BORDER_COLOR,
            )
            btn_step_test.configure(
                command=lambda c=cb, b=btn_step_test: run_gemini_test(
                    vars_dict["key"],
                    lambda cb_ref=c: self._display_to_model_id(cb_ref.get()),
                    b,
                    "テスト",
                )
            )
            btn_step_test.pack(side="left")

            current_plan_combos.append(cb)

        # モデル更新・主力フィルタ・公式リンク行
        action_row1 = ctk.CTkFrame(model_frame, fg_color="transparent")
        action_row1.pack(fill="x", padx=12, pady=(6, 4))

        def fetch_models_ui(btn):
            key = vars_dict["key"].get().strip()
            if not key:
                messagebox.showwarning(
                    "警告", "最新モデルを取得するには、APIキーを入力してください。", parent=self
                )
                return
            btn.configure(text="取得中...", state="disabled")

            def _do_fetch():
                try:
                    new_models, new_meta = fetch_gemini_models_with_meta(key)
                    if not new_models:
                        self.after(
                            0,
                            lambda: messagebox.showinfo(
                                "情報", "取得可能なモデルが見つかりませんでした。", parent=self
                            ),
                        )
                        return

                    def _apply_models():
                        self.models_list = [[m[0], m[1]] for m in new_models]
                        self.gemini_model_meta.update(new_meta)
                        self._refresh_gemini_combos()

                        vis_cnt = len(self._get_visible_gemini_models())
                        mode_note = (
                            f"（主力モデル {vis_cnt}件 を表示中。すべて見る場合は「主力モデルのみ表示」のチェックを外してください）"
                            if self.gemini_mainstream_only_var.get()
                            else f"（全 {vis_cnt}件 を表示中）"
                        )

                        # 無料枠・課金枠の両方でパフォーマンス設定・AI抽出パラメータを最新モデル仕様に合わせて自動最適化
                        changes_summary = []
                        for target_plan in ("free", "paid"):
                            t_is_free = target_plan == "free"
                            t_vars = self.gemini_vars[target_plan]
                            t_step_models = [
                                t_vars["model_step1"].get().strip(),
                                t_vars["model_step2"].get().strip(),
                                t_vars["model_step3"].get().strip(),
                            ]

                            # 1. 全モデルのパフォーマンス設定（RPM / スレッド）を最適化
                            for mid, p_v in t_vars["perf_vars"].items():
                                rec = get_recommended_gemini_settings(
                                    mid, is_free=t_is_free, model_meta=self.gemini_model_meta
                                )
                                old_rpm = safe_int(p_v["rpm"], rec["rpm"], 1, 2000)
                                old_th = safe_int(p_v["threads"], rec["threads"], 1, 20)
                                p_v["rpm"].set(str(rec["rpm"]))
                                p_v["threads"].set(str(rec["threads"]))
                                if target_plan == plan_type and mid in t_step_models:
                                    if old_rpm != rec["rpm"] or old_th != rec["threads"]:
                                        changes_summary.append(
                                            f"・{mid}: RPM {old_rpm}→{rec['rpm']} / スレッド {old_th}→{rec['threads']}"
                                        )

                            # 2. AI抽出パラメータ設定（Temp / 最大トークン / 安全フィルタ）を最適化
                            opt_params = get_optimized_gemini_params_for_steps(
                                t_step_models,
                                is_free=t_is_free,
                                model_meta=self.gemini_model_meta,
                            )
                            old_tp = safe_float(t_vars["temp"], opt_params["temperature"], 0.0, 2.0)
                            old_tok = safe_int(t_vars["tokens"], opt_params["max_tokens"], 1024, 2097152)
                            old_sf = bool(t_vars["safety"].get())
                            t_vars["temp"].set(str(opt_params["temperature"]))
                            t_vars["tokens"].set(str(opt_params["max_tokens"]))
                            t_vars["safety"].set(opt_params["disable_safety"])

                            if target_plan == plan_type:
                                if abs(old_tp - opt_params["temperature"]) > 1e-6:
                                    changes_summary.append(
                                        f"・Temp: {old_tp} → {opt_params['temperature']}"
                                    )
                                if old_tok != opt_params["max_tokens"]:
                                    changes_summary.append(
                                        f"・最大トークン: {old_tok} → {opt_params['max_tokens']}"
                                    )
                                if old_sf != opt_params["disable_safety"]:
                                    changes_summary.append(
                                        f"・安全フィルタ無効化: {'ON' if opt_params['disable_safety'] else 'OFF'}"
                                    )

                        if changes_summary:
                            msg = (
                                f"最新のモデルリスト (全{len(self.models_list)}件) を Gemini API から取得し、\n"
                                f"パフォーマンス設定およびAI抽出パラメータを自動最適化しました！\n"
                                f"{mode_note}\n\n"
                                f"【最適化された設定】\n"
                                + "\n".join(changes_summary)
                            )
                        else:
                            msg = (
                                f"最新のモデルリスト (全{len(self.models_list)}件) を Gemini API から取得しました！\n"
                                f"{mode_note}\n\n"
                                f"✅ パフォーマンス設定（RPM / スレッド）およびAI抽出パラメータは最新モデルの最適値に調整済みです。"
                            )
                        messagebox.showinfo("更新・最適化完了", msg, parent=self)

                    self.after(0, _apply_models)
                except Exception as e:
                    diag = diagnose_api_error("gemini", e)
                    self.after(0, lambda d=diag: show_api_error_dialog(self, d))
                finally:
                    self.after(
                        0,
                        lambda: btn.winfo_exists()
                        and btn.configure(text="🌐 最新モデルを更新", state="normal"),
                    )

            threading.Thread(target=_do_fetch, daemon=True).start()

        btn_fetch = ctk.CTkButton(
            action_row1,
            text="🌐 最新モデルを更新",
            width=155,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="white",
        )
        btn_fetch.configure(command=lambda b=btn_fetch: fetch_models_ui(b))
        btn_fetch.pack(side="left", padx=(0, 10))

        ctk.CTkCheckBox(
            action_row1,
            text="主力モデルのみ表示",
            variable=self.gemini_mainstream_only_var,
            command=self._refresh_gemini_combos,
            text_color=TEXT_COLOR,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            border_color="#ADB5BD",
            checkbox_width=18,
            checkbox_height=18,
        ).pack(side="left")

        lbl_cnt = ctk.CTkLabel(
            action_row1,
            text=f"({len(visible_models)} / {len(self.models_list)}件)",
            text_color=MUTED_TEXT,
            font=ctk.CTkFont(size=11),
        )
        lbl_cnt.pack(side="left", padx=(4, 0))
        self.gemini_count_labels.append(lbl_cnt)

        action_row2 = ctk.CTkFrame(model_frame, fg_color="transparent")
        action_row2.pack(fill="x", padx=12, pady=(0, 8))

        lbl_link = ctk.CTkLabel(
            action_row2,
            text="🔗 各モデルの特徴を確認する (公式)",
            text_color=PRIMARY,
            cursor="hand2",
            font=ctk.CTkFont(size=12, underline=True),
        )
        lbl_link.pack(side="right")
        lbl_link.bind(
            "<Button-1>",
            lambda e: webbrowser.open_new("https://ai.google.dev/gemini-api/docs/models/gemini"),
        )

        # --- ③ AI抽出パラメータ設定 (右側) ---
        param_frame = ctk.CTkFrame(
            middle_frame,
            fg_color=CARD_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        param_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        ctk.CTkLabel(
            param_frame,
            text=" ③ AI抽出パラメータ設定 ",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=PRIMARY,
        ).pack(anchor="w", padx=12, pady=(8, 4))

        param_row1 = ctk.CTkFrame(param_frame, fg_color="transparent")
        param_row1.pack(fill="x", padx=12, pady=(10, 6))

        ctk.CTkLabel(
            param_row1, text="Temp:", font=ctk.CTkFont(size=12, weight="bold"), text_color=TEXT_COLOR
        ).pack(side="left")
        entry_temp = ctk.CTkEntry(
            param_row1,
            textvariable=vars_dict["temp"],
            width=62,
            fg_color="white",
            border_color="#CED4DA",
            text_color=TEXT_COLOR,
        )
        entry_temp.pack(side="left", padx=(6, 14))

        ctk.CTkLabel(
            param_row1,
            text="最大トークン:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(side="left")
        entry_tokens = ctk.CTkEntry(
            param_row1,
            textvariable=vars_dict["tokens"],
            width=86,
            fg_color="white",
            border_color="#CED4DA",
            text_color=TEXT_COLOR,
        )
        entry_tokens.pack(side="left", padx=(6, 0))

        param_row2 = ctk.CTkFrame(param_frame, fg_color="transparent")
        param_row2.pack(fill="x", padx=12, pady=6)
        ctk.CTkCheckBox(
            param_row2,
            text="安全フィルタ無効化 (エラー回避)",
            variable=vars_dict["safety"],
            text_color=TEXT_COLOR,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            border_color="#ADB5BD",
            checkbox_width=18,
            checkbox_height=18,
        ).pack(side="left")

        param_row3 = ctk.CTkFrame(param_frame, fg_color="transparent")
        param_row3.pack(fill="x", padx=12, pady=(12, 10))

        def reset_param():
            opt = get_optimized_gemini_params_for_steps(
                [
                    vars_dict["model_step1"].get(),
                    vars_dict["model_step2"].get(),
                    vars_dict["model_step3"].get(),
                ],
                is_free=is_free,
                model_meta=self.gemini_model_meta,
            )
            vars_dict["temp"].set(str(opt["temperature"]))
            vars_dict["tokens"].set(str(opt["max_tokens"]))
            vars_dict["safety"].set(opt["disable_safety"])

        ctk.CTkButton(
            param_row3,
            text="🔄 推奨値",
            width=95,
            command=reset_param,
            fg_color="#E9ECEF",
            hover_color="#DEE2E6",
            text_color=TEXT_COLOR,
            border_width=1,
            border_color=BORDER_COLOR,
        ).pack(side="right")

        # --- ④ パフォーマンス設定 (モデル別 API制限) ---
        perf_container = ctk.CTkFrame(
            parent_tab,
            fg_color=CARD_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        perf_container.pack(fill="x", padx=6, pady=5)

        ctk.CTkLabel(
            perf_container,
            text=" ④ パフォーマンス設定 (モデル別 API制限 / RPM・スレッド) ",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=PRIMARY,
        ).pack(anchor="w", padx=12, pady=(8, 4))

        perf_list_frame = ctk.CTkScrollableFrame(
            perf_container,
            fg_color="#F8FAFC",
            border_width=1,
            border_color=BORDER_COLOR,
            height=155,
        )
        perf_list_frame.pack(fill="x", padx=12, pady=(2, 6))

        saved_perf_data = self.settings.get(f"{plan_type}_model_perf", {}) or {}

        def _ensure_perf_var(m_id):
            if m_id not in vars_dict["perf_vars"]:
                saved_item = saved_perf_data.get(m_id)
                rec = get_recommended_gemini_settings(
                    m_id, is_free=is_free, model_meta=self.gemini_model_meta
                )
                def_rpm = saved_item.get("rpm", rec["rpm"]) if isinstance(saved_item, dict) else rec["rpm"]
                def_thr = (
                    saved_item.get("threads", rec["threads"])
                    if isinstance(saved_item, dict)
                    else rec["threads"]
                )
                vars_dict["perf_vars"][m_id] = {
                    "rpm": ctk.StringVar(value=str(def_rpm)),
                    "threads": ctk.StringVar(value=str(def_thr)),
                }

        def render_perf_settings():
            for w in perf_list_frame.winfo_children():
                w.destroy()

            active_ids = {
                vars_dict["model_step1"].get().strip(),
                vars_dict["model_step2"].get().strip(),
                vars_dict["model_step3"].get().strip(),
            }
            active_ids.discard("")

            # 表示対象 = 現在のフィルタ適用後モデル ＋ ステップ選択中のモデル
            rows_to_show = []
            seen_ids = set()
            for m_disp, m_id in self._get_visible_gemini_models():
                rows_to_show.append((m_disp, m_id))
                seen_ids.add(m_id)
            for mid in sorted(active_ids):
                if mid not in seen_ids:
                    rows_to_show.append((self._model_id_to_display(mid), mid))
                    seen_ids.add(mid)

            for m_disp, m_id in rows_to_show:
                _ensure_perf_var(m_id)
                is_used = m_id in active_ids

                row_frame = ctk.CTkFrame(
                    perf_list_frame,
                    fg_color="#E7F1FF" if is_used else "transparent",
                    corner_radius=6,
                )
                row_frame.pack(fill="x", padx=4, pady=2)

                label_prefix = "★ " if is_used else "   "
                ctk.CTkLabel(
                    row_frame,
                    text=f"{label_prefix}{m_disp}",
                    anchor="w",
                    width=340,
                    font=ctk.CTkFont(size=12, weight="bold" if is_used else "normal"),
                    text_color=PRIMARY if is_used else TEXT_COLOR,
                ).pack(side="left", padx=(6, 10), pady=3)

                p_vars = vars_dict["perf_vars"][m_id]
                ctk.CTkLabel(
                    row_frame, text="RPM:", font=ctk.CTkFont(size=11, weight="bold"), text_color=TEXT_COLOR
                ).pack(side="left")
                ctk.CTkEntry(
                    row_frame,
                    textvariable=p_vars["rpm"],
                    width=64,
                    height=26,
                    fg_color="white",
                    border_color="#CED4DA",
                    text_color=TEXT_COLOR,
                ).pack(side="left", padx=(4, 18), pady=3)

                ctk.CTkLabel(
                    row_frame,
                    text="スレッド:",
                    font=ctk.CTkFont(size=11, weight="bold"),
                    text_color=TEXT_COLOR,
                ).pack(side="left")
                ctk.CTkEntry(
                    row_frame,
                    textvariable=p_vars["threads"],
                    width=54,
                    height=26,
                    fg_color="white",
                    border_color="#CED4DA",
                    text_color=TEXT_COLOR,
                ).pack(side="left", padx=(4, 6), pady=3)

                if is_used:
                    ctk.CTkLabel(
                        row_frame,
                        text="ステップ使用中",
                        font=ctk.CTkFont(size=10, weight="bold"),
                        text_color=PRIMARY,
                    ).pack(side="right", padx=8)

        render_perf_settings()

        perf_action_inner = ctk.CTkFrame(perf_container, fg_color="transparent")
        perf_action_inner.pack(fill="x", padx=12, pady=(2, 10))

        ctk.CTkButton(
            perf_action_inner,
            text="ℹ️ 制限と仕様を確認",
            fg_color="#E9ECEF",
            hover_color="#DEE2E6",
            text_color=TEXT_COLOR,
            border_width=1,
            border_color=BORDER_COLOR,
            command=lambda f=is_free: self._show_gemini_limit_info(f),
            width=155,
        ).pack(side="left")

        def reset_perf_only():
            for m_id, p_vars in vars_dict["perf_vars"].items():
                rec = get_recommended_gemini_settings(
                    m_id, is_free=is_free, model_meta=self.gemini_model_meta
                )
                p_vars["rpm"].set(str(rec["rpm"]))
                p_vars["threads"].set(str(rec["threads"]))

        ctk.CTkButton(
            perf_action_inner,
            text="🔄 パフォーマンスの推奨値セット",
            fg_color="#E9ECEF",
            hover_color="#DEE2E6",
            text_color=TEXT_COLOR,
            border_width=1,
            border_color=BORDER_COLOR,
            command=reset_perf_only,
            width=210,
        ).pack(side="right")

        def _reset_gemini_tab_all():
            rec_m1, rec_m2, rec_m3 = pick_recommended_gemini_step_models(self.models_list)
            defaults = [rec_m1, rec_m2, rec_m3]
            for idx, (v_key, def_id) in enumerate(zip(var_keys, defaults)):
                vars_dict[v_key].set(def_id)
                current_plan_combos[idx].set(self._model_id_to_display(def_id))
                _ensure_perf_var(def_id)
            reset_perf_only()
            reset_param()
            render_perf_settings()
            messagebox.showinfo(
                "推奨値に戻しました",
                f"使用モデル（Step1: {rec_m1} / Step2: {rec_m2} / Step3: {rec_m3}）・\n"
                f"RPM・スレッド・パラメータ設定をすべて最新の最適化推奨値に戻しました。",
                parent=self,
            )

        self.tab_reset_handlers[tab_key] = _reset_gemini_tab_all
        self.gemini_tab_controls[plan_type] = {
            "combos": current_plan_combos,
            "render_perf": render_perf_settings,
        }

        # --- ⑤ カスタムプロンプト ---
        self._build_prompt_section(
            parent_tab,
            " ⑤ 独自の追加指示 (カスタムプロンプト) - 任意 ",
            lambda: vars_dict["prompts"],
            lambda items: vars_dict.__setitem__("prompts", items),
        )

    # ==============================
    # Azure OpenAI Service タブ構築
    # ==============================
    def _build_azure_tab(self, parent_tab):
        az = self.azure_vars

        # --- ① Azure OpenAI 接続設定 ---
        conn_frame = ctk.CTkFrame(
            parent_tab,
            fg_color=CARD_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        conn_frame.pack(fill="x", padx=6, pady=5)

        ctk.CTkLabel(
            conn_frame,
            text=" ① Azure OpenAI 接続設定 (エンドポイント / APIキー / デプロイ名) ",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=PRIMARY,
        ).pack(anchor="w", padx=12, pady=(8, 4))

        # エンドポイントURL
        row_ep = ctk.CTkFrame(conn_frame, fg_color="transparent")
        row_ep.pack(fill="x", padx=12, pady=3)
        ctk.CTkLabel(
            row_ep,
            text="エンドポイントURL:",
            width=140,
            anchor="w",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(side="left")
        entry_ep = ctk.CTkEntry(
            row_ep,
            textvariable=az["endpoint"],
            placeholder_text="https://your-resource.openai.azure.com/",
            fg_color="white",
            border_color="#CED4DA",
            text_color=TEXT_COLOR,
        )
        entry_ep.pack(side="left", fill="x", expand=True, padx=(0, 8))
        bind_context_menu(entry_ep)
        ctk.CTkLabel(
            row_ep,
            text="例: https://your-resource.openai.azure.com/",
            text_color=MUTED_TEXT,
            font=ctk.CTkFont(size=11),
        ).pack(side="left")

        # APIキー
        row_key = ctk.CTkFrame(conn_frame, fg_color="transparent")
        row_key.pack(fill="x", padx=12, pady=3)
        ctk.CTkLabel(
            row_key,
            text="APIキー:",
            width=140,
            anchor="w",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(side="left")
        entry_az_key = ctk.CTkEntry(
            row_key,
            textvariable=az["key"],
            show="*",
            placeholder_text="KEY 1 または KEY 2 を貼り付け...",
            fg_color="white",
            border_color="#CED4DA",
            text_color=TEXT_COLOR,
        )
        entry_az_key.pack(side="left", fill="x", expand=True, padx=(0, 6))
        bind_context_menu(entry_az_key)

        btn_toggle_az = ctk.CTkButton(
            row_key,
            text="確認",
            width=64,
            fg_color="#E9ECEF",
            hover_color="#DEE2E6",
            text_color=TEXT_COLOR,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        btn_toggle_az.pack(side="left")

        def toggle_az_key():
            if entry_az_key.cget("show") == "*":
                entry_az_key.configure(show="")
                btn_toggle_az.configure(text="隠す")
            else:
                entry_az_key.configure(show="*")
                btn_toggle_az.configure(text="確認")

        btn_toggle_az.configure(command=toggle_az_key)

        def _normalize_azure_fields(target_step_key=None):
            raw_ep = az["endpoint"].get().strip()
            dep_val = az[target_step_key].get().strip() if target_step_key else az["deployment_step1"].get().strip()
            ver_val = az["api_version"].get().strip()
            base_ep, resolved_dep, resolved_ver = parse_azure_endpoint_info(raw_ep, dep_val, ver_val)
            if base_ep:
                az["endpoint"].set(base_ep)
            if resolved_dep:
                if target_step_key and not dep_val:
                    az[target_step_key].set(resolved_dep)
                elif not target_step_key:
                    for sk in ("deployment_step1", "deployment_step2", "deployment_step3"):
                        if not az[sk].get().strip():
                            az[sk].set(resolved_dep)
            if resolved_ver:
                az["api_version"].set(resolved_ver)
            return base_ep, resolved_dep, resolved_ver

        def run_azure_test(step_key, btn, default_text="🔌 接続テスト"):
            base_ep, resolved_dep, resolved_ver = _normalize_azure_fields(step_key)
            key = az["key"].get().strip()

            if not base_ep:
                messagebox.showwarning("警告", "エンドポイントURLを入力してください。", parent=self)
                return
            if not key:
                messagebox.showwarning("警告", "APIキーを入力してください。", parent=self)
                return
            if not resolved_dep:
                messagebox.showwarning("警告", "デプロイ名（モデルデプロイ名）を入力してください。", parent=self)
                return

            btn.configure(state="disabled", text="通信中...")

            def _worker():
                try:
                    test_azure_connection(key, base_ep, resolved_dep, resolved_ver)
                    self.after(
                        0,
                        lambda: messagebox.showinfo(
                            "テスト成功",
                            f"Azure OpenAI Service への接続に成功しました！\n\n"
                            f"・エンドポイント: {base_ep}\n"
                            f"・デプロイ名: {resolved_dep}\n"
                            f"・APIバージョン: {resolved_ver}",
                            parent=self,
                        ),
                    )
                except Exception as e:
                    diag = diagnose_api_error(
                        "azure", e, model_or_deployment=resolved_dep, endpoint=base_ep
                    )
                    self.after(0, lambda d=diag: show_api_error_dialog(self, d))
                finally:
                    self.after(
                        0,
                        lambda: btn.winfo_exists() and btn.configure(state="normal", text=default_text),
                    )

            threading.Thread(target=_worker, daemon=True).start()

        # ステップ別デプロイ名
        visible_az = self._get_visible_azure_models()
        step_defs = [
            ("Step1 (一覧抽出):", "deployment_step1"),
            ("Step2 (要素抽出):", "deployment_step2"),
            ("Step3 (最終検証):", "deployment_step3"),
        ]

        def _get_current_az_optimized():
            return get_optimized_azure_settings_for_steps(
                az["deployment_step1"].get(),
                az["deployment_step2"].get(),
                az["deployment_step3"].get(),
                detected_api_version=self.detected_azure_api_ver,
                model_meta=self.azure_model_meta,
            )

        def _on_az_dep_select(selected_val, step_key):
            az[step_key].set(selected_val)
            rec = _get_current_az_optimized()
            az["rpm"].set(str(rec["rpm"]))
            az["threads"].set(str(rec["threads"]))
            az["tokens"].set(str(rec["max_tokens"]))

        def _copy_step1_to_all():
            val = az["deployment_step1"].get().strip()
            if not val:
                return
            az["deployment_step2"].set(val)
            az["deployment_step3"].set(val)
            for cb in self.azure_combos:
                cb.set(val)
            rec = _get_current_az_optimized()
            az["rpm"].set(str(rec["rpm"]))
            az["threads"].set(str(rec["threads"]))
            az["tokens"].set(str(rec["max_tokens"]))

        for idx, (label_text, s_key) in enumerate(step_defs):
            row_dep = ctk.CTkFrame(conn_frame, fg_color="transparent")
            row_dep.pack(fill="x", padx=12, pady=2)

            ctk.CTkLabel(
                row_dep,
                text=label_text,
                width=140,
                anchor="w",
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=TEXT_COLOR,
            ).pack(side="left")

            cb_dep = ctk.CTkComboBox(
                row_dep,
                variable=az[s_key],
                values=visible_az,
                width=260,
                fg_color="white",
                border_color="#CED4DA",
                button_color="#E9ECEF",
                button_hover_color="#DEE2E6",
                dropdown_fg_color="white",
                dropdown_hover_color="#E7F1FF",
                dropdown_text_color=TEXT_COLOR,
                text_color=TEXT_COLOR,
                command=lambda val, k=s_key: _on_az_dep_select(val, k),
            )
            cb_dep.pack(side="left", padx=(0, 6))
            bind_context_menu(cb_dep)
            self.azure_combos.append(cb_dep)

            btn_test = ctk.CTkButton(
                row_dep,
                text="🔌 接続テスト" if idx == 0 else "テスト",
                width=110 if idx == 0 else 64,
                fg_color=PRIMARY if idx == 0 else "#E9ECEF",
                hover_color=PRIMARY_HOVER if idx == 0 else "#DEE2E6",
                text_color="white" if idx == 0 else TEXT_COLOR,
                border_width=0 if idx == 0 else 1,
                border_color=BORDER_COLOR,
            )
            btn_test.configure(
                command=lambda k=s_key, b=btn_test, t=("🔌 接続テスト" if idx == 0 else "テスト"): run_azure_test(
                    k, b, t
                )
            )
            btn_test.pack(side="left", padx=(0, 6))

            if idx == 0:
                ctk.CTkButton(
                    row_dep,
                    text="全ステップへ同じデプロイ名を反映",
                    command=_copy_step1_to_all,
                    fg_color="#E9ECEF",
                    hover_color="#DEE2E6",
                    text_color=TEXT_COLOR,
                    border_width=1,
                    border_color=BORDER_COLOR,
                    width=210,
                ).pack(side="left")

        # APIバージョン・最新モデル更新・主力フィルタ行
        row_ver = ctk.CTkFrame(conn_frame, fg_color="transparent")
        row_ver.pack(fill="x", padx=12, pady=(6, 4))

        ctk.CTkLabel(
            row_ver,
            text="APIバージョン:",
            width=140,
            anchor="w",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(side="left")

        combo_ver = ctk.CTkComboBox(
            row_ver,
            variable=az["api_version"],
            values=AZURE_API_VERSION_CANDIDATES,
            width=185,
            fg_color="white",
            border_color="#CED4DA",
            button_color="#E9ECEF",
            button_hover_color="#DEE2E6",
            dropdown_fg_color="white",
            dropdown_hover_color="#E7F1FF",
            dropdown_text_color=TEXT_COLOR,
            text_color=TEXT_COLOR,
        )
        combo_ver.pack(side="left", padx=(0, 10))
        bind_context_menu(combo_ver)

        def fetch_azure_models_ui(btn):
            base_ep, resolved_dep, resolved_ver = _normalize_azure_fields("deployment_step1")
            key = az["key"].get().strip()

            if not base_ep:
                messagebox.showwarning(
                    "警告", "最新モデルを取得するには、エンドポイントURLを入力してください。", parent=self
                )
                return
            if not key:
                messagebox.showwarning(
                    "警告", "最新モデルを取得するには、APIキーを入力してください。", parent=self
                )
                return

            btn.configure(text="取得中...", state="disabled")

            def _do_fetch():
                try:
                    new_az_models, best_ver, new_az_meta = fetch_azure_models_detailed(
                        key, base_ep, resolved_ver
                    )
                    if not new_az_models:
                        self.after(
                            0,
                            lambda: messagebox.showinfo(
                                "情報",
                                "取得可能なチャット対応モデルが見つかりませんでした。",
                                parent=self,
                            ),
                        )
                        return

                    def _apply_az():
                        self.detected_azure_api_ver = best_ver
                        self.azure_models_list = list(new_az_models)
                        self.azure_model_meta.update(new_az_meta)
                        self._refresh_azure_combos()
                        vis_models = self._get_visible_azure_models()
                        if vis_models:
                            for sk in ("deployment_step1", "deployment_step2", "deployment_step3"):
                                if not az[sk].get().strip():
                                    az[sk].set(vis_models[0])

                        rec = _get_current_az_optimized()
                        changes = []
                        cur_ver = az["api_version"].get().strip()
                        cur_rpm = safe_int(az["rpm"], rec["rpm"], 1, 2000)
                        cur_th = safe_int(az["threads"], rec["threads"], 1, 20)
                        cur_tp = safe_float(az["temp"], rec["temperature"], 0.0, 2.0)
                        cur_tok = safe_int(az["tokens"], rec["max_tokens"], 1024, 128000)

                        if cur_ver != rec["api_version"]:
                            changes.append(f"・APIバージョン: {cur_ver or '未設定'} → {rec['api_version']}")
                        if cur_rpm != rec["rpm"]:
                            changes.append(f"・RPM: {cur_rpm} → {rec['rpm']}")
                        if cur_th != rec["threads"]:
                            changes.append(f"・スレッド: {cur_th} → {rec['threads']}")
                        if abs(cur_tp - rec["temperature"]) > 1e-6:
                            changes.append(f"・Temp: {cur_tp} → {rec['temperature']}")
                        if cur_tok != rec["max_tokens"]:
                            changes.append(f"・最大トークン: {cur_tok} → {rec['max_tokens']}")

                        # APIバージョン・パフォーマンス設定・AI抽出パラメータを最新仕様に合わせて自動最適化
                        az["api_version"].set(rec["api_version"])
                        az["rpm"].set(str(rec["rpm"]))
                        az["threads"].set(str(rec["threads"]))
                        az["temp"].set(str(rec["temperature"]))
                        az["tokens"].set(str(rec["max_tokens"]))

                        mode_note = (
                            f"（主力モデル {len(vis_models)}件 を表示中。すべて見る場合は「主力モデルのみ表示」のチェックを外してください）"
                            if self.azure_mainstream_only_var.get()
                            else f"（全 {len(vis_models)}件 を表示中）"
                        )

                        if changes:
                            msg = (
                                f"最新のモデルリスト (全{len(self.azure_models_list)}件) を Azure から取得し、\n"
                                f"APIバージョン・パフォーマンス設定・AI抽出パラメータを自動最適化しました！\n"
                                f"{mode_note}\n\n"
                                f"【最適化された設定】\n"
                                + "\n".join(changes)
                                + "\n\n（※設定済みのデプロイ名は維持されています）"
                            )
                        else:
                            msg = (
                                f"最新のモデルリスト (全{len(self.azure_models_list)}件) を Azure から取得しました！\n"
                                f"{mode_note}\n\n"
                                f"✅ APIバージョン・パフォーマンス設定・AI抽出パラメータは最新仕様の最適値に調整済みです。"
                            )
                        messagebox.showinfo("更新・最適化完了", msg, parent=self)

                    self.after(0, _apply_az)
                except Exception as e:
                    diag = diagnose_api_error(
                        "azure", e, model_or_deployment=resolved_dep, endpoint=base_ep
                    )
                    self.after(0, lambda d=diag: show_api_error_dialog(self, d))
                finally:
                    self.after(
                        0,
                        lambda: btn.winfo_exists()
                        and btn.configure(text="🌐 最新モデルを更新", state="normal"),
                    )

            threading.Thread(target=_do_fetch, daemon=True).start()

        btn_fetch_az = ctk.CTkButton(
            row_ver,
            text="🌐 最新モデルを更新",
            width=155,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="white",
        )
        btn_fetch_az.configure(command=lambda b=btn_fetch_az: fetch_azure_models_ui(b))
        btn_fetch_az.pack(side="left", padx=(0, 12))

        ctk.CTkCheckBox(
            row_ver,
            text="主力モデルのみ表示",
            variable=self.azure_mainstream_only_var,
            command=self._refresh_azure_combos,
            text_color=TEXT_COLOR,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            border_color="#ADB5BD",
            checkbox_width=18,
            checkbox_height=18,
        ).pack(side="left")

        self.azure_count_label = ctk.CTkLabel(
            row_ver,
            text=f"({len(visible_az)} / {len(self.azure_models_list)}件)",
            text_color=MUTED_TEXT,
            font=ctk.CTkFont(size=11),
        )
        self.azure_count_label.pack(side="left", padx=(4, 0))

        ctk.CTkLabel(
            conn_frame,
            text="※独自のデプロイ名を付けている場合は直接入力も可能です。エンドポイントに長いターゲットURI全体を貼っても自動補正されます。",
            text_color=MUTED_TEXT,
            font=ctk.CTkFont(size=11),
        ).pack(anchor="w", padx=12, pady=(2, 8))

        # --- 中段左右分割 ---
        middle_frame = ctk.CTkFrame(parent_tab, fg_color="transparent")
        middle_frame.pack(fill="x", padx=6, pady=5)
        middle_frame.grid_columnconfigure(0, weight=3)
        middle_frame.grid_columnconfigure(1, weight=2)

        # --- ② パフォーマンス・並列処理設定 ---
        perf_frame = ctk.CTkFrame(
            middle_frame,
            fg_color=CARD_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        perf_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        ctk.CTkLabel(
            perf_frame,
            text=" ② パフォーマンス・並列処理設定 ",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=PRIMARY,
        ).pack(anchor="w", padx=12, pady=(8, 4))

        speed_row = ctk.CTkFrame(perf_frame, fg_color="transparent")
        speed_row.pack(fill="x", padx=12, pady=6)

        ctk.CTkLabel(
            speed_row,
            text="RPM (1分あたりの最大リクエスト数):",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(side="left")
        ctk.CTkEntry(
            speed_row,
            textvariable=az["rpm"],
            width=70,
            fg_color="white",
            border_color="#CED4DA",
            text_color=TEXT_COLOR,
        ).pack(side="left", padx=(6, 16))

        ctk.CTkLabel(
            speed_row,
            text="スレッド (並列数):",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(side="left")
        ctk.CTkEntry(
            speed_row,
            textvariable=az["threads"],
            width=58,
            fg_color="white",
            border_color="#CED4DA",
            text_color=TEXT_COLOR,
        ).pack(side="left", padx=(6, 0))

        link_row = ctk.CTkFrame(perf_frame, fg_color="transparent")
        link_row.pack(fill="x", padx=12, pady=(2, 4))
        lbl_az_link = ctk.CTkLabel(
            link_row,
            text="🔗 Azure OpenAI 公式ドキュメントを確認する",
            text_color=PRIMARY,
            cursor="hand2",
            font=ctk.CTkFont(size=12, underline=True),
        )
        lbl_az_link.pack(side="right")
        lbl_az_link.bind(
            "<Button-1>",
            lambda e: webbrowser.open_new("https://learn.microsoft.com/ja-jp/azure/ai-services/openai/"),
        )

        perf_actions = ctk.CTkFrame(perf_frame, fg_color="transparent")
        perf_actions.pack(fill="x", padx=12, pady=(6, 10))

        ctk.CTkButton(
            perf_actions,
            text="ℹ️ 設定ガイドを確認",
            command=self._show_azure_guide,
            fg_color="#E9ECEF",
            hover_color="#DEE2E6",
            text_color=TEXT_COLOR,
            border_width=1,
            border_color=BORDER_COLOR,
            width=155,
        ).pack(side="left")

        def reset_az_perf():
            rec = _get_current_az_optimized()
            az["api_version"].set(rec["api_version"])
            az["rpm"].set(str(rec["rpm"]))
            az["threads"].set(str(rec["threads"]))

        ctk.CTkButton(
            perf_actions,
            text="🔄 推奨値",
            command=reset_az_perf,
            fg_color="#E9ECEF",
            hover_color="#DEE2E6",
            text_color=TEXT_COLOR,
            border_width=1,
            border_color=BORDER_COLOR,
            width=95,
        ).pack(side="right")

        # --- ③ AI抽出パラメータ設定 ---
        param_frame = ctk.CTkFrame(
            middle_frame,
            fg_color=CARD_BG,
            corner_radius=8,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        param_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        ctk.CTkLabel(
            param_frame,
            text=" ③ AI抽出パラメータ設定 ",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=PRIMARY,
        ).pack(anchor="w", padx=12, pady=(8, 4))

        param_row1 = ctk.CTkFrame(param_frame, fg_color="transparent")
        param_row1.pack(fill="x", padx=12, pady=6)

        ctk.CTkLabel(
            param_row1, text="Temp:", font=ctk.CTkFont(size=12, weight="bold"), text_color=TEXT_COLOR
        ).pack(side="left")
        ctk.CTkEntry(
            param_row1,
            textvariable=az["temp"],
            width=62,
            fg_color="white",
            border_color="#CED4DA",
            text_color=TEXT_COLOR,
        ).pack(side="left", padx=(6, 14))

        ctk.CTkLabel(
            param_row1,
            text="最大トークン:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(side="left")
        ctk.CTkEntry(
            param_row1,
            textvariable=az["tokens"],
            width=86,
            fg_color="white",
            border_color="#CED4DA",
            text_color=TEXT_COLOR,
        ).pack(side="left", padx=(6, 0))

        ctk.CTkLabel(
            param_frame,
            text="※推論モデル(GPT-5/6, o4等)は自動で最適パラメータに調整されます",
            text_color=MUTED_TEXT,
            font=ctk.CTkFont(size=11),
            wraplength=360,
            justify="left",
        ).pack(anchor="w", padx=12, pady=(2, 4))

        param_row3 = ctk.CTkFrame(param_frame, fg_color="transparent")
        param_row3.pack(fill="x", padx=12, pady=(6, 10))

        def reset_az_param():
            rec = _get_current_az_optimized()
            az["temp"].set(str(rec["temperature"]))
            az["tokens"].set(str(rec["max_tokens"]))

        ctk.CTkButton(
            param_row3,
            text="🔄 推奨値",
            command=reset_az_param,
            fg_color="#E9ECEF",
            hover_color="#DEE2E6",
            text_color=TEXT_COLOR,
            border_width=1,
            border_color=BORDER_COLOR,
            width=95,
        ).pack(side="right")

        def _reset_azure_tab_all():
            reset_az_perf()
            reset_az_param()
            rec = _get_current_az_optimized()
            messagebox.showinfo(
                "推奨値に戻しました",
                f"Azure OpenAI Service の設定をすべて最新の最適化推奨値に戻しました。\n\n"
                f"・APIバージョン: {rec['api_version']}\n"
                f"・RPM: {rec['rpm']} / スレッド: {rec['threads']}\n"
                f"・Temp: {rec['temperature']} / 最大トークン: {rec['max_tokens']}",
                parent=self,
            )

        self.tab_reset_handlers[TAB_AZURE] = _reset_azure_tab_all

        # --- ④ カスタムプロンプト ---
        self._build_prompt_section(
            parent_tab,
            " ④ 独自の追加指示 (カスタムプロンプト) - 任意 ",
            lambda: az["prompts"],
            lambda items: az.__setitem__("prompts", items),
        )

    # ==============================
    # 仕様・制限ガイドウィンドウ (Gemini / Azure)
    # ==============================
    def _create_info_table(self, parent, headers, rows, col_weights, header_bg=HEADER_BG, header_fg=TEXT_COLOR):
        frame = ctk.CTkFrame(parent, fg_color=BORDER_COLOR, corner_radius=4)
        frame.pack(fill="x", padx=14, pady=(4, 14))
        for i, weight in enumerate(col_weights):
            frame.grid_columnconfigure(i, weight=weight)
        for col_idx, text in enumerate(headers):
            lbl = ctk.CTkLabel(
                frame,
                text=text,
                font=ctk.CTkFont(size=12, weight="bold"),
                fg_color=header_bg,
                text_color=header_fg,
                corner_radius=0,
                padx=8,
                pady=8,
                justify="left",
                anchor="w",
            )
            lbl.grid(row=0, column=col_idx, sticky="nsew", padx=1, pady=1)
        for row_idx, row_data in enumerate(rows, 1):
            row_bg = CARD_BG if row_idx % 2 == 1 else "#F8FAFC"
            for col_idx, text in enumerate(row_data):
                lbl = ctk.CTkLabel(
                    frame,
                    text=text,
                    font=ctk.CTkFont(size=11, weight="bold" if col_idx == 0 else "normal"),
                    fg_color=row_bg,
                    text_color=TEXT_COLOR,
                    corner_radius=0,
                    justify="left",
                    anchor="nw",
                    padx=8,
                    pady=8,
                    wraplength=280 if col_idx > 0 else 180,
                )
                lbl.grid(row=row_idx, column=col_idx, sticky="nsew", padx=1, pady=1)

    def _show_gemini_limit_info(self, is_free):
        info_win = ctk.CTkToplevel(self)
        info_win.title("Gemini API 制限と仕様一覧")
        info_win.configure(fg_color=BG_COLOR)
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        center_window(info_win, self, min(980, screen_w - 40), min(720, screen_h - 80))
        info_win.transient(self)
        info_win.grab_set()

        def _close_info():
            info_win.destroy()
            try:
                if self.winfo_exists():
                    self.grab_set()
            except Exception:
                pass

        info_win.protocol("WM_DELETE_WINDOW", _close_info)

        scroll = ctk.CTkScrollableFrame(info_win, fg_color=BG_COLOR)
        scroll.pack(fill="both", expand=True, padx=8, pady=8)

        ctk.CTkLabel(
            scroll,
            text="Gemini API 仕様・制限一覧",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=PRIMARY,
        ).pack(pady=(10, 12))

        ctk.CTkLabel(
            scroll,
            text="▼ プラン比較",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(anchor="w", padx=14)
        self._create_info_table(
            scroll,
            ["比較項目", "無料枠 (Free Tier)", "課金枠 (Paid Tier)"],
            [
                ["利用料金", "完全無料（クレジットカード登録不要）", "従量課金（トークンと呼ばれるデータ量に応じて支払い）"],
                [
                    "利用できるモデル",
                    "Gemini 4 Hydrogen / 3.8 Flash / 3.5 Flash-Lite など",
                    "すべての最新モデル（Gemini 4 Carbon / Neon / Argon 含む）",
                ],
                [
                    "データのプライバシー",
                    "入力データがGoogleのAI学習に利用される可能性がある",
                    "入力データはAI学習に利用されない",
                ],
            ],
            [1, 3, 3],
        )

        known_model_info = {
            "gemini-4-hydrogen": {
                "limit": ["15 RPM目安\n[推奨: 12 RPM / 直列(1)]", "高スループット\n[推奨: 300 RPM / 並列(5〜)]"],
                "desc": [
                    "Gemini 4 標準主力モデル",
                    "汎用性・スピード・マルチモーダル精度の高いバランス",
                    "一般的な機械図面の部品・要素抽出（Step1 / Step2 のデフォルト推奨）",
                ],
            },
            "gemini-4-argon": {
                "limit": ["制限目安は公式参照\n[推奨: 5 RPM / 直列(1)]", "時期・枠により変動\n[推奨: 200 RPM / 並列(5)]"],
                "desc": [
                    "高効率フロンティア推論",
                    "高い知能と高速応答・トークン効率の両立",
                    "複雑な三面図の解析と処理速度を両立したい場合",
                ],
            },
            "gemini-4-neon": {
                "limit": ["制限目安は公式参照\n[推奨: 2 RPM / 直列(1)]", "時期・枠により変動\n[推奨: 150 RPM / 並列(5)]"],
                "desc": [
                    "構造化抽出・高度推論モデル",
                    "複雑な構造化データ抽出と段階的思考",
                    "見落とし検証や複雑な図面の精密解析（Step3 のデフォルト推奨）",
                ],
            },
            "gemini-4-carbon": {
                "limit": ["非常に厳しい\n[推奨: 1 RPM / 直列(1)]", "時期・枠により変動\n[推奨: 100 RPM / 並列(3〜5)]"],
                "desc": [
                    "最上位フロンティアモデル",
                    "極めて複雑な図面構造解析、高度な立体推論",
                    "絶対にミスが許されない複雑なアセンブリ図面の解析",
                ],
            },
            "gemini-3.8-flash": {
                "limit": ["15 RPM目安\n[推奨: 12 RPM / 直列(1)]", "1000 RPM\n[推奨: 300 RPM / 並列(5〜)]"],
                "desc": [
                    "高速・軽量Flash最新世代",
                    "低遅延・大量部品処理・マルチモーダル対応",
                    "多数の部品を高速に並列解析したい場合",
                ],
            },
            "gemini-3.5-flash-lite": {
                "limit": ["15 RPM目安\n[推奨: 12 RPM / 直列(1)]", "1000 RPM\n[推奨: 300 RPM / 並列(5〜)]"],
                "desc": [
                    "最軽量・低コストモデル",
                    "圧倒的な処理スピードと低コスト",
                    "鮮明なCAD図面の簡易抽出、コストを抑えたい場合",
                ],
            },
        }

        data_limit = []
        data_model = []
        for m_display, m_id in self._get_visible_gemini_models():
            info = known_model_info.get(m_id)
            if not info:
                for known_id, known_data in known_model_info.items():
                    if known_id in m_id:
                        info = known_data
                        break
            if info:
                data_limit.append([m_display, info["limit"][0], info["limit"][1]])
                data_model.append([m_display, info["desc"][0], info["desc"][1], info["desc"][2]])
            else:
                data_limit.append([m_display, "詳細は公式ドキュメントを参照", "詳細は公式ドキュメントを参照"])
                data_model.append([m_display, "APIから取得した追加モデル", "-", "最新機能を試したい場合"])

        ctk.CTkLabel(
            scroll,
            text="▼ 各モデルの制限目安 (RPM と 推奨スレッド数)",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(anchor="w", padx=14)
        self._create_info_table(
            scroll,
            ["モデル名", "無料枠の制限目安\n(RPM / スレッド数)", "課金枠の制限目安\n(RPM / スレッド数)"],
            data_limit,
            [2, 3, 3],
        )

        ctk.CTkLabel(
            scroll,
            text="▼ 各モデルの特徴と適した用途",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(anchor="w", padx=14)
        self._create_info_table(
            scroll, ["モデル名", "特徴", "得意なこと", "適した用途"], data_model, [2, 2, 3, 3]
        )

        plan_type = "free" if is_free else "paid"
        v = self.gemini_vars[plan_type]
        status_text = (
            f"【現在、このタブで選択中の設定】\n"
            f"プラン: {'無料枠 (Free Tier)' if is_free else '課金枠 (Paid Tier)'}\n"
            f"Step1: {v['model_step1'].get()}  ／  Step2: {v['model_step2'].get()}  ／  Step3: {v['model_step3'].get()}"
        )
        ctk.CTkLabel(
            scroll,
            text=status_text,
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=PRIMARY,
            fg_color=CARD_BG,
            corner_radius=8,
            padx=12,
            pady=10,
        ).pack(fill="x", padx=14, pady=10)

        ctk.CTkButton(
            scroll,
            text="閉じる",
            command=_close_info,
            width=150,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="white",
        ).pack(pady=(4, 16))

    def _show_azure_guide(self):
        info_win = ctk.CTkToplevel(self)
        info_win.title("Azure OpenAI Service 仕様・設定ガイド")
        info_win.configure(fg_color=BG_COLOR)
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        center_window(info_win, self, min(980, screen_w - 40), min(680, screen_h - 80))
        info_win.transient(self)
        info_win.grab_set()

        def _close_info():
            info_win.destroy()
            try:
                if self.winfo_exists():
                    self.grab_set()
            except Exception:
                pass

        info_win.protocol("WM_DELETE_WINDOW", _close_info)

        scroll = ctk.CTkScrollableFrame(info_win, fg_color=BG_COLOR)
        scroll.pack(fill="both", expand=True, padx=8, pady=8)

        ctk.CTkLabel(
            scroll,
            text="🔷 Azure OpenAI Service 設定・モデルガイド",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=PRIMARY,
        ).pack(pady=(10, 12))

        ctk.CTkLabel(
            scroll,
            text="▼ 1. 接続設定に必要な項目（Azure Portal / Azure AI Foundry で確認）",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(anchor="w", padx=14)
        self._create_info_table(
            scroll,
            ["設定項目", "入力内容の例", "確認場所・ポイント"],
            [
                [
                    "エンドポイントURL",
                    "https://xxxx.openai.azure.com/",
                    "Azure Portal の「キーとエンドポイント」画面に表示されるURLです。\n※誤って長いターゲットURI全体を貼り付けても自動でベースURLを抽出します。",
                ],
                ["APIキー", "32文字以上の英数字キー", "「KEY 1」または「KEY 2」のいずれかをコピーして貼り付けます。"],
                [
                    "デプロイ名\n(Step1〜3)",
                    "gpt-5 / gpt-5-mini / gpt-5.2 など",
                    "Azure AI Foundry の「デプロイ」画面で作成したモデルの名前です。全ステップ同じデプロイ名でも、Step3だけ高精度モデルにすることも可能です。",
                ],
                [
                    "APIバージョン",
                    f"{DEFAULT_AZURE_API_VERSION} (推奨・最新標準)",
                    "最新の次世代 v1 API（日付指定不要の /openai/v1/）が標準です。従来の日付形式（2025-04-01-preview 等）が必要な環境ではプルダウンから選択できます。「🌐 最新モデルを更新」を押すと利用可能な最新方式を自動判定します。",
                ],
            ],
            [2, 3, 5],
            header_bg=AZURE_HEADER_BG,
            header_fg=PRIMARY,
        )

        ctk.CTkLabel(
            scroll,
            text="▼ 2. 画像解析（Vision）に対応した推奨モデル一覧",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(anchor="w", padx=14)
        self._create_info_table(
            scroll,
            ["モデル種別", "特徴", "おすすめの用途", "推奨設定目安 (RPM / 並列数)"],
            [
                [
                    "gpt-5.2 / gpt-5.1\ngpt-5 / gpt-6",
                    "最新主力フラグシップ",
                    "複雑な三面図の立体復元、見落とし検証（Step3 に最もおすすめ）",
                    "60 RPM / 並列 3〜5",
                ],
                [
                    "gpt-5-mini\ngpt-5-nano",
                    "高速・低コスト・軽量主力",
                    "部品一覧抽出・多数部品の並列解析（Step1 / Step2 におすすめ）",
                    "120 RPM / 並列 5〜10",
                ],
                [
                    "o4-mini",
                    "高度な推論モデル",
                    "複雑な寸法計算や多段推論が必要な図面（※自動で推論モデル用パラメータへ切替）",
                    "30 RPM / 並列 1〜3",
                ],
            ],
            [2, 2, 4, 2],
            header_bg=AZURE_HEADER_BG,
            header_fg=PRIMARY,
        )

        ctk.CTkButton(
            scroll,
            text="閉じる",
            command=_close_info,
            width=150,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            text_color="white",
        ).pack(pady=(6, 16))

    # ==============================
    # アクション: リセット・収集・保存・キャンセル
    # ==============================
    def reset_current_tab_all(self):
        try:
            current_tab = self.tabview.get()
            handler = self.tab_reset_handlers.get(current_tab)
            if handler:
                handler()
        except Exception:
            pass

    def _collect_settings(self, normalize_azure=True):
        prof = self.active_profile_var.get()
        if prof == "azure":
            engine = "azure"
            plan = self._last_gemini_plan
        elif prof == "gemini_paid":
            engine = "gemini"
            plan = "paid"
        else:
            engine = "gemini"
            plan = "free"

        collected = {
            "engine": engine,
            "plan": plan,
            "models_list": [list(m) for m in self.models_list],
            "gemini_model_meta": copy.deepcopy(self.gemini_model_meta),
            "gemini_mainstream_only": bool(self.gemini_mainstream_only_var.get()),
            "azure_models_list": list(self.azure_models_list),
            "azure_model_meta": copy.deepcopy(self.azure_model_meta),
            "azure_mainstream_only": bool(self.azure_mainstream_only_var.get()),
            "saved_prompts": list(self.saved_prompts),
        }

        # Gemini Free / Paid
        for plan_type in ("free", "paid"):
            is_free = plan_type == "free"
            v = self.gemini_vars[plan_type]
            cbs = self.gemini_tab_controls[plan_type]["combos"]

            for step_idx, step_name in enumerate(["step1", "step2", "step3"]):
                mid = self._display_to_model_id(cbs[step_idx].get())
                collected[f"{plan_type}_model_{step_name}"] = mid

            perf_dict = copy.deepcopy(self.settings.get(f"{plan_type}_model_perf", {}) or {})
            for m_id, p_vars in v["perf_vars"].items():
                rec = get_recommended_gemini_settings(
                    m_id, is_free=is_free, model_meta=self.gemini_model_meta
                )
                perf_dict[m_id] = {
                    "rpm": safe_int(p_vars["rpm"], rec["rpm"], 1, 2000),
                    "threads": safe_int(p_vars["threads"], rec["threads"], 1, 20),
                }
            collected[f"{plan_type}_model_perf"] = perf_dict
            collected[f"{plan_type}_key"] = v["key"].get().strip()
            collected[f"{plan_type}_temp"] = safe_float(v["temp"], 0.0, 0.0, 2.0)
            collected[f"{plan_type}_tokens"] = safe_int(v["tokens"], 65535, 1024, 2097152)
            collected[f"{plan_type}_safety"] = bool(v["safety"].get())
            collected[f"{plan_type}_prompts"] = list(v["prompts"])

        # Azure OpenAI Service
        az = self.azure_vars
        ep_val = az["endpoint"].get().strip()
        dep1 = az["deployment_step1"].get().strip()
        dep2 = az["deployment_step2"].get().strip()
        dep3 = az["deployment_step3"].get().strip()
        raw_ver_val = az["api_version"].get().strip()

        if normalize_azure and ep_val:
            base_ep, parsed_dep, parsed_ver = parse_azure_endpoint_info(ep_val, dep1, raw_ver_val)
            ep_val = base_ep
            if parsed_dep and not dep1:
                dep1 = parsed_dep
            if parsed_dep and not dep2:
                dep2 = parsed_dep
            if parsed_dep and not dep3:
                dep3 = parsed_dep
            ver_val = parsed_ver or DEFAULT_AZURE_API_VERSION
        elif normalize_azure:
            ver_val = raw_ver_val or DEFAULT_AZURE_API_VERSION
        else:
            ver_val = raw_ver_val

        collected["azure_endpoint"] = ep_val
        collected["azure_api_key"] = az["key"].get().strip()
        collected["azure_deployment_step1"] = dep1
        collected["azure_deployment_step2"] = dep2 or dep1
        collected["azure_deployment_step3"] = dep3 or dep1
        collected["azure_api_version"] = ver_val
        collected["azure_rpm"] = safe_int(az["rpm"], 60, 1, 2000)
        collected["azure_threads"] = safe_int(az["threads"], 3, 1, 20)
        collected["azure_temp"] = safe_float(az["temp"], 0.0, 0.0, 2.0)
        collected["azure_tokens"] = safe_int(az["tokens"], 8192, 1024, 128000)
        collected["azure_prompts"] = list(az["prompts"])

        return collected

    def _has_changes(self):
        return self._collect_settings(normalize_azure=False) != self._initial_snapshot

    def cancel_and_close(self):
        if self._has_changes():
            if not messagebox.askyesno(
                "確認",
                "変更が適用されていません。\n破棄して設定画面を閉じますか？",
                parent=self,
            ):
                return
        self.destroy()

    def save_and_close(self):
        new_settings = self._collect_settings(normalize_azure=True)
        self.on_save_callback(new_settings)
        self.destroy()
